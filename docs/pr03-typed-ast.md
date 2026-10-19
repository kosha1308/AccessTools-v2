# PR-03 — контракт типизированного AST

Основание: задание PR-03, main b620947 после PR-02, expression-model PR-01 и ТЗ v0.15 (прочитано через Drive). Этот контракт конкретизирует уже заданную модель; выполнения AST нет.

## План и границы

1. Сохранить 96 регрессий PR-02 и наблюдать RED отсутствующего AST API — выполнено.
2. Добавить один record-класс узла, типизированный каталог метаданных и статический валидатор; проверить арность, граф, scope, типы и бюджеты — выполнено.
3. Расширить закрытый XML явной схемой 2.0.0; 1.0.0 остаётся legacy deferred без переинтерпретации. Два независимых примера и отрицательные фикстуры — выполнено.
4. Выполнить VBA в Excel x64 и независимое ревью — выполнено; результат передаётся отдельным PR без merge. PR-04 не выполняется.

## Представление

Expression содержит RootNodeId, ScopeTemplateId и плоскую коллекцию CAtNode. Узел: Id, Kind, ResultType, Values (закрытые строковые свойства), Parameters (имя → Type/Value), Bindings (ordered BindingId), Children (ordered NodeId). Такой контейнер позволяет позднее добавлять ColumnRef/collection nodes с теми же IDs и лексическими bindings; они сейчас отвергаются как неизвестные kind. Нет хранения формул/макросов/handler names.

Одна Expression — дерево: NodeId уникальны в Expression, root имеет ноль родителей, прочие узлы — ровно одного. Циклы, unreachable nodes, повторное использование дочернего ID и dangling ссылки запрещены. Для повторения значения есть Let/ValueRef; положение узла не является его ID. Порядок детей сохраняется независимо от порядка записей Nodes.

ScopeTemplateId явно закрепляет документный контекст FieldRef; Rule обязан принадлежать этому Template. Let: первые N children — значения N упорядоченных Bindings, последний child — body. Значение видит только внешние и предыдущие bindings; body видит все. BindingId уникальны во всей Expression, одинаковые DisplayName не участвуют в разрешении. Nested Let создаёт локальное расширение scope, не изменяет sibling/outer scope. ReturnType Let = body. ValueRef — BindingId, не имя поля. Неявного захвата row alias нет.

TypedLiteral: State=OK или EMPTY; Value обязателен (пустая строка для EMPTY). OK payload — каноническая строка объявленного типа, без преобразования Text в другие типы. StateIs: один child, States — список через `|`, выход Bool. Call: OperationId и точная OperationVersion; args в Children; параметры типизированы и совпадают с сигнатурой. AllOf/AnyOf требуют минимум один Bool child, Not — один Bool; When — Bool gate/then/else с одинаковыми типами ветвей и результата.

## Состояния — контракт, не вычисление

| State | Значение |
| --- | --- |
| OK | Есть корректный payload указанного типа |
| EMPTY | Существующее scalar значение пусто; не FALSE/0 |
| MISSING | Нет обязательного поля/locator/member |
| INVALID | Значение не соответствует грамматике/типу |
| UNAVAILABLE | Нет доступа к источнику/ресурсу |
| ERROR | Ошибка вычисления/предпосылки |
| CONFLICT | Неоднозначность/несовместимые результаты |
| LIMIT | Исчерпан бюджет |
| NOT_APPLICABLE | Явно неприменимое вычисление |

Контракт будущего Evaluator: strict Call/Not распространяют non-OK; When не выбирает else при non-OK gate. AllOf/AnyOf short-circuit только по OK Bool, без coercion ошибок. StateIs наблюдает состояние, но не снимает техническую неполноту и не даёт разрешение PASS. ValueRef/Let сохраняют envelope и diagnostics. EMPTY не равен MISSING; NOT_EVALUATED — будущий trace event, не EvalState. Эффекты completeness/trace не вычисляются PR-03 и обязательны до заявлений о Runtime.

Типы выражений ограничены существующими scalar именами PR-02. OK literal поддерживает Text, Bool, Int64 (включая весь signed 64-bit диапазон строкой), Decimal (канонический десятичный текст до 28 значащих цифр и до 28 цифр дробной части; ведущие нули не расходуют precision), Date и DateTime (UTC). Числа не допускают ведущий плюс/нули, экспоненту, `-0` или завершающие дробные нули; ноль пишется `0`. Date поддерживает годы 0100–9999 по контракту VBA Date; DateTime требует UTC `Z`. Email/Hostname/IpSet/PortSet допустимы для FieldRef/сигнатур и EMPTY, но OK literal для них требует будущего отдельного codec и сейчас отклоняется. List/Record/табличные scopes ещё не реализованы; контейнер не требует замены для их добавления.

## Доверенный каталог

AtSignature(ReturnType, ordered ArgumentTypes) создаёт метаданные: ReturnType, Arguments, Parameters (name→type), States и Propagation=Strict. States — непустой набор известных собственных состояний операции при OK-аргументах, обязательно OK; Strict отдельно требует распространения любого non-OK входа, независимо от этого списка. Только хост регистрирует Dictionary `OperationId@exactVersion → signature`; XML хранит ссылки и значения параметров, не сигнатуры и не код. Сигнатура проверяется до использования; типы/версия/арность/обязательные и лишние параметры диагностируются. Примеры каталога в тестах синтетические, алгоритмы отсутствуют.

Для старой схемы 1.0.0 сохранён каталог строк ReturnType PR-02. Он не считается типизированной сигнатурой AST; Call с таким каталогом отклоняется. Хост может использовать полноценную сигнатуру и при проверке legacy refs, но это не повышает legacy до AST-ready.

## XML и безопасность

SchemaVersion 2.0.0 использует тот же container namespace `urn:accesstools:definitions:1`, различается строго по SchemaVersion. У каждой Expression вместо Deferred=true: Deferred=false и единственная группа `<ast root="…" template="…">`; узлы плоские, каждый содержит ровно groups values/parameters/bindings/children. У прочих Kind ast запрещён. Namespace обозначает семейство контейнера, SchemaVersion — точный контракт. Версия 1.0.0 не допускает ast или in-memory AST; её writer не добавляет группы и не меняет версию. Upgrade — отдельное явное действие вне PR-03; старый reader обязан отказать 2.0.0.

PR-02 XML ограничения, запрет DTD/external entities, scalar escaping и append-only verified save сохраняются. XML shape error AST включает ObjectId и NodeId, если они уже прочитаны. Статические diagnostics AST всегда содержат ObjectId Expression/Rule и NodeId (root/проблемный узел); узел доступен будущему пошаговому редактору напрямую. Основной UI остаётся понятным редактором условий поверх дерева, не ручным XML/AST программированием.

Бюджеты хоста: до 512 nodes/Expression, 4096 nodes/package, depth 64; параметры API могут только уменьшать лимиты. XML по-прежнему ≤1 MiB/depth32, определения ≤2000. Лимиты проверяются до рекурсивного обхода; root walk ограничен visited и depth. Это бюджеты валидации, не runtime time/comparison budget. Структурная корректность не заменяет публикацию, подпись, capability checks или ручную приёмку.

## Проверка и приёмка

19.10.2026, реальный Excel **16.0 build 17932, Office64**, Windows x64: `PASS|96|0||Office64` и `PASS|AST|76|0` — **172 assertions PASS, 0 FAIL**. Исходники импортированы в собственную пустую тестовую книгу, VBA действительно исполнен. Это не запуск Evaluator и не ручная приёмка UI. Пробы MSXML6 x86/x64 сохранены из PR-02; Office **x86 VBA не проверен**.

Проверены оба положительных XML ([simple](../tests/fixtures/simple-ast.xml), [equipment-return](../tests/fixtures/equipment-return-ast.xml)) и оба отрицательных ([type](../tests/fixtures/simple-ast-invalid.xml), [scope](../tests/fixtures/equipment-return-ast-invalid.xml)). Один и тот же pipeline load/validate/save/reload: типы/параметры/ссылки, Unicode/CR/LF/TAB/percent, порядок детей и bindings, перемещение Field без изменения RuleId/NodeId. В EQUIPMENT_RETURN сохраняются Table/Source/Binding из PR-02, но AST ещё не читает строки таблицы.

Негативные тесты: неизвестные OperationId/version, legacy строка вместо полной сигнатуры, арность и тип аргумента/результата, лишний/неверный/отсутствующий параметр, недопустимые states/Propagation, неизвестные узлы/типы/поля/версии, missing child, дубли NodeId, цикл/осиротевший узел, forward/self ValueRef, утечка вложенного binding в sibling, дубли BindingId, отсутствующий scope. Проверены уменьшенные host budgets и реальные hard limits 512 nodes / depth64 / 4096 nodes per package. Неизвестная schema, AST под schema1, DTD, неожиданный AST XML и duplicate parameters отклонены. Граничные Int64, Decimal precision/scale, даты и запрет неявного Text→IP покрыты отдельно.

Все 96 регрессий PR-02 сохранены; единственная актуализация теста старой схемы — значение *неподдерживаемой* версии заменено с `2.0.0` на `9.0.0`, поскольку `2.0.0` теперь явно реализована. Legacy fixtures не переписаны. Неполный Draft/publish-ready остаются ограничениями PR-02, не обходятся AST type checking.

Независимое source review: найден Decimal precision defect, воспроизведены два FAIL до исправления, после исправления полный PASS; повторное ревью подтвердило исправление. Остановка первого AST-прогона на зарезервированном VBA `Local` устранена до положительного результата. Рецензент не заявляет отдельного запуска Office. Локальные host/hashes/assertion logs находятся в игнорируемом `artifacts/run-…`, Office binaries в PR не добавлены.

Воспроизведение из checkout: `./tests/Probe-MSXML.ps1`, затем `./tests/Run-Tests.ps1`. Требования и политика доступа к VBProject — [PR-02 validation](pr02-validation.md); runner не меняет Trust Center, открывает только собственную пустую книгу. Дополнительно к прежнему result.txt создаются ast-result.txt и ast-assertions.txt. Ожидается 96 legacy и 76 AST PASS. Для приёмки сравнить root/ordered children в созданных XML и проверить NodeId в отрицательных fixtures. Ручная работа Admin не тестируется, потому что UI не создан.

Технические решения для ревью: flat tree со стабильными ID, schema2 без неявного upgrade, фиксированные сигнатуры/параметры из доверенного хоста, последовательный lexical Let, бюджеты и ограниченный literal codec. Нет поддержки выполнения AST, операций, source/table reading, scope инстанцирования строк, completeness effects, publish/digest/rollback или XLAM UI. Эти ограничения явные; validation success не разрешает Runtime/PASS. PR-04 не начат, v1/архив/Google Drive не изменены.
