# Условия, типы и вычисление: проект PR-01

Статус: предложение к архитектурному ревью. Здесь описана **модель дерева**, а не реализованный язык, parser, VBA или рабочий пакет проверок. Основания — [rule-model](rule-model.md), ТЗ v0.15 §3.5–3.6 и ADR-007. Схемы в [матрице](expressiveness-matrix.md) — читаемое обозначение таких деревьев, не исполняемые конфигурации.

## 1. Представление и альтернативы

Рекомендуется authoritative typed AST: визуальный редактор напрямую создаёт узлы с устойчивыми NodeId и ссылками. Публикация хранит каноническую структурную форму; текстовый preview строится из неё и пока read-only. Это сохраняет одинаковую семантику сложных условий в Admin и Runtime. Формат физической сериализации (JSON/XML/Excel tables) — отдельное решение, AST от него не зависит.

Альтернатива «текстовый DSL как источник» требует parser/roundtrip/позиционных diagnostics и двух полноценных UI; расширяет первый этап. Альтернатива плоским condition rows проста для AND, но неудобна для lexical aliases, correlated ANY/ALL и scopes. Рекомендуемый AST использует таблицы только как возможное хранение, а не как ограничение выражений. Выбор ещё подлежит ревью; PR-01 не вводит свободное редактирование исполняемого текста.

## 2. Типы и значения

| Семейство | Типы / payload | Ограничения |
| --- | --- | --- |
| Scalar | Bool, Int64, Decimal, Text, Date, DateTime, Email, Hostname, Enum<DomainId> | Число/дата не преобразуются из строки без Call parse; текст хранит Unicode |
| Domain | IpSet, PortSet, AddressParse, DeclarationSet | IpSet = нормализованные непересекающиеся закрытые интервалы UInt32; PortSet — интервалы 1…65535 |
| Structured | Record<SchemaId>, List<T>, Set<T>, Map<K,V>, Group<K,T>, LookupMatches<T>, Coverage<T> | Статическая схема members, тип ключа и comparator закреплены |
| Provenance | CellRef, RowRef, SourceRowRef, TargetRef, DiagnosticRef | Только immutable identifiers/metadata; живых Excel-объектов нет |
| Input envelope | EvalResult<T> | State раздельно с payload; отсутствие не заменяется 0/False/пустой строкой |

`Int64` и `Decimal` — целевые смысловые типы, способ безопасного хранения в VBA предстоит проверить отдельно; нельзя использовать signed Long для UInt32 адреса или счётчика /0. Мощность IPv4 /0 — 4 294 967 296. Date — календарная дата без часового пояса; ContextRef(CLOCK_DATE) фиксируется на начало run в настроенной локали/зоне. Excel serial/epoch/error value преобразует Resolver с явной диагностикой, а не математическая операция.

Каждый член Record<SchemaId> уже сейчас содержит EvalResult<MemberType>, а сама collection может быть OK при отдельных пустых/невалидных cells. Отсутствующая обязательная колонка — MISSING/schema diagnostic, не сотни пустых строк. Пустая существующая cell — EMPTY, плохое содержимое — INVALID с RowRef. Для LEFT Join без match правая alias имеет EMPTY<RightRecord>, Member/ColumnRef распространяют EMPTY; ненайденное имя member в schema — MISSING/type diagnostic, не unmatched join. Проверка StateIs правой alias позволяет отличить отсутствие match от ошибки чтения источника.

Text сравнивается с объявленным comparator (binary/casefold/нормализация); никакого неявного глобального Trim/split. Email/hostname нормализуются соответствующей операцией с заданными параметрами, а DisplayText сохраняет исходник. Поле free-text может содержать переносы и запятые без превращения в список. Enum допускает только domain values; неизвестное значение INVALID, не новый молчаливый вариант.

IpSet не хранит форму ввода: IP, /32 и single-element range могут иметь одинаковое множество. DeclarationSet хранит Kind, raw/normalized declaration, prefix/base и provenance. OP_DECLARATIONS_COMPARE принимает policy сопоставления; OP_IP_EQUAL сравнивает множества. Это разные типы и операции. Окончательный выбор Q-07 не требуется для самого существования обоих типов.

## 3. EvalResult, пустота и ошибки

| State | Смысл | Допустимый следующий шаг |
| --- | --- | --- |
| OK | Типизированное value доступно | Обычный Call/predicate/member access |
| EMPTY | Scalar/объявленное optional значение найдено, но пусто | Явный StateIs / IsEmpty; не обычное сравнение |
| MISSING | Обязательный locator/ref/member отсутствует | Диагностика resolver/schema; отдельная явно обработанная ветвь |
| INVALID | Значение есть, но не соответствует типу/грамматике | Ошибка пользовательского ввода или конфигурации с provenance |
| UNAVAILABLE | Источник/ресурс недоступен или snapshot неустойчив | CANNOT_EVALUATE затронутого вычисления, retry отдельно |
| ERROR | Вычислительное исключение / нарушенная предпосылка | Диагностика, без Boolean coercion |
| CONFLICT | Неоднозначный locator/selection/incompatible structured outputs | Ручной выбор/исправление данных, не первый случайный кандидат |
| LIMIT | Исчерпан объявленный бюджет | Остановка зависимого вычисления, incomplete result |
| NOT_APPLICABLE | Явная неприменимость дочернего вычисления | Только объявленная policy узла; не глобальный False |

List с нулём элементов — **OK([])**, Set без элементов — OK(∅), Count=0. Они не EMPTY. EMPTY scalar не равен MISSING locator или INVALID token. Понятие NULL: отдельного значения SQL-null и трёхзначной SQL-логики нет; `TypedLiteral(Empty<T>)` возвращает EMPTY<T>. Сравнение EMPTY=EMPTY не выдаёт TRUE; «оба пусты» выражается двумя StateIs. Member Record имеет собственный envelope, а не неописанный Variant Null.

Diagnostic содержит Code, category (input/source/schema/computation/output), originating NodeId, source/row/field refs, redacted parameters и handled-by trace. Наружу не уходят секретные исходные значения без разрешённого Explain-профиля. При нескольких обязательных ошибках rule собирает диагностические записи отдельно от пользовательских замечаний.

Дополнительный effect `EvaluationRequirements` переносится вместе с результатами, в том числе OK Bool после StateIs: `(RequirementId, origin NodeId, dependency, required branch, classification, disposition)`. Required технический MISSING/UNAVAILABLE/ERROR/CONFLICT/LIMIT остаётся UNRESOLVED для rule/run completeness; ни StateIs, ни Message Active=false, ни valid-subset projection его не снимают. Когда зависимость не достигнута по short-circuit, requirement не создаётся. Принимаемая бизнес-классификация EMPTY/INVALID может быть завершена явным finding FAIL либо явно разрешённым domain case (например, текстовая группа как группа); это не разрешение подтвердить невалидный IpSet. Valid-subset сохраняет requirement полноты исходного запроса, пока invalid часть не классифицирована. Нельзя переназначить required source в optional только ради успеха потребителя его данных: compiler выводит requiredness из фактически используемой decision dependency; optional enrichment может отсутствовать только без утверждений, зависящих от него.

Таким образом ASSERT StateIs(SourceRows(requiredBinding),UNAVAILABLE) может иметь Bool TRUE, но его RuleOutcome всё равно CANNOT_EVALUATE из unresolved requirement. Аналогичный gate не может сделать ошибку NOT_APPLICABLE задним числом. Операции/узлы имеют contract propagation этих effects; проверка полноты — обязанность валидатора и evaluator, а не цвета/текста Action. Подробная реализация effect tracking остаётся отдельным инженерным этапом, этот контракт нужен до её начала.

Обычные операции strict: non-OK аргумент распространяет состояние; они не вызываются с неизвестным payload. Исключения — introspection StateIs и специализированный parser, чья сигнатура явно принимает raw input. Partial AddressParse может иметь **OK Record** с valid intervals, groups, auxiliary links, invalidTokens и ambiguity diagnostics; извлечение valid intervals не означает, что invalidTokens были успешно проверены. Rule обязано отдельно обработать эти members, иначе validator completeness contract не разрешает выдавать полный PASS.

## 4. Узлы AST и scopes

Каждый узел содержит NodeId, NodeKind, exact type/ref, children в определённом порядке, parameters schema и source provenance. TypeChecker выводит/проверяет тип и все ссылки до публикации.

| NodeKind | Вход / выход | Правило |
| --- | --- | --- |
| TypedLiteral | type + typed payload либо Empty<T> → T envelope | Канонические даты/числа; никаких формул Excel |
| FieldRef | FieldId → EvalResult<T> | Разрешён только в scope шаблона/заявки |
| ColumnRef | alias + ColumnId → EvalResult<T> | Alias обязан существовать в лексическом scope строки |
| ContextRef | CLOCK_DATE / LOCALE / RUN_PROFILE и другие зарегистрированные keys → T | Значение закреплено RunContext, не считывает ОС заново |
| ValueRef | BindingId → EvalResult<T> | Lexical Let/collection binding, scalar либо record, immutable |
| Member | record expression + static MemberId → T | Схема известна, dynamic arbitrary property запрещена |
| Record | SchemaId + static MemberId → child AST expressions → Record<SchemaId> | Конструирует typed record для Map/score/join, каждый member сохраняет свой envelope и requirements |
| ListExpr | ElementType + ordered child AST expressions → List<ElementType> | Конструирует список вычисленных значений, одинаковый T; non-OK обязательного элемента распространяется, не отбрасывается |
| SourceRows | BindingId + schema → List<Record> | Объявленная lazy dependency; source adapter вне pure операции |
| RefExpression | ExpressionId@version + typed args → T | Параметры и return type проверяются, рекурсия запрещена |
| OutputRef | RuleRef + OutputId + InstanceKey → typed output envelope | Только объявленный same-run DAG producer; не предыдущий UI run/counter и не форматированный Message |
| Call | OperationId@exact version + args + typed parameters → T | Только зарегистрированные сигнатуры, никаких Eval/Application.Run из пользовательской строки |
| Let | ordered bindings + body → type(body) | Каждое имя связывается один раз; forward/self-reference запрещены |
| When | Bool gate + then + else → общий тип | Оценивает только выбранную ветвь; non-OK gate не выбирает else |
| AllOf / AnyOf / Not | Bool children → Bool | Короткое замыкание по правилам ниже |
| StateIs | expression + список states → Bool OK | Явно наблюдает envelope; raw diagnostic остаётся в trace |
| Filter / Map | List<T> + alias + predicate/projection → List<T>/List<U> | Alias только внутри дочернего дерева; row provenance сохраняется |
| Any / All / None | collection + alias + Bool predicate → Bool | Пустое множество и ошибки определены ниже |
| Count | List/Set/Group → Int64 | Count non-OK не становится 0 |
| GroupBy | collection + alias + typed key → List<Group<K,T>> | Comparator явный, invalid key не схлопывается в пустой ключ |
| Join | left/right collections + два aliases + typed predicate + join kind → List<Record> | INNER/LEFT; multiple matches сохраняются, если cardinality policy не требует уникальности |

`Exists` — UI-оператор для OK collection Count>0; `IsEmpty` — UI-оператор StateIs(EMPTY) для scalar или Count=0 для collection, с разными явными типами. EQ/NE/LT/LE/GT/GE, IN, CONTAINS, STARTS_WITH, arithmetic UI nodes строятся типизированным Call соответствующей зарегистрированной сигнатуры. Они не добавляют неописанные semantics.

Scope RuleDefinition = Document, Field, Row(TableId), Table, Group(collection/key). Scope — способ инстанцирования, не новый evaluator. Aliases `t`, `a`, `p` существуют только в соответствующих Filter/Map/Join/quantifier; вложенные aliases имеют новые BindingId, одинаковое DisplayName не создаёт захват внешнего значения. NamedExpression не захватывает ambient row: получает row/значения явными аргументами.

Обозначения вроде `p.Allowed` в документах — сокращение Member(ValueRef(P), ALLOWED); `@FIELD` — FieldRef, `t.PORT` — ColumnRef с alias binding. Это не утверждённая лексическая grammar. AST references не хранят экранные подписи или положение в дереве вместо ID.

Alias коллекции scalar читается ValueRef(aliasBindingId), без выдуманной колонки TOKEN. Для record alias ColumnRef(alias,ColumnId) эквивалентен Member(ValueRef(alias),ColumnId). TypedLiteral(List<T>,payload) содержит только literal values; список результатов нескольких Map строится ListExpr(List<T>,children), затем OP_LIST_CONCAT. Call вычисляет аргументы eagerly в установленном порядке и использует заявленный в каталоге приоритет non-OK; только управляющие узлы дают short-circuit.

## 5. Булева логика, вычислительная полнота и коллекции

AllOf слева направо: OK(TRUE) продолжает, OK(FALSE) возвращает FALSE без чтения последующих children, non-OK возвращает это состояние без его преобразования в FALSE. AnyOf: OK(FALSE) продолжает, OK(TRUE) возвращает TRUE, non-OK распространяется. Not инвертирует только OK Bool. Узлы пропущенной ветви получают trace `NOT_EVALUATED`; это trace event, не дополнительный EvalState.

| Выражение | Результат / trace |
| --- | --- |
| AllOf(FALSE, SourceRows недоступного источника → predicate) | OK(FALSE); источник не нужен и не открыт |
| AllOf(TRUE, UNAVAILABLE) | UNAVAILABLE; rule CANNOT_EVALUATE |
| AnyOf(TRUE, INVALID) | OK(TRUE); правая ветвь NOT_EVALUATED |
| AnyOf(FALSE, INVALID) | INVALID, не FALSE и не TRUE |
| Not(EMPTY) | EMPTY; не TRUE |
| When(UNAVAILABLE, X, Y) | UNAVAILABLE, обе ветви NOT_EVALUATED |

Пропущенный optional source не делает run incomplete. Но уже полученная ошибка обязательного вычисления не исчезает из trace/diagnostics после настройки цвета или скрытия пользовательского сообщения. Admin не может настроить общий error policy «источник недоступен → PASS». Явная бизнес-ветвь «нет записи в успешно прочитанном справочнике → разрешить» — нормальная Bool policy, отличная от UNAVAILABLE. Исторические PASS в linked DB/IP из Q-10 представлены как migration gap, не предлагаются default Runtime.

Any(empty)=FALSE, All(empty)=TRUE, None(empty)=TRUE. Для matched-IP ALLROWS нужен явный guard Count(relevant)>0; без него пустая группа логически удовлетворяет All, но не доказывает применимость бизнес-правила. Required table emptiness проверяется отдельной clause. Filter с non-OK predicate не удаляет строку молча: возвращает non-OK с row provenance; Map/Join/GroupBy аналогично. Для обработки невалидных строк надо **явно** предварительно классифицировать состояния, опубликовать diagnostics и создать отдельную валидную подколлекцию. Порядок стабильный по snapshot order/явному Sort; join never silently first-match.

Для resource budgets операции коллекций и интервальный index ограничивают число comparisons/nodes/output rows/depth. Индекс может ускорять Join, но не менять результаты/ordering. Превышение бюджета — LIMIT, не усечённый успешный результат. Huge IpSet не превращается в collection всех адресов ради All: используются intersection/difference/coverage. Подробные сигнатуры/ошибки — [каталог](operation-catalog.md).

## 6. Правило и маршруты результата

Rule gate FALSE → NOT_APPLICABLE, причина объясняется. Gate non-OK → CANNOT_EVALUATE. После TRUE создаются instances scope; Let/clauses читают snapshots и формируют typed findings. Clause ASSERT TRUE = без finding, FALSE = настроенное замечание/согласование; EMIT TRUE = настроенный результат, FALSE = без результата. Контекст EMIT не является автоматическим PASS всей проверки. Независимые clauses допускают несколько issues на одной цели.

StateIs используется в explicit state branches: EMPTY поля может создать пользовательский FAIL; MISSING locator может создать отдельное техническое предупреждение NoCell и incomplete flag, но не подтвердить значение поля. Такой обработанный MISSING не обязан останавливать независимую проверку title; downstream requiring field value остаётся CANNOT_EVALUATE. `handled` означает объяснённую диагностику, а не доказанную корректность данных.

В RuleOutcome сохраняются findings всех уровней, `evaluationCompleteness` и status. При неполной обязательной оценке status=CANNOT_EVALUATE даже при известных FAIL findings; findings не теряются. При полной оценке предлагаемый summary порядок FAIL > NEEDS_APPROVAL > WARN > INFO > PASS, а список approvals/errors сохраняется отдельно; NOT_APPLICABLE не участвует как PASS. Это новый summary-contract для ревью, не изменение существующих severity v1 и не решение Q-05.

Числовой priority влияет только на явно указанную selection/conflict policy. Sequential clauses не дают «последний цвет победил». Найденные условия Rulebook выражаются как Filter source rows по predicate AST; OP_RANK_MAX выбирает максимальные ranks и сохраняет ties; OP_UNIQUE_VALUE выявляет incompatible structured output, а collect-all approvals/comments остаются отдельной проекцией. Deny/exception — такие же выражения с заданным эффектом, не магическое правило имени.

## 7. Показательный структурный фрагмент

Документационный пример без runnable serialization: поле EMAIL читается FieldRef; StateIs(EMPTY) вызывает отдельную required-clause. При OK Email Let(P) связывает SourceRows по Binding; Filter с OP_EMAIL_SELECTOR_SCORE выбирает совпавшие строки; OP_RANK_MAX оставляет максимальный score; Map извлекает ALLOWED и DENIED IpSet; OP_IP_UNION объединяет их отдельно; OP_IP_COVERAGE(request, allow) и OP_IP_OVERLAPS(request, deny) дают признаки нарушения. ASSERT требует отсутствия deny-overlap и полного покрытия при непустом allow. ActionDef отдельно связывает finding с полем IP, сообщением и каналом.

Если выбранных строк нет в успешно прочитанной policy, отдельная When-ветвь реализует принятую для этой policy NoMatchAction. Если SourceRows UNAVAILABLE, вычисление не попадает в ветвь «нет строк». Exact-email-over-domain, scope и no-match настраиваются в дереве/источнике; библиотека не читает лист «Инициаторы» сама. До согласования Q-03 конкретный проект правила остаётся design-only.

## 8. Обязательные проверки проекта до программирования

| Проверка | Ожидание архитектуры |
| --- | --- |
| Field renamed/moved; reused twice | Все ссылки по ID сохранены; новый snapshot адрес, оба target обновлены |
| Enum/Text argument вместо IpSet | Type error до публикации, explicit parse нужен |
| Alias вне вложенного Join | Scope error с NodeId, без fallback на document field |
| Cycle именованных выражений/outputs | Publish blocked с цепочкой цикла |
| Boolean ошибки/EMPTY/пустые коллекции | Контракты таблиц выше, никакой неявной истинности |
| ADR-007 .1–.10, deny .5 | Overlap TRUE, uncovered при allow .1–.8 = .9–.10 |
| Superset имеет только .1 и .10 | MatchedRows имеются, но coverage PARTIAL, uncovered .2–.9 |
| List .1,.10 против .1–.10 | OP_IP_EQUAL FALSE; декларации сравниваются отдельно |
| Very large range | Нормализованные интервалы, без materialize-all; LIMIT только по объявленному бюджету |
| Same merge, разные targets | Canonical target group до Writer; все findings сохранены |

Это ожидаемые design witnesses и негативные проверки будущего валидатора, **не запущенные тесты**. Ни AST interpreter, ни type checker, ни Excel writer в этом PR не реализованы.
