# PR-04 — универсальный scalar Evaluator

База: main `42839a5a709b40dc0559e580c376cc202ad9f314`. Ветка `feat/pr-04-evaluator`. Основание — задание владельца PR-04 и [expression-model](expression-model.md). Реализуется только выполнение уже существующего scalar AST; PR-05 не выполняется.

## API и границы доверия

| API | Точный контракт |
| --- | --- |
| `AtEvalValue(typeName, state, value, dependency="")` | Проверенный input envelope. Известный тип/одно из девяти состояний; OK содержит канонический scalar текст PR-03, non-OK не имеет payload. Неверный host input → VBA error 5, без догадок/преобразований |
| `AtEvalContext(packageId, packageVersion, templateId, fields)` | Immutable snapshot: Dictionary FieldId → CAtEvalResult; копируются envelopes и scalar records. Package/version/template закреплены. ReadField возвращает новую копию, отсутствующий FieldId → Nothing |
| `CAtEvaluator.Evaluate(package, expressionId, context, maxSteps=1024, maxDepth=64, maxDiagnostics=128, maxTrace=2048, maxText=65536)` | Сначала проверка schema2, identity и всей структуры с каталогом runtime; затем scalar AST. Ошибки preflight → ERROR и пустая trace; бюджеты не позволяют запустить непроверенное дерево |
| `CAtOperationRegistry.Catalog` | Новый независимый Dictionary exact ID@version → signature, совместимый с PR-03. Правка возвращённой метаинформации не меняет registry |
| `CAtOperationRegistry.ExecuteOperation(id, version, arguments, parameters)` | Только compiled whitelist и строгие повторные проверки типов/арности/параметров. Неизвестная exact version → ERROR. Нет callbacks, имени макроса/формулы, динамической регистрации из конфигурации |

Результат CAtEvalResult: ResultType, State, Value, HasValue, Diagnostics, Requirements, Provenance, Trace, IsComplete. Публичные getters коллекций возвращают глубокие копии scalar records; наружу нет setters. Friend builders предназначены только для доверенного кода того же VBA-проекта. Внутренние scalar значения хранятся каноническими строками: Bool — `true`/`false`, Int64/Decimal без Double rounding. Не происходит Text→Bool/IP/Date coercion. **Value читается только при HasValue=True (State=OK)**; пустая служебная строка при non-OK обозначает отсутствие payload, а не успешный Text(""). OK Text("") и EMPTY различаются.

Context не принимает предыдущие результаты исполнения с Requirements/Trace и противоречивый OK envelope с Error diagnostics. Исходная диагностика и provenance non-OK input сохраняются. Ограничения конструктора snapshot: ≤2000 fields, ≤65536 символов scalar value, ≤128 diagnostic/provenance records на поле, суммарно ≤1048576 символов value/metadata. Ошибки построения context — нарушения host contract до Evaluate, а не частичный успешный результат. Сам Evaluator не получает Workbook/Range и не создаёт Excel COM: использует VBA records/Collections и существующий Windows Scripting.Dictionary.

## Исполнение и состояния

- TypedLiteral, FieldRef, ValueRef, Let, AllOf, AnyOf, Not, When, StateIs, Call исполнены. FieldRef читает snapshot; отсутствующее поле → MISSING, несовпадение типа → ERROR. Никакого ActiveWorkbook или I/O.
- AllOf/AnyOf идут слева направо, short-circuit только на OK false/true соответственно; первый достигнутый non-OK распространяется без преобразования. Not инвертирует только OK Bool. When исполняет только выбранную ветвь, non-OK gate не выбирает else.
- Let вычисляет ordered bindings один раз, в каждой инициализации видны только внешние и предыдущие bindings. ValueRef возвращает копию уже вычисленного envelope. Effects всех достигнутых bindings сохраняются, включая неиспользованный впоследствии binding. Вложенный scope не протекает в sibling.
- Call eager: все аргументы посещаются в порядке Children. Если есть non-OK аргументы, backend не вызывается; состояние — первого такого аргумента, diagnostics/requirements — **всех** достигнутых аргументов. Скрытых приоритетов состояний нет.
- Все девять состояний PR-03 различаются: OK, EMPTY, MISSING, INVALID, UNAVAILABLE, ERROR, CONFLICT, LIMIT, NOT_APPLICABLE. Только OK имеет payload. StateIs возвращает OK Bool, сохраняя effects ребёнка.

## EvaluationRequirements / completeness

Requirement record: RequirementId, ObjectId, OriginNodeId, Dependency, RequiredBranch, State, Classification, Disposition. RequiredBranch — NodeId достигнутого обязательного наблюдения; путь восстанавливается по AST/trace. Dependency — opaque host dependency, либо локальный AST/Node anchor; missing field указывает Field/FieldId. RequirementId детерминирован внутри одного вычисления, записи deduplicate при распространении и повторном ValueRef.

На фактически прочитанном non-OK поле создаётся UNRESOLVED requirement. MISSING/UNAVAILABLE/ERROR/CONFLICT/LIMIT — Technical; EMPTY/INVALID — Input; NOT_APPLICABLE — Applicability. Собственная ошибка Call создаёт отдельный requirement даже рядом с уже унаследованной ошибкой. Propagation/StateIs не создают заменяющий requirement и не удаляют исходный. Бюджет добавляет Technical requirement с Dependency=Budget.

**IsComplete = OK + отсутствие unresolved requirements + отсутствие Error diagnostics**. Это не RuleOutcome/PASS и не механизм публикации. Например, StateIs(required field UNAVAILABLE) → OK(true), но IsComplete=False, raw diagnostic и requirement сохраняются. When/Let/скрытие результата не делают его успешным. Правка возвращённой коллекции не может снять требование. Для недостигнутого по short-circuit поля requirement не создаётся.

TypedLiteral(EMPTY) — намеренная константа, не обязательный внешний вход; она не порождает внешнее requirement, в том числе через Let/ValueRef/strict propagation. EMPTY **прочитанного поля** остаётся незавершённым Input requirement. Это не разрешение подтвердить плохой IP или скрыть source error.

PR-04 не реализует RuleOutcome, findings и разрешённую domain-классификацию Input requirements. Все достигнутые non-OK входы консервативно остаются UNRESOLVED; публичного `resolve technical error`/optional switch нет. Будущий классификатор правил может завершать согласованные EMPTY/INVALID cases по expression-model, но не технические обязательства ради PASS.

## Четыре операции

| Exact key | Сигнатура и политика |
| --- | --- |
| OP_TEXT_EQUAL@1.0.0 | (Text, Text) → Bool, обязательный Text parameter MODE=`binary`; StrComp vbBinaryCompare, без Trim/casefold. Другое значение MODE → INVALID |
| OP_TEXT_TRIM@1.0.0 | (Text) → Text; только крайние ASCII spaces U+0020, без скрытой Unicode-нормализации |
| OP_BOOL_EQUAL@1.0.0 | (Bool, Bool) → Bool; равенство канонических Bool |
| OP_INT64_LE@1.0.0 | (Int64, Int64) → Bool; точное signed сравнение строк по знаку/длине/digits, весь диапазон Int64 |

Новые compiled операции добавляются отдельной разработкой: одновременно whitelist, signature и тесты; конфигурация выбирает только exact ID/version и typed parameters. Старые OP_TEST* fixtures не становятся автоматически executable: runtime-тесты явно связывают их с реальными четырьмя операциями только в тестовой модели. Бизнес-операций NETWORK_ACCESS нет.

## Ресурсы, trace и совместимость

Параметры budget могут только уменьшать указанные hard caps. Steps считают входы в узлы и backend calls; skipped subtree не расходует вычисления, но расходует trace. Depth — глубина EvalNode. Trace records: Sequence, ObjectId, NodeId, Kind, Event, Reason; EXECUTED / OPERATION / NOT_EVALUATED показывают реальный порядок и причину пропуска, без текстовых payload. Переполнение steps/depth/diagnostics/trace/scalar payload → LIMIT, HasValue=False, IsComplete=False. При trace cap оставшиеся записи опускаются с явной E_LIMIT причиной; trace не выдаётся за полную. При diagnostic cap сохраняется bounded summary, **requirements не отбрасываются**. Не происходит возврата усечённого успешного значения.

Новая SchemaVersion не нужна: PR-04 не меняет XML/AST форму. Все регрессии read/write схем 1.0.0 и 2.0.0 сохранены; legacy deferred не исполняется. No implicit upgrade. Runtime preflight требует только schema2 с реальными зарегистрированными сигнатурами.

## Контракт для следующего этапа

Будущий Resolver отдельно и безопасно читает Excel, нормализует типы и формирует новый snapshot для каждого запуска. Dictionary keys — FieldId; значения — envelopes с точной declared type и состоянием. Он различает EMPTY cell / MISSING locator / INVALID input / UNAVAILABLE resource / CONFLICT selection; не заменяет их успешной пустой строкой. PackageId/Version и TemplateId должны совпасть с определениями. Dependency/provenance должны содержать неприватные opaque snapshot/source revision IDs, не исходные секретные значения/пути. Результаты разных context/run не объединять без RunId namespace снаружи. Сейчас OK input/literal codec ограничен scalar Text/Bool/Int64/Decimal/Date/DateTime PR-03; OK domain payloads и табличные scopes требуют отдельного контракта и сейчас отклоняются. Известные non-OK domain states можно передать без payload.

Host вызывает Evaluate синхронно, читает State/HasValue и IsComplete/Requirements, затем отдельно формирует findings/ActionPlan. Не использовать один Value или Bool как PASS. Registry не открывает источники, Evaluator не вызывает Writer. В этом PR нет Resolver, source adapters, табличного Evaluator, Writer, Admin UI, Rulebook/Superset и PR-05.

## Проверка

Реальный Excel 16.0 build 17932 Office64 / Windows x64: `PASS|96|0||Office64`, `PASS|AST|76|0`, `PASS|Evaluator|84|0` — **256 assertions PASS, 0 FAIL**. Проверены все 172 старые регрессии, исполнение двух исходных examples и одинакового AST для двух разных TemplateId, TRUE/FALSE, все non-OK states, short-circuit/eager order, Let once/nesting, missing field, wrong types/operations/parameters, StateIs requirements, собственная ошибка Call рядом с inherited requirement (RED → GREEN), все runtime caps, immutable snapshots/records, Int64 boundaries и отсутствие побочных изменений тестовой книги. Независимое source review помогло выявить последний requirement-origin gap; до исправления тест был отрицательным. Office x86 VBA и ручной UI не проверены; прохождение тестов не означает ручную приёмку XLAM.

Воспроизведение: `./tests/Run-Tests.ps1` по [условиям PR-02](pr02-validation.md), дополнительно `./tests/Probe-MSXML.ps1`. Runner создаёт/закрывает только собственный скрытый Excel, не меняет Trust Center. Result/host/hashes/trace assertion evidence находятся в игнорируемом artifacts/run-…; добавлены eval-result.txt и eval-assertions.txt. Никаких реальных заявок/Office binaries в PR. AccessTools 1, архив и Google Drive не изменены. Бизнес-семантика не менялась.
