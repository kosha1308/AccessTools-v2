# Admin и Runtime: графические рабочие процессы

Статус: **предлагаемый дизайн PR-01, только документация**. Формы, кнопки, AST, журнал Writer и публикация ниже не реализованы; ручная приёмка не выполнена. Основания: [архитектура](architecture.md), [ADR-001…007](decisions.md), [критерии](acceptance-criteria.md), [аудит и Q-01…Q-11](v1-audit-findings.md), [Excel-контракт v1](v1-excel-output-contract.md), [ТЗ v0.15](https://docs.google.com/document/d/1skCtOzRTq4_wuNNkZHSaeyWsh_sf_RjYq1OKDezmBKM/edit), разделы 2–6, 8–10. Названия экранов проектные; технология Excel UI и сериализация этим документом не выбираются.

## 1. Навигация и идентичность

Старт Admin: «Шаблоны», «Общие источники», «Библиотека операций», «Пакеты и версии», «Проверка и диагностика». Карточка шаблона: «Распознавание», «Поля и таблицы», «Именованные выражения», «Правила», «Действия и сообщения», «Зависимости», «Проверить», «Версии». Хлебные крошки показывают объект и draft version; справа свойства, использующие объекты и ошибки с переходом к узлу. Условия редактируются деревом, выбор Excel-области — отдельным временным режимом с возвратом в карточку.

Все изменения проходят через UI: без VBA, ручной правки ini, служебных формул или внутренних файлов пакета. Возможности ограничены структурными узлами и зарегистрированной библиотекой; новый произвольный алгоритм пользователь не программирует.

TemplateId, FieldId, TableId, ColumnId, SourceId, RuleId, ActionId и идентификаторы именованных выражений — стабильные неизменяемые uppercase ключи. UI проверяет уникальность и лексический контракт схемы. DisplayName изменяем, может быть русским; выбор показывает также тип/шаблон для устранения неоднозначности. Исторический `NET.*` хранится как `LegacyRuleId`. Копия получает новый ID; перенос поля/переименование label сохраняют ID. История опубликованных версий не переписывается.

## 2. Общий жизненный цикл

| Действие | Шаги UI | Проверки и результат |
|---|---|---|
| Создать | «Создать» → тип → имя → свойства → сохранить draft | Новый ID, незавершённость в диагностике |
| Клонировать | Объект/версия → «Копия» → имя/область → зависимости | Новые собственные ID; внутренние refs клона согласованно перепривязаны; shared refs показаны явно |
| Редактировать | Published object → «Черновик новой версии» → правки → diff | Runtime продолжает прежний immutable package |
| Отключить | Impact preview → состояние в draft | Не равно удалить; активная неразрешимая зависимость блокирует publish |
| Архивировать | «Архивировать» → consumers → сохранить | Объект скрыт из новых выборов; исторические пакеты/результаты сохранены |
| Восстановить | «Архив» → «Восстановить в draft» | Тот же ID; коллизии и несовместимые зависимости требуют разрешения |
| Удалить draft object | «Используется в…» → заменить/удалить refs → удалить | Нет молчаливого cascade; published objects архивируются |

Цикл применяется к шаблонам, правилам, полям/колонкам, источникам и действиям. Отключение шаблона исключает новые запуски из новой публикации; старые выпущенные файлы не изменяет. Отключение rule исключает его вычисления/действия. Для field/source/action отображаются все consumers; их нужно перенастроить/отключить, а не трактовать missing dependency как пустоту. Shared source имеет собственную версию; пакеты принимают новую binding/schema явно.

Копирование шаблона предлагает сохранить общие источники по ссылке либо создать новые definitions. Пути реальных файлов и fixture-данные не распространяются автоматически. Копирование rule в другой шаблон требует mapping каждого field/table/source/action; одинаковые DisplayName не доказывают идентичность. Архивирование action не очищает уже отмеченные книги без journal reconciliation.

## 3. Шаблон, поле и выбор реального диапазона Excel

1. «Шаблоны → Создать»: название, TemplateId, описание; выбрать разрешённый образец файлом. Показать полный путь, книгу и листы. Чтение образца не меняет его.
2. «Распознавание»: выбрать реальные ячейки признаков, ожидаемые метки/структуру и признаки версии. Preview показывает неоднозначность между шаблонами; имя файла не единственное доказательство. На старте Excel все источники не перебираются.
3. «Поля → Добавить»: название, тип, scalar/table column → «Выбрать в Excel». Admin сворачивается в панель выбора, пользователь активирует нужную книгу/лист и выделяет область. «Принять» возвращает selection; «Отмена» не меняет draft.
4. Подтверждение показывает **книгу, sheet identity, видимое имя листа, A1 адрес, размер, MergeArea/anchor и превью**. Переключение книги обнаруживается; другая книга не принимается молча. Runtime привязывает лист внутри распознанного экземпляра, а не к временному пути образца/ActiveSheet. Переименование/замена листа проверяется resolver-ом; неоднозначность — ошибка сопоставления.
5. Scalar внутри merge нормализуется к anchor и полной MergeArea, оба показаны. Частичное пересечение merge требует исправить selection либо явно расширить с повторным подтверждением. Смешанный диапазон не превращается в один scalar. Разорванные selection допускаются только для модели списка целей, иначе нужны отдельные bindings.
6. Для table выбрать header row, ColumnId, начало, правило окончания/включения и logical row identity. Полностью/частично merged строки, пустые промежутки и формулы с пустым результатом видны в preview. «До первой пустой» не скрытый universal default. Неоднозначный merge требует явной схемы logical rows; каждая физическая строка не объявляется бизнес-записью автоматически.
7. «Проверить привязки» показывает распознанные поля/строки и EMPTY/MISSING/INVALID с причинами. Relative offsets, dynamic selectors и column mappings редактируются тем же мастером с проверкой типов.

### Перенос без переписывания правил

`EMPLOYEE_EMAIL`, DisplayName «Почта сотрудника», расположен D7. Два rules используют FieldRef(EMPLOYEE_EMAIL); action — TargetFieldRef(EMPLOYEE_EMAIL). Admin открывает поле → «Расположение» → выбирает F9:G9 → видит merge anchor F9 → просматривает три использования → сохраняет draft. Меняется binding, не refs. После публикации оба rules и action разрешают F9:G9; поиск/замена D7 в текстах не нужен. Preview показывает старую/новую цель. Старую отметку D7 очищают по journal её run, а не по новому binding.

Переименование в «Корпоративная почта» меняет label. Retype Text→Date показывает consumers и несовместимые операции/подстановки; неявное приведение до PASS запрещено. Invalid draft сохраняется, publish блокируется до явного преобразования/правок consumers. Удаление field/column показывает ссылки из rules, recognizers, DAG expressions, bindings, actions и reports, включая неактивные drafts; пользователь заменяет ссылки либо архивирует потребителей. Dangling refs запрещены в опубликованном пакете.

## 4. Общие источники

«Общие источники → Создать»: SourceId, название, поддерживаемый тип, locator. Выбор файла → лист/таблица → header row → key/result columns и типы. Preview различает missing file, read failure, schema error, duplicate key, EMPTY и INVALID value. Настраиваются schema/version, freshness, duplicate policy и поддерживаемый cache contract. Новая политика allow/deny при ошибке источника не придумывается.

В template выбирается общий source и mapping logical fields к schema columns. Переименование файла/листа/колонки выполняется здесь с повторным preview; «Используется» показывает затронутые пакеты. Изменение definition не правит сам источник. Совместимые bindings/schema/cache позволяют нескольким шаблонам использовать один индекс.

Preview фиксирует snapshot/version/freshness и provenance. Несохранённая открытая книга и диск не смешиваются молча: фактически прочитанный snapshot показан, несовместимость даёт UNAVAILABLE/CONFLICT по контракту. Runtime lazy-читает только зависимости фактически исполняемых правил роли, не весь каталог.

## 5. Графический редактор правила

«Правила → Создать»: название, бизнес-цель, scope Document/Field/Row/Table/Group, execution roles, входы, condition tree, outcome mapping, исключения, приоритет, actions. Название и порядок карточек не являются скрытыми условиями/приоритетами.

Авторитетное представление — **typed structural AST**. Кнопка «Добавить» предлагает допустимые узлы `Let`, `When`, `AllOf`, `AnyOf`, `Not`, `StateIs`, `Filter`, `Map`, `GroupBy`, `Join`, `Any`, `All`, `None`, `Count`, `Call OP_*@version`, refs и typed literals. Текстовый preview опционален и **только для чтения**. Произвольные Eval, VBA и исполняемый пользовательский текст запрещены.

| Редактирование | Действие администратора | Проверка |
|---|---|---|
| Группы | Добавить AllOf/AnyOf/Not, переместить детей | Boolean types и scopes; trace отражает порядок |
| Условие | Выбрать field/column, оператор и typed operand/другое поле | Тип/cardinality; missing не приводится к пустой строке |
| Повторное выражение | Let/именованное выражение, refs входов | DAG, cycles, неизвестные refs; label rename сохраняет identity |
| Таблицы | Filter → Any/All/None/Count; Map/GroupBy | Row aliases, scope, empty-set policy |
| Корреляция | Join по явно выбранным keys/операции | Не случайный Cartesian join; multiplicity/duplicates/missing и budget явны |
| Сложный алгоритм | Выбрать OP_*@version из библиотеки, назначить параметры | Версия/сигнатура/статусы/лимиты; нет создания arbitrary алгоритма |
| Исключение | Условие + область + основание + результат | Видно в trace; не скрывает техническую ошибку |
| Приоритет | Выбрать согласованную conflict policy | Tie не first-row-wins: композиция либо CONFLICT |
| Выход | Outcome и несколько actions | Цели/severity/messages/channels отдельно от вычисления |

Исключение — видимый узел/ссылка, а не скрытая VBA-ветка. NOT_APPLICABLE или другое преобразование требует явного контракта; исключение не превращает UNAVAILABLE/ERROR в PASS. Подавленные совпадения объяснимы. Domain/email policy, hostname specificity, deny/allow, ALLROWS и declaration comparison не получают новую семантику только от вида UI.

Join сохраняет provenance левой/правой строк. «IP и сервис этой же строки» ссылается на один alias/явный join, не на любой сервис книги. Поведение All на пустом наборе, Count при частично недоступных строках, Not над ошибкой явно показано в контракте и preview. Неразрешённая обязательная семантика блокирует publish. StateIs различает состояния без boolean coercion. Перестановка узлов не разрешает менять бизнес-политику скрыто.

EvalResult: typed value, state **OK | EMPTY | MISSING | INVALID | UNAVAILABLE | ERROR | CONFLICT | LIMIT | NOT_APPLICABLE**, provenance, trace. Rule outcomes: **PASS | FAIL | WARN | INFO | NEEDS_APPROVAL | NOT_APPLICABLE | CANNOT_EVALUATE**. Это разные измерения; severity отдельно. Mapping задаётся rule contract. Source read failure, required missing, conflict и budget exhaustion не PASS; «источник доступен, записи нет» отличается от «источник не прочитан».

Узел Explain показывает ожидаемый тип, фактическое состояние, причину, field/source/operation и physical/logical row identity. Short-circuit и «не вычислялось» не изображают успешный дочерний результат. Необработанная ошибка операции даёт ERROR/CANNOT_EVALUATE с кодом, не произвольный fallback. Библиотека показывает version/signature/states/limits/examples. Запрос нового алгоритма передаётся разработчику; до регистрации версии rule остаётся draft. Общие операции используются без проверки имени шаблона внутри evaluator.

## 6. Draft → preview → publish → rollback

«Сохранить черновик» допускает незавершённый/invalid AST, который можно однозначно загрузить, с сохранённой диагностикой. «Проверить» валидирует schema/types/refs/cycles/scopes/versions/budgets, затем sandbox использует **копию snapshot** книги и источников. Preview не меняет заявку/source/Rulebook/published package. Без безопасной копии он прекращается с причиной.

Explain показывает raw/normalized values, узлы/aliases, исключения, versions, states, provenance, план каждого канала; переход к полю и visual plan работают на копии. Отличаются «запланировано», «применено в preview» и Runtime output. Sandbox success не подтверждает Excel Writer manual.

«Опубликовать» показывает diff объектов, bindings, roles, sources, operations, texts/actions и intentional changes. Блокируют errors, dangling/disabled dependencies, unsupported versions, conflicts, необходимые пакету unresolved semantic choices и отсутствие обязательной согласованной приёмки. Warning не снимается молча; ограничения фиксируются. Отсутствие schema errors не закрывает Q автоматически.

Создаётся immutable package с versions/schema/dependency fingerprints. Активация только после полной проверки; при сбое старый пакет активен. Runtime pin-ит версию на run. Формат storage/signatures и механизм атомарного переключения требуют отдельного контракта; Excel-транзакция не обещается. Draft невидим Runtime.

«Версии → Сравнить → Откатить» выбирает прежний совместимый пакет, показывает diff и проверяет зависимости. Откат переключает active package, не переписывает историю/источники и не стирает чужие правки annotations. Старые marks обслуживает journal своего run. Disable/archive не означают немедленного удаления marks.

## 7. Роли и каналы независимы

ADMIN редактирует; FIRST_LINE/COORDINATOR исполняют capabilities. Rule выбирает execution roles, action — audience/channel: заявка, окно, пользовательский готовый текст/отчёт, coordinator output, result workbook, скрытый diagnostic trace. Severity не назначает аудиторию. NEEDS_APPROVAL может показываться в поле и отсутствовать в тексте исправлений только как явная настройка.

FIRST_LINE не исполняет coordinator-only rules и не читает их sources. Общий rule может работать в обеих ролях с разными каналами. Скрытие в окне не удаляет canonical issue; исключение из user report не исключает coordinator. Техническая невозможность проверки доступна оператору и не маскируется отсутствием бизнес-ошибок. Отключённый rule не имеет скрытого legacy execution path.

## 8. Issue, Action Planner, Writer и восстановление

Это **целевой проект**, не гарантия legacy markers. Issue: identity, RunId/PackageVersion/RuleId, outcome/severity, channel messages, logical targets, provenance, audience. ActionDefinition: вид, FieldRef/ColumnRef/row alias либо layout target, текст/стиль/маршрут. Evaluator только выдаёт данные. Planner разрешает цели и строит проверяемый план до записи.

### Canonical physical target и конфликт

Перед группировкой разрешаются workbook instance → sheet identity → **полная физическая MergeArea/anchor** либо обычная область. Разные logical refs одной merge объединяются. Пересечения немerged ranges разбиваются на однозначные участки либо дают CONFLICT; результат не зависит от dictionary order. Layout/merge/address ambiguity не приводит к записи в угаданную ячейку.

Все самостоятельные сообщения сохраняются. Только согласованная semantic dedup с provenance всех исходных issues может объединять повторы. Цвет ERROR > WARNING > INFO; outcome получает severity через action contract. Порядок: display priority → RuleId → issue identity → stable row identity. Это presentation order, не allow/deny priority. Некомпозируемые actions дают CONFLICT, не исчезновение одного результата.

### Владение и сохранность

Preflight проверяет версию Excel, protection, journal, comment type и стиль/target conflicts. **Защищённые цели и неподдерживаемые threaded comments, обнаруженные после вычисления, дают ActionReceipt=FAILED с diagnostic CAPABILITY_UNSUPPORTED и output incomplete; destructive fallback запрещён.** RuleOutcome от ошибки применения не меняется: FAIL мог вычислиться, но annotation не примениться. Если вычисление target/payload вернуло non-OK EvalResult, зависимый action не допускается к планированию; сохраняются diagnostic и output incomplete. Run не сообщает полную успешную запись.

Ownership journal хранит RunId, package/action/issue identities, fingerprint книги/листа/target, исходный снимок только затрагиваемого оформления/note/metadata, применённое состояние, expected-after и step status. Формат/долговечное storage/защита/size limits требуют прототипа. Маркер `AccessTools.` сам по себе владение не доказывает. Персональные данные не дублируются без необходимости.

Пользовательские note/style/data/formulas сохраняются, слой надстройки отделён. Проверяются rich formatting, CF priorities, shape metadata, merged patterned fills. До первой мутации capability проверяет возможность полного захвата и сериализации затрагиваемых пользовательских note/style, включая нужные свойства восстановления. Если точное безопасное сохранение не поддержано, Writer отказывается от этой записи с причиной; destructive partial fallback запрещён. Полная обратимость всех объектов Excel заранее не обещается.

При clear/repeat текущий объект сравнивается с expected-after. Правка пользователем note/style/value/formula/merge после run не затирается: CONFLICT, адрес, ручной разбор/копия. Старые targets берутся из journal, не текущего binding. Disable rule не даёт права удалить неподтверждённые объекты.

### Частичный вывод и crash

Стадии: полный plan → preflight → подготовленный journal → перед шагом повтор fingerprint → apply → verify expected-after → записать step status → итог. Между apply/ack возможен crash. Это **защищённая поэтапная запись, не ACID-транзакция Excel**.

Ошибка останавливает зависимую цепочку. Состояния journal lifecycle INTENT/PENDING/APPLIED/RECOVERY_REQUIRED видны отдельно от terminal ActionReceipt APPLIED/SKIPPED/CONFLICT/FAILED/ROLLED_BACK; partial output маркируется неполным. Компенсация допустима только при совпадении expected-after и тоже журналируется. Application flags восстанавливаются доступным exception handler; после crash нужна recovery диагностика.

При следующем открытии незавершённый run показывает список шагов, предлагает безопасное восстановление подтверждённых собственных изменений либо копию для разбора. До сверки journal запись поверх unknown state блокируется. Отсутствующий/повреждённый journal не заменяется похожим цветом/marker. Journal и durable result не удаляются как ephemeral preview. Recovery должен быть доказан прототипом/тестами/Excel manual.

## 9. Отчёты и result workbook

User text, окно, notes и coordinator instructions имеют отдельные templates с типизированными подстановками; missing placeholders видны в preview. SKIPPED/недоступная обязательная проверка запрещает ложный итог «проверено, замечаний нет». Полнота вычисления и business issues отражаются независимо.

Будущий NETWORK_ACCESS report profile воспроизводит [зафиксированную модель v1](v1-excel-output-contract.md): «Сводка координатора» + фактические отфильтрованные строки «Superset», связь «№ из сводки», несколько номеров при необходимости. Планируемые 11 полей: №; IP / сеть из заявки; Статус Superset; Найдено строк; Сервис; Инструкция; Имя хоста (Superset); Сервис (Superset); Полное имя ИС (Superset); СКУФ ГРУППА ТП (Superset); КОМАНДА ПОДДЕРЖКИ (Superset). Первые шесть отличаются от source-derived последних пяти. Неоднозначный hostname, not-found, имена/порядок требуют Q-06/эталонов. Source не изменяется; результат строится в отдельной snapshot-копии. Generic Writer не проверяет TemplateId=NETWORK_ACCESS.

**Q-09 остаётся открыт:** глобальные coordinator messages сохраняются в canonical model, но конечный канал/лист/секция не выбирается здесь. Preview показывает «Глобальные сообщения: канал не согласован». Пакет, обязанный их публиковать, блокируется до решения. Не разрешены ни повторение потери «Общее», ни самовольный новый пользовательский лист. Сохранность пользовательской заявки и полнота публикации — разные инварианты.

## 10. Runtime: знакомый цикл

1. Открыть заявку → «Проверить» в доступной роли. Показать распознанные template/version/package. Unknown/ambiguous форма не получает случайный профиль.
2. «Источники»: имена, freshness, schema/version/status. Разрешённую выгрузку можно выбрать через UI, бизнес-правила недоступны Runtime. Read failure — CANNOT_EVALUATE зависимых checks, не PASS.
3. Run: чтение/evaluation/plan → Writer preflight/apply; защита от параллельного repeat, lazy sources. Показать готовность каналов и partial failures.
4. Results: issues и отдельная полнота выполнения; details/trace/provenance в объёме роли. «Перейти» активирует конкретную книгу/лист/canonical target, независимо от ActiveWorkbook. Closed/changed book сообщает причину.
5. «Все проблемы/Копировать» строит текст аудитории. «Открыть результат» открывает временный view последнего run; видна пометка временности, source неизменён.
6. «Повторить» использует именно книгу окна, новый RunId и явно показанные package/source versions. Старые marks reconcile-ятся по journal; user edits — conflicts, прошлый snapshot не смешивается с новым.
7. «Очистить отметки» показывает собственные actions, восстанавливает подтверждённые и сообщает conflicts. Не означает повторную проверку/исправление issues. Reports/durable copies не удаляет без отдельного выбранного действия.
8. «Сохранить копию результата»: путь/overwrite → отдельная долговечная книга/отчёт с run/package/source metadata → подтверждение фактической записи. Сохранить исходную заявку с annotations — отдельная операция Excel. Ephemeral view может очищаться только по объявленной политике; durable copy/journal не считаются view. Save failure — «не сохранено», путь не доказательство успеха.

## 11. Независимый synthetic EQUIPMENT_RETURN

Это **дизайнерский пример**, не действующая бизнес-конфигурация/политика. Реальный UI и Excel ещё не выполнены.

1. Создать `EQUIPMENT_RETURN` («Возврат оборудования»), выбрать искусственную книгу/лист «Возврат» и признаки распознавания. Через Excel selection добавить `EMPLOYEE_ID` Text, `RETURN_DATE` Date, table `RETURN_ITEMS`: `ASSET_ID` Text, `QUANTITY` Number, `CONDITION` Text. Адреса независимы от IDs.
2. Создать synthetic shared `ASSET_REGISTRY`: asset key, employee key, количество/состояние демонстрационной схемы. Preview duplicate key и missing file — разные состояния. Связать источник; lookup operation пригодна другим формам.
3. Создать демонстрационную обязательность EMPLOYEE_ID и проверку типа RETURN_DATE через StateIs. Настроить EMPTY/INVALID outcomes и actions по field refs; это не реальное кадровое требование.
4. Row-rule alias ITEM: synthetic quantity predicate и membership CONDITION. Выбрать общую lookup `OP_*@version`, назначить ITEM.ASSET_ID; Join asset/employee с явными aliases. Not-found и read failure раздельны. Новый EquipmentReturnController не нужен.
5. Filter включаемых строк, All/Any/None/Count, явная empty-set semantics; именованное DAG expression используется двумя rules. Демонстрационное исключение synthetic category имеет основание и trace.
6. Actions: annotation ITEM.QUANTITY и user text; отдельная coordinator instruction. Execution roles и audience выбираются независимо. Два issues одной MergeArea сохраняются оба, цвет детерминирован. Неподдерживаемый threaded target после evaluation даёт ActionReceipt=FAILED с CAPABILITY_UNSUPPORTED и output incomplete, сохраняя RuleOutcome.
7. Перенести RETURN_DATE через selection master, изменить label: оба зависимых refs сохраняются. Удаление ASSET_ID показывает join/lookup/action consumers; retype QUANTITY показывает type errors. Clone rule, disable original, archive/restore clone → diff и отсутствие выключенной ветки.
8. Сохранить invalid draft, показать блокировку publish; исправить errors/required semantic choices; preview копии и согласованный publish gate. Runtime run → navigate → repeat → clear → durable copy; сравнить versions/rollback, draft невидим Runtime. Всё перечисленное — будущая приёмка, **NOT STARTED**.

## 12. Открытые решения, проверка и риски

Q-01…Q-11 остаются открытыми: fixtures Q-02; policy Q-03/Q-10; Writer Q-04; audience Q-05; summary Q-06/Q-09; полнота Q-08; port lookup Q-11. ADR-007 утверждает **полный включительный IPv4-интервал**, но не закрывает весь Q-07: TOP↔LOWER declaration equality и compact grammar требуют отдельных контрактов. Address-set intersection/coverage и declaration comparison — разные типизированные операции. Список крайних адресов не равен диапазону; внутренний deny обнаруживается без обязательного полного перечисления.

При документальной приёмке проверить AC-01…AC-10/AC-16: CRUD/архивирование каждого объекта; перенос field один раз; nested/correlated/exception tree; source error против not-found; все каналы/plan; publish gates/rollback. Excel cases — [PR-00 Writer cases](v1-excel-output-contract.md) плюс user edit after run, orphan journal, crash write/ack, protected/threaded targets, partial result book/Save failure. Проверять отдельно сохранность чужих объектов и полноту coordinator публикации.

**Проверено:** документальное сопоставление с источниками. **Не проверено:** UI, Excel, Writer/recovery, publish/rollback, future tests и performance. Риски: rich Excel formats/comments; durable journal при crash; sheet/row identity; scopes/empty-set semantics; согласование Q-09/Q-10; минимальная Office version и performance limits. Исполняемые компоненты/бизнес-правила этим документом не создаются. Реализация требует утверждённых контрактов и отдельного задания.
