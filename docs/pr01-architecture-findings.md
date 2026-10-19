# PR-01: архитектурное ревью, границы и решения

Статус: **предложения**, не разрешение реализации. Ветка создаётся от main `2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3` после объединения [PR-00 #32](https://github.com/kosha1308/AccessTools-v2/pull/32). Утверждённые ADR-001…ADR-007 сохраняются. Исходники AccessTools 1, архив пилота, рабочие заявки и Google Drive не изменяются.

## 1. Цель и входы

Нужна модель, в которой администратор без VBA создаёт и полностью меняет условия и действия произвольного поддерживаемого Excel-шаблона. Доказательство — конкретное представление 17 RuleId и 26 дополнительных путей PR-00 плюс независимый EQUIPMENT_RETURN. Документирование модели не создаёт реальных RuleDefinition и не является настройкой производственных проверок.

Обязательные материалы прочитаны: AGENTS, README, architecture, decisions, implementation-plan, acceptance-criteria и шесть документов PR-00. [ТЗ v0.15](https://docs.google.com/document/d/1skCtOzRTq4_wuNNkZHSaeyWsh_sf_RjYq1OKDezmBKM/edit) доступно через Google Drive: 141 параграф, revision `AHj4eMSB7bvLUA4xgcgrlYmj87dzfQzI0NJMY34eIxkyMHXFok8RtGNMhmamIyEGxt5njvuqQcBF5De3O2aMAEzVlzYao0zNtWn614YQ_kI`. Это идентификатор прочитанного снимка, не публичный immutable URL. Старые ТЗ и [выводы PR-00](v1-audit-findings.md) остаются историческими свидетельствами; их текст не переписывается задним числом под новую модель.

Приоритет ADR-007: дефисный диапазон включает все адреса; запрет внутри, пробел покрытия и отсутствие промежуточных адресов в источнике нельзя проверять только по endpoints. Эта часть исторического Q-07 уточнена владельцем; вопрос сопоставления CIDR-деклараций и compact записи остаётся открытым. Новый документ не называет весь Q-07 решённым.

## 2. Состав предложения и навигация

| Документ | Что фиксирует |
| --- | --- |
| [rule-model](rule-model.md) | Сущности, поля/locators, источники, версии, граф ссылок, Rule/Issue/Action/Report и журнал Writer |
| [expression-model](expression-model.md) | Типы, EvalResult, AST/scopes, ошибки и short-circuit, квантификаторы, rule outcomes |
| [admin-workflows](admin-workflows.md) | Редактирование всего через UI, выбор диапазона в Excel, test/publish/rollback, familiar Runtime и безопасный вывод |
| [expressiveness-matrix](expressiveness-matrix.md) | 43 отдельных схемы, входы/условия/операции/действия/параметры/свидетельства и второй шаблон |
| [operation-catalog](operation-catalog.md) | Предлагаемые OperationId/version, типизированные контракты, ошибки/параметры, связь с C01…C21 |
| Этот документ | Альтернативы, ограничения, перенос открытых Q и гейты перед кодом |

## 3. Рассмотренные альтернативы

| Область | Рекомендуемое предложение | Альтернативы и цена |
| --- | --- | --- |
| Условия | Структурный typed AST; UI редактирует его напрямую, text preview read-only | Text-first DSL даёт удобный экспорт, но требует grammar/parser/roundtrip и второго редактора; плоские rows усложняют вложенные/correlated scopes |
| Хранение | Логическая схема независима; выбрать сериализацию до кодирования валидатора | JSON удобен для вложенных AST, XML имеет штатные инструменты Windows, Excel tables удобны для аудита, но требуют explicit node/edge tables. Ни формат, ни доступность runtime parser здесь не утверждены |
| Числовые/IP данные | Нормализованные интервалы без перечисления всех адресов; DeclarationSet отдельно | Массив всех IPv4 не ограничен разумным бюджетом; единственный строковый Equals смешивает policy/lookup/declarations |
| Источники | Shared catalog + Binding + immutable run snapshot + общий versioned cache | Перезагрузка файла каждым правилом противоречит lazy reuse; live Workbook внутри операции нарушает детерминизм и разделение вычисления/IO |
| Вывод | ActionPlan + journal + compare-before-restore + receipts, canonical MergeArea | Старые marker-only notes/цвет не доказывают владение; «всегда всё очищать перед проверкой» теряет старый результат при сбое |
| Ошибки | Состояние вычисления отдельно от business outcome и output receipt | Превращение ошибки в False/0/empty/PASS скрывает невозможность достоверного утверждения |
| Новые алгоритмы | Registered pure operation c параметрами; правила — AST | Специальный dispatcher NETWORK_ACCESS или исходный VBA в expression нарушают constructor-first |

Это предложения для согласования, а не новые принятые ADR. Сопоставление legacy-поведения и целевых ограничений не маскируется словом «совместимо».

## 4. Открытые вопросы PR-00: состояние после PR-01

| Вопрос | Что позволяет модель | Что всё ещё требуется от владельца |
| --- | --- | --- |
| Q-01 baseline | LegacyRuleId/Path + source refs, pinned package semantics; разные варианты представимы | Выбрать эталон конфликтных сценариев; слияние PR-00 не приняло `58d8789` автоматически |
| Q-02 рабочие источники | Версионированная schema/Binding, provenance, обезличенные golden vectors | Фактический Rulebook/схема/версия, сообщения и допустимые данные приёмки |
| Q-03 IP/DB policy | Приоритет selector, union/difference/coverage, primary gate, account predicate редактируются | Подтвердить спорные linked/DB трактовки; не устанавливаем новую default политику от имени владельца |
| Q-04 Writer | Нормализация merge, ownership journal, collision-safe planning | Принять физический прототип и ограничения Office; документация не доказывает сохранность в Excel |
| Q-05 аудитории/согласования | ExecutionProfile, Audience и Channel отдельны; messages/severity задаются данными | Утвердить конкретные маршруты; общий workbook note не является приватным каналом |
| Q-06 Superset/сводка | Проекции сохраняют found/not-found, metadata, row links и отдельную coverage | Утвердить labels, порядок 11 колонок, показ PARTIAL/неполного охвата и неоднозначного hostname |
| Q-07 IP-контексты | IpSet vs DeclarationSet, независимая comparison policy | ADR-007 уже принят только для полного диапазона; strict CIDR/plain и compact notation требуют решения |
| Q-08 полнота сценариев | Все 43 строки аудита имеют схемы и witnesses; второй шаблон независим | Принять полноту повседневного workflow и добавить неизвестные случаи |
| Q-09 «Общее» | Global projection не привязана к IP и не теряется из-за IP grouping | Выбрать канал и оформление глобальных инструкций; добавление нового листа не принято автоматически |
| Q-10 technical failures | CANNOT_EVALUATE, diagnostics, incomplete/output receipts выражены | Разрешить расхождение исторических PASS/fallback с v0.15; публикация затронутого пакета до выбора blocked |
| Q-11 port lookup | Схема может позднее получить PortSet column+join, без нового template controller | Самостоятельная портовая схема Superset не подтверждена; пока потенциальная функция, не часть доказанного v1-паритета |

Решённость вопроса хранится как DecisionRequirement с evidence. Не требуется закрыть все Q для обсуждения **структурной** модели, но релиз конкретного зависимого бизнес-сценария не публикуется с невыбранной семантикой. Отсутствие runtime источника — не то же, что открытое архитектурное решение: это отдельная диагностика в момент запуска.

## 5. Архитектурные риски и предлагаемые ограничения

| ID | Риск / граница | Предлагаемая мера / проверка до обещания |
| --- | --- | --- |
| A01 | «Вся логика редактируема» может превратиться в opaque operation, внутри которой скрыт Rulebook | Запрет операций вида EvaluateNetworkAccess/Rulebook; составные policy и selection видны как AST, алгоритм лишь вычисляет типизированный примитив |
| A02 | Если неизвестного алгоритма нет в библиотеке, UI не создаст его произвольно | Показать capability gap и новый contract/version; не обещать генерацию любого алгоритма без кода |
| A03 | Журнал и snapshot не дают настоящей транзакции Excel автоматически | Write-ahead progress, before/after verification, конфликт при чужой правке, separate report fallback; прототип merge/CF/legacy+threaded comments обязателен |
| A04 | Partial parser value может выглядеть как полностью валидный запрос | Invalid/ambiguous tokens остаются typed members+diagnostics; explicit valid-subset projection не даёт полного PASS |
| A05 | Matched-any Superset превращается в «всё найдено» | Каждому запросу coverage и uncovered IpSet; любой matched row не заменяет full coverage. Семантика старого match fact сохраняется отдельно |
| A06 | Signed VBA типы/Decimal/Date/Unicode и regex budgets могут не обеспечить контракт | Потребуются инженерные доказательства диапазона /0, locale boundaries, limits; никаких runtime библиотек/зависимостей не установлено в PR-01 |
| A07 | Техническая ошибка скрывается из-за audience/override либо branch | Диагностика и полнота сохраняются независимо от visible issues; no-source != no-match, error-to-PASS запрещён |
| A08 | Открытая грязная книга источника/замена файла вызывает смешанный run | Заморозить snapshot/revision; changed-during-read → retry/UNAVAILABLE; все consumers используют тот же token |
| A09 | Переименование/архивирование/rollback разрушает ссылки или предыдущие annotations | Immutable IDs, retention graph, version pin; disabled actions снимаются по journal; rollback не откатывает пользовательские данные |
| A10 | Shared source начинает грузиться у первой линии только из-за наличия coordinator report | Graph по активному profile+действительно достигнутым ветвям, без I/O при старте |
| A11 | Optional source/projection превращается в обязательный для всех правил | Requiredness задаётся по dependency branch, а не один флаг всего шаблона; объяснение неоткрытых источников |
| A12 | Сериализация/подпись ошибочно воспринимаются как решённая безопасность | Digest = контроль целостности, не аутентичность; выбрать deployment trust policy отдельно; данные не исполняются |
| A13 | Дополнительный путь DB.schema в v1 автоматически добавляет/сохраняет колонку Rulebook; чистое вычисление не выражает такой побочный эффект | **Архитектурный action gap:** нужен отдельный SourceMigration contract (допустимые изменения схемы, ownership, backup, explicit intent/receipts). Вариант — Admin-only preparation вместо Runtime write, но это несовместимость, требующая решения, не принятое упрощение |

Кандидат решения A13, **не утверждён**: `SourceSchemaMigrationPlan` содержит SourceDefinitionRef, from/to SchemaVersion, ожидаемый before digest, список ограниченных шагов AddColumn/AddValidation/InitializeMetadata, backup path/digest, preconditions (own-opened, not read-only, no unsaved external edits), postconditions и rollback/receipt policy. Вычислитель может только сообщить необходимость такого плана. Применение — отдельный зарегистрированный IO action adapter, не pure Call и не произвольный VBA из конфигурации. Разрушительные DeleteColumn/overwrite пользовательских данных в этот кандидат не входят. Изменение source-version инвалидирует snapshots/cache только для **следующего** run; текущий run не смешивает версии.

Два варианта для владельца: подготовка/миграция через Admin перед публикацией либо совместимое Runtime действие с явно разрешённой source-write capability и preflight. Рекомендуется сначала Admin-only preparation, но перенос auto-save поведения v1 этим не объявляется эквивалентным и до решения остаётся GAP_ACTION. Точный способ сохранения и восстановления source workbook требует отдельного инженерного контракта/приёмки.

## 6. Проверки выразительности и границы утверждения

Проверка PR-01 — статическая: каждая legacy строка имеет явные входы, tree composition, primitive calls, настраиваемое действие, editable business parameters, witness и пометку ограничений. Наличие слова «представимо» не засчитывается без схемы. Проверяются точные множества 17 RuleId и 26 LegacyPath относительно объединённой матрицы PR-00, а operation IDs должны разрешаться в каталоге. Нельзя сокращать проверку до 17 общих имён.

Итог статического покрытия: **43/43 строки учтены; 42 композиции и 1 GAP_ACTION (`DB.schema`)**. Conditional не означает принятую policy: до решений Q остаются ограничения публикации. Второй пример — семь схем EQUIPMENT_RETURN (required/email/lookup/duplicates/date/approval/report) на тех же core entities, AST и действиях.

Проверка пакета документации: восемь Markdown-файлов (шесть результатов + navigation/proposals в architecture/decisions), 17 точных LegacyRuleId, 26 точных LegacyPath без дублей, 41 определённая операция, 33 из них упомянуты матрицей; 86 pinned blob-ссылок на 22 repository/commit/file цели. Проверены наличие локальных целей/anchors, границы source-строк, соответствие каждой ссылки LegacyPath её строке PR-00, количество колонок Markdown-таблиц и отсутствие вызовов запрещённых opaque operations. Ошибок адресации/покрытия не найдено. Это проверка документов, не исполнение typed sketches.

Независимый обзор и последующая сверка исправили смешение ActionReceipt/journal lifecycle, идентичность action instance, origin issue/projection, вложенные member envelopes, completeness obligations после StateIs, short-circuit, scalar aliases, MatchScore-vs-number и strict parser для invalid raw keys. Значения native Date не передаются в текстовую нормализацию; отдельный общий read-profile сохраняет raw text/native date. Имена parser/normalization profiles в схемах являются кандидатными ссылками, их эксплуатационные значения не утверждены.

Witness ADR-007: .1–.10 имеет cardinality 10; deny .5 пересекает; allow .1–.8 оставляет .9–.10; источник с .1 и .10 оставляет .2–.9; list .1,.10 не равен диапазону; huge range обрабатывается интервалами. Эти ожидаемые результаты выводятся математически и документируются для всех применимых операций, а не выдаются за выполненный parser test. Declaration witness отдельно сохраняет raw Kind/prefix; утверждение о его эквивалентности не подменяет IpSet.

Независимый пример EQUIPMENT_RETURN содержит возвращающего сотрудника/email/date и таблицу asset/quantity/condition + реестр имущества. Он использует FieldRef, ColumnRef, SourceRows, сравнения, Join/GroupBy/Any/All и тот же Action/Report/Writer. Для его описания не нужны NETWORK_ACCESS, IP, порты, Superset или обязательный специализированный adapter. Его значения искусственные, пример не реализован и не является новым рабочим шаблоном.

Проверка документации не равна прогону движка: ни один Runtime статус не объявляется AUTOMATED PASS или EXCEL MANUAL PASS. Исходные AC остаются NOT STARTED. Существующие исторические доказательства PR-00 не превращаются в новые тесты PR-01.

## 7. Гейт перед программированием

1. Рассмотреть шесть документов и принять/изменить предлагаемые entity/AST/error/result contracts, альтернативы хранения, exact operation versions и правила публикации.
2. Для первой технической реализации отдельно согласовать SchemaVersion, ID/serialization convention, type storage/budgets и совместимость локального VBA без сетевых runtime зависимостей.
3. Для Writer отдельно выбрать journal storage/capability support/recovery contract и протокол ручной приёмки. Нельзя обещать восстановление неподдерживаемого форматирования.
4. Для реальных NETWORK_ACCESS-пакетов получить решения affected Q и версии эталонных Rulebook/обезличенных кейсов; не вставлять неизвестные значения в defaults.
5. Перед реализацией конкретного этапа — отдельное задание владельца. Ни core, ни формы, ни реальные правила в PR-01 не разрабатываются.

Порядок ревью: rule-model → expression-model → operation-catalog → 43 схемы → admin-workflows → таблица вопросов этого документа. Проверить один пример изменения координаты, один correlated ALLROWS, один интервал с внутренним запретом, один конфликт MergeArea и второй шаблон. Неразрешённая policy и невозможность Writer capability должны оставаться заметны; их нельзя закрыть общим утверждением о полной совместимости.
