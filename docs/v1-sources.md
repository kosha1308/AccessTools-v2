# Источники AccessTools 1: аудит PR-00

Статус: наблюдения по коду, не проектирование PR-01. Последний рабочий обязательный эталон: integration 58d8789d1e60860c96cb35ba0e780a6424510130. Параллельный исторический эталон main: d9aa9b88f9193c0d2935d99c5cc8ecca3865687c. Нижеследующая исходная инвентаризация main дополнена обязательными поздними отличиями integration; main не заменяет integration. ТЗ v0.13 и v0.14 рассматриваются как требования, а не доказательство реализации. Содержимое рабочих строк не включено.

## Конфигурация и фактические файлы

[Конфигурация источников](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/config/templates/network_access/data_sources.ini#L1)

| SourceId | Файл / папка | Роль |
|---|---|---|
| RULEBOOK | workspace/network_access/settings/01_Правила_NETWORK_ACCESS.xlsx | Единая рабочая книга правил; редактируема |
| ALLOWED_PROTOCOLS | та же книга, Протоколы | Видимый SourceId без отдельного редактирования |
| IP_REGISTRY | workspace/network_access/superset/*.xlsx | Папка выгрузок Superset; редактируема, interval cache |
| ALLOWED_INITIATORS, INITIATOR_APPROVAL_RULES | та же книга, Инициаторы | Совместимые внутренние ID |
| IP_EXCEPTIONS, HOSTNAME_RULES | та же книга, Правила и исключения | Совместимые внутренние ID |
| USER_MESSAGES | та же книга, Проблемы и Общий текст | Сообщения и параметры |

Runtime бизнес-логики читает RULEBOOK: отдельный File override старого ID не является новым независимым источником. В репозитории также есть legacy defaults hostname_rules.xlsx, initiator_approval_rules.xlsx, ip_exceptions.xlsx, network_ip_registry.xlsx и allowed_initiators.txt / allowed_protocols.txt; наличие файла не доказывает его участие в текущем маршруте.

## Схема Rulebook

Заголовки исторического репозиторного XLSX main проверены read-only openpyxl; строки данных не читались в отчёт. Эта проверка подтверждает структуру файла, но не поведение Excel.

| Лист | Заголовки / метаданные |
|---|---|
| Содержание | титул; E1 DEMO_DEFAULTS; F1 SUPPORT_V2 — служебные маркеры, не бизнес-правила |
| Инициаторы | Активно; Домены / почты; Требуемое согласование; Уровень сообщения; Комментарий / инструкция; Служебный режим |
| Правила и исключения | Активно; Название правила; IP; Область IP; Hostname/FQDN; Тип совпадения hostname; Порт назначения; Протокол назначения; Сервис — условие; Сервис — результат; Группа назначения; Требуемое согласование; Комментарий / инструкция; Уровень сообщения |
| Протоколы | Активно; Протокол; Комментарий |
| Проблемы | A:I: Поле заявки; Что проверяется; Штатное сообщение; Сообщение пользователю; Сообщение координатору; Уровень сообщения; Ошибка пользователя; Подсвечивать в заявке; Активно. J пустой; K:M Ключ; Текст для пользователя; Текст уведомления |
| Общий текст | Ключ; Значение |

[LoadNetworkUnifiedRulesFlex](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkRulebookFlex.bas#L12): поиск заголовков в строке 1, строки данных со 2-й. Поддерживаются старые aliases «Тип совпадения», «Порт», «Протокол», «Сервис условие», «Сервис», «Согласующий», «Примечание». Без Активно / листа / открываемой книги получается пустая Collection; исключение даёт Nothing. Неактивные строки пропускаются; невалидный непустой IP пропускается; строка без условий не добавляется. Все заполненные условия одной строки соединены AND. Валидатор строже загрузчика: проверяет наличие условий и результатов, зависимость scope от IP, типа от hostname, формат IP/портов, активные протоколы, severity. Эти разные выходы нельзя сводить к одному «нет совпадения».

[LoadNetworkInitiatorRulesV2](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkSupportSheets.bas#L579): aliases Условие, Разрешён, Согласование, Уровень, Комментарий; каноническая активная строка разрешает домен/почту, кроме служебного approval-only. Пустое условие пропускается. Служебный режим сохраняет семантику старых строк согласования и не является разрешением.

[Direct schema](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemSchemaV5.bas#L4) и [Support migration](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkSupportSheets.bas#L18): схема «Проблемы» эволюционировала; INI и миграционные aliases не заменяют фактический A:M layout. Автоматический запуск миграции в этом аудите не выполнялся.

## Superset: поиск и приоритет

[FindRegistryIntervalHeaders](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modRegistryIntervalIndex.bas#L754): перебор листов в порядке книги, поиск в первых 30 визуальных строках; обязательны в одной строке SOURCE:ТЕЛЕГРАФ, SOURCE: ЦХД, SOURCE: AWX. FQDN и Сервис — дополнительные поля обогащения, а не обязательные SOURCE-колонки. IP объединяются из трёх SOURCE-полей с сохранением связи с исходной строкой.

[RegistryLatestCandidatePath / validation](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modCheckRegistry.bas#L545): runtime выбирает первый кандидат; валидатор перебирает кандидатов до книги с нужными колонками. Это может дать OK для старой книги при непригодной новейшей. [CandidateFiles](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modRegistryCacheUnified.bas#L361) документирует фильтр *.xlsx, исключение ~$ и demo_network_ip_registry.xlsx, сортировку FileDateTime по убыванию; равные даты не задают отдельного приоритета. Старый unified parser не следует путать с текущим interval index.

[EnsureNetworkRegistryIntervalCache](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modRegistryIntervalIndex.bas#L27): NO_FILES, UNSAVED_SOURCE, HIT, BUILD_FAILED, REBUILT, REBUILT_NOT_SAVED, ERROR различаются. Кэш network_ip_registry.intervals.txt и .meta, версия 2; session index зависит от пути, размера, времени изменения. CIDR сохраняет сеть и исходный host alias. Непустой кэш не доказывает актуальность источника.

[AcquireSourceWorkbook](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modDataSources.bas#L19): уже открытая книга переиспользуется, включая несохранённые значения; иначе ReadOnly=True, UpdateLinks=0. Release закрывает только книгу, открытую надстройкой, без сохранения. Поэтому проверка дискового файла и проверка открытого Excel могут различаться.

## Доказательства и пробелы

Код прочитан, схема XLSX прочитана без записи. VBA regression entry points и сценарии developer/tests — имеющиеся тесты, их исполнение не подтверждено. Ручная проверка Excel, миграций, блокировки файлов, обновления и просмотра совпавших строк не выполнялась.

Конфликты: permissive loader против strict validator; latest runtime против validator fallback; INI USER_MESSAGES против direct schema V5; отдельные legacy files против единого Rulebook. ТЗ v0.14 требует SourceBinding и независимости операций от фиксированных листов; v1 содержит жёсткие имена листов/заголовков. Это исходное ограничение для паритета, не разрешение менять поведение.


## Обязательный рабочий эталон integration 58d8789

Все ссылки этого раздела закреплены на integration. Доступный XLSX выше не подтверждает материализованную после миграций книгу integration: её физическая схема отдельно не установлена. Ниже — доказанные кодом runtime-расширения; неизвестные позиции колонок не выдуманы.

| Источник / слой | Схема и контракт integration |
|---|---|
| Rulebook: Ограничения IP инициаторов | Заголовки строки 1: Активно; Email / домен; Область IP; Разрешённые IP; Запрещённые IP (НЕ); Комментарий / инструкция. Первые пять обязательны. Данные со строки 2. Runtime добавляет лист в существующую книгу при отсутствии, если она не открыта пользователем; если создание невозможно — NOT_CONFIGURED с пустой Collection. Неполная схема — INVALID, ошибка строки — Nothing + INVALID с номером строки. Активная строка пропускается при пустых Email / домен + Разрешённые IP + Запрещённые IP (НЕ), независимо от заполненности Scope/Comment; отсутствие selector, плохой scope, пустые allow+deny или неверные selectors — ошибка конфигурации. |
| Связь политики с инициатором | Логическая связь Email / домен с активным разрешением на листе Инициаторы; отдельной колонки «Связь политики с инициатором» в прочитанном коде нет. Dropdown _AT_INITIATOR_POLICY_SELECTORS берёт условия Инициаторы; сама policy не даёт права подачи заявки. |
| DB policy | В том же policy-листе добавляется «Учетная запись для портов БД». Условие по email/domain выбирается по specificity; запрет требует пустого поля заявки «Учетная запись» в блоке «Поля обязательные для портов БД». Это не новый отдельный источник БД. |
| Simple Rulebook | Сохраняются runtime-колонки прежнего листа; справа добавляется Инструкция. При первом создании копируется из Требуемое согласование; далее Инструкция — источник истины и зеркалится в старое runtime-поле. Group/approver/comment/severity/legacy note скрываются, не удаляются. Открытая пользователем книга не мигрируется; изменения сохраняются только для собственной открытой книги. |
| Superset detail | DetailIpHeaders: IP, IP адрес, IP-адрес, Адрес, IP хоста, Адрес IP, IPv4. DetailHostnameHeaders: Имя хоста, Hostname, FQDN, Host, Имя узла. DetailServiceHeaders: Сервис, Service, Название сервиса. DetailFullSystemHeaders: варианты Полное имя/наименование/название ИС и информационной системы. Заголовки нормализуются по регистру, пробелам, переносам и уточнениям в скобках. Exact IP+hostname, fallback exact IP; это enrichment, не замена SOURCE index. |
| Связь результата Superset со сводкой | Superset: № из сводки; координаторская сводка: №. Агрегируются обычные FQDN/Сервис/Полное имя ИС/СКУФ ГРУППА ТП/КОМАНДА ПОДДЕРЖКИ. Целевые колонки имеют суффикс (Superset). Hostname публикуется лишь при единственном FQDN; остальные значения объединяются без дублей. Старые промежуточные AccessTools: ... не источник сводки. |

Код схемы и loader: [policy loader](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L340), [policy schema](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L502), [linked gate и DB policy](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modInitiatorPolicyExt.bas#L14), [SimpleSchema](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRulebookSimpleSchema.bas#L19), [detail INI](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/config/templates/network_access/data_sources.ini#L149), [summary join](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L25).

### Locator и mapping-конфигурация (integration)

[template.ini](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/config/templates/network_access/template.ini#L1): CurrentTitle, InstructionUrl; Recognition MinMarkers=3 из пяти маркеров «Система в ИЭП?», «IP источника», «IP назначения», «Адрес», «Имя хоста». Тип проверки выбирается вручную; markers служат защитой от чужого файла.

[fields.ini](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/config/templates/network_access/fields.ini#L1): INITIATOR_FULL_NAME, DEPARTMENT, POSITION, EMAIL, CONTACT читаются BELOW под label; REQUEST_TYPE, SYSTEM_IEP, TOP_JUSTIFICATION, VALIDITY — AFTER справа от объединённого label. Техническая таблица: SourceIpHeader, DestinationIpHeader, ProtocolHeader, PortHeader, ReasonHeader. Нижняя таблица: SystemHeader, AddressHeader, HostnameHeader, RoleHeader. Это locator schema, а не фиксированные координаты. [Extractor](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkExtractor.bas#L19) читает MergeArea, границы обеих таблиц и logical rows; несколько физических строк не обязательно несколько независимых source records.

[checks.ini](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/config/templates/network_access/checks.ini#L1) задаёт включение/порядок check IDs, но поздние wrappers диспетчера меняют фактический маршрут; одного файла недостаточно. [Runtime mapping](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRulebookFlex.bas#L12) связывает условия IP/scope/hostname/type/port/protocol/service-condition с Services/Groups/Approvers/Comments, SourceRow и severity; Service condition и Service result различны. Все заполненные условия строки — AND, список hostname — OR. SimpleSchema меняет представление, не переносит столбцы вычислителя.

### Приоритет требований и конфликты

ТЗ13 H–M и O фиксируют позднюю policy, нижние адреса при группе, merge/blank handling, координаторскую сводку и T34–T39. В ТЗ13 H NoPolicyAction открыт; последняя ТЗ14 §7.3 устанавливает ALLOW, а integration при отсутствии selected policy завершает проверку без нарушения. Это не ALLOW при ошибке чтения источника.

Существенный конфликт ТЗ14 §7.3 с последней рабочей реализацией и ТЗ13 O: ТЗ14 описывает полный explicit range как два endpoints и compact form как будущее; integration CompactCoverage сравнивает ranges/списки по полному фактическому покрытию и поддерживает compact. В ТЗ14 одновременно сказано сравнивать CIDR по охвату; TOP↔LOWER integration оставляет CIDR строгой декларацией. Требуется решение владельца для будущего переноса; аудит не заменяет реализацию формулировкой ТЗ. ТЗ13 O заявляет PASS gates и ручное подтверждение; это исторические сведения документа, не выполненная здесь проверка.

### Адреса и порты Superset: подтверждённая граница

По integration подтверждён IP/интервальный lookup в SOURCE-колонках, затем уточнение hostname и получение сервиса/metadata. Поиск порта в самой выгрузке Superset **не подтверждён**: [BuildRegistryIntervalIndex](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryIntervalIndex.bas#L242) строит index адресов, а [контекст](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L117) читает IP/FQDN/Service без port-column binding. Порт — предикат технической строки против Rulebook ([RuntimePortMatches](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRulebookFlexCache.bas#L426)), а не доказанное поле lookup Superset. Формулировка ТЗ14 §7.2 «IP+порт в Superset» остаётся требованием/пробелом до отдельного доказательства; её нельзя представить как уже реализованную операцию.

Hostname narrowing безопасно сохраняет unhinted rows, а если совпадений hostname вообще нет — сохраняет все IP-matched rows. При наличии IP+hostname совпадений удаляются лишь несовпавшие строки с явно привязанным к этому IP hostname. [Точная ветка](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L214). Это не строгий отказ при hostname mismatch.
