# ADR-008 — XML-контракт определений и безопасная запись редакций

Дата: 19.10.2026. **Проект решения для ревью PR-02**; реализация ограничена структурной моделью. Основание: отдельное задание владельца PR-02 после объединения PR #33, ТЗ v0.15, [модель](../rule-model.md), [выражения](../expression-model.md). ADR-007 и открытые Q-01…Q-11 не изменены.

## Выбор формата

Выбран XML через late-bound `Msxml2.DOMDocument.6.0`. На доступном Windows реально выполнены активация, Unicode-разбор, установка защитных параметров и запрет DTD в **32- и 64-битных процессах**. Модули VBA загружены и исполнены в Excel 16.0 **x64**; Office x86 ещё не проверен.

| Критерий | XML / MSXML6 | JSON |
| --- | --- | --- |
| Windows Excel VBA | COM API парсера/DOM, без reference в VBE | В языке VBA нет встроенного JSON codec; нужен новый собственный парсер или проверенная библиотека |
| x86/x64 | Late binding без WinAPI-деклараций; оба COM-host проверены | Возможен чистый VBA, но отдельная работа по Unicode, числам, глубине и ошибкам |
| Надёжность прототипа | Проверены загрузка, UTF-8 round-trip, отрицательные XML и сбои записи в Excel | Исполняемый прототип JSON в этот PR не включён; надёжность не заявляется |
| Зависимости | Компоненты Windows MSXML6 и Scripting.Dictionary/FSO | Дополнительная библиотека либо поддерживаемый проектом парсер |
| Ограничения | Windows/COM, явный запрет DTD, особенность CR/LF в DOM | Нет оснований выбирать только ради удобства генерации файлов |

Это локальный формат проекта; не XLSX-внутренности и не исполняемый DSL. Никаких ScriptControl, Eval, Application.Run, Evaluate, макросов или формул из конфигурации в производственных модулях нет. Тестовый запуск вызывает только известную процедуру тестов.

## Физическое размещение

Предлагаемый корень: `%LOCALAPPDATA%\AccessTools2\`, с возможностью явно выбрать иной доверенный корень в будущем Admin. `Drafts\<PackageId>\<revision>.xml` — новые редакции черновика; `Published\<PackageId>\<PackageVersion>\package.xml` — неизменяемый опубликованный пакет. Runtime в будущем читает только Published, не выбирает автоматически свежий Draft. ACL и способ доверенного распространения определяются отдельно.

**PR-02 не создаёт эти каталоги автоматически**, не выбирает active package и не публикует/откатывает пакеты. `AtSaveNew` принимает конкретный путь нового файла; тесты используют только игнорируемый `artifacts/run-…`. Не сохранять конфигурации рядом с заявкой или в источнике неявно. Сам файл не содержит рабочих данных/учётных сведений.

## Модель 1.0.0

Две record-класса вместо десятка одинаковых классов: `CAtPackage` представляет PackageDefinition; `CAtDefinition.Kind` выбирает закрытую схему остальных сущностей. Нет диспетчера бизнес-обработчиков.

| Kind | Модель | Свойства Values | Обязательные ссылки |
| --- | --- | --- | --- |
| Template | TemplateDefinition | Нет | Нет; принадлежность полей/таблиц/правил задают обратные OwnerTemplate; допустимы явные Rule refs |
| Field | FieldDefinition | DataType | OwnerTemplate, Locator |
| Locator | LocatorDefinition | Strategy, Sheet; параметры стратегии ниже | OwnerTemplate; для TableColumn также Table и ровно одна Column |
| Table | TableDefinition | Нет | OwnerTemplate, Locator |
| Column | ColumnDefinition | DataType | Ровно один OwnerTable **или** OwnerSource |
| Source | SourceDefinition | Adapter=Excel, FileName, Sheet | Нет; источник общий и независим от шаблона |
| Binding | SourceBinding | Нет | OwnerTemplate, Source, минимум одна Column этого Source |
| Expression | Структурная NamedExpression-ссылка | ReturnType, Deferred=true | Нет; AST-тело не поддерживается |
| Rule | RuleDefinition | Нет | OwnerTemplate, Expression с ожидаемым Bool |
| Action | ActionDefinition | Trigger, Audience, Channel | OwnerTemplate, Rule |

У всех определений `Id`, `Version`, `DisplayName`, коллекции `References` и `Dependencies`; у пакета `Id`, `Version`, `DisplayName`, `SchemaVersion`, `State`, `PublishedBy`, `PublishedAt`, `Definitions`. Названия и значения — Unicode-строки. Все скаляры Values — строки с проверяемой лексической формой; произвольных пользовательских свойств нет.

ID: `[A-Z][A-Z0-9_]*`, длина 1–64, регистр значим; уникальность в namespace **Kind внутри пакета**. Одновременно хранится одна версия каждого Kind/Id. PackageId — отдельная идентичность контейнера. DisplayName не является ключом. Все версии в этом минимальном формате — три канонических числа 0…65535 без ведущих нулей, например `1.0.0`; диапазоны, prerelease/build и автоматический выбор версии не поддержаны. EntityVersion в этой реализации записывается той же лексической формой, но остаётся отдельной от SchemaVersion/OperationVersion/PackageVersion сущностью.

Reference хранит `Role, Kind, Id, Version, ExpectedType`. Версии exact. Для Field/Column/Expression/Operation обязателен ExpectedType; допустимые объявленные типы: Text, Bool, Int64, Decimal, Date, DateTime, Email, Hostname, IpSet, PortSet. Это **имена типов**, не реализованные IP/port-операции и не AST type checker. Не выполняются coercion или чтение значений Excel.

Дополнительные ссылки Rule: Field, Table, Operation; Action: Field; Expression: Operation. Владельцы должны совпадать, Column binding принадлежит выбранному Source, Column locator принадлежит выбранной Table. Rule зависимости допустимы только внутри своего Template. Expression зависимости относятся к общим Expression. Ссылки владения не становятся рёбрами исполнения; иначе естественная связь Template↔Rule ошибочно образовала бы цикл. `Dependencies` — ID внутри того же Kind и единственной закреплённой пакетной версии; для прочих Kind они запрещены. Циклы и глубина больше 128 диагностируются.

Хост передаёт в `AtValidate` доверенный Dictionary `OperationId@exactVersion → ReturnType`. Конфигурация не может объявить операцию установленной. `OP_TEST` существует только как синтетическая метаинформация в тестах, без реализации алгоритма. Полные входные сигнатуры, AST, scope и вычисление — отдельные этапы.

Locator: обязательный Sheet (буквальное имя 1–31 символ, без управляющих символов и : \\ / ? * [ ], без начального/конечного apostrophe) и одна стратегия. Такой же синтаксис Sheet проверяется у Source. A1 — верхний регистр, необязательные `$`, одна ячейка или возрастающий прямоугольник до XFD1048576. DefinedName — простое ASCII-имя (буква/underscore, далее буквы/цифры/underscore/точка), не A1 и не R/C. LabelOffset — непустой Label, целые RowOffset/ColumnOffset в границах размеров Excel. TableColumn — Table/Column refs. Параметры и ссылки другой стратегии запрещены. Наличие реального листа/имени, нахождение метки, объединения и окончательная координата не проверяются без будущего Resolver.

Action допускает Trigger Error/Warning/Info/Approval; Audience Initiator/FirstLine/Coordinator; Channel Results/Comment/Report. Это только структурные метаданные, без сообщений, Writer, маршрутизации или изменения книг.

Published требует непустые PublishedBy и корректный UTC PublishedAt `YYYY-MM-DDTHH:MM:SSZ`. Draft не должен содержать метаданные публикации. Это минимальная проверка структуры, **не разрешение публикации**: подпись/digest, доверенный автор, profiles, capabilities, AST и бизнес-решения здесь не проверяются. Любой результат содержит информационную диагностику `AST_DEFERRED`, если проверка дошла до этого этапа.

## XML, Unicode и совместимость

Namespace `urn:accesstools:definitions:1`, root `package`, обязательный `schemaVersion="1.0.0"`. Заголовок содержит остальные свойства пакета. Единственный дочерний `definitions` содержит `definition` с атрибутами kind/id/version/displayName и ровно по одной группе `properties`, `references`, `dependencies`. `property` — name/value; `ref` — role/kind/id/version/expectedType; `dependency` — id. Пустые группы и пустые обязательные атрибуты формы сохраняются явно. Полные примеры — [simple.xml](../../tests/fixtures/simple.xml), [equipment-return.xml](../../tests/fixtures/equipment-return.xml).

Неизвестные элементы/атрибуты, дубли групп/свойств, комментарии, непустой текст и processing instructions кроме XML declaration отвергаются. Структурные дубли ID/refs выявляет валидатор. SchemaVersion читается до интерпретации определений. Поддерживается только **точно 1.0.0**; новая версия требует нового reader/явной миграции, текущий reader не игнорирует незнакомые поля. XSD не требуется: правила формы проверяются кодом, затем семантическая структура — валидатором.

MSXML загружает байты согласно XML encoding declaration, DOM.Save записывает UTF-8 (при отсутствии declaration encoding это стандартная XML-кодировка); строки VBA остаются UTF-16 BSTR. Нет промежуточной ANSI-записи XML через VBA Print/FSO TextStream. Сериализатор создаёт XML declaration; свойство DOM.XML может опустить encoding, что не означает ANSI.

Технический прототип показал нормализацию CRLF → LF даже при MSXML attribute character references. Для точного сохранения скаляров контракт 1.0.0 перед передачей в DOM экранирует **процент `%25`, CR `%0D`, LF `%0A`, TAB `%09`**; обычный Unicode и `<>&"` обрабатывает XML DOM. При чтении обратное преобразование и проверка канонической записи; literal `%0D` сохраняется как `%250D`. Это не URL-кодирование всей строки и не исполняемый текст. XML, созданный вручную, также должен соблюдать это правило. Изменение encoding в будущем требует смены SchemaVersion.

Ограничения: локальный файл до 1 MiB, текстовый API до 524288 UTF-16 code units, до 2000 определений, DOM depth ≤32. `MaxXMLSize=1024` — блоки по 1024 символа ([Microsoft](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms765399(v=vs.85))), отдельный байтовый лимит проверяется до чтения. DTD запрещён, external resolution и validateOnParse отключены; конфигурация не загружает схемы, XSL или внешние сущности.

## API и диагностика

| API | Контракт |
| --- | --- |
| AtNewPackage / AtNewDefinition / AtAddReference | Построить record-модель, без I/O |
| AtValidate(package, catalog) | Collection диагностик; AtHasErrors проверяет Severity=Error |
| AtToXml(package) | Детерминированная XML-строка с сохранением порядка; не заменяет валидацию |
| AtFromXml(text, catalog, diagnostics) / AtLoad(path, catalog, diagnostics) | Сначала безопасный разбор/форма/версия, затем валидатор |
| AtSaveNew(package, path, catalog, diagnostics) | Проверить и сохранить **новый** файл редакции; Boolean означает подтверждённый commit файла |

Диагностика: `Code, ObjectId, Property, Reason, Message, Severity`. Message — понятный текст на английском; локализация будущего Admin отдельно. Основные семейства: ID_DUPLICATE/ID_INVALID/VERSION_INVALID, REF_*, TYPE_UNKNOWN, LOCATOR_*, DEPENDENCY_*, OPERATION_UNKNOWN, PUBLISHED_REQUIRED, SCHEMA_VERSION, XML_INVALID/XML_SHAPE/XML_RUNTIME, FILE_LIMIT, IO_READ/IO_WRITE/FILE_EXISTS.

Повреждённый XML, неизвестная версия или неверная форма → Nothing + ошибки. Синтаксически допустимая, но структурно неверная модель → модель + ошибки, чтобы будущий Admin мог показать причины. Нельзя использовать одну проверку `Not package Is Nothing` как признак валидности. `AtSaveNew` в PR-02 допускает только модели без структурных ошибок; сохранение неполного редактируемого черновика с диагностикой из PR-01 ещё не реализовано и потребует отдельного draft-save API. Info AST_DEFERRED не блокирует запись тестовых определений.

## Безопасное сохранение

1. Валидировать модель с внешним каталогом. Существующий целевой файл/каталог → FILE_EXISTS, никаких изменений.
2. Эксклюзивно создать принадлежащий этому вызову временный файл в том же каталоге (`CreateTextFile(..., False)`). DOM.Save пишет только в этот staging-файл.
3. Загрузить его тем же reader, повторить структурную валидацию и сравнить повторную сериализацию с исходной. Ошибка/потеря значения → не commit.
4. `Name temp As target` переименовывает в пределах тома и отказывает при существующем target; нет `Kill target`/overwrite или copy-over-existing.
5. При ошибке удалить только собственный staging-файл. Предыдущие редакции никогда не открываются для записи.

Функция не обещает crash-durable транзакцию ОС/сетевого хранилища. После аварийного завершения может остаться staging; он не является активной редакцией. Внешние вмешательства/ACL каталога, reparse points и конкурентный недоверенный автор не решаются файловым именем: рабочий каталог должен быть доверенным. Контроль active pointer, manifest/digest, retention, rollback, административная миграция Rulebook и восстановление после питания здесь отсутствуют. Минимальный безопасный контракт — новая проверенная редакция при неизменности предыдущих; full publication отдельным этапом.

## Ограничения и решения для ревью

Подтвердить XML 1.0.0 и scalar escaping; две generic record-модели с закрытой схемой Kind; append-only сохранение с отдельными файлами редакций; предлагаемое размещение Drafts/Published. Windows MSXML/Scripting могут блокироваться политиками организации: это диагностируется, обход настроек безопасности не выполняется. Office x86, Mac/Web, сетевые диски/авария питания и реальный пользовательский сценарий не сертифицированы. Полный отчёт и воспроизведение — [pr02-validation](../pr02-validation.md).
