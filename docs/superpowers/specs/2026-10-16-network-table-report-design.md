# Таблицы заявки и отдельный отчёт Excel

Цель продукта — полный сценарий заявка/Superset/координатор; данный этап даёт
локальные проверки заполнения таблиц и видимый русский отчёт. Технические решения
делегированы, inline execution без повторных согласований внутреннего кода.
Семантика только read-only reference AccessTools d9aa9b8, без копирования кода.

## Проверки

ATNetworkTableChecks.Run(source As ATNetworkReadResult, allowedProtocols As Collection)
As ATCheckReport. Внутри существующая normalization DMY_TEXT/Бессрочно. Scope
NETWORK_TABLE_INPUT; outcomes NET.TECHNICAL_TABLE, NET.PROTOCOLS, NET.ADDRESS_TABLE.
Pure snapshot, ни Range/Worksheet, ни автоматического открытия справочников.

Верхняя таблица: source/destination/protocol/port обязательны; порт 1..65535,
списки/диапазоны через существующий parser. Математический core по-прежнему
разрешает порт0; бизнес-ограничение применяется только здесь. Пустое обоснование
WARNING. Дубликаты первых четырёх исходных полей после header normalization —
WARNING с номером первой строки; collision-safe length-prefixed key.
Нижняя: address/hostname обязательны, system/role пустые WARNING. Проверка формата
hostname по Rulebook и TOP↔LOWER ещё не выполняется.

Чистые поддерживаемые IP expressions проверяются существующим parser; numeric
invalid выдаёт ERROR. Непустое нечисловое выражение вне IP-алфавита может быть
группой: WARNING ADDRESS_MANUAL_REQUIRED и outcome SKIPPED, а не ложный PASS
или запрет группы. Это консервативная ручная проверка, не полноценный legacy
mixed-group classifier. При наличии обычной ERROR outcome FAIL имеет приоритет.
Все CVErr — отдельные ERROR, не пустота. Raw refs и row identity сохраняются.

Список протоколов передаётся явно: элементы непустые ASCII literals до128символов
(letters/digits + _-./+), case-insensitive, max1000. Нет каталога/Nothing/пустой:
SKIPPED + WARNING; остальные проверки продолжаются. Неверный каталог — весь
табличный результат ERROR без частичных PASS. В заявке comma/semicolon/pipe,
ASCII whitespace разделяют протоколы; каждый token должен быть в каталоге.
Пропущенный протокол отмечается обязательностью, не дублируется в membership.

Empty table даёт одну ERROR с ref текущего заголовка, не фиктивные строки.
Reader добавляет CellRef к RAW collections (верхний первый/нижний адресный header),
normalizer сохраняет refs. Missing refs/schema/system failures блокируют результат.
Скан до10000строк, общий runtime budget200000steps/65536text. Max990issues в table
report (10 зарезервированы для document); превышение блокирует весь результат.

ATNetworkLocalChecks.Run(source,allowedProtocols) объединяет6document+3table
outcomes, Scope NETWORK_LOCAL. При ERROR любого этапа публикуются только blocked
issues/checks того этапа, без положительных частичных результатов. Все этапы
offline, source snapshot не меняется. Полная проверка NETWORK_ACCESS не заявляется.

## Отчёт

ATExcelReportWriter.Create(report As ATCheckReport, source As Worksheet) As Workbook.
Создаёт только новую книгу, листы «Отчёт» и «Проверки», ничего не сохраняет/не
перезаписывает автоматически. Исходная книга не меняется. Сохраняет/восстанавливает
EnableEvents даже при исключении, ошибочную собственную книгу закрывает без сохранения.
При повторе создаёт новую книгу. source=Nothing допускается, ссылок тогда нет.

Видимая часть целиком на русском: заголовок, счётчики, scope, статусы, названия
проверок/полей. Всегда заметна оговорка: проверена только часть заявки, Superset
и сводка не выполнялись. Колонки №/Важность/Проверка/Поле/Сообщение/Ячейка;
текст переносится, severity цветом, фильтр, печать landscape/A4 в ширину1страницы.
Отдельный лист хранит outcomes и причины пропуска. Нулевые issues не объявляют
всю заявку корректной. Разметка сдержанная navy/teal, без внутренних ID в UI.

Все строки перед записью форматируются Text: сообщения/refs не могут стать
формулами. Hyperlinks только к проверенному A1 в переданном source, sheet name
должен точно совпасть; никаких URL/path из report. Saved source: file address +
sheet SubAddress; unsaved source: open-workbook SubAddress с bracket book name.
Невалидный/чужой ref выводится обычным текстом без ссылки. Не следовать ссылкам
из недоверенного текста. Закрытый/перемещённый источник требует повторного запуска.

## Приёмка

TDD real Excel: rows/spacers/merges/empty tables/fresh anchors, port0, invalid IP,
groups/manual skip, protocol multi-values/missing/bad catalog, duplicates with
delimiter collisions, CVErr, source rejection, local6+3 aggregation, issue limit.
Writer: Russian labels/counts/layout, separate books, unchanged source, copied
report, formula-like text, trusted/foreign refs, events restore, repeat, no save.
Synthetic report XLSX + PDF render for visual QA; source-free sample may be delivered
only as example, not working app. Independent whole-branch review, gates7/7+9/9,
original ReadOnly/hash, clean checkout, draft PR base feat/network-document-checks.
