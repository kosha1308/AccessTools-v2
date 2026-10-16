# Единое чтение формы NETWORK_ACCESS

Цель — получить один согласованный immutable контекст заявки после проверки
реального профиля v3.0. Используем существующие readers; ядро остаётся универсальным.
Порядок: Verify -> проверка параметров запуска -> свежие layouts -> ограниченные
таблицы и11 одиночных полей -> snapshot. Отказ профиля останавливает всё до данных.

## Контракт

ATNetworkRequestReader.Read(ws,runId,runDate,timezoneId) As ATNetworkReadResult.
Результат содержит State, Diagnostics (копия), Context (snapshot только при OK).
Структурная/техническая ошибка не выдаёт частичный Context. Ошибки пользовательских
ячеек сохраняются как ERROR соответствующих полей в успешно извлечённом контексте:
следующий слой проверок сможет выдать несколько понятных ошибок с правильными адресами.
Одна ошибка шаблона остаётся одним русским сообщением.

11 document fields имеют суффикс _RAW: INITIATOR_FULL_NAME, DEPARTMENT, POSITION,
EMAIL, CONTACT, REQUEST_TYPE, SYSTEM_IEP, DB_ACCOUNT, PERSONAL_DATA, JUSTIFICATION,
VALIDITY. Их описания возвращает AT_NetworkDocumentFieldsV3() As Collection.
DB account читается из C10, а не из метки C9. Значения TEXT остаются целиком.

AT_NetworkTechnicalSpecV3() и AT_NetworkAddressSpecV3() создают новые описания
таблиц из тех же заголовков, что profile factory. TECHNICAL_ROWS_RAW содержит
ROW<TECHNICAL_TABLE_RAW>: SOURCE_RAW, DESTINATION_RAW, PROTOCOL_RAW, PORT_RAW,
REASON_RAW. ADDRESS_ROWS_RAW содержит ROW<ADDRESS_TABLE_RAW>: SYSTEM_RAW,
ADDRESS_RAW, HOSTNAME_RAW, ROLE_RAW. Все поля исходные TEXT/EMPTY/ERROR.
Upper заканчивается перед текущим нижним заголовком; lower — на LastValueRow
только рабочих колонок. Пробелы/merge обрабатываются имеющимися readers.
Нет кеша по имени листа или переиспользования границ между запусками.

## Значения Excel и дата

Чтобы будущая нормализация даты не путала числовую дату Excel со строкой цифр,
ATExcelFieldValue сохраняет immutable scalar RawValue2 (включая Empty/CVErr).
Обычный Result остаётся прежним TEXT: совместимость существующего DSL не меняется.
В document context добавлены WORKBOOK_DATE1904 (BOOL), VALIDITY_VALUE2_KIND (TEXT)
и VALIDITY_SERIAL (NUMBER, MISSING если исходник не numeric; ERROR при ошибке ячейки).
Нативные String/Double/Boolean/Empty/Error не превращаются друг в друга.
Никакой даты по локали, TODAY или default policy не угадывается. Date1900/1904
и keywords будут нормализованы отдельно по опубликованной политике.

## Сохранность и приёмка

Нет записи в книгу, настройки Excel и исходный эталон. Снимок работает после
закрытия/изменения источника, не хранит Worksheet/Range. Run context явный:
locale INVARIANT, comparison BINARY; runId/date/timezone передаёт вызывающий код.
Лимиты reader/profile сохраняются. Невалидный run context не публикуется.

Реальные временные листы: все11 полей/адреса, C10 regression, изменённая структура,
пустая форма, upper/lower spacer/merge, перенос нижнего заголовка, CVErr документа/
таблицы, данные после пустой строки, форматирование конца листа, Date1904 и numeric
vs string validity, immutable snapshot, compiled COUNT/ALLROWS и cell trace.
Отдельно — read-only запуск на предоставленном оригинальном XLSX с unchanged SHA256.

Нормализация бизнес-типов, полный набор checks/Rulebook, IP matcher/cache Superset,
writer сводки и пользовательский XLAM остаются следующими этапами.
