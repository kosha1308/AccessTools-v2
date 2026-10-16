# Подготовка данных NETWORK_ACCESS к проверкам

Полный пользовательский сценарий остаётся целью: проверки заявки, Superset,
сводка координатора. Этот этап превращает подтверждённый raw snapshot в typed
контекст без чтения или изменения Excel. Ядро остаётся независимым от шаблона.
Технические решения делегированы пользователем; выполняем inline без повторного
согласования внутренних классов.

Выбран отдельный adapter над существующими readers и domain parsers. Изменять
raw reader нельзя: исходный текст и ошибки нужны для объяснений. Встраивать
знание формы в generic evaluator тоже нельзя. Отдельный шаг позволяет повторно
обрабатывать тот же snapshot и проверять каждый слой независимо.

## Даты

ATExcelDateAdapter.Normalize(raw As Variant, date1904 As Boolean, textFormat As String)
As ATEvalResult. Явные форматы ISO (yyyy-mm-dd) и DMY_DOT (dd.mm.yyyy), без
угадывания по локали. Внешние ASCII-пробелы обрезаются, внутренние не исправляются.
Числовые Value2: целые дни, 1900 serial1..2958465 кроме фиктивного 60;
1904 serial0..2957003. Отрицательные/дробные/выходящие за диапазон INVALID.
Строка цифр не становится serial. Неподдерживаемый формат — CONFIG_DATE_FORMAT
ERROR, Empty — EMPTY, Null — MISSING, CVErr — EXCEL_CELL_ERROR ERROR,
bool/array/object/Date — INVALID. DATE хранится как Gregorian VBA Date через
существующий независимый от локали ordinal parser; дата/время компьютера не читаются.

## Нормализатор заявки

ATNetworkRequestNormalizer.Normalize(source As ATNetworkReadResult,
dateTextFormat As String, unlimitedKeyword As String) As ATNetworkReadResult.
Оба параметра policy обязательны: формат только ISO/DMY_DOT; keyword непустой,
не дата в выбранном формате и не строка цифр. Значение keyword сравнивается
через уже принятую Excel header normalization (case/whitespace). Никаких новых
синонимов, значений enum, правил revoke или допустимых протоколов не придумываем.
Если source ERROR — его диагностика возвращается прежде проверки policy.

Сохраняются 11 document RAW fields и их source refs, обе RAW collections,
WORKBOOK_DATE1904, VALIDITY_VALUE2_KIND, VALIDITY_SERIAL. Документные TEXT поля
без суффикса RAW сохраняют исходные значения; отдельное семантическое разрешение
EMAIL/ENUM/HOSTNAME остаётся будущим этапом.

VALIDITY_DATE: typed DATE; VALIDITY_IS_UNLIMITED: BOOL. Keyword даёт True и
MISSING дату; дата — False и DATE; пустота — False и EMPTY дату; invalid — False
и INVALID дату. Нативный BOOL срока INVALID, исходный CVErr сохраняет ERROR в
обоих результатах. Все полученные поля ссылаются на исходную C12. Состояния
MISSING/EMPTY/INVALID и ERROR не смешиваются; сбой данных не прерывает остальные
поля. Системный/конфигурационный сбой не выдаёт частичного Context.

TECHNICAL_ROWS: ROW<TECHNICAL_TABLE>, поля TECH_SOURCE_IP/TECH_DEST_IP
IP_EXPRESSION, TECH_DEST_PORT PORT_SET через действующие parsers; TECH_PROTOCOL
TEXT с внешним ASCII Trim и ASCII uppercase (это канонизация, не проверка списка);
TECH_REASON TEXT без изменений. ADDRESS_ROWS: ROW<ADDRESS_TABLE>, SYSTEM,
HOSTNAME, ROLE TEXT без изменений, ADDRESS_IP IP_EXPRESSION. RowId/порядок/CellRef
сохраняются, строки не добавляются/не удаляются, неверный token не спасается
частично. Пустые и ошибочные поля сохраняют своё состояние в typed поле.
Никаких сопоставлений между таблицами, Superset или бизнес-решений здесь нет.

Снимок frozen, независим от исходника. Source не мутируется. Лимиты: максимум
10000 строк на коллекцию, 200000 общих шагов и 65536 символов на значение
из штатного runtime budget (лимит чтения 1million cells остаётся в reader);
чужой malformed snapshot даёт NETWORK_INPUT_INVALID, не частичные данные.

## Приёмка

Полный real Excel compile/tests. Границы обеих эпох, serial60, дроби, leap days,
неверные/неподдерживаемые форматы, locale independence, numeric-vs-text, errors,
blank; реальный reader -> normalizer -> compiled domain/date rules + trace;
invalid mixed tokens, row errors, merged/spacer metadata, immutable snapshot,
policy validation and source error priority. Независимый обзор, static/gate,
чистая копия и draft PR на предыдущий этап. Не объявлять пользовательскую готовность.
