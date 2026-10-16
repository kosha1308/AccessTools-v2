# Точное сопоставление объявлений верх/низ

Продолжаем fea4c22/PR13. Проверки заполнения и writer завершены и не переписываются.
Новая проверка NET.ADDRESS_MATCH добавляется к локальному отчёту. Полный сценарий
Superset/координатор остаётся целью; технические решения делегированы, native inline.

ATNetworkAddressMatch.Run(source As ATNetworkReadResult) As ATCheckReport.
Работает по RAW snapshot: SOURCE_RAW/DESTINATION_RAW верхней таблицы и ADDRESS_RAW
нижней. Нет чтения/записи Excel, внешних источников, нормализации всего контекста.
StateERROR источника имеет приоритет; malformed schema/ref/system failure блокирует
результат без partial PASS. Только EXCEL_CELL_ERROR является data error: при ней
сопоставление целиком SKIPPED (ошибки ячеек уже показывает table input checker).

Разделители/пробелы/Unicode dash — существующий IP tokenizer, без изменения core.
Каждый token валидируется существующим AT_ParseIPToken с prefix для сокращений.
Однако ключ строится отдельно от интервальной математики: IP:canonical-address,
RANGE:canonical-low-high, CIDR:введённая canonical база/маска. Ни объединения
пересечений, ни сведения IP к /32. Одноточечный диапазон отличается от IP.
Развёрнутый/сокращённый диапазон эквивалентны, порядок/повторы несущественны.

Неверный IP-подобный token сохраняется целиком как INVALID:uppercase-token;
не обрезается и не спасается частично. После invalid prefix сбрасывается, чтобы
не фабриковать адрес из следующего голого числа. Это явное консервативное правило.
Текстовые группы, неподдержанная семантика/маски x и неоднозначные нечисловые
записи требуют SKIPPED всего сравнения, без выдачи ложных несовпадений. Token,
начинающийся цифрой и содержащий точку или slash, считается IP-подобным даже
с буквенным хвостом; его нельзя потерять как текстовую группу. Colon/IPv6 — manual.
Неподдержанные маски не запрещаются, просто не подтверждаются этим comparator.

Индекс хранит ключи и уникальные refs каждого объявления. Проверка сначала
указывает лишние нижние объявления, затем недостающие внизу верхние, с ERROR
на каждую исходную ячейку. Дубли одного ref не создают повторных issues.
Пустой верхний набор → SKIPPED; пустой нижний при непустом верхнем → missing errors.
Одинаковые invalid tokens могут совпасть как объявления, но это не снимает
ошибки формата из другой проверки и не означает валидность заявки.

Scope NETWORK_ADDRESS_MATCH, один outcome. Limits10000rows,10000keys,
65536chars/value, shared200000steps, max990issues. Limit → один BLOCKED report.
ATNetworkLocalChecks объединяет существующие9+1 outcomes; общий max1000issues.
При техническом отказе нового этапа/переполнении итог не публикует частичные PASS.
Writer получает только новую русскую подпись NET.ADDRESS_MATCH; алгоритм не меняется.

Тесты: пары IP/CIDR/ranges/invalid/groups, typed input failure, refs/merges/spacers,
пропуски и лишние значения, предел issues, frozen source и closed workbook;
local10outcomes и существующий writer. Full real Excel compile, one independent
review, static7/7 gate9/9, original ReadOnly/hash and clean checkout. Draft PR base
feat/network-table-report, no merge/release. Дальше loaders, Superset matcher,
координаторский writer и пользовательский пакет.
