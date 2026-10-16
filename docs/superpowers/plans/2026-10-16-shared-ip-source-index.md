# Общий индекс IP источника

Base91a079a/PR19. Следующий узел Superset-потока, общий для NETWORK_ACCESS/VPN.
Готовые14 checks, Rulebook engine и Excel writer не меняются.

## Контракт

ATIPSourceCompiler.Compile(raw, sourceId, addressFields, fingerprint) возвращает
ATIPSourceResult с immutable Index либо Error/Diagnostics. RAW — типизированные
строки любого источника; sourceId задаёт ожидаемый тип SOURCEID_RAW.
Superset задаёт три SOURCE_* колонки; универсальный индекс не знает шаблона заявки.

Индекс хранит интервалы IPv4/CIDR/range/mask/compact list и исходные позиции строк.
Никакого разворачивания сетей. Сортировка low + prefix maximum high обеспечивает
binary search и остановку поиска без полного обхода источника на обычном запросе.
Query(expression, callerBudget) возвращает ATIPMatchResult:
State/Matches/SourceComplete/Diagnostics. Совпадение — математический overlap,
не full policy coverage. Каждая исходная строка возвращается один раз в порядке
источника, со всеми полями и refs. Неверный запрос — ManualReview, без salvage.

Повреждённая непустая IP-ячейка не разбирается частично: вся ячейка пропускается,
но другие корректные ячейки/строки остаются индексированы. Каждая такая ячейка
видна в DataIssues (field/value/code/ref); SourceComplete=False во всех query
results. Это не разрешает выдавать окончательное «не найдено» без предупреждения.
Ошибки структуры RAW/неверные типы/лимиты блокируют весь compile без partial index.
Пустые IP-ячейки допустимы. RAW уже включает скрытые/отфильтрованные строки.

Build limits:100000 rows,200000 intervals,20m work,65536 chars/cell.
Query использует callerBudget, включая предел результатов; нет тихого усечения.
Fingerprint — явная непрозрачная сигнатура файла/схемы от source controller;
обязан включать путь/размер/mtime. Файлы здесь не открываются. UI-слой актуальности
и расчёт fingerprint по файлу остаются следующей интеграцией.

ATIPSourceCache.GetOrCompile: session cache по SourceId/набору колонок/fingerprint.
Одинаковая сигнатура возвращает тот же frozen index двум шаблонам. Изменение
fingerprint удаляет прежний индекс до compile; ошибка не возвращает старые данные.
До4 source/schema entries, FIFO eviction, без startup I/O.

## Выполнение inline

1. Meaningful RED: list/CIDR/range boundaries, 0/32/wholeIPv4, compact source,
   пересекающиеся источники, строгие malformed/ExcelError/dataIssues, dedup/order,
   immutable rows, query limits and source schema. Реализовать compiler/index/
   guarded result models, bottom-up sorts.
2. Cache tests RED/GREEN: NETWORK/VPN reuse, fingerprint/schema invalidation,
   invalid replacement never stale success, bounded eviction. Scale10001 rows
   with low query budget proves index search avoids full scan.
3. Excel entire regression, provided empty SupersetReadOnly, independent review
   fixes RED/GREEN, static7/gate9/clean60s. Docs/readiness, draft PR over19.

## Review focus

Prefix maxima with nested intervals; no lost duplicates at0/4294967295; partial
bad cells cannot disappear; empty/invalid query never successful no-match; same
signature reuse versus changed/bad source invalidation; all output work bounded.
Network effective-address collection, hostname refinement, registry check,
Superset copy and coordinator summary remain next composition step.
