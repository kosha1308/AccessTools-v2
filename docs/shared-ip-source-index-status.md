# Общий IP-индекс источника

Продолжение PR19. Добавлен единый механизм для Superset/IP registry, который
сможет использовать NETWORK_ACCESS, VPN и другие шаблоны. Готовые14 проверок,
Rulebook и Excel writer не изменены. Это ещё не сетевой registry runner и не
итоговая координаторская книга.

## API

- ATIPSourceCompiler.Compile(raw, sourceId, addressFields, fingerprint) возвращает
  ATIPSourceResult (State/Index/Diagnostics). Тип RAW — SOURCEID_RAW; колонки
  адресов передаются метаданными. Никаких Excel/file операций внутри индекса.
- Index.Query(expression, callerBudget) возвращает ATIPMatchResult
  (State/Matches/SourceComplete/Diagnostics). Matches содержит исходные immutable
  строки со всеми полями и refs, один раз в порядке источника.
- Index.SourceId/Fingerprint/SourceRowCount/IntervalCount/IssueCount и DataIssues
  сохраняют происхождение и неполноту данных.
- ATIPSourceCache.GetOrCompile — session cache по SourceId, схеме адресных колонок
  и fingerprint. Повторный consumer получает тот же frozen index. Изменение
  fingerprint удаляет старую запись до compile; неудачная замена не возвращает
  устаревший индекс. До4 source/schema entries, FIFO eviction, без startup I/O.

Fingerprint здесь — явный аргумент source controller, не вычисленная индексом
свежесть файла. Controller должен включить путь, размер и mtime и учитывать
несохранённые правки. Эта привязка к фактическому файлу/UI ещё не выполнена.

## Семантика и границы

IP/CIDR/range/маски x/х и compact lists используют прежнюю интервальную математику.
Совпадение — overlap, а не полный policy-допуск. Сети не разворачиваются в IP.
Индекс сортирует интервалы по low, хранит prefix maximum high, ищет binary search
с обратным проходом до доказанного отсутствия пересечений. Очень широкие
перекрывающиеся интервалы могут потребовать длинного прохода, но он ограничен budget.

Пустая IP-ячейка допустима. Непустая неверная ячейка пропускается целиком, включая
правильный фрагмент рядом с неверным; ошибка Excel также сохраняется в DataIssues.
Другие корректные ячейки продолжают участвовать. SourceComplete=False сопровождает
каждый query result при наличии таких проблем: потребитель не должен выдавать
безусловное «не найдено» по неполному источнику. DataIssues содержит SOURCE_ROW,
SOURCE_FIELD, VALUE (с фактическим cell ref), CODE. Ошибка RAW-схемы/типа/лимит
блокирует весь compile; неверный запрос даёт ManualReview без частичного результата.

Пределы:100000 RAW rows,200000 интервалов,20m build work,65536 символов на IP-ячейку;
query использует бюджет вызывающей стороны, включая число совпадений. Переполнение
не усекает результаты. Исходные скрытые/filtered строки включены RAW reader;
видимость итоговой копии будет управляться последующим view, отдельно от индекса.

## Верификация

RED10846/2 -> GREEN10890/0. Добавлен finite brute-force corpus, маска с
непоследним x, реальная схема трёх SOURCE колонок и сохранение enrichment fields,
schema-sensitive cache. Явная компиляция PASS,11124/0. Независимый обзор не нашёл
Critical/Important замечаний; незакрытых замечаний нет.

Есть тест10001 RAW rows с поиском последней строки при budget1000: линейный обход
в него не поместился бы. Проверены вложенные/общие интервалы,0/4294967295,
dedup/source order, строгий разбор неверных ячеек, ошибки Excel, отсутствие
source-field, output/list/work limits, frozen rows, смена signature и4-entry bound.

Первый полный прогон60s остановился возле старого report writer. Диагностический
прогон90s завершился11124/0; измеренные новые группы заняли0.297s и5.414s,
включая10001 rows. Default timeout не менялся. Финальная чистая проверка
прошла на чистой копии01715ca с явным90s: compilePASS,11124/0,
report excel-0ba0b31dd217478fabb675dd6a64e252.
Контрольная предыдущая a098fea без нового индекса тоже заняла77.22s и прошла10845/0
(report excel-409c3484114a467085a937bf72f13215). Изменение времени связано с текущей
средой; тесты и default60 не менялись. Для повторения в этой среде используйте
`powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/Test.ps1 -TimeoutSeconds 90`.
Static tooling7/7, Excel gate integration9/9.

Предоставленная пустая Superset-книга прочитана ReadOnly: RAW0, индекс0, query0
и SourceComplete=True; отдельная проверка3/0. Report
excel-e89142dd5ef149bfb61881452a6162e6. SHA256 до/после:
2462F1CC10CD802BC24C5F49A440BAF0514D4A6FE9A4808B98FB7158F5BD6023.
Стандартные тесты используют только вымышленные данные.

Дальше: effective-address collection из верхних/нижних IP с дедупликацией,
использование query результатов в NET.IP_REGISTRY, уточнение hostname без потери
IP-совпадений, Rules/service context, безопасная ATVIEW_SUPERSET копия и сводка
координатора с пятью enrichment-полями. Полный пользовательский выпуск ещё не готов.
