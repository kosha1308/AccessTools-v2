# Этап 2B.2 — правила по строкам

На 2026-10-16 реализовано и проверено в Excel: Compile VBAProject PASS,
**9679 PASS / 0 FAIL**. Static tooling7/7, Excel gate9/9; справочник не изменён.
Evidence чистой копии записывается в progress.md.

## Что работает

- Компилятор сохраняет FieldId/TableId/RowAlias и scope именованного определения.
  Runtime не выбирает подходящую строку по имени поля повторно.
- Неизменяемые строки с RowId и metadata ячеек; типизированные коллекции строк.
  Снимки изолированы от изменения исходного списка/строки.
- ALLROWS, ANYROW, NOROWS, FILTER, COUNT и EXISTS. Вложенные aliases и сравнения
  между таблицами поддержаны, включая IP-предикаты.
- Пустые коллекции: ALLROWS/NOROWS TRUE, ANYROW/EXISTS FALSE, COUNT0.
- DOCUMENT определения/defaults не захватывают caller row scope и могут
  использовать свой alias с тем же именем; ROW определения получают caller scope.
- Non-OK predicate сохраняет состояние/первичную диагностику. FILTER не выдаёт
  частичный результат. Short circuit пропускает оставшиеся строки с trace marker.
- Trace содержит RowId/TableId/FieldId/CellRef, исходный span и SourceId без raw values.
- Количество строк ограничивается до клонирования; отдельные поля валидируются
  только при чтении. COUNT/EXISTS и decisive skip не читают большие/ошибочные поля.

## Проверки

Metadata binding: missing API RED compile →9571/0. Row model: missing class RED →9580/0.
Traversal/trace: missing trace fields RED; после реализации шесть прежних negative
fixtures для unsupported rows стали давать ожидаемый FIELD_MISSING, поскольку
операции теперь поддержаны. Эти fixtures перенесены на такие же row operations
над пока неподдержанным SOURCE; новым операциям соответствуют положительные tests.
Результат9660/0 включает nested/correlated IP, empty laws, non-OK states, missing/type/
INVALID values, skipped rows/cell trace, depth/steps/rows/trace limits.

Независимое статическое ревью не выявило конкретных ошибок реализации. Его три
пункта покрытия закрыты тестами: oversized unused TEXT/list, DOCUMENT default
с собственным квантификатором/alias, missing row field и повторное использование
после nested error/decisive skip. Полный результат **9679/0**.
Замечаний ревью без решения не осталось.

## Граница

Это вычисление на переданных in-memory строках. Чтение формы Excel, объединённые
ячейки/пустые строки заявки, источники, Superset, Rulebook, координаторская сводка
и пользовательская надстройка ещё не подключены. DISTINCT_COUNT, collection
equality, SOURCE/LOOKUP и safe patterns остаются явно unsupported.
Критерий первого использования остаётся полным сценарием из first-use-network-access.md.
