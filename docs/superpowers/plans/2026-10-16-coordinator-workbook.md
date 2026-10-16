# Coordinator workbook Implementation Plan

Goal: отдельная двухлистовая книга из frozen registry plan, без изменения источников.
Base:2ece9de/PR21. Inline execution; existing design from ТЗ appendixL and PR21
next-stage contract, user repeatedly delegates technical implementation.

1. ATCoordinatorWorkbookWriter.Create(registry) As Workbook builds a normalized
   ATVIEW_SUPERSET value snapshot of matched original eight business columns plus
   exact SUMMARY_LINK. Export-only unnamed index and unrelated sheets excluded.
   Rows dedup by physical source RowId, sorted ascending; links retain every ID.
   Max10000 unique source rows matches existing enricher physical-row budget;
   max100000links,32767celltext; overflow stops without partial book. Source
   errors retained as Excel error; relevant enrichment errors block final book.
2. First sheet Сводка координатора headerrow5, second Superset headerrow5.
   Summary base:№,IP/сеть из заявки,Статус Superset,Найдено строк,Сервис,Инструкция;
   five existing enrichment fields plus Ячейки заявки. Blank rule fields clearly
   labelled pending in banner, no fabricated annotations. Use existing enricher
   on actual output Superset ordinary columns; multiple services newline,
   hostname only unambiguous. No-hit and unconfirmed entries remain.
   Source snapshots immutable, output text escaped/no formulas/external links,
   events restored, created book closed on any failure. Never save implicitly.
3. TDD initial failing writer scenario: actual two lowerIP, overlapping /24 and
   singleton, shared sourcerow linked to both, no duplicate rows, exact1vs10,
   safe formula-like sourcevalue, missing retained, five fields, empty/error.
   Test output filtering with existing enricher; snapshotindependent. Extra
   tests sourceCVErr, limits, failurecleanup/applicationstate, inputblocked.
4. Full Excel explicit90s, static7/gate9, fresh independent review, cleanclone,
   sample workbook + visualQA, docs/readiness, draftPR basePR21, no merge/release.

Ruling: the workbook uses eight recognized source business columns as values;
no unknown columns or VBA/formulas copied. This follows RAW contract; preserving
export index or arbitrary extra columns is not advertised. Source freshness,
Rulebook service/hostname integration, user-run/file saving remain next stage.
