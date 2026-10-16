# Проверки одиночных полей и отчёт

Цель полного продукта остаётся прежней: привычный сценарий заявка → Superset →
сводка координатора. Этот этап добавляет шесть локальных document checks и
неизменяемый русский отчёт. UI, полный набор NETWORK_ACCESS и внешние источники
не объявляются готовыми. Технические решения делегированы; выполняем inline.

Предпочтён отдельный проверяющий adapter над raw reader/normalizer, а не условия
формы в generic evaluator. Отчёт — универсальные immutable значения для будущего
writer/UI. Копирование старого приложения не требуется: semantic reference только
из действующих правил; реализация новая, без VBScript.RegExp и locale parsing.

## Совместимость правил

Read-only reference kosha1308/AccessTools commit d9aa9b88f9193c0d2935d99c5cc8ecca3865687c:
четыре обязательных поля инициатора (ФИО, подразделение, должность, почта);
контакт не обязательный. REQUEST_TYPE: предоставить доступ/предоставить,
отозвать доступ/отозвать, без учёта регистра и внешних пробелов. SYSTEM_IEP — Да.
Верхнее обоснование пустое → WARNING. Синтаксис email сохраняет действующий
ASCII-контракт local [A-Z0-9._%+-]+, один @, domain [A-Z0-9.-]+, точка и
минимум две буквы TLD; это совместимость, не новый RFC validator.

Срок: при отзыве игнорируется даже CVErr/пустой/невалидный. При неизвестном типе
SKIPPED, чтобы не предполагать предоставление. При предоставлении обязателен:
бессрочно либо дата не раньше явной RunDate. TODAY/локальные настройки не читаются.
Ошибки данных полей сохраняются, но не дублируются: пустая/ошибочная почта
отмечается required check, email check SKIPPED с понятной причиной.

## Явный формат даты DMY_TEXT

ATExcelDateAdapter получает дополнительный textFormat DMY_TEXT. Для строки:
строгий ISO целиком либо последняя корректная дата в тексте, например «до
30.12.2026», «с 1/10/2026 до 31-12-2026». День/месяц 1–2 ASCII цифры,
разделители . / -, год ровно 4 цифры, 1900..9999; примыкающие цифры запрещены.
Границы слова не нужны. Неверные кандидаты пропускаются, последняя корректная
дата используется независимо от более позднего неверного кандидата (legacy).
Нет IsDate/CDate(text)/regex. Числовые Value2 и ISO/DMY_DOT сохраняют прежний
контракт. Scan линейный с константной длиной кандидата, max65536 символов;
слишком длинный текст даёт RUNTIME_LIMIT_TEXT ERROR, не INVALID.
Keyword validation использует тот же canonical text, что matching, как ранее.

## API и результат

ATNetworkDocumentChecks.Run(source As ATNetworkReadResult) As ATCheckReport.
Вход — результат verified reader. Внутри normalization с DMY_TEXT/Бессрочно,
параметры опираются на подтверждённую семантику формы. Нет Excel/Range/внешних
источников/записи. Source failure останавливает запуск до checks.

ATCheckIssue immutable CheckId/Code/Severity/FieldId/CellRef/Message.
ATCheckOutcome immutable CheckId/Status/Reason, Status PASS/FAIL/SKIPPED/BLOCKED.
ATCheckReport immutable State/ScopeId/Issues/Checks/ErrorCount/WarningCount/
SkippedCount/HasErrors/Summary. ScopeId NETWORK_DOCUMENT, не полный NETWORK_ACCESS.
Collections копируются, объекты frozen после Initialize; повторная Initialize
запрещена. State OK означает выполненный запуск в указанном scope, не отсутствие
бизнес-ошибок; HasErrors считает severity ERROR и State ERROR.

Детерминированный порядок: REQUIRED_FIELDS, INITIATOR_EMAIL, REQUEST_TYPE,
SYSTEM_IEP, VALIDITY, TOP_JUSTIFICATION. Неизвестный тип пропускает VALIDITY;
остальные независимые проверки продолжаются. Issue всегда содержит адрес
исходной ячейки; шаблон/технический отказ допускает пустой адрес.

Всего шесть проверок. Ошибки Excel в обычном поле → одна ERROR issue; только
EXCEL_CELL_ERROR принимается как data failure. Неверный тип/пропущенное поле/
пустой ref/внутренняя диагностика → State ERROR, одна русская blocked issue,
ни частичных checks, ни положительных исходов. Валидация схемы до проверок.
Для отзыва ошибка срока не становится бизнес-ошибкой; при этом системный сбой
нормализации всё равно не маскируется. Максимум1000 issues/checks в generic report.

## Приёмка

TDD real Excel: новые даты/free text/bounds, шесть проверок на blank/valid/multi-error
forms; request aliases/case/whitespace, required contact exception, warning severity,
email accepted/rejected corpus, date today/past/future/revoke/unknown, CVErr and
structural blocking, counts/copied collections/frozen objects/source close/re-run.
Независимый обзор whole branch, static7/7 gate9/9, clean-checkout compile/full suite,
реальный эталон ReadOnly с неизменным hash. Draft PR base предыдущий этап, без merge.
