# AccessTools 1: контракт Excel-вывода и пользовательского результата

## Область и статус доказательств

Аудит исходного кода, а не подтверждение работы Excel. Прочитан `AGENTS.md`. Исторический baseline **main d9aa9b88f9193c0d2935d99c5cc8ecca3865687c**, последняя рабочая версия по заданию **integration 58d8789d1e60860c96cb35ba0e780a6424510130** репозитория `kosha1308/AccessTools`. Обе версии представлены отдельно; выбор иной бизнес-семантики этим аудитом не производится. Ссылки ниже закреплены на immutable SHA. Исходные снимки прочитаны как UTF-8; реальные адреса, домены и инструкции справочников не перепубликуются: точный текст следует читать в указанном источнике.

**Код изучен статически; Excel manual не выполнен; VBA/COM/release tests не запускались в этом аудите.** Наличие regression-функций и исторические PASS в ТЗ не доказывают запуск на текущем окружении.

ТЗ v0.13 (проверенная ревизия с фиксацией 58d8789 от 2026-10-16, приложение L) нормативно требует финальную двухслойную модель, 11 полей сводки, фактические строки Superset, неизменность исходного источника. Дополнение об ALLROWS требует верхние + нижние IP. Это внешний документ, а не файл immutable GitHub: [источник ТЗ v0.13](https://docs.google.com/document/d/1FHhbIQIzjlqrntIJKUlOOtDqn0EDctEU/edit). Ссылка не фиксирует версию; проверенная здесь ревизия зафиксирована выше; снимок в поставку документа не входит. Это ограничение трассируемости нельзя скрывать под GitHub-ссылкой на код.

## 1. Действующий путь и контракты публикации

Manifest выбирает `ApplyNetworkProblemCatalogV5`, `NetworkRequestAnnotationIssuesConfigured`, RULEBOOK как источник пользовательских сообщений и действие `OpenNetworkResultViewSnapshot` ([manifest, main](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/config/templates/network_access/manifest.ini#L1-L17)). Старые Cached/Audience функции присутствуют в коде, но их существование не делает их главным transform для этого manifest.

`ExecuteTemplatePipeline`: очистка старых отметок → проверки → transform → фильтрация аннотаций → writer → summary → отдельный фильтр окна → показ; context сохраняет полные Issues/Results. `applyAnnotations=False` выключает и очистку, и запись. Ошибка transform логируется, ошибка annotation filter возвращает **все** issues: это расширяющий публикацию fallback, не гарантированный fail-closed ([pipeline, main](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modTemplateRunner.bas#L6-L150), функции `ExecuteTemplatePipeline`, `ApplyTemplateIssueTransform`, `TemplateAnnotationIssues`).

| Слой | Решение / данные | Следствие |
|---|---|---|
| Примечание заявки | AnnotateRequest, затем SheetName/Address; Message после transform | Пользовательский шаблон сообщения не является отдельным текстом примечания |
| Цвет заявки | максимальный Severity по writer-группе | ERROR > WARNING > INFO |
| Окно результатов | ShowInResultsWindow | Скрытие не удаляет issue из coordinator context |
| «Все проблемы» | IncludeInUserReport, видимые строки, UserMessageTemplate и UserReportBaseMessage | Не тождественно полному Message |
| Superset | инфраструктурные CheckId и захваченные coordinator issues | Отдельный канал; не эквивалентен флагу AnnotateRequest |

Основания таблицы: [writer](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L175-L284), [V5 audience](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemRuntimeV5.bas#L9-L123), [direct transform](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemDirect.bas#L28-L113), [routing](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkIssueRouting.bas#L6-L18).

## 2. Тексты, severity, аудитория: точный порядок

`ApplyNetworkProblemCatalogV5` сначала готовит direct schema/routing view, вызывает V4/direct, затем принудительную аудиторию окна и добавление ссылок на ячейки только для user report. На ошибке повторяет V4/audience/references под подавлением ошибок ([V5](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemRuntimeV5.bas#L9-L33)).

`ApplyNetworkProblemDirectCore` читает лист «Проблемы»: «Ключ», «Сообщение пользователю», «Сообщение координатору», «Уровень сообщения», «Ошибка пользователя», «Подсвечивать в заявке», «Активно». Выбор по исходным CheckId/Severity/Field: **CheckId|severity|field → severity|field → severity|* → *|***. Поздняя одинаковая запись ключа замещает предыдущую в dictionary. `Активно=Нет` удаляет issue целиком, затем пересчитываются Results. Непустой coordinator template заменяет Message, user template сохраняется отдельно; пустой user template становится `{message}`, базой остаётся оригинальный Message проверки. Только затем применяется новый Severity и два независимых флага публикации ([DirectCore/loader/lookup](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemDirect.bas#L4-L267)). Точные default-тексты и раскрываемые подстановки — в `ExpandDirectProblemTextFromBase`, а не в этой копии документа.

Severity принимает «ошибка»/error, «предупреждение»/warning/warn, «информация»/info; нераспознанное значение не заменяет исходный уровень. Defaults исключают инфраструктуру и согласование из пользовательских проблем, но согласование сохраняет отметку заявки; fallback сохраняет исходный IncludeInUserReport, если он задан ([fallback/defaults/severity](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemDirect.bas#L270-L390)). Это нельзя свести к одному Audience enum без потери поведения.

После каталога V5 скрывает NET.IP_REGISTRY полностью; NET.IP_EXCEPTIONS/NET.HOSTNAME_RULES — только INFO. Text-group INFO принудительно снимает все пользовательские флаги и переводит CheckId в NET.IP_EXCEPTIONS для coordinator pipeline. UiErrors/UiWarnings/UiInfos уменьшаются, реальные счётчики и CountsToken сохраняются ([ApplyNetworkResultsWindowAudience](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemRuntimeV5.bas#L35-L123)).

«Все проблемы» строится по отображаемым строкам; INFO по умолчанию исключён, нормализованные одинаковые итоговые строки дедуплицируются. SKIPPED не становится замечанием инициатору, но запрещает ложный окончательный «замечаний нет». Приоритет текста: direct UserMessageTemplate → legacy severity|field → severity|* → *|* → `{message}`; у legacy сначала ShortText, затем Text. UserReportBaseMessage имеет приоритет над первым абзацем details ([BuildTemplateUserReportText/BuildUserReportItemBody/FindUserReportTemplate](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/modUserReport.bas#L8-L179)). Ссылки на адреса, группировка по UiGroupKey/catalog/message и лимит 220 символов формируются только для пользовательского текста ([ApplyNetworkUserReportCellReferences](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemRuntimeV5.bas#L125-L245)).

## 3. Writer, merge, коллизии и сохранность

| Условие | Запись / очистка | Ограничение |
|---|---|---|
| Пустые SheetName или Address | Issue не попадает в writer groups | Наличие в результатах не гарантирует отметку |
| Несколько issues на одном literal SheetName\|Address | Одно примечание; пункты в порядке issues; общий максимальный цвет | Нет семантической дедупликации Message |
| Разные адреса одной MergeArea | Группировка выполняется **до** нормализации к merge anchor | Возможны последовательное замещение блока/цвета; глобальный max по MergeArea не гарантирован |
| Обычная ячейка | Условное форматирование с формулой-маркером AccessTools.SEVERITY, SetFirstPriority | Не меняет Interior напрямую, но меняет приоритет CF |
| MergeArea | Примечание на первом cell; исходная заливка в Shape.AlternativeText; прямой цвет всей области | Снимок из anchor, не полноформатный snapshot всей области |
| Metadata не записалась | Без исходной заливки разрешён direct fill; иначе CF fallback | Ошибка скрывается SafeExit |
| Пользовательское примечание | Старый AT-блок удаляется, пользовательский текст + два CRLF + новый блок | Существующий размер сохраняется; новое примечание AutoSize |
| Блок AT без END | RemoveAccessToolsBlock возвращает исходный текст | Повреждённый блок сохраняется; повтор может добавить ещё один |
| Очистка | Обходит Comments с NOTE_BEGIN; убирает CF с AccessTools. и блок; пустой остаток удаляет Comment | Не сканирует все UsedRange CF; осиротевшие CF без note могут остаться |
| Старая merge-отметка без metadata | Убирает solid fill, только если цвет совпал с текущим configured ERROR/WARNING/INFO | Пользовательская заливка того же цвета неотличима; при смене конфигурации старый цвет может остаться |

Основания всех строк: [Clear/Resolve/Delete/Remove](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L4-L173); [Apply/HighestSeverity/Formatting](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L175-L284); [Save/Restore metadata](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L286-L450); [Append/Build note](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L453-L527).

В note перечислены номер, Severity и Message каждого issue. В конце один «Инструкция по сервису» с первым непустым InstructionUrl. Это не объединение всех URL ([BuildAccessToolsNoteBlock](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modAnnotations.bas#L501-L527)). Формулы и Value2 заявки writer не присваивает; это ограниченное доказательство по writer, не обещание полной неизменности всех модулей. Cleanup/Append/Formatting подавляют ошибки, а значит успешный возврат не подтверждает фактическое восстановление защищённого листа. Очистка происходит **до** проверок, поэтому при последующем сбое нет транзакционного восстановления предыдущих результатов ([pipeline](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/core/modTemplateRunner.bas#L6-L33)).

## 4. Рабочий процесс пользователя

Источники открываются через настроенный template/source, меню имеет отдельные source/config/maintenance действия ([OpenTemplateDataSource/EnsureConfiguredDataSource/OpenEditableFile](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/modMenu.bas#L379-L484)). Обычная проверка запускает UI и annotations; regression entry point выключает их. Перед новым анализом сбрасывается прошлый result snapshot и coordinator state ([RunNetworkAccessCheck/RunNetworkAccessAnalysis](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkEngine.bas#L6-L197)).

Форма: «Все проблемы», «Повторить проверку», «Очистить отметки», Copy/Go и template action. Go активирует связанную книгу/лист и Application.Goto по issue Address; если книга закрыта — сообщение, если адрес недоступен — обработанная ошибка. Repeat ищет именно книгу окна, активирует её и запускает EntryPoint manifest; busy запрещает действие. Clear снимает annotations, не запускает проверки заново и не доказывает отсутствие issues ([ResultsNavigate](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/modResultsWindow.bas#L399-L446), [ResultsClearAnnotations/ResultsRepeatCheck](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/modResultsWindow.bas#L579-L628), [form handlers](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/frmResults.code.txt#L329-L433)).

Отдельное автоматическое сохранение исходной заявки в pipeline отсутствует. Сохранение результата Superset — другой путь: CaptureNetworkResultViewSnapshot сохраняет временный ATVIEW, пытается сменить read-only на read-write, при оставшемся read-only только логирует warning. Snapshot path всё равно запоминается. Следующий reset закрывает snapshot без сохранения и удаляет временный файл; пользовательский долговременный результат требует сохранения отдельной копии средствами Excel ([Reset/Capture/Open snapshot](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkResultView.bas#L7-L126)).

## 5. Superset: main и integration не взаимозаменяемы

В main ATVIEW строится лениво при «Открыть Superset», может существовать и при нуле найденных строк. Порядок: pending interval view либо empty coordinator view → hostname enrichment → conflict warnings → coordinator comments → AccessTools_Итог → snapshot. На время сборки отключаются repaint/events/auto calculation, сохранённое состояние восстанавливается; ошибка отдельных enrichment-процедур, которые подавляют исключения, не гарантирует отказ всей сборки ([OpenNetworkResultViewSnapshot](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkResultView.bas#L98-L246)).

### Схема main

| Лист / слой | Поля и смысл |
|---|---|
| Отфильтрованные строки исходной выгрузки | Исходные колонки сохраняются; `AccessTools: IP из заявки` связывает строки с запросом |
| Rulebook enrichment | `AccessTools: Сервис из правила`, `AccessTools: Группа назначения`, `AccessTools: Требуемое согласование`, `AccessTools: Комментарий / инструкция` |
| Coordinator enrichment | `AccessTools: Комментарии координатору` |
| `AccessTools_Итог` | `IP / сеть из заявки`, `Статус Superset`, `Уровень`, `Комментарии координатору` |

Точные источники: [ShowPending/write request links](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modCheckRegistryIntervals.bas#L240-L315), [EnrichNetworkRegistryViewWithHostnameRules](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modRegistryViewEnrichment.bas#L10-L170), [ApplyNetworkCoordinatorCommentsToRegistryView/BuildNetworkCoordinatorSummarySheet](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkCoordinatorOutput.bas#L304-L449).

Summary содержит по entry строке; статусы «НЕ ПРОВЕРЕНО», «НАЙДЕН», «НЕ НАЙДЕН В SUPERSET». Для глобальных сообщений добавляется «Общее». Это наблюдаемый main, но **не финальный контракт приложения L ТЗ13** ([BuildNetworkCoordinatorSummarySheet](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkCoordinatorOutput.bas#L358-L449)).

### Финальная схема integration (соответствует форме приложения L)

Integration добавляет после старой сводки `FinalizeNetworkCoordinatorSummaryV2`, затем `MoveRegistryDetailToCoordinatorSummary`, и только затем snapshot ([порядок](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkResultView.bas#L176-L190)).

| № поля | Заголовок `Сводка координатора` | Источник / агрегация |
|---|---|---|
| 1 | № | Номер entry с 1 |
| 2 | IP / сеть из заявки | DisplayText старой сводки, исключая «Общее» |
| 3 | Статус Superset | Статус старой сводки |
| 4 | Найдено строк | Число видимых совпавших registry rows для entry |
| 5 | Сервис | Rulebook output, deduplicated set |
| 6 | Инструкция | Требуемое согласование / извлечённый опубликованный output, deduplicated set |
| 7 | Имя хоста (Superset) | Единственное уникальное непустое значение; несколько → пусто |
| 8 | Сервис (Superset) | Все уникальные значения фактических видимых строк, по новой строке |
| 9 | Полное имя ИС (Superset) | Аналогично |
| 10 | СКУФ ГРУППА ТП (Superset) | Аналогично; обычная колонка Superset |
| 11 | КОМАНДА ПОДДЕРЖКИ (Superset) | Аналогично; обычная колонка Superset |

Основания: [CoordinatorV2ReadOldSummary/Aggregate](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L47-L165), [RewriteSummary](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L233-L315), [MoveRegistryDetailToCoordinatorSummary](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L4-L210), [dedup/unique helpers](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L302-L400).

Детальный лист переименовывается в `Superset`, если имя свободно. `AccessTools: IP из заявки` превращается в `№ из сводки` (одна строка может иметь несколько номеров), скрытые строки получают пустую связь; временные Rulebook/coordinator-колонки удаляются справа налево. Это преобразования временной копии, не исходного справочника ([CleanRegistrySheet/ReplaceRequestTextsWithNumbers](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L167-L231), [точный whitelist удаления](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L453-L464)).

**Риск потери глобальных сообщений:** V2 reader исключает строку «Общее», затем RewriteSummary очищает старый лист. Это наблюдаемая потеря прежнего канала summary, а не разрешение удалить бизнес-информацию в v2. Нужен отдельный выбор владельца о сохранении глобальной coordinator-информации ([ReadOldSummary](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L75-L102), [RewriteSummary](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L233-L254)).

## 6. Автотесты, ручные сценарии, неизвестные

Код содержит regression hooks, **не запущенные здесь**: [RunNetworkProblemDirectRegression](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemDirect.bas#L426-L444), [RunUserReportSkippedOnlyRegression](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/ui/modUserReport.bas#L73-L88), [RunNetworkCoordinatorSummaryV2Regression](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkCoordinatorSummaryV2.bas#L505-L536), [RunRegistryDetailSummaryRegression](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L453-L500). Эти hooks не заменяют визуальные проверки Excel.

Все следующие synthetic cases **предложены, не выполнены**; данные должны быть искусственными (например 192.0.2.10, host-a.example.invalid), без реальных исходников.

| Case | Подготовка / действие | Проверяемый результат |
|---|---|---|
| W01 | Пользовательский note до/после валидного AT-блока; run → repeat → clear | Пользовательский текст и размер note сохраняются, нет дублирования AT-блока |
| W02 | Merge с исходной patterned fill, ERROR+INFO в одном адресе | ERROR; после clear исходный fill; Value2/formula неизменны |
| W03 | Два разных адреса одной MergeArea, ERROR затем INFO и обратный порядок | Зафиксировать коллизию групп; не выдавать желаемый global max за доказанный |
| W04 | Необычный CF, повреждённый note marker, удалённый note, повреждённые AlternativeText | Измерить реальные остаточные CF/fill и сохранность пользовательского текста |
| W05 | Protected sheet / недоступная Shape.AlternativeText | Проверить fallback и фактическую частичную запись, отсутствие ложного UI успеха |
| M01 | Один CheckId с direct и wildcard строками; изменить severity, Active, два audience flags | Direct приоритет, выключенный issue исчезает; окно/annotation/report независимы |
| M02 | SKIPPED-only + INFO-only + одинаковые тексты с разными адресами | Нет ложного полного успеха; INFO default; адреса в user report, без изменения Message |
| U01 | Две книги, переключить ActiveWorkbook, Go/Repeat/Clear из окна первой | Операция относится к книге результатов; закрытая книга даёт обработанное сообщение |
| S01 | Явный верхний IP + текстовая группа + нижний IP + duplicate + not-found | Уникальные IP доступны в summary; not-found остаётся; результаты сверять с ТЗ13, отдельно main/integration |
| S02 | Два пересекающихся entries и одна фактическая строка Superset | В `№ из сводки` оба номера; count и enrichment корректны |
| S03 | Несколько hostname, duplicate service, пробелы в новых заголовках | Hostname пусто при неоднозначности; поля 8–11 уникальны и многострочны |
| S04 | Глобальное INFO coordinator + not-found + отсутствующий source | Проверить потерю «Общее», НЕ ПРОВЕРЕНО, пустые source-derived поля |
| S05 | Закрыть ATVIEW; открыть повторно; затем новая проверка | Snapshot работает до reset; отдельная сохранённая копия сохраняется; исходный source hash неизменен |
| S06 | Ошибка сохранения / read-only / ошибка enrichment | Зафиксировать частичный результат и warning; не считать snapshot path доказательством успешного Save |

Неизвестные до ручной приёмки: версия Excel и поведение legacy Comments/merged CF, защищённые листы, богатое форматирование notes, конкурирующее изменение источника, фактическая сохранность пользовательских стилей при legacy cleanup. Выше приведены статические риски из конкретных ветвей кода, а не утверждение о воспроизведённых дефектах. Долговечность ATVIEW, аудитория, потеря «Общее» и конфликт адресов одной MergeArea должны быть явно решены перед реализацией parity; новые бизнес-правила этим документом не вводятся.
