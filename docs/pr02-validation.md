# PR-02 — реализация и проверка

База: main `f735d6363550ce088f71a771a6a0442e82b11374` после PR #33. Ветка `feat/pr-02-schema-validator`. Контракт и предложения для согласования — [ADR-008](adr/008-schema-storage.md). Прочитаны ТЗ v0.15 и материалы PR-01; Google Drive, AccessTools 1 и архивный пилот не изменены.

## Состав

Пять производственных файлов в `src`: CAtPackage, CAtDefinition, modAtModel, modAtValidation, modAtXml. Generic Kind-схема представляет десять требуемых моделей; дополнительный Expression — только deferred ссылка. Два синтетических XML используют общий pipeline; OP_TEST — только тестовая запись доверенного каталога. Нет AST-интерпретатора, библиотеки операций, бизнес-проверок, UI, Excel Writer, миграции Rulebook или публикации.

## Реально выполнено 19.10.2026

- `Probe-MSXML.ps1`: PASS в Windows PowerShell x64 и x86 — создание MSXML6, Unicode, защитные свойства и отказ DTD. Это не запуск 32-битного Excel.
- `Run-Tests.ps1`: импорт исходников в отдельную пустую книгу и исполнение VBA в **Excel 16.0 build 17932, Windows (64-bit) NT 10.00, Office64**. **77 assertions PASS, 0 FAIL**. Сводка `PASS|77|0||Office64`. Это реальное выполнение VBA, не статический анализ.
- RED bootstrap до реализации API; затем ошибки лимита MSXML и CR/LF выявлены реальными запусками. Регрессии ревью: 3 FAIL до исправления → полный PASS после исправления.
- Независимое source review проверило XML/save и validator. Исправлены несовпадение владельца Template→Rule, чужая Rule dependency, Table/Column refs у A1 locator. Ревью не заменяет выполнение; повторное ревью не нашло блокирующих замечаний.

| Обязательный сценарий | Автоматическое доказательство |
| --- | --- |
| Оба примера | Загрузка simple и equipment-return, общий AtValidate, сохранение, повторная загрузка |
| Round-trip | Равенство повторно сериализованной модели: ID/версии/значения/порядок/ссылки; отдельно Unicode, `<&`, CR/LF/TAB, literal percent sequences |
| Перемещение поля | B8 → $XFD$1048576 и DisplayName change; ID FIELD_NAME в Rule reference прежний, ошибок нет |
| Обязательные/отсутствующие ссылки | MISSING Field, удалённый Locator, неверная exact version, неправильный владелец/Column Source |
| Дубли и ID | Повторный FieldId, неверная грамматика, duplicate refs и XML property |
| Зависимости | Self-cycle, цикл из двух правил, корректный DAG, неизвестная dependency, чужой Template |
| Тип/версия | DataType=Bogus, Text→Int64 при expected Text, Version=1.x, SchemaVersion=2.0.0/9.0.0 |
| Операции | Неизвестные OperationId и exact OperationVersion при внешнем каталоге |
| Locator | A0/XFE1/обратный диапазон, DefinedName и запрет A1-name, offsets и конфликт настроек, TableColumn и две Column |
| Published | Отсутствующая публикационная metadata, корректная UTC date, невозможное 30 февраля; deferred-AST диагностика остаётся |
| Повреждённая конфигурация | Некорректные XML text и файл; неизвестные атрибут/элемент/property, запрещённая PI и DTD |
| Ограничения | Text >524288 символов, parser depth >32 |
| Ошибка записи | Existing target → FILE_EXISTS; отсутствующий parent → IO_WRITE. Предыдущий файл сравнивается **побайтово**, затем загружается снова |
| Ошибка чтения / памяти | Неизвестный путь → IO_READ; Nothing model → MODEL_INVALID |

Тестовые результаты, assertion labels, версия/сборка Excel, hashes исходников, ID собственного процесса и временные XML записываются в `artifacts/run-<guid>/`; каталог игнорируется Git. Логи не содержат пользовательских книг. Не коммитить локальные артефакты или Office binary. Исходники ASCII для воспроизводимого импорта в VBE; Unicode в тесте строится через ChrW, а XML — UTF-8.

## Воспроизведение / ручная приёмка

Требуется Windows + desktop Excel с VBA. Тестовый runner использует автоматизацию и доступ к VBProject; если действующая политика запрещает его, запуск завершается ошибкой — runner не меняет Trust Center/реестр. Выполнять только доверенный код этого репозитория. Тестовые процессы PowerShell получают ExecutionPolicy Bypass только на время запуска, машинная политика не меняется.

Из корня checkout:

```powershell
./tests/Probe-MSXML.ps1
./tests/Run-Tests.ps1
```

Если PowerShell запрещает локальные скрипты, можно запустить конкретный доверенный runner через `powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tests/Run-Tests.ps1`; не менять глобальную политику. В sandbox может потребоваться разрешённый запуск Excel вне sandbox; это ограничение среды, не дефект XML. Runner создаёт отдельный скрытый Excel, пустую книгу, не открывает рабочие книги; закрывает только собственную книгу/приложение. При timeout 90 s завершает только процесс с сохранёнными ID + StartTime и свой worker.

При ручной приёмке: проверить PASS и битность в result.txt; сравнить fixtures с созданными XML; проверить assertion labels про move/rename и previous bytes. Для Office32 повторить весь Run-Tests на машине с установленным **32-битным Excel**. Вариант без разрешения VBProject: создать свою тестовую книгу, импортировать два `.cls`, три `.bas` и `tests/AtSchemaTests.bas` в VBE, запустить `AtRunSchemaTests` с путями к checkout и отдельному пустому scratch-каталогу через Immediate. Не импортировать в рабочую заявку. Код тестов не сохраняет книгу.

## Не проверено / границы результата

Не выполнены Office x86 VBA, другие версии Office/Windows и организационные политики, ручной Admin/Runtime UI, настоящий Excel Resolver/Writer и пользовательские заявки. Нет теста power-loss, сетевого share, ACL/reparse attack или конкурентного процесса в момент rename; сохранность предыдущих файлов обеспечивается отсутствием операций записи в них, но crash-durability не заявляется. Тесты не подтверждают бизнес-паритет AccessTools 1 или полноценную публикацию.

Полный AST type checking/operation signatures и вычисление остаются PR-03/следующими этапами по отдельному заданию. Сохранение неполного Draft в Admin, publish/digest/active pointer/rollback, фактическое чтение Source и разрешение Locator не реализованы. Принять Published metadata как структурно корректную не означает доверять файлу или разрешать запуск.

На согласование: ADR-008 (формат и escaping, generic records, exact версии, append-only файлы, размещение Draft/Published). Новых решений по бизнес-семантике нет. PR-03 не начат.
