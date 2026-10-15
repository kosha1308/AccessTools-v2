# AccessTools-v2 — проектное решение первого этапа

Основание: AccessTools_Конструктор_проверок_ТЗ_реализация_v0.13, приложения N–P. Дата: 2026-10-15.
Статус: проектное решение согласовано пользователем 2026-10-15 сообщением «ДЕЛАЙ ДАЛЬШЕ».

## Цель и границы
Создать самостоятельно тестируемую основу языка правил AccessTools-v2. Результат этапа — разбор выражения, проверка идентификаторов и типов, диагностика и каноническая сериализация. Выполнение бизнес-проверок, Admin UI, публикация Runtime и изменение AccessTools 0.2.x в этот этап не входят.

## Выбор реализации
Рекомендуется VBA как язык поставляемого ядра: он соответствует Excel Runtime без Python/Node/DLL. Тестовый runner на Windows запускает Excel через PowerShell COM; PowerShell используется только разработчиком для импорта, сборки и тестирования. Production Runtime не зависит от этих инструментов.
Альтернатива — сначала написать эталонный parser на Python/TypeScript — ускоряет отдельные тесты, но создаёт вторую реализацию и риск расхождения при переносе в VBA. Для первого этапа выбирается единое VBA-ядро.

## Компоненты
- schema: версии DSL/Schema и контракты сущностей.
- registry: типы, операторы и функции с signatures, semantic version, статусом, evaluator binding, примерами и test contract.
- diagnostics: стабильный код, категория, severity, исходная позиция и длина, понятное сообщение.
- lexer: токены со spans; строгие identifiers, strings, numbers и delimiters.
- parser/AST: literals, references, calls, arithmetic, relations, AND/OR/NOT, collections и scoped aliases.
- serializer: одна детерминированная каноническая запись AST.
- resolver/type checker: разрешение IDs и scopes, signatures, запрет неоднозначности и циклов Named Expressions.
- validation limits: размер входа, количество токенов/узлов, глубина выражений и коллекций.

Ядро не знает названий Excel-заголовков NETWORK_ACCESS и не открывает внешние источники. Metadata и тестовые контексты передаются через явные контракты. Реестр проверяет синтаксическую допустимость и типы; подключение реального evaluator является следующим этапом.

## Контракты
Entity IDs: [A-Z][A-Z0-9_]*, canonical uppercase. Field references: @FIELD_ID; row reference: @t.FIELD_ID; named expressions: $EXPRESSION_ID(...). Aliases имеют лексическую область. Keywords регистронезависимы.
Строки заключаются в двойные кавычки с документированным escaping; числа используют десятичную точку; даты — DATE("YYYY-MM-DD"); списки — [443,8443].
Короткая и полная запись приводятся к одному AST при явно заданном contextual FieldId. Legacy import policy задаётся metadata и не угадывается.
Приоритет: calls/скобки, unary +/-, */%, +/-, relations, NOT, AND, OR.
EvalResult<T> заранее описывается как Value + DataType + State + Diagnostics; State: OK/ERROR/SKIPPED/MANUAL_REVIEW. TRUE/FALSE допустимы только при OK. EMPTY, MISSING и INVALID различаются. Реальное вычисление и propagation policy будут реализованы перед этапом evaluator и должны иметь завершённую спецификацию до перехода к нему.
Set semantics: ANY/OVERLAPS — пересечение; COVER — A subset R; CONTAINS — R subset A; EXACT — равенство; NONE — непересечение.
Actions отделены от выражений. Произвольный код и I/O из DSL запрещены.

## Порядок реализации
1. Минимальная структура репозитория, README, версии схемы, runner и первый доказательный тест.
2. Registry и диагностическая модель с положительными и отрицательными тестами.
3. Lexer, затем parser/AST и serializer с round-trip tests.
4. Identifier resolver, aliases, type checker и Named Expression dependency graph.
5. Limits, negative/fuzz/property/golden tests и сверка примеров приложения N.

Первый проверяемый инкремент: структура + версии + registry/diagnostics contracts + работающий тестовый runner. Следующие инкременты не объявляются завершёнными заранее.

## Проверка
Каждый функциональный инкремент начинается с теста ожидаемого поведения. Проверяются precedence, скобки, NOT, contextual/full equivalence, literals/lists, references, nested aliases, signatures, cycles, unknown IDs и malformed inputs. Serializer проходит parse → serialize → parse с эквивалентным AST. В тестах источники представлены metadata без чтения пользовательских данных.
Сборка проверяет допустимость и уникальность имён VBA-компонентов, включая ограничение 31 символ, импорт модулей и компиляцию проекта. Если реальный Excel или доступ к VBAProject недоступен, это явно фиксируется: статическая проверка не заменяет Excel compile/test.

## Готовность этапа
Независимая сборка из AccessTools-v2; синхронизированные grammar/spec/registry; детерминированный AST/serializer; зелёные tests; примеры N компилируются либо дают предусмотренные diagnostics; завершены semantic contracts. До этого evaluator и shadow integration не подключаются.

## Ограничения и дальнейшие решения
NoPolicyAction для email/domain → IP не выбирается молча и остаётся обязательной настройкой перед публикацией такой policy. Первый этап не зависит от выбора этой бизнес-политики.
После согласования этого решения составляется подробный план реализации с точными файлами, тестами и командами. Репозиторий GitHub на момент проверки содержит только README.md.

