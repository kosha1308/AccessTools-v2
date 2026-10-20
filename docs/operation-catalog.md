# Каталог типизированных операций — проект PR-01

Дополнение PR-08: реализованная чистая `OP_TEXT_REQUIRED@1.0.0` имеет сигнатуру Text→Bool без параметров, разрешает только EMPTY Input-обязательства; остальные состояния и технические ошибки сохраняет. Точная семантика и совместимость описаны в [контракте PR-08](pr08-multiple-rules.md). Нижеследующий проект 41 операции не означает их реализации.

Статус: **проект для ревью; операции не реализованы и не зарегистрированы**. Основание: ТЗ v0.15, ADR-003/ADR-007 и шесть документов PR-00. `OP_*@1.0.0` — предлагаемые неизменяемые ссылки будущего registry. Принятый ADR-007 требует полного включительного смысла IPv4-диапазонов. Q-01…Q-11 не утверждаются этим каталогом. Исторические C01–C21 и SHA-ссылки сохранены ниже отдельно от нового проекта.

## 1. Границы

Предлагается библиотека небольших чистых алгоритмов + структурный AST. Альтернативы — закрытая операция «evaluate policy» или перенос каждого legacy check отдельным контроллером — скрыли бы бизнес-логику от Admin. Здесь применение операций, условия, selector priority, строки и решения остаются видимыми узлами дерева.

Чистые операции не принимают Workbook/Range, имена шаблонов/листов, CheckId, пути или исполняемый текст; не читают файлы/часы, не выбирают latest source и не записывают Excel. Нет `OP_EVAL_POLICY`, `OP_EVAL_RULEBOOK`, `OP_GET_CELL`, `OP_WRITE_CELL`. Имена legacy-листов допустимы только в исторической трассировке.

Resolver разрешает FieldId/TableId/ColumnId/SourceBindingId, координаты, MergeArea, logical rows, aliases/headers, границы и visual-empty policy; выдаёт immutable typed snapshots + provenance. Evaluator исполняет AST. Outcome mapping создаёт смысловой результат; Action Planner связывает логические targets с provenance; Writer отдельно пишет/восстанавливает собственные аннотации. Текстовая операция не выбирает recipient/цвет/физический Range.

## 2. Общий контракт

`EvalResult<T>{state,value?,diag:List<Diagnostic>,provenance}`; value присутствует **только при OK**. `false`, 0, `List<T>[]` и пустое математическое множество — допустимые OK-значения.

| State | Значение |
|---|---|
| OK | Типизированное значение T вычислено |
| EMPTY | Пустой исходный ввод, не Bool(false) |
| MISSING | Логическое поле/привязка/элемент отсутствует |
| INVALID | Значение/параметр не соответствует типу/грамматике |
| UNAVAILABLE | Resolver не предоставил необходимый источник/snapshot |
| ERROR | Непредвиденная вычислительная ошибка |
| CONFLICT | Несовместимые равноценные значения без принятой tie policy |
| LIMIT | Лимит объёма/работы превышен; truncated value не выдаётся |
| NOT_APPLICABLE | Явная ветвь неприменимости, не универсальный fallback |

`Diagnostic={code,nodeId,operationRef,argumentId?,tokenSpan?,rowIds?,expectedType?,actualType?,safeDetails}`. Коды стабильны в major версии; trace не публикует секретные строки источников. Provenance содержит логические input refs/source revision tokens/RowIds и путь преобразования. Физические адреса — opaque refs Resolver/Action; они не участвуют в сравнении значений. Distinct/Join не теряют provenance исходных строк.

Сигнатуры далее показывают `f(args):T`; фактически каждая возвращает EvalResult<T>. Обычный Call eager вычисляет аргументы слева направо, затем при non-OK не вызывает алгоритм; diagnostics сохраняются. Для нескольких non-OK аргументов предлагается технический порядок ERROR > LIMIT > CONFLICT > UNAVAILABLE > INVALID > MISSING > EMPTY > NOT_APPLICABLE, затем порядок аргументов. Это проектная конвенция, не business severity и не порядок исполнения логических узлов. Обработка отсутствия — StateIs/When, без `EMPTY→false`, `UNAVAILABLE→[]` или автоматического ALLOW.

Каждая операция дополнительно может вернуть ERROR/INTERNAL, LIMIT/BUDGET_EXCEEDED. Несовместимые типы/enum/params/версии блокируют publish; повреждённый runtime пакет получает INVALID/TYPE_MISMATCH, INVALID/PARAMETER_INVALID или ERROR/OPERATION_VERSION_UNAVAILABLE. Остальные собственные ошибки перечислены у операций. UNAVAILABLE приходит от Resolver и пропагируется, а не означает отсутствие matches.

`RuleOutcome=PASS|FAIL|WARN|INFO|NEEDS_APPROVAL|NOT_APPLICABLE|CANNOT_EVALUATE` — отдельный тип. Bool false сам по себе не FAIL; опубликованная ветвь задаёт смысл. Required technical non-OK → CANNOT_EVALUATE; исключение требует явного доменного решения. NEEDS_APPROVAL не становится ошибкой инициатора. Q-10 о legacy PASS/read failures остаётся открытым.

## 3. Типы

| Тип | Нормализация / граница |
|---|---|
| Scalar | Text, Bool, Int64, Decimal, LocalDate (тот же календарный тип Date в expression-model), Email, Hostname, Enum<Id>; никаких неявных String→IP/Email |
| List<T>, Record<SchemaId>, Row<SchemaId> | Пустой List OK/count0; record members имеют собственный EvalResult<T>; поля по immutable MemberId/ColumnId, RowId стабилен внутри snapshot |
| IpSet | Отсортированные непересекающиеся включительные uint32-интервалы [lo,hi], смежные слиты; ∅ OK; не массив всех адресов |
| DeclarationSet | Набор исходных форм IP/CIDR/RANGE с введённым base/prefix, endpoints, raw-token provenance и profile version; не IpSet |
| AddressParse | Record{items:List<AddressItem>,groups:List<Text>,links:List<Text>,invalidTokens:List<InvalidAddressToken>,ambiguous:Bool,ambiguityReasons:List<Diagnostic>}; InvalidAddressToken={raw:Text,span:TokenSpan,code:Enum,provenance:ProvenanceRef}; raw — вычислительные данные, не автоматический dump в diagnostics/log. Форма/spans сохраняются |
| PortSet | Включительные интервалы 1..65535; отдельный от IpSet тип |
| MatchScore | Record{matches:Bool,rank:Int64,specificity:Int64,negativeMatched:Bool}; no-match корректный OK |
| Coverage<TSet> | Record{isCovered:Bool,covered:TSet,uncovered:TSet}; requested/available названы явно |
| IntervalMatches<Row> | List<Record{requestKey,row:Row,matched:IpSet}>; any overlap не обещает полной найденности |

AddressItem содержит kind (IPV4/CIDR/RANGE), raw, normalized, span, ipSet, declaration и provenance. Uppercase MemberId в матрице — фиксированный mapping registry: RAW_TOKEN→raw, TOKEN_SPAN→span, IP_SET→ipSet, DECLARATION→declaration, PROVENANCE→provenance, AMBIGUOUS→ambiguous; это не dynamic lookup по DisplayName. `ambiguous` описывает распознанную неоднозначность грамматики/смешивания, не стирает invalidTokens; случаи закреплены parser profile. Некорректный profile даёт INVALID, неоднозначность пользовательского ввода остаётся member+diagnostic для отдельного clause.

Для OP_INTERVAL_LOOKUP один requested IpSet — одна математическая query. `requestKey` — canonical value key этой query, не FieldId/RowId и не номер сводки. Caller хранит RequestEntry{ENTRY_ID,DISPLAY_TEXT,DECLARATIONS,IP_SET,PROVENANCE}: выполняет lookup на entry либо join cached candidate rows обратно по каждому entry. IP_UNION используется для batch query optimization, но не заменяет RequestEntry identity и не сливает соседние строки отчёта автоматически.

Profiles нормализации имеют immutable ID@version: trim/case/dash/separators/grammar явно выбраны; IP-разделители не глобальны. IPv6 и dotted netmask не объявлены поддерживаемыми в 1.0.0. DisplayName не используется как ссылка.

## 4. AST: не opaque evaluation operations

Все узлы имеют immutable NodeId, статический result type; named refs закреплены exact version. Это форма модели без выбора сериализации.

| Узел | Контракт |
|---|---|
| TypedLiteral(type,value) | OK false/0/[] допустимы; list literal содержит только константы, не unevaluated AST; source EMPTY моделируется state |
| ListExpr(T,children) | Производный typed List<T> из child expressions, в том числе List<List<IpSet>> для LIST_CONCAT(Map…,Map…). Non-OK required child пропагируется; порядок children фиксирован |
| FieldRef(FieldId) | Значение typed Resolver; без координат |
| ColumnRef(alias,ColumnId) | Только объявленный alias Record/Row и существующая typed column; sugar Member(ValueRef(alias),ColumnId). Scalar collection alias читается ValueRef(alias) напрямую |
| Member(recordExpr,MemberId) | Типизированное поле record возвращает собственный member envelope, не динамическое имя из строки |
| Record(SchemaId,MemberId→child AST) | Typed record construction со статическими slots; сохраняет EvalResult каждого member. Outer OK record не доказывает OK всех required members; obligations учитывает completeness validator |
| ContextRef(CLOCK_DATE/LOCALE/RUN_PROFILE) | Pin execution snapshot, без I/O и скрытого Date/Now |
| SourceRows(SourceBindingId) | Lazy EvalResult<List<Row<S>>>; недоступный source не OK[] |
| Call(OperationId@version,args,params) | Registry whitelist, полная проверка типов/версии/закрытой params schema |
| Let(BindingId,expr,body), ValueRef(BindingId) | Immutable lexical binding, одно вычисление; ValueRef разрешает Let и collection alias; refs вне scope/циклы запрещены |
| Filter(rows,alias,predicate:Bool) | Выбор List<Row>; non-OK predicate не удаляет строку тихо, а делает всю выборку non-OK |
| Map(rows,alias,projection:T) | List<T>, стабильный порядок; projection expr может быть Record со score + original row. Non-OK projection → non-OK Map; OK Record с non-OK member сохраняет этот envelope/obligation, не превращает его в default value |
| All/Any/None(rows,alias,predicate:Bool) | Для OK[] True/False/True; required presence отдельно Count>0. Слева направо short-circuit: All останавливается на false, Any на true, None на первом true с результатом false. Уже вычисленный non-OK немедленно пропагируется, последующие rows NOT_EVALUATED |
| Count(rows) | Int64 длины OK списка; недоступный список не становится 0 |
| GroupBy(rows,alias,key) | List<Record{key,rows}>; typed Scalar/Record key, стабильный порядок по первому RowId |
| Join(left,right,aliases,predicate,kind) | INNER/LEFT, List<Record{left:Row<L>,right:Row<R>}>; multiple matches дают все пары. LEFT unmatched right member=EMPTY<Row<R>>, не MISSING schema и не выдуманный Row. Predicate — Bool AST |
| When(condition,then:T,else:T) | Ленивые ветви одного типа; non-OK condition пропагируется |
| StateIs(expr,state) | Bool OK сравнения состояния, provenance/required technical obligations сохраняются; TRUE об ошибке не разрешает полный PASS |
| AllOf/AnyOf/Not | AllOf/AnyOf short-circuit слева направо: false/true соответственно останавливают вычисление. Вычисленный non-OK немедленно пропагируется и не маскируется следующей ветвью. Not не инвертирует state. Empty AllOf/AnyOf проектно True/False, бизнес-пустота требует явной ветви |
| RefExpression(ExpressionId@version,args) | Published named AST, проверенный DAG и сигнатура, не lexical ValueRef/свободная рекурсия |
| OutputRef(RuleRef,OutputId,InstanceKey) | Typed output того же run с explicit DAG edge, pinned RuleRef; отсутствующий required output→MISSING. Не читает counters, editable report cells или результаты прошлого запуска |

Граф зависимостей строится до исполнения, с conditional edges. Ссылка/тип невыбранной When либо short-circuit ветви валидируются, но её source не открывается. Невычисленная ветвь отмечается в trace NOT_EVALUATED (не EvalResult state), не создаёт runtime obligation и не диагностирует гипотетическую ошибку. Уже вычисленная ошибка не может быть замаскирована последующим false/true. Поэтому порядок логических узлов и snapshot rows семантически значим и фиксируется; индекс/optimizer не вправе его менять. Ограничения depth/node/rows/join cardinality явны; числовые пороги ещё не утверждены.

## 5. Общие scalar/text/selector/aggregate операции

Все строки — предложенные **@1.0.0**. Params доступны typed controls Admin; profileRef тоже ID@exactVersion, не строка VBA/regex-кода. Общие ошибки/пропагация §2 действуют для каждой строки.

| Operation / signature → output | Params Admin / семантика | Собственные states/errors; reuse |
|---|---|---|
| OP_TEXT_NORMALIZE(text:Text):Text | profileRef: trim, case invariant/preserve, newline, trailing-dot, allowed dash mapping | После trim пусто → EMPTY/BLANK_TEXT; INVALID/PROFILE_INVALID. Нормализация полей любого шаблона |
| OP_TEXT_TOKENS(text:Text):List<Text> | separatorProfileRef, normalizeProfileRef, deduplicate:Bool | Пустой ввод/только separators → EMPTY/NO_INPUT_TOKENS; INVALID/TOKEN_INVALID; LIMIT/TOKEN_COUNT. Протоколы, enum, selectors |
| OP_COMPARE<T:Scalar>(left:T,right:T,relation):Bool | EQ/NE всех Scalar; LT/LE/GT/GE ordered types; Text collationRef explicit | Неравенство OK false; INVALID/RELATION_TYPE_MISMATCH; records/set не Scalar. Сравнения значений, count, numeric priority |
| OP_NUMBER_BINARY<T:Int64/Decimal>(left:T,right:T,operator):T | ADD/SUB/MUL для обоих; DIV только Decimal, MOD только Int64; Decimal precision/scale/roundingProfileRef и numeric limit profile закреплены. Опциональный общий кандидат проекта, не новый бизнес-алгоритм | INVALID/NUMBER_OPERATOR_TYPE; INVALID/DIVIDE_BY_ZERO (DIV/MOD); ERROR/NUMBER_OVERFLOW при непредставимом результате; LIMIT/NUMERIC_BUDGET. Нет implicit Int64→Decimal coercion. Расчёты, приоритетные числовые score, интервальные метрики |
| OP_TEXT_MATCH(text:Text,pattern:Text,mode):Bool | EXACT/PREFIX/CONTAINS, explicit case/collation profile | Blank text EMPTY; пустой pattern INVALID/EMPTY_PATTERN. Свободный glob/regex не обещан. Строковый поиск |
| OP_EMAIL_PARSE(text:Text):Email | grammarProfileRef, domain normalization | EMPTY; INVALID/EMAIL_SYNTAX. Только синтаксис, без допуска субъекта |
| OP_SELECTOR_PARSE(text:Text,kind:EMAIL/HOSTNAME/IP):Record{positive:List<Text>,negative:List<Text>} | grammarProfileRef задаёт separators и NOT/НЕ/! prefixes; typed side extraction, не исполнение bool expression и не allow/deny policy | EMPTY/SELECTOR_INPUT_EMPTY; INVALID/SELECTOR_SYNTAX при malformed NOT/запрещённых separators. Negative-only допустим как parsed record, но score/applicability задаётся отдельно. Общие selectors любого источника |
| OP_EMAIL_SELECTOR_SCORE(email:Email,positive:List<Text>,negative:List<Text>):MatchScore | exact email rank2; domain/@domain/*@domain rank1, specificity=0 (не добавляет legacy-неизвестный length tie-break), selectorProfileRef; negative match запрещает совпадение строки | Positive OK[]→OK no-match rank0/specificity0; negative match→matches false/negativeMatched true/rank0; INVALID/SELECTOR_SYNTAX. Приоритетные справочники/ограничения |
| OP_HOSTNAME_MATCH_SCORE(host:Hostname,patterns:List<Text>,mode):MatchScore | EXACT rank3/PREFIX2/CONTAINS1, specificity=длина, normalizeProfileRef | patterns OK[]→OK no-match; INVALID/PATTERN_INVALID. Ranking имён ресурсов, не whole rule evaluation |
| OP_RANK_MAX<TRecord>(rows:List<TRecord>,rankFields:List<ColumnId>):List<TRecord> | Лексикографический порядок, direction по полям; возвращает все ties в стабильном порядке | OK[]→OK[]; INVALID/RANK_FIELD_TYPE. Приоритетный выбор, без first-row-wins |
| OP_UNIQUE_VALUE<TRecord>(rows:List<TRecord>,compareFields:List<ColumnId>):TRecord | Сравнение заданных полей по explicit normalization; предварительно PROJECT | 0 rows→EMPTY/NO_CANDIDATE; 1 distinct→OK representative со всеми refs; >1→CONFLICT/MULTIPLE_VALUES. Structured conflicts/однозначные metadata |
| OP_DISTINCT<T>(items:List<T>,keyFields?):List<T> | Scalar value либо typed record keys, normalization profile | OK[]→OK[]; INVALID/KEY_SCHEMA. Provenance дублей объединяется. Списки/отчёты |
| OP_LIST_CONCAT<T>(lists:List<List<T>>):List<T> | Один список списков одного T; сохраняет outer/inner порядок, не deduplicate | OK[]→OK[]; non-OK inner envelope пропагируется, LIMIT/OUTPUT_COUNT. Объединение derived collections без hidden array syntax |
| OP_SORT<TRecord>(rows:List<TRecord>,keys):List<TRecord> | Typed ordered fields/direction, explicit text profile, RowId tie-break | OK[]→OK[]; INVALID/SORT_KEY_TYPE. Детерминированный output |
| OP_PROJECT<TRecord,TOutRecord>(record:TRecord,mapping):TOutRecord | Только existing record projection: immutable target member IDs→source member paths, типосовместимые привязки, без исполняемого текста; member envelopes/provenance сохраняются | Отсутствующий member в schema INVALID/PROJECTION_SCHEMA (publish error); существующий MISSING/EMPTY member остаётся таким в OK record и при Member извлечении. Map/Join transformation |
| OP_TEXT_JOIN(items:List<Text>,separator:Text):Text | unique:Bool, normalizeProfileRef, input order/явный Sort | OK[]→OK("") как вычисленный текст; LIMIT/OUTPUT_LENGTH. Metadata/инструкции |
| OP_LIST_CONTAINS<T:Scalar>(items:List<T>,value:T):Bool | comparison/case profile | OK[]→OK false; INVALID/COMPARISON_PROFILE. Membership/enum validation |
| OP_HOSTNAME_EQUIVALENT(left:Hostname,right:Hostname):Bool | EXACT_FQDN/SHORT_NAME profile; trailing-dot/case explicit | No equality→OK false; INVALID/PROFILE_INVALID. Join уточнение, не DNS lookup |
| OP_TEXT_FORMAT(template:MessageTemplate,parameters:Record):Text | Typed immutable template **value**, предварительно разрешённый pinned ref компилятором/Resolver; placeholders и formatting locale explicit | MISSING/MESSAGE_PARAMETER; INVALID/FORMAT_TYPE; LIMIT/OUTPUT_LENGTH. Нет lookup по ref внутри pure call; только текст без recipient/target/style |

Ranking использует `(rank,specificity)` как два поля вместо legacy `rank*100000+length`. Дополнительный business priority — поле и порядок OP_RANK_MAX, не константа matcher. Равные лучшие structured results проверяет UNIQUE_VALUE. Broad approvals/comments строятся Map **всех matched rows отдельно**, не только winners: ранжирование структурного результата не удаляет инструктивные выходы. Raw selectors проходят SELECTOR_PARSE; Member positive/negative передаются score operation. Их преобразование не скрывается в адаптере источника. Произвольная логика NOT остаётся AST.

Tuple `(rank,length)` — предложенный способ порядка, **не безусловно эквивалентный** legacy numeric score. Эквивалентность `rank*100000+length` доказывается только при закреплённом profile bound `length<100000`; без такого ограничения возможен переход через rank band. Если нужен точный legacy numeric порядок, AST явно вычисляет score через NUMBER_BINARY(MUL,ADD) и RANK_MAX по score. Выбор tuple/legacy numeric mode и максимальная длина относятся к migration review; скрытого изменения приоритета здесь не утверждено.

Для OP_RANK_MAX/OP_SORT schema Record поля адресуются ColumnId typed relation schema; для OP_PROJECT output members — MemberId. Эти refs задаёт editor/schema, не пользовательские строки. OP_COMPARE records не сравнивает: scalar members или UNIQUE_VALUE(compareFields) обязательны.

SELECTOR_PARSE имеет закрытый `SelectorGrammarProfile@version`: `separators:List<SeparatorEnum>`, `negativePrefixes:List<NOT_WORD|NE_WORD|BANG>`, `prefixBoundary:SPACE_OR_COLON`, `trim:Bool`, `case:INVARIANT`, `maxTokens:Int`, `maxTokenLength:Int`. Это проект schema, не принятый default эксплуатации. Split выполняется по выбранным separators, затем снимается не более одного negative prefix, после чего проверяется selector kind. EMAIL token — exact email либо domain/@domain/*@domain; HOSTNAME token — непустой pattern для выбранного match mode, не executable wildcard expression. IP token сохраняет полный range/CIDR/x-selector как Text, включая пробелы вокруг дефиса; после lexical splitting каждый token отдельно проходит OP_IP_PARSE с выбранным grammar profile. SELECTOR_PARSE не присваивает allow/deny смысл: положительные/отрицательные части разрешающей и запрещающей колонок объединяет явный AST. Пустые split fragments игнорируются; bare NOT/НЕ/!, повторный negative prefix, скобочная boolean expression, malformed email/domain и неразрешённый separator дают INVALID/SELECTOR_SYNTAX с tokenSpan. Слишком длинный ввод/число токенов — LIMIT/TOKEN_LENGTH или TOKEN_COUNT, без partial value. Negative-only list допустим, но не означает «разрешить всех»: score требует positive match, если AST отдельно не задаёт иное. Positive/negative lists дедуплицируются независимо с объединением refs; одинаковый selector на обеих сторонах сохраняется, negative match имеет приоритет. Ни SourceBinding adapter, ни OP_TEXT_TOKENS не исполняют NOT скрыто.

## 6. Даты и clock

| Operation / signature | Params Admin | Собственные states/errors; reuse |
|---|---|---|
| OP_DATE_PARSE(raw:Text/LocalDate):LocalDate | mode STRICT_DMY_TEXT либо proposed LEGACY_ISDATE_FIRST; separators; last-valid-match; fallback year range; localeRef/conversionProfileRef | EMPTY/DATE_INPUT_EMPTY; INVALID/DATE_SYNTAX или DATE_CALENDAR; строгий DateSerial-style roundtrip, без переноса 31 февраля. LocalDate input возвращается без текстового parse. В legacy-first принимаемые конверсией даты могут быть вне fallback года; explain фиксирует ветвь. Сроки любых документов |
| OP_DATE_COMPARE(left:LocalDate,right:LocalDate,relation):Bool | EQ/NE/LT/LE/GT/GE; today отдельный ContextRef CLOCK_DATE | Нет Date/Now внутри pure call; missing clock propagates MISSING, не default date. Сроки/даты/интервальные условия |

Excel Date/serial/Error Variant читает Resolver с явным read profile и приводит к Text/LocalDate либо INVALID. Clock date/locale/timezone захватываются один раз в execution snapshot; нет смены результата от полуночи между Calls. Legacy compatibility — предложение, не утверждение всех locale quirks. «Бессрочно», revoke, request type и missing locator — When/StateIs/сравнения AST, не скрытые ветви date algorithm. Q-10 о legacy технических исходах открыт.

## 7. Mixed address parse, IpSet и DeclarationSet

| Operation / signature | Params Admin | Собственные states/errors; reuse |
|---|---|---|
| OP_ADDRESS_PARSE_MIXED(text:Text):AddressParse | grammarProfileRef/separatorProfileRef, allowGroups/allowLinks, compactMode DISABLED/APPROVED_PROFILE, reverseRange REJECT/NORMALIZE, suffixWildcardMode DISABLED/X_SUFFIX | Пусто EMPTY/ADDRESS_INPUT_EMPTY; при распознанном mixed вводе **OK** record со всеми items/groups/links/**invalidTokens** и diagnostics; invalid fragments не исчезают. Structural profile invalid→INVALID/GRAMMAR_PROFILE; LIMIT/TOKEN_COUNT или INTERVAL_COUNT без partial value. Mixed fields |
| OP_IP_PARSE(text:Text):IpSet | Versioned grammar/separators/range/wildcard profiles, groups/links запрещены | Пусто EMPTY; group/link/non-IP или malformed fragment INVALID/IP_SYNTAX/IP_EXPECTED; ни одного «частично корректного» IP из неверного октета. Inclusive intervals без enumeration. Адресные sets |
| OP_IP_FROM_ADDRESS_PARSE(parsed:AddressParse):IpSet | invalidPolicy STRICT/EXPLICIT_VALID_SUBSET; subset-mode только явно, output explain маркирует incomplete input | STRICT при invalidTokens nonempty→INVALID/ADDRESS_HAS_INVALID; EXPLICIT_VALID_SUBSET→OK valid union + warning diag, не основание PASS полного запроса. Только группы/links→OK∅. INVALID/ADDRESS_ITEM_SCHEMA для повреждённого typed record. Parse reuse |
| OP_DECLARATIONS_FROM_ADDRESS_PARSE(parsed:AddressParse):DeclarationSet | declarationProfileRef: preserve CIDR entered base/prefix, canonical range endpoints, order/dedup; invalidPolicy только STRICT | invalidTokens→INVALID/ADDRESS_HAS_INVALID; только groups→OK empty set; damaged items INVALID/DECLARATION_SCHEMA. Не получает IpSet: исходная форма потеряна |
| OP_DECLARATIONS_COMPARE(left:DeclarationSet,right:DeclarationSet):Record{equal:Bool,leftOnly:DeclarationSet,rightOnly:DeclarationSet} | Profile STRICT_FORMS либо PROPOSED_EXPLICIT_COVERAGE_WITH_STRICT_CIDR, exact version | OK false + differences — бизнес-неравенство; INVALID/DECLARATION_PROFILE_MISMATCH. Выбор business profile Q-07 **открыт**, enum не означает принятия |

AddressParse OK означает корректное извлечение структурированного разбора, **не валидность всей адресной ячейки**. Правило обязано явно проверить Count(Member(parsed,invalidTokens))=0 до positive outcome полного запроса; диагностический valid-subset не подменяет этот gate. Необработанный INVALID/UNAVAILABLE required parser input → CANNOT_EVALUATE.

Исторический MATCH.invalid с raw invalid declaration не ослабляет STRICT: Map invalidTokens→raw Text, DISTINCT и взаимные LIST_CONTAINS образуют отдельную симметричную разность raw keys. Параллельно Record может явно построить AddressParse **valid-items branch** из original items/groups/links и пустого invalidTokens для строгого извлечения только валидных деклараций. Original invalid list/provenance/format FAIL сохраняются отдельной required clause; такая диагностическая ветвь не удостоверяет весь ввод. Одинаковый invalid raw token может не дать mismatch, но всё равно не даёт корректного адреса или общего PASS.

ADR-007 действует во всех профилях: `a-b` содержит **все** промежуточные адреса включительно. Reverse normalization отдельный параметр. Compact `x.x.x.x,y` — отдельная неутверждённая grammar: DISABLED даёт INVALID/UNAPPROVED_COMPACT_FORM в strict parse, либо invalid token в mixed result. Historical wildcard x в **целых последних октетах** — другой profile, не inference из текста. Dotted netmask/произвольная wildcard mask не обещаны.

CIDR `host/prefix` в IpSet округляется к сети; DeclarationSet сохраняет введённую форму по профилю. STRICT_FORMS может различать range/list и CIDR/plain, но не объявляет диапазон только endpoints. Предлагаемый профиль explicit coverage + strict CIDR требует Q-07/fixtures; он не default всех шаблонов. Semantic types не преобразуются неявно.

## 8. Интервалы и lookup

Все интервалы включительны, нормализованы, без обязательного перечисления адресов. OK∅ допустимо. Общие ошибки §2 действуют; повреждённые bounds дают INVALID/INTERVAL_BOUNDS (для PortSet PORT_INTERVAL_BOUNDS). LIMIT всегда без truncated value.

| Operation @1.0.0 / signature | Точная семантика; Admin/reuse |
|---|---|
| OP_IP_UNION(sets:List<IpSet>):IpSet | Union с merge overlaps/adjacency; OK[]→OK∅; объединение таблиц/источников |
| OP_IP_INTERSECTION(left:IpSet,right:IpSet):IpSet | A∩B; deny/match details |
| OP_IP_DIFFERENCE(left:IpSet,right:IpSet):IpSet | A−B, внутренние holes сохраняются; allowed−denied |
| OP_IP_OVERLAPS(left:IpSet,right:IpSet):Bool | A∩B≠∅; empty→false; matched-any, не permission |
| OP_IP_COVERAGE(requested:IpSet,available:IpSet):Coverage<IpSet> | covered=A∩B, uncovered=A−B, isCovered=uncovered∅; empty request→true, обязательность отдельным cardinality gate |
| OP_IP_EQUAL(left:IpSet,right:IpSet):Bool | Равенство адресов независимо от declaration forms |
| OP_IP_CARDINALITY(set:IpSet):Int64 | Σ(hi−lo+1); /0=4294967296 без массива |
| OP_INTERVAL_LOOKUP<Row>(requested:IpSet,rows:List<Row>,ipColumn:ColumnId):IntervalMatches<Row> | Каждая row с непустым intersection, matched subset, стабильный request/RowId порядок. Params source column binding + limit profile; no latest file choice. MISSING/COLUMN_MISSING, INVALID/LOOKUP_COLUMN_TYPE, LIMIT/MATCH_COUNT. Source OK[]→OK[] |

Полное покрытие lookup строится **отдельно**: Map(matches→matched) → IP_UNION → IP_COVERAGE(requested,union). Count(matches)>0 — matched-any. Source содержит только .1 и .10 при request .1–.10: два matches, но uncovered .2–.9. Deny сравнивается со **всем requested**, независимо от найденных source rows: .5 внутри запроса даёт overlap=true. Allowed−Denied с hole .5 не полностью покрывает запрос, даже если оба конца разрешены.

| Port operation @1.0.0 / signature | Params / states/errors / reuse |
|---|---|
| OP_PORT_PARSE(text:Text):PortSet | separatorProfileRef, reverseRange REJECT/NORMALIZE, grammarProfile; 1..65535. EMPTY/PORT_INPUT_EMPTY, INVALID/PORT_SYNTAX/PORT_BOUNDS; tokens не теряются |
| OP_PORT_UNION(sets:List<PortSet>):PortSet | Union; empty list→OK∅; объединение условий |
| OP_PORT_INTERSECTION(left:PortSet,right:PortSet):PortSet | A∩B; детали overlap |
| OP_PORT_DIFFERENCE(left:PortSet,right:PortSet):PortSet | A−B с holes; exclusions |
| OP_PORT_OVERLAPS(left:PortSet,right:PortSet):Bool | Any intersection; отдельный ANY предикат |
| OP_PORT_COVERAGE(requested:PortSet,available:PortSet):Coverage<PortSet> | uncovered=requested−available; positive/exemption требует full coverage, не overlap |
| OP_PORT_EQUAL(left:PortSet,right:PortSet):Bool | Normalized set equality |

Port math имеет те же empty/limits законы, что IP math, но другой тип. Протоколы: TEXT_TOKENS + LIST_CONTAINS/Any/All над typed enum/catalog. Пустое условие «любой» задаёт AST When/StateIs, не parser default. Slash/backslash разрешены только explicit port profile и не распространяются на CIDR/обычный текст.

Q-11: исторически доказаны IP lookup и отдельно port predicate условия. Port lookup **в самой выгрузке** не доказан и не обещан OP_INTERVAL_LOOKUP. Возможный Join IP+PortSet — потенциальная композиция только после принятой SourceBinding schema/портовой колонки/ANY-vs-COVERAGE policy, не существующий паритет.

## 9. Сложные случаи как открытая композиция AST

Это доказательство формы выразимости, не исполняемая конфигурация реального правила.

- Primary eligibility отдельно → EMAIL_SELECTOR_SCORE по source rows → Filter matches → RANK_MAX всех ties → scope Filter → Map allowed/denied → union → deny overlap/full coverage. No secondary match — явная When ветвь опубликованной политики, не ALLOW при UNAVAILABLE.
- Matched-key ALLROWS: Filter/GroupBy rows по anchor intersection в scope → Count>0 → All(тот же predicate). Посторонние IP вне anchor не участвуют. Predicate — видимый AST, не callback. Исторический поддержанный subset и предполагаемое расширение не смешиваются; suppress positive Action описывается ветвью, а не удалением issues внутри pure op.
- Hostname/service: Join interval matches с typed hints; если есть hostname matches — Filter matching + unhinted; иначе явно выбранный IP-only fallback. Service берётся из этой же row. RANK_MAX/UNIQUE_VALUE структурных выходов отдельно от Map всех broad approvals/comments.
- Mixed groups/lower addresses: группы отдельны от IpSet; upper+lower IpSet объединяется без раскрытия группы. Недоказуемое «лишний нижний IP» при неизвестном составе группы — applicability/needs source, не выдуманная membership. Сторона односторонней policy не угадывается.
- Merge/duplicates: Logical TableResolver задаёт RowIds, поля, targets и partial/full merge policy. GroupBy(normalized tuple)+Count>1 даёт evidence повторов. OP_DISTINCT не сканирует Excel.
- Metadata: Join по typed key → PROJECT/DISTINCT/TEXT_JOIN; UNIQUE_VALUE для однозначного имени. CONFLICT обрабатывается явной ветвью; финальные колонки/порядок/recipient и сохранение файла принадлежат OutputProfile/Action, не pure aggregate.
- Срок: When request-type/infinite literal/StateIs missing → соответствующая outcome ветвь; иначе DATE_PARSE и DATE_COMPARE с ContextRef CLOCK_DATE. Date operation не получает шаблонный переключатель типа заявки.

## 10. Admin, registry, version pin и миграции

OperationDefinition: immutable OperationId, exact SemVer, InputTypes/OutputType/generic constraints, закрытая ParameterSchema и versions профилей, state/error table, limit dimensions, deterministic flag, complexity description, UI label/help/examples, compatibility/deprecation metadata. Все OP_* здесь pure/deterministic при одинаковых typed inputs/params. Источник, часы и локаль находятся снаружи.

Admin фильтрует операции по типу slot; показывает typed входы, параметры enum/profile, состояния/ошибки, пример и dependencies. Условия коллекций редактируются деревом, поля/колонки выбираются по IDs, literal проверяется по типу. Никакого macro name, свободного Eval/VBA/SQL/regex-code. «Где используется» и Test/Explain доступны до publish. Количество строк/matches и covered/uncovered видны отдельно от финального outcome.

PublishedPackage закрепляет exact OperationId@version, params profile refs, AST schema version, ExpressionId refs, SourceBinding versions. Patch не подменяется автоматически. Patch — исправление без изменения утверждённой семантики; minor — совместимое расширение; major — изменение результата/грамматики/приоритета/state/errors. При любом changed golden нужен migration review; DisplayName можно менять без смены ID. Нет нужной версии → publish blocked, runtime ERROR, без latest fallback.

Миграция: draft copy → old/new typed definitions diff → обе версии на обезличенных golden → explain semantic deltas → утверждение спорных Q → новый immutable package. Старый пакет остаётся для rollback. C01–C21 — trace labels, не автоматически OperationId. ADR-007 уже принят; endpoints-only expected legacy fixture нельзя тихо заменить: intentional delta фиксируется отдельно.

Dependency planner строит граф Field/Table/Source refs, operation/expression versions, ролей и conditional edges. Источники ленивы по реально нужным ветвям; старт приложения не открывает все книги. Resolver закрепляет immutable snapshot/index по `(SourceId,BindingVersion,revisionToken,schemaVersion,indexVersion)`; token должен описывать подтверждённую ревизию, не только имя файла. Несохранённая открытая книга требует собственного captured revision token. Cache pure Call: exact operation version + canonical typed inputs + params/profile versions; не Excel addresses. Смена source внутри запуска — новый snapshot либо technical state, не смешение ревизий.

Clock/locale/run profile — ContextRefs в execution snapshot и explain. Pure operation не запрашивает token/диск/Date/Now; index — взаимозаменяемая оптимизация typed lookup, не другой бизнес-алгоритм. Limit policy pin входит в package; LIMIT не выдаёт частичные rows как полные. Конкретные числовые performance ceilings остаются предметом измерений.

Resolver отдельно фиксирует headers/aliases, formula/read profiles, visual empty, hidden rows, table end, merged blocks/partial merges, deterministic RowId и provenance. Его контракт `EvalResult<List<Row>>`: отсутствующая таблица MISSING, прочитанная пустая OK[]. Физические target Range и Writer не замораживаются этим каталогом: Q-04 открыт. Даты Excel Error Variant → INVALID через read profile, а не исключение свободного CStr в pure operation.

## 11. AC-16 и дополнительные проектные векторы

**Ожидаемые результаты дизайна, NOT EXECUTED.** Это синтетические случаи для будущих contract/golden тестов; не заявления PASS VBA/Excel.

| Вектор | Ожидаемый typed результат / граница |
|---|---|
| Parse 10.0.0.1–10.0.0.10 по approved dash profile | IpSet [.1,.10], CARDINALITY=10, не 2 endpoints; ASCII '-' то же |
| Request .1–.10, deny=.5 | OVERLAPS=true, intersection=.5, difference=[.1,.4]∪[.6,.10] |
| Request .1–.10, available .1–.8 | Coverage false, uncovered=.9–.10 |
| Lookup .1–.10, source только .1 и .10 | OK matches count2; matched-any true; coverage false/uncovered=.2–.9 |
| Список .1,.10 против диапазона .1–.10 | IP_EQUAL=false; cardinality2 vs10; DeclarationSet тест отдельно |
| Крупный /0 | CARDINALITY=4294967296, interval count1; без enumeration; lookup limit→LIMIT/no value |
| Allow [.1,.4]∪[.6,.10], request .1–.10 | Coverage false; hole .5 сохраняется, оба endpoints не спасают |
| TOP declaration plain .1 vs .1/32 | IP_EQUAL=true; STRICT_FORMS DECLARATIONS_COMPARE=false. Business profile Q-07 открыт |
| Range vs полный explicit list | IP_EQUAL=true по ADR-007; DeclarationCompare expected зависит от **выбранного** profile, не implicit coercion |
| Compact .1,2 | DISABLED strict→INVALID; mixed→OK invalidTokens; approved grammar expected ещё требует решения |
| Mixed valid IP + invalid octet + group | ADDRESS_PARSE_MIXED OK record с invalidTokens/diag; STRICT extraction INVALID; нельзя положительно проверить полный запрос по subset |
| Только допустимая группа | Mixed OK groups; IP extraction OK∅; обязательность отдельная ветвь |
| Ports 443/8443, 443\8443 | Approved slash profile→{443,8443}; другой профиль INVALID; CIDR не разбивается |
| Request ports443,9000 vs allowed443,8443 | OVERLAPS=true, COVERAGE=false/uncovered9000 |
| Source UNAVAILABLE vs OK[] | Required expression CE для UNAVAILABLE; OK[] lookup no-match, business branch явно выбирает outcome |
| EMPTY/MISSING/INVALID даты, сегодня/вчера | States различны; today≥clock today true, yesterday<today true; clock неизменен внутри запуска |
| ALLROWS anchor A, unrelated B | Все A true→true; один A false→false; B исключён; нет A→Count0 и явная applicability |
| Top ties с разными structured values | RANK_MAX возвращает обе; UNIQUE_VALUE CONFLICT; broad outputs не исчезают |
| Non-OK predicate в фактически вычисленной row | Filter/All/Any пропагируют non-OK и RowIds; если short-circuit уже остановился до этой row, она NOT_EVALUATED, obligation не возникает |
| Два revisions при одинаковом SourceId | Разные cache keys; невозможность capture→technical state, не mixed snapshot OK |

Дополнительно: все C01–C21 legacy fixtures, NOT selectors, exact/domain priority, suffix wildcard invalid middle x, shared source двух независимых шаблонов, переименование координаты без изменения операции, циклы/типы/версии/лимиты, partial output failure. Excel UI/Writer и текущие release gates не запускались.

## 12. Трассировка исторических C01–C21

В таблице сокращённые operation names обозначают полные `OP_<name>@1.0.0` из §§5–8; AST nodes — §4. Это проектная миграция, не уже работающая новая реализация.

| Trace | Proposed primitives / композиция | Вне pure calls |
|---|---|---|
| C01 | ADDRESS_PARSE_MIXED → IP_FROM_ADDRESS_PARSE; invalid count gate | FieldResolver/group applicability |
| C02 | IP_PARSE, UNION/INTERSECTION/DIFFERENCE/OVERLAPS/COVERAGE/EQUAL/CARDINALITY | Scope выбирается AST |
| C03 | PORT_PARSE | Field binding/separator profile |
| C04 | PORT_OVERLAPS отдельно PORT_COVERAGE | Необязательное условие/positive decision |
| C05 | TEXT_TOKENS, LIST_CONTAINS, All/Any | SourceRows catalog/enum binding |
| C06 | HOSTNAME_MATCH_SCORE, RANK_MAX, UNIQUE_VALUE | Broad outputs отдельная ветвь |
| C07 | EMAIL_PARSE, EMAIL_SELECTOR_SCORE, Filter/Any | Primary eligibility отдельно secondary policy |
| C08 | INTERVAL_LOOKUP, IP_COVERAGE | SourceResolver/index/view Action |
| C09 | Join/Filter/Map + typed IP/PORT predicates + hostname ranking | Весь condition tree в AST |
| C10 | PROJECT/TEXT_FORMAT + StateIs/When | Metadata sources/routing/Actions |
| C11 | SELECTOR_PARSE, EMAIL_SELECTOR_SCORE, Filter, RANK_MAX | Active/no-policy branch |
| C12 | IP_UNION/DIFFERENCE/OVERLAPS/COVERAGE | RuleOutcome/recipient |
| C13 | SELECTOR_PARSE(kind IP), Map IP_PARSE approved suffixWildcard profile, IP_UNION | Grammar acceptance; allow/deny объединение и NOT смысл в AST |
| C14 | EMAIL_SELECTOR_SCORE/COMPARE/When | Account FieldResolver/applicability |
| C15 | Filter/GroupBy/All/Count/IP_OVERLAPS | Legacy syntax migration/Action branches |
| C16 | DECLARATIONS_FROM_ADDRESS_PARSE/DECLARATIONS_COMPARE отдельно IP math | Q-07 profile/physical refs |
| C17 | TEXT_NORMALIZE, GroupBy/Count/DISTINCT | Logical TableResolver/merge/targets |
| C18 | IP_UNION/INTERVAL_LOOKUP/When | Group membership и source side не угадываются |
| C19 | INTERVAL_LOOKUP/HOSTNAME_EQUIVALENT/Join/Filter | IP-only fallback видимая ветвь |
| C20 | Join/PROJECT/DISTINCT/UNIQUE_VALUE/TEXT_JOIN | OutputProfile/Writer |
| C21 | DATE_PARSE/DATE_COMPARE/When/StateIs | Clock/locale context/revoke/infinite policy |

Q-01 baseline, Q-02 fixtures, Q-03 policy, Q-04 Writer, Q-05 routing, Q-06 summary, Q-07 declarations, Q-08 полнота, Q-09 global messages, Q-10 technical outcomes, Q-11 port lookup остаются открытыми в [реестре PR-00](v1-audit-findings.md). В частности, предлагаемая structured conflict policy и декларационные profiles требуют ревью; наличие каталога не снимает решение владельца. ADR-007 уже утверждён и имеет приоритет независимо от состояния остальных Q.

---

# Приложение: историческая трассировка PR-00 C01–C21

Ниже сохранён audit trail AccessTools 1. Он не нормативный контракт новых OP_*, не утверждает Q-01…Q-11 и не означает прохождения тестов нового продукта. ADR-007/ТЗ15 имеют приоритет над прежними конфликтами range/compact.

## Предварительный каталог сложных операций (C), историческая редакция

PR-00: инвентаризация алгоритмов v1 для обсуждения паритета. Имена ниже — ярлыки аудита, не утверждённые OperationId, версии или дизайн PR-01. Обязательный последний рабочий эталон: integration 58d8789d1e60860c96cb35ba0e780a6424510130; исторический main d9aa9b88f9193c0d2935d99c5cc8ecca3865687c рассмотрен параллельно. C01–C10 ниже описывают исторический слой main; C11–C20 фиксируют поздние обязательные семантики integration и уточняют старые контракты; v0.14 §4 требует регистрацию операций с типами/параметрами/статусами, но такие контракты здесь не изобретены. Admin-параметры означают наблюдаемые бизнес-настройки или вопросы переноса, а не готовый интерфейс.

| C / назначение | Вход → результат | Пусто / ошибка / ограничения | Admin-параметры; повторное применение | Код |
|---|---|---|---|---|
| C01 IPv4 и списки заявки | Text → Ips, InvalidTokens, Groups | IPv4 4 октета 0..255; CIDR /0..32; одиночный CIDR в старом парсере даёт host, не сеть. CR/LF, ; , и pipe; сокращённые последние октеты. Диапазон разворачивается, обратные границы меняются местами, лимит 65536; пустой текст даёт пустые словари | Семантика списка/групп; лимит сейчас константа. Проверки заявки и импорт | [AnalyzeNetworkAddressText](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkParsers.bas#L46) |
| C02 Числовые интервалы IPv4/CIDR | Text → множество START:END; membership, overlap, intersection text | CIDR округляет к сети через 2^(32-prefix); /0 без перечисления адресов. Неверные токены отделяются от допустимых; это не полная валидация текста. Ключ/парсер имеют собственные проверки | Scope source/destination/оба; валидация отдельно. Заявка, Rulebook, Superset | [NetworkCidrBounds / interval API](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkIpIntervals.bas#L191) |
| C03 Порты строгой заявки | Text → intervals + invalidTokens | 1..65535; обратный диапазон invalid; пробелы/табуляции/переносы/; , pipe. Пусто → пустые наборы | Разделители и границы сейчас код; импорт/проверки | [AnalyzeRequestPorts](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkParsers.bas#L389) |
| C04 Порты условий / покрытие | request intervals + rule text → Boolean | Legacy ParsePortIntervals извлекает числовые подстроки, разворачивает обратные границы. PortIntervalsMatch = любое пересечение, пустое условие True. Runtime Flex требует полного покрытия всех портов строки; частичное пересечение не освобождает строку | Условия порта Rulebook. Обязательно различать overlap и covers, а не один универсальный match | [RuntimePortMatches / covers](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkRulebookFlexCache.bas#L426) |
| C05 Токены протоколов / membership | request text + rule tokens → Boolean | Без учёта регистра; ParseTextTokens делит также по / + пробелу; пустое правило True; совпадает любой токен. Наличие в активном разрешённом списке проверяется отдельно | Активные протоколы; повторно membership в справочнике | [ParseTextTokens / ProtocolMatches](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkParsers.bas#L498) |
| C06 hostname и специфичность | hostname + Patterns + MatchType → match, score | contains=1, starts=2, иначе exact=3; score=rank*100000+длина совпавшего pattern, максимум. Нет совпадения =0. Маска ТЗ не равна автоматически glob/regex v1 | Тип, patterns; выбор наиболее специфичного правила, конфликты при равном приоритете | [HostnameSpecificity](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkRulebookV2.bas#L733) |
| C07 Политика инициатора | email + active rule collection → разрешение, согласования | Lower/Trim; email с @; точное совпадение полного адреса или домена, optional ведущий @. Поддомены не включаются автоматически. approval-only не разрешает; неактивная строка не совпадает | Домены/почты, активность, approval, severity, comment, служебный режим. Policy, допуски, согласования | [Initiator policy](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkSupportSheets.bas#L655) |
| C08 Superset IP lookup | интервалы заявки + index → строки источника | Три SOURCE-поля объединяются; возвращается связь с исходной строкой. Кэш v2 хранит сеть и host alias CIDR. Отсутствие источника/невалидная схема/пустой index различимы на уровне loader | SourceBinding, выбор выгрузки, заголовки; поиск в реестрах и объяснение результата | [BuildRegistryIntervalIndex](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modRegistryIntervalIndex.bas#L242) |
| C09 Join IP→FQDN/сервис и составное правило | technical row + registry context + rules → match / structured services, groups, approvers, comments | Условия AND; сервис связывается с IP этой же строки, а не с любым сервисом книги. Hostname specificity и конфликт результатов не заменяются first-row-wins | IP scope, hostname, port, protocol, service-condition; результирующие mapping-поля. Обогащение таблиц | [RuntimeFastMatchesTechnicalRow](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkRulebookFlexCache.bas#L303) |
| C10 Метаданные support / messages | key + source settings → тексты, severity, маршрутизация | Заголовки и aliases мигрируют; отсутствующий ключ, выключенная строка, пустой override имеют отдельные места обработки. Это не IP-алгоритм и не direct writer | Ключи, тексты пользователя/координатора, ошибка пользователя, подсветка, активно. Каталоги сообщений других шаблонов | [Direct problem metadata](https://github.com/kosha1308/AccessTools/blob/d9aa9b88f9193c0d2935d99c5cc8ecca3865687c/src/checks/network_access/modNetworkProblemSchemaV5.bas#L4) |

## Границы и несоответствия

«Маска» требует уточнения: в прочитанных IPv4-функциях реализован числовой CIDR prefix, но dotted mask 255.255.255.0 не установлена как поддерживаемый контракт; hostname содержит exact/starts/contains, поэтому обещать произвольные wildcard-маски нельзя. Старый ExtractIPv4s теряет сетевой охват CIDR; interval API сохраняет его. Эти API не взаимозаменяемы.

Для C09 порядок применения результатов зависит от технического/hostname/Superset контекста: одинаковая специфичность с разными структурными результатами требует явного конфликта, а не выбора произвольной строки. Сервис результата отличается от сервиса условия. Перенос требует сохранить provenance строки источника, а вычислитель не должен писать ячейки (требование v0.14).

Для будущей регистрации отсутствуют утверждённые: OperationId/version, единая схема типизированных ошибок, null/empty/invalid контракты для каждого слоя и согласованный перенос констант в Admin. Это вопросы следующего этапа; данная таблица не делает permissive v1-функцию строгой и не вводит новую policy.

## Проверка

Код: функции прочитаны по указанному SHA. Тесты: в исходниках присутствуют RunNetworkCidrRegression, RunRegistryIntervalParserRegression, RunNetworkInitiatorRulesV2Regression, RunNetworkRulebookFlexCacheRegression и RunHostnameRuleNotesRegression; наличие не означает прохождение. В этом аудите VBA и регрессии не исполнялись. Ручной Excel: не проверен.

Минимальные эталоны для последующего согласования: /0,/32, host bits в CIDR, обратный и чрезмерный диапазон; неполный IP и группа; порт 0/65536, обратный диапазон, частичное против полного покрытия; пустой протокол и смешанный регистр; exact против starts/contains и tie; точная почта против домена/поддомена/approval-only; новый невалидный Superset при старом валидном; несохранённая открытая книга; отсутствующий key и выключенный override. Это предложенные проверки, не выполненные тесты.



## Поздний обязательный слой integration: C11–C20

Следующие ссылки закреплены на 58d8789. Здесь «ошибка» означает именно наблюдаемый выход функции/проверки, а не проект нового Status enum. Настройки Admin перечислены как переносимые данные, не как утверждённая форма UI.

| C / назначение | Точные входы → выход | Empty / missing / invalid / error | Admin-параметры; reuse; доказательство |
|---|---|---|---|
| C11 Выбор policy email/domain с NOT | email String, policies Collection с EmailSelector → Collection всех строк максимальной specificity | exact email=2, @domain / *@domain / domain=1; совпавший отрицательный selector даёт 0 всей строке; pipe недопустим; нет подходящих → пустая Collection, проверка без нарушения | selectors, active, scope; приоритетный справочный выбор для иных полей. [SelectNetworkInitiatorIpPolicies](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L630) |
| C12 IP policy Allowed−Denied и full coverage | requested интервалы одной стороны technical row; выбранные policy с AllowedIntervals/DeniedIntervals; side SOURCE/DESTINATION → violation Boolean и ERROR на SourceAddress/DestinationAddress | Allow объединяется по выбранным applicable rows; Deny имеет приоритет при любом пересечении; если Allow непуст, требуется полное покрытие union всех запросов. Только Deny не вводит пустой allow-list. Пустой/нераспознанный request пропускается: формат проверяет базовый check. Ошибка загрузки/INVALID → служебная ERROR, не отсутствие совпадения | allowed/denied IP, scope, comments; общий доступ к диапазонам ресурсов. [side check](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L114), [union coverage](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L974) |
| C13 Policy selectors IPv4/CIDR/range/x | selector text → positive и negative interval dictionaries + problem String / Boolean success | НЕ/NOT/! в разрешённом поле переносит в deny; в отдельной deny-колонке любой токен DENY. x и кириллическая х — только целые последние октеты; числовой октет после x invalid, значения 0..255. Это suffix octet mask, не dotted netmask. Некорректная активная policy отвергает загрузку, не молча исчезает | selector grammar, негативные условия; reusable IP-set parsing с отдельным контрактом. [ParsePolicyIpSelectors](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L798), [wildcard](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckInitiatorIpPolicy.bas#L894) |
| C14 Связанный допуск и DB account restriction | заявка ws, email, primary initiator rules, selected policy + DB-account field → CheckResult + issues | Вторичная policy не разрешает инициатора. Нет/invalid email, нет источника или primary допуска → wrapper завершает без второй ошибки email. Ошибки настройки policy отвязываются от cell email. DB запрет действует только для primary allowed и best selector; forbidden/deny/нет/no aliases требуют пустого значения; нет найденного DB field → без issue; ошибка чтения запрета возвращает False — фактическое ограничение v1 | primary source binding; DB запрет, selector; межполевые eligibility/field policy. [linked](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modInitiatorPolicyExt.bas#L14), [DB](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modInitiatorPolicyExt.bas#L73), [forbidden values](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modInitiatorPolicyExt.bas#L378) |
| C15 matched-IP ALLROWS positive gate | IP text с optional ALLROWS:, current technical row, rule, technicalRows, ipHostnameMap → clean IP + AllRows flag; разрешение публикации положительного результата | Пустое тело ALLROWS invalid. Для каждого anchor matched IP учитываются лишь строки, пересекающие anchor в scope SOURCE/DESTINATION/BOTH; каждая должна пройти то же полное правило. Нет relevant rows → False. Поддержаны IP/port/protocol; hostname/service специально исключены из hotfix. При исключении wrapper возвращает базовые issues и логирует fallback | AllRows, selector, scope, port/protocol; групповой квантификатор для совпавшего ключа. [parser](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkAllRowsHotfix.bas#L13), [supported subset](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkAllRowsHotfix.bas#L153), [group evaluation](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkAllRowsHotfix.bas#L196) |
| C16 TOP↔LOWER compact / range-list equivalence | верхние source/destination cell texts и нижние Address texts + cell provenance → strict declarations + explicit coverage sets, missing/extra issues | Full ranges/ordinary IP/compact endpoints и списки сравниваются по покрытию, порядок/дубли не влияют. CIDR остаётся строгой исходной декларацией, включая /32; invalid IPv4-like остаётся INVALID strict value. Ambiguous worksheet → SKIPPED; пустой распознанный TOP → SKIPPED; group/link отправляется legacy маршрутом | declaration-vs-coverage policy, separators, cell bindings; сверка двух произвольных таблиц. [check](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modAddressCompactCoverage.bas#L6), [normalizer](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modAddressCompactCoverage.bas#L201), [pair API](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modAddressCompactCoverage.bas#L409) |
| C17 Merge-aware table extraction / duplicates | Worksheet + locator header columns/bounds → Collection dictionaries Row, SourceAddress, DestinationAddress, PortAddress, SourceText, DestinationText, Protocol, SourceIps, DestinationIps, PortIntervals | CellArea берёт MergeArea. CollectTechnicalRows перебирает непустые физические строки в границах: не обещает полную дедупликацию merge block. Активный маршрут PortSeparatorV3 → NetworkTechnicalMergeHotfixV2 → base отдельно снимает ложные merge warnings/blank-primary errors и пересчитывает дубли через MergeFixAddDuplicateWarnings; старый RealWorld wrapper не обозначает активный путь. Пустые промежутки не завершают таблицу; ненайденная table → пустая collection | locator/границы/primary columns; reusable logical table interpretation, provenance. [extract](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkExtractor.bas#L551), [active V2](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkTechnicalMergeHotfix.bas#L4), [MergeFixAddDuplicateWarnings](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkTechnicalMergeHotfix.bas#L158) |
| C18 Effective upper+lower IPs / groups | верхние интервалы + все валидные нижние Address cells → deduplicated effective request для Superset/pending view и дополнительная policy-проверка | Текстовая группа не раскрывается без source; нижний IP нельзя считать лишним лишь из-за группы. Невалидные/неоднозначные lower cells не проходят дополнительный policy путь. Нижняя сторона не угадывается: применяются BOTH policy, отсутствующие в upper ключи; отсутствующий Superset IP сохраняется как not-found в output | group policy, BOTH scope, source binding; объединение заявленных/фактических объектов. [registry hotfix](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRealWorldHotfix.bas#L266), [lower policy](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRealWorldHotfix.bas#L366) |
| C19 hostname→IP narrowing и связанный сервис | request interval matches + lower hostname hints (включая paired lines) + registry rows → relevant row flags и serviceContext | Только найденные index rows; hostname уточняет кандидатов для данного IP, а service берётся из этого же контекста, не произвольного IP. Short/FQDN equivalence и normalization — отдельные функции. Не готовый context/нет matches → пустые наборы или False TryGet; это не доказательство отсутствия всех IP | hostname aliases, headers, mapping; join lookup с provenance. [context](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L117), [paired hints](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L409), [equivalence](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L554) |
| C20 Superset support metadata join / aggregate | результирующий Superset с № из сводки и обычными source columns + Summary № → пять колонок (Superset) | Service/full system/СКУФ ГРУППА ТП/КОМАНДА ПОДДЕРЖКИ уникально агрегируются по номеру. Hostname только при единственном FQDN; несколько → пусто. Нет sheets/№/link header → выход без записи; исключения логируются. Это output adapter v1 с прямой записью, для v2 надо разделить computation/action | source header aliases, target mapping, grouping key и uniqueness; reusable joins/aggregates. [MoveRegistryDetailToCoordinatorSummary](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L25), [unique host](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryDetailSummary.bas#L378) |

### Дополнения к C03/C04 и C15

[AnalyzeRequestPortsCompat](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkPortSeparatorHotfix.bas#L20) переводит `/` и `\` в запятую и вызывает strict AnalyzeRequestPorts. Дефис остаётся диапазоном. Вход Text → intervals + invalidTokens; пусто — пустые словари, некорректные tokens не исчезают. Admin: separators для конкретного поля, а не глобальная грамматика текста; reuse: parsing lists. Wrapper удаляет лишь ложные separator errors базовой технической проверки. Полное port coverage для positive exemption сохраняется независимо от разделителя.

ALLROWS эталоны A/B/C (обозначения кейсов аудита, не новые поля Rulebook): A — все строки одного matched IP удовлетворяют одному rule → positive result допустим; B — одна строка того же matched IP содержит непокрытый порт → positive result блокируется для этого IP; C — посторонний IP вне anchor/scope не блокирует matched IP. Строки с разными сервисами/hostname не следует автоматически включать в поддерживаемый hotfix. Без ALLROWS старое поведение сохраняется. Этот matched-key квантификатор отличается от универсального ALLROWS(empty)=TRUE, рекомендуемого будущим DSL ТЗ13 N.

### Требования, конфликты и проверка позднего слоя

ТЗ13 H–M/O описывает рабочие кейсы policy Allowed−Denied/full containment, real groups + lower IP, blank rows/merge, summary и T34–T39. ТЗ14 §7 требует общих зарегистрированных операций, SourceBinding и отсутствия ссылок на фиксированные NETWORK_ACCESS-листы внутри вычислителя. Таблица выделяет кандидатов повторного применения; не переносит wrappers в новую архитектуру и не утверждает PR-01.

ТЗ14 §7.3 NoPolicyAction=ALLOW совпадает с отсутствием selected policy в integration; открытый вопрос из ТЗ13 H устарел по этому пункту. Ошибка/нет источника — отдельный контракт и не должна маскироваться словом ALLOW. ТЗ14 endpoints-only range и compact-as-future конфликтуют с интеграционным T36/T37 и ТЗ13 O: full range/list coverage, compact support, строгий CIDR TOP↔LOWER. Требуется явное разрешение владельца; нельзя молча выбрать main или новую формулировку.

Код integration прочитан; примеры настоящих заявок здесь не публикуются. Исторические сведения ТЗ13 о 85 PASS / 0 FAIL / 0 SKIP и реальной заявке не являются текущим запуском. Ни release gates, ни Excel UI, ни миграция runtime XLSX в этом аудите не выполнялись. Дополнительные обязательные эталоны: C11 exact-over-domain/NOT; C12 split-union coverage + deny; C13 suffix x и invalid middle x; C14 ниже расположенная реальная account cell; C15 A/B/C; C16 range/list/compact и CIDR/plain; C17 merge+blank; C18 группа+found/notfound lower; C19 несколько FQDN одного IP и сервис чужого IP; C20 несколько support values и неоднозначный host; slash/backslash ports.

C08/C19 не означают доказанный port lookup в Superset: подтверждены SOURCE IP interval index, hostname narrowing и сервисный join; порт проверяется отдельно по Rulebook против технической строки. [Index builder](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modRegistryIntervalIndex.bas#L242) не задаёт port-column schema. ТЗ14 §7.2 IP+port — неподтверждённое требование, не текущий контракт. C19 narrowing сохраняет unhinted rows, а при отсутствии всех hostname matches возвращается к IP-only rows; [ветка фильтра](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modNetworkRegistryMatchContext.bas#L214).

## C21 Legacy date extraction и локальные часы

Назначение: NET.VALIDITY; переносимый кандидат date parsing + compare с независимой clock dependency. [CheckNetworkValidity](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckValidity.bas#L6), [TryExtractNetworkValidityDate](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckValidity.bas#L66), [TryBuildStrictNetworkDate](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckValidity.bas#L114).

Вход: rawValue Variant из VALIDITY.Cells(1,1).Value; нормализованный REQUEST_TYPE; наличие обоих locator ranges; локальная VBA Date компьютера. Выход parser API — Boolean success и ByRef extractedDate As Date. Выход check — CheckResult и issues. «infinite/revoke/expired» ниже являются смысловыми исходами аудита, не фактическими полями результата VBA.

Порядок: MISSING REQUEST_TYPE → SKIPPED; «отозвать» / «отозвать доступ» → PASS без чтения VALIDITY. Любой иной request type, включая EMPTY/INVALID, идёт в срок. MISSING VALIDITY → SKIPPED; Trim(CStr(rawValue)) пусто → ERROR на VALIDITY; NormalizeText «бессрочно» → PASS (infinite). Затем parser сначала вызывает IsDate: успех → DateValue(CDate(rawValue)), поэтому локаль и VBA-конверсия имеют приоритет над строгой текстовой грамматикой.

Только при IsDate=False применяется regexp `(^|[^0-9])([0-3]?[0-9])[\.\/-]([01]?[0-9])[\.\/-]([0-9]{4})([^0-9]|$)`; берётся последняя корректная дата среди фактически найденных non-overlapping matches, а не обязательно последнее вхождение, визуально похожее на дату. Strict builder проверяет год 1900..9999, месяц 1..12, день 1..31 и Year/Month/Day roundtrip после DateSerial; 31 февраля отвергается. Ограничение года относится к strict fallback, не ко всем значениям, принятым IsDate. Нет корректного match / parser exception → False. Check выдаёт ERROR формата, если parser=False; extractedDate < локальная Date → ERROR истёкшего срока; сегодня/будущее → PASS.

Ошибки различаются: EMPTY и INVALID срока — разные сообщения; MISSING locator — SKIPPED; revoke не требует срока; parser exception сводится к False, но check не имеет собственного общего error handler для каждого неподдерживаемого Variant (например, Error Variant), поэтому нельзя обещать тотальную обработку произвольного Variant как штатной INVALID. Исходный raw value и полученная дата должны быть доступны будущему explain отдельно.

Admin/параметры для инвентаризации: locator binding REQUEST_TYPE/VALIDITY, revoke aliases и infinite literal, grammar/locale policy, диапазон fallback года, сравнение «раньше сегодняшнего дня», routing/messages. Сейчас grammar/границы/aliases зашиты в код, clock и locale берутся из окружения; это не уже существующие редактируемые настройки. Reuse: сроки действия любого доступа/документа, извлечение даты из поясняющего текста, сравнение с внедрёнными часами. Для сохранения паритета нельзя заменить IsDate-first строгим ISO-only parser без отдельного решения.

Тесты: RunNetworkValidityDateRegression присутствует [в исходнике](https://github.com/kosha1308/AccessTools/blob/58d8789d1e60860c96cb35ba0e780a6424510130/src/checks/network_access/modCheckValidity.bas#L137), не запускался. Требуются epoch фиксированных часов, сегодня/вчера/завтра, revoke+invalid срок, infinite, MISSING vs EMPTY, несколько корректных/некорректных дат, locale-dependent IsDate и даты до 1900 через IsDate, Error Variant. Ручная Excel/локаль не подтверждены.
