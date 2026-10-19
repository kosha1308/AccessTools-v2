# PR-01 — матрица выразимости конструктора

## Статус и источники

Документационный проект: **17 LegacyRuleId + 26 точных LegacyPath PR-00 = 43 независимые строки покрытия**, затем второй независимый шаблон. Здесь нет исполняемых конфигураций, бизнес-правил VBA, интерпретатора или результатов выполнения.

**STRUCTURAL**: из предложенных типов, AST и чистых операций можно построить явную композицию без закрытого сетевого контроллера. Это не runtime PASS, утверждённая сигнатура, выполненный тест или перенос через Admin. **CONDITIONAL**: композиция выразима, но публикацию draft блокируют названный Q или выбор параметров. **GAP_ACTION**: для точного воспроизведения побочного эффекта не хватает согласованного action contract; описать его диагностикой недостаточно. Во всех строках technical failure не может стать PASS.

Источники: [PR-00 rules](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md), [parity](v1-parity-matrix.md), [sources](v1-sources.md), [вывод](v1-excel-output-contract.md), [findings/Q-01…Q-11](v1-audit-findings.md), [исторический каталог C01–C21](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/operation-catalog.md). Предлагаемые OP сигнатуры — [каталог PR-01](operation-catalog.md).

[ADR-007](decisions.md) и ТЗ v0.15 принимают полный включительный интервал. Q-01…Q-11 этим документом не закрываются. Q-07 оставляет TOP↔LOWER декларации/compact условными; Q-10 запрещает незаметно перенести исторический PASS при failed source. Superset MATCHED_ANY не означает IS_COVERED: .1–.10 при реестре только .1 и .10 имеет покрытие 2/10, uncovered=.2–.9.

## Общий контракт и нотация

Semantic IDs immutable uppercase; DisplayName, locator и binding меняются отдельно. RuleId NET_INITIATOR_ALLOWED сохраняет LegacyRuleId "NET.INITIATOR_ALLOWED"; LegacyPath хранит точную историческую строку. Новая семантика требует версии определения. Ссылки Rule/Expression/Source/Operation/Action/Message закрепляются immutable PublishedPackage; @1.0.0 ниже — предложенные версии примера, не опубликованный пакет.

EvalResult<T>: state OK|EMPTY|MISSING|INVALID|UNAVAILABLE|ERROR|CONFLICT|LIMIT|NOT_APPLICABLE; value есть **только при OK**. Пустой список — OK [], не EMPTY. RuleOutcome: PASS|FAIL|WARN|INFO|NEEDS_APPROVAL|NOT_APPLICABLE|CANNOT_EVALUATE. State и status различаются. Gate false → NOT_APPLICABLE; gate technical state → CANNOT_EVALUATE. ClauseKind ASSERT: true passes, false emits configured issue; EMIT: true emits action. Technical states не подменяются false или empty list. Required ошибки/недоступность источника нельзя перевести настройкой в PASS.

Это notional AST sketches, не JSON, сериализация или готовые конфигурации. Узлы: TypedLiteral, FieldRef, ColumnRef(alias), SourceRows, Call, Let, ValueRef(BindingId), Member(record,MemberId), Record(SchemaId,static members), ListExpr(ElementType,children), ContextRef, AllOf/AnyOf/Not, When, StateIs, Any/All/None/Count/Filter/Map/GroupBy/Join, RefExpression(ExpressionId@version,args). Let — lexical immutable binding; ValueRef читает его в scope. RefExpression — reusable expression с явными параметрами и DAG, не counters предыдущего запуска. ContextRef(CLOCK_DATE/LOCALE/RUN_PROFILE) получает pinned context, без live Date/Now или workbook.

Сокращения лишь для печати: L(T,v)=TypedLiteral(T,v); F(X)=FieldRef(X); C(a,X)=ColumnRef(a,X); V(X)=ValueRef(X); M(x,p)=Member(x,p); R(X,args)=RefExpression(X@1.0.0,args); EQ(a,b)=Call(OP_COMPARE@1.0.0,a,b,L(RELATION,EQ)); GT/GE/LT/NE аналогичны; N(x)=Call(OP_TEXT_NORMALIZE@1.0.0,args=[x],params={profileRef:PROFILE_TEXT_V1@1.0.0}). Ни одного нового узла через shorthand не вводится. Все OP ниже имеют proposed exact @1.0.0. Int=Int64, Date=LocalDate; PROFILE и RELATION — зарегистрированные enum/ref domains. Any(xs) без predicate в prose = GT(Count(xs),L(Int64,0)), а не canonical Any с пропущенным аргументом. OP_X(args;params {...}) = Call(OP_X@1.0.0,args,typed parameters): profiles не лишние positional аргументы. Scalar aliases коллекций читаются ValueRef; ColumnRef только record alias. Bool true/false и numeric0 в кратких sketches — TypedLiteral. Uppercase display MemberId явно соответствует contract members: MATCHES→matches,RANK→rank,SPECIFICITY→specificity,NEGATIVE_MATCHED→negativeMatched,POSITIVE→positive,NEGATIVE→negative,IS_COVERED→isCovered,COVERED→covered,UNCOVERED→uncovered,EQUAL→equal,LEFT_ONLY→leftOnly,RIGHT_ONLY→rightOnly,INVALID_TOKENS→invalidTokens. Это display mapping статической схемы, не автоматическое переименование произвольных полей. Record не сравнивается scalar OP_COMPARE: AllOf по полям либо OP_UNIQUE_VALUE с compareFields.

### Logical inputs и snapshots

NETWORK_ACCESS: TITLE, INITIATOR_FULL_NAME, INITIATOR_DEPARTMENT, INITIATOR_POSITION, INITIATOR_EMAIL, REQUEST_TYPE, SYSTEM_IEP, VALIDITY, TOP_JUSTIFICATION, DB_ACCOUNT. LABEL_OBSERVATIONS: LABEL_ID,PRESENT,provenance. TECH_ROWS: ROW_ID,SOURCE_TEXT,DESTINATION_TEXT,PROTOCOL_TEXT,PORT_TEXT,REASON_TEXT,provenance,PRIMARY_GROUP_ID. ADDRESS_ROWS: ROW_ID,SYSTEM_TEXT,ADDRESS_TEXT,HOSTNAME_TEXT,ROLE_TEXT,provenance. Resolver/TableDefinition определяет пустоту/границы/merge; FieldId не A1, provenance не передаётся как Range в операции.

Текстовые примеры используют поля Text. REQUEST_TYPE остаётся raw Text, а enum-validity проверяется отдельным assertion: это позволяет независимо описать legacy срок даже при неверном типе заявки. Для VALIDITY предлагается общий read profile `RAW_DATE_INPUT@1.0.0`: пустая cell → EMPTY; иначе OK Record{TEXT_VIEW:Text,NATIVE_DATE:LocalDate envelope}. TEXT_VIEW сохраняет текстовое представление, NATIVE_DATE=OK только для исходного Excel date value, для обычного текста EMPTY; ошибка чтения/serial conversion диагностируется отдельно, не маскируется текстом. Это generic typed raw-date view, не NET-specific reader. Native date не передаётся в OP_TEXT_NORMALIZE. Конкретная совместимость такого reader с Excel .Value/локалью — часть conditional date-profile приёмки.

SourceBindings @1.0.0 (логические схемы предложения, не заявление о наличии одноимённых физических колонок в рабочем Rulebook; физические aliases и преобразования связываются явно по PR-00. Derived ALLOW/PRIORITY/ALL_ROWS создаются открытой AST-проекцией из доступных флагов/условий либо явным literal policy, не новым скрытым загрузчиком):
- SRC_INITIATORS: ACTIVE,SELECTOR,ALLOW,APPROVAL,COMMENT,SEVERITY,PRIORITY.
- SRC_IP_POLICY: ACTIVE,SELECTOR,SCOPE,ALLOWED_TEXT,DENIED_TEXT,DB_FORBIDDEN,COMMENT,PRIORITY.
- SRC_UNIFIED_RULES: ACTIVE,IP_CONDITION,IP_SCOPE,HOST_PATTERNS,HOST_MATCH_TYPE,PORT_CONDITION,PROTOCOL_CONDITION,SERVICE_CONDITION,SERVICES,GROUPS,APPROVERS,COMMENT,SEVERITY,ALL_ROWS,PRIORITY.
- SRC_PROTOCOLS: TOKEN,ACTIVE.
- SRC_REGISTRY: RECORD_ID,IP_INTERVALS,FQDN,SERVICE,FULL_SYSTEM,SUPPORT_GROUP,SUPPORT_TEAM.
- SRC_MESSAGES: KEY,USER_TEXT,COORDINATOR_TEXT,SEVERITY,CHANNEL_FLAGS,ACTIVE,PRIORITY.

SourceRows(SRC_X@1.0.0,SNAP_RB_001) и SourceRows(SRC_REGISTRY@1.0.0,SNAP_REGISTRY_001) используют **синтетические immutable snapshot handles**. Request=SNAP_REQUEST_001; clock SNAP_CLOCK_001 (2026-10-19), locale=SNAP_LOCALE_001. В настоящем пакете нужны content digest, schema fingerprint, SourceDefinition/Binding versions, index version и resolver snapshot ID; эти значения не выдуманы как реальные хеши. Все зависимые выражения одного run используют те же snapshots. Sources lazy: FIRST_LINE без registry-dependent clauses не открывает Superset. OK [] отличается от unavailable file и invalid schema. Fallback template protocol list — отдельный pinned binding, а не hidden code.

Параллельно математическому EXPR_EFFECTIVE_IPS сохраняется List<RequestEntry>{ENTRY_ID,DISPLAY_TEXT,DECLARATIONS,IP_SET,PROVENANCE}, построенный Map/Record/OP_DISTINCT по объявленному declaration key. Это исходные уникальные записи для сводки, **не** интервалы после IP_UNION. Lookup делается по IP_SET каждой entry либо через batch union + Join candidates обратно к entry; номер сводки присваивает projection по стабильному порядку ENTRY_ID. .1 и .2 не превращаются самовольно в одну отображаемую строку .1–.2, пересекающиеся entries могут ссылаться на одну source row. OP_INTERVAL_LOOKUP.requestKey идентифицирует query value, ENTRY_ID и физические refs сохраняет вызывающий Map/Join.

### Reusable выражения с открытой композицией

**EXPR_EMAIL**: Call(OP_EMAIL_PARSE@1.0.0,args=[F(INITIATOR_EMAIL)],params={grammarProfileRef:PROFILE_EMAIL_V1@1.0.0}). EMPTY/INVALID parent email классифицируется required/format clause; failed source не становится invalid email.

**EXPR_INITIATOR_MATCHES**:
Filter(SourceRows(SRC_INITIATORS@1.0.0,SNAP_RB_001),I,
 When(EQ(C(I,ACTIVE),L(Bool,true)),
  Let(SEL,Call(OP_SELECTOR_PARSE@1.0.0,args=[C(I,SELECTOR),L(SELECTOR_KIND,EMAIL)],params={grammarProfileRef:PROFILE_INITIATOR_SELECTOR_V1@1.0.0}),
   Let(SCORE,Call(OP_EMAIL_SELECTOR_SCORE@1.0.0,args=[R(EXPR_EMAIL),M(V(SEL),POSITIVE),M(V(SEL),NEGATIVE)],params={selectorProfileRef:PROFILE_INITIATOR_SELECTOR_V1@1.0.0}),
    M(V(SCORE),MATCHES))),
  L(Bool,false))).
EXPR_PRIMARY_ALLOWED = Any(R(EXPR_INITIATOR_MATCHES),I,EQ(C(I,ALLOW),L(Bool,true))). Пустой raw selector активной строки обрабатывается schema clause, не скрыто score0. Approval-only не даёт primary допуска.

**EXPR_SELECTED_POLICY**: Filter active SRC_IP_POLICY; Map P →
Let(SEL,OP_SELECTOR_PARSE(C(P,SELECTOR),EMAIL;params grammarProfileRef),
 Let(S,OP_EMAIL_SELECTOR_SCORE(R(EXPR_EMAIL),M(V(SEL),POSITIVE),M(V(SEL),NEGATIVE)),
  Record(SCHEMA_POLICY_SCORE,{ROW:V(P),MATCHES:M(V(S),MATCHES),RANK:M(V(S),RANK),SPECIFICITY:M(V(S),SPECIFICITY),PRIORITY:C(P,PRIORITY)}))).
Filter derived MATCHES=true; OP_RANK_MAX(rows,L(List<ColumnId>,[RANK,SPECIFICITY])) возвращает все top ties; Map извлекает ROW. Exact rank2>domain1; email SPECIFICITY=0, не неизвестный length tie-break. Scope фильтруется **после** best rank. Все best applicable rows сохраняются для allow/deny union.

PRIORITY отдельное Int64; большее выше. Где business selection включает его, rankFields=[RANK,SPECIFICITY,PRIORITY] с явным direction; legacy policy профиль не добавляет PRIORITY как скрытый фильтр best-specificity union. Ties не first-row-wins. Неактивные source строки не парсятся через strict Call — explicit Filter/When gate. Required technical states сохраняют unresolved evaluation requirement.

**EXPR_TECH_PARSE**: per included row Map(F(TECH_ROWS),T,
Let(SP,OP_ADDRESS_PARSE_MIXED(C(T,SOURCE_TEXT);params grammarProfileRef),
 Let(DP,OP_ADDRESS_PARSE_MIXED(C(T,DESTINATION_TEXT);params grammarProfileRef),
  Record(SCHEMA_PARSED_TECH,{ROW_ID:C(T,ROW_ID),SOURCE_PARSE:V(SP),DESTINATION_PARSE:V(DP),PORTS:OP_PORT_PARSE(C(T,PORT_TEXT);params separatorProfileRef),PROTOCOLS:OP_TEXT_TOKENS(C(T,PROTOCOL_TEXT);params separatorProfileRef),PROVENANCE:C(T,PROVENANCE)})))).
Record создаёт derived значения, OP_PROJECT только existing member projections. Для EMPTY cell сперва state clause; parser в nonempty ветви не получает EMPTY payload. OK AddressParse с invalidTokens не valid address: Count(INVALID_TOKENS)=0 и no ambiguous — обязательные clauses. Positive IP consumers OP_IP_FROM_ADDRESS_PARSE invalidPolicy=STRICT; диагностический EXPLICIT_VALID_SUBSET только явно + форматный FAIL/обязательство полноты original input.

**EXPR_EFFECTIVE_IPS**:
OP_IP_UNION(
 OP_LIST_CONCAT(
  ListExpr(List<IpSet>,[
   Map(R(EXPR_TECH_PARSE),T,OP_IP_FROM_ADDRESS_PARSE(C(T,SOURCE_PARSE);params invalidPolicy=STRICT)),
   Map(R(EXPR_TECH_PARSE),T,OP_IP_FROM_ADDRESS_PARSE(C(T,DESTINATION_PARSE);params invalidPolicy=STRICT)),
   Map(F(ADDRESS_ROWS),A,OP_IP_FROM_ADDRESS_PARSE(OP_ADDRESS_PARSE_MIXED(C(A,ADDRESS_TEXT);params grammarProfileRef);params invalidPolicy=STRICT))
  ]))).
ListExpr содержит три List<IpSet>, CONCAT возвращает List<IpSet>, UNION получает **один** List<IpSet>. Параллельный provenance Join сохраняет все refs; mathematical union не уничтожает адреса источников. Group-only parse даёт OK∅, не состав группы.

**EXPR_POLICY_IP_SIDES**: per selected policy row P:
Let(AL,OP_SELECTOR_PARSE(C(P,ALLOWED_TEXT),IP;params grammarProfileRef),Let(DE,OP_SELECTOR_PARSE(C(P,DENIED_TEXT),IP;params grammarProfileRef),
 Record(SCHEMA_POLICY_IP_SIDES,{
 ALLOW:OP_IP_UNION(Map(M(V(AL),POSITIVE),TOK,OP_IP_PARSE(V(TOK);params grammarProfileRef))),
 DENY:OP_IP_UNION(OP_LIST_CONCAT(ListExpr(List<IpSet>,[
  Map(M(V(AL),NEGATIVE),TOK,OP_IP_PARSE(V(TOK);params grammarProfileRef)),
  Map(M(V(DE),POSITIVE),TOK,OP_IP_PARSE(V(TOK);params grammarProfileRef)),
  Map(M(V(DE),NEGATIVE),TOK,OP_IP_PARSE(V(TOK);params grammarProfileRef))
 ])))
 }))).
EMPTY allowed/denied raw field → When(StateIs(field,EMPTY),Record(SCHEMA_SELECTOR_SIDES,{POSITIVE:L(List<Text>,[]),NEGATIVE:L(List<Text>,[])}),OP_SELECTOR_PARSE(...)), чтобы optional blank не вызывал EMPTY parse. Deny column обе стороны объединяются явно; отрицательный токен в allow идёт только DENY. Selector IP grammar сохраняет диапазон целиком; suffix wildcard grammar отдельный профиль, не dotted netmask. Не скрытая политика внутри IP_PARSE.

**EXPR_RULE_MATCH(T,R)**: AllOf пяти независимых Bool expressions:
- When(StateIs(C(R,IP_CONDITION),EMPTY),true,OP_IP_OVERLAPS(scope-selected valid requested IPs,OP_IP_PARSE(C(R,IP_CONDITION);params grammarProfileRef))).
- When(StateIs(C(R,PORT_CONDITION),EMPTY),true,M(OP_PORT_COVERAGE(requested ports,OP_PORT_PARSE(rule ports;params grammarProfileRef)),IS_COVERED)).
- When(StateIs(C(R,PROTOCOL_CONDITION),EMPTY),true,Any(request protocol tokens,Q,OP_LIST_CONTAINS(rule tokens,V(Q)))).
- When(StateIs(C(R,HOST_PATTERNS),EMPTY),true,Any(IP-related hostname hints,H,M(OP_HOSTNAME_MATCH_SCORE(C(H,HOSTNAME),rule pattern List<Text>,C(R,HOST_MATCH_TYPE);params normalizeProfileRef),MATCHES))).
- When(StateIs(C(R,SERVICE_CONDITION),EMPTY),true,Any(IP-related registry context,S,OP_LIST_CONTAINS(rule service tokens,C(S,SERVICE)))).
Missing/invalid/error condition не EMPTY wildcard. Service context — Join/Filter matched IP+hostname, не все source rows; invalid-subset не снимает original format FAIL. Это открытая AST-композиция, не OP_EVAL_RULEBOOK/POLICY.

**EXPR_HOST_OUTPUTS**: Map matched rows через Record(MATCHES,RANK,SPECIFICITY,ROW,PRIORITY) с Member(OP_HOSTNAME_MATCH_SCORE,…); OP_RANK_MAX по [RANK,SPECIFICITY] или явно принятому priority tuple. Затем Map(top,S,OP_PROJECT(M(V(S),ROW),mapping={SERVICES:SERVICES,GROUPS:GROUPS})) даёт records с непосредственными SERVICES/GROUPS для OP_UNIQUE_VALUE(rows,[SERVICES,GROUPS];params normalization profile). CONFLICT отдельная state clause. Broad COMMENT/APPROVERS Map всех matches→OP_DISTINCT сохраняются. Severity rank maximum по шкале INFO10/WARNING20/ERROR30 всех visible outputs, не alphabetical string max. Tuple ranking — проектный профиль; точный legacy `rank*100000+length`, если выбран, строится OP_NUMBER_BINARY с явными MUL/ADD и rankField SCORE, а не объявляется автоматически эквивалентным tuple при любой длине.

Во всех shared projections StateIs обрабатывает EMPTY/INVALID input до dependent strict Call. If значение требуется downstream, MISSING/ERROR не превращается в empty branch success; unresolved required completeness сохраняется. Состояние AddressParse OK говорит об извлечённом разборе, не об успешной проверке всех элементов.

### Действия и повтор

ActionDef — декларативные EmitIssue, Annotate, ReportRow, SourceView; это не Call операции. ActionId@version, Rule/Clause provenance, MessageTemplate@version, параметры текста через OP_TEXT_FORMAT, target FieldId или TableId+RowId+ColumnId+resolver provenance. Physical адрес находит Planner/Writer. Severity независимо от каналов REQUEST_ANNOTATION/RESULTS_WINDOW/USER_REPORT/COORDINATOR/SOURCE_VIEW и ролей FIRST_LINE/COORDINATOR.

В карточках ISSUE(code,target,status) сокращает EmitIssue и разрешённый Annotate; ROW(report,mapping)=ReportRow; VIEW(source,recordIds)=SourceView. Shared planner отвечает за ownership, merge/dedup/idempotence; Evaluator не удаляет issue и не меняет counters. Legacy suppression записывается predicate/gate до actions. Numeric rule/clause priorities, explicit DAG и stable clause IDs обеспечивают порядок, а не VBA dispatch order. Все тексты/targets/channel flags редактируются Admin; при source failure/technical state блокируется ложный PASS. Неизвестные actual Rulebook тексты Q-02 остаются unknown.

## 17 основных правил

### 01 — NET_TEMPLATE

LegacyRuleId NET.TEMPLATE; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#nettemplate).

- Inputs TITLE Text как optional locator observation и LABEL_OBSERVATIONS List; expected title/13 label IDs из TemplateDefinition@version. Gate recognition OK. Отсутствие optional title классифицировано WARN; ошибка чтения/неоднозначность required recognition даёт CE, а не «title отсутствует».
- AST: EMIT StateIs(F(TITLE),MISSING); ASSERT EQ(N(F(TITLE)),N(L(Text,expectedTitle))) при OK; Filter(labels,L,EQ(C(L,PRESENT),false)) → Count>0 EMIT. Нет title и missing labels — два независимых clauses.
- Actions WARN TITLE или document target; structure text с missing labels; другие rules продолжаются. Admin title/labels/priority/text/channel/target.
- Witness нет title и IP source label → два WARN; absent title не даёт физического адреса. Resolver/manual Q-04 не проверен.

### 02 — NET_REQUIRED_FIELDS

LegacyRuleId NET.REQUIRED_FIELDS; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netrequired_fields).

- Inputs четыре INITIATOR_* поля, source none, gate true. Отдельный clause на поле, не format validation ФИО.
- AST EMIT StateIs(F(X),EMPTY) → FAIL; EMIT StateIs(F(X),MISSING) → historical WARN/document target плюс техническая completeness diagnosis. OK любое непустое принято.
- Actions required ISSUE на X; Admin required field list, text/routing/target. Missing mandatory locator не может hidden warning превратить общий completeness в PASS.
- Witness четыре EMPTY → четыре FAIL; MISSING full name → WARN без cell + CE completeness. Не проверено отображение target, Q-04/Q-10 политики missing остаются открыты.

### 03 — NET_INITIATOR_EMAIL

LegacyRuleId NET.INITIATOR_EMAIL; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netinitiator_email).

- Input INITIATOR_EMAIL; source none; gate Not(StateIs(email,EMPTY)). MISSING CE.
- AST Let(E,R(EXPR_EMAIL),ASSERT(StateIs(V(E),OK))); INVALID parse → format FAIL, technical ERROR/LIMIT → CE, не invalid input.
- Actions EMAIL_FORMAT на email; Admin grammar/profile version, binding, text/channels. Default preserves legacy regexp, не обещает RFC полноту.
- Witness a@b FAIL; a@example.test PASS; EMPTY → NOT_APPLICABLE format clause, required всё равно FAIL. Исторически format-check EMPTY PASS, здесь статус granular clause отличается явно.

### 04 — NET_INITIATOR_ALLOWED

LegacyRuleId NET.INITIATOR_ALLOWED; **CONDITIONAL Q-02,Q-03,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netinitiator_allowed).

- Inputs email, SRC_INITIATORS ACTIVE/ALLOW/SELECTOR snapshot. Gate valid email.
- AST Filter(active allow rows) → Count; если >0 ASSERT R(EXPR_PRIMARY_ALLOWED); no active rows отдельная «not configured» CE, не unmatched denial.
- Actions unmatched existing list → FAIL email; unavailable source CE. Admin selectors/active/ALLOW/profile/priority/no-rows business policy/text/channel.
- Witness active example.test, u@other.test → FAIL; source missing → CE. No-policy ALLOW вторичного IP check не подменяет primary closed list.

### 05 — NET_INITIATOR_APPROVAL_RULES

LegacyRuleId NET.INITIATOR_APPROVAL_RULES; **CONDITIONAL Q-02,Q-05,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netinitiator_approval_rules).

- Inputs email, SRC_INITIATORS APPROVAL/COMMENT/SEVERITY; gate valid email. ALLOW не gate: approval-only applies.
- AST Filter(R(EXPR_INITIATOR_MATCHES),I,StateIs(C(I,APPROVAL),OK)); per entity EMIT true; OP_TEXT_FORMAT(MessageTemplateValue pinned на MessageDefinition@version,Record(SCHEMA_APPROVAL_TEXT,{EMAIL:R(EXPR_EMAIL),APPROVAL:C(I,APPROVAL),COMMENT:C(I,COMMENT)})); numeric severity rank отдельно. Optional EMPTY comment заранее преобразуется явно в literal пустой Text либо исключается optional placeholder branch; техническая ошибка member не скрывается.
- Actions NEEDS_APPROVAL на email; historical USER_REPORT=false, REQUEST_ANNOTATION=true независимы. Admin Approval/Comment/severity/profile/text/targets/channels.
- Witness две matched approval строки → две actions, rule NEEDS_APPROVAL; no approvals OK [] → PASS; unavailable ≠ []. Routing Q-05 не утверждён.

### 06 — NET_REQUEST_TYPE

LegacyRuleId NET.REQUEST_TYPE; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netrequest_type).

- Input REQUEST_TYPE, no source, true gate.
- ASSERT OP_LIST_CONTAINS(L(List<Text>,[предоставить доступ,отозвать доступ,предоставить,отозвать]),N(F(REQUEST_TYPE))); EMPTY required FAIL, MISSING CE.
- Action FAIL REQUEST_TYPE; Admin literals/aliases/binding/text/roles/channels.
- Witness «отозвать» PASS, «продлить» FAIL. Исторический missing SKIPPED не target PASS.

### 07 — NET_SYSTEM_IEP

LegacyRuleId NET.SYSTEM_IEP; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netsystem_iep).

- SYSTEM_IEP Text, no source. ASSERT EQ(N(F(SYSTEM_IEP)),L(Text,да)); EMPTY FAIL, missing/error CE.
- Action FAIL SYSTEM_IEP, сообщение о регистрации первой линией. Это не доказательство recipient FIRST_LINE.
- Admin expected literal, binding/title/text/channels/priority.
- Witness «Нет» FAIL, «Да» PASS. Technical error не преобразуется в «Нет».

### 08 — NET_VALIDITY

LegacyRuleId NET.VALIDITY; **CONDITIONAL date profile/locale,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netvalidity).

- REQUEST_TYPE, VALIDITY raw scalar, ContextRef(CLOCK_DATE/LOCALE), no source. Gate NOT revoke aliases; historical invalid request type тоже non-revoke.
- AST: сначала EMIT StateIs(F(VALIDITY),EMPTY) → FAIL. Только при OK raw-date record Let(T,M(F(VALIDITY),TEXT_VIEW),When(EQ(N(V(T)),L(Text,бессрочно)),true,Let(D,When(StateIs(M(F(VALIDITY),NATIVE_DATE),OK),M(F(VALIDITY),NATIVE_DATE),When(StateIs(M(F(VALIDITY),NATIVE_DATE),EMPTY),OP_DATE_PARSE(V(T);params mode=LEGACY_ISDATE_FIRST,localeRef=ContextRef(LOCALE)),M(F(VALIDITY),NATIVE_DATE))),OP_DATE_COMPARE(V(D),ContextRef(CLOCK_DATE),GE))))) даёт Bool для ASSERT. Ветка text parse допустима при EMPTY native date; INVALID/ERROR native read не выбирают успешный text fallback. INVALID parsed date → format FAIL, ERROR/LIMIT → CE. Истечение, формат и пустота имеют разные clauses/codes, а не один boolean fallback.
- Actions three codes EMPTY/expired/format target VALIDITY. Admin revoke aliases/infinite/date grammar+locale policy/clock/messages/channels.
- Witness revoke+garbage → N/A; clock19.10.2026, date18 → FAIL,19 → PASS;31.02 invalid. 1900..9999 only strict fallback, IsDate-first может принять иное; ISO-only silently не parity.

### 09 — NET_TOP_JUSTIFICATION

LegacyRuleId NET.TOP_JUSTIFICATION; **STRUCTURAL**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#nettop_justification).

- TOP_JUSTIFICATION, no source; EMIT StateIs(F(TOP_JUSTIFICATION),EMPTY) → WARN; OK nonempty PASS; missing resolver CE.
- Action WARN на поле, не blocking FAIL; Admin optional/required policy, text/target/channel.
- Witness EMPTY WARN с «не блокирует исполнение», nonempty PASS.
- Gap: optional absence vs mandatory locator resolution должны быть явны, а не inferred из NotFound Excel.

### 10 — NET_TECHNICAL_TABLE

LegacyRuleId NET.TECHNICAL_TABLE; **CONDITIONAL Q-04,Q-07,parser profiles**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#nettechnical_table).

- TECH_ROWS, SRC_PROTOCOLS snapshot; fallback literal list отдельный pinned binding. Gate table resolver OK. Used rows = Filter(any five fields nonempty), не StopAtFirstBlank.
- AST R(EXPR_TECH_PARSE); независимые source/destination EMPTY/ambiguous/invalid clauses; ASSERT port parse OK+Count(intervals)>0; All(protocol tokens,P,OP_LIST_CONTAINS(active tokens,V(P))); reason EMPTY WARN. Count(rows)=0 → four required clauses. Duplicates GroupBy normalized [source,destination,protocol,port], Count>1, reason вне key.
- Actions corresponding columns; subsequent duplicate SOURCE_TEXT with original RowId UiGroup; quiet group profile suppresses only INFO, composition groups unknown. Admin required columns/protocol binding/fallback/port separators/duplicate key/reason/merge profile/text/channels.
- Witness wholly empty → four FAIL; 80/443 compat OK; 0 invalid; .1–.10 includes .5. Resolver merge Q-04 и compact Q-07 не решены; suppression до ActionPlan, не remove/recount.

### 11 — NET_ADDRESS_TABLE

LegacyRuleId NET.ADDRESS_TABLE; **CONDITIONAL Q-04,parser profile**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netaddress_table).

- ADDRESS_ROWS; no source; gate resolver OK. Used row AnyOf four nonempty.
- AST Map rows → OP_ADDRESS_PARSE_MIXED; system/role EMPTY WARN; address/hostname EMPTY FAIL; ambiguous/invalid/no recognized address FAIL. GroupBy logical field provenance before actions avoids merged duplicate; group-only quiet, not resolved set.
- Actions empty table system WARN + address/hostname FAIL first-entry logical targets; no invented role warning for absent row. Admin mandatory columns/parser profile/dedup/text/channel.
- Witness row only SYSTEM → address/hostname FAIL, role WARN. Group accepted as group does not grant arbitrary IP; physical merge UI Q-04 unverified.

### 12 — NET_ADDRESS_MATCH

LegacyRuleId NET.ADDRESS_MATCH; **CONDITIONAL Q-07,Q-04**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netaddress_match).

- TECH source/destination, ADDRESS address, parse diagnostics/groups/aux links; gate no ambiguous; blocking group → cannot compare, not PASS.
- AST для валидных объявлений: OP_DECLARATIONS_FROM_ADDRESS_PARSE обеих сторон → OP_DECLARATIONS_COMPARE(left,right;params declarationProfileRef=PROFILE_DECLARATIONS_V1@1.0.0); OP_IP_DIFFERENCE(topExplicit,lowerExplicit) и наоборот для ordinary IP/ranges. Raw invalid keys сравниваются отдельным деревом пути MATCH.invalid, не строгим parser API. Map differences сохраняет каждую provenance; groups/aux flags — отдельные clauses; fallback выбран явным профилем, не закрытым controller.
- Actions lower extra FAIL ADDRESS_TEXT; upper missing below FAIL source/destination; aux INFO document. Admin declaration-vs-coverage profile, grammar/group policy/text/channel/targets.
- Witness range .1–.10 differs endpoints list .1,.10 (ADR007). Legacy strict IP vs /32 differs; explicit range vs full list equal by coverage. Default declaration/compact Q-07 remains conditional, not approved by this sketch.

### 13 — NET_IP_REGISTRY

LegacyRuleId NET.IP_REGISTRY; **CONDITIONAL Q-02,Q-06,Q-10,Q-11**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netip_registry).

- R(EXPR_EFFECTIVE_IPS), SRC_REGISTRY pinned rows/index; gate no ambiguous, selected channels need source. Port binding absent: no invented registry-port lookup.
- AST OP_INTERVAL_LOOKUP(requested,SourceRows(SRC_REGISTRY@1.0.0),L(ColumnId,IP_INTERVALS)) → typed matches с source rows и matched IpSet; Map(matches,M,M(V(M),MATCHED)) → OP_IP_UNION → matchedUnion; OP_IP_COVERAGE(requested,matchedUnion) → covered/uncovered. Any(matches)=MATCHED_ANY отдельно. Join provenance+IDs; hostname Filter+OP_HOSTNAME_EQUIVALENT сохраняет unhinted/fallback IP-only.
- Actions historical no-match WARN all refs; ReportRow found/partial/not-found + covered/uncovered; SourceView IDs; Join support metadata. Partial severity/report policy Q-06 unresolved, false full-found forbidden ADR007. Admin bindings/headers/index/narrowing/text/channels/mappings.
- Witness request .1–.10 registry .1,.10 → MATCHED_ANY true, IS_COVERED false, uncovered .2–.9, no full-found. This expands explicit coverage reporting beyond historical overlap-only warning; cannot call it unchanged legacy behavior.

### 14 — NET_IP_EXCEPTIONS

LegacyRuleId NET.IP_EXCEPTIONS; **CONDITIONAL Q-02,Q-03,Q-05,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netip_exceptions).

- Parsed TECH_ROWS, ADDRESS hints, SRC_UNIFIED_RULES; SRC_REGISTRY lazy only SERVICE_CONDITION. Gate required snapshots OK; safe/ALLROWS historical gates differ, no silent rewrite.
- AST Join(T rows,R active rules,R(EXPR_RULE_MATCH,T,R)); all filled conditions AND. Service Join only matched IP/hostname; absent context technical clause. EXPR_HOST_OUTPUTS selects structured top, all comments/approvals retained. ALL_ROWS When(flag,All(Filter(rows,S,OP_IP_OVERLAPS(anchor,scope IPs)),S,R(EXPR_RULE_MATCH,S,R)),true) plus Count(relevant)>0; supported subset no host/service.
- Actions before plan gate positive exemption; target matched IP else PORT_TEXT if port-condition else DESTINATION_TEXT. Conflict issue; severity rank max visible row outputs; GroupBy severity+target+message dedup. Admin all conditions/scope/ALL_ROWS/specificity/priority/results/severity/text/channels.
- Witness same IP ports80,443 with ALLROWS80 → no positive exemption; unrelated IP outside anchor does not block. Error CE, not historical base fallback silent success Q-10.

### 15 — NET_HOSTNAME_RULES

LegacyRuleId NET.HOSTNAME_RULES; **CONDITIONAL Q-02,Q-05**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#nethostname_rules).

- ADDRESS HOSTNAME_TEXT, SRC_UNIFIED_RULES; gate active hostname-only rows without technical/service conditions.
- AST OP_TEXT_TOKENS profile newline/comma/pipe/trailing dot; Join score>0; GroupBy host+provenance; R(EXPR_HOST_OUTPUTS). Specificity exact>prefix>contains then length; numeric priority separate explicit tie policy; unique structured output CONFLICT if incompatible.
- Actions hostname ISSUE, ReportRow service/group, NEEDS_APPROVAL outputs independently; broad comments/approvals OP_DISTINCT retained; severity numeric maximum. Admin patterns/matchtype/specificity/priority/compareFields/output/text/channels.
- Witness exact serviceA + broad approvalB → both; equal exact serviceA/serviceB → CONFLICT with no arbitrary service. OK [] no matches PASS; source failure CE.

### 16 — NET_INITIATOR_IP_ACCESS

LegacyRuleId NET.INITIATOR_IP_ACCESS; **CONDITIONAL Q-02,Q-03,Q-07,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netinitiator_ip_access).

- Email, primary allowed, selected policy; SRC_INITIATORS/SRC_IP_POLICY same RB snapshot, TECH + lower-only ADDRESS. Gate primary allowed; failed source CE, not assumed false.
- AST R(EXPR_SELECTED_POLICY); Filter by side scope; Map R(EXPR_POLICY_IP_SIDES,row) извлекает positive/negative raw selectors, затем allow/deny IpSet; OP_IP_UNION(Map(sideRecords,Member ALLOW)) и аналогично DENY. Per request ASSERT AllOf(Not(OP_IP_OVERLAPS(request,deny)),When(GT(OP_IP_CARDINALITY(allow),L(Int64,0)),Member(OP_IP_COVERAGE(request,allow),IS_COVERED),true)). Lower-only extra uses absent upper keys and ANY/BOTH selection separately.
- Actions FAIL source/destination or lower ADDRESS_TEXT; config CE document target. OK [] selected = no secondary restriction, primary still required. Admin NOT/selectors/score/tie/scope/allow/deny/comment/priority-profile/text/channels.
- Witness request .1–.10 allowed .1–.8 => FAIL .9–.10 uncovered; deny .5 fails even in allowed /24; adjacent allowed intervals union covers → PASS. No source ≠ no selected policy. Q-03 inheritance/DB association unresolved.

### 17 — NET_INITIATOR_DB_ACCOUNT

LegacyRuleId NET.INITIATOR_DB_ACCOUNT; **CONDITIONAL Q-02,Q-03,Q-10**. [PR00](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-business-rules.md#netinitiator_db_account).

- Email, DB_ACCOUNT logical field; SRC_INITIATORS/IP_POLICY DB_FORBIDDEN. Gate primary allowed; actual requested DB-port not historical condition.
- AST selected best score first; Any(selected,P,OP_LIST_CONTAINS(L(List<Text>,[запрещено,запрещен,запрещён,forbidden,deny,denied,нет,no]),N(C(P,DB_FORBIDDEN)))) → ASSERT StateIs(F(DB_ACCOUNT),EMPTY). Any forbidden in best ties wins; lower-score domain discarded before DB flag.
- Actions nonempty FAIL DB_ACCOUNT; read/missing field historical PASS explicitly unresolved Q-10, target CE not silent success. Rulebook schema creation/Save not operation side effect; separate Admin/source maintenance flow.
- Admin forbidden aliases/profile, binding/selectors/priority/messages/channels. Witness exact unrestricted overrides domain deny; equal-best one deny + svc_test → FAIL; source failed CE. Adding DB-port applicability would require new business decision.

## 26 дополнительных путей — отдельные clauses/projections

Имена LegacyPath ниже скопированы точно из PR-00; это 26 строк, а не новые LegacyRuleId. Semantic IDs здесь принадлежат Clause/Projection/validation entities. Все inherits общие snapshots/context, typed-state contract и ActionDef; номер source-строки соответствует immutable PR-00, а не изменяемой копии документа.

### 18 — NET_TEMPLATE_TITLE_STRUCTURE

LegacyPath "TEMPLATE.title + structure"; parent NET_TEMPLATE; **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L49).

Inputs TITLE и LABEL_OBSERVATIONS; Template expected labels. Gate recognition OK. AST AllOf не объединяет факты в одну ошибку: два EMIT predicates StateIs(TITLE,MISSING) и GT(Count(Filter(labels,PRESENT=false)),0). ASSERT title version только при OK title. Actions два WARN, target document при absent title. Admin independent messages/target fallback/labels/priority. Witness no title+2 labels absent → два issues, list labels внутри второго; не один ambiguous «плохая форма». Gap missing required values completeness остаётся отдельной technical diagnosis.

### 19 — NET_REQUIRED_MISSING

LegacyPath "REQUIRED.missing"; parent NET_REQUIRED_FIELDS; **CONDITIONAL Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L50).

Input envelope каждой mandatory FieldRef; no source, gate true. EMIT StateIs(F(X),MISSING) создаёт legacy NoCell WARN и technical incomplete; StateIs(EMPTY) остаётся другой clause с FAIL. Admin label locator, missing message и document target; не позволяет missing required locator=PASS. Witness ФИО label отсутствует → нет ложной подсветки пустой соседней ячейки, incomplete CE. Q-10: сохранение исторического WARN vs технического общего status требует решения, но отсутствие значения не доказано успешным.

### 20 — NET_EMAIL_EMPTY

LegacyPath "EMAIL.empty"; parent NET_INITIATOR_EMAIL; **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L51).

Input INITIATOR_EMAIL; source none. Gate Not(StateIs(email,EMPTY)); required clause independent dependency by snapshot, не чтение previous result. AST When(StateIs(email,EMPTY),NOT_APPLICABLE,format expression). Это status annotation, не expression literal enum вместо Boolean. Actions format suppressed, required FAIL остаётся. Admin разделение mandatory/format и texts; technical state не выбирает empty branch. Witness EMPTY → один required issue, ноль EMAIL_FORMAT. Исторический PASS формата явно отображён как nonapplicability target, не «поле правильно».

### 21 — NET_APPROVAL_USER_REPORT

LegacyPath "APPROVAL.user-report"; parent NET_INITIATOR_APPROVAL_RULES; **CONDITIONAL Q-05,Q-02**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L52).

Inputs matched Approval+Comment+Severity, SRC_MESSAGES channel flags. Gate valid matched approval. AST EMIT approval-present; routing separate from Boolean conditions; OP_TEXT_FORMAT creates two message payloads for initiator/coordinator, not hide whole issue. Actions NEEDS_APPROVAL, Annotate EMAIL true, USER_REPORT false fallback; RESULTS/COORDINATOR independently specified. Admin per-channel flags/templates; severity does not infer recipient. Witness approval required → visible email note, not user-error list under historical fallback. Catalog override fixture absent Q-02; target channels Q-05 remain conditional.

### 22 — NET_VALIDITY_REVOKE

LegacyPath "VALIDITY.revoke"; parent NET_VALIDITY; **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L53).

Input REQUEST_TYPE only for gate, no source; N(type) in literal revoke aliases → gate false. AST Not(OP_LIST_CONTAINS(revoke aliases,N(F(REQUEST_TYPE)))) guards date parse and its dependencies; validity not needed. Action no date issue/annotation; rule N/A vs historical PASS, reason explicit. Admin aliases and applicability; no hidden call to today's date. Witness revoke + invalid/missing VALIDITY → no date parse required; MISSING REQUEST_TYPE cannot pretend revoke and remains CE.

### 23 — NET_TECH_NO_USED_ROWS

LegacyPath "TECH.no-used-rows"; parent NET_TECHNICAL_TABLE; **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L54).

Inputs resolver OK TECH_ROWS, primary five column inclusion definition; no registry. AST EQ(Count(Filter(rows,any business data)),0), four independent EMIT clauses. Actions FAIL source/destination/protocol/port at FIRST_ENTRY_SLOT logical target supplied by TableDefinition; row need not physically exist as data record. Admin empty-table messages/target slot and mandatory columns. Witness OK [] → four FAIL issues, not vacuous All([])=true PASS. Missing header MISSING → CE, not empty table.

### 24 — NET_TECH_GROUPS

LegacyPath "TECH.groups"; parent NET_TECHNICAL_TABLE; **CONDITIONAL Q-02,Q-08**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L55).

Inputs AddressParse.GROUPS for TECH source/destination; no group membership source. AST Any(groups) EMIT INFO in historical-main profile, EMIT predicate false in quiet integration output profile; group value retained as unknown content. Actions no pointless INFO in integration projection, format failures independent. Admin profile/channel/message; cannot label unknown group resolved/allowed. Witness GROUP_A alone accepted text group; no composition checking, explicit lower IP still proceeds through separate paths. Baseline choice Q-01/Q-08 and operational text Q-02 remain separate from expressibility.

### 25 — NET_TECH_DUPLICATE

LegacyPath "TECH.duplicate"; parent NET_TECHNICAL_TABLE; **CONDITIONAL Q-04**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L56).

Inputs logical TECH row/provenance/PRIMARY_GROUP_ID and normalized source,destination,protocol,port; no source. AST Map→OP_PROJECT key fields; GroupBy four keys; Filter Count>1; OP_SORT by snapshot row order; first row original, subsequent ForEach EMIT. Reason excluded. Merge whole-block members filtered using explicit TableDefinition grouping before duplicate computation. Actions WARN subsequent SOURCE_TEXT, UiGroup original RowId; Admin key/rank/order/target/text/merge profile. Witness two equal rows different reason → duplicate; same whole-primary merge block → not duplicate. No scalar record EQ needed; physical interpretation Q-04 unresolved.

### 26 — NET_TECH_MERGE

LegacyPath "TECH.merge"; parent NET_TECHNICAL_TABLE; **CONDITIONAL Q-04,Q-08**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L57).

Input resolver merge topology metadata, primary columns and row inclusion predicates; no operation accepts Worksheet. TableDefinition chooses whole-primary block logical row vs partial inherited cells. AST Filter(rows,any primary business data) and GroupBy(PRIMARY_GROUP_ID) feed normal clauses; no special EvaluateNetworkMerge. Historical V2 removed legacy merge WARN/blank-primary errors; target never generates those findings when configured inclusion says non-record. Actions duplicates/required after projection only. Admin primary cols/merge interpretation/empty boundaries/target policy. Witness visually blank spacer/reason-only merge not new connectivity; valid two real rows remain two. Unknown adapter extraction capability is honest schema gap, publication blocked until resolved.

### 27 — NET_TECH_PORT_SEPARATORS

LegacyPath "TECH.port-separators"; parent NET_TECHNICAL_TABLE; **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L58).

Input PORT_TEXT; source none, gate row included. AST Let(P,OP_PORT_PARSE(text,PROFILE_PORT_COMPAT_V1),ASSERT(StateIs(V(P),OK))) with params separators include slash/backslash; hyphen is range, 1..65535. Pure parser rather than generating error then deleting it. Actions invalid FAIL port, OK no issue; Admin separator profile/range grammar/text/target, scope only port columns. Witness 80/443 and 80\443 OK two ports, 8000-8100 interval, 0 invalid. Does not alter IPv4 CIDR slash or arbitrary text globally.

### 28 — NET_ADDRESS_GROUPS

LegacyPath "ADDRESS.groups"; parent NET_ADDRESS_TABLE; **CONDITIONAL Q-08**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L59).

Input ADDRESS_TEXT mixed parse; no source resolving groups. AST group count>0 and diagnostics clauses separate; quiet integration suppresses INFO action, not Group state/provenance. Address required asserts recognized IP or group; no false full IpSet extraction from group. Admin profile/channel/group messages; runtime group resolution needs separate declared source if added later. Witness GROUP_A alone no informational issue in integration, hostname required clause still applies. Composition unknown flagged in address-match scope, not silently allowed arbitrary addresses.

### 29 — NET_ADDRESS_MERGED_DEDUP

LegacyPath "ADDRESS.merged-dedup"; parent NET_ADDRESS_TABLE; **CONDITIONAL Q-04**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L60).

Inputs ADDRESS field-provenance canonical group IDs from resolver, no source. AST GroupBy([FIELD_ID,PROVENANCE_GROUP_ID]) → one evaluated field instance; system/address/hostname/role different field IDs retain independent clauses. ActionPlanner canonical target grouping subsequently combines issues, never evaluator writes one note. Admin dedup scope/merge read-target policy/text/channels. Witness one merged hostname spanning 3 rows → one hostname-empty issue, not three; two separate physical fields same text remain two issues. Cannot use OP_DISTINCT(text) because that destroys address provenance; Q-04 manual pending.

### 30 — NET_MATCH_INVALID

LegacyPath "MATCH.invalid"; parent NET_ADDRESS_MATCH; **CONDITIONAL Q-07**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L61).

Inputs AddressParse invalidTokens/items+raw provenance обеих таблиц; source none. AST Map(invalidTokens,T,Record(SCHEMA_INVALID_KEY,{KIND:L(Text,INVALID),RAW:M(V(T),RAW_TOKEN),PROVENANCE:M(V(T),PROVENANCE)})); OP_DISTINCT по KIND/RAW сохраняет refs. Lower unmatched = Filter(lowerInvalid,L,None(upperInvalid,U,EQ(C(L,RAW),C(U,RAW)))), upper unmatched симметрично. Это scalar Text comparison исходных невалидных токенов, **не** OP_DECLARATIONS_FROM_ADDRESS_PARSE: strict operation при invalidTokens вернёт INVALID. Валидные items можно отдельно спроецировать Record(SCHEMA_ADDRESS_PARSE,{ITEMS:validItems,GROUPS:groups,LINKS:links,INVALID_TOKENS:literal empty list}) с сохранённым original parse provenance/обязательством полноты и применить strict declaration extraction лишь к этому обозначенному subset. Actions format FAIL исходного invalid остаётся; mismatch FAIL по unmatched raw keys — отдельные targets. Admin raw-key normalization, comparison/profile/message/routing; Q07 открыт. Witness 10.0.0.9999 против 10.0.0.9 не становится частично распознанным равенством; тот же invalid token с обеих сторон не даёт mismatch, но format FAIL остаётся. Overall invalid request не получает полного PASS.

### 31 — NET_MATCH_AUX_LINK

LegacyPath "MATCH.aux-link"; parent NET_ADDRESS_MATCH; **CONDITIONAL Q-07,Q-05**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L62).

Inputs parsed IP/declarations/AUXILIARY_LINKS/BLOCKING_GROUP flags; no external URL fetching. AST AllOf(no blocking groups,Any(auxiliaryLinks)) guards comparison + separate EMIT true INFO. Compare IP sets/declarations per profile; link text excluded from equality, not declared verified. Action INFO document-level, code ADDRESS_MATCH_LINKS; per-IP differences retain normal targets. Admin mixed-content grammar, info template/channels, declaration profile. Witness 10.0.0.1+URL above and same IP below → compare succeeds, INFO explicitly says link not checked. Pure text group still blocks full comparison; cannot equate group and link.

### 32 — NET_MATCH_CIDR

LegacyPath "MATCH.CIDR"; parent NET_ADDRESS_MATCH; **CONDITIONAL Q-07**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L63).

Inputs typed CIDR declarations, lower actual IPs, groups; no registry. AST When(profile group/CIDR fallback,OP_IP_COVERAGE(lowerActual,upperCIDR).IS_COVERED,OP_DECLARATIONS_COMPARE(strict TOP,strict LOWER).EQUAL) with explicit Member and Let, never one generic equality. Conditional pattern is a selectable documented alternative, not approved default. Actions missing/extra FAIL target relevant row, no incorrect «fully matched» for unknown group coverage. Admin exact CIDR/base/type profile and fallback policy. Witness plain IP vs /32 differs under strict declarations; actual lower IP in top CIDR permitted only chosen fallback; mathematical covered does not settle declarative Q-07.

### 33 — NET_REGISTRY_LOWER_ONLY

LegacyPath "REGISTRY.lower-only"; parent NET_IP_REGISTRY; **CONDITIONAL Q-06,Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L64).

Inputs lower intervals+provenance, upper keys, SRC_REGISTRY; gate source OK. AST Filter lower absent exact upper key (Join anti membership over declaration/interval key); OP_INTERVAL_LOOKUP + Count matches=0 EMIT WARN. Coverage computed separately, partial≠full. Actions WARNING each lower ref, ReportRow found/partial/not-found; missing source CE. Admin key dedup definition, binding/no-match text/channels/partial policy. Witness IP only in lower not registry → lower WARN even if upper group; lower-only found → no no-match issue. Do not silently omit actual lower addresses.

### 34 — NET_REGISTRY_CONTEXT

LegacyPath "REGISTRY.context"; parent NET_IP_REGISTRY; **CONDITIONAL Q-06,Q-09**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L65).

Inputs effective TOP+LOWER IpSet/provenance, registry snapshots/hints. AST OP_IP_UNION→OP_INTERVAL_LOOKUP→Join match IDs+source rows; hostname Filter retains unhinted/IP fallback; Map→OP_PROJECT report columns; GroupBy summary number, OP_DISTINCT/OP_TEXT_JOIN aggregate support fields; OP_UNIQUE_VALUE FQDN controls hostname uniqueness. Action SourceView selected record IDs, ReportRow summary metadata; no direct pending workbook manipulation. Admin projection mapping, output order, no-match/partial/ambiguous-host policy, lazy channels. Witness lower-only found ID included source view and coordinator support metadata; multiple FQDN keep all matched source rows but do not fabricate unique hostname. Global messages Q09 not discarded.

### 35 — NET_EXCEPTIONS_SERVICE_UNAVAILABLE

LegacyPath "EXCEPTIONS.service-unavailable"; parent NET_IP_EXCEPTIONS; **CONDITIONAL Q-02,Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L66).

Inputs candidates whose nonservice prerequisites match, service source snapshot/context; need not load for rules without service condition. AST When(AllOf(service-required,other-prerequisites),When(StateIs(serviceContext,UNAVAILABLE),technical EMIT,When(EQ(Count(relevant source rows),0),no-context EMIT,ordinary match))). Distinguish source unavailable from successfully read no relevant row; neither proves rule false due service content. Action legacy WARN category service-context at source/destination target; target CE completeness for required source, message data selectable. Admin dependency requiredness/prerequisites/target/text/channel. Witness candidate IP+port passes, registry unreadable → service warning+CE, not PASS/no matches.

### 36 — NET_EXCEPTIONS_STRUCTURED_CONFLICT

LegacyPath "EXCEPTIONS.structured-conflict"; parent NET_IP_EXCEPTIONS; **CONDITIONAL Q-02,Q-05**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L67).

Inputs fully matched technical candidates + host ranks + SERVICES/GROUPS; no extra source beyond unified rules. AST OP_RANK_MAX matched host scores/length → OP_UNIQUE_VALUE(top,compareFields=[SERVICES,GROUPS]); EMIT StateIs(result,CONFLICT); When OK publish structured; independently collect all comments/approvals. Action ERROR conflict at matched IP, retain instructive outputs, status CE conflict completeness unless business conflict treated FAIL explicitly approved. Admin compareFields/priority/specificity/tie/output text/channels. Witness equal host specificity with groups G1/G2 → no arbitrary group; both comments retained.

### 37 — NET_EXCEPTIONS_ALLROWS

LegacyPath "EXCEPTIONS.ALLROWS"; parent NET_IP_EXCEPTIONS; **CONDITIONAL Q-03,Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L68).

Inputs matched IP anchor, technical rows, rule scope/IP/port/protocol/ALL_ROWS; host/service excluded supported profile. AST relevant=Filter(rows,S,OP_IP_OVERLAPS(anchor,scope-selected S IPs)); AllOf(GT(Count(relevant),0),All(relevant,S,R(EXPR_RULE_MATCH,S,R))). For ALL_ROWS flag positive exemption action requires this predicate; default All([])=true alone insufficient. No output deletion/count mutation. Admin flag/scope/relevant-key/conditions/supported profile/failure channels. Witness same IP ports80+443 vs rule80 blocks positive output for this IP; unrelated IP443 outside scope does not block. Historical error fallback base retained is unresolved Q10, target CE not apparent approval.

### 38 — NET_HOSTNAME_BROAD_OUTPUT

LegacyPath "HOSTNAME.broad-output"; parent NET_HOSTNAME_RULES; **CONDITIONAL Q-05**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L69).

Inputs all matched host rows and top structured winner; SRC_UNIFIED_RULES. AST top structural selection separate expression from Map(allMatches,COMMENT/APPROVERS)→OP_DISTINCT; actions joined by host provenance, not Filter top across whole output. Action NEEDS_APPROVAL/comment ReportRows plus winner Service/Group; severity max all visible. Admin result-column split, templates/channels/specificity/priority/dedup. Witness exact service SAP and broad contains approval SECURITY → both survive; top selection never deletes broad approval. Broad generic literal example synthetic, not real operational rule.

### 39 — NET_HOSTNAME_CONFLICT

LegacyPath "HOSTNAME.conflict"; parent NET_HOSTNAME_RULES; **CONDITIONAL Q-02,Q-05**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L70).

Inputs hostname-only candidates, ranks, structural fields; gate best score exists. AST OP_RANK_MAX→OP_UNIQUE_VALUE(compareFields) and StateIs(CONFLICT); separate EMIT conflict, When OK structured output; other comments/approvals still collected. Actions conflict hostname target; no fabricated service, no arbitrary row order. Admin specificity/tie compareFields/numeric priority/affected channels/text. Witness identical exact pattern outputs S1/S2 → conflict; duplicate identical outputs → one unique structure no conflict. User routing/actual Rulebook fixture open; CONFLICT cannot be hidden PASS.

### 40 — NET_IPPOLICY_CONFIG

LegacyPath "IPPOLICY.config"; parent NET_INITIATOR_IP_ACCESS; **CONDITIONAL Q-02,Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L71).

Inputs SourceRows SRC_IP_POLICY, typed schema/parser diagnostics, parent primary eligibility; sources same pinned RB. Gate source actually needed, then StateIs(INVALID/UNAVAILABLE/ERROR/CONFLICT/LIMIT) explicit branches. INVALID active selector/scope/allow-deny row not silently filtered away; OK [] no-policy separate. Action config technical EmitIssue document-level, not email blame; CE blocks PASS. Admin source schema/mapping/validation profile, diagnostics text/channel; required source error semantics not user-selectable PASS. Witness malformed active policy → document technical error with source RowId, no email annotation; primary denial independently reports email issue.

### 41 — NET_IPPOLICY_LOWER_ONLY

LegacyPath "IPPOLICY.lower-only"; parent NET_INITIATOR_IP_ACCESS; **CONDITIONAL Q-03,Q-07,Q-10**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L72).

Inputs lower parsed valid intervals absent upper keys, best selected ANY/BOTH policy; SRC_INITIATORS/IP_POLICY. Gate primary allowed+valid cell; invalid/ambiguous lower records produce existing format diagnostics and do not enter guessed scope. AST Filter(selected policy,SCOPE ANY/BOTH)→allow/deny unions; per lower-only interval ASSERT no deny overlap and allowed full coverage if nonempty. Dedupe by physical provenance once before actions. Action FAIL ADDRESS_TEXT with lower-specific message; Admin scope/key/selector/grammar/text/target. Witness upper GROUP_A, lower .20, policy ANY allows .1–.10 → lower FAIL; SOURCE-only policy does not guess lower's side. Full intervals ADR007, exact absent-key policy migration Q07 conditional.

### 42 — NET_DB_SCHEMA

LegacyPath "DB.schema"; parent NET_INITIATOR_DB_ACCOUNT; **GAP_ACTION Q-02,Q-03,Q-10,source maintenance**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L73).

Inputs source schema capabilities (DB_FORBIDDEN column missing, owned/read-only/opened-by-user flags); not part of pure Request rule evaluation. AST validation StateIs(SourceRows schema,MISSING/INVALID) produces CE/diagnostic, optional ReportRow required schema change; **no Call creates column/dropdown or Save**. Proposed Admin SourceDefinition migration preview/version transition, outside Evaluator and requiring separate write-capability/ownership contract; current actions do not claim full source-maintenance implementation. Witness legacy could create column and save own-opened Rulebook; target design only signals migration need until that adapter exists. Admin bindings/validation/dropdown metadata editable; historical runtime side effect not silently lost or accepted. Honest architectural capability gap, not STRUCTURAL «supported».

### 43 — CONFIG_UNKNOWN_RULE

LegacyPath "UNKNOWN_CHECK"; no implemented LegacyRuleId (dynamic unknown text retained diagnostic); **STRUCTURAL**. [PR00 path](https://github.com/kosha1308/AccessTools-v2/blob/2a04ae7c2cd92c20b470d5a4aa0c39f8320188b3/docs/v1-parity-matrix.md#L74).

Inputs immutable package RuleRefs/ExpressionRefs/OperationRefs; no Excel/source needed. This is package validation, not eighteenth NET business rule. Schema reference check: None(registered rule refs,R,EQ(C(R,ID),requestedId)) → missing-reference diagnostic; no free dynamic macro dispatch. Actions publish blocked; runtime corrupt package → CE diagnostic, not PASS. Admin removes/repairs unresolved reference; cannot substitute arbitrary code. Witness unknown NET.UNKNOWN in draft detected before publish; legacy generated SKIPPED «Неизвестный ID…», target makes failure visible without inventing rule.

## Второй независимый шаблон — EQUIPMENT_RETURN

Синтетический проект «Акт возврата оборудования», не готовая конфигурация и не новое нормативное требование к эксплуатации. Нужен для проверки независимости модели до NETWORK_ACCESS. Ни один OperationId этого раздела не относится к адресам/сетям; template-specific evaluator и hardcoded asset adapter не предполагаются. **STRUCTURAL** означает только выразимость на бумаге.

TemplateId EQUIPMENT_RETURN@1.0.0, поля ER_EMPLOYEE_EMAIL:Text, ER_RETURN_DATE:Date. Таблица ER_ITEMS: ER_ROW_ID, ER_ASSET_ID:Text, ER_QUANTITY:Int, ER_CONDITION:Text, resolver provenance. Locator может быть DefinedName или TableColumn; изменение координат не меняет RuleId. SourceDefinition ASSET_REGISTRY общий, Binding ER_ASSET_BINDING@1.0.0: ASSET_ID:Text, OWNER_EMAIL:Text, EXPECTED_QUANTITY:Int, DUE_DATE:Date, ACTIVE:Bool. Синтетический source snapshot SNAP_ASSETS_001, request SNAP_RETURN_001; clock/locale pinned. Отсутствие source/schema → CE, не «актив не найден».

Synthetic source examples:
- EQ_001, owner u@example.test, expected1, due2026-10-20, active true.
- EQ_002, same owner, expected2, due2026-10-18, active true.

Synthetic request: ER_EMPLOYEE_EMAIL=u@example.test, ER_RETURN_DATE=2026-10-19; rowA EQ_001 quantity1 condition GOOD; rowB EQ_002 quantity2 condition DAMAGED. Не публикуются реальные сотрудники/активы.

### ER_REQUIRED — STRUCTURAL

Inputs employee/date и ER_ITEMS; no external source. Gate resolver schema OK. EMIT StateIs(F(ER_EMPLOYEE_EMAIL),EMPTY), StateIs(F(ER_RETURN_DATE),EMPTY), Count(ER_ITEMS)=0; per entity EMIT EMPTY asset id/quantity/condition. MISSING locator/INVALID typed Int/Date — отдельная technical/value диагностика, не false success. ASSERT GT(C(I,ER_QUANTITY),L(Int,0)) только при OK Int.

Actions EmitIssue+Annotate нужные logical fields/cells, FAIL обязательности; empty table target FIRST_ENTRY_SLOT. Admin mandatory field/column list, quantity relation/threshold, texts/channels/targets/priority. Witness quantity0 → FAIL; empty item list → table FAIL, а не All([])=true PASS. Reuses StateIs/Count/ForEach/OP_COMPARE и общие actions required сетевых rules без сетевого типа.

### ER_EMAIL_FORMAT — STRUCTURAL

Input ER_EMPLOYEE_EMAIL; gate nonempty; Let(E,Call(OP_EMAIL_PARSE@1.0.0,F(ER_EMPLOYEE_EMAIL),L(PROFILE,PROFILE_EMAIL_V1)),ASSERT(StateIs(V(E),OK))). INVALID format → FAIL employee email; parse ERROR→CE. Same operation as NET_INITIATOR_EMAIL, другой binding и message. Admin grammar profile/target/channel; witness u@example.test accepted, broken@ fails. EMPTY required handled separately, не две одинаковые ошибки.

### ER_ASSET_LOOKUP — STRUCTURAL

Inputs nonempty asset IDs, ER_ASSET_BINDING snapshot. AST Join(ER_ITEMS,I,Filter(SourceRows(ER_ASSET_BINDING@1.0.0,SNAP_ASSETS_001),A,EQ(C(A,ACTIVE),true)),A,EQ(N(C(I,ER_ASSET_ID)),N(C(A,ASSET_ID))),LEFT). For each item Filter matched source rows; Count0→FAIL unknown asset; Count>1→source-key CONFLICT/CE; Count1 supplies single typed record. ASSERT EQ(N(employeeEmail),N(source OWNER_EMAIL)) если owner-policy включён отдельным clause; это явное бизнес-условие, не lookup side effect.

Actions unknown/mismatched-owner на ER_ASSET_ID/ER_EMPLOYEE_EMAIL по Admin target, technical duplicate registry key document diagnostic; VIEW optional registry matched IDs. Admin binding/key normalization/active predicate/owner clause/messages/channels. Witness absent EQ_404 → FAIL только если source successfully OK; missing registry → CE. Reuses Filter/Join/Count/OP_TEXT_NORMALIZE/OP_COMPARE/OP_UNIQUE_VALUE; сетевой OP_INTERVAL_LOOKUP здесь не нужен.

### ER_ITEM_DUPLICATES — STRUCTURAL

Input ER_ITEMS asset IDs + RowId/provenance; no source. AST Map(normalized asset id)→GroupBy(ER_ASSET_ID_KEY)→Filter Count>1; OP_SORT by snapshot row order; EMIT subsequent member. Quantity не включается в duplicate key по явному синтетическому решению, а не общему закону.

Actions WARN либо FAIL выбранного policy на second ER_ASSET_ID, UiGroup первого logical row; отдельное count/quantity reconciliation, если понадобится, новая clause. Admin key columns, allowed duplicate policy, order/target/severity/message. Witness два EQ_001 quantity1 и2 → один duplicate issue на второй строке; координаты переставлены — provenance всё равно корректен. Reuses GroupBy/OP_SORT/provenance actions NET_TECH_DUPLICATE без source/destination/IP.

### ER_DUE_DATE — STRUCTURAL

Input OK ER_RETURN_DATE Date и unique asset registry DUE_DATE; gate valid nonempty date+single successful lookup. ASSERT Call(OP_DATE_COMPARE@1.0.0,F(ER_RETURN_DATE),C(A,DUE_DATE),L(RELATION,LE)); false WARN overdue. Clock comparison отдельно, если бизнес выберет «просрочен на сегодня»: OP_DATE_COMPARE(ContextRef(CLOCK_DATE),DUE_DATE,GT), не live Now. В этом примере сравнивается фактическая дата акта, clock не подменяет её.

Actions WARN ER_RETURN_DATE с item provenance и due date параметром, ReportRow overdue true. Admin reference field/relation/optional grace policy как typed date-profile parameter только при принятом контракте, texts/channels/target. Witness EQ_001 due20 vs returned19 PASS; EQ_002 due18 vs returned19 WARN. Missing/INVALID due date не «не просрочено», CE. Reuses DateCompare NET_VALIDITY без срока доступа/отзыва.

### ER_DAMAGE_APPROVAL — STRUCTURAL

Input ER_CONDITION; gate included item and OK condition. ASSERT Call(OP_LIST_CONTAINS@1.0.0,L(List<Text>,[good,damaged,incomplete]),N(C(I,ER_CONDITION))) после versioned canonical normalization (Trim + invariant lowercase). Separate EMIT EQ(N(condition),L(Text,damaged)) → NEEDS_APPROVAL с approval literal FINANCE_REVIEW@version; severity/channels independent. Нормализованы и ввод, и допустимые literals; uppercase display values не сравниваются с lowercase без канонизации.

Actions annotation на ER_CONDITION, COORDINATOR ReportRow approval; USER_REPORT independently configurable, no fake input error «сотрудник недопустим». Admin condition enum, approval recipients/templates/channel flags/severity/numeric priority. Witness rowB DAMAGED → NEEDS_APPROVAL даже при correct asset/quantity; no refusal invented. Reuses AllOf/When/EmitIssue/messages/channels из NET_INITIATOR_APPROVAL_RULES и collect-all approval projection.

### ER_RETURN_REPORT — STRUCTURAL

Projection, не дополнительная mandatory проверка. Inputs Join item/request/unique asset, named expressions for overdue/approvals/findings; source snapshot refs explicit. AST Map→Record schema {EMPLOYEE,RETURN_DATE,ASSET_ID,QUANTITY,CONDITION,DUE_DATE,OVERDUE,APPROVALS}; GroupBy employee if output profile needs aggregate; OP_DISTINCT approvals + OP_TEXT_JOIN; OP_PROJECT maps result schema, OP_SORT sets stable columns/rows. Не читать editable Excel results as truth.

Action ReportRow в копии output workbook с declared report columns; optional SourceView shows matched registry rows, no writes original asset registry. Admin sheet/display/column order/mapping/group/sort/templates/role/channel/ownership; Writer same как NETWORK_ACCESS. Witness two sample rows → two report rows, only EQ_002 overdue+approval, employee text repeated/projection configurable. Annotations/report source fail cannot leave overall PASS «всё верно», ActionReceipt separate from RuleOutcome.

### Что именно доказано вторым шаблоном

Те же Field/Column/Source refs, snapshots, StateIs, scalar compare, EmailParse, DateCompare, Filter/Map/GroupBy/Join, Record/projection, ranked selection/unique-value (если registry неоднозначен), EmitIssue/Annotate/ReportRow/SourceView выражают независимые required/lookup/duplicates/date/approval/report. Источник выбирается binding, не именем NETWORK_ACCESS. Schema позволяет rebind ASSET_REGISTRY в другой шаблон, а NET messages/targets не переиспользуются случайно. Это **не** доказательство, что UI уже умеет создать такой шаблон: AC-01…AC-10 требуют будущего реального Admin и Excel.

## Открытые границы и критерий завершения матрицы

1. **Q-01/Q-02/Q-08:** integration — заданный рабочий audit baseline, но конкретный versioned production Rulebook/полнота повседневных случаев ещё не подтверждены. Синтетические источники выше не закрывают этот пробел.
2. **Q-03:** no-policy/primary eligibility, scope inheritance и DB association допускают явное AST описание наблюдения; публикация конкретного варианта ждёт решения. No-match OK [] не одинаков unavailable source.
3. **Q-04:** shared resolver table/merge/provenance и Writer ownership требуют отдельного принятого контракта и Excel witnesses. Бумажный GroupBy не доказывает правильный COM-read/write.
4. **Q-05/Q-06/Q-09:** target recipient/channel, hidden UI issues, global coordinator output, report schema/hostname ambiguity/partial coverage severity — DecisionRequirement в package. Не выводить «готово» из наличия EmitIssue.
5. **Q-07:** диапазон inclusive уже принят ADR007; declaration equality, IP vs /32, compact grammar и group fallback остаются conditional. AST и OP profiles не утверждают default.
6. **Q-10:** read failure/transform/writer exceptions исторически могли дать PASS/fallback. Это записано как migration release blocker: required non-OK CE не выключается цветом/visibility. Legacy fallback воспроизводить как скрытый успех нельзя.
7. **Q-11:** портовой схемы registry нет; отдельно Rulebook PortSet coverage достаточно для доказанного поведения. Новый port lookup только после явной source/schema/business-policy спецификации.
8. **NET_DB_SCHEMA:** текущие ActionKind не обещают migration/edit/save исходного Rulebook. Admin source-schema maintenance capability — отдельный gap; не заменять исходный эффект молчаливой read-only «поддержкой».
9. **Operation contracts:** OP version/signature, parser profiles, exact locale behavior, limits, ContextRef snapshot, expression alias/type checking и derived Record construction должны быть проверены общим schema validator. Профиль с именем V1 — ссылка на ещё проектируемое определение, не принятая грамматика.
10. **Выход:** типизированный факт и ActionIntent ≠ выполненное действие. ActionReceipt/OwnershipJournal и manual double-run/cleanup/rollback tests ещё отсутствуют.

Матрица учитывает **43 audit entities** и синтетический EQUIPMENT_RETURN. Для 42 entities дана композиция в модели (с условными бизнес-профилями там, где остаются Q); для DB.schema зафиксирован отдельный GAP_ACTION точного source write-back. Для каждого указан formulation, witness, sources/dependencies, selectable Admin business data и unresolved boundary. Доказательство runtime expressiveness появится только после компиляции этих typed sketches в общий interpreter, создания/редактирования через Admin и отдельной приёмки Excel; этот PR таких утверждений не делает.
