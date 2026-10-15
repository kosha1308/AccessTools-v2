# AccessTools compiler contracts — DSL/Schema 1.0.0

This is the supported compiler profile. Runtime evaluator/source adapters, UI and
publication are separate stages. A successful Compile result is a typed expression,
not a business PASS and not permission to change a workbook.

## Identity, syntax and versions

Entity IDs use ASCII [A-Z][A-Z0-9_]* and canonical uppercase. Lowercase input
canonicalizes; localized/punctuated IDs are rejected. Stable FieldIds do not depend
on Excel headings. Schema/DSL versions are retained on ATCompileResult. Versioned
registry entries are the source of signatures, syntax, help, examples and bindings.

Text -> lexer -> shared AST -> identifier/type/dependency checking -> typed AST.
Contextual syntax and explicit field syntax share this pipeline. Parentheses,
NOT/AND/OR and arithmetic have the precedence recorded in grammar.ebnf. Relations
include binary operators, postfix state inspectors and three-operand BETWEEN.
NOT IN lowers to NOT(IN(...)); aliases retain a canonical primitive identity.

A colon does not blindly make a list: a scalar signature accepts one scalar value;
a collection signature may lower one contextual value to a singleton LIST. This is
contextual syntactic sugar, not a global scalar-to-list implicit conversion.
Canonical list syntax is [1,2]; function arguments use brackets for one list argument.
No user text is passed to EVAL, formulas, shell or dynamic macro dispatch.

## Registry and literal contracts

Entries have stable IDs/categories, readable descriptions, signatures, syntax,
precedence/associativity, evaluator binding IDs, version/status, examples and test
contracts. Registered primitives/signatures freeze. Category/syntax conflicts and
unknown operand/result types are rejected. LIST/SET/ROW/COLLECTION types use bounded
structural type expressions. The generic type variable is T.

DataType entries include NormalizerId, LiteralPolicy and ValidationPolicy.
TEXT preserves one free-text value; NUMBER uses invariant decimal spelling; BOOL
uses explicit Boolean/mapping policy; DATE uses Gregorian ISO years 0100–9999;
PORT/PORT_SET use unsigned 16-bit integer validation. Domain types without a literal
constructor are FIELD_ONLY. A string is not guessed to be an IP, email or hostname.

Canonical numeric text strips insignificant leading/trailing zeroes without
floating-point conversion. Runtime numeric conversion/range handling follows the
evaluation contract in semantics.md. DATE literals are validated without locale or
host calendar dependence. Registered NUMBER/LIST<NUMBER> -> PORT_SET conversions
validate known numeric literals (including named arguments/defaults) at compile
time. Values coming from fields/functions still require the runtime target
normalizer; compile-time type compatibility is not a substitute for value checks.
Port 0 is an unsigned numeric port value; service/business restrictions are separate.

The current operator/function catalog is generated in catalog.md. All non-DataType
registry examples compile against a defined fixture context. Tests reject drift
between the registry and the committed generated catalog. Regenerate explicitly:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/Test.ps1 -UpdateCatalog

## Compiled model

AST nodes retain DataType, canonical PrimitiveId and the selected immutable
signature when applicable. Registered implicit conversions record the target
normalizer binding. Named calls retain their typed Expansion and ParameterBindings,
including parsed defaults and declared target types. Runtime will not need to
reparse defaults or original DSL to recover these bindings. Canonical serialization
uses the expression syntax, not the internal expansion.

Rule root type defaults to BOOL. Unknown names, mismatches, invalid known literals
and limits produce diagnostics and no usable root. Warnings may accompany a valid
typed root. No source workbook is opened by compilation.

## Scope, named expressions and dependencies

ATCompileContext binds document fields, table fields, shared source schemas and
named expressions. All metadata types are validated first. Row collections use
COLLECTION<ROW<TABLE_ID>>. Quantifier aliases are lexical, mandatory and cannot
shadow an active alias. An unqualified field requires exactly one candidate;
document and row candidates are included, and ambiguity is never resolved silently.

Named Scope=DOCUMENT is the default and does not capture caller row aliases.
Scope=ROW explicitly permits the active row context. Supplied arguments resolve in
the caller. Defaults resolve in definition scope and may reference only earlier
declared parameters. Direct, indirect and default recursion are rejected.
Deprecated expressions remain compatible with a warning. Dependencies and UsedBy
edges include named expressions and the fields/sources they expand to.

SOURCE binds a registered schema. LOOKUP/SOURCE_VALUE require literal registered
SourceId and ColumnId, an explicitly declared lookup-key type, and a compatible key
expression. Their result type comes from the source column schema. Source/column
identities canonicalize; SOURCE_VALUE serializes as LOOKUP. Data access and handling
of not-found/ambiguous/stale sources are reserved for the source adapter stage.

## Diagnostics

ATDiagnostic carries stable Code/Category/Severity, SourceId, Start, Length and a
human message. Spans are 1-based UTF-16 offsets in the original expression. EOF may
have Length=0. SourceId identifies RULE, SCHEMA, NAMED:<id> or DEFAULT:<id>:<parameter>.
Nested parser errors retain their actual code/source. Unexpected native errors map
to stable internal codes while preserving details; localized exception text is not
used as a diagnostic ID.

Known follow-up for evaluator/explain: some synthetic/schema literal AST nodes do
not yet carry SourceId, and nested deprecation warnings use RULE provenance.
These do not affect compilation results; they are listed in stage1-status.md.
