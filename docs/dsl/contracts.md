# Compiler contracts (DSL/Schema 1.0.0)

Entity IDs are ASCII identifiers matching [A-Z][A-Z0-9_]*. Lowercase ASCII input
canonicalizes to uppercase; whitespace, punctuation and localized letters are
rejected. Registry keys are unique across categories. Types register before
operators/functions; every signature references registered DataTypes.

ATPrimitive contains a stable ID, category, description, evaluator binding ID,
semantic version, active/deprecated status, determinism policy, examples, test
contract and one or more ATSignature values. EvaluatorId is metadata only: it is
never passed to Application.Run. Registry entries and signatures freeze after
registration; returned collections are copies. Runtime evaluators are not yet implemented.

ATDiagnostic carries code/category/severity, Start and Length, and a human message.
Spans count UTF-16 code units (VBA String indexing), starting at 1. End-of-input
diagnostics may use Length=0. Invalid spans and incomplete diagnostic metadata fail
explicitly. Positions refer to original input, not canonical output.

EvalState (OK, ERROR, SKIPPED, MANUAL_REVIEW) and ValueState (PRESENT, EMPTY,
MISSING, INVALID) are separate axes. Only EvalState OK may carry a usable Boolean.
EMPTY means a resolved field with no user value, MISSING means failed resolution,
INVALID means present data incompatible with its declared type. None may be
silently coerced into False, zero or an empty string. Evaluator implementation and
propagation tests belong after the compiler acceptance gate.

The registry includes scalar/domain types, port/IP set-operation signatures,
arithmetic/comparison/Boolean operators, LOWER, DATE, COUNT, DISTINCT_COUNT, EXISTS,
SOURCE and collection quantifiers/FILTER. Parser relation recognition uses the
registered syntax (including multiword syntax), not primitive IDs. Conflicting
syntax in one category is rejected. DATE's function primitive ID is DATE_LITERAL,
with Syntax=DATE, keeping its identity distinct from the DATE DataType.

ATCompileContext binds document/table fields, shared source metadata and named
expressions. Every metadata type is validated before compilation. No source file
is opened. A row collection uses COLLECTION<ROW<TABLE_ID>>; a row alias is mandatory
for a quantifier and may not shadow an active alias. Unqualified row references
resolve only when there is exactly one candidate, including document candidates.

Compile(text, context, contextualFieldId, requiredType, limits) returns a typed AST,
diagnostics and deduplicated dependencies/UsedBy edges. Rule root type defaults to
BOOL. Errors return no usable root. Deprecation warnings preserve the typed root.
Named parameters are checked against declared types; explicit arguments are checked
in the caller context, defaults/body under a cycle guard. The default/body expansion
is type-checked and retained on the named AST node. Direct/indirect/default recursion
and missing required arguments fail compilation.

Registered conversions currently allow NUMBER or LIST<NUMBER> to PORT_SET.
These are compile-time compatibility declarations, not implemented port parsers.
Runtime must apply the target domain validator before producing an OK value.
Compiler front-end success is therefore not authorization to publish/execute a rule.
The full primitive/evaluator and domain-normalizer catalog remains a subsequent gate.
