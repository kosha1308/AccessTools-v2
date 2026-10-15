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

Initial registry includes scalar/domain type metadata, port/IP set-operation
signatures and LOWER(TEXT). It is a tested foundation, not the complete language
catalog from appendix N. Functions/operators are added with their consuming parser
and compiler tests; no claim of complete DSL support is made at this point.
