# Evaluation contract reserved for the next implementation stage

This document defines the contract; no evaluator executes these operations yet.
Compiler acceptance and Runtime readiness are separate gates.

## Results and failure propagation

EvalResult<T> carries Value, DataType, EvalState and diagnostics. Value is usable
only when EvalState=OK. Each check run fixes its run identity and timestamp in one
EvaluationContext; contextual functions receive that snapshot. Registry metadata
defines determinism and cache policy. No primitive performs arbitrary I/O or actions.

Evaluation is left to right. For evaluated operands, the first non-OK result
(ERROR, SKIPPED or MANUAL_REVIEW) stops that expression and propagates unchanged;
it is never coerced into a Boolean. Primitive exceptions become ERROR diagnostics.
An explicit source policy may map source failures into SKIPPED/MANUAL_REVIEW before
propagation. This preserves a deterministic diagnostic and prevents silent PASS.

| Left operand | AND | OR |
| --- | --- | --- |
| OK/FALSE | OK/FALSE; right NOT_EVALUATED | evaluate right |
| OK/TRUE | evaluate right | OK/TRUE; right NOT_EVALUATED |
| non-OK | propagate left; right NOT_EVALUATED | propagate left; right NOT_EVALUATED |

NOT flips only an OK Boolean. NOT_EVALUATED is an explain-trace marker, not a
failure state. Actions may consume only an OK/TRUE rule result and are executed
by a separate action layer. False or non-OK results do not silently fire actions.

EMPTY is an existing resolved empty value, MISSING is failed schema/template
resolution, INVALID is present data rejected by its type normalizer. EMPTY checks
inspect an explicitly empty resolved value. MISSING follows template/config error
policy; INVALID follows domain-validation policy. NOT_EMPTY must not treat invalid
data as a valid filled business value. A future nullable/source-null primitive must
declare its own signature; null is not silently equated with these three states.

## Sets and collections

A denotes requested/actual values; R denotes allowed/expected values.
ANY/OVERLAPS: intersection nonempty. NONE: intersection empty. COVER: A subset R.
CONTAINS: R subset A. EXACT: normalized sets equal. Normalization must preserve
the established legacy distinction between declaration form and interval coverage.
IP, CIDR, ranges and whole-octet x masks require domain normalizers; string heuristics
and partial overlap must not replace full coverage. These algorithms are not newly
implemented or claimed tested by compiler-only assertions.

ALLROWS(empty)=TRUE, ANYROW(empty)=FALSE, NOROWS(empty)=TRUE. FILTER(empty) returns
an empty collection and COUNT(empty)=0. Rules requiring data must add EXISTS.
Each evaluated predicate follows the same non-OK propagation contract; a decisive
Boolean may short circuit and remaining rows are marked NOT_EVALUATED in explain.

## Current limits

Default input length: 65536 UTF-16 code units. Tokens: 16384 including EOF.
AST nodes: 8192. Syntactic/AST/expanded traversal nesting: 64. Nested collection
operators: 8. Named expansion depth: 64. Total compiler traversal steps: 200000.
Metadata type nesting: 16. Engineering limits are configurable through ATLimits;
invalid/unsafe recursion settings fail with diagnostics. They are not performance
guarantees for arbitrary real workbooks. No regex evaluator is registered yet.

## Explicit unsupported capabilities

LOOKUP/SOURCE_VALUE adapters, runtime IP/port normalization/evaluation, the complete
text/date/function catalog, legacy ALLROWS import, production extraction, explain
execution, Admin UI, publishing and release profiles are not implemented in this
milestone. Unregistered syntax/functions/types fail before execution. A scalar
IP field passed to DISTINCT_COUNT is rejected; its template must provide a typed
collection binding. No automatic column-to-collection projection is guessed.
