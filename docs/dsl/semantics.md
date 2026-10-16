# Evaluation contract reserved for the next implementation stage

This document defines the contract. A partial stage-2A evaluator is drafted but
has not passed Excel compilation or execution tests; see ../runtime-core-status.md.
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
data as a valid filled business value. NULL has its own reserved ValueState=NULL;
it is not EMPTY, MISSING or INVALID. Ordinary scalar operators require a usable
PRESENT value. Explicit state inspectors may inspect EMPTY/NULL or a domain-invalid
normalizer result; they never suppress a source/schema ERROR, SKIPPED or MANUAL_REVIEW.

| Inspector | PRESENT | EMPTY | NULL | domain INVALID | MISSING |
| --- | --- | --- | --- | --- | --- |
| EMPTY | FALSE | TRUE | FALSE | ERROR | ERROR |
| NOT_EMPTY | TRUE | FALSE | FALSE | ERROR | ERROR |
| IS NULL | FALSE | FALSE | TRUE | ERROR | ERROR |
| IS NOT NULL | TRUE | TRUE | FALSE | ERROR | ERROR |
| VALID | TRUE | FALSE | FALSE | FALSE | ERROR |
| INVALID | FALSE | FALSE | FALSE | TRUE | ERROR |

These are explicitly state-aware primitives, not implicit error-to-Boolean casts.
COALESCE has two same-typed arguments: only explicit NULL selects its fallback.
EMPTY remains empty; non-OK results propagate. The fallback is evaluated lazily.

## Scalar and source evaluator contracts

Runtime NUMBER values are finite IEEE-754 Double values. Decimal literal spelling
is preserved exactly by the compiler; conversion happens in the typed normalizer.
Overflow, nonfinite values and nonzero underflow to zero are diagnostic errors.
Arithmetic errors/division by zero propagate as ERROR. ROUND uses ties-to-even and
an integral precision from 0 through 28. Runtime checks value-dependent function
constraints; the compiler does not evaluate arithmetic/function calls.

DATE is a Gregorian date in years 0100–9999; DATE_DIFF(start,end) returns signed
whole days end minus start. TODAY uses the fixed run timestamp and configured
timezone. Locale/timezone/normalization settings belong to EvaluationContext and
its cache identity; absence is CONFIG_CONTEXT_REQUIRED, never a guessed system
setting. DATETIME is a field/source-only normalized instant with explicit timezone
policy; no localized or offset-free datetime literal is guessed by this profile.

TEXT comparisons use the field's published NormalizeMode. Legacy compatibility
uses Trim plus case-insensitive matching; literal operands use the same comparison
policy. Case conversion uses the configured normalization locale. LEN counts UTF-16
units of its input. TRIM and comparison normalization never tokenize free text.
LIKE uses the registered wildcard policy; MATCHES requires a safe bounded evaluator,
not arbitrary host regex/EVAL. No regex engine executes during this compiler stage.

MIN/MAX on empty lists are AGGREGATE_EMPTY errors; SUM(empty)=0, AVG(empty) is an
AGGREGATE_EMPTY error, DISTINCT_COUNT(empty)=0. Mixed-element lists are type errors.
MIN/MAX support two compatible scalar arguments or one declared typed list.

SOURCE/LOOKUP compiler bindings contain only schemas/dependencies. The runtime
adapter must have explicit policies for NOT_FOUND, AMBIGUOUS, STALE, UNAVAILABLE and
INVALID_SOURCE; missing policies are CONFIG_SOURCE_POLICY_REQUIRED. Not-found or
ambiguous rows never become business PASS. Multiple rows are not reduced to the
first one. Source data is loaded only through the active dependency graph. Returned
values are normalized/validated against the declared column type before use.

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
guarantees for arbitrary real workbooks. Regex signatures are metadata only;
an executable safe-regex handler and its work/time budgets belong to stage 2.

## Explicit unsupported capabilities

Runtime source data adapters, verified normalization/evaluation, legacy ALLROWS
import, production extraction, explain execution, Admin UI, publishing and release
profiles are not implemented here. LOOKUP/SOURCE_VALUE compiler schema binding
and operator/function signatures are implemented. Unregistered syntax/functions/
types fail before execution. A scalar
IP field passed to DISTINCT_COUNT is rejected; its template must provide a typed
collection binding. No automatic column-to-collection projection is guessed.
