# AccessTools DSL 1.0.0 language catalog

Generated from AT_CreateRegistry; edit registry metadata, then regenerate this file.

These are compiler signatures and versioned evaluator/normalizer binding contracts. Runtime evaluators are not enabled by this catalog.

## TEXT

- Category: DATATYPE
- Display name: TEXT
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: TEXT
- Canonical ID/syntax: TEXT / TEXT
- Negated alias: FALSE
- Signatures: () -&gt; TEXT
- Evaluator binding: TEXT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: TEXT_V1
- Literal policy: STRING
- Validation policy: PRESERVE_FREE_TEXT
- Example: TEXT
- Test contract: positive; negative; empty; boundary; malformed

## NUMBER

- Category: DATATYPE
- Display name: NUMBER
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: NUMBER
- Canonical ID/syntax: NUMBER / NUMBER
- Negated alias: FALSE
- Signatures: () -&gt; NUMBER
- Evaluator binding: NUMBER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: DECIMAL_TEXT_V1
- Literal policy: DECIMAL
- Validation policy: INVARIANT_DECIMAL
- Example: NUMBER
- Test contract: positive; negative; empty; boundary; malformed

## BOOL

- Category: DATATYPE
- Display name: BOOL
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: BOOL
- Canonical ID/syntax: BOOL / BOOL
- Negated alias: FALSE
- Signatures: () -&gt; BOOL
- Evaluator binding: BOOL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: BOOL_V1
- Literal policy: TRUE_FALSE
- Validation policy: EXPLICIT_MAPPING_ONLY
- Example: BOOL
- Test contract: positive; negative; empty; boundary; malformed

## DATE

- Category: DATATYPE
- Display name: DATE
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: DATE
- Canonical ID/syntax: DATE / DATE
- Negated alias: FALSE
- Signatures: () -&gt; DATE
- Evaluator binding: DATE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: ISO_DATE_V1
- Literal policy: DATE_CALL
- Validation policy: ISO_GREGORIAN_0100_9999
- Example: DATE
- Test contract: positive; negative; empty; boundary; malformed

## DATETIME

- Category: DATATYPE
- Display name: DATETIME
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: DATETIME
- Canonical ID/syntax: DATETIME / DATETIME
- Negated alias: FALSE
- Signatures: () -&gt; DATETIME
- Evaluator binding: DATETIME
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: DATETIME_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: DATETIME
- Test contract: positive; negative; empty; boundary; malformed

## EMAIL

- Category: DATATYPE
- Display name: EMAIL
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: EMAIL
- Canonical ID/syntax: EMAIL / EMAIL
- Negated alias: FALSE
- Signatures: () -&gt; EMAIL
- Evaluator binding: EMAIL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: EMAIL_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: EMAIL
- Test contract: positive; negative; empty; boundary; malformed

## PROTOCOL

- Category: DATATYPE
- Display name: PROTOCOL
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: PROTOCOL
- Canonical ID/syntax: PROTOCOL / PROTOCOL
- Negated alias: FALSE
- Signatures: () -&gt; PROTOCOL
- Evaluator binding: PROTOCOL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: PROTOCOL_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: PROTOCOL
- Test contract: positive; negative; empty; boundary; malformed

## HOSTNAME

- Category: DATATYPE
- Display name: HOSTNAME
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: HOSTNAME
- Canonical ID/syntax: HOSTNAME / HOSTNAME
- Negated alias: FALSE
- Signatures: () -&gt; HOSTNAME
- Evaluator binding: HOSTNAME
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: HOSTNAME_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: HOSTNAME
- Test contract: positive; negative; empty; boundary; malformed

## PORT

- Category: DATATYPE
- Display name: PORT
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: PORT
- Canonical ID/syntax: PORT / PORT
- Negated alias: FALSE
- Signatures: () -&gt; PORT
- Evaluator binding: PORT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: PORT_SET_V1
- Literal policy: REGISTERED_NUMERIC_CONVERSION
- Validation policy: UNSIGNED_16_BIT_INTEGER
- Example: PORT
- Test contract: positive; negative; empty; boundary; malformed

## PORT_SET

- Category: DATATYPE
- Display name: PORT_SET
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: PORT_SET
- Canonical ID/syntax: PORT_SET / PORT_SET
- Negated alias: FALSE
- Signatures: () -&gt; PORT_SET
- Evaluator binding: PORT_SET
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: PORT_SET_V1
- Literal policy: REGISTERED_NUMERIC_CONVERSION
- Validation policy: UNSIGNED_16_BIT_INTEGER
- Example: PORT_SET
- Test contract: positive; negative; empty; boundary; malformed

## IP

- Category: DATATYPE
- Display name: IP
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: IP
- Canonical ID/syntax: IP / IP
- Negated alias: FALSE
- Signatures: () -&gt; IP
- Evaluator binding: IP
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: IP_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: IP
- Test contract: positive; negative; empty; boundary; malformed

## IP_EXPRESSION

- Category: DATATYPE
- Display name: IP_EXPRESSION
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: IP_EXPRESSION
- Canonical ID/syntax: IP_EXPRESSION / IP_EXPRESSION
- Negated alias: FALSE
- Signatures: () -&gt; IP_EXPRESSION
- Evaluator binding: IP_EXPRESSION
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: IP_EXPRESSION_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: IP_EXPRESSION
- Test contract: positive; negative; empty; boundary; malformed

## ENUM

- Category: DATATYPE
- Display name: ENUM
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: ENUM
- Canonical ID/syntax: ENUM / ENUM
- Negated alias: FALSE
- Signatures: () -&gt; ENUM
- Evaluator binding: ENUM
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: ENUM_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: ENUM
- Test contract: positive; negative; empty; boundary; malformed

## FIELD_REF

- Category: DATATYPE
- Display name: FIELD_REF
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: FIELD_REF
- Canonical ID/syntax: FIELD_REF / FIELD_REF
- Negated alias: FALSE
- Signatures: () -&gt; FIELD_REF
- Evaluator binding: FIELD_REF
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: FIELD_REF_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: FIELD_REF
- Test contract: positive; negative; empty; boundary; malformed

## SOURCE_REF

- Category: DATATYPE
- Display name: SOURCE_REF
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: SOURCE_REF
- Canonical ID/syntax: SOURCE_REF / SOURCE_REF
- Negated alias: FALSE
- Signatures: () -&gt; SOURCE_REF
- Evaluator binding: SOURCE_REF
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Normalizer binding: SOURCE_REF_V1
- Literal policy: FIELD_ONLY
- Validation policy: INVALID_IS_ERROR
- Example: SOURCE_REF
- Test contract: positive; negative; empty; boundary; malformed

## ANY

- Category: OPERATOR
- Display name: ANY
- Description: True when normalized actual and expected sets have a nonempty intersection
- Syntax: ANY
- Canonical ID/syntax: ANY / ANY
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL
- Evaluator binding: ANY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS ANY: 443
- Test contract: positive; negative; empty; boundary; malformed

## NONE

- Category: OPERATOR
- Display name: NONE
- Description: True when normalized actual and expected sets have no intersection
- Syntax: NONE
- Canonical ID/syntax: NONE / NONE
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL
- Evaluator binding: NONE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS NONE: 443
- Test contract: positive; negative; empty; boundary; malformed

## COVER

- Category: OPERATOR
- Display name: COVER
- Description: Full coverage: every actual/requested element is inside the expected/allowed set
- Syntax: COVER
- Canonical ID/syntax: COVER / COVER
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL
- Evaluator binding: COVER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS COVER: 443
- Test contract: positive; negative; empty; boundary; malformed

## CONTAINS

- Category: OPERATOR
- Display name: CONTAINS
- Description: Contains expected text or all expected set elements according to the selected typed signature
- Syntax: CONTAINS
- Canonical ID/syntax: CONTAINS / CONTAINS
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: CONTAINS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS CONTAINS: 443
- Test contract: positive; negative; empty; boundary; malformed

## EXACT

- Category: OPERATOR
- Display name: EXACT
- Description: Equality of normalized sets; ordering and duplicates do not change set equality
- Syntax: EXACT
- Canonical ID/syntax: EXACT / EXACT
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL
- Evaluator binding: EXACT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS EXACT: 443
- Test contract: positive; negative; empty; boundary; malformed

## OVERLAPS

- Category: OPERATOR
- Display name: OVERLAPS
- Description: True when normalized actual and expected sets have a nonempty intersection
- Syntax: OVERLAPS
- Canonical ID/syntax: OVERLAPS / OVERLAPS
- Negated alias: FALSE
- Signatures: (PORT_SET, PORT_SET) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (LIST&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, SET&lt;T&gt;) -&gt; BOOL; (SET&lt;T&gt;, LIST&lt;T&gt;) -&gt; BOOL
- Evaluator binding: OVERLAPS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @PORTS OVERLAPS: 443
- Test contract: positive; negative; empty; boundary; malformed

## LOWER

- Category: FUNCTION
- Display name: LOWER
- Description: Text case conversion under the registered normalization policy
- Syntax: LOWER
- Canonical ID/syntax: LOWER / LOWER
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; TEXT
- Evaluator binding: LOWER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: LOWER(@NAME)
- Test contract: positive; negative; empty; boundary; malformed

## INTERSECTS

- Category: OPERATOR
- Display name: INTERSECTS
- Description: True when normalized actual and expected sets have a nonempty intersection
- Syntax: INTERSECTS
- Canonical ID/syntax: INTERSECTS / INTERSECTS
- Negated alias: FALSE
- Signatures: (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL
- Evaluator binding: INTERSECTS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @IP INTERSECTS @OTHER
- Test contract: positive; negative; empty; boundary; malformed

## CONTAINED_BY

- Category: OPERATOR
- Display name: CONTAINED_BY
- Description: Full coverage: every actual/requested element is inside the expected/allowed set
- Syntax: CONTAINED_BY
- Canonical ID/syntax: CONTAINED_BY / CONTAINED_BY
- Negated alias: FALSE
- Signatures: (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL
- Evaluator binding: CONTAINED_BY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @IP CONTAINED_BY @ALLOWED
- Test contract: positive; negative; empty; boundary; malformed

## IN

- Category: OPERATOR
- Display name: IN
- Description: Typed membership, or full domain containment for IP expressions
- Syntax: IN
- Canonical ID/syntax: IN / IN
- Negated alias: FALSE
- Signatures: (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL; (T, LIST&lt;T&gt;) -&gt; BOOL; (T, SET&lt;T&gt;) -&gt; BOOL
- Evaluator binding: IN
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @IP IN SOURCE("ALLOWED_NETWORKS")
- Test contract: positive; negative; empty; boundary; malformed

## OP_EQUALS

- Category: OPERATOR
- Display name: OP_EQUALS
- Description: Equality/inequality of compatible typed operands
- Syntax: =
- Canonical ID/syntax: OP_EQUALS / =
- Negated alias: FALSE
- Signatures: (T, T) -&gt; BOOL
- Evaluator binding: OP_EQUALS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: 1 = 2
- Test contract: positive; negative; empty; boundary; malformed

## OP_NOT_EQUALS

- Category: OPERATOR
- Display name: OP_NOT_EQUALS
- Description: Equality/inequality of compatible typed operands
- Syntax: !=
- Canonical ID/syntax: OP_NOT_EQUALS / !=
- Negated alias: FALSE
- Signatures: (T, T) -&gt; BOOL
- Evaluator binding: OP_NOT_EQUALS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: 1 != 2
- Test contract: positive; negative; empty; boundary; malformed

## AND

- Category: OPERATOR
- Display name: AND
- Description: Boolean conjunction with deterministic left-to-right short circuit
- Syntax: AND
- Canonical ID/syntax: AND / AND
- Negated alias: FALSE
- Signatures: (BOOL, BOOL) -&gt; BOOL
- Evaluator binding: AND
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 7 / LEFT
- Example: TRUE AND FALSE
- Test contract: positive; negative; empty; boundary; malformed

## OR

- Category: OPERATOR
- Display name: OR
- Description: Boolean disjunction with deterministic left-to-right short circuit
- Syntax: OR
- Canonical ID/syntax: OR / OR
- Negated alias: FALSE
- Signatures: (BOOL, BOOL) -&gt; BOOL
- Evaluator binding: OR
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 8 / LEFT
- Example: TRUE OR FALSE
- Test contract: positive; negative; empty; boundary; malformed

## NOT

- Category: OPERATOR
- Display name: NOT
- Description: Negates an OK Boolean; non-OK states are propagated without Boolean coercion
- Syntax: NOT
- Canonical ID/syntax: NOT / NOT
- Negated alias: FALSE
- Signatures: (BOOL) -&gt; BOOL
- Evaluator binding: NOT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 6 / RIGHT
- Example: NOT TRUE
- Test contract: positive; negative; empty; boundary; malformed

## ADD

- Category: OPERATOR
- Display name: ADD
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: +
- Canonical ID/syntax: ADD / +
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: ADD
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 4 / LEFT
- Example: 1 + 2
- Test contract: positive; negative; empty; boundary; malformed

## SUBTRACT

- Category: OPERATOR
- Display name: SUBTRACT
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: -
- Canonical ID/syntax: SUBTRACT / -
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: SUBTRACT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 4 / LEFT
- Example: 1 - 2
- Test contract: positive; negative; empty; boundary; malformed

## MULTIPLY

- Category: OPERATOR
- Display name: MULTIPLY
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: *
- Canonical ID/syntax: MULTIPLY / *
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: MULTIPLY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 3 / LEFT
- Example: 1 * 2
- Test contract: positive; negative; empty; boundary; malformed

## DIVIDE

- Category: OPERATOR
- Display name: DIVIDE
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: /
- Canonical ID/syntax: DIVIDE / /
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: DIVIDE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 3 / LEFT
- Example: 1 / 2
- Test contract: positive; negative; empty; boundary; malformed

## REMAINDER

- Category: OPERATOR
- Display name: REMAINDER
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: %
- Canonical ID/syntax: REMAINDER / %
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: REMAINDER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 3 / LEFT
- Example: 1 % 2
- Test contract: positive; negative; empty; boundary; malformed

## UNARY_PLUS

- Category: OPERATOR
- Display name: UNARY_PLUS
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: UNARY_PLUS
- Canonical ID/syntax: UNARY_PLUS / UNARY_PLUS
- Negated alias: FALSE
- Signatures: (NUMBER) -&gt; NUMBER
- Evaluator binding: UNARY_PLUS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 2 / RIGHT
- Example: +1
- Test contract: positive; negative; empty; boundary; malformed

## UNARY_MINUS

- Category: OPERATOR
- Display name: UNARY_MINUS
- Description: Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero
- Syntax: UNARY_MINUS
- Canonical ID/syntax: UNARY_MINUS / UNARY_MINUS
- Negated alias: FALSE
- Signatures: (NUMBER) -&gt; NUMBER
- Evaluator binding: UNARY_MINUS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 2 / RIGHT
- Example: -1
- Test contract: positive; negative; empty; boundary; malformed

## GREATER

- Category: OPERATOR
- Display name: GREATER
- Description: Ordered comparison of compatible numbers, dates or text
- Syntax: &gt;
- Canonical ID/syntax: GREATER / &gt;
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; BOOL; (DATE, DATE) -&gt; BOOL; (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: GREATER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @A &gt; @B
- Test contract: positive; negative; empty; boundary; malformed

## LESS

- Category: OPERATOR
- Display name: LESS
- Description: Ordered comparison of compatible numbers, dates or text
- Syntax: &lt;
- Canonical ID/syntax: LESS / &lt;
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; BOOL; (DATE, DATE) -&gt; BOOL; (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: LESS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @A &lt; @B
- Test contract: positive; negative; empty; boundary; malformed

## GREATER_EQUAL

- Category: OPERATOR
- Display name: GREATER_EQUAL
- Description: Ordered comparison of compatible numbers, dates or text
- Syntax: &gt;=
- Canonical ID/syntax: GREATER_EQUAL / &gt;=
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; BOOL; (DATE, DATE) -&gt; BOOL; (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: GREATER_EQUAL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @A &gt;= @B
- Test contract: positive; negative; empty; boundary; malformed

## LESS_EQUAL

- Category: OPERATOR
- Display name: LESS_EQUAL
- Description: Ordered comparison of compatible numbers, dates or text
- Syntax: &lt;=
- Canonical ID/syntax: LESS_EQUAL / &lt;=
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; BOOL; (DATE, DATE) -&gt; BOOL; (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: LESS_EQUAL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @A &lt;= @B
- Test contract: positive; negative; empty; boundary; malformed

## DATE_LITERAL

- Category: FUNCTION
- Display name: DATE_LITERAL
- Description: Locale-independent Gregorian ISO date constructor for years 0100 through 9999
- Syntax: DATE
- Canonical ID/syntax: DATE_LITERAL / DATE
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; DATE
- Evaluator binding: DATE_LITERAL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: DATE("2026-10-15")
- Test contract: positive; negative; empty; boundary; malformed

## ALLROWS

- Category: QUANTIFIER
- Display name: ALLROWS
- Description: Typed row predicate quantifier with lexical aliases and explicit empty-collection semantics
- Syntax: ALLROWS
- Canonical ID/syntax: ALLROWS / ALLROWS
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;, BOOL) -&gt; BOOL
- Evaluator binding: ALLROWS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: ALLROWS(@ROWS AS t, TRUE)
- Test contract: positive; negative; empty; boundary; malformed

## ANYROW

- Category: QUANTIFIER
- Display name: ANYROW
- Description: Typed row predicate quantifier with lexical aliases and explicit empty-collection semantics
- Syntax: ANYROW
- Canonical ID/syntax: ANYROW / ANYROW
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;, BOOL) -&gt; BOOL
- Evaluator binding: ANYROW
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: ANYROW(@ROWS AS t, TRUE)
- Test contract: positive; negative; empty; boundary; malformed

## NOROWS

- Category: QUANTIFIER
- Display name: NOROWS
- Description: Typed row predicate quantifier with lexical aliases and explicit empty-collection semantics
- Syntax: NOROWS
- Canonical ID/syntax: NOROWS / NOROWS
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;, BOOL) -&gt; BOOL
- Evaluator binding: NOROWS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: NOROWS(@ROWS AS t, TRUE)
- Test contract: positive; negative; empty; boundary; malformed

## FILTER

- Category: QUANTIFIER
- Display name: FILTER
- Description: Select rows satisfying a typed Boolean predicate while preserving their schema
- Syntax: FILTER
- Canonical ID/syntax: FILTER / FILTER
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;, BOOL) -&gt; COLLECTION&lt;ROW&lt;T&gt;&gt;
- Evaluator binding: FILTER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: FILTER(@ROWS AS t, TRUE)
- Test contract: positive; negative; empty; boundary; malformed

## COUNT

- Category: FUNCTION
- Display name: COUNT
- Description: Count collection elements or distinct normalized typed list elements
- Syntax: COUNT
- Canonical ID/syntax: COUNT / COUNT
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;) -&gt; NUMBER
- Evaluator binding: COUNT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: COUNT(@ROWS)
- Test contract: positive; negative; empty; boundary; malformed

## DISTINCT_COUNT

- Category: FUNCTION
- Display name: DISTINCT_COUNT
- Description: Count collection elements or distinct normalized typed list elements
- Syntax: DISTINCT_COUNT
- Canonical ID/syntax: DISTINCT_COUNT / DISTINCT_COUNT
- Negated alias: FALSE
- Signatures: (LIST&lt;T&gt;) -&gt; NUMBER
- Evaluator binding: DISTINCT_COUNT
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: DISTINCT_COUNT(@VALUES)
- Test contract: positive; negative; empty; boundary; malformed

## EXISTS

- Category: FUNCTION
- Display name: EXISTS
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: EXISTS
- Canonical ID/syntax: EXISTS / EXISTS
- Negated alias: FALSE
- Signatures: (COLLECTION&lt;ROW&lt;T&gt;&gt;) -&gt; BOOL
- Evaluator binding: EXISTS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: EXISTS(@ROWS)
- Test contract: positive; negative; empty; boundary; malformed

## SOURCE

- Category: FUNCTION
- Display name: SOURCE
- Description: Bind to a registered source schema and add a dependency without opening source data
- Syntax: SOURCE
- Canonical ID/syntax: SOURCE / SOURCE
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; T
- Evaluator binding: SOURCE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: SOURCE("SOURCE_ID")
- Test contract: positive; negative; empty; boundary; malformed

## STARTS_WITH

- Category: OPERATOR
- Display name: STARTS_WITH
- Description: Text relation; locale-independent registered evaluator required
- Syntax: STARTS_WITH
- Canonical ID/syntax: STARTS_WITH / STARTS_WITH
- Negated alias: FALSE
- Signatures: (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: STARTS_WITH
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @NAME STARTS_WITH "text"
- Test contract: positive; negative; empty; boundary; malformed

## ENDS_WITH

- Category: OPERATOR
- Display name: ENDS_WITH
- Description: Text relation; locale-independent registered evaluator required
- Syntax: ENDS_WITH
- Canonical ID/syntax: ENDS_WITH / ENDS_WITH
- Negated alias: FALSE
- Signatures: (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: ENDS_WITH
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @NAME ENDS_WITH "text"
- Test contract: positive; negative; empty; boundary; malformed

## LIKE

- Category: OPERATOR
- Display name: LIKE
- Description: Text relation; locale-independent registered evaluator required
- Syntax: LIKE
- Canonical ID/syntax: LIKE / LIKE
- Negated alias: FALSE
- Signatures: (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: LIKE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @NAME LIKE "text"
- Test contract: positive; negative; empty; boundary; malformed

## MATCHES

- Category: OPERATOR
- Display name: MATCHES
- Description: Text relation; locale-independent registered evaluator required
- Syntax: MATCHES
- Canonical ID/syntax: MATCHES / MATCHES
- Negated alias: FALSE
- Signatures: (TEXT, TEXT) -&gt; BOOL
- Evaluator binding: MATCHES
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @NAME MATCHES "text"
- Test contract: positive; negative; empty; boundary; malformed

## EMPTY

- Category: OPERATOR
- Display name: EMPTY
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: EMPTY
- Canonical ID/syntax: EMPTY / EMPTY
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: EMPTY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD EMPTY
- Test contract: positive; negative; empty; boundary; malformed

## NOT_EMPTY

- Category: OPERATOR
- Display name: NOT_EMPTY
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: NOT_EMPTY
- Canonical ID/syntax: NOT_EMPTY / NOT_EMPTY
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: NOT_EMPTY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD NOT_EMPTY
- Test contract: positive; negative; empty; boundary; malformed

## VALID

- Category: OPERATOR
- Display name: VALID
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: VALID
- Canonical ID/syntax: VALID / VALID
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: VALID
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD VALID
- Test contract: positive; negative; empty; boundary; malformed

## INVALID

- Category: OPERATOR
- Display name: INVALID
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: INVALID
- Canonical ID/syntax: INVALID / INVALID
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: INVALID
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD INVALID
- Test contract: positive; negative; empty; boundary; malformed

## IS_NULL

- Category: OPERATOR
- Display name: IS_NULL
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: IS NULL
- Canonical ID/syntax: IS_NULL / IS NULL
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: IS_NULL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD IS NULL
- Test contract: positive; negative; empty; boundary; malformed

## IS_NOT_NULL

- Category: OPERATOR
- Display name: IS_NOT_NULL
- Description: Explicit inspection of resolved value state; never silently coerce MISSING or INVALID
- Syntax: IS NOT NULL
- Canonical ID/syntax: IS_NOT_NULL / IS NOT NULL
- Negated alias: FALSE
- Signatures: (T) -&gt; BOOL
- Evaluator binding: IS_NOT_NULL
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @FIELD IS NOT NULL
- Test contract: positive; negative; empty; boundary; malformed

## BETWEEN

- Category: OPERATOR
- Display name: BETWEEN
- Description: Inclusive lower and upper bounds; AND separates bounds
- Syntax: BETWEEN
- Canonical ID/syntax: BETWEEN / BETWEEN
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER, NUMBER) -&gt; BOOL; (DATE, DATE, DATE) -&gt; BOOL
- Evaluator binding: BETWEEN
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @VALUE BETWEEN 1 AND 10
- Test contract: positive; negative; empty; boundary; malformed

## NOT_IN

- Category: OPERATOR
- Display name: NOT_IN
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: NOT IN
- Canonical ID/syntax: IN / IN
- Negated alias: TRUE
- Signatures: (T, LIST&lt;T&gt;) -&gt; BOOL; (T, SET&lt;T&gt;) -&gt; BOOL; (IP_EXPRESSION, IP_EXPRESSION) -&gt; BOOL
- Evaluator binding: NOT_IN
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 5 / NONE
- Example: @VALUE NOT IN [1,2]
- Test contract: positive; negative; empty; boundary; malformed

## UPPER

- Category: FUNCTION
- Display name: UPPER
- Description: Text case conversion under the registered normalization policy
- Syntax: UPPER
- Canonical ID/syntax: UPPER / UPPER
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; TEXT
- Evaluator binding: UPPER
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: UPPER(@TEXT)
- Test contract: positive; negative; empty; boundary; malformed

## TRIM

- Category: FUNCTION
- Display name: TRIM
- Description: Remove leading/trailing permitted whitespace without tokenizing free text
- Syntax: TRIM
- Canonical ID/syntax: TRIM / TRIM
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; TEXT
- Evaluator binding: TRIM
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: TRIM(@TEXT)
- Test contract: positive; negative; empty; boundary; malformed

## LEN

- Category: FUNCTION
- Display name: LEN
- Description: Length of text in UTF-16 code units
- Syntax: LEN
- Canonical ID/syntax: LEN / LEN
- Negated alias: FALSE
- Signatures: (TEXT) -&gt; NUMBER
- Evaluator binding: LEN
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: LEN(@TEXT)
- Test contract: positive; negative; empty; boundary; malformed

## ABS

- Category: FUNCTION
- Display name: ABS
- Description: Absolute numeric value
- Syntax: ABS
- Canonical ID/syntax: ABS / ABS
- Negated alias: FALSE
- Signatures: (NUMBER) -&gt; NUMBER
- Evaluator binding: ABS
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: ABS(@NUMBER)
- Test contract: positive; negative; empty; boundary; malformed

## ROUND

- Category: FUNCTION
- Display name: ROUND
- Description: Numeric rounding to the requested integral decimal precision; ties to even
- Syntax: ROUND
- Canonical ID/syntax: ROUND / ROUND
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER
- Evaluator binding: ROUND
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: ROUND(@NUMBER,2)
- Test contract: positive; negative; empty; boundary; malformed

## MIN

- Category: FUNCTION
- Display name: MIN
- Description: Minimum/maximum of two compatible scalars or one typed list; empty lists require explicit policy
- Syntax: MIN
- Canonical ID/syntax: MIN / MIN
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER; (DATE, DATE) -&gt; DATE; (LIST&lt;NUMBER&gt;) -&gt; NUMBER; (LIST&lt;DATE&gt;) -&gt; DATE
- Evaluator binding: MIN
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: MIN(@A,@B)
- Test contract: positive; negative; empty; boundary; malformed

## MAX

- Category: FUNCTION
- Display name: MAX
- Description: Minimum/maximum of two compatible scalars or one typed list; empty lists require explicit policy
- Syntax: MAX
- Canonical ID/syntax: MAX / MAX
- Negated alias: FALSE
- Signatures: (NUMBER, NUMBER) -&gt; NUMBER; (DATE, DATE) -&gt; DATE; (LIST&lt;NUMBER&gt;) -&gt; NUMBER; (LIST&lt;DATE&gt;) -&gt; DATE
- Evaluator binding: MAX
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: MAX(@A,@B)
- Test contract: positive; negative; empty; boundary; malformed

## SUM

- Category: FUNCTION
- Display name: SUM
- Description: Sum/arithmetic mean of a numeric list; AVG on empty input is an error
- Syntax: SUM
- Canonical ID/syntax: SUM / SUM
- Negated alias: FALSE
- Signatures: (LIST&lt;NUMBER&gt;) -&gt; NUMBER
- Evaluator binding: SUM
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: SUM(@NUMBERS)
- Test contract: positive; negative; empty; boundary; malformed

## AVG

- Category: FUNCTION
- Display name: AVG
- Description: Sum/arithmetic mean of a numeric list; AVG on empty input is an error
- Syntax: AVG
- Canonical ID/syntax: AVG / AVG
- Negated alias: FALSE
- Signatures: (LIST&lt;NUMBER&gt;) -&gt; NUMBER
- Evaluator binding: AVG
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: AVG(@NUMBERS)
- Test contract: positive; negative; empty; boundary; malformed

## COALESCE

- Category: FUNCTION
- Display name: COALESCE
- Description: Two-argument typed default for explicit NULL; does not suppress ERROR, SKIPPED or MANUAL_REVIEW
- Syntax: COALESCE
- Canonical ID/syntax: COALESCE / COALESCE
- Negated alias: FALSE
- Signatures: (T, T) -&gt; T
- Evaluator binding: COALESCE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: COALESCE(@VALUE,@DEFAULT)
- Test contract: positive; negative; empty; boundary; malformed

## TODAY

- Category: FUNCTION
- Display name: TODAY
- Description: Date from fixed EvaluationContext run timestamp and configured timezone
- Syntax: TODAY
- Canonical ID/syntax: TODAY / TODAY
- Negated alias: FALSE
- Signatures: () -&gt; DATE
- Evaluator binding: TODAY
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: CONTEXT / RUN_CONTEXT
- Precedence/associativity: 1 / LEFT
- Example: TODAY()
- Test contract: positive; negative; empty; boundary; malformed

## DATE_DIFF

- Category: FUNCTION
- Display name: DATE_DIFF
- Description: Signed whole Gregorian days from start date to end date
- Syntax: DATE_DIFF
- Canonical ID/syntax: DATE_DIFF / DATE_DIFF
- Negated alias: FALSE
- Signatures: (DATE, DATE) -&gt; NUMBER
- Evaluator binding: DATE_DIFF
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: DATE_DIFF(@START,@END)
- Test contract: positive; negative; empty; boundary; malformed

## LOOKUP

- Category: FUNCTION
- Display name: LOOKUP
- Description: Source-schema lookup: key type and output column type are checked before execution
- Syntax: LOOKUP
- Canonical ID/syntax: LOOKUP / LOOKUP
- Negated alias: FALSE
- Signatures: (TEXT, T, TEXT) -&gt; T
- Evaluator binding: LOOKUP
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: LOOKUP("SOURCE_ID",@KEY,"COLUMN_ID")
- Test contract: positive; negative; empty; boundary; malformed

## SOURCE_VALUE

- Category: FUNCTION
- Display name: SOURCE_VALUE
- Description: Versioned registered capability; signatures and normalizer/validation policy define its compiler contract
- Syntax: SOURCE_VALUE
- Canonical ID/syntax: LOOKUP / LOOKUP
- Negated alias: FALSE
- Signatures: (TEXT, T, TEXT) -&gt; T
- Evaluator binding: SOURCE_VALUE
- Version/status: 1.0.0 / ACTIVE
- Determinism/cache: PURE / SESSION
- Precedence/associativity: 1 / LEFT
- Example: SOURCE_VALUE("SOURCE_ID",@KEY,"COLUMN_ID")
- Test contract: positive; negative; empty; boundary; malformed

