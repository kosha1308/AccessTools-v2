Attribute VB_Name = "modATDomainParseTests"
Option Explicit
Public Sub AT_TestDomainParse()
    Dim result As ATEvalResult, other As ATEvalResult, expression As Variant, budget As New ATEvalBudget
    Dim limits As New ATEvalLimits, value As ATRuntimeValue, caught As Boolean, ignored As Variant
    budget.Initialize limits
    Set result = AT_ParseDomain("IP_EXPRESSION", "0.0.0.0/0", Nothing, budget)
    AT_Assert result.Value.Domain.Item(1).High = 4294967295#, "CIDR0 covers IPv4 universe"
    For Each expression In Array("192.0.2.1-3", "192.0.2.1,2,3", "192.0.2.1 ; 192.0.2.2" & vbLf & "192.0.2.3", "192.0.2.3-192.0.2.1", "192.0.2.1 " & ChrW$(&H2013) & " 192.0.2.3")
        Set result = AT_ParseDomain("IP_EXPRESSION", CStr(expression), Nothing, budget)
        AT_Assert result.State = atEvalOK, "domain parser state"
        AT_Assert result.Value.ValueState = atValuePresent, "domain parser valid compact/range"
        If result.Value.ValueState = atValuePresent Then AT_Assert result.Value.Domain.Item(1).High - result.Value.Domain.Item(1).Low = 2#, "compact equivalent coverage"
    Next expression
    For Each expression In Array("192.0.2.256", "192.0.2.12345", "192.0.2.1junk", "192.0.2.1,999", "192.0.2.1/33", "192.0.2.1/032", "192.0.2", "192.0.2.x1", "groupname", "192.0.2.1;bad", "2001:db8::1", "-1.0.0.0")
        Set result = AT_ParseDomain("IP_EXPRESSION", CStr(expression), Nothing, budget)
        AT_Assert result.Value.ValueState = atValueInvalid, "malformed entire IP token rejected"
    Next expression
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.2.x", Nothing, budget)
    AT_Assert result.Value.Domain.Count = 1 And result.Value.Domain.Item(1).High - result.Value.Domain.Item(1).Low = 255#, "trailing wildcard compressed"
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.x.1", Nothing, budget)
    AT_Assert result.Value.Domain.Count = 256, "internal wildcard exact disjoint intervals"
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1/32", Nothing, budget)
    Set other = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1", Nothing, budget)
    AT_Assert Not AT_IPDeclarationsEqual(result.Value, other.Value, budget), "declaration preserves CIDR32 versus plain"
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.2.57/24", Nothing, budget)
    Set other = AT_ParseDomain("IP_EXPRESSION", "192.0.2.0/24", Nothing, budget)
    AT_Assert Not AT_IPDeclarationsEqual(result.Value, other.Value, budget), "declaration preserves CIDR base"
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1,2,3", Nothing, budget)
    Set other = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1-3", Nothing, budget)
    AT_Assert AT_IPDeclarationsEqual(result.Value, other.Value, budget), "nonCIDR declaration coverage equivalent"
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1,3-5", Nothing, budget)
    Set other = AT_ParseDomain("IP_EXPRESSION", "192.0.2.1,192.0.2.3-5", Nothing, budget)
    AT_Assert result.Value.ValueState = atValuePresent, "compact range continuation uses prior prefix"
    If result.Value.ValueState = atValuePresent Then AT_Assert AT_IPDeclarationsEqual(result.Value, other.Value, budget), "compact range continuation equivalent coverage"
    Set result = AT_ParseDomain("PORT_SET", "0,443/8443\65535;100 - 102", Nothing, budget)
    AT_Assert result.Value.ValueState = atValuePresent, "port separators and numeric boundaries"
    For Each expression In Array("65536", "1.5", "1-0", "443oops", "-1", "443;bad", "000443")
        Set result = AT_ParseDomain("PORT_SET", CStr(expression), Nothing, budget)
        AT_Assert result.Value.ValueState = atValueInvalid, "invalid port token not salvaged"
    Next expression
    Set result = AT_ParseDomain("PORT_SET", " ", Nothing, budget)
    AT_Assert result.Value.ValueState = atValueEmpty, "empty domain distinct from invalid"
    Set result = AT_ParseDomain("PORT_SET", "443", Nothing, budget)
    Set value = result.Value
    On Error Resume Next
    value.Domain.Add 80#, 80#, budget
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "runtime domain payload immutable"
    On Error Resume Next
    ignored = value.Scalar
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "domain not silently exposed as scalar"
    limits.MaxSteps = 100
    budget.Initialize limits
    Set result = AT_ParseDomain("IP_EXPRESSION", "x.1.x.2", Nothing, budget)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "wildcard expansion work bounded"
    limits.MaxSteps = 200000
    limits.MaxListItems = 1
    budget.Initialize limits
    Set result = AT_ParseDomain("IP_EXPRESSION", "192.0.x.1", Nothing, budget)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "wildcard interval allocation bounded"
End Sub
