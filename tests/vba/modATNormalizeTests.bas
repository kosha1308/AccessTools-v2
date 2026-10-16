Attribute VB_Name = "modATNormalizeTests"
Option Explicit
Public Sub AT_TestNormalize()
    Dim result As ATEvalResult, item As Variant
    Set result = AT_ParseDecimal("1.25", Nothing)
    AT_Assert result.Value.Scalar = 1.25, "invariant decimal"
    For Each item In Array("1,25", "1x", "NaN", "Infinity", String$(400, "9"), "0." & String$(400, "0") & "1")
        Set result = AT_ParseDecimal(CStr(item), Nothing)
        AT_Assert result.Value.ValueState = atValueInvalid, "invalid or out of range decimal"
    Next item
    Set result = AT_ParseDecimal("0." & String$(400, "0"), Nothing)
    AT_Assert result.Value.Scalar = 0#, "long exact zero"
    Set result = AT_NormalizeScalar("BOOL", "TRUE", Nothing)
    AT_Assert result.Value.ValueState = atValueInvalid, "Boolean text rejected"
    Set result = AT_NormalizeScalar("BOOL", True, Nothing)
    AT_Assert result.Value.Scalar = True, "Boolean native accepted"
    For Each item In Array("2024-02-29", "0100-01-01", "9999-12-31")
        Set result = AT_NormalizeScalar("DATE", CStr(item), Nothing)
        AT_Assert result.Value.ValueState = atValuePresent, "Gregorian boundary date"
    Next item
    For Each item In Array("1900-02-29", "2026-02-30", "0099-01-01", "16/10/2026")
        Set result = AT_NormalizeScalar("DATE", CStr(item), Nothing)
        AT_Assert result.Value.ValueState = atValueInvalid, "Gregorian invalid date"
    Next item
    AT_Assert AT_TextKey(" AbC ", "ASCII_TRIM_CI") = "abc", "ASCII trim comparison"
    AT_Assert AT_TextKey(" AbC ", "BINARY") = " AbC ", "binary preserves text"
    AT_Assert AT_AsciiCase(ChrW$(&H410) & "aZ", True) = ChrW$(&H410) & "AZ", "Unicode preserved in ASCII case conversion"
End Sub
