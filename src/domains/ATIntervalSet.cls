VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIntervalSet"
Option Explicit
Private mMaximum As Double, mInitialized As Boolean, mFrozen As Boolean
Private mItems As Collection
Public Sub Initialize(ByVal maximum As Double)
    If mInitialized Then Err.Raise vbObjectError + 2600, , "DOMAIN_FROZEN"
    If maximum < 0# Or maximum > 4294967295# Or maximum <> Fix(maximum) Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    mMaximum = maximum: mInitialized = True
    Set mItems = New Collection
End Sub
Public Sub Add(ByVal low As Double, ByVal high As Double, ByVal budget As ATEvalBudget)
    Dim result As New Collection, existing As ATIntervalRange, range As ATIntervalRange, inserted As Boolean
    If mFrozen Then Err.Raise vbObjectError + 2600, , "DOMAIN_FROZEN"
    If Not mInitialized Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    If low < 0# Or high < low Or high > mMaximum Or low <> Fix(low) Or high <> Fix(high) Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    budget.Step
    For Each existing In mItems
        budget.Step
        If existing.High + 1# < low Then
            Append result, existing, budget
        ElseIf high + 1# < existing.Low Then
            If Not inserted Then
                Set range = New ATIntervalRange
                range.Initialize low, high
                Append result, range, budget
                inserted = True
            End If
            Append result, existing, budget
        Else
            If existing.Low < low Then low = existing.Low
            If existing.High > high Then high = existing.High
        End If
    Next existing
    If Not inserted Then
        Set range = New ATIntervalRange
        range.Initialize low, high
        Append result, range, budget
    End If
    Set mItems = result
End Sub
Private Sub Append(ByVal target As Collection, ByVal range As ATIntervalRange, ByVal budget As ATEvalBudget)
    budget.CheckList target.Count + 1
    target.Add range
End Sub
Public Sub Freeze()
    If Not mInitialized Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    mFrozen = True
End Sub
Public Property Get Frozen() As Boolean
    Frozen = mFrozen
End Property
Public Property Get Count() As Long
    Count = mItems.Count
End Property
Public Property Get Maximum() As Double
    Maximum = mMaximum
End Property
Public Function Item(ByVal index As Long) As ATIntervalRange
    Set Item = mItems(index)
End Function
Public Function Overlaps(ByVal other As ATIntervalSet, ByVal budget As ATEvalBudget) As Boolean
    Dim a As Long, b As Long, left As ATIntervalRange, right As ATIntervalRange
    a = 1: b = 1
    Do While a <= mItems.Count And b <= other.Count
        budget.Step
        Set left = mItems(a): Set right = other.Item(b)
        If left.Low <= right.High And right.Low <= left.High Then
            Overlaps = True
            Exit Function
        End If
        If left.High < right.High Then a = a + 1 Else b = b + 1
    Loop
End Function
Public Function IsSubsetOf(ByVal other As ATIntervalSet, ByVal budget As ATEvalBudget) As Boolean
    Dim a As Long, b As Long, left As ATIntervalRange, right As ATIntervalRange
    b = 1
    For a = 1 To mItems.Count
        budget.Step
        Set left = mItems(a)
        Do While b <= other.Count
            budget.Step
            Set right = other.Item(b)
            If right.High >= left.Low Then Exit Do
            b = b + 1
        Loop
        If b > other.Count Then Exit Function
        Set right = other.Item(b)
        If right.Low > left.Low Or right.High < left.High Then Exit Function
    Next a
    IsSubsetOf = True
End Function
