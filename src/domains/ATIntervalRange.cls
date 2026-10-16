VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIntervalRange"
Option Explicit
Private mLow As Double, mHigh As Double, mInitialized As Boolean
Public Sub Initialize(ByVal low As Double, ByVal high As Double)
    If mInitialized Then Err.Raise vbObjectError + 2600, , "DOMAIN_FROZEN"
    If low < 0# Or high < low Or high > 4294967295# Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    If low <> Fix(low) Or high <> Fix(high) Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    mLow = low: mHigh = high: mInitialized = True
End Sub
Public Property Get Low() As Double
    If Not mInitialized Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    Low = mLow
End Property
Public Property Get High() As Double
    If Not mInitialized Then Err.Raise vbObjectError + 2601, , "DOMAIN_INVALID_INTERVAL"
    High = mHigh
End Property
