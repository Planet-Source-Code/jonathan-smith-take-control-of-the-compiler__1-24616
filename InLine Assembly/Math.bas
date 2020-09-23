Attribute VB_Name = "Math"

Public Function DivideBy2Normally(lngDividend As Long, lngIterations As Long) As Long
    Dim xIterations As Long
    For xIterations = 1 To lngIterations
        DivideBy2Normally = lngDividend / 2
    Next
End Function

Public Function DivideBy2ByShifting(lngDividend As Long, lngIterations As Long) As Long
    Dim xIterations As Long
    For xIterations = 1 To lngIterations
        DivideBy2ByShifting = lngDividend / 2
    Next
End Function

