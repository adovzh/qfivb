Attribute VB_Name = "Qfi_Dates"
Public Function QFI_DAYADD(d As Date, lag As Integer) As Date
    QFI_DAYADD = DateAdd("d", lag, d)
End Function

Public Function QFI_RELDATEADD(d As Date, reldate As String) As Variant
    Dim num As Double, numStr As String
    Dim reldateLen As Integer
    
    reldateLen = Len(reldate)
    
    If reldateLen < 2 Then GoTo ErrSection
    
    numStr = Left(reldate, reldateLen - 1)
    If Not IsNumeric(numStr) Then GoTo ErrSection
    
    num = Val(numStr)
    
    Select Case Right(reldate, 1)
    Case "D"
    QFI_RELDATEADD = DateAdd("d", num, d)
    End Select
    Exit Function
ErrSection:
    QFI_RELDATEADD = CVErr(xlErrValue)
End Function
