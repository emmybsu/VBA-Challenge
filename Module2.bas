Attribute VB_Name = "Module2"
Sub alpha2()
Dim ws As Worksheet

Dim Yearly As Double
'Yearly = (ws.Cells(i, 6).Value) - (opening_price)
For Each ws In Worksheets
    For i = 2 To 70000

    If ws.Cells(i, 11).Value >= 0 Then
                       ws.Cells(i, 11).Interior.ColorIndex = 4
                  Else

                       ws.Cells(i, 11).Interior.ColorIndex = 3
                 '  MsgBox (Yearly)
                 
        End If
    
        Next i
    Next ws
    
    
    
End Sub

