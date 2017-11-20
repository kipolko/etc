Attribute VB_Name = "Module1"
Sub findInSheet()

Dim c As Range
Dim firstAdd As String

On Error Resume Next
Set c = ThisWorkbook.Sheets(1).UsedRange.Find(what:="g", after:=Cells(1, 1))
On Error GoTo 0

firstaddress = c.Address

If c Is Not Nothing Then
    Do
        On Error Resume Next
        Set c = ThisWorkbook.Sheets(1).UsedRange.FindNext
        On Error GoTo 0
    Loop While c.Address <> firstAdd
End If


End Sub


Sub findInVBE()


With ThisWorkbook.VBProject.VBComponents
    For i = 1 To .Count
        Debug.Print .Item(i).Name
    Next i
End With

End Sub
