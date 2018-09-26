Option Explicit


Function delete_all_sheets_except(ByVal match As String)
Dim ws As Worksheet

Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each ws In ActiveWorkbook.Sheets
    If Not UCase(Left(ws.Name, Len(match))) = UCase(match) Then
        ws.Delete
    End If
Next

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Function
