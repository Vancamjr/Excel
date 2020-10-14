Attribute VB_Name = "JVC_Tools"
Public Sub Renumber_Tasks()
Attribute Renumber_Tasks.VB_ProcData.VB_Invoke_Func = "R\n14"
'Notes:  This code will run through the first column and renumber each cell that contains a value from row 3 to 500
Dim i, j, k, oldNum, newNum As Integer
Dim checkString As String
Dim oldStatusBar As String
Dim ws As Worksheet
Dim r As Range
Dim endRow As Integer
On Error Resume Next
Application.ScreenUpdating = False
Set ws = Application.ActiveSheet
Set r = Application.Range("A3:A300")
'find the last row for updating
For Each c In r.Cells
    If Len(c.Value) > 0 Then
        endRow = c.Row
    End If
Next c
i = 1
For Each c In r.Cells
    If Len(c.Value) > 0 Then
        oldNum = c.Value
        newNum = i
        c.Value = newNum
        If oldNum <> newNum Then
            For j = 3 To endRow  'run through and update the dependent tasks column
                checkString = ws.Cells(j, 4).Value  'the string we run through
                If oldNum <> newNum And Len(ws.Cells(j, 4).Value) > 0 Then
                    ws.Cells(j, 4).Value = UpdatePrecident(oldNum, newNum, ws.Cells(j, 4).Value)
                End If
            Next j
        End If
        i = i + 1
    End If
Next c
Application.ScreenUpdating = True
End Sub
Private Function UpdatePrecident(ByVal oldNum As Integer, ByVal newNum As Integer, ByVal checkString As String) As String
Dim tString, oldNumHash As String
Dim i, j As Integer
If oldNum = newNum Then
    UpdatePrecident = checkString
Else
    checkString = Trim(checkString)
    tString = "#"
    For i = 1 To Len(checkString)
        If InStr(1, "0123456789", Mid(checkString, i, 1), vbTextCompare) Then
            tString = tString & Mid(checkString, i, 1)
        ElseIf InStr(1, ",-", Mid(checkString, i, 1), vbTextCompare) Then
            tString = tString & "#" & Mid(checkString, i, 1) & "#"
        End If
    Next i
    tString = tString & "#"
    'Debug.Print "#1: ", checkString, tString
    oldNumHash = "#" & oldNum & "#"
    j = InStr(1, tString, oldNumHash, vbTextCompare)
    If j > 0 Then
        tString = Left(tString, j - 1) & "#" & newNum & "#" & Mid(tString, j + Len(oldNumHash))
    End If
    'Debug.Print "#2: ", checkString, tString

    UpdatePrecident = ""
    For i = 1 To Len(tString)
        If InStr(1, "#", Mid(tString, i, 1), vbTextCompare) < 1 Then
            UpdatePrecident = UpdatePrecident & Mid(tString, i, 1)
        End If
    Next i
End If
End Function
Public Function TestIdentLevel(r As Range) As Integer
Debug.Print r.Value
TestIdentLevel = r.IndentLevel
End Function
Public Sub Old_Renumber_Tasks()
'Notes:  This code will run through the first column and renumber each cell that contains a value from row 3 to 500
Dim i As Integer
On Error Resume Next
Dim r As Range
Set r = Application.Range("A3:A500")
i = 1
For Each c In r.Cells
    If Len(c.Value) > 0 Then
        c.Value = i
        i = i + 1
    End If
Next c
End Sub

