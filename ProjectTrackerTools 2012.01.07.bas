Attribute VB_Name = "JVC_Tools"
Public Sub Renumber_Tasks()
Attribute Renumber_Tasks.VB_ProcData.VB_Invoke_Func = "R\n14"
'Notes:  This code will run through the first column and renumber each cell that contains a value from row 3 to 500
Dim i, j, k, imax, oldNum, newNum As Integer
Dim checkString As String
Dim oldStatusBar As String
Dim data(300, 2) As Integer 'array (newnum,oldnum)
Dim ws As Worksheet
Dim r As Range
Dim endRow As Integer
On Error Resume Next
Application.ScreenUpdating = False
Set ws = Application.ActiveSheet
Set r = Application.Range("A3:A300") 'expand this to A300
'updates item numbers and find the last row
i = 1
For Each c In r.Cells
    If Len(c.Value) > 0 Then
        data(i, 1) = c.Value    'old value
        data(i, 2) = i          'new value
        'c.Value = i
        endRow = c.Row
        i = i + 1
    End If
Next c
imax = i - 1
''Update precidents for the tasks (rows) changing index numbers
For i = 1 To imax
    oldNum = data(i, 1)
    newNum = data(i, 2)
    Debug.Print data(i, 1), data(i, 2)
    If (newNum <> oldNum) And (oldNum <> 0) Then
        'these have changed
        For j = 3 To endRow
            If Len(ws.Cells(j, 4).Value) > 0 Then
                If ws.Cells(j, 4).Value <> UpdatePrecident(oldNum, newNum, ws.Cells(j, 4).Value) Then
                    Debug.Print oldNum, newNum, "From: " & ws.Cells(j, 4).Value, "To: " & UpdatePrecident(oldNum, newNum, ws.Cells(j, 4).Value)
                    'ws.Cells(j, 4).Value = UpdatePrecident(oldNum, newNum, ws.Cells(j, 4).Value)
                End If
            End If
        Next j
    End If
Next i
Application.ScreenUpdating = True
End Sub
Public Function UpdatePrecident(ByVal oldNum As Integer, ByVal newNum As Integer, ByVal checkString As String) As String
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

