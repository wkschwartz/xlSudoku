Attribute VB_Name = "Main"
' Copyright William Schwartz 2013.

Option Base 0
Option Explicit

Public Sub SolveSudoku(square As Range, Optional timeOutput As Range)
    Dim sud As New Sudoku, timeStart As Single, timeEnd As Single
    Dim feasible As Boolean
    Dim r As Long, c As Long
    Dim initPos() As Long
    
    ReDim initPos(1 To sud.Size(), 1 To sud.Size()) As Long
    
    If Not sud.IsValidSquare(square) Then
        Call MsgBox(square.Address & " does not contain a valid Sudoku square :(", _
            vbExclamation + vbOKOnly, "Non-Valid Sudoku Square")
        Exit Sub
    End If
    For r = 1 To sud.Size()
        For c = 1 To sud.Size()
            initPos(r, c) = square(r, c)
        Next c
    Next r

    timeStart = Timer
        feasible = sud.Init(initPos)
    timeEnd = Timer
    
    If IsMissing(timeOutput) Then
        Debug.Print timeEnd - timeStart
    Else
        With timeOutput
            .Value = timeEnd - timeStart
            .NumberFormat = "0.00000" ' Singles have <= 7 significant digits
        End With
    End If
    If feasible Then
        Call OutputToExcelRange(sud, square)
    Else
        Call MsgBox(square.Address & " has no feasible solution :(", vbExclamation + vbOKOnly, "Not Feasible")
    End If
End Sub

Private Sub OutputToExcelRange(sud As Sudoku, square As Range)
    ' Write the result back to square
    Dim r As Long, c As Long
    For r = 1 To sud.Size()
        For c = 1 To sud.Size()
            If sud.Answer(r, c) > 0 Then
                square(r, c) = sud.Answer(r, c)
            End If
        Next c
    Next r
End Sub
