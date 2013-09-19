VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright William Schwartz 2013.

' Class Sudoku represents and solves a 9x9 sudoku problem. To use it, create a
' Sudoku instance (you can make as many as you like since this is class module)
' with
'
'   Dim s as new Sudoku
'
' Then call s.Init(a) is a 9x9 array of Longs 0 to 9. The Init method will
' will solve the sudoku and return whether a feasible solution exists. If so,
' query the result with s.Answer(r, c) where r, c are the 1-indexed row and
' column positions, or call s.OutputToExcelRange(square) to write the solution
' back to a Range square.

' The solver works by keeping a set of possible values for each cell in the
' domains array. Once it knows the answer to one cell, it can go through the
' cells in the row, column, and subsquare, and rule out that answer for those
' "neighbors" with the ruleOut method. This removes that answer from those
' neighbors' domains. Once a domain has only one element left, that's its
' answer. That answered cell is pushed onto the answeredStack so its
' constraints can be propogated next. During this process, if it ever happens
' that some row, column, or subsquare has a given value in only one cell, that
' cell is answered by that value (the mustBe method), and that cell can also be
' pushed onto the stack.

Option Explicit

Const MAX_ELEM As Long = 9 ' max value allowed in a cell
Const BAD_INIT_POSITIONS As Long = 514 ' error value

' First index of each array is the named by the attribute; second index is the
' particular cell value being counted.
Private Type Counts
    rows(1 To MAX_ELEM, 1 To MAX_ELEM) As Long
    columns(1 To MAX_ELEM, 1 To MAX_ELEM) As Long
    subsqs(1 To MAX_ELEM, 1 To MAX_ELEM) As Long
End Type

Private domains(1 To MAX_ELEM, 1 To MAX_ELEM) As SmallIntSet ' Stores possible answers for each cell
Private answers(1 To MAX_ELEM, 1 To MAX_ELEM) As Long ' Stores the answers
Private answeredStack As Stack ' Stores the index of cells to be propogated
Private valueCounts As Counts ' For ensuring each row/col/subsq has at least one of each value
Private feasible As Boolean

Public Function Init(initPos() As Long) As Boolean
    ' Initialize the Sudoku object with the MAX_ELEM x MAX_ELEM array of
    ' initial positions. Use 0s for empty cells. Solve the auction. Return true
    ' if a feasible solution was found, false otherwise.
    
    ' the subsquare-related functions aren't written flexibly, unfortunately.
    Debug.Assert MAX_ELEM = 9
    
    If LBound(initPos, 1) <> 1 Or LBound(initPos, 2) <> 1 Or _
            UBound(initPos, 1) <> MAX_ELEM Or UBound(initPos, 2) <> MAX_ELEM Then
        Err.Raise Number:=vbObjectError + BAD_INIT_POSITIONS, _
            Description:="initPos must be a square " & CStr(MAX_ELEM) & " on a side"
    End If
    
    feasible = True
    Set answeredStack = New Stack
    answeredStack.Init
    ' Read in the square and set up the domains
    Dim r As Long, c As Long, cellVal As Long
    For r = 1 To MAX_ELEM
        For c = 1 To MAX_ELEM
            Set domains(r, c) = New SmallIntSet
            Call domains(r, c).Init(1, MAX_ELEM)
            cellVal = initPos(r, c)
            If cellVal < 0 Or cellVal > MAX_ELEM Then
                Err.Raise Number:=vbObjectError + BAD_INIT_POSITIONS, _
                    Description:="initPos must contain only nubers 0 to " & CStr(MAX_ELEM)
            End If
            If cellVal > 0 Then
                Call Add(r, c, cellVal)
                answers(r, c) = cellVal
                Call answeredStack.Push(rcToIndex(r, c))
            Else ' When a cell has no initial position, it could be any number 1-9
                For cellVal = 1 To MAX_ELEM
                    Call Add(r, c, cellVal)
                Next cellVal
                answers(r, c) = 0
            End If
        Next c
    Next r
    
    Call Solve
    Init = feasible
End Function

Private Sub Add(r As Long, c As Long, v As Long)
    ' Add v to the domain of cell (r, c)
    If Not domains(r, c).Has(v) Then
        valueCounts.rows(r, v) = valueCounts.rows(r, v) + 1
        valueCounts.columns(c, v) = valueCounts.columns(c, v) + 1
        valueCounts.subsqs(rcToSubsq(r, c), v) = valueCounts.subsqs(rcToSubsq(r, c), v) + 1
    End If
    Call domains(r, c).Add(v)
End Sub

Private Sub Del(r As Long, c As Long, v As Long)
    ' Delete v from the domain of cell (r, c)
    If domains(r, c).Has(v) Then
        valueCounts.rows(r, v) = valueCounts.rows(r, v) - 1
        valueCounts.columns(c, v) = valueCounts.columns(c, v) - 1
        valueCounts.subsqs(rcToSubsq(r, c), v) = valueCounts.subsqs(rcToSubsq(r, c), v) - 1
    End If
    Call domains(r, c).Del(v)
End Sub

Public Function Answer(r As Long, c As Long) As Long
    ' Return the solution found for cell (r, c) where r and c are 1-indexed
    ' positions in the Sudoku grid. Return 0 if a solution was not found.
    Answer = answers(r, c)
End Function

Private Sub Solve()
    ' Propogate the row, column, and subsquare constraints from answered cells
    ' to neighboring cells until there are no more answered cells.
    Dim answr As Long, r As Long, c As Long
    Dim cellIndex As Long, cellR As Long, cellC As Long
    Do While answeredStack.Count() > 0
        cellIndex = answeredStack.Pop()
        cellR = indexToRow(cellIndex)
        cellC = indexToColumn(cellIndex)
        answr = answers(cellR, cellC)
        'Debug.Print "Propogating color " & CStr(answer) & " from cell (" & _
        '    CStr(cellR) & ", " & CStr(cellC) & ")"
        ' Rule out the cell's answer from the cell's row
        For c = 1 To MAX_ELEM
            If c <> cellC Then
                If Not ruleOut(cellR, c, answr, "row constraint") Then
                    Exit Sub
                End If
            End If
        Next c
        ' Rule out the cell's answer from the cell's column
        For r = 1 To MAX_ELEM
            If r <> cellR Then
                If Not ruleOut(r, cellC, answr, "col constraint") Then
                    Exit Sub
                End If
            End If
        Next r
        ' Rule out the cell's answer from the cell's subsquare
        For r = subSqStart(cellR) To subSqEnd(cellR)
            If r <> cellR Then
                For c = subSqStart(cellC) To subSqEnd(cellC)
                    If c <> cellC Then
                        If Not ruleOut(r, c, answr, "subsquare constraint") Then
                            Exit Sub
                        End If
                    End If
                Next c
            End If
        Next r
    Loop
End Sub

Private Function ruleOut(cellR As Long, cellC As Long, bad As Long, reason As String) As Boolean
    ' Mark that squareCell cannot have value bad and return whether the sudoku
    ' is still feasible. If it is not, set the instance-wide feasible flag to
    ' false. If cell squareCell now has an answer, set the answers vector and
    ' add squareCell to the stack. Reason is a string for debugging output
    If answers(cellR, cellC) = bad Then
        feasible = False
        ruleOut = False
        'Debug.Print "Cell (" & CStr(cellR) & ", " & CStr(cellC) & ") has no " & _
        '    "possible value: Sudoku infeasible"
        Exit Function
    End If

    'Debug.Print "  value((" & CStr(cellR) & ", " & CStr(cellC) & ")) != " & CStr(bad) & ": " & reason

    Call Del(cellR, cellC, bad)
    'Debug.Assert domains(cellR, cellC).Card() > 0
    Dim answr As Long
    ruleOut = True
    If domains(cellR, cellC).Card() = 1 And answers(cellR, cellC) = 0 Then
        Call answeredStack.Push(rcToIndex(cellR, cellC))
        ' Figure out what that one remaining feasible answer is
        For answr = 1 To MAX_ELEM
            If domains(cellR, cellC).Has(answr) Then
                answers(cellR, cellC) = answr
                Exit For
            End If
        Next answr
    End If
    If Not atLeastOneOfEachConstraint(cellR, cellC, bad) Then
        ' Not strictly necessary since it's the last statement in the function
        ' but do it anyway in case we add anything later
        Exit Function
    End If
End Function

Private Function mustBe(cellR As Long, cellC As Long, good As Long, reason As String) As Boolean
    ' Mark that (cellR, cellC) must have value good and return whether the
    ' sudoku is still feasible. If it is not, set the instance-wide feasible
    ' flag to false. If it is, cell (cellR, cellC) now has an answer, set the
    ' answers vector and add (cellR, cellC) to the stack. Reason is a string
    ' for debugging output
    mustBe = True
    If answers(cellR, cellC) <> 0 Then
        If answers(cellR, cellC) <> good Then
            mustBe = False
            feasible = False
            'Debug.Print "infeasible: value(" & CStr(squareCell) & ") != " & _
            '    CStr(good) & " because already == " & CStr(answers(squareCell))
        End If
        Exit Function
    End If

    'Debug.Print "  value((" & CStr(cellR) & ", " & CStr(cellC) & ")) = " & CStr(good) & ": " & reason

    Dim bad As Long
    For bad = 1 To MAX_ELEM
        If bad <> good Then
            Call Del(cellR, cellC, bad)
            ' Would be great to check the at-least-one-of-each-value in each
            ' row, column, and subsquare here just as we do after a call to
            ' ruleOut, but it always seems to run out of stack space. Perhaps
            ' a custom stack would solve this problem like we do with
            ' answeredStack
            'If Not atLeastOneOfEachConstraint(cellR, cellC, bad) Then
            '    Exit Function
            'End If
        End If
    Next bad
    'Debug.Assert domains(cellR, cellC).Card() = 1
    'Debug.Assert domains(cellR, cellC).Has(good)
    answers(cellR, cellC) = good
    Call answeredStack.Push(rcToIndex(cellR, cellC))
End Function

Private Function atLeastOneOfEachConstraint(cellR As Long, cellC As Long, bad As Long) As Boolean
    ' Did the deletion cause another cell in the row, column, or
    ' or subsquare to become the last man standing holding `bad` in
    ' its domain?
    Dim bail As Boolean, c As Long, r As Long, lastR As Long, lastC As Long
    ' Search the row
    If valueCounts.rows(cellR, bad) = 1 Then
        For c = 1 To MAX_ELEM
            If domains(cellR, c).Has(bad) Then
                lastC = c
                Exit For
            End If
        Next c
        If lastC <> cellC Then
            If Not mustBe(cellR, lastC, bad, "row requirement") Then
                Exit Function
            End If
        End If
    End If
    ' Search the column
    If valueCounts.columns(cellC, bad) = 1 Then
        For r = 1 To MAX_ELEM
            If domains(r, cellC).Has(bad) Then
                lastR = r
                Exit For
            End If
        Next r
        If lastR <> cellR Then
            If Not mustBe(lastR, cellC, bad, "col requirement") Then
                Exit Function
            End If
        End If
    End If
    ' Search the subsquare
    If valueCounts.subsqs(rcToSubsq(cellR, cellC), bad) = 1 Then
        For r = subSqStart(cellR) To subSqEnd(cellR)
            For c = subSqStart(cellC) To subSqEnd(cellC)
                If domains(r, c).Has(bad) Then
                    lastR = r
                    lastC = c
                    bail = True
                    Exit For
                End If
            Next c
            If bail Then
                Exit For
            End If
        Next r
        If lastR <> cellR And lastC <> cellC Then
            If Not mustBe(lastR, lastC, bad, "subsq requirement") Then
                Exit Function
            End If
        End If
    End If
    atLeastOneOfEachConstraint = True
End Function

Private Function rcToIndex(r As Long, c As Long) As Long
    ' Accept (row, column) values and return a index
    rcToIndex = (r - 1) * MAX_ELEM + c
End Function

Private Function indexToColumn(index As Long) As Long
    ' Return the column offset for square cell index.
    indexToColumn = 1 + ((index - 1) Mod MAX_ELEM)
End Function

Private Function indexToRow(index As Long) As Long
    ' Return the row offset for square cell index.
    indexToRow = 1 + ((index - 1) \ MAX_ELEM)
End Function

Private Function subSqStart(i As Long) As Long
    ' Return the row or column that begins the subsquares that contain row or
    ' column i
    Dim subsqsize As Long
    subsqsize = Sqr(MAX_ELEM)
    subSqStart = subsqsize * ((i - 1) \ 3) + 1
End Function

Private Function subSqEnd(i As Long) As Long
    ' Return the row or column that ends the subsquares that contain row or
    ' column i
    Dim subsqsize As Long
    subsqsize = Sqr(MAX_ELEM)
    subSqEnd = subSqStart(i) + (subsqsize - 1)
End Function

Private Function rcToSubsq(r As Long, c As Long) As Long
    ' Return the index of the subsquare for the given cell
    Dim subsqsize As Long
    subsqsize = Sqr(MAX_ELEM)
    rcToSubsq = subsqsize * ((r - 1) \ subsqsize) + (c - 1) \ subsqsize + 1
End Function

Public Function IsValidSquare(square As Range) As Boolean
    ' Return whether square is valid
    IsValidSquare = True
    If square.rows.Count <> 9 Or square.columns.Count <> 9 Then
        IsValidSquare = False
        Exit Function
    End If
    ' There should be at least 17 filled in cells for the Sudoku to have a
    ' unique solution
    Dim num As Integer, c As Range
    num = 0
    For Each c In square
        If Not IsEmpty(c.Value) And IsNumeric(c.Value) Then
            If c.Value > 0 And c.Value < 10 Then
                num = num + 1
            Else
                IsValidSquare = False
                Exit Function
            End If
        End If
    Next c
    If num < 17 Then
        IsValidSquare = False
    End If
End Function

Public Function Size() As Long
    ' Return the size of the sudoku, which is both the length of the side
    ' and the maximum value a cell can have.
    Size = MAX_ELEM
End Function
