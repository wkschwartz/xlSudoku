Attribute VB_Name = "Tests"
' Copyright William Schwartz 2013.

Option Explicit

Private Sub TestStack()
    ' Test that stack pops items off in the order reverse of when they're
    ' pushed in.
    Dim s As New Stack, source As Variant, i As Long, sourceAnswer As Long, stackAnswer As Long
    Call s.Init
    source = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    For i = LBound(source) To UBound(source)
        sourceAnswer = source(i)
        Call s.Push(sourceAnswer)
    Next i
    For i = UBound(source) To LBound(source) Step -1
        sourceAnswer = source(i)
        stackAnswer = s.Pop()
        Debug.Assert stackAnswer = sourceAnswer
    Next i
    Debug.Print "Test passed."
End Sub
