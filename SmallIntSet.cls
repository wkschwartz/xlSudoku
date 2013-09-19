VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmallIntSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright William Schwartz 2013.

' SmallIntSet represents mathematical sets of small integers.

Option Explicit

Const ElementOutOfBounds As Long = 513

Private elem() As Boolean
Private Count As Long

Public Sub Init(min As Long, max As Long)
    ' Initialize a SmallIntSet that can contain elements e such that
    ' min <= e <= max
    Dim i As Long
    ReDim elem(min To max)
    For i = min To max
        elem(i) = False
    Next i
    Count = 0
End Sub

Public Sub Add(element As Long)
    ' Add element to the set.
    If element < LBound(elem) Or element > UBound(elem) Then
        Err.Raise Number:=vbObjectError + ElementOutOfBounds, _
            Description:="element " & CStr(element) & " out of bounds"
    End If
    If Not elem(element) Then
        Count = Count + 1
        elem(element) = True
    End If
End Sub

Public Sub Del(element As Long)
    ' Delete element from the set.
    If element < LBound(elem) Or element > UBound(elem) Then
        Err.Raise Number:=vbObjectError + ElementOutOfBounds, _
            Description:="element " & CStr(element) & " out of bounds"
    End If
    If elem(element) Then
        Count = Count - 1
        elem(element) = False
    End If
End Sub

Public Function Has(element As Long) As Boolean
    ' Return whether the set contains element.
    If element < LBound(elem) Or element > UBound(elem) Then
        Err.Raise Number:=vbObjectError + ElementOutOfBounds, _
            Description:="element " & CStr(element) & " out of bounds"
    End If
    Has = elem(element)
End Function

Public Function Card() As Long
    ' Return the cardinality (size) of the set.
    Card = Count
End Function

Public Function Str() As String
    ' Return a string representation of the contents of this set.
    Dim i As Long, c As Long
    c = 0
    Str = ""
    For i = LBound(elem) To UBound(elem)
        If elem(i) Then
            c = c + 1
            Str = Str & CStr(i)
            If c < Count Then
                Str = Str & ","
            End If
        End If
    Next i
    Str = Str
End Function

