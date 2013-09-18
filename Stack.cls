VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright William Schwartz 2013.

' Stack is a singly linked list implementing a stack interface. Insert
' values with Push. Remove values from with Pop. Count the size of the stack
' with Count. Call Init immediately after making a new stack.

Option Explicit

Private head As node
Private length As Long
Const EmptyStackErrorCode As Long = 513

Public Sub Init()
    ' Initialize a new stack.
    Set head = Nothing
    length = 0
End Sub

Public Function Count() As Long
    ' Return the number of elements in the stack
    Count = length
End Function

Public Sub Push(v As Long)
    ' Place v on the top of the stack.
    Dim newNode As New node
    newNode.Value = v
    If head Is Nothing Then
        Set newNode.nextNode = Nothing
    Else
        Set newNode.nextNode = head
    End If
    Set head = newNode
    length = length + 1
End Sub

Public Function Pop() As Long
    ' Return the number stored at the top of the stack
    If head Is Nothing Then
        Err.Raise Number:=vbObjectError + EmptyStackErrorCode, source:="Stack", Description:="Pop from empty stack."
    Else
        Dim newHead As node
        Pop = head.Value
        Set newHead = head.nextNode
        Set head.nextNode = Nothing
        Set head = newHead
        length = length - 1
    End If
End Function
