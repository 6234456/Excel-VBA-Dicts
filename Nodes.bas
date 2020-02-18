Option Explicit


Private pIndex As Long
Private pLeft As Nodes  ' the maximal length of array object
Private pRight As Nodes     ' the length of current List Object
Private pVal As Variant      ' the type of the class


Public Function toString() As String
    toString = "<Nodes #" & pIndex & " " & pVal & ">"
End Function

Public Property Get value() As Variant
    If IsObject(pVal) Then
        Set value = pVal
    Else
        value = pVal
    End If
End Property


Public Property Let leftNode(ByRef n As Nodes)
    
    Set pLeft = n
    
End Property

Public Property Let index(ByVal i As Long)
    
    pIndex = i

End Property

Public Property Let RightNode(ByRef n As Nodes)
    
    Set pRight = n
    
End Property

Public Property Get index() As Long
    
    index = pIndex

End Property

Public Property Get leftNode() As Nodes
    
    Set leftNode = pLeft
    
End Property

Public Property Get RightNode() As Nodes
    
    Set RightNode = pRight
    
End Property


Public Function init(ByRef l As Nodes, ByRef r As Nodes, ByVal i As Long, ByVal val As Variant) As Nodes
    
    Set pLeft = l
    Set pRight = r
    pIndex = i
    pVal = val
    
    Set init = Me
End Function
