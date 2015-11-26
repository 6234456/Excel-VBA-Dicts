Option Explicit


Private pList As Lists              ' the underlying array object
Private pRoot As Nodes

Public Property Get sign() As String
    sign = "TreeSets"
End Property

Public Property Get root() As Nodes
    Set root = pRoot
End Property

Public Function toString() As String
    If pList.length = 0 Then
        toString = "{}"
    Else
        Dim i
        
        Dim res As String
        res = "{"
        
        For Each i In Me.toArray
            res = res & i & ", "
        Next i
        
        toString = left(res, Len(res) - 2) & "}"
    End If
    
End Function

Public Property Get length() As Integer
    length = pList.length
End Property

Public Property Get list() As Lists
    Set list = pList
End Property

Public Function init()
    Set pList = New Lists
    pList.init
End Function

Public Function add(ByVal val)
    Dim n As New Nodes
    
    n.init(Nothing, Nothing, pList.length, val).e
    
    If Me.length > 0 Then
        
      Call appendNode(n, pRoot)

    Else
        Set pRoot = n
    End If
    
    pList.add n
    Set n = Nothing

End Function

Private Sub appendNode(ByRef n As Nodes, ByRef root As Nodes)
    
    Dim targVal, rootVal
    
    targVal = n.value
    rootVal = root.value
    
    If targVal > rootVal Then
        If Not root.RightNode Is Nothing Then
            Call appendNode(n, root.RightNode)
        Else
            root.RightNode = n
        End If
    ElseIf targVal < rootVal Then
        If Not root.leftNode Is Nothing Then
            Call appendNode(n, root.leftNode)
        Else
            root.leftNode = n
        End If
    End If

End Sub

Public Function addAll(ByVal val)
    Dim arr
    Dim i
    
    If isInstance(val, Array("Lists", "Sets")) Then
        arr = val.toArray
    Else
        arr = val
    End If
    
    For Each i In arr
        Me.add i
    Next i
    
End Function

Public Function min() As Variant
    Dim tmp As New Nodes
    Set tmp = Me.root
    
    Do While Not tmp.leftNode Is Nothing
        Set tmp = tmp.leftNode
    Loop
    
    min = tmp.value
End Function

Public Function max() As Variant
    Dim tmp As New Nodes
    Set tmp = Me.root
    
    Do While Not tmp.RightNode Is Nothing
        Set tmp = tmp.RightNode
    Loop
    
    max = tmp.value
End Function


Private Function isInstance(ByVal obj, ByVal sign) As Boolean
    On Error GoTo listhandler
    
    Dim res As Boolean
    res = False
    
    Dim myType As String
    myType = obj.sign
    
listhandler:
    If Err.Number = 0 Then
        If Not IsArray(sign) Then
            res = (myType = sign)
        Else
            Dim e
    
            For Each e In sign
                If e = myType Then
                    res = True
                    Exit For
                End If
            Next e
        End If
    End If
    
    isInstance = res
End Function

Public Function toArray()
    toArray = toArray_(Me.root).toArray
End Function

Private Function toArray_(ByRef n As Nodes) As Lists
    Dim l As New Lists
    l.init
    
    If Not n Is Nothing Then
        l.addAll(toArray_(n.leftNode)).add(n.value).addAll toArray_(n.RightNode)
    End If
    
    Set toArray_ = l
End Function

