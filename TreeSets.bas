Option Explicit


Private pLen  As Integer
Private pRoot As Nodes
Private pCnt As Integer

Public Property Get sign() As String
    sign = "TreeSets"
End Property

Public Property Get root() As Nodes
    Set root = pRoot
End Property

Public Function toString(Optional ByVal replaceDecimalPoint As Boolean = True) As String
    If pLen = 0 Then
        toString = "{}"
    Else
        Dim i
        
        Dim res As String
        res = "{"
        
        If replaceDecimalPoint Then
            For Each i In Me.toArray
                res = res & Replace("" & i, ",", ".") & ", "
            Next i
        Else
            For Each i In Me.toArray
                res = res & i & ", "
            Next i
        End If
        
        toString = left(res, Len(res) - 2) & "}"
    End If
    
End Function

Public Function p()
    
    Debug.Print Me.toString

End Function

Public Function clear()
    pLen = 0
    pCnt = 0
    Set pRoot = Nothing
End Function

Public Function ceiling(ByVal e) As Variant
    
    Dim res As New Nodes
    res.init(Nothing, Nothing, -1, 0).e
    
    Call ceiling_(e, pRoot, res)
    
    If res.index = -1 Then
        Set ceiling = Nothing
    Else
        ceiling = res.value
    End If

End Function

Public Function contains(ByVal e) As Boolean
    
    Dim res As Boolean
    res = False
    
    If Not isNothing(Me.ceiling(e)) Then
        res = Me.ceiling(e) = e
    End If
    
    contains = res

End Function

Public Function isEmpty() As Boolean
    isEmpty = pLen = 0
End Function

Public Function size() As Integer
    size = pLen
End Function

Public Function subSet(ByVal fromElement, ByVal toElement) As TreeSets
    Dim tree As New TreeSets
    tree.init
    
    tree.addAll toArray_(pRoot).filter("AND(_>=" & Replace("" & fromElement, ",", ".") & ", _<" & Replace("" & toElement, ",", ".") & ")")
    Set subSet = tree
End Function

Public Function pollFirst() As Variant

    Dim res
    res = Me.first()
    Me.remove res
    pollFirst = res
    
End Function

Public Function pollLast() As Variant

    Dim res
    res = Me.last()
    Me.remove res
    pollLast = res
    
End Function

Public Function remove(ByVal e) As Boolean
    
    Dim parent As Nodes
    Set parent = Nothing
    
    Dim res As Boolean
    res = False
    
    Dim tmp As Integer
    
    Call remove_(e, pRoot, parent, res, tmp)
    
    If res Then
        pLen = pLen - 1
    End If
    
    remove = res

End Function


Private Sub remove_(ByVal e, ByRef n As Nodes, ByRef parent As Nodes, ByRef res As Boolean, ByRef tmp As Integer)
    
    If e = n.value Then
        ' appendNode will add one back
        res = True
        
        If Not parent Is Nothing Then
            If parent.value > n.value Then
                parent.leftNode = Nothing
            Else
                parent.RightNode = Nothing
            End If
            
            If Not n.leftNode Is Nothing Then
                Call appendNode(n.leftNode, pRoot, tmp)
            End If
            
             If Not n.RightNode Is Nothing Then
                Call appendNode(n.RightNode, pRoot, tmp)
             End If
        Else
            'root element
            If n.leftNode Is Nothing Then
                Set pRoot = n.RightNode
            ElseIf n.RightNode Is Nothing Then
                 Set pRoot = n.leftNode
            Else
                Call appendNode(n.leftNode, n.RightNode, tmp)
                Set pRoot = n.RightNode
            End If
        End If
    ElseIf e > n.value Then
        
        If Not n.RightNode Is Nothing Then
            Call remove_(e, n.RightNode, n, res, tmp)
        End If
        
         
    Else
        If Not n.leftNode Is Nothing Then
            Call remove_(e, n.leftNode, n, res, tmp)
        End If
    End If
    

End Sub

Private Sub ceiling_(ByVal e, ByRef n As Nodes, ByRef res As Nodes)
    If e > n.value Then
        If Not n.RightNode Is Nothing Then
            Call ceiling_(e, n.RightNode, res)
        End If
    ElseIf e = n.value Then
        Set res = n
    Else
        If Not n.leftNode Is Nothing Then
            If e > n.leftNode.value Then
                res = e
            Else
                Call ceiling_(e, n.leftNode, res)
            End If
        Else
            Set res = n
        End If
    End If
End Sub

Public Property Get length() As Integer
    length = pLen
End Property

Public Property Get count() As Integer
    count = pCnt
End Property


Public Function init()
    pLen = 0
    pCnt = 0
End Function

Public Function add(ByVal val)
    Dim n As New Nodes
    
    n.init(Nothing, Nothing, pCnt, val).e
    
    If pLen > 0 Then
        
      Call appendNode(n, pRoot, pLen)

    Else
        Set pRoot = n
        pLen = 1
        pCnt = 0
    End If
    
    Set n = Nothing
    pCnt = pCnt + 1

End Function

Private Sub appendNode(ByRef n As Nodes, ByRef root As Nodes, ByRef l As Integer)
    
    Dim targVal, rootVal
    
    targVal = n.value
    rootVal = root.value
    
    If targVal > rootVal Then
        If Not root.RightNode Is Nothing Then
            Call appendNode(n, root.RightNode, l)
        Else
            root.RightNode = n
            l = l + 1
           
        End If
    ElseIf targVal < rootVal Then
        If Not root.leftNode Is Nothing Then
            Call appendNode(n, root.leftNode, l)
        Else
            root.leftNode = n
            l = l + 1
           
        End If
    End If

End Sub

Public Function addAll(ByVal val)
    Dim arr
    Dim i
    
    If isInstance(val, Array("Lists", "Sets", "TreeSets")) Then
        arr = val.toArray
    Else
        arr = val
    End If
    
    For Each i In arr
        Me.add i
    Next i
    
End Function

Public Function first() As Variant
    Dim tmp As New Nodes
    Set tmp = Me.root
    
    Do While Not tmp.leftNode Is Nothing
        Set tmp = tmp.leftNode
    Loop
    
    first = tmp.value
End Function

Public Function last() As Variant
    Dim tmp As New Nodes
    Set tmp = Me.root
    
    Do While Not tmp.RightNode Is Nothing
        Set tmp = tmp.RightNode
    Loop
    
    last = tmp.value
End Function
Private Function isNothing(ByVal e) As Boolean
    
    On Error GoTo handler1
    Dim res As Boolean
    
    res = e Is Nothing

handler1:
    If Err.Number <> 0 Then
        isNothing = False
    Else
        isNothing = res
    End If

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
