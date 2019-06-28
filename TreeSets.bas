 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class TreeSets, java TreeSet API implemented with VBA
'@dependency                               Lists, Nodes
'@author                                   Qiou Yang
'@license                                  MIT
'@lastUpdate                               28.06.2019
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


Private pLen  As Integer
Private pRoot As Nodes
Private pCnt As Integer

Private Sub Class_Initialize()
    init
End Sub

Private Sub Class_Terminate()
    clear
End Sub

Public Property Get root() As Nodes
    Set root = pRoot
End Property

Public Function toString(Optional ByVal replaceDecimalPoint As Boolean = True) As String
    If pLen = 0 Then
        toString = "{ }"
    Else
        Dim i
        
        Dim res As String
        res = "{ "
        
        For Each i In Me.toArray
            If replaceDecimalPoint And IsNumeric(i) Then
                res = res & Replace("" & i, ",", ".") & ", "
            Else
                res = res & i & ", "
            End If
        Next i
       
        toString = Left(res, Len(res) - 2) & " }"
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

Public Function ceiling(ByVal e, Optional ByVal asNode As Boolean = False) As Variant
    
    If Me.Count = 0 Then
        If asNode Then
          Set ceiling = Nothing
        Else
          ceiling = Null
        End If
    Else
        Dim res As New Nodes
        res.init Nothing, Nothing, -1, 0
        
        Call ceiling_(e, pRoot, res)
        
        If res.index = -1 Then
            If asNode Then
              Set ceiling = Nothing
            Else
              ceiling = Null
            End If
        Else
            If asNode Then
                Set ceiling = res
            Else
                ceiling = res.value
            End If
        End If
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
    Me.Remove res
    pollFirst = res
    
End Function

Public Function pollLast() As Variant

    Dim res
    res = Me.last()
    Me.Remove res
    pollLast = res
    
End Function

Public Function Remove(e As Variant) As Boolean
    
    Dim parent As Nodes
    Set parent = Nothing
    
    Dim res As Boolean
    res = False
    
    Dim tmp As Integer
    
    Call remove_(e, pRoot, parent, res, tmp)
    
    If res Then
        pLen = pLen - 1
    End If
    
    Remove = res

End Function


Private Function remove_(ByVal e, ByRef n As Nodes, ByRef parent As Nodes, ByRef res As Boolean, ByRef tmp As Integer)
    
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
End Function

Private Function ceiling_(ByVal e, ByRef n As Nodes, ByRef res As Nodes)
    If e > n.value Then
        If Not n.RightNode Is Nothing Then
            Call ceiling_(e, n.RightNode, res)
        End If
    ElseIf e = n.value Then
        Set res = n
    Else
        If Not n.leftNode Is Nothing Then
            Set res = n
            Call ceiling_(e, n.leftNode, res)
        Else
            Set res = n
        End If
    End If
End Function

Public Property Get length() As Integer
    length = pLen
End Property

Public Property Get Count() As Integer
    Count = pCnt
End Property

Public Function init()
    pLen = 0
    pCnt = 0
End Function

Public Function add(val As Variant, Optional ByVal updateIfDuplicated As Boolean = False)
    Dim n As New Nodes
    
    n.init Nothing, Nothing, pCnt, val
    
    If pLen > 0 Then
        
      Call appendNode(n, pRoot, pLen, updateIfDuplicated)

    Else
        Set pRoot = n
        pLen = 1
        pCnt = 0
    End If
    
    Set n = Nothing
    pCnt = pCnt + 1

End Function

Private Function appendNode(ByRef n As Nodes, ByRef root As Nodes, ByRef l As Integer, Optional ByVal updateIfDuplicated As Boolean = False)
    
    Dim targVal, rootVal
    
    targVal = n.value
    rootVal = root.value
    
    If targVal > rootVal Then
        If Not root.RightNode Is Nothing Then
            Call appendNode(n, root.RightNode, l, updateIfDuplicated)
        Else
            root.RightNode = n
            l = l + 1
        End If
    ElseIf targVal < rootVal Then
        If Not root.leftNode Is Nothing Then
            Call appendNode(n, root.leftNode, l, updateIfDuplicated)
        Else
            root.leftNode = n
            l = l + 1
        End If
    Else
        If updateIfDuplicated Then
            root.index = n.index
        End If
    End If

End Function

Public Function addAll(ParamArray val() As Variant)
    Dim arr
    Dim i
    
    For Each i In val
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
Private Function isNothing(e) As Boolean
    
 isNothing = TypeName(e) = "Nothing"

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
