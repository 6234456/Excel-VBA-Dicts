 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class HashSets
'@dependency                               Lists
'@author                                   Qiou Yang
'@license                                  MIT
'@lastUpdate                               07.07.2020
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private pCol As New Lists


Private Sub Class_Initialize()
    init
End Sub

Private Sub Class_Terminate()
    clear
End Sub

Public Function toString(Optional ByVal replaceDecimalPoint As Boolean = True) As String
    If pCol.length = 0 Then
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
    Set pCol = Nothing
End Function


Public Function contains(ByVal e) As Boolean

   contains = pCol.contains(e)

End Function

Public Function isEmpty() As Boolean
    isEmpty = pLen = 0
End Function

Public Function size() As Long
    size = pCol.length
End Function

Public Function ceiling(ByVal e, Optional ByVal asNode As Boolean = False) As Variant
    
    If Me.Count = 0 Then
        If asNode Then
          Set ceiling = Nothing
        Else
          ceiling = Null
        End If
    Else
        If pCol.indexOf(e) = -1 Then
            If asNode Then
              Set ceiling = Nothing
            Else
              ceiling = Null
            End If
        Else
            ceiling = pCol.indexOf(e)
        End If
    End If

End Function


Public Function Remove(e As Variant) As Boolean
    
    If Me.contains(e) Then
        pCol.Remove (e)
        Remove = True
    Else
        Remove = False
    End If

End Function

Public Property Get length() As Long
    length = pCol.length
End Property

Public Property Get Count() As Long
    Count = pCol.length
End Property

Public Function init()

End Function

Public Function add(val As Variant, Optional ByVal updateIfDuplicated As Boolean = False) As Boolean
   add = False
    If Not pCol.contains(val) Then
        pCol.add val
        
        add = True
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
    If IsObject(pCol.getVal(0)) Then
        Set first = pCol.getVal(0)
    Else
        first = pCol.getVal(0)
    End If
End Function

Public Function last() As Variant
    If IsObject(pCol.last()) Then
        Set last = pCol.last()
    Else
        last = pCol.last()
    End If
End Function
Private Function isNothing(e) As Boolean
    
 isNothing = TypeName(e) = "Nothing"

End Function

Public Function toArray()
    toArray = pCol.toArray
End Function
