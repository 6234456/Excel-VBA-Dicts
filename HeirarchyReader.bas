'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@ desc: to Read the Heirarchy in the excel worksheet into dicts
'@ dependency: Dicts
'@ since: 17.12.2014
'@ lastUpdate: 27.02.2015  reduce : reduce to first level
'@ author: Qiou Yang
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Private pDict As New Dicts
Private pLevels As Integer
Private pDataDict As New Dicts


' construct pLvl Dicts to hold the elements
' the last level dict Name -> Coefficient
' other levels  Name -> Dict of SubLevel
' eg dict_Ultra.dict("Summe sonst.betr. Erträge").dict("Mieterlöse").dict("4869 sonst. Mietertr. o. Steue")
' summary added
' the last but one lvl

Public Property Get dict() As Dicts
    Set dict = pDict
End Property

Public Property Let dict(d As Dicts)
    Set pDict = d
End Property

Public Property Get level() As Integer
    level = pLevels
End Property

Public Property Let level(l As Integer)
    pLevels = l
End Property



Public Function loadStructure(Optional ByVal targSht As String = "", Optional ByVal levels As Integer = 3, Optional ByVal startRow As Integer = 1, Optional ByVal startCol As Integer = 1, Optional ByVal endRow As Integer)
    
    Dim alteSht As Worksheet
    Set alteSht = ActiveSheet
    
    If targSht <> "" Then
        Worksheets(targSht).Activate
    End If
    
    If IsMissing(endRow) Or endRow = 0 Then
        endRow = Cells(Rows.Count, startCol + levels - 1).End(xlUp).Row
    End If
    
    Columns(startCol).Insert
    Cells(startRow, startCol).Value = "Root"
    
    Set pDict = loadStructure__(levels + 1, startRow, startCol, endRow)
    pLevels = levels
    
    Columns(startCol).Delete
    
End Function


Public Function loadData(ByVal targSht As String, ByVal keyCol As Integer, ByVal valCol As Integer, Optional ByVal startRow As Integer, Optional ByVal endRow As Integer, Optional ByVal reg As Variant, Optional ByVal ignoreNullVal As Boolean, Optional ByVal setNullValto As Variant)

    Call pDataDict.load(targSht, keyCol, valCol, startRow, endRow, reg, ignoreNullVal, setNullValto)
    
   ' Call print_d__(pDataDict, 1, 0)
    
    Call loadData__(pDataDict, pDict, pLevels)

End Function

Public Function loadDataFromDict(ByVal d As Dicts)
    
    Call loadData__(d, pDict, pLevels)
    
End Function

Public Function clone() As HeirarchyReader
    Dim res As HeirarchyReader
    Set res = New HeirarchyReader
    
    res.dict = clone__(pDict, pLevels)
    res.level = pLevels
   
    Set clone = res

End Function

Public Function summary() As Dicts
   Dim dict As New Dicts
   Call dict.ini
   Set summary = reduce__(pDict, pLevels, dict)
    
End Function

Public Function reduce() As Dicts
    Dim dict As New Dicts
    Call dict.ini
    Dim k
    
    For Each k In pDict.dict.keys
        dict.dict(k) = dict2Sum__(pDict.dict(k))
    Next k
    
    Set reduce = dict
   
End Function

Public Function toJSON(Optional ByVal ky As String = "root") As String

    Dim res As String
    res = "{""name"": """ & ky & """," & Chr(13) & """children"": [" & Chr(13)
    
    Dim k
    
  '  For Each k In pDict.dict.keys
        res = res & dict2json__(pDict) & ","
  '  Next k
    
    If Right(res, 1) = "," Then
        res = Left(res, Len(res) - 1)
    End If
    
    res = res & "]}"
    
    toJSON = res

End Function


Public Function print_d()
    Call print_d__(pDict, pLevels, 0)
End Function

' l  the outmost level
Private Function reduce__(ByVal d As Dicts, ByVal l As Integer, ByVal res As Dicts) As Dicts
    Dim k
    
    If l = 2 Then
        For Each k In d.dict.keys
           res.dict(k) = d.dict(k).reduce("")
        Next k
    Else
        For Each k In d.dict.keys
           Set res = reduce__(d.dict(k), l - 1, res)
        Next k
    End If
    
    Set reduce__ = res


End Function

'outmost level is 3
'lowest is 1
Private Function dict2Sum__(ByVal d As Dicts) As Double
    Dim k
    Dim res As Double
    
    For Each k In d.dict.keys
        If Not isDicts(d.dict(k)) Then
            res = res + d.reduce("")
            Exit For
        Else
            res = res + dict2Sum__(d.dict(k))
        End If
    Next k
    
    
    dict2Sum__ = res
End Function

Private Function dict2json__(ByVal d As Dicts) As String
    Dim k
    Dim j
    Dim cnt As Integer
    Dim res As String
    Dim toExit As Boolean
    toExit = False
    res = ""
    cnt = 1
    
    For Each k In d.dict.keys
        For Each j In d.dict(k).dict.keys
            
            If Not isDicts(d.dict(k).dict(j)) And cnt = 1 Then
                res = res & d.dict(k).toJSON(k) & ","
                toExit = True
                cnt = cnt + 1
            End If
            
        Next j
        
        cnt = 1
        
        If toExit Then
            toExit = False
            'Exit For
        Else
            res = res & "{""name"": """ & k & """, ""children"":[" & dict2json__(d.dict(k)) & "]},"
        End If
        
    Next k
    
    If Right(res, 1) = "," Then
        res = Left(res, Len(res) - 1)
    End If
    
    dict2json__ = res
End Function



Private Function print_d__(ByVal d As Dicts, ByVal l As Integer, ByVal cnt As Integer)
    Dim k
    
    If l > 1 Then
         For Each k In d.dict.keys
            Debug.Print String(cnt, "-") & k
            Call print_d__(d.dict(k), l - 1, cnt + 1)
         Next k
    Else
        For Each k In d.dict.keys
            Debug.Print k & "  " & d.dict(k)
        Next k
    End If


End Function

Private Function loadData__(ByVal srcDict As Dicts, ByVal targDict As Dicts, ByVal l As Integer)
     Dim k
    
     If l > 1 Then
         For Each k In targDict.dict.keys
            Call loadData__(srcDict, targDict.dict(k), l - 1)
         Next k
    Else
        For Each k In targDict.dict.keys
            ' if not in srcDict set val to 0
            If Not srcDict.dict.exists(Trim(CStr(k))) Then
                targDict.dict(k) = 0
            Else
                targDict.dict(k) = targDict.dict(k) * srcDict.dict(k)
            End If
        Next k
    End If
    

End Function

Public Function filter()

    Call filterNonePositive__(pDict, pLevels)

End Function

Private Function filterNonePositive__(ByVal targDict As Dicts, ByVal l As Integer)
     Dim k
    
     If l > 1 Then
         For Each k In targDict.dict.keys
            Call filterNonePositive__(targDict.dict(k), l - 1)
         Next k
    Else
        For Each k In targDict.dict.keys
            ' if not in srcDict set val to 0
            If targDict.dict(k) <= 0 Then
                targDict.dict.Remove (k)
            End If
        Next k
    End If

End Function



Private Function clone__(ByVal d As Dicts, ByVal l As Integer) As Dicts
    Dim res As New Dicts
    Dim k
    
    Call res.ini
    
    If l > 1 Then
         For Each k In d.dict.keys
            Set res.dict(k) = clone__(d.dict(k), l - 1)
         Next k
    Else
        For Each k In d.dict.keys
            res.dict(k) = d.dict(k)
        Next k
    End If
    
    Set clone__ = res

End Function


Private Function loadStructure__(ByVal levels As Integer, Optional ByVal startRow As Integer = 1, Optional ByVal startCol As Integer = 1, Optional ByVal endRow As Integer) As Dicts
    
    Dim res As Dicts
    
    Set res = New Dicts
    Call res.ini
    
    
    Dim rngEnd As Long
    Dim i
    Dim targVal
   
    
    If levels > 1 Then
        ' from the first cell down to next cell
        
        rngEnd = startRow + 1
        
        ' second not empty cell
        Do While rngEnd <= endRow
            If Not IsEmpty(Cells(rngEnd, startCol)) Then
                rngEnd = rngEnd - 1
                Exit Do
            Else
                rngEnd = rngEnd + 1
            End If
        Loop
        
        For i = startRow To rngEnd
            If levels > 2 Then
                targVal = Cells(i, startCol + 1).Value
                If Not IsEmpty(targVal) Then
                    ' to avoid the redundency in the lvl 1
                    Set res.dict(Trim(CStr(targVal))) = loadStructure__(levels - 1, i, startCol + 1, rngEnd)
                End If
            Else
                Set res = loadStructure__(1, startRow, startCol + 1, rngEnd)
                Exit For
            End If
        Next i
        
        
    Else
        
        Call res.load("", startCol, startCol + 1, startRow, endRow, , False, 1)

    End If
    
    
    Set loadStructure__ = res
    
End Function


Private Function isDicts(ByRef d As Variant) As Boolean

On Error GoTo errIsDicts
        
    Dim tmpDict As Object
    Set tmpDict = d.reg("\d{4}")
    
    
errIsDicts:
    If Err.Number <> 0 Then
        isDicts = False
    Else
        isDicts = True
    End If
    
End Function
