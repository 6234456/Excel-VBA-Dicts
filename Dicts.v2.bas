
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                          Util Class Dicts
'@author                        Qiou Yang
'@lastUpdate                    23.02.2017
'                               minor bugfix
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


Private pDict As Object
Private pRngCol As Long
Private pIsDictFilled As Boolean
Private pStrictMode As Boolean
Private pStrictModeReg As Object
Private pReversedMode As Boolean
Private pLevel As Long
Private pIsNamed As Boolean
Private pNamedArray As Dicts


Public Property Get dict() As Object
    Set dict = pDict
End Property

Public Property Get named() As Dicts
    If pIsNamed Then
        Set named = pNamedArray
    Else
        Set named = Nothing
    End If
End Property

Public Function setNamed(ByVal rng As Variant) As Dicts
   On Error GoTo namedArrayHdl
   
   Dim s As String
   Dim c
   Dim cnt As Long
   
   cnt = 0
   
   Dim d As New Dicts
   Call d.ini
   
   s = rng.Address
   
namedArrayHdl:

    If Err.Number = 0 Then
        For Each c In rng.Cells
            d.dict(Trim(CStr(c.Value))) = cnt
            cnt = cnt + 1
        Next c
        
        Me.setNamed d
    Else
        Set pNamedArray = rng
    End If
    
   pIsNamed = True
   
   Set setNamed = Me
End Function

Public Property Let columnRange(ByVal rng As Long)
   pRngCol = rng
End Property

Public Property Get Count() As Long
    On Error GoTo cntArrayHdl

    Count = pDict.Count

cntArrayHdl:
    If Err.Number <> 0 Then
        Count = 0
    End If
End Property

Public Property Get keysArr() As Variant
    
    Dim res() As String
    
    If Me.Count > 0 Then
        ReDim res(0 To Me.Count - 1)
        
        Dim k
        Dim cnt As Long
        cnt = 0
        
        For Each k In Me.Keys
            res(cnt) = CStr(k)
            cnt = cnt + 1
        Next k
    End If
    
    keysArr = res
End Property

Public Property Get Keys() As Variant
    Keys = pDict.Keys
End Property


Public Property Let dict(ByRef dict As Object)
    Set pDict = dict
    'pIsDictFilled = True
End Property

Public Property Let strictModeReg(mode As Object)
    If Not pStrictMode Then
        pStrictMode = True
    End If
    
    Set pStrictModeReg = mode
    'pIsDictFilled = True
End Property

Public Property Let strictMode(mode As Boolean)
    On Error GoTo errhandler2
    Dim a As Boolean
    a = pStrictModeReg.test("")

errhandler2:
    If Err.Number = 0 And Not mode Then
        Set pStrictModeReg = Nothing
    End If

     pStrictMode = mode
    'pIsDictFilled = True
End Property


Public Property Let reversedMode(mode As Boolean)
   pReversedMode = mode
End Property

Public Property Let appendMode(mode As Boolean)
    If mode Then
        Call Me.ini
        pIsDictFilled = mode
    End If
    
End Property

Public Sub ini()
    
    On Error GoTo Errhandler1
    
    Dim a As Long
    a = pDict.Count
    
      
Errhandler1:
    If Err.Number <> 0 Then
        Set pDict = CreateObject("scripting.dictionary")
        pDict.compareMode = vbTextCompare
    End If
    
    ' pIsDictFilled = True

End Sub

' to add the shtName just through dict.productX("""'src'!{*}""").p
Public Sub loadAddress(ByVal targSht As String, ByVal targKeyCol As Long, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant, Optional isR1C1 As Boolean = False)
    
  ' store the name of current sheet

    Dim tmpname As String
    Dim i As Long
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    With Worksheets(targSht)
    
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.compareMode = vbTextCompare
        
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            targRowEnd = .Cells(Rows.Count, targKeyCol).End(xlUp).row
        End If
        
        Dim hasReg As Boolean
        hasReg = Not IsMissing(reg)
        Dim test As Boolean
        test = True
        
        
        Dim myKey As Variant
        Dim myVal As Variant
        
        ' pReversedMode
        Dim startOrder
        Dim endOrder
        Dim stepOrder
        
        
        If targRowBegine < targRowEnd Then
            Dim arr1()
            arr1 = .Cells(targRowBegine, targKeyCol).Resize(targRowEnd - targRowBegine + 1, 1).Value
            
            
            If pReversedMode Then
                startOrder = UBound(arr1)
                endOrder = LBound(arr1)
                stepOrder = -1
            Else
                endOrder = UBound(arr1)
                startOrder = LBound(arr1)
                stepOrder = 1
            End If
    
            For i = startOrder To endOrder Step stepOrder
                myKey = Trim(CStr(arr1(i, 1)))
                
                If Not isR1C1 Then
                    myVal = .Cells(i + targRowBegine - 1, targValCol).Address(0, 0)
                Else
                    myVal = .Cells(i + targRowBegine - 1, targValCol).Address(ReferenceStyle:=xlR1C1)
                End If
                
                If myKey <> "" Then
                
                    If hasReg Then
                       test = reg.test(myKey)
                    End If
                   
    
                    If test Then
                         dict(myKey) = myVal
                    End If
                    
                End If
                
                test = True
            Next
        ElseIf targRowBegine = targRowEnd Then
            myKey = Trim(CStr(.Cells(targRowBegine, targKeyCol).Value))
            myVal = .Cells(targRowBegine, targValCol).Address(0, 0)
            
            If myKey <> "" Then
                
                If hasReg Then
                   test = reg.test(myKey)
                End If
               
    
                If test Then
                     dict(myKey) = myVal
                End If
                    
            End If
            
        Else
            Err.Raise 8888, , "endRow must be bigger than startRow!"
        End If
   
    
    End With
    
    
    ' strictMode
    Dim k As Variant
   
    Dim tmpDict As Object
    Set tmpDict = CreateObject("scripting.dictionary")
    tmpDict.compareMode = vbTextCompare
    
    If pStrictMode Then
        If Not IsReg(pStrictModeReg) Then
        
            Dim defaultReg As Object
            Set defaultReg = CreateObject("vbscript.regexp")
            
            With defaultReg
                .pattern = "[_\W]"
                .Global = True
            End With
        
            For Each k In dict.Keys
                If defaultReg.test(k) Then
                    tmpDict(defaultReg.Replace(k, "")) = dict(k)
                Else
                    tmpDict(k) = dict(k)
                End If
            Next k
        Else
            For Each k In dict.Keys
                If pStrictModeReg.test(k) Then
                    tmpDict(pStrictModeReg.Execute(k)(0).submatches(0)) = dict(k)
                Else
                    tmpDict(k) = dict(k)
                End If
            Next k
        End If
        Set dict = tmpDict
    End If
    
    
    
    If Not pIsDictFilled Then
        Set pDict = dict
    Else
        Dim k1 As Variant
        For Each k1 In dict.Keys
            pDict(k1) = dict(k1)
        Next k1
    End If
    
    pLevel = 1
    
    
End Sub


Public Sub load(ByVal targSht As String, ByVal targKeyCol As Long, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant, Optional ByVal ignoreNullVal As Boolean, Optional ByVal setNullValTo As Variant)
    
  ' store the name of current sheet

    Dim tmpname As String
    Dim i As Long
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    With Worksheets(targSht)
    
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.compareMode = vbTextCompare
        
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            targRowEnd = .Cells(Rows.Count, targKeyCol).End(xlUp).row
        End If
        
        Dim hasReg As Boolean
        hasReg = Not IsMissing(reg)
        Dim test As Boolean
        test = True
        
        
        Dim hasIgnoreNull As Boolean
        hasIgnoreNull = (Not IsMissing(ignoreNullVal)) And ignoreNullVal
        
        Dim hasNullVal As Boolean
        hasNullVal = (Not IsMissing(setNullValTo))
        
       
        
        Dim myKey As Variant
        Dim myVal As Variant
        
        ' pReversedMode
        Dim startOrder
        Dim endOrder
        Dim stepOrder
        
        
        If targRowBegine < targRowEnd Then
            Dim arr1()
            Dim arr2()
            arr1 = .Cells(targRowBegine, targKeyCol).Resize(targRowEnd - targRowBegine + 1, 1).Value
            
            If Not IsArray(targValCol) Then
                arr2 = .Cells(targRowBegine, targValCol).Resize(targRowEnd - targRowBegine + 1, 1).Value
            Else
                arr2 = rngCol(targRowBegine, targRowEnd, targValCol)
            End If
            
            
            If pReversedMode Then
                startOrder = UBound(arr1)
                endOrder = LBound(arr1)
                stepOrder = -1
            Else
                endOrder = UBound(arr1)
                startOrder = LBound(arr1)
                stepOrder = 1
            End If
            
        
            For i = startOrder To endOrder Step stepOrder
                myKey = Trim(CStr(arr1(i, 1)))
                myVal = arr2(i, 1)
                
                If myKey <> "" Then
                
                    If hasReg Then
                       test = reg.test(myKey)
                    End If
                    
                    If test And hasIgnoreNull Then
                        test = (Trim(CStr(myVal)) <> "" And myVal <> 0)
                    End If
                    
                    If test Then
                        If hasNullVal And (Trim(CStr(myVal)) = "" Or myVal = 0) Then
                            dict(myKey) = setNullValTo
                            Else: dict(myKey) = myVal
                        End If
                    End If
                End If
                
                test = True
            Next
        Else
            myKey = Trim(CStr(.Cells(targRowBegine, targKeyCol).Value))
            
            If Not IsArray(targValCol) Then
                myVal = .Cells(targRowBegine, targValCol).Value
            Else
                myVal = rngCol(targRowBegine, targRowEnd, targValCol)(1, 1)
            End If
    
            
            If myKey <> "" Then
            
                If hasReg Then
                   test = reg.test(myKey)
                End If
                
                If test And hasIgnoreNull Then
                    test = (Trim(CStr(myVal)) <> "" And myVal <> 0)
                End If
                
                If test Then
                    If hasNullVal And (Trim(CStr(myVal)) = "" Or myVal = 0) Then
                        dict(myKey) = setNullValTo
                        Else: dict(myKey) = myVal
                    End If
                End If
            End If
        End If
   
    End With
    
    
    ' strictMode
    Dim k As Variant
   
    Dim tmpDict As Object
    Set tmpDict = CreateObject("scripting.dictionary")
    tmpDict.compareMode = vbTextCompare
    
    If pStrictMode Then
        If Not IsReg(pStrictModeReg) Then
        
            Dim defaultReg As Object
            Set defaultReg = CreateObject("vbscript.regexp")
            
            With defaultReg
                .pattern = "[_\W]"
                .Global = True
            End With
        
            For Each k In dict.Keys
                If defaultReg.test(k) Then
                    tmpDict(defaultReg.Replace(k, "")) = dict(k)
                Else
                    tmpDict(k) = dict(k)
                End If
            Next k
        Else
            For Each k In dict.Keys
                If pStrictModeReg.test(k) Then
                    tmpDict(pStrictModeReg.Execute(k)(0).submatches(0)) = dict(k)
                Else
                    tmpDict(k) = dict(k)
                End If
            Next k
        End If
        Set dict = tmpDict
    End If
    
    
    
    If Not pIsDictFilled Then
        Set pDict = dict
    Else
        Dim k1 As Variant
        For Each k1 In dict.Keys
            pDict(k1) = dict(k1)
        Next k1
    End If
    
    pLevel = 1
    
    
End Sub

Public Sub loadStruct(ByVal targSht As String, ByVal targKeyCol1 As Long, ByVal targKeyCol2 As Long, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant)
      ' store the name of current sheet

    Dim tmpname As String
    Dim i As Long
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    With Worksheets(targSht)
    
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.compareMode = vbTextCompare
        
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            targRowEnd = .Cells(Rows.Count, targKeyCol2).End(xlUp).row
        End If
        
        Dim hasReg As Boolean
        hasReg = Not IsMissing(reg)
        Dim test As Boolean
        test = True
        
        If IsArray(targValCol) Then
            ' the number of cols
            pRngCol = UBound(targValCol) - LBound(targValCol) + 1
            
            If pRngCol = 1 Then
                targValCol = targValCol(LBound(targValCol))
            End If
        Else
            pRngCol = 1
        End If
        
        Dim tmpPreviousRow As Long
        Dim tmpCurrentRow As Long
        Dim tmpDict As New Dicts
        
        tmpPreviousRow = targRowEnd
        tmpCurrentRow = tmpPreviousRow
        
        Do While tmpCurrentRow > targRowBegine
            tmpCurrentRow = .Cells(tmpCurrentRow, targKeyCol1).End(xlUp).row
            
            If pRngCol = 1 Then
                Call tmpDict.load(targSht, targKeyCol2, targValCol, tmpCurrentRow + 1, tmpPreviousRow, reg, True)
            Else
                Call tmpDict.loadRng(targSht, targKeyCol2, targValCol, tmpCurrentRow + 1, tmpPreviousRow, reg)
            End If
            
            Set dict(Trim(CStr(.Cells(tmpCurrentRow, targKeyCol1).Value))) = tmpDict
            
            Set tmpDict = Nothing
            
            tmpPreviousRow = tmpCurrentRow - 1
        Loop
    
    End With

    If Not pIsDictFilled Then
        Set pDict = dict
    Else
        Dim k As Variant
        For Each k In dict.Keys
            pDict(k) = dict(k)
        Next k
    End If
    
    pLevel = 2
    
End Sub


Public Sub loadRng(ByVal targSht As String, ByVal targKeyCol As Long, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant)
    
  ' store the name of current sheet
    Dim tmpname As String
    Dim i As Long
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    Worksheets(targSht).Activate
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.compareMode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.Count, targKeyCol).End(xlUp).row
    End If
    
    Dim hasReg As Boolean
    hasReg = Not IsMissing(reg)
    Dim test As Boolean
    test = True
    
    ' the number of cols
    pRngCol = UBound(targValCol) - LBound(targValCol) + 1
    
    Dim myKey As Variant
    Dim myVal As Variant
    Dim startOrder As Long
    Dim endOrder As Long
    Dim stepOrder As Long
    
    If targRowBegine < targRowEnd Then
        Dim arr1()
        Dim arr2()
        arr1 = Range(Cells(targRowBegine, targKeyCol), Cells(targRowEnd, targKeyCol))
        arr2 = rngArr(targRowBegine, targRowEnd, targValCol)
        
        If pReversedMode Then
            startOrder = UBound(arr1)
            endOrder = LBound(arr1)
            stepOrder = -1
        Else
            endOrder = UBound(arr1)
            startOrder = LBound(arr1)
            stepOrder = 1
        End If
    
    
        For i = startOrder To endOrder Step stepOrder
            myKey = Trim(CStr(arr1(i, 1)))
            myVal = arr2(i, 1)
            
            If myKey <> "" Then
            
                If hasReg Then
                   test = reg.test(myKey)
                End If
                
                If test Then
                    dict(myKey) = myVal
                End If
            End If
            
            test = True
        Next
    Else
        myKey = Trim(CStr(Cells(targRowBegine, targKeyCol).Value))

        myVal = rngArr(targRowBegine, targRowEnd, targValCol)(1, 1)
  
        If myKey <> "" Then
        
            If hasReg Then
               test = reg.test(myKey)
            End If

            
            If test Then
                dict(myKey) = myVal
            End If
        
        End If
    End If
   
    
    Worksheets(tmpname).Activate

    If Not pIsDictFilled Then
        Set pDict = dict
    Else
        Dim k As Variant
        For Each k In dict.Keys
            pDict(k) = dict(k)
        Next k
    End If
    
    pLevel = 1
End Sub

' rng can be Range Object or an array
Public Function frequencyCount(ByRef rng) As Dicts

    Dim res As New Dicts
    Call res.ini
    
    Dim k
   

    If Not IsArray(rng) Then
        For Each k In rng.Cells
            If Len(Trim(CStr(k.Value))) > 0 Then
                If res.exists(k.Value) Then
                    res.dict(CStr(k.Value)) = res.dict(CStr(k.Value)) + 1
                Else
                    res.dict(CStr(k.Value)) = 1
                End If
            End If
        Next k
    Else
         For Each k In rng
            If Len(Trim(CStr(k))) > 0 Then
                If res.exists(k) Then
                    res.dict(CStr(k)) = res.dict(CStr(k)) + 1
                Else
                    res.dict(CStr(k)) = 1
                End If
            End If
        Next k
    End If

    Set frequencyCount = res

End Function



Public Sub unload(ByVal shtName As String, ByVal keyCol As Long, ByVal startingRow As Long, ByVal startingCol As Long, Optional ByVal endRow As Long, Optional ByVal endCol As Long)

    Dim tmpname As String
    tmpname = ActiveSheet.Name
    
    If Trim(shtName) = "" Then
        shtName = tmpname
    End If

    
    With Worksheets(shtName)
    
    
       If IsMissing(endRow) Or endRow = 0 Then
           endRow = .Cells(Rows.Count, keyCol).End(xlUp).row
       End If
       
       Dim c
       
       
       If IsMissing(endCol) Or endCol = 0 Then
    
           For Each c In .Cells(startingRow, keyCol).Resize(endRow - startingRow + 1, 1).Cells
               If pDict.exists(Trim(CStr(c.Value))) Then
                   .Cells(c.row, startingCol).Value = pDict(Trim(CStr(c.Value)))
               End If
           Next c
       Else
           
           Dim tmpC As Long
           
           If pRngCol = 0 Then
               tmpC = endCol - startingCol + 1
           Else
               tmpC = pRngCol
           End If
           
           For Each c In .Cells(startingRow, keyCol).Resize(endRow - startingRow + 1, 1).Cells
               If pDict.exists(Trim(CStr(c.Value))) Then
                   .Cells(c.row, startingCol).Resize(1, tmpC) = pDict(Trim(CStr(c.Value)))
               End If
           Next c
       
       End If
       
    End With

End Sub


Public Sub dump(ByVal shtName As String, Optional ByVal keyCol As Long = 1, Optional ByVal startingRow As Long = 1, Optional ByVal startingCol As Long, Optional ByVal endCol As Long)


    If IsMissing(startingCol) Or startingCol = 0 Then
        startingCol = keyCol + 1
    End If
    
    'unload the key
    Worksheets(shtName).Cells(startingRow, keyCol).Resize(Me.Count, 1) = Application.WorksheetFunction.Transpose(Me.keysArr)
    
    Call Me.unload(shtName, keyCol, startingRow, startingCol, , endCol)

End Sub

Public Function exists(ByVal k) As Boolean
    
    exists = pDict.exists(Trim(CStr(k)))
    
End Function

Public Function item(ByVal k) As Variant
    
    item = pDict(Trim(CStr(k)))

End Function

Public Function getNamedVal(ByVal nm As String) As Dicts
        
    If pIsNamed Then
        Dim i As Long
        i = pNamedArray.item(nm)
        
        Set getNamedVal = Me.reduceRngX("if({i}=" & i & ",{v}+{*},{v})")
        
    Else
        Set getNamedVal = Nothing
    End If
   

End Function


' ________________________________________Class Collection Functions___________________________________________

Public Function minus(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
        If Not dict2.dict.exists(k) Then
            res.dict(k) = pDict(k)
        End If
    Next k
    
    Set minus = res
End Function

'
Public Function add(dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts

    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    
    res.dict = pDict
    
    For Each k In dict2.dict.Keys
        If Not pDict.exists(k) Then
            res.dict(k) = dict2.dict(k)
        ElseIf Not keepOriginalVal Then
            res.dict(k) = dict2.dict(k)
        End If
    Next k
    
    Set add = res
End Function

Public Function update(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
        If Not dict2.dict.exists(k) Then
            res.dict(k) = pDict(k)
        ElseIf pDict(k) <> dict2.dict(k) Then
            res.dict(k) = dict2.dict(k)
        Else
            res.dict(k) = pDict(k)
        End If
    Next k
    
    Set update = res

End Function

Public Function reduce(ByVal sign As String) As Variant
    Dim res As Variant
    Dim k
    
    
    If sign = "" Or sign = "+" Then
        res = 0
        For Each k In pDict.Keys
            res = res + pDict(k)
        Next k
    ElseIf sign = "*" Then
        res = 1
        For Each k In pDict.Keys
            res = res * pDict(k)
        Next k
    End If
    
    reduce = res
End Function

Public Function mapKey(ByRef d As Dicts) As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        If d.exists(k) Then
            res.dict(d.dict(k)) = pDict.item(k)
        End If
    Next k

    Set mapKey = res

End Function

'''''''
'@param   re: RegExp-Obj with group
'         pos: the position of the group which is designated as the new key
''''''''
Public Function mapKeyReg(ByRef re As Object, Optional ByVal pos As Long = 0) As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        If re.test(k) Then
            res.dict(re.Execute(k)(0).submatches(pos)) = pDict.item(k)
        End If
    Next k

    Set mapKeyReg = res

End Function


Public Function mapKeyX(ByVal operation As String, Optional ByVal placeholder As String = "{*}") As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        res.dict(Application.Evaluate(Replace(operation, placeholder, CStr(k)))) = pDict.item(k)
    Next k

    Set mapKeyX = res

End Function


'''''''
'@param   re: RegExp-Obj with group
'         pos: the position of the group which is designated as the new val
''''''''
Public Function mapValReg(ByRef re As Object, Optional ByVal pos As Long = 0) As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        If re.test(pDict.item(k)) Then
            res.dict(k) = re.Execute(pDict.item(k))(0).submatches(pos)
        End If
    Next k

    Set mapValReg = res

End Function

' dict(k) -> Array(1,1,1,1,1)  =>  dict(k) -> 5
Public Function reduceRng(ByVal sign As String) As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        res.dict(k) = reduceArray(pDict(k), sign)
    Next k
   
    
    Set reduceRng = res
    
End Function

' dict(k) -> Array(1,1,1,1,1)  =>  dict(k) -> 5
Public Function reduceRngX(ByVal operation As String, Optional ByVal initVal As Variant = 0, Optional ByVal placeholder As String = "{*}", Optional ByVal index As String = "{i}", Optional ByVal cumVal As String = "{v}", Optional ByVal hasThousandSep As Boolean = True) As Dicts
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    Dim k

    For Each k In pDict.Keys
        res.dict(k) = reduceArrayX(pDict(k), operation, initVal, placeholder, index, cumVal, hasThousandSep)
    Next k
   
    
    Set reduceRngX = res
    
End Function

Public Function reduceRngVertical(ByVal sign As String) As Variant
    Dim k
    Dim i
    Dim tmpCnt As Long
    tmpCnt = 1
    Dim arr()
    
    Dim u As Long
    Dim l As Long

    For Each k In pDict.Keys
        If tmpCnt = 1 Then
            u = UBound(pDict(k))
            l = LBound(pDict(k))
            ReDim arr(l To u)
            tmpCnt = 2
            
            If sign = "+" Then
                For i = l To u
                    arr(i) = 0
                Next i
            Else
                For i = l To u
                    arr(i) = 1
                Next i
            End If
            
        End If
        
        If sign = "+" Then
            For i = l To u
                arr(i) = arr(i) + pDict(k)(i)
            Next i
        Else
            For i = l To u
                arr(i) = arr(i) * pDict(k)(i)
            Next i
        End If

    Next k
   
    
    reduceRngVertical = arr


End Function

Private Function reduceArray(ByVal arr, ByVal sign As String) As Variant
    Dim res As Variant
    Dim k
    
    
    If sign = "" Or sign = "+" Then
        res = 0
        For Each k In arr
            res = res + k
        Next k
    ElseIf sign = "*" Then
        res = 1
        For Each k In arr
            res = res * k
        Next k
    End If
    
    reduceArray = res
    
End Function

Public Function reduceArrayX(ByVal arr, ByVal operation As String, Optional ByVal initVal As Variant = 0, Optional ByVal placeholder As String = "{*}", Optional ByVal index As String = "{i}", Optional ByVal cumVal As String = "{v}", Optional ByVal hasThousandSep As Boolean = True) As Variant
    
    Dim k
    Dim v
    Dim tmp As String
    
    
    If hasThousandSep Then
        For k = LBound(arr) To UBound(arr)
            tmp = Replace(arr(k) & "", ",", ".")
            initVal = Replace(initVal & "", ",", ".")
            initVal = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, tmp), index, k), cumVal, initVal))
        Next k
    Else
        For k = LBound(arr) To UBound(arr)
            initVal = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, arr(k) & ""), index, k), cumVal, initVal))
        Next k
    End If

    reduceArrayX = initVal
   

End Function


Public Function filterVal(ByVal operation As String, Optional ByVal placeholder As String = "{*}", Optional ByVal hasThousandSep As Boolean = True) As Dicts
    Dim k
    Dim tmp As String
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    If hasThousandSep Then
        For Each k In pDict.Keys
            tmp = Replace(pDict(k) & "", ",", ".")
            
            If Application.Evaluate(Replace(operation, placeholder, tmp)) Then
                res.dict(k) = pDict(k)
            End If
        Next k
    Else
        For Each k In pDict.Keys
            If Application.Evaluate(Replace(operation, placeholder, pDict(k) & "")) Then
                res.dict(k) = pDict(k)
            End If
        Next k
    End If

    Set filterVal = res
    
End Function

Public Function filterExclude(ByVal reg As Object) As Dicts
    
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
      If Not reg.test(k) Then
        res.dict(k) = pDict(k)
      End If
    Next k
    
    Set filterExclude = res
    
End Function

Public Function filterInclude(ByVal reg As Object) As Dicts
    
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
      If reg.test(k) Then
        res.dict(k) = pDict(k)
      End If
    Next k
    
    Set filterInclude = res
    
End Function

''''''''''''''''''''
'set all the elements to a constant
'default to be 1
''''''''''''''''''''

Public Function constDict(Optional ByVal constant As Variant) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    If IsMissing(constant) Then
        constant = 1
    End If
    
    For Each k In pDict.Keys
        res.dict(k) = constant
    Next k
    
    Set constDict = res

End Function




'''''''''''''''''''
'@param operand2 can be either number or Dicts
'       operation supports only the string
'''''''''''''''''''

Public Function product(ByVal operand2 As Variant, ByVal operation As String, Optional ByVal IsNumericOperation As Boolean = True) As Dicts
    Dim k
   
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

    
    If IsNumeric(operand2) Then
        ' if the second operand is numeric
        For Each k In Me.dict.Keys
            If IsNumericOperation Then
                res.dict(k) = Application.Evaluate(Application.WorksheetFunction.Substitute(pDict(k) & operation & operand2, ",", "."))
            Else
                res.dict(k) = Application.Evaluate(pDict(k) & operation & operand2)
            End If
        Next k
    Else
        For Each k In Me.dict.Keys
            If IsNumericOperation Then
               If operand2.dict.exists(k) Then
                    res.dict(k) = Application.Evaluate(Application.WorksheetFunction.Substitute(pDict(k) & operation & operand2.dict(k), ",", "."))
               End If
            Else
                If operand2.dict.exists(k) Then
                    res.dict(k) = Application.Evaluate(pDict(k) & operation & operand2.dict(k))
                End If
            End If
        Next k
    End If
   
    Set product = res
    
End Function


'''''''''''''''''''
'@param operation is the string to be converted, placeholder is {*} by default
'
'''''''''''''''''''

Public Function productX(ByVal operation As String, Optional ByVal placeholder As String = "{*}", Optional ByVal hasThousandSep As Boolean = True) As Dicts
    Dim k
    Dim tmp As String
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini

            If hasThousandSep Then
                For Each k In pDict.Keys
                    tmp = Replace(pDict(k) & "", ",", ".")
                    res.dict(k) = Application.Evaluate(Replace(operation, placeholder, tmp))
                Next k
            Else
                For Each k In pDict.Keys
                    res.dict(k) = Application.Evaluate(Replace(operation, placeholder, pDict(k) & ""))
                Next k
            End If
        
   
    Set productX = res
    
End Function

Public Function clone() As Dicts
        Dim res As Dicts
       Set res = clone__(Me, pLevel)
       
       With res
            .appendMode = pIsDictFilled
            .reversedMode = pReversedMode
       
       
            If pStrictMode Then
                 .strictMode = True
                 .strictModeReg = pStrictModeReg
            End If
       
       End With
       
       Set clone = res

End Function

Private Function clone__(ByVal d As Dicts, ByVal l As Long) As Dicts
    Dim res As New Dicts
    Dim k
    
    Call res.ini
    
    If l > 1 Then
         For Each k In d.dict.Keys
            Set res.dict(k) = clone__(d.dict(k), l - 1)
         Next k
    Else
        For Each k In d.dict.Keys
            res.dict(k) = d.dict(k)
        Next k
    End If
    
    Set clone__ = res

End Function



Public Function productRng(ByVal operand2 As Variant, ByVal operation As String) As Dicts
    Dim k
    Dim i
   
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    
    If IsNumeric(operand2) Then
        ' if the second operand is numeric

        For Each k In pDict.Keys
            res.dict(k) = productArr(pDict(k), operation, operand2)
        Next k
    Else
    
        For Each k In pDict.Keys
          
            If operand2.dict.exists(k) Then
                res.dict(k) = productArr(pDict(k), operation, operand2.dict(k))
            End If

        Next k
    End If
   
    Set productRng = res

End Function


Private Function productArr(ByVal arr1 As Variant, ByVal operation As String, ByVal arr2 As Variant) As Variant
    Dim res
    Dim i
    ReDim res(LBound(arr1) To UBound(arr1))
    
    If IsNumeric(arr2) Then
        For i = LBound(arr1) To UBound(arr1)
            res(i) = Application.Evaluate(Replace(arr1(i) & operation & arr2, ",", "."))
        Next i
    Else
        For i = LBound(arr1) To UBound(arr1)
            res(i) = Application.Evaluate(Replace(arr1(i) & operation & arr2(i), ",", "."))
        Next i
    End If
    
    productArr = res

End Function


' ______________________________ Print______________________________________________

Public Function p()
    
    ' check if the val is array
    Dim is_a As Boolean
    Dim k
    
    For Each k In Me.dict.Keys
        is_a = IsArray(Me.dict(k))
        Exit For
    Next k
    
    If is_a Then
         For Each k In Me.dict.Keys
            Debug.Print k & "  " & a_toString(Me.item(k))
        Next k
    Else
        For Each k In Me.dict.Keys
            Debug.Print k & "  " & Me.item(k)
        Next k
    End If
    
   

End Function

Private Function a_toString(ByVal arr As Variant) As String
    Dim res As String
    Dim i
    res = "["
    
    For Each i In arr
        res = res & Replace(" " & i, ",", ".") & ", "
    Next i
    
    res = Left(res, Len(res) - 2)
    
    
    a_toString = res & " ]"

End Function


Public Function pk()

    Dim k
    For Each k In Me.dict.Keys
        Debug.Print k
    Next k

End Function

Public Function ps(Optional ByVal lvl As Long = 1, Optional ByVal cnt As Long = 0)
    
    Dim k
    
    If cnt = lvl Then
        For Each k In Me.dict.Keys
            Debug.Print String(cnt, Chr(9)) & k & Chr(9) & Me.dict(k)
        Next k
    Else
        For Each k In Me.dict.Keys
            Debug.Print String(cnt, Chr(9)) & k
            Me.dict(k).ps lvl, cnt + 1
        Next k
    End If
    

End Function

Public Function toJSON(Optional ByVal k As String = "root") As String
    Dim res As String
    res = "{""name"":""" & k & """," & Chr(13)
    res = res & """children"":[" & Chr(13)
    
    Dim ky
    For Each ky In pDict.Keys
        res = res & "{""name"":""" & Replace(CStr(ky), """", "") & """, " & """size"": " & Replace(CStr(pDict(ky)), ",", ".") & "}," & Chr(13)
    Next ky
    
    toJSON = Left(res, Len(res) - 2) & Chr(13) & "]}"
    
    
End Function

' ________________________________________Util Functions____________________________________________
Public Function reg(ByVal pattern As String, Optional ByVal flag As String) As Object
    Dim obj As Object
    Set obj = CreateObject("vbscript.regexp")
    
    obj.pattern = pattern
    
    If IsMissing(flag) Then
        obj.IgnoreCase = False
    Else
    ' "gi"
        If InStr(StrConv(flag, vbLowerCase), "g") > 0 Then
            obj.Global = True
        End If
        
        ' i by default to false
        If InStr(StrConv(flag, vbLowerCase), "i") > 0 Then
            obj.IgnoreCase = True
        End If
    End If
    
    Set reg = obj
End Function

Public Function rng(ByVal start As Long, ByVal ending As Long)
    Dim res()
    ReDim res(0 To ending - start)
    
    Dim i As Long
    For i = start To ending
        res(i - start) = i
    Next i
    
    rng = res
End Function

Public Function y(Optional ByVal sht As String = "", Optional ByVal col As Long = 1, Optional ByVal wb As String = "") As Long
    
    y = getTargetWorksheet(sht, wb).Cells(Rows.Count, col).End(xlUp).row
    
End Function

Public Function x(Optional ByVal sht As String = "", Optional ByVal row As Long = 1, Optional ByVal wb As String = "") As Long
    
    x = getTargetWorksheet(sht, wb).Cells(row, Columns.Count).End(xlToLeft).Column
    
End Function

Private Function getTargetWorksheet(Optional ByVal sht As String = "", Optional ByVal wb As String = "") As Worksheet
    Dim shtObj As Worksheet

    If sht = "" Then
        Set shtObj = ActiveSheet
    Else
        If wb = "" Then
            Set shtObj = Worksheets(sht)
        Else
            Set shtObj = Workbooks(wb).Worksheets(sht)
        End If
    End If
    
    Set getTargetWorksheet = shtObj

End Function


' ________________________________________Util Functions End____________________________________________

' summe vom Range
Private Function rngCol(ByVal startRow As Long, ByVal endRow As Long, ByVal arrCol As Variant)
    Dim res()
    ReDim res(1 To endRow - startRow + 1, 1 To 1)
    
    Dim i As Long
    Dim j As Long
    
    Dim sum As Double
    
    
    For i = startRow To endRow
        For j = 0 To UBound(arrCol)
            If IsNumeric(Cells(i, arrCol(j)).Value) Then
             sum = sum + Cells(i, arrCol(j)).Value
            End If
        Next j
        
        res(i - startRow + 1, 1) = sum
        sum = 0
    Next i
    
    rngCol = res
    
End Function

Private Function rngArr(ByVal startRow As Long, ByVal endRow As Long, ByVal arrCol As Variant)
    Dim res()
    ReDim res(1 To endRow - startRow + 1, 1 To 1)
    
    Dim i As Long
    Dim j As Long
    
    Dim sum()
    ReDim sum(0 To UBound(arrCol))
    
    
    For i = startRow To endRow
        For j = 0 To UBound(arrCol)
            sum(j) = Cells(i, arrCol(j)).Value
        Next j
        
        res(i - startRow + 1, 1) = sum
        ReDim sum(0 To UBound(arrCol))
    Next i
    
    rngArr = res
    
End Function

Private Function IsReg(testObj As Object) As Boolean
    On Error GoTo errhandler3
    
    Dim a As Boolean
    a = testObj.test("")
    
errhandler3:
    If Err.Number = 0 Then
        IsReg = True
    Else
        IsReg = False
    End If

End Function

Public Function getTargetColumn(ByVal targSht As String, ByVal targCol As Long, Optional ByVal targRowBegine, Optional ByVal targRowEnd) As Range
    Dim tmpname As String
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If

    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Worksheets(targSht).Cells(Rows.Count, targCol).End(xlUp).row
    End If
    
    Set getTargetColumn = Worksheets(targSht).Cells(targRowBegine, targCol).Resize(targRowEnd - targRowBegine + 1, 1)

End Function

