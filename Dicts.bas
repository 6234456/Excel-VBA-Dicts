'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc          Util Class Dicts
'@lastUpdate    23.06.2014
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


Private pDict As Object
Private pRngCol As Integer
Private pIsDictFilled As Boolean
Private pStrictMode As Boolean
Private pStrictModeReg As Object
Private pReversedMode As Boolean


Public Property Get dict() As Object
    Set dict = pDict
End Property


Public Property Let dict(ByVal dict As Object)
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
        pIsDictFilled = True
    Else
        pIsDictFilled = False
    End If
    

End Property


Public Sub ini()
    
    On Error GoTo Errhandler1
    
    Dim a As Integer
    a = pDict.count
    
      
Errhandler1:
    If Err.Number <> 0 Then
        Set pDict = CreateObject("scripting.dictionary")
        pDict.comparemode = vbTextCompare
    End If
    
    ' pIsDictFilled = True

End Sub


Public Sub load(ByVal targSht As String, ByVal targKeyCol As Integer, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant, Optional ByVal ignoreNullVal As Boolean, Optional ByVal setNullValto As Variant)
    
  ' store the name of current sheet

    Dim tmpname As String
    Dim i As Integer
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    Worksheets(targSht).Activate
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.comparemode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.count, targKeyCol).End(xlUp).Row
    End If
    
    Dim hasReg As Boolean
    hasReg = Not IsMissing(reg)
    Dim test As Boolean
    test = True
    
    
    Dim hasIgnoreNull As Boolean
    hasIgnoreNull = (Not IsMissing(ignoreNullVal)) And ignoreNullVal
    
    Dim hasNullVal As Boolean
    hasNullVal = (Not IsMissing(setNullValto))
    
   
    
    Dim myKey As Variant
    Dim myVal As Variant
    
    ' pReversedMode
    Dim startOrder
    Dim endOrder
    Dim stepOrder
    
    
    If targRowBegine < targRowEnd Then
        Dim arr1()
        Dim arr2()
        arr1 = Range(Cells(targRowBegine, targKeyCol), Cells(targRowEnd, targKeyCol))
        
        If Not IsArray(targValCol) Then
            arr2 = Range(Cells(targRowBegine, targValCol), Cells(targRowEnd, targValCol))
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
                        dict(myKey) = setNullValto
                        Else: dict(myKey) = myVal
                    End If
                End If
            End If
            
            test = True
        Next
    Else
        myKey = Trim(CStr(Cells(targRowBegine, targKeyCol).Value))
        
        If Not IsArray(targValCol) Then
            myVal = Cells(targRowBegine, targValCol).Value
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
                    dict(myKey) = setNullValto
                    Else: dict(myKey) = myVal
                End If
            End If
        End If
    End If
   
    
    Worksheets(tmpname).Activate
    
    
    ' strictMode
    Dim k As Variant
   
    Dim tmpDict As Object
    Set tmpDict = CreateObject("scripting.dictionary")
    tmpDict.comparemode = vbTextCompare
    
    If pStrictMode Then
        If Not IsReg(pStrictModeReg) Then
        
            Dim defaultReg As Object
            Set defaultReg = CreateObject("vbscript.regexp")
            
            With defaultReg
                .pattern = "[_\W]"
                .Global = True
            End With
        
            For Each k In dict.keys
                If defaultReg.test(k) Then
                    tmpDict(defaultReg.Replace(k, "")) = dict(k)
                Else
                    tmpDict(k) = dict(k)
                End If
            Next k
        Else
            For Each k In dict.keys
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
        For Each k1 In dict.keys
            pDict(k1) = dict(k1)
        Next k1
    End If
    
    
    
End Sub
Public Sub loadRng(ByVal targSht As String, ByVal targKeyCol As Integer, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant)
    
  ' store the name of current sheet
    Dim tmpname As String
    Dim i As Integer
    
    tmpname = ActiveSheet.Name
    If Trim(targSht) = "" Then
        targSht = tmpname
    End If
    
    Worksheets(targSht).Activate
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.comparemode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.count, targKeyCol).End(xlUp).Row
    End If
    
    Dim hasReg As Boolean
    hasReg = Not IsMissing(reg)
    Dim test As Boolean
    test = True
    
    ' the number of cols
    pRngCol = UBound(targValCol) - LBound(targValCol) + 1
    
    Dim myKey As Variant
    Dim myVal As Variant
    
    If targRowBegine < targRowEnd Then
        Dim arr1()
        Dim arr2()
        arr1 = Range(Cells(targRowBegine, targKeyCol), Cells(targRowEnd, targKeyCol))
        
        
        arr2 = rngArr(targRowBegine, targRowEnd, targValCol)
    
    
        For i = LBound(arr1) To UBound(arr1)
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
        For Each k In dict.keys
            pDict(k) = dict(k)
        Next k
    End If
    
End Sub



Public Sub unload(ByVal shtName As String, ByVal keyCol As Long, ByVal startingRow As Long, ByVal startingCol As Long, Optional ByVal endRow As Long, Optional ByVal endCol As Long)

    Dim tmpname As String
    tmpname = ActiveSheet.Name
    
    If Trim(shtName) = "" Then
        shtName = tmpname
    End If

    
    Worksheets(shtName).Select
    
    
    If IsMissing(endRow) Or endRow = 0 Then
        endRow = Worksheets(shtName).Cells(Rows.count, keyCol).End(xlUp).Row
    End If
    
    Dim c
    
    
    If IsMissing(endCol) Or endCol = 0 Then
 
        For Each c In Range(Cells(startingRow, keyCol), Cells(endRow, keyCol)).Cells
            If pDict.exists(Trim(CStr(c.Value))) Then
                Cells(c.Row, startingCol).Value = pDict(Trim(CStr(c.Value)))
            End If
        Next c
    Else
        
        Dim tmpC As Integer
        
        If endCol <> 0 And pRngCol > endCol - startingCol + 1 Then
            tmpC = endCol - startingCol + 1
        Else
            tmpC = pRngCol
        End If
        
        For Each c In Range(Cells(startingRow, keyCol), Cells(endRow, keyCol)).Cells
            If pDict.exists(Trim(CStr(c.Value))) Then
                Cells(c.Row, startingCol).Resize(1, tmpC) = pDict(Trim(CStr(c.Value)))
            End If
        Next c
    
    End If
    Worksheets(tmpname).Activate

End Sub




' ________________________________________Class Collection Functions___________________________________________

Public Function minus(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.keys
        If Not dict2.dict.exists(k) Then
            res.dict(k) = pDict(k)
        End If
    Next k
    
    Set minus = res
End Function

'
Public Function add(dict2 As Dicts, Optional keepOriginalVal As Boolean) As Dicts

    If IsMissing(keepOriginalVal) Then
        keepOriginalVal = True
    End If
    
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    
    res.dict = pDict
    
    For Each k In dict2.dict.keys
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
    
    For Each k In pDict.keys
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
    
    For Each k In pDict.keys
        res.dict(k) = constant
    Next k
    
    Set constDict = res

End Function


'''''''''''''''''''
'@param operand2 can be either number or Dicts
'       operation supports only the string
'''''''''''''''''''

Public Function product(ByVal operand2 As Variant, ByVal operation As String, Optional ByVal IsNumericOperation As Boolean) As Dicts
    Dim k
    Dim isNum As Boolean
    isNum = True
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    If Not IsMissing(IsNumericOperation) Then
        If Not IsNumericOperation Then
            isNum = False
        End If
    Else
        isNum = True
    End If

   
    
    If IsNumeric(operand2) Then
        ' if the second operand is numeric
        
         
        For Each k In pDict.keys
            If Not isNum Then
               
                res.dict(k) = Application.Evaluate(Application.WorksheetFunction.Substitute(pDict(k) & operation & operand2, ",", "."))
            Else
                res.dict(k) = Application.Evaluate(pDict(k) & operation & operand2)
            End If
        Next k
    Else
    
        For Each k In pDict.keys
            If Not isNum Then
               
                res.dict(k) = Application.Evaluate(Application.WorksheetFunction.Substitute(pDict(k) & operation & operand2.dict(k), ",", "."))
            Else
                res.dict(k) = Application.Evaluate(pDict(k) & operation & operand2.dict(k))
            End If
        Next k
    End If
   
    Set product = res
    
End Function



' ________________________________________Util Functions____________________________________________
Public Function reg(ByVal pattern As String, Optional ByVal flag As String) As Object
    Dim obj As Object
    Set obj = CreateObject("vbscript.regexp")
    
    obj.pattern = pattern
    
    If IsMissing(flag) Then
        obj.IgnoreCase = True
    Else
    ' "gi"
        If InStr(StrConv(flag, vbLowerCase), "g") > 0 Then
            obj.Global = True
        End If
        
        ' i by default to true
        If InStr(StrConv(flag, vbLowerCase), "i") > 0 Then
            obj.IgnoreCase = False
        End If
    End If
    
    Set reg = obj
End Function

Public Function rng(ByVal start As Integer, ByVal ending As Integer)
    Dim res()
    ReDim res(0 To ending - start)
    
    Dim i As Integer
    For i = start To ending
        res(i - start) = i
    Next i
    
    rng = res
End Function


' ________________________________________Util Functions End____________________________________________

' summe vom Range
Private Function rngCol(ByVal startRow As Integer, ByVal endRow As Integer, ByVal arrCol As Variant)
    Dim res()
    ReDim res(1 To endRow - startRow + 1, 1 To 1)
    
    Dim i As Integer
    Dim j As Integer
    
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

Private Function rngArr(ByVal startRow As Integer, ByVal endRow As Integer, ByVal arrCol As Variant)
    Dim res()
    ReDim res(1 To endRow - startRow + 1, 1 To 1)
    
    Dim i As Integer
    Dim j As Integer
    
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



