'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                          Util Class Dicts
'@author                        Qiou Yang
'@lastUpdate                    26.10.2015
'                               add dump
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

Public Property Let columnRange(ByVal rng As Integer)
   pRngCol = rng
End Property

Public Property Get Count() As Integer
    Count = pDict.Count
End Property

Public Property Get keysArr() As Variant
    Dim res()
    ReDim res(0 To Me.Count - 1)
    
    Dim k
    Dim cnt As Integer
    cnt = 0
    
    For Each k In Me.Keys
        res(cnt) = k
        cnt = cnt + 1
    Next k
    
    keysArr = res
End Property

Public Property Get Keys() As Variant
    Keys = pDict.Keys
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
    a = pDict.Count
    
      
Errhandler1:
    If Err.Number <> 0 Then
        Set pDict = CreateObject("scripting.dictionary")
        pDict.compareMode = vbTextCompare
    End If
    
    ' pIsDictFilled = True

End Sub

' to add the shtName just through dict.productX("""'src'!{*}""").p
Public Sub loadAddress(ByVal targSht As String, ByVal targKeyCol As Integer, ByVal targValCol, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByVal reg As Variant, Optional isR1C1 As Boolean = False)
    
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
    dict.compareMode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.Count, targKeyCol).End(xlUp).Row
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
        arr1 = Range(Cells(targRowBegine, targKeyCol), Cells(targRowEnd, targKeyCol))
        
        
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
                myVal = Cells(i + targRowBegine - 1, targValCol).Address(0, 0)
            Else
                myVal = Cells(i + targRowBegine - 1, targValCol).Address(ReferenceStyle:=xlR1C1)
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
        myKey = Trim(CStr(Cells(targRowBegine, targKeyCol).Value))
        myVal = Cells(targRowBegine, targValCol).Address(0, 0)
        
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
   
    
    Worksheets(tmpname).Activate
    
    
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
    dict.compareMode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.Count, targKeyCol).End(xlUp).Row
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
    dict.compareMode = vbTextCompare
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.Count, targKeyCol).End(xlUp).Row
    End If
    
    Dim hasReg As Boolean
    hasReg = Not IsMissing(reg)
    Dim test As Boolean
    test = True
    
    ' the number of cols
    pRngCol = UBound(targValCol) - LBound(targValCol) + 1
    
    Dim myKey As Variant
    Dim myVal As Variant
    Dim startOrder As Integer
    Dim endOrder As Integer
    Dim stepOrder As Integer
    
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
    
End Sub



Public Sub unload(ByVal shtName As String, ByVal keyCol As Long, ByVal startingRow As Long, ByVal startingCol As Long, Optional ByVal endRow As Long, Optional ByVal endCol As Long)

    Dim tmpname As String
    tmpname = ActiveSheet.Name
    
    If Trim(shtName) = "" Then
        shtName = tmpname
    End If

    
    Worksheets(shtName).Activate
    
    
    If IsMissing(endRow) Or endRow = 0 Then
        endRow = Worksheets(shtName).Cells(Rows.Count, keyCol).End(xlUp).Row
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
Public Function add(dict2 As Dicts, Optional keepOriginalVal As Boolean) As Dicts

    If IsMissing(keepOriginalVal) Then
        keepOriginalVal = True
    End If
    
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

Public Function reduceRngVertical(ByVal sign As String) As Variant
    Dim k
    Dim i
    Dim tmpCnt As Integer
    tmpCnt = 1
    Dim arr()
    
    Dim u As Integer
    Dim l As Integer

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

Public Function filterExklude(ByVal reg As Object) As Dicts
    
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
      If Not reg.test(k) Then
        res.dict(k) = pDict(k)
      End If
    Next k
    
    Set filterExklude = res
    
End Function

Public Function filterInklude(ByVal reg As Object) As Dicts
    
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    Call res.ini
    
    For Each k In pDict.Keys
      If reg.test(k) Then
        res.dict(k) = pDict(k)
      End If
    Next k
    
    Set filterInklude = res
    
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
        
         
        For Each k In pDict.Keys
            If Not isNum Then
               
                res.dict(k) = Application.Evaluate(Application.WorksheetFunction.Substitute(pDict(k) & operation & operand2, ",", "."))
            Else
                res.dict(k) = Application.Evaluate(pDict(k) & operation & operand2)
            End If
        Next k
    Else
    
        For Each k In pDict.Keys
            If Not isNum Then
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
            Debug.Print k & "  " & a_toString(Me.dict(k))
        Next k
    Else
        For Each k In Me.dict.Keys
            Debug.Print k & "  " & Me.dict(k)
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
