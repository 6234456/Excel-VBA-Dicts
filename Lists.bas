Option Explicit


Private pArr()              ' the underlying array object
Private pMaxLen As Integer  ' the maximal length of array object
Private pLen As Integer     ' the length of current List Object
Private pType As String     ' the type of the class


Public Property Get length() As Integer
    length = pLen
End Property

Public Property Get sign() As String
    sign = pType
End Property


Public Function init() As Lists
    
    pMaxLen = 20
    pLen = 0
    pType = "Lists"
    ReDim pArr(0 To pMaxLen - 1)
    
    Set init = Me
End Function

Private Sub check()
    If pLen > pMaxLen * 0.75 Then
        pMaxLen = Int(pMaxLen * 1.5)
        
        ReDim Preserve pArr(0 To pMaxLen - 1)
    End If
End Sub

Private Sub override(ByRef list As Lists)
    pLen = list.length
    pArr = list.toArray
    pMaxLen = UBound(pArr) + 1
End Sub

Public Function isEmpty() As Boolean
    
    isEmpty = (pLen = 0)

End Function

Private Function isInstance(ByVal obj, ByVal sign As String) As Boolean
    On Error GoTo listhandler
    
    Dim res As Boolean
    res = False
    
    Dim myType As String
    myType = obj.sign
    
listhandler:
    If Err.Number = 0 Then
        res = (myType = sign)
    End If
    
    isInstance = res
End Function

Public Function fromRng(ByRef rng As Range, Optional ByVal orientation As String = "v") As Lists
    Dim res As New Lists
    res.init
    
    Dim rowNum As Integer
    rowNum = rng.Rows.Count
    
    Dim i
    
    If rowNum = 1 Then
        res.addAll rng.Value
    Else
        Dim tmp As New Lists
        
        For i = 1 To rowNum
            tmp.init
            tmp.addAll rng.Rows(i).Value
            res.add tmp
            Set tmp = Nothing
        Next i
    End If
    
    If orientation = "h" Or orientation = "H" Then
        Set res = res.zipMe
    End If
    
    Call override(res)
    Set fromRng = Me

End Function

Public Function add(ByVal ele) As Lists
   
    Call check
    
    If isInstance(ele, "Lists") Then
        Set pArr(pLen) = ele
    Else
        pArr(pLen) = ele
    End If
    
    pLen = pLen + 1
    Set add = Me
End Function


Public Function remove(ByVal ele) As Lists
    If Me.contains(ele) Then
        Set remove = Me.removeAt(Me.indexOf(ele))
    Else
        Set remove = Me
    End If
End Function

Public Function removeAt(ByVal index As Integer) As Lists
    
    Dim res As New Lists
    res.init
    
    Set res = Me.slice(, index).addList(Me.slice(index + 1))
    Call override(res)
    
    Set removeAt = Me

End Function

Public Function addAt(ByVal ele, ByVal index As Integer) As Lists
    Dim res As Lists
    Set res = Me.slice(, index).add(ele).addList(Me.slice(index))
    Call override(res)
    Set addAt = Me
End Function

Public Function addAllAt(ByVal eles, ByVal index As Integer) As Lists
    Dim res As Lists
    Set res = Me.slice(, index).addAll(eles).addList(Me.slice(index))
    Call override(res)
    Set addAllAt = Me
End Function

Public Function replaceAllAt(ByVal eles, ByVal index As Integer) As Lists
    Dim res As Lists
    Set res = Me.slice(, index).addAll(eles).addList(Me.slice(index + 1))
    Call override(res)
    Set replaceAllAt = Me
End Function

Public Function addAll(ByVal arr) As Lists
    
    Dim i

    If isInstance(arr, "Lists") Then
        For Each i In arr.toArray
            Me.add i
        Next i
    Else
        For Each i In arr
            Me.add i
        Next i
    End If
    
    Set addAll = Me
End Function

Public Function addList(ByRef l As Lists) As Lists
    
    If l.length > 0 Then
        Me.addAll (l.toArray)
    End If
    Set addList = Me
End Function

Public Function zip(ParamArray l() As Variant) As Lists
    Dim res As New Lists
    res.init
    
    Dim targLen As Integer  ' the length of res
    targLen = pLen
    
    Dim tmp
    Dim i
    
    For Each tmp In l
        If targLen > tmp.length Then
            targLen = tmp.length
        End If
    Next tmp
    
    For i = 0 To targLen - 1
        Dim tmpList As New Lists
        tmpList.init
        
        tmpList.add pArr(i)
        
        For Each tmp In l
            tmpList.add tmp.getVal(i)
        Next tmp
        
        res.add tmpList
        Set tmpList = Nothing
    Next i
    
    Set zip = res

End Function

' zip the Lists within the Lists
Public Function zipMe() As Lists
    If pLen <= 1 Then
       Set zipMe = Me
    Else
        Dim i
        Dim j
        Dim res As New Lists
        res.init
        
        Dim lenArr As New Lists
        lenArr.init
        
        For i = 0 To pLen - 1
            lenArr.add pArr(i).length
        Next i
        
        For j = 0 To lenArr.min - 1
            Dim tmp As New Lists
            tmp.init
            
            For i = 0 To pLen - 1
                tmp.add pArr(i).getVal(j)
            Next i
            
            res.add tmp
            Set tmp = Nothing
        Next j
        
        Set zipMe = res
    End If
End Function

Public Function getVal(ByVal index As Integer, Optional ByVal index2) As Variant
    If index >= pLen Or index < 0 Then
        Err.Raise 8888, , "ArrayIndexOutOfBoundException"
    End If
    
    
    If Not isInstance(pArr(index), "Lists") Then
        getVal = pArr(index)
    Else
        If IsMissing(index2) Then
            Set getVal = pArr(index)
        Else
            getVal = pArr(index).getVal(index2)
        End If
    End If

End Function


Public Function setVal(ByVal index As Integer, ByVal ele As Variant) As Lists
    If index >= pLen Or index < 0 Then
        Err.Raise 8888, , "ArrayIndexOutOfBoundException"
    End If
    
    If isInstance(ele, "Lists") Then
        Set pArr(index) = ele
    Else
        pArr(index) = ele
    End If
    
    Set setVal = Me
End Function

Public Function indexOf(ByVal ele) As Integer
    Dim i As Integer
    Dim hasFound As Boolean
    hasFound = False
    
    For i = 0 To pLen
        If pArr(i) = ele Then
            hasFound = True
            Exit For
        End If
    Next i
    
    If hasFound Then
        indexOf = i
    Else
        indexOf = -1
    End If
End Function

Public Function contains(ByVal ele) As Boolean
    contains = Me.indexOf(ele) > -1
End Function


Public Function min() As Variant
    Dim res
    res = pArr(0)
    
    Dim i As Integer

    For i = 1 To pLen - 1
        If pArr(i) < res Then
            res = pArr(i)
        End If
    Next i
    
    min = res
End Function

Public Function max() As Variant
    Dim res
    res = pArr(0)
    
    Dim i As Integer

    For i = 1 To pLen - 1
        If pArr(i) > res Then
            res = pArr(i)
        End If
    Next i
    
    max = res
End Function

Public Function avg() As Double
    avg = Me.reduce("?+_", 0) / pLen
End Function
Public Function containsAll(ByVal arr) As Boolean
    Dim res As Boolean
    res = True
    
    Dim i
    For Each i In arr
        If Not Me.contains(i) Then
            res = False
            Exit For
        Else
    Next i
    
    containsAll = res
End Function

Public Function subList(ByVal fromIndex As Integer, ByVal toIndex As Integer) As Lists
    Set subList = Me.slice(fromIndex, toIndex, 1)
End Function

Public Function every(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal replaceDecimalPoint As Boolean = True) As Boolean
    Dim res As Boolean
    res = True
    
    Dim i
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            If Not Application.Evaluate(Replace(judgement, placeholder, Replace("" & i, ",", "."))) Then
                res = False
                Exit For
            End If
        Next i
    Else
        For Each i In Me.toArray
            If Not Application.Evaluate(Replace(judgement, placeholder, "" & i)) Then
                res = False
                Exit For
            End If
        Next i
    End If
    
    every = res

End Function

Public Function some(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal replaceDecimalPoint As Boolean = True) As Boolean
    Dim res As Boolean
    res = False
    
    Dim i
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            If Application.Evaluate(Replace(judgement, placeholder, Replace("" & i, ",", "."))) Then
                res = True
                Exit For
            End If
        Next i
    Else
        For Each i In Me.toArray
            If Application.Evaluate(Replace(judgement, placeholder, "" & i)) Then
                res = True
                Exit For
            End If
        Next i
    End If
    
    some = res

End Function



''''''''''''
'@param     operation:              string to be evaluated, e.g. _*2 will be interpreated as ele * 2
'           placeholder:            placeholder to be replaced by the value
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
''''''''''''
Public Function map(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    
    Dim res As New Lists
    res.init
    
    Dim i
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            res.add (Application.Evaluate(Replace(operation, placeholder, Replace("" & i, ",", "."))))
        Next i
    Else
        For Each i In Me.toArray
            res.add (Application.Evaluate(Replace(operation, placeholder, "" & i)))
        Next i
    End If
    
    Set map = res
End Function

Public Function mapList(ByVal operation As String, Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    
    Dim tmpStr As String
    tmpStr = operation
    
    Dim res As New Lists
    res.init
    
    Dim idx As New Lists
    idx.init
    
    Dim i
    Dim j
    
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    
    With re
        .Global = True
        .Pattern = "#(\d+)(\D|\b)"
    End With
    
    For Each i In re.Execute(operation)
        idx.add CInt(i.submatches(0))
    Next i
    
    idx.sort
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            For Each j In idx.toArray
                tmpStr = Replace(tmpStr, "#" & j, Replace("" & i.getVal(j), ",", "."))
            Next j
            res.add Application.Evaluate(tmpStr)
            tmpStr = operation
        Next i
    Else
        For Each i In Me.toArray
            For Each j In idx.toArray
                operation = Replace(operation, "#" & j, "" & i.getVal(j))
            Next j
            res.add Application.Evaluate(tmpStr)
            tmpStr = operation
        Next i
    End If
    
    Set mapList = res
End Function

''''''''''''
'@param     judgement:              string to be evaluated and return Boolean, e.g. _>2 will be interpreated as ele > 2
'           placeholder:            placeholder to be replaced by the value
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
''''''''''''
Public Function filter(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    Dim res As New Lists
    res.init
    
    Dim i
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            If Application.Evaluate(Replace(judgement, placeholder, Replace("" & i, ",", "."))) Then
                res.add i
            End If
        Next i
    Else
        For Each i In Me.toArray
            If Application.Evaluate(Replace(judgement, placeholder, "" & i)) Then
                res.add i
            End If
        Next i
    End If
    
    Set filter = res
End Function


Public Function reduce(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True) As Variant
    Dim res
    Dim i
    
    res = initialVal
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            res = Application.Evaluate(Replace(Replace(operation, placeholder, Replace("" & i, ",", ".")), placeholderInitialVal, Replace("" & res, ",", ".")))
        Next i
    Else
        For Each i In Me.toArray
            res = Application.Evaluate(Replace(Replace(operation, placeholder, "" & i), placeholderInitialVal, "" & res))
        Next i
    End If
    
     reduce = res
End Function

Public Function product(ByVal operation As String, ByRef list2 As Lists, Optional ByVal placeholder As String = "_", Optional ByVal placeholderOther As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    Dim res As New Lists
    res.init
    
    Dim i As Integer
    
    Dim targLen As Integer
    targLen = Application.WorksheetFunction.min(pLen, list2.length) - 1
    
    If replaceDecimalPoint Then
        For i = 0 To targLen
            res.add Application.Evaluate(Replace(Replace(operation, placeholder, Replace("" & pArr(i), ",", ".")), placeholderOther, Replace("" & list2.getVal(i), ",", ".")))
        Next i
    Else
        For i = 0 To targLen
            res.add Application.Evaluate(Replace(Replace(operation, placeholder, "" & pArr(i)), placeholderOther, "" & list2.getVal(i)))
        Next i
    End If
    
     Set product = res
End Function



Public Function slice(Optional ByVal fromIndex, Optional ByVal toIndex, Optional ByVal step) As Lists

    Dim res As New Lists
    res.init
    
    If IsMissing(fromIndex) Then
        fromIndex = 0
    End If
    
    If IsMissing(toIndex) Then
        toIndex = pLen
    End If
    
     If IsMissing(step) Then
        step = 1
    End If
    
    If fromIndex < 0 Then
        fromIndex = pLen + fromIndex
    End If
    
    If toIndex < 0 Then
        toIndex = pLen + toIndex
    End If
    
    If fromIndex <> toIndex Then
        Dim i As Integer
        
        If step > 0 Then
            For i = fromIndex To toIndex - 1 Step step
                res.add pArr(i)
            Next i
        Else
            For i = toIndex - 1 To fromIndex Step step
                res.add pArr(i)
            Next i
        End If
    End If
    
    Set slice = res
End Function

Public Function toArray() As Variant
    On Error GoTo handler1
    Dim arr()
    
    If pLen > 0 Then
        ReDim arr(0 To pLen - 1)
        Dim i As Integer
        
        Dim tmp As Integer
        tmp = pArr(0).length
        
handler1:
        If Err.Number <> 0 Then
            For i = 0 To pLen - 1
                arr(i) = pArr(i)
            Next i
        Else
            For i = 0 To pLen - 1
                Set arr(i) = pArr(i)
            Next i
        End If
    Else
        arr = Array()
    End If
    
    toArray = arr

End Function


Public Function toString()
    On Error GoTo handler
    If pLen = 0 Then
        toString = "[]"
    Else
        Dim res As String
        res = "["
        
        Dim tmp As Integer
        tmp = pArr(0).length
        
        Dim i As Integer
        
handler:
        If Err.Number <> 0 Then
            For i = 0 To pLen - 1
                res = res & pArr(i) & ", "
            Next i
        Else
            For i = 0 To pLen - 1
                res = res & pArr(i).toString() & ", "
            Next i
        End If
        
        toString = Left(res, Len(res) - 2) & "]"
    End If
   
End Function

Public Function sort() As Lists
    Dim res As New Lists
    res.init
    
    Dim arr
    
    arr = Me.toArray
    Call QuickSort(arr, 0, pLen - 1)
    res.addAll arr
    
    Call override(res)
    
    Set sort = Me

End Function

Public Function reverse() As Lists
    
    Set reverse = Me.slice(, , -1)
    
End Function

Public Function p()
    Debug.Print Me.toString
End Function

Public Function e()
End Function

Public Function copy() As Lists
    Dim res As New Lists
    res.init
    
    res.addAll (Me.toArray)
    
    Set copy = res
    
End Function


Private Sub QuickSort(vArray As Variant, ByVal inLow As Integer, ByVal inHi As Integer)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Integer
  Dim tmpHi   As Integer

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) > pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot > vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub
