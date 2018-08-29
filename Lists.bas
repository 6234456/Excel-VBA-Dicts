 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class Lists
'@author                                   Qiou Yang
'@lastUpdate                               29.08.2018
'                                          integrate with Dicts
'                                          minor bug-fix
'@TODO                                     optional params
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Option Explicit


Private pArr()              ' the underlying array object
Private pMaxLen As Integer  ' the maximal length of array object
Private pLen As Integer     ' the length of current List Object


Public Property Get length() As Integer
    length = pLen
End Property

Public Property Get sign() As String
    sign = "Lists"
End Property

Public Function init() As Lists
    
    pMaxLen = 20
    pLen = 0
    ReDim pArr(0 To pMaxLen - 1)
    
    Set init = Me
End Function

Public Function toDict() As Dicts
    Dim res As New Dicts
    
    Dim i
    For i = 0 To pLen - 1
        res.dict.add Me.getVal(i), i
    Next i
    
    Set toDict = res
    Set res = Nothing

End Function


Public Function clear() As Lists
    Me.init
End Function

Private Sub Class_Initialize()
    Me.init
End Sub

Private Sub Class_Terminate()

End Sub

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

Public Function isEmptyList() As Boolean
    
    isEmptyList = True
    
    Dim i
    For i = 0 To pLen - 1
        If Not IsEmpty(pArr(i)) Then
            isEmptyList = False
            Exit For
        End If
    Next i

End Function

Private Function isInstance(obj, ByVal sign) As Boolean
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

Public Function isLists(testObj As Variant) As Boolean
   isLists = TypeName(testObj) = "Lists"
End Function

Private Function isObj(ByVal obj) As Boolean
    On Error GoTo listhandler
    
    Dim res As Boolean
    res = False
    
    Dim myType As String
    myType = obj.sign
    
listhandler:
    isObj = (Err.Number = 0)

End Function

'in case of 1 * N or N * 1 matrix, return 1-dimensional array

Public Function fromRng(ByRef rng As Range, Optional ByVal orientation As String = "v") As Lists
    Dim res As New Lists
    
    
    Dim rowNum As Integer
    rowNum = rng.Rows.count
    Dim colNum As Integer
    colNum = rng.Columns.count
    
    Dim i
    
    If rowNum = 1 Or colNum = 1 Then
        res.addAll rng.Value
    Else
        Dim tmp As New Lists
        
        For i = 1 To rowNum
            tmp.init
            If IsArray(rng.Rows(i).Value) Then
                tmp.add rng.Rows(i).Value
            Else
                tmp.addAll rng.Rows(i).Value
            End If
            
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

Public Function toRng(ByRef rng As Range)
    
    If Me.length > 0 Then
        Dim y
        y = pLen
        
        If y = 1 Then
            rng.Resize(1, pArr(0).length).Value = Me.toArray
        Else
            Dim lenArr As New Lists
            
            Dim i
            For i = 0 To pLen - 1
                If isInstance(pArr(i), "Lists") Then
                    lenArr.add pArr(i).length
                Else
                    lenArr.add 1
                End If
            Next i
            
            Dim maxLen As Integer
            maxLen = lenArr.max
            
            rng.Resize(1, maxLen).Cells.clear
            
            For i = 0 To pLen - 1
                If isInstance(pArr(i), "Lists") Then
                    rng.Offset(i, 0).Resize(1, pArr(i).length).Value = pArr(i).toArray
                Else
                    rng.Offset(i, 0).Value = pArr(i)
                End If
            Next i
        End If
    End If
    
End Function

Public Function fromArray(arr, Optional ByVal iter As Boolean = True) As Lists
    Dim l As New Lists
    
    If iter Then
        Dim i
        For Each i In arr
            If IsArray(i) Then
                l.add fromArray(i)
            Else
                l.add i
            End If
        Next i
        
        Set fromArray = l
    Else
        Set fromArray = l.addAll(arr)
    End If
    
    Set l = Nothing

End Function

Private Function serial(ByVal start As Long, ByVal ending As Long, Optional ByVal steps As Long = 1) As Variant
    Dim res()
    Dim cnt As Long
    cnt = -1
    Dim i As Long
    
    For i = start To ending Step steps
        cnt = cnt + 1
    Next i
    
    If cnt > -1 Then
        ReDim res(0 To cnt)
        
        Dim cnt1 As Long
        cnt1 = 0
        
        For i = start To ending Step steps
            res(cnt1) = i
            cnt1 = cnt1 + 1
        Next i
    End If
    
    serial = res
End Function

Public Function fromSerial(ByVal start As Long, ByVal ending As Long, Optional ByVal steps As Long = 1) As Lists
    
    Me.clear
    
    Set fromSerial = Me.addAll(serial(start, ending, steps))
End Function

Public Function ones(ByVal n As Long) As Lists
    
    Set ones = Me.fromSerial(1, n).map("_/({i}+1)")
    
End Function

Public Function add(ByVal ele, Optional ByVal keepOldElements As Boolean = True) As Lists
   
    Call check
    
    If Not keepOldElements Then
        Me.clear
    End If
    
    If isObj(ele) Then
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

Public Function addAll(ByVal arr, Optional ByVal keepOldElements As Boolean = True) As Lists

    If Not keepOldElements Then
        Me.clear
    End If
    
    Dim i

    If isInstance(arr, "Lists") Then
        For Each i In arr.toArray
            Me.add i
        Next i
    ElseIf IsArray(arr) Then
        If Not isArrayEmpty(arr) Then
            For Each i In arr
                Me.add i
            Next i
        End If
    Else
        Me.add arr
    End If
    
    Set addAll = Me
End Function

Private Function isArrayEmpty(arr) As Boolean
    
    On Error GoTo hdl
    
    Dim res  As Boolean
    res = True
    
    Dim tmp
    
    If IsArray(arr) Then
        tmp = arr(LBound(arr))
    End If
    
hdl:
    If Err.Number = 0 Then
        res = False
    End If
    
    isArrayEmpty = res

End Function

Public Function addList(ByRef l As Lists) As Lists
    
    If l.length > 0 Then
        Me.addAll (l.toArray)
    End If
    Set addList = Me
End Function

Public Function zip(ParamArray l() As Variant) As Lists
    Dim res As New Lists
    
    
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
    If pLen = 0 Then
        Set zipMe = Me
    ElseIf pLen = 1 Then
       If Me.isLists(Me.getVal(0)) Then
        Dim k
        Dim l As New Lists
        Dim tmp1 As New Lists
        
        For k = 0 To Me.getVal(0).length - 1
            l.add tmp1.add(Me.getVal(0).getVal(k))
            Set tmp1 = Nothing
        Next k
        
        Set zipMe = l
        Set l = Nothing
        Set tmp1 = Nothing
        
       Else
       
        Set zipMe = Me
       End If
       
    Else
        Dim i
        Dim j
        Dim res As New Lists
        
        
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
    
    
    If Not isObj(pArr(index)) Then
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
    
    If isObj(ele) Then
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

Public Function every(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True) As Boolean
    Dim res As Boolean
    res = True
    
    Dim i
    
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            If Not Application.Evaluate(Replace(Replace(judgement, placeholder, Replace("" & i, ",", ".")), idx, cnt)) Then
                res = False
                Exit For
            End If
            
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            If Not Application.Evaluate(Replace(Replace(judgement, placeholder, "" & i), idx, cnt)) Then
                res = False
                Exit For
            End If
            
            cnt = cnt + 1
        Next i
    End If
    
    every = res

End Function

Public Function some(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True) As Boolean
    Dim res As Boolean
    res = False
    
    Dim i
    
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            If Application.Evaluate(Replace(Replace(judgement, placeholder, Replace("" & i, ",", ".")), idx, cnt)) Then
                res = True
                Exit For
            End If
            
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            If Application.Evaluate(Replace(Replace(judgement, placeholder, "" & i), idx, cnt)) Then
                res = True
                Exit For
            End If
        Next i
        
        cnt = cnt + 1
    End If
    
    some = res

End Function

''''''''''''
'@param     operation:              string to be evaluated, e.g. _*2 will be interpreated as ele * 2
'           placeholder:            placeholder to be replaced by the value
'           idx:                    index of the element
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
''''''''''''
Public Function map(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0) As Lists
    
    Dim res As New Lists
    
    
    Dim i
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            i = IIf(IsEmpty(i), setNullValTo, i)
            res.add (Application.Evaluate(Replace(Replace(operation, placeholder, Replace("" & i, ",", ".")), idx, cnt & "")))
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            i = IIf(IsEmpty(i), setNullValTo, i)
            res.add (Application.Evaluate(Replace(Replace(operation, placeholder, "" & i), idx, cnt & "")))
            cnt = cnt + 1
        Next i
    End If
    
    Set map = res
End Function

' first reduce then map

Public Function mapList(ByVal operation As String, ByVal reduceOp As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal initialVal = 1, Optional ByVal placeholderInitialVal As String = "?") As Lists
    Dim res As New Lists
    Dim tmp As New Lists
    
    Dim i
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            i = tmp.addAll(i).reduce(reduceOp, initialVal, placeholder, placeholderInitialVal, idx, replaceDecimalPoint)
            res.add (Application.Evaluate(Replace(Replace(operation, placeholder, Replace("" & i, ",", ".")), idx, cnt & "")))
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            i = tmp.addAll(i).reduce(reduceOp, initialVal, placeholder, placeholderInitialVal, idx, replaceDecimalPoint)
            res.add (Application.Evaluate(Replace(Replace(operation, placeholder, "" & i), idx, cnt & "")))
            cnt = cnt + 1
        Next i
    End If
    
    Set mapList = res
    Set tmp = Nothing
   
End Function

''''''''''''
'@param     judgement:              string to be evaluated and return Boolean, e.g. _>2 will be interpreated as ele > 2
'           placeholder:            placeholder to be replaced by the value
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
''''''''''''
Public Function filter(ByVal judgement As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0) As Lists
    Dim res As New Lists
    
    Dim i
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            i = IIf(IsEmpty(i), setNullValTo, i)
            If Application.Evaluate(Replace(Replace(judgement, placeholder, Replace("" & i, ",", ".")), idx, cnt)) Then
                res.add i
            End If
            
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            i = IIf(IsEmpty(i), setNullValTo, i)
            If Application.Evaluate(Replace(Replace(judgement, placeholder, "" & i), idx, cnt)) Then
                res.add i
            End If
            
            cnt = cnt + 1
        Next i
    End If
    
    Set filter = res
End Function


Public Function take(ByVal n As Long) As Lists
    n = IIf(n >= 0, n, Me.length + n)
    Set take = Me.filter("{i}<" & n)
End Function

Public Function drop(ByVal n As Long) As Lists
    n = IIf(n >= 0, n, Me.length + n)
    Set drop = Me.filter("{i}>=" & n)
End Function

Public Function filterWith(arr As Variant) As Lists
    Dim i
    Dim cnt As Long
    cnt = 0
    
    Dim res As New Lists
    
    For Each i In xToArray(arr)
        If i Then
            res.add Me.getVal(cnt)
        End If
        
        cnt = cnt + 1
        If cnt = Me.length Then
            Exit For
        End If
    Next i
    
    Set filterWith = res
End Function

Public Function filterReg(ByRef reg As Object) As Lists

    Set filterReg = Me.filterWith(Me.judgeReg(reg))
    
End Function

'@desc map to boolean array based on reg
Public Function judgeReg(ByVal reg As Object) As Lists
    Dim i As Long
    Dim res As New Lists
    
    For i = 0 To Me.length - 1
        res.add reg.test(Me.getVal(i))
    Next i
    
    Set judgeReg = res
    Set res = Nothing
End Function

'@desc RegExp should contain at least one group, if matched the first group will be mapped to the list,  otherwise keep the original value
Public Function mapReg(ByVal reg As Object) As Lists
    Dim i As Long
    Dim res As New Lists
    
    For i = 0 To Me.length - 1
        If reg.test(Me.getVal(i)) Then
            res.add reg.Execute(Me.getVal(i))(0).submatches(0)
        Else
            res.add Me.getVal(i)
        End If
    Next i
    
    Set mapReg = res
    Set res = Nothing
End Function

Public Function nullVal(Optional setValTo As Variant) As Lists
    Dim res As New Lists
    
    Dim i
    
    'setValTo missing, left out empty value
    If IsMissing(setValTo) Then
        For Each i In Me.toArray
            If Not IsEmpty(i) Then
                res.add i
            End If
        Next i
    Else
        For Each i In Me.toArray
            res.add IIf(IsEmpty(i), setValTo, i)
        Next i
    End If

    Set nullVal = res
End Function

Public Function reduce(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True) As Variant
    Dim res
    Dim i
    
    res = initialVal
    
    Dim cnt As Long
    cnt = 0
    
    If replaceDecimalPoint Then
        For Each i In Me.toArray
            res = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, Replace("" & i, ",", ".")), placeholderInitialVal, Replace("" & res, ",", ".")), idx, cnt))
            cnt = cnt + 1
        Next i
    Else
        For Each i In Me.toArray
            res = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, "" & i), placeholderInitialVal, "" & res), idx, cnt))
            cnt = cnt + 1
        Next i
    End If
    
     reduce = res
End Function

Public Function reduceRight(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True) As Variant
     reduceRight = Me.reverse().reduce(operation, initialVal, placeholder, placeholderInitialVal, idx, replaceDecimalPoint)
End Function

Public Function product(ByVal operation As String, ByRef list2 As Lists, Optional ByVal placeholder As String = "_", Optional ByVal placeholderOther As String = "?", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    Dim res As New Lists
    
    
    Dim i As Integer
    
    Dim cnt As Long
    cnt = 0
    
    Dim targLen As Integer
    targLen = Application.WorksheetFunction.min(pLen, list2.length) - 1
    
    If replaceDecimalPoint Then
        For i = 0 To targLen
            res.add Application.Evaluate(Replace(Replace(Replace(operation, placeholder, Replace("" & pArr(i), ",", ".")), placeholderOther, Replace("" & list2.getVal(i), ",", ".")), idx, cnt))
            cnt = cnt + 1
        Next i
    Else
        For i = 0 To targLen
            res.add Application.Evaluate(Replace(Replace(Replace(operation, placeholder, "" & pArr(i)), placeholderOther, "" & list2.getVal(i)), idx, cnt))
            cnt = cnt + 1
        Next i
    End If
    
     Set product = res
End Function

Public Function slice(Optional ByVal fromIndex, Optional ByVal toIndex, Optional ByVal step) As Lists

    Dim res As New Lists
    
    
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

Private Function xToArray(x As Variant) As Variant
    If IsArray(x) Then
        xToArray = x
    ElseIf isInstance(x, "Lists") Then
        xToArray = x.toArray
    Else
        xToArray = Array(x)
    End If
End Function

Public Function toArray() As Variant
    Dim arr()
    
    If pLen > 0 Then
        ReDim arr(0 To pLen - 1)
        Dim i As Integer
     
        For i = 0 To pLen - 1
            If Not isObj(pArr(i)) Then
                arr(i) = pArr(i)
            Else
                If Me.isLists(pArr(i)) Then
                    arr(i) = pArr(i).toArray()
                Else
                    Set arr(i) = pArr(i)
                End If
            End If
        Next i
    Else
        arr = Array()
    End If
    
    toArray = arr

End Function

Public Function toString()

    If pLen = 0 Then
        toString = "[]"
    Else
        Dim res As String
        res = "["

        Dim i As Integer

       
        For i = 0 To pLen - 1
            If IsArray(pArr(i)) Then
                Dim t As New Lists
                t.init
                res = t.addAll(pArr(i)).toString
            ElseIf Not isObj(pArr(i)) Then
                res = res & pArr(i) & ", "
            Else
                res = res & pArr(i).toString() & ", "
            End If
        Next i

        toString = Left(res, Len(res) - 2) & "]"
    End If
   
End Function

Public Function sort(Optional ByVal isAscending As Boolean = True) As Lists
    Dim res As New Lists
    
    Dim arr
    
    arr = Me.toArray
    Call QuickSort(arr, 0, pLen - 1)
    res.addAll arr
    
    If isAscending Then
        Set res = res.reverse
    End If
    
    Call override(res)
    
    Set sort = Me

End Function

Public Function reverse() As Lists
    
    Set reverse = Me.slice(, , -1)
    
End Function

Public Function p()
    Debug.Print Me.toString
End Function

Public Function unique() As Lists
    Dim tmp As Object
    Set tmp = CreateObject("scripting.dictionary")
    tmp.compareMode = vbTextCompare
    
    Dim k
    
    For Each k In Me.toArray
        tmp(k) = 1
    Next k
    
    Me.clear
    Set unique = Me.addAll(tmp.keys)
    
    Set tmp = Nothing
End Function

Public Function copy() As Lists
    Dim res As New Lists
    
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
