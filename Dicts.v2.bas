 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class Dicts
'@author                                   Qiou Yang
'@lastUpdate                               27.08.2018
'                                          code refactor
'                                          integrate load/reduce/map/filter into single function
'                                          new feature: load horizontally: set isVertical to false
'@TODO                                     add comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declaration compulsory
Option Explicit

'___________private variables_____________
'scripting.Dictionary Object
Private pDict As Object

' heirachry key level
Private pLevel As Long

' has column label
Private pIsNamed As Boolean

' column label as Dicts, label -> index
Private pNamedArray As Dicts

' target workbook
Private pWb As Workbook

Private pList As Lists

' enum for the parameters in filter/reduce/map
Enum ProcessWith
    Key = 0
    Value = 1
    RangedValue = 2
End Enum

' aggregate method for the function ranged
Enum AggregateMethod
    AggMap = 0
    AggReduce = 1
    Aggfilter = 2
End Enum

' unified property of Dicts/Lists/Nodes
Public Property Get sign() As String
    sign = "Dicts"
End Property


Private Sub Class_Initialize()
    ini
    Set pList = New Lists
End Sub

Private Sub Class_Terminate()
    Set pWb = Nothing
    Set pDict = Nothing
    Set pNamedArray = Nothing
    Set pList = Nothing
End Sub

' get/set target workbook
Public Property Get wb() As Workbook
    Set wb = pWb
End Property

Public Property Let wb(ByRef wkb As Workbook)
   Set pWb = wkb
End Property


' get the underlying Dicitionary-Object
Public Property Get dict() As Object
    Set dict = pDict
End Property

' get/set column labels
Public Property Get named() As Dicts
    If pIsNamed Then
        Set named = pNamedArray
    Else
        Set named = Nothing
    End If
End Property

Public Property Let named(ByVal rng As Variant)
    setNamed rng
End Property

'''''''''''
'@desc:     set the column/row labels to the underlying Dicts
'@return:   this Dicts
'@param:    rng either as Dicts or as Range
'''''''''''
Public Function setNamed(ByVal rng As Variant) As Dicts
   
   On Error GoTo namedArrayHdl
   
   Dim s As String
   Dim c
   Dim cnt As Long
   
   cnt = 0
   
   Dim d As New Dicts
   
   ' test if rng is a Range-Object
   s = rng.Address
   
namedArrayHdl:

    ' if rng is a Range-Object
    If Err.Number = 0 Then
        For Each c In rng.Cells
            d.dict(Trim(CStr(c.Value))) = cnt
            cnt = cnt + 1
        Next c
        
        Me.setNamed d
    Else
        'if rng is a Dicts-Object
        Set pNamedArray = rng
    End If
    
   pIsNamed = True
   
   Set setNamed = Me
   
End Function

' get length of the key-value pairs
Public Property Get count() As Long
    count = pDict.count
End Property

' get keys as Array, if no element return null-Array
Public Property Get keysArr() As Variant
    
    Dim res() As String
    
    If Me.count > 0 Then
        ReDim res(0 To Me.count - 1)
        
        Dim k
        Dim cnt As Long
        cnt = 0
        
        For Each k In Me.keys
            res(cnt) = CStr(k)
            cnt = cnt + 1
        Next k
    End If
    
    keysArr = res
    
End Property

' get keys as Array, if no element return null-Array
Public Property Get valsArr() As Variant
    
    Dim res()
    
    If Me.count > 0 Then
        ReDim res(0 To Me.count - 1)
        
        Dim k
        Dim cnt As Long
        cnt = 0
        
        For Each k In Me.keys
            res(cnt) = Me.dict(k)
            cnt = cnt + 1
        Next k
    End If
    
    valsArr = res
    
End Property

' get keys as iterable-object
Public Property Get keys() As Variant
    keys = pDict.keys
End Property

' set underlying scripting.Dictionary Object
Public Property Let dict(ByRef dict As Object)
    Set pDict = dict
End Property

' initiate the Dictionary-Object
Private Sub ini()
    
    On Error GoTo Errhandler1
    
    Dim a As Long
    a = pDict.count
    
Errhandler1:
    ' if not yet initiated, set pDict
    If Err.Number <> 0 Then
        Set pDict = CreateObject("scripting.dictionary")
        pDict.compareMode = vbTextCompare
    End If
    
    Set pWb = ThisWorkbook
    
End Sub

'''''''''''''''''''''''''''
'@desc:     get Worksheet
'@return:   the target sht
'@param:    targSht         sheet name in string, by default the activesheet
'           wb              the workbook which contains the targSht
'''''''''''''''''''''''''''
Function getTargetSht(Optional ByVal targSht As String = "", Optional ByRef wb As Workbook) As Worksheet
    
    Dim tmpWb As Workbook
    Set tmpWb = IIf(wb Is Nothing, pWb, wb)
    
    With tmpWb
        Dim tmpname As String
        
        tmpname = ActiveSheet.Name
        If Trim(targSht) = "" Then
            targSht = tmpname
        End If
        
       Set getTargetSht = .Worksheets(targSht)
    End With
    
    Set tmpWb = Nothing
    
End Function
    

'''''''''''''''''''''''''''
'@desc:     load the content of range
'@return:   the target range
'@param:    targSht         sheet name in string, by default the activesheet
'           targKeyCol      target key column, default to be 1
'           targValCol      target value column, the column to be read from, default to be the key column
'           targRowBegine   row number to begin
'           targRowEnd      row number ends, by default the last none-empty row of key column
'           isVertical      if true, data entries ranged vertically, i.e. model vlookup;  if false,  targKeyCol means actually targKeyRow and targValCol targValRow
'''''''''''''''''''''''''''
Function getRange(Optional ByVal targSht As String = "", Optional ByVal targKeyCol As Long = 1, Optional ByVal targValCol = 1, Optional targRowBegine As Variant, Optional ByVal targRowEnd As Variant, Optional ByRef wb As Workbook, Optional ByVal isVertical As Boolean = True) As Range
    
    ' if the targValCol is single number, put it into array
    If Not IsArray(targValCol) Then
        targValCol = Array(targValCol)
    End If
    
    ' get the target Range
    With getTargetSht(targSht, wb)
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        If IsMissing(targRowEnd) Then
            If isVertical Then
                targRowEnd = .Cells(.Rows.count, targKeyCol).End(xlUp).row
            Else
                targRowEnd = .Cells(targKeyCol, .Columns.count).End(xlToLeft).Column
            End If
        End If
        
        If isVertical Then
            Set getRange = .Range(Cells(targRowBegine, targValCol(LBound(targValCol))), Cells(targRowEnd, targValCol(UBound(targValCol))))
        Else
            Set getRange = .Range(Cells(targValCol(LBound(targValCol)), targRowBegine), Cells(targValCol(UBound(targValCol)), targRowEnd))
        End If
    End With
    
End Function

'''''''''''
'@desc:     get one-dimensional array based on the range
'@return:   one-dimensional array containing value or address
'@param:    rng             as target Range
'           isVertical      if true, data entries ranged vertically, i.e. model vlookup
'           asAddress       keep the address as the content of the array
'''''''''''
Public Function rngToArr(ByRef rng As Range, Optional ByVal isVertical As Boolean = True, Optional ByVal asAddress As Boolean = False) As Variant
    
    Dim i
    Dim res()
    Dim cnt As Long
    cnt = 0
    Dim arr()   ' multi-dimensional array containing either value or address
    
    ' fill in the arr
    If rng.Cells.count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = IIf(asAddress, rng.Address, rng.Value)
    Else
        If asAddress Then
            arr = rngToAddress(rng)
        Else
            arr = rng.Value
        End If
    End If
    
    ' slice the 2-dimensional array based on the direction specified
    If isVertical Then
        ReDim res(0 To rng.Rows.count - 1)
        For i = LBound(arr, 1) To UBound(arr, 1)
            res(cnt) = sliceArr(arr, i, isVertical)
            cnt = cnt + 1
        Next i
    Else
        ReDim res(0 To rng.Columns.count - 1)
        For i = LBound(arr, 2) To UBound(arr, 2)
            res(cnt) = sliceArr(arr, i, isVertical)
            cnt = cnt + 1
        Next i
    End If
    
    ' if the result array contains only one element, return the result
    If UBound(res) = LBound(res) Then
        rngToArr = res(0)
    Else
        rngToArr = res
    End If
    
End Function


'''''''''''
'@desc:     get two-dimensional array with the address of the target range
'@return:   two-dimensional array with the address of the target range
'@param:    rng             as target Range
'''''''''''
Public Function rngToAddress(ByRef rng As Range) As Variant
    
    Dim fst As Range
    Set fst = rng.Cells(1, 1)
    
    Dim lst As Range
    Set lst = fst.Offset(rng.Rows.count - 1, rng.Columns.count - 1)
        
    Dim i As Long
    Dim j As Long
    
    Dim res()
    ReDim res(1 To rng.Rows.count, 1 To rng.Columns.count)
   
    For i = fst.row To lst.row
        For j = fst.Column To lst.Column
            res(i - fst.row + 1, j - fst.Column + 1) = Cells(i, j).Address
        Next j
    Next i
    
    rngToAddress = res

End Function

'''''''''''
'@desc:     slice two-dimensional array into one-dimensional array based on the direction specified
'@return:   one-dimensional array containing values of specific row or column
'@param:    arr             two-dimensional array
'           n               the n-th row, if isVertical else the n-th column
'           isVertical      if true, data entries ranged vertically, i.e. model vlookup
'''''''''''
Private Function sliceArr(arr, ByVal n As Long, Optional ByVal isVertical As Boolean = True) As Variant
    
    Dim i
    Dim res
    Dim cnt As Long
    cnt = 0
    
    If isVertical Then
        ReDim res(0 To UBound(arr, 2) - LBound(arr, 2))
        ' n is row number, dimension 1
        For i = LBound(arr, 2) To UBound(arr, 2)
            res(cnt) = arr(n, i)
            cnt = cnt + 1
        Next i
    Else
         ReDim res(0 To UBound(arr, 1) - LBound(arr, 1))
        ' n is col number, dimension 2
        For i = LBound(arr, 1) To UBound(arr, 1)
            res(cnt) = arr(i, n)
            cnt = cnt + 1
        Next i
    End If
    
    sliceArr = res
    
End Function

'''''''''''
'@return:   dimension of the array
'@param:    arr             target array
'''''''''''
Private Function arrDimension(arr) As Long

    On Error GoTo hdl:
    Dim res As Long
    res = 0
    Dim cnt As Long
    cnt = 1
    
    If IsArray(arr) Then
        Dim e
        
        Do While True
            e = UBound(arr, cnt)
            cnt = cnt + 1
        Loop
hdl:
        res = cnt - 1
    End If
    
    arrDimension = res
    
End Function


'''''''''''
'@return:   length of the one-dimensional array
'@param:    arr             target array
'''''''''''
Private Function arrLen(arr) As Long
    
    arrLen = UBound(arr) - LBound(arr) + 1

End Function


'''''''''''
'@return:   Dicts obj
'@param:    keyArr             keys in one-dimensional array
'           valArr             vals in one-dimensional array, which can contain arrays of value as its element
'           isReversed         read from bottom up if true
'           keyCstr            whether transfer the keys into trimmed string
'''''''''''
Function arrToDict(keyArr, valArr, Optional ByVal isReversed As Boolean = False, Optional ByVal keyCstr As Boolean = False) As Object
    
    ' combine the key-value pair in a zipped mode
    If arrLen(keyArr) <> arrLen(valArr) Then
        Err.Raise 8888, "", "Arrays with different length can not be combined"
    End If
    
    Dim res As Object
    Set res = CreateObject("scripting.dictionary")
    res.compareMode = vbTextCompare
    
    Dim i
    
    If isReversed Then
        For i = UBound(keyArr) To LBound(keyArr) Step -1
            If Len(Trim(CStr(keyArr(i)))) > 0 Then
                res(IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i))) = valArr(UBound(valArr) + i - UBound(keyArr))
            End If
        Next i
    Else
        For i = LBound(keyArr) To UBound(keyArr)
            If Len(Trim(CStr(keyArr(i)))) > 0 Then
                res(IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i))) = valArr(LBound(valArr) + i - LBound(keyArr))
            End If
        Next i
    End If
    
    Set arrToDict = res
    Set res = Nothing
    
End Function

'''''''''''
'@desc:     map the range to dictionary
'@return:   dictionary-object
'@param:    arr             two-dimensional array
'           n               the n-th row, if isVertical else the n-th column
'''''''''''
Function rngToDict(ByRef keyRng As Range, ByRef valRng As Range, Optional ByVal isReversed As Boolean = False, Optional ByVal asAddress As Boolean = False) As Object
    
    ' if the keyRng contains only one column, it is vertical
    Dim isVertical As Boolean
    isVertical = keyRng.Columns.count = 1
    
    Set rngToDict = arrToDict(rngToArr(keyRng, Not isVertical), IIf(IIf(isVertical, valRng.Columns.count, valRng.Rows.count) = 1, rngToArr(valRng, Not isVertical, asAddress), rngToArr(valRng, isVertical, asAddress)), isReversed)
End Function
 
' to add the shtName just through dict.productX("""'src'!{*}""").p
Public Function load(Optional ByVal Sht As String = "", Optional ByVal KeyCol As Long = 1, Optional ByVal ValCol = 1, Optional RowBegine As Variant = 1, Optional ByVal RowEnd As Variant, Optional ByVal reg As Variant, Optional ByVal ignoreNullVal As Boolean, Optional ByVal setNullValTo As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False, Optional ByRef asAddress As Boolean = False, Optional appendMode As Boolean = False, Optional ByVal isVertical As Boolean = True) As Dicts
    Dim keyRng As Range
    Set keyRng = getRange(Sht, KeyCol, KeyCol, RowBegine, RowEnd, wb, isVertical)
    
    Dim valRng As Range
    Set valRng = getRange(Sht, KeyCol, ValCol, RowBegine, RowEnd, wb, isVertical)
    
    If pDict.count = 0 Or Not appendMode Then
        Set pDict = rngToDict(keyRng, valRng, Reversed, asAddress)
    Else
       Set pDict = Me.union(createInstance(Me.rngToDict(keyRng, valRng, Reversed, asAddress))).dict
    End If
    
    Set load = Me
    
    Set keyRng = Nothing
    Set valRng = Nothing
End Function

Public Function of(ByRef dictObj As Object) As Dicts
    Set pDict = dictObj
    Set of = Me
End Function

Public Function createInstance(ByRef dictObj As Object) As Dicts
    Dim res As New Dicts
    Set createInstance = res.of(dictObj)
    Set res = Nothing

End Function

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
            targRowEnd = .Cells(Rows.count, targKeyCol2).End(xlUp).row
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

End Sub

' rng can be Range Object or an array
Public Function frequencyCount(ByRef rng) As Dicts

    Dim res As New Dicts
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
    Set res = Nothing
End Function

Public Sub unload(ByVal shtName As String, ByVal KeyCol As Long, ByVal startingRow As Long, ByVal startingCol As Long, Optional ByVal endRow As Long, Optional ByVal endCol As Long)

    Dim tmpname As String
    tmpname = ActiveSheet.Name
    
    If Trim(shtName) = "" Then
        shtName = tmpname
    End If

    With Worksheets(shtName)
       If IsMissing(endRow) Or endRow = 0 Then
           endRow = .Cells(Rows.count, KeyCol).End(xlUp).row
       End If
       
       Dim c

       If IsMissing(endCol) Or endCol = 0 Then
           For Each c In .Cells(startingRow, KeyCol).Resize(endRow - startingRow + 1, 1).Cells
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
           
           For Each c In .Cells(startingRow, KeyCol).Resize(endRow - startingRow + 1, 1).Cells
               If pDict.exists(Trim(CStr(c.Value))) Then
                   .Cells(c.row, startingCol).Resize(1, tmpC) = pDict(Trim(CStr(c.Value)))
               End If
           Next c
       
       End If
       
    End With

End Sub


Public Sub dump(ByVal shtName As String, Optional ByVal KeyCol As Long = 1, Optional ByVal startingRow As Long = 1, Optional ByVal startingCol As Long, Optional ByVal endCol As Long)

    If IsMissing(startingCol) Or startingCol = 0 Then
        startingCol = KeyCol + 1
    End If
    
    If shtName = "" Then
        shtName = ActiveSheet.Name
    End If
    
    'unload the key
    Worksheets(shtName).Cells(startingRow, KeyCol).Resize(Me.count, 1) = Application.WorksheetFunction.Transpose(Me.keysArr)
    
    Call Me.unload(shtName, KeyCol, startingRow, startingCol, , endCol)

End Sub

Public Function exists(ByVal k) As Boolean
    
    exists = pDict.exists(Trim(CStr(k)))
    
End Function

' 1 param get the item
' 2 params set the value to the key
Public Function item(ByVal k, Optional v) As Variant
    
    If IsMissing(v) Then
        If IsObject(pDict(Trim(CStr(k)))) Then
            Set item = pDict(Trim(CStr(k)))
        Else
            item = pDict(Trim(CStr(k)))
        End If
    Else
        If IsObject(v) Then
            Set pDict(Trim(CStr(k))) = v
        Else
            pDict(Trim(CStr(k))) = v
        End If
    End If

End Function

Public Function clear()
    pDict.RemoveAll
End Function


' if delete if all the elements are empty
' if value specified, set all empty value in the range to the value
Public Function nulls(Optional ByVal toVal, Optional isRanged As Boolean = False) As Dicts
    
    Dim k
    
    If Not isRanged Then
        isRanged = isRanged_(Me)
    End If
    
    If Not isRanged Then
        If IsMissing(toVal) Then
             For Each k In Me.keys
                If IsEmpty(Me.dict(k)) Then
                    Me.dict.remove k
                End If
            Next k
        Else
            For Each k In Me.keys
                If IsEmpty(Me.dict(k)) Then
                    Me.dict(k) = toVal
                End If
            Next k
        End If
    Else
        Dim l As New Lists
        If IsMissing(toVal) Then
            For Each k In Me.keys
                If l.addAll(Me.dict(k), False).isEmptyList Then
                    Me.dict.remove k
                End If
            Next k
        Else
            For Each k In Me.keys
                Me.dict(k) = l.addAll(Me.dict(k), False).nullVal(toVal).toArray
            Next k
        End If
    End If
    
    Set nulls = Me

End Function

Private Function isRanged_(ByRef obj As Dicts) As Boolean
    
    Dim k
    For Each k In obj.keys
        isRanged_ = IsArray(obj.dict(k))
        Exit For
    Next k
    
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


Public Function sliceWithName(ByVal nm As String) As Dicts
    If pIsNamed Then
        Dim i As Long
        i = pNamedArray.item(nm)
        
        Set sliceWithName = Me.filterRngX("{i}=" & i)
    Else
        Set sliceWithName = Nothing
    End If
End Function


' ________________________________________Class Collection Functions___________________________________________
Public Function diff(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    
    For Each k In pDict.keys
        If Not dict2.dict.exists(k) Then
            res.dict(k) = pDict(k)
        End If
    Next k
    
    Set diff = res
    Set res = Nothing
End Function

'@desc      get the union of two dicts
'@params
Public Function union(dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts
    Dim k
    
    Dim res As New Dicts
    res.dict = pDict
    
    For Each k In dict2.dict.keys
        If Not pDict.exists(k) Then
            res.dict(k) = dict2.dict(k)
        ElseIf Not keepOriginalVal Then
            res.dict(k) = dict2.dict(k)
        End If
    Next k
    
    Set union = res
    Set res = Nothing
End Function

Public Function intersect(dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts
    Dim k
    
    Dim res As New Dicts
    res.dict = pDict
    
    For Each k In dict2.dict.keys
        If pDict.exists(k) Then
            If Not keepOriginalVal Then
                res.dict(k) = dict2.dict(k)
            Else
                res.dict(k) = pDict.dict(k)
            End If
        End If
    Next k
    
    Set intersect = res
    Set res = Nothing
End Function


Public Function update(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As New Dicts
    
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

Public Function reduce(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal reduceWith As Long = ProcessWith.Value) As Variant
     Dim l As New Lists
     
     If Len(reduceValRangeOp) > 0 Then
        reduceWith = ProcessWith.RangedValue
     End If
     
     If reduceWith = ProcessWith.Value Then
        reduce = l.addAll(Me.valsArr).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
     ElseIf reduceWith = ProcessWith.Key Then
        reduce = l.addAll(Me.keysArr).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
     ElseIf reduceWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set l = Nothing
End Function

''''''''''''
'@param     operation:              string to be evaluated, e.g. _*2 will be interpreated as ele * 2
'           placeholder:            placeholder to be replaced by the value
'           idx:                    index of the element
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
'@example   get length of the valsArray -> ProcessWith.Value
''''''''''''
Public Function map(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal mapWith As Long = ProcessWith.Value) As Dicts
     Dim l As New Lists
     
     If mapWith = ProcessWith.Value Then
        Set map = Me.updateFromArray(l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray, mapWith)
     ElseIf mapWith = ProcessWith.Key Then
        Set map = Me.updateFromArray(l.addAll(Me.keysArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray, mapWith)
     ElseIf mapWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set l = Nothing
End Function


Public Function ranged(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal initialVal As Variant = 1, Optional ByVal aggregate As Long = AggregateMethod.AggReduce) As Dicts
    
    Dim k
    Dim res As New Dicts
    Dim l As New Lists
    
    If aggregate = AggregateMethod.AggReduce Then
        For Each k In Me.keys
            res.dict(k) = l.addAll(Me.dict(k), False).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
        Next k
    ElseIf aggregate = AggregateMethod.AggMap Then
         For Each k In Me.keys
            res.dict(k) = l.addAll(Me.dict(k), False).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray
        Next k
    ElseIf aggregate = AggregateMethod.Aggfilter Then
        For Each k In Me.keys
            res.dict(k) = l.addAll(Me.dict(k), False).filter(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray
        Next k
    Else
        Err.Raise 8889, , "unknown aggregate parameter"
    End If
    
    Set ranged = res
    Set res = Nothing
    Set l = Nothing
End Function


Public Function filter(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo As Variant = 0, Optional ByVal filterWith As Long = ProcessWith.Value) As Dicts
     Dim l As New Lists
     Dim tmp As New Lists
     
     If filterWith = ProcessWith.Value Then
        ' map to true or false
        Set tmp = l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        Set pDict = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
     ElseIf filterWith = ProcessWith.Key Then
        Set tmp = l.addAll(Me.keysArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        Set pDict = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
     ElseIf filterWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set filter = Me
     Set l = Nothing
     Set tmp = Nothing
End Function

Public Function updateFromArray(ByVal arr, Optional ByVal updateWith As Long = ProcessWith.Value) As Dicts
    Dim keyArr
    Dim valArr
    
    keyArr = Me.keysArr
    valArr = Me.valsArr
    
    If arrLen(arr) <> arrLen(keyArr) Then
        Err.Raise 8888, , "Input Array should be the same length with the Dict"
    End If
    
    Dim l As New Lists
    
    If updateWith = ProcessWith.Key Then
        Set pDict = arrToDict(arr, valArr, , False)
    Else
        Set pDict = arrToDict(keyArr, arr, , False)
    End If
    
    Set updateFromArray = Me
    keyArr = Array()
    valArr = Array()
    Set l = Nothing

End Function




Public Function reduceRngVertical(ByVal sign As String) As Variant
    Dim k
    Dim i
    Dim tmpCnt As Long
    tmpCnt = 1
    Dim arr()
    
    Dim u As Long
    Dim l As Long

    For Each k In pDict.keys
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

Private Function ifEmpty(ByVal targetVal As Variant, ByVal valIfNull As Variant) As Variant
    
   ifEmpty = IIf(IsEmpty(targetVal), valIfNull, targetVal)

End Function


Private Function reduceArray(ByVal arr, ByVal sign As String, Optional ByVal valIfNull As Variant = 0) As Variant
    Dim res As Variant
    Dim k
    
    
    If sign = "" Or sign = "+" Then
        res = 0
        For Each k In arr
            res = res + ifEmpty(k, valIfNull)
        Next k
    ElseIf sign = "*" Then
        res = 1
        For Each k In arr
            res = res * ifEmpty(k, valIfNull)
        Next k
    End If
    
    reduceArray = res
    
End Function

'''''''''''''''''''''''''''
'@desc:     reduceArrayX -> reduce the array as value through the operation defined
'           ref. reduceRngX
'@param:    arr             array to be reduced
'           operation       operation to be performed on the array, e.g. get the sum of array "{v}+{*}"
'           initVal         the inital value of the reduction, e.g. get the sum of array 0
'           placeholder     placeholder of the value
'           index           placeholder of the index, starting from 0
'           cumVal          the accumlator
'           hasThousandSep  relevant for "." as thousand sep
'           valIfNull       set value if the array position is null
'''''''''''''''''''''''''''
Private Function reduceArrayX(ByVal arr, ByVal operation As String, Optional ByVal initVal As Variant = 0, Optional ByVal placeholder As String = "{*}", Optional ByVal index As String = "{i}", Optional ByVal cumVal As String = "{v}", Optional ByVal hasThousandSep As Boolean = True, Optional ByVal valIfNull As Variant = 0) As Variant
    Dim k
    Dim v
    Dim tmp As String
    
    If hasThousandSep Then
        For k = LBound(arr) To UBound(arr)
            tmp = Replace(ifEmpty(arr(k), valIfNull) & "", ",", ".")
            initVal = Replace(initVal & "", ",", ".")
            initVal = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, tmp), index, k), cumVal, initVal))
        Next k
    Else
        For k = LBound(arr) To UBound(arr)
            initVal = Application.Evaluate(Replace(Replace(Replace(operation, placeholder, ifEmpty(arr(k), valIfNull) & ""), index, k), cumVal, initVal))
        Next k
    End If

    reduceArrayX = initVal
End Function

''''''''''''''''''''
'set all the elements to a constant
'default to be 1
''''''''''''''''''''

Public Function constDict(Optional ByVal constant As Variant = 1) As Dicts
    Dim k
    Dim res As New Dicts
    
    For Each k In pDict.keys
        res.dict(k) = constant
    Next k
    
    Set constDict = res
    Set res = Nothing
End Function


'deep copy of this-Dicts-Object
Public Function clone() As Dicts
    Set clone = clone__(Me, pLevel)
End Function

Private Function clone__(ByVal d As Dicts, ByVal l As Long) As Dicts
    Dim res As New Dicts
    Dim k

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
    Set res = Nothing

End Function

' ______________________________ Print______________________________________________

'print the key=>value pairs of this Dicts
Public Function p()
    Debug.Print Me.toString()
End Function

Public Function toString() As String
    toString = x_toString(Me)
End Function

' print iterables to screen
Private Function a_toString(ByVal arr As Variant, Optional ByVal lvl As Integer = 0) As String
    
    If arrLen(arr) = 0 Then
        a_toString = "[ ]"
    Else
        Dim res As String
        Dim i
        res = "["
        
        For Each i In arr
            If Not IsNumeric(i) Then
                res = res & x_toString(i, lvl + 1) & ", "
            Else
                res = res & Replace(" " & i, ",", ".") & ", "
            End If
        Next i
        
        res = Left(res, Len(res) - 2)

        a_toString = res & " ]"
    End If

End Function

Private Function dicts_toString(d As Variant, Optional ByVal lvl As Integer = 0) As String

    If d.count = 0 Then
        dicts_toString = "{}"
    Else
        Dim res As String
        Dim k
        res = "{" & Chr(10)
        
        For Each k In d.dict.keys
            res = res & String(lvl, Chr(9)) & k & Chr(9) & "=>" & Chr(9) & x_toString(d.dict(k), lvl + 1) & "," & Chr(10)
        Next k
        
        res = Left(res, Len(res) - 2)
        
        dicts_toString = res & Chr(10) & String(lvl, Chr(9)) & "}"
    End If

End Function

Public Function x_toString(x As Variant, Optional ByVal lvl As Integer = 0) As String
        
    If IsArray(x) Then
        x_toString = a_toString(x, lvl)
    ElseIf Me.isDict(x) Then
        x_toString = dicts_toString(x, lvl)
    Else
        If pList.isLists(x) Then
            x_toString = x.toString
        Else
            x_toString = CStr(x)
        End If
    End If

End Function

Public Function pk()

    Dim k
    For Each k In Me.dict.keys
        Debug.Print k
    Next k

End Function

Public Function ps(Optional ByVal lvl As Long = 1, Optional ByVal cnt As Long = 0)
    
    Dim k
    
    If cnt = lvl Then
        For Each k In Me.dict.keys
            Debug.Print String(cnt, Chr(9)) & k & Chr(9) & Me.dict(k)
        Next k
    Else
        For Each k In Me.dict.keys
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
    For Each ky In pDict.keys
        res = res & "{""name"":""" & Replace(CStr(ky), """", "") & """, " & """size"": " & Replace(CStr(pDict(ky)), ",", ".") & "}," & Chr(13)
    Next ky
    
    toJSON = Left(res, Len(res) - 2) & Chr(13) & "]}"
    
End Function

' ________________________________________Util Functions____________________________________________

' return the RegExp-Object
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
    Set obj = Nothing
End Function

' return a consective sequence of the integer numbers
Public Function rng(ByVal start As Long, ByVal ending As Long, Optional ByVal steps As Long = 1)
    Dim res()
    Dim cnt As Long
    cnt = -1
    Dim i As Long
    
    For i = start To ending Step steps
        cnt = cnt + 1
    Next i
    
    ReDim res(0 To cnt)
    
    For i = start To ending Step steps
        res(i - start) = i
    Next i
    
    rng = res
End Function

Public Function y(Optional ByVal Sht As String = "", Optional ByVal col As Long = 1, Optional ByVal wb As String = "") As Long
    
    y = getTargetWorksheet(Sht, wb).Cells(Rows.count, col).End(xlUp).row
    
End Function

Public Function x(Optional ByVal Sht As String = "", Optional ByVal row As Long = 1, Optional ByVal wb As String = "") As Long
    
    x = getTargetWorksheet(Sht, wb).Cells(row, Columns.count).End(xlToLeft).Column
    
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

Public Function ClassHashID() As String
    ClassHashID = "#Dicts_W3I89DWX897HH7NC9"
End Function

Public Function isDict(o As Variant) As Boolean
    On Error GoTo errhandler_d
    
    Dim a As Boolean
    a = (o.ClassHashID = "#Dicts_W3I89DWX897HH7NC9")
    
errhandler_d:
    If Err.Number = 0 Then
        isDict = a
    Else
        isDict = False
    End If

End Function

' is Instance of Dicts, Lists or Nodes
Private Function isObj(ByVal obj) As Boolean
    On Error GoTo listhandler
    
    Dim res As Boolean
    res = False
    
    Dim myType As String
    myType = obj.sign
    
listhandler:
    isObj = (Err.Number = 0)

End Function
