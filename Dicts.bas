 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class Dicts
'@author                                   Qiou Yang
'@license                                  MIT
'@lastUpdate                               07.08.2019
'                                          minor bugfix / load method now does not change the status of  underlying object
'                                          feed/reset method -> feed the value recursively to the struct
'@TODO                                     add comments
'                                          unify the Exception-Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declaration compulsory
Option Explicit

'___________private variables_____________
'implement TreeMaps Object

Private pKeys As TreeSets
Private pVals As Collection
Const pUpdate As Boolean = True ' to update if duplicated

' has column label
Private pIsLabeled As Boolean

' column label as Dicts, label -> index
Private pLabeledArray As Dicts

' target workbook
Private pWb As Workbook

Private pList As Lists

' enum for the parameters in filter/reduce/map
Enum ProcessWith
    key = 0
    value = 1
    RangedValue = 2
End Enum

' aggregate method for the function ranged
Enum AggregateMethod
    AggMap = 0
    AggReduce = 1
    Aggfilter = 2
End Enum

Private Sub Class_Initialize()
    Set pWb = ThisWorkbook
    Set pList = New Lists
    Set pKeys = New TreeSets
    Set pVals = New Collection
End Sub

Private Sub Class_Terminate()
    Set pWb = Nothing

    Set pLabeledArray = Nothing
    Set pList = Nothing
    
    Set pKeys = Nothing
    Set pVals = Nothing
End Sub

' get/set target workbook
Public Property Get wb() As Workbook
    Set wb = pWb
End Property

Public Property Let wb(ByRef wkb As Workbook)
   Set pWb = wkb
End Property

' get the underlying Dicitionary-Object
Public Property Get dict() As Dicts
    Set dict = Me
End Property

Public Function add(k, v) As Dicts
    pKeys.add k, pUpdate
    pVals.add v
    Set add = Me
End Function

Public Property Let Item(key As Variant, value As Variant)
    add key, value
End Property
   
Public Property Get Item(key As Variant) As Variant
    Dim tmp As Nodes
    Set tmp = pKeys.ceiling(key, True)
    
    If tmp Is Nothing Then
        Item = Null
    Else
        If IsObject(pVals.Item(tmp.index + 1)) Then
            Set Item = pVals.Item(tmp.index + 1)
        Else
            Item = pVals.Item(tmp.index + 1)
        End If
    End If
    
    Set tmp = Nothing
End Property

Public Function exists(key As Variant) As Boolean
    Dim tmp As Nodes
    Set tmp = pKeys.ceiling(key, True)
    exists = False
    
    If Not tmp Is Nothing Then
        exists = tmp.value = key
    End If
    
End Function

Public Function RemoveAll()
    pKeys.clear
    Set pVals = New Collection
End Function

Public Function Remove(e)
    pKeys.Remove e
End Function

Public Function clear()
    RemoveAll
End Function

' get/set column labels
Public Property Get label() As Dicts
    If pIsLabeled Then
        Set label = pLabeledArray
    Else
        Set label = Nothing
    End If
End Property

Public Property Let label(ByVal rng As Variant)
    setLabel rng
End Property

Public Function hasLabel() As Boolean
    hasLabel = pIsLabeled
End Function

Public Function copyLabel(ByRef src As Dicts, ByRef targ As Dicts)
    If src.hasLabel Then
        targ.label = src.label
    Else
        targ.label = Nothing
    End If
End Function

'''''''''''
'@desc:     set the column/row labels to the underlying Dicts
'@return:   this Dicts
'@param:    rng either as Array, Dicts or as Range
'''''''''''
Public Function setLabel(ByVal rng As Variant) As Dicts
   
   If isInstanceOf(rng, "Nothing") Then
        pIsLabeled = False
        Set pLabeledArray = Nothing
   Else
        
        Dim c
        Dim cnt As Long
        cnt = 0
        
        Dim d As New Dicts
        
        If isInstanceOf(rng, "Range") Then
             For Each c In rng.Cells
                 d.Item(Trim(CStr(c.value))) = cnt
                 cnt = cnt + 1
             Next c
             
             Me.setLabel d
         Else
             If isDict(rng) Then
                 Set pLabeledArray = rng
             ElseIf IsArray(rng) Then
                 Dim k
                 
                 For k = 0 To UBound(rng) - LBound(rng)
                     d.Item(rng(k)) = k
                 Next k
                 
                 Me.setLabel d
             End If
         End If
         
        pIsLabeled = True
   End If
   
   Set setLabel = Me
   
End Function

'@desc      get element by key and label
'@return    the target element

Public Function getByLabel(ByRef k As Variant, ByRef label As String) As Variant
    
    If Not pIsLabeled Then
        Err.Raise 99760, , "LabelAbsentException: please specify the label first"
    End If
    
    If Not Me.exists(k) Then
        Err.Raise 89760, , "ElementNotFoundException: the key does not exist"
    End If
    
    If Not pLabeledArray.exists(label) Then
        Err.Raise 89760, , "ElementNotFoundException: the label '" & label & "' does not exist"
    End If
    
    If IsObject(Item(k)(pLabeledArray.Item(label))) Then
        Set getByLabel = Item(k)(pLabeledArray.Item(label))
    Else
        getByLabel = Item(k)(pLabeledArray.Item(label))
    End If
    
End Function

' get length of the key-value pairs
' if recursive set to true, count the keys of all child-dicts
' allLevels only relevant in recursive-mode, count all the keys in the structure
Public Function Count(Optional ByVal recursive As Boolean = False, Optional ByVal allLevels As Boolean = True) As Long
    
    If Not recursive Then
        Count = pKeys.size
    Else
        If isDicted_(Me) Then
            Dim k
            Dim res As Long
            
            For Each k In Me.Keys
                res = IIf(allLevels, 1, 0) + res + Me.Item(k).Count(True)
            Next k
            
            Count = res
        Else
            Count = Me.Count
        End If
    End If
End Function

' get keys as Array, if no element return null-Array
Public Property Get keysArr() As Variant
    keysArr = Me.Keys
End Property

' get keys as Array, if no element return null-Array
Public Property Get valsArr() As Variant
    
    Dim res()
    
    If Me.Count > 0 Then
        ReDim res(0 To Me.Count - 1)
        
        Dim k
        Dim cnt As Long
        cnt = 0
        
        For Each k In Me.Keys
            res(cnt) = Me.Item(k)
            cnt = cnt + 1
        Next k
    End If
    
    valsArr = res
    Erase res
End Property

' get keys as iterable-object
Public Property Get Keys() As Variant
    Keys = pKeys.toArray
End Property

Public Function fromArray(ByRef arr) As Dicts
    If IsArray(arr) Then
       Set fromArray = pList.addAll(arr).toDict
    ElseIf TypeName(arr) = "Lists" Then
       Set fromArray = arr.toDict
    Else
        Err.Raise 9876, , "Unknown Parameter Type!"
    End If
End Function

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
    
    ' get the target Range
    With getTargetSht(targSht, wb)
        If IsMissing(targRowBegine) Then
            targRowBegine = 1
        End If
        
        ' if the targValCol is single number, put it into array
        If Not IsArray(targValCol) Then
            targValCol = Array(targValCol)
        End If

        If IsMissing(targRowEnd) Then
            If isVertical Then
                targRowEnd = .Cells(.Rows.Count, targKeyCol).End(xlUp).row
            Else
                targRowEnd = .Cells(targKeyCol, .Columns.Count).End(xlToLeft).Column
            End If
        End If
        
        If isVertical Then
            Set getRange = .Cells(targRowBegine, targValCol(LBound(targValCol))).Resize(targRowEnd - targRowBegine + 1, targValCol(UBound(targValCol)) - targValCol(LBound(targValCol)) + 1)
        Else
            Set getRange = .Cells(targValCol(LBound(targValCol)), targRowBegine).Resize(targValCol(UBound(targValCol)) + 1 - targValCol(LBound(targValCol)), targRowEnd + 1 - targRowBegine)
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
    If rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = IIf(asAddress, rng.Address, rng.value)
    Else
        If asAddress Then
            arr = rngToAddress(rng)
        Else
            arr = rng.value
        End If
    End If
    
    ' slice the 2-dimensional array based on the direction specified
    If isVertical Then
        ReDim res(0 To rng.Rows.Count - 1)
        For i = LBound(arr, 1) To UBound(arr, 1)
            res(cnt) = sliceArr(arr, i, isVertical)
            cnt = cnt + 1
        Next i
    Else
        ReDim res(0 To rng.Columns.Count - 1)
        For i = LBound(arr, 2) To UBound(arr, 2)
            res(cnt) = sliceArr(arr, i, isVertical)
            cnt = cnt + 1
        Next i
    End If
    
    ' if the result array contains only one element and the element is not array itself, return the result
    If UBound(res) = LBound(res) Then
        rngToArr = res(0)
    Else
        rngToArr = res
    End If
    
    Erase res
    Erase arr
    
End Function


'''''''''''
'@desc:     get two-dimensional array with the address of the target range
'@return:   two-dimensional array with the address of the target range
'@param:    rng             as target Range
'''''''''''
Public Function rngToAddress(ByRef rng As Range, Optional ByVal withShtName As Boolean = True, Optional ByVal withWbName As Boolean = False) As Variant
    
    Dim fst As Range
    Set fst = rng.Cells(1, 1)
    
    Dim lst As Range
    Set lst = fst.offSet(rng.Rows.Count - 1, rng.Columns.Count - 1)
    
    Dim shtName As String
    Dim wbName As String
    shtName = "'" & fst.Worksheet.Name & "'!"
    wbName = "'[" & fst.Worksheet.parent.Name & "]" & fst.Worksheet.Name & "'!"
        
    Dim i As Long
    Dim j As Long
    
    Dim res()
    ReDim res(1 To rng.Rows.Count, 1 To rng.Columns.Count)
   
    For i = fst.row To lst.row
        For j = fst.Column To lst.Column
            res(i - fst.row + 1, j - fst.Column + 1) = IIf(withWbName, wbName & Cells(i, j).Address(0, 0), IIf(withShtName, shtName & Cells(i, j).Address(0, 0), Cells(i, j).Address(0, 0)))
        Next j
    Next i
    
    rngToAddress = res
    
    Set fst = Nothing
    Set lst = Nothing
    Erase res

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
    Dim res()
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
    Erase res
    
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
Function arrToDict(keyArr, valArr, Optional ByVal isReversed As Boolean = False, Optional ByVal keyCstr As Boolean = False) As Dicts
    
    Dim res As New Dicts

    ' combine the key-value pair in a zipped mode
    If arrLen(keyArr) = 0 Then
    
    ElseIf arrLen(keyArr) = 1 And arrLen(valArr) > 1 Then
        res.Item(keyArr(LBound(keyArr))) = valArr
    Else
        If arrLen(keyArr) <> arrLen(valArr) Then
            Err.Raise 8888, "", "Arrays with different length can not be combined"
        End If
        
        Dim i, k
        
        If isReversed Then
            For i = UBound(keyArr) To LBound(keyArr) Step -1
                If Len(Trim(CStr(keyArr(i)))) > 0 Then
                    k = IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i))
                    If res.exists(k) Then
                        res.Remove k
                    End If
                    
                    res.add k, valArr(i)
                End If
            Next i
        Else
            For i = LBound(keyArr) To UBound(keyArr)
                If Len(Trim(CStr(keyArr(i)))) > 0 Then
                    k = IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i))
                    If res.exists(k) Then
                        res.Remove k
                    End If
                
                    res.add k, valArr(i)
                End If
            Next i
        End If
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
Function rngToDict(ByRef keyRng As Range, ByRef valRng As Range, Optional ByVal isReversed As Boolean = False, Optional ByVal asAddress As Boolean = False) As Dicts
    
    ' if the keyRng contains only one column, it is vertical
    Dim isVertical As Boolean
    isVertical = keyRng.Columns.Count = 1
    
    Set rngToDict = arrToDict(rngToArr(keyRng, Not isVertical), IIf(IIf(isVertical, valRng.Columns.Count, valRng.Rows.Count) = 1, rngToArr(valRng, Not isVertical, asAddress), rngToArr(valRng, isVertical, asAddress)), isReversed)
End Function

'''''''''''
'@desc:     read data from xlSht to Dicts-Collection
'@return:   Dicts-object
'@param:    Sht             Name of the target Worksheet
'           KeyCol          the position of key,  the n-th column if isVertical, else the n-th row
'           ValCol          the position of value,  can be either a number or an array representing the n-th column if isVertical, else the n-th row
'           RowBegine       the firstRow of the data entries
'           RowEnd          the lastRow of the data entries, by default the last none-empty row
'           wb              the Workbook-Object which contains the Sht, by default thisWorkbook
'           Reversed        read from bottom up if true.
'           asAddress       load the addresses
'           appendMode      keep the old data if evoked multiple times
'           isVertical      vlook if true
'''''''''''
Public Function load(Optional ByVal sht As String = "", Optional ByVal KeyCol As Long = 1, Optional ByVal valCol = 1, Optional RowBegine As Variant = 1, Optional ByVal RowEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False, Optional ByRef asAddress As Boolean = False, Optional appendMode As Boolean = False, Optional ByVal isVertical As Boolean = True) As Dicts
    Dim keyRng As Range
    Set keyRng = getRange(sht, KeyCol, KeyCol, RowBegine, RowEnd, wb, isVertical)
    
    Dim valRng As Range
    Set valRng = getRange(sht, KeyCol, valCol, RowBegine, RowEnd, wb, isVertical)
    
    If Count = 0 Or Not appendMode Then
        Set load = rngToDict(keyRng, valRng, Reversed, asAddress)
    Else
       Set load = union(createInstance(Me.rngToDict(keyRng, valRng, Reversed, asAddress)))
    End If

    Set keyRng = Nothing
    Set valRng = Nothing
    
End Function

Public Function loadH(Optional ByVal sht As String = "", Optional ByVal KeyRow As Long = 1, Optional ByVal ValRow = 1, Optional ColBegine As Variant = 1, Optional ByVal ColEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False, Optional ByRef asAddress As Boolean = False, Optional appendMode As Boolean = False) As Dicts
    Set loadH = load(sht:=sht, KeyCol:=KeyRow, valCol:=ValRow, RowBegine:=ColBegine, RowEnd:=ColEnd, wb:=wb, Reversed:=Reversed, asAddress:=asAddress, appendMode:=appendMode, isVertical:=False)
End Function

'@desc update self with new dictionary obj
'@deprecated only for the legacy code
Public Function of(ByRef dictObj As Dicts) As Dicts
    Set of = dictObj
End Function

'@desc create a new instance with the dictionary obj
'@deprecated only for the legacy code
Public Function createInstance(ByRef dictObj As Dicts) As Dicts
    Dim res As New Dicts
    Set createInstance = res.of(dictObj)
    Set res = Nothing
End Function

'@deprecated only for the legacy code
Public Function emptyInstance() As Dicts
    Dim res As New Dicts
    Set emptyInstance = res
    Set res = Nothing
End Function


Public Function loadStruct(ByVal sht As String, ByVal KeyCol1 As Long, ByVal KeyCol2 As Long, ByVal valCol, Optional RowBegine As Variant, Optional ByVal RowEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False) As Dicts

    With getTargetSht(sht, wb)
    
        Dim dict As New Dicts
        
        If IsMissing(RowBegine) Then
            RowBegine = 1
        End If
        
        If IsMissing(RowEnd) Then
            RowEnd = .Cells(Rows.Count, KeyCol2).End(xlUp).row
        End If
        
        Dim tmpPreviousRow As Long
        Dim tmpCurrentRow As Long
        Dim tmpDict As New Dicts
        
        tmpPreviousRow = RowEnd
        tmpCurrentRow = tmpPreviousRow
        
        Do While tmpCurrentRow > RowBegine
            tmpCurrentRow = .Cells(tmpCurrentRow, KeyCol1).End(xlUp).row
            
            dict.add .Cells(tmpCurrentRow, KeyCol1).value, tmpDict.load(sht, KeyCol2, valCol, tmpCurrentRow + 1, tmpPreviousRow, wb, Reversed)
            Set tmpDict = Nothing
            
            tmpPreviousRow = tmpCurrentRow - 1
        Loop
        
        Set loadStruct = dict
        Set dict = Nothing
    End With

End Function

Public Function reset(Optional ByVal v As Variant = 0) As Dicts
    
    Dim k
    
    For Each k In Me.Keys
       If isDict(Me.Item(k)) Then
            Me.Item(k).reset v
        Else
            Me.Item(k) = v
       End If
    Next k
    
    Set reset = Me
    
End Function

' incremental based on the data Dict feed
Public Function feed(ByRef d As Dicts, Optional ByVal isIncremental As Boolean = False) As Dicts
    
    Dim k
    
    For Each k In Me.Keys
       If isDict(Me.Item(k)) Then
            Me.Item(k).feed d
        Else
            If d.exists(k) Then
                If isIncremental Then
                    Me.Item(k) = Me.Item(k) + d.Item(k)
                Else
                    Me.Item(k) = d.Item(k)
                End If
            End If
       End If
    Next k
    
    Set feed = Me

End Function


' rng can be Range Object or an array
Public Function frequencyCount(ByRef rng) As Dicts

    Dim res As New Dicts
    Dim k

    If TypeName(rng) = "Range" Then
        For Each k In rng.Cells
            If Len(k.value) > 0 Then
                If res.exists(k.value) Then
                    res.Item(k.value) = res.Item(k.value) + 1
                Else
                    res.Item(k.value) = 1
                End If
            End If
        Next k
    Else
         For Each k In rng
            If Len(k) > 0 Then
                If res.exists(k) Then
                    res.Item(k) = res.Item(k) + 1
                Else
                    res.Item(k) = 1
                End If
            End If
        Next k
    Else
        ' type undefined
    End If

    Set frequencyCount = res
    Set res = Nothing
End Function

Public Sub unload(ByVal shtName As String, ByVal keyPos As Long, ByVal startingRow As Long, ByVal startingCol As Long, Optional ByVal endRow As Long, Optional ByVal endCol As Long, Optional ByRef wb As Workbook, Optional ByVal isVertical As Boolean = True)
    
    Dim c
    Dim tmp
    Dim l
    
    With getTargetSht(shtName, wb)
        If isVertical Then
        
            If IsMissing(endRow) Or endRow = 0 Then
               endRow = .Cells(.Rows.Count, keyPos).End(xlUp).row
            End If
            
            For Each c In .Cells(startingRow, keyPos).Resize(endRow - startingRow + 1, 1).Cells
                If exists(c.value) Then
                    tmp = Item(c.value)
                    If IsArray(tmp) Then
                        If IsMissing(endCol) Or endCol = 0 Then
                            .Cells(c.row, startingCol).Resize(1, arrLen(tmp)).value = tmp
                        Else
                             l = pList.fromArray(tmp, False).take(endCol - startingCol + 1).toArray
                            .Cells(c.row, startingCol).Resize(1, arrLen(l)).value = l
                        End If
                    Else
                        .Cells(c.row, startingCol).value = tmp
                    End If
                End If
            Next c
            
        Else
            If IsMissing(endCol) Or endCol = 0 Then
               endCol = .Cells(keyPos, .Columns.Count).End(xlToLeft).Column
            End If
            
            For Each c In .Cells(keyPos, startingCol).Resize(1, endCol - startingCol + 1).Cells
                If exists(c.value) Then
                    tmp = Item(c.value)
                    If IsArray(tmp) Then
                        If IsMissing(endRow) Or endRow = 0 Then
                            .Cells(startingRow, c.Column).Resize(arrLen(tmp), 1).value = Application.WorksheetFunction.Transpose(tmp)
                        Else
                            l = pList.fromArray(tmp, False).take(endCol - startingCol + 1).toArray
                            .Cells(startingRow, c.Column).Resize(arrLen(l), 1).value = Application.WorksheetFunction.Transpose(l)
                        End If
                    Else
                        .Cells(startingRow, c.Column).value = tmp
                    End If
                End If
            Next c
        End If
    End With

End Sub

Public Sub dump(ByVal shtName As String, Optional ByVal keyPos As Long = 1, Optional ByVal startingRow As Long = 1, Optional ByVal startingCol As Long = 2, Optional ByVal endRow As Long, Optional ByVal endCol As Long, Optional ByRef wb As Workbook, Optional ByVal isVertical As Boolean = True, Optional ByVal trailingRows As Long = 0, Optional ByVal withLabel As Boolean = False)

    With getTargetSht(shtName, wb)
        
        If Me.Count > 0 Then
            If isDicted_(Me) Then
                Dim k
                Dim cnt As Long
                
                For Each k In Me.Keys
                    If isVertical Then
                        .Cells(startingRow + cnt, keyPos) = k
                    Else
                        .Cells(keyPos, startingCol + cnt) = k
                    End If
                
                    Me.Item(k).dump shtName, keyPos + 1, startingRow + cnt + 1, startingCol + 1, startingRow + cnt + Me.Item(k).Count(True), endCol, wb, isVertical, trailingRows, withLabel
                    cnt = cnt + Me.Item(k).Count(True) + 1
                Next k
            Else
                 'unload the key
                If isVertical Then
                    .Cells(startingRow, keyPos).Resize(Me.Count, 1) = Application.WorksheetFunction.Transpose(Me.keysArr)
                Else
                    .Cells(keyPos, startingCol).Resize(1, Me.Count) = Me.keysArr
                End If
            
                Me.unload shtName, keyPos, startingRow, startingCol, endRow, endCol, wb, isVertical
            End If
        End If
        
        If withLabel And Me.hasLabel Then
            
            .Rows(startingRow).Insert Shift:=xlDown
            .Range(.Cells(startingRow, startingCol), .Cells(startingRow, startingCol + Me.label.Count - 1)) = Me.label.keysArr
            
        End If
        
    End With
    
    
End Sub

' if delete if all the elements are empty
' if value specified, set all empty value in the range to the value
Public Function nulls(Optional ByVal toVal, Optional isRanged As Boolean = False) As Dicts
    
    Dim k
    
    If Not isRanged Then
        isRanged = isRanged_(Me)
    End If
    
    If Not isRanged Then
        If IsMissing(toVal) Then
             For Each k In Me.Keys
                If isEmpty(Me.Item(k)) Then
                    Me.Remove k
                End If
            Next k
        Else
            For Each k In Me.Keys
                If isEmpty(Me.Item(k)) Then
                    Me.Item(k) = toVal
                End If
            Next k
        End If
    Else
        Dim l As New Lists
        If IsMissing(toVal) Then
            For Each k In Me.Keys
                If l.addAll(Me.Item(k), False).isEmptyList Then
                    Me.Remove k
                End If
            Next k
        Else
            For Each k In Me.Keys
                Me.Item(k) = l.addAll(Me.Item(k), False).nullVal(toVal).toArray
            Next k
        End If
    End If
    
    Set nulls = Me

End Function

' if containing arrays as element
Private Function isRanged_(ByRef obj As Dicts) As Boolean
    
    Dim k
    For Each k In obj.Keys
        isRanged_ = IsArray(obj.Item(k))
        Exit For
    Next k
    
End Function

' if containing Dicts as elements
Private Function isDicted_(ByRef obj As Dicts) As Boolean
    
    Dim k
    For Each k In obj.Keys
        isDicted_ = isDict(obj.Item(k))
        Exit For
    Next k
    
End Function

Public Function sliceWithLabel(nm) As Dicts
    If pIsLabeled Then
        Dim res As Dicts
        
        If TypeName(nm) = "String" Then
            Dim i As Long
            i = pLabeledArray.Item(nm)
            
            Set res = Me.ranged("{i}=" & i, aggregate:=AggregateMethod.Aggfilter)
            res.setLabel Array(nm)
        ElseIf IsArray(nm) Then
        
            Dim idxList As New Lists
            Dim k
            Dim d As New Dicts
            
            Dim tmp As New Lists
            
            For Each k In nm
                idxList.add pLabeledArray.Item(k)
            Next k
            
            For Each k In Me.Keys
                d.add k, tmp.addAll(Me.Item(k), False).filterIndex(idxList).toArray
            Next k
            
            Set res = d
            Set d = Nothing
            Set idxList = Nothing
            res.setLabel nm
        ElseIf TypeName(nm) = "Lists" Then
            Set res = sliceWithLabel(nm.toArray)
        Else
            Err.Raise "5678", , "TypeError: the parameter of sliceWithLabel should be either String, Array or Lists. " & TypeName(nm) & "  found."
        End If
        
        Set sliceWithLabel = res
        Set res = Nothing
    Else
        Set sliceWithLabel = Nothing
    End If
End Function


' ________________________________________Class Collection Functions___________________________________________

'@desc get elements not contained in dict2 but in this dict
Public Function diff(ByVal dict2 As Dicts) As Dicts
    Dim k
    
    Dim res As Dicts
    Set res = New Dicts
    
    For Each k In Me.Keys
        If Not dict2.exists(k) Then
            res.Item(k) = Me.Item(k)
        End If
    Next k
    
    copyLabel Me, res
    
    Set diff = res
    Set res = Nothing
End Function

'@desc      get the union of two dicts
'@params
Public Function union(dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts
    Dim k
    Dim res As New Dicts
    
    For Each k In Keys
        res.add k, Item(k)
    Next k
    
    For Each k In dict2.Keys
        If Not Me.exists(k) Or Not keepOriginalVal Then
            res.add k, dict2.Item(k)
        End If
    Next k
    
    copyLabel Me, res
    
    Set union = res
    Set res = Nothing
End Function

Public Function intersect(ByRef dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts
    Dim k
    
    Dim res As New Dicts
    
    For Each k In dict2.Keys
        If Me.exists(k) Then
            If Not keepOriginalVal Then
                res.Item(k) = dict2.Item(k)
            Else
                res.Item(k) = Me.Item(k)
            End If
        End If
    Next k
    
    copyLabel Me, res
    
    Set intersect = res
    Set res = Nothing
End Function


Public Function update(ByVal dict2 As Dicts) As Dicts

    Set update = Me.union(dict2, False)
    
End Function

Public Function reduce(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal reduceWith As Long = ProcessWith.value) As Variant
     Dim l As New Lists
     
     If reduceWith = ProcessWith.value Then
        reduce = l.addAll(Me.valsArr).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
     ElseIf reduceWith = ProcessWith.key Then
        reduce = l.addAll(Me.keysArr).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
     ElseIf reduceWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set l = Nothing
End Function

Public Function reduceKey(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal reduceWith As Long = ProcessWith.value) As Variant
     reduceKey = reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint, ProcessWith.key)
End Function

Public Function reduceRngVertical(Optional ByVal operation As String = "?+_", Optional ByVal initialVal As Variant = 0, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True) As Lists
    Set pList = pList.fromArray(Me.valsArr).zipMe

    Dim l As New Lists
    Dim k

    For k = 0 To pList.length - 1
        l.add pList.getVal(k).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
    Next k
   
    Set reduceRngVertical = l
    Set l = Nothing
    pList.clear
End Function

''''''''''''
'@param     operation:              string to be evaluated, e.g. _*2 will be interpreated as ele * 2
'           placeholder:            placeholder to be replaced by the value
'           idx:                    index of the element
'           replaceDecimalPoint:    whether the Germany Decimal Point should be replace by "."
'@example   get length of the valsArray -> ProcessWith.Value
''''''''''''
Public Function map(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal mapWith As Long = ProcessWith.value) As Dicts
     
     If (Not IsReg(operation)) And TypeName(operation) <> "String" Then
        Err.Raise 8889, , "ParameterTypeError: 'operation' should be either String or RegExp!"
     End If
     
     Dim l As New Lists
     
     If mapWith = ProcessWith.value Then
        If IsReg(operation) Then
            Set map = Me.updateFromArray(l.addAll(Me.valsArr).mapReg(operation).toArray, mapWith)
        Else
            Set map = Me.updateFromArray(l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray, mapWith)
        End If
     ElseIf mapWith = ProcessWith.key Then
         If IsReg(operation) Then
            Set map = Me.updateFromArray(l.addAll(Me.keysArr).mapReg(operation).toArray, mapWith)
        Else
            Set map = Me.updateFromArray(l.addAll(Me.keysArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray, mapWith)
        End If
     ElseIf mapWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set l = Nothing
End Function

Public Function mapKey(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0) As Dicts
     Set mapKey = map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo, ProcessWith.key)
End Function

Public Function mergeMap(ByVal operation, ByRef other As Dicts, Optional ByVal placeholder As String = "_", Optional ByVal elemSelf As String = "{1}", Optional ByVal elemOther As String = "{2}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal mapWith As Long = ProcessWith.value) As Dicts
    
    Dim k
    Dim mt, ot
    Dim res As New Dicts
    
    For Each k In Me.Keys
   
        ot = IIf(other.exists(k), other.Item(k), setNullValTo)

        ot = IIf(replaceDecimalPoint, Replace("" & ot, ",", "."), ot)
        mt = Me.Item(k)
        mt = IIf(replaceDecimalPoint, Replace("" & mt, ",", "."), mt)
            
        res.Item(k) = Application.Evaluate(Replace(Replace(operation, elemSelf, mt), elemOther, ot))
        
    Next k
    
    copyLabel Me, res
    
    Set mergeMap = res
    Set res = Nothing


End Function

Public Function ranged(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal initialVal As Variant = 0, Optional ByVal aggregate As Long = AggregateMethod.AggReduce) As Dicts
    
    Dim k
    Dim res As New Dicts
    Dim l As New Lists
    
    If aggregate = AggregateMethod.AggReduce Then
        For Each k In Me.Keys
            res.Item(k) = l.addAll(Me.Item(k), False).reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint)
        Next k
    ElseIf aggregate = AggregateMethod.AggMap Then
         For Each k In Me.Keys
            res.Item(k) = l.addAll(Me.Item(k), False).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray
        Next k
    ElseIf aggregate = AggregateMethod.Aggfilter Then
        For Each k In Me.Keys
            res.Item(k) = l.addAll(Me.Item(k), False).filter(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray
        Next k
    Else
        Err.Raise 8889, , "unknown aggregate parameter"
    End If
    
    Set ranged = res
    Set res = Nothing
    Set l = Nothing
End Function

Public Function filter(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo As Variant = 0, Optional ByVal filterWith As Long = ProcessWith.value) As Dicts
     
     If (Not IsReg(operation)) And TypeName(operation) <> "String" Then
        Err.Raise 8889, , "ParameterTypeError: 'operation' should be either String or RegExp!"
     End If
     
     Dim l As New Lists
     Dim tmp As New Lists
     Dim res As New Dicts
     
     
     If filterWith = ProcessWith.value Then
     
        If IsReg(operation) Then
            Set tmp = l.addAll(Me.valsArr).judgeReg(operation)
        Else
            Set tmp = l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        End If
        
        Set res = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
        
     ElseIf filterWith = ProcessWith.key Then
        
        If IsReg(operation) Then
            Set tmp = l.addAll(Me.keysArr).judgeReg(operation)
        Else
            Set tmp = l.addAll(Me.keysArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        End If
     
        Set res = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
     ElseIf filterWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     copyLabel Me, res
     
     Set filter = res
     Set l = Nothing
     Set tmp = Nothing
End Function

Public Function filterKey(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo As Variant = 0) As Dicts
    Set filterKey = filter(operation, placeholder, idx, replaceDecimalPoint, setNullValTo, ProcessWith.key)
End Function

Public Function groupBy(ByRef attr, ByVal valCol As Long, Optional ByVal aggregateBy = xlSum) As Dicts
        
    If Not isRanged_(Me) Then
        Err.Raise 9997, , "more than one attribute in data set expected"
    End If
    
    If Not IsArray(attr) Then
        Err.Raise 9999, , "attribute array should be specified!"
    End If

    Dim l As Long
    l = arrLen(attr)

    If l = 0 Then
        Err.Raise 9998, , "at least one attribute should be contained"
    End If

    If aggregateBy <> xlSum And aggregateBy <> xlCount Then
        Err.Raise 9996, , "aggregateMethod unkown, should be xlCount or xlSum"
    End If
    
    ' lists of attributes
    Set pList = pList.fromArray(Me.valsArr).zipMe

    
    ' get value array where the aggregate method to be operated on
    Dim valArr()
    valArr = pList.getVal(valCol).toArray
    
    ' extract target attributes specified in the attr
    Dim nl As New Lists
    For l = LBound(attr) To UBound(attr)
        nl.add pList.getVal(attr(l))
    Next l
    
    ' transpose to the target data entry
    Set pList = nl.zipMe
    Set nl = Nothing
    
    Dim k, i, arr(), e
    Dim parent As Dicts
    Dim ub As Long
    Dim cnt As Long
    Dim res As New Dicts
    
    cnt = 0
    
    ' loop through the entries to update the result Dicts
    For Each k In Me.Keys
    
        arr = pList.getVal(cnt).toArray
        
        l = arrLen(arr)
        ub = UBound(arr)
        
        ' process the attributes of the entry
        ' i stands for the level of the dicts, level 0 is the root, on the top most level to perform the aggregate
        For i = LBound(arr) To UBound(arr)
        
            Set nl = Nothing
            
            Set parent = cascading(res, nl.addAll(arr, False).take(i - LBound(arr)).toArray)
            e = arr(i)
            
            If parent.exists(e) Then
                If i = ub Then
                    parent.Item(e) = parent.Item(e) + IIf(aggregateBy = xlSum, valArr(cnt), IIf(aggregateBy = xlCount, 1, 0))
                End If
            Else
                If i = ub Then
                    parent.Item(e) = IIf(aggregateBy = xlSum, valArr(cnt), IIf(aggregateBy = xlCount, 1, 0))
                Else
                    parent.add e, emptyInstance
                End If
            End If
            
            Set nl = Nothing
            
        Next i
        
        cnt = cnt + 1
        
    Next k
    
    Set groupBy = res
    
    pList.clear
    Set nl = Nothing
    Set res = Nothing
    Set parent = Nothing
    Erase valArr
    Erase arr
    
End Function

Public Function groupByLabel(attr, Optional ByVal valCol As String = "value", Optional ByVal aggregateBy = xlSum) As Dicts
        
    If Not pIsLabeled Then
        Err.Raise 9994, , "attribute labels should be specified!"
    End If
    
    If Not IsArray(attr) Then
        Err.Raise 9704, , "ParameterTypeError, attr as Array expected, " & TypeName(attr) & " provided."
    End If
    
    Dim k
    Dim col As Long
    
    col = pLabeledArray.Item(valCol)
    
    Dim attrCol()
    ReDim attrCol(0 To arrLen(attr) - 1)
    
    For k = 0 To arrLen(attr) - 1
        attrCol(k) = pLabeledArray.Item(attr(k + LBound(attr)))
    Next k
    
    Dim res As Dicts
    Set res = groupBy(attrCol, col, aggregateBy)
    res.setLabel attr
    
    Set groupByLabel = res
    Set res = Nothing
    Erase attrCol
        
End Function

' search for the sub-dicts through the array content
Private Function cascading(ByRef dict As Dicts, arr) As Dicts
    Dim e
    Dim tmp As Dicts
    Set tmp = dict
    
    For Each e In arr
       Set tmp = tmp.Item(e)
    Next e
    
    Set cascading = tmp
    Set tmp = Nothing
End Function

''''''''''''
'@desc      replace keys or vals with a new array
'@param     arr:                    the new array
'           updateWith:             keys or values to be replaced
'@return    self after update
''''''''''''
Public Function updateFromArray(ByVal arr, Optional ByVal updateWith As Long = ProcessWith.value) As Dicts
    Dim keyArr
    Dim valArr
    Dim res As Dicts
    
    keyArr = Me.keysArr
    valArr = Me.valsArr
    
    If arrLen(arr) <> arrLen(keyArr) Then
        Err.Raise 8888, , "Input Array should be the same length with the Dict"
    End If
    
    Dim l As New Lists
    
    If updateWith = ProcessWith.key Then
        Set res = arrToDict(arr, valArr, , False)
    Else
        Set res = arrToDict(keyArr, arr, , False)
    End If
    
    Set updateFromArray = res
    Erase keyArr
    Erase valArr
    Set l = Nothing
    Set res = Nothing

End Function


Public Function sort(Optional ByVal isAscending As Boolean = True, Optional ByVal sortRecursively As Boolean = True) As Dicts
    
    Dim res As New Dicts
    Dim l As New Lists
    Set l = l.addAll(Me.keysArr).sort(isAscending)
    
    Dim i
    
    For i = 0 To l.length - 1
        If sortRecursively Then
            If isDict(Me.Item(l.getVal(i))) Then
                res.add l.getVal(i), Me.Item(l.getVal(i)).sort(isAscending, sortRecursively)
            Else
                res.add l.getVal(i), Me.Item(l.getVal(i))
            End If
        Else
            res.add l.getVal(i), Me.Item(l.getVal(i))
        End If
    Next i
    
    copyLabel Me, res
    
    Set sort = res
    Set res = Nothing
    Set l = Nothing

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
        res = "[ "
        
        For Each i In arr
            If isBool(i) Or isEmpty(i) Or Not IsNumeric(i) Then
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

    If d.Count = 0 Then
        dicts_toString = "{ }"
    Else
        Dim res As String
        Dim k
        res = "{" & Chr(10)
        
        For Each k In d.Keys
            res = res & String(lvl, Chr(9)) & """" & k & """" & Chr(9) & ":" & Chr(9) & x_toString(d.Item(k), lvl + 1) & "," & Chr(10)
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
            x_toString = a_toString(x.toArray, lvl)
        Else
            If IsNull(x) Then
                x_toString = "null"
            Else
                If isNothing(x) Or isEmpty(x) Then
                    x_toString = "null"
                Else
                    If IsDate(x) Then
                        x_toString = """" & Format(x, "yyyy-mm-dd") & """"
                    Else
                        If TypeName(x) = "Boolean" Then
                            x_toString = IIf(x, "true", "false")
                        ElseIf IsNumeric(x) Then
                            x_toString = CStr(x)
                        Else
                            x_toString = """" & CStr(x) & """"
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function pk()
    Dim k
    For Each k In Me.Keys
        Debug.Print k
    Next k
End Function

Private Function getLabeledSubDict(k) As Dicts
    
    If Me.exists(k) Then
        If Me.hasLabel Then
        
            Dim res As Dicts
            Set res = Me.Item(k)
            res.label = pList.addAll(Me.label.keysArr, False).drop(1).toDict
            
            Set getLabeledSubDict = res
            Set res = Nothing
            pList.clear
        Else
            Set getLabeledSubDict = Me.Item(k)
        End If
    Else
        Set getLabeledSubDict = Nothing
    End If
    
End Function

Public Function toJSON(Optional ByVal exportTo As String) As String
    Dim res As String
    res = x_toString(Me)
    
    toJSON = res
    If Not IsMissing(exportTo) Then
        Dim fso As Object
        Set fso = CreateObject("scripting.filesystemobject")
        
        Dim targPath As String
        targPath = ThisWorkbook.Path & "\" & exportTo
        
        Dim ts As Object
        Set ts = fso.createtextfile(targPath)
        
        ts.writeline res
        ts.Close
        
        Set ts = Nothing
        Set fso = Nothing
        targPath = ""
    End If
    
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
' start/ending can be column label in string
Public Function rng(ByVal start, ByVal ending, Optional ByVal steps As Long = 1)

    If isInstanceOf(start, "String") Then
        start = letterToColNum(start)
    End If
    
     If isInstanceOf(ending, "String") Then
        ending = letterToColNum(ending)
    End If

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

Public Function y(Optional ByVal sht As String = "", Optional ByVal col As Long = 1, Optional ByVal wb As Workbook) As Long
    y = getTargetSht(sht, wb).Cells(Rows.Count, col).End(xlUp).row
End Function

Public Function x(Optional ByVal sht As String = "", Optional ByVal row As Long = 1, Optional ByVal wb As Workbook) As Long
    x = getTargetSht(sht, wb).Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Private Function IsReg(testObj) As Boolean
    IsReg = TypeName(testObj) = "IRegExp2"
End Function

Public Function isDict(testObj As Variant) As Boolean
   isDict = TypeName(testObj) = "Dicts"
End Function

Public Function isBool(testObj As Variant) As Boolean
   isBool = TypeName(testObj) = "Boolean"
End Function
Public Function isNothing(testObj As Variant) As Boolean
   If IsObject(testObj) Then
        isNothing = testObj Is Nothing
    Else
        isNothing = False
   End If
   
End Function
Public Function isInstanceOf(testObj, typeArr) As Boolean
    Dim s As String
    s = TypeName(testObj)
    
    If TypeName(typeArr) = "String" Then
        isInstanceOf = s = typeArr
    ElseIf IsArray(typeArr) Then
        Dim k
        
        Dim res As Boolean
        res = False
        
        For Each k In typeArr
            If s = k Then
                res = True
                Exit For
            End If
        Next k
        isInstanceOf = res
        
    ElseIf isInstanceOf(typeArr, "Lists") Then
        isInstanceOf = isInstanceOf(testObj, typeArr.toArray)
    Else
        Err.Raise 9980, , "ParameterTypeErrorException: typeArr should be either String, Array or Lists"
    End If
End Function

Public Function withSameType(obj1, obj2) As Boolean
    withSameType = TypeName(obj1) = TypeName(obj2)
End Function

Public Function letterToColNum(ByVal l As String) As Integer
    letterToColNum = Range(l & "1").Column
End Function

Public Function fromString(ByRef s As String, Optional ByRef i As Long = 1) As Variant
    
    skipSpace s, i
    
    Select Case Mid$(s, i, 1)
    Case "["
        Set fromString = listFromString(s, i)
    Case "{"
        Set fromString = dictFromString(s, i)
    Case """", "'"
        fromString = strFromString(s, i)
    Case Else
        fromString = elementFromString(s, i)
    End Select

End Function

' element at i is "{"
Public Function dictFromString(ByRef s As String, Optional ByRef i As Long = 1) As Dicts
    
    Dim stack As New Lists
    Dim res As New Dicts
    Dim k As String
    
    skipSpace s, i
    
    stack.add i
    
    i = i + 1
    
    Do
    Select Case Mid$(s, i, 1)
        Case "{"
            stack.add i
            i = i + 1
            
            skipSpace s, i
            k = strFromString(s, i)
            skipSpace s, i
            res.add k, fromString(s, i)
            
            If i >= Len(s) Then GoTo endFunc
          '  i = i - 1
        Case ",", " ", VBA.vbCr, VBA.vbTab
            i = i + 1
        Case "}"
            Set stack = stack.dropLast(1)
            
            i = i + 1
            If stack.length = 0 Then GoTo endFunc
        Case ":"
            i = i + 1
            res.add k, fromString(s, i)
        Case Else
            k = strFromString(s, i)
    End Select
    
    Loop While i < Len(s)
    
endFunc:
    Set dictFromString = res

End Function

' element at i is "["
Public Function listFromString(ByRef s As String, Optional ByRef i As Long = 1) As Lists
    
    Dim stack As New Lists
    Dim res As New Lists
    
    skipSpace s, i
    
    stack.add i
    
    i = i + 1
    
    Do
    Select Case Mid$(s, i, 1)
        Case "["
            stack.add i
            res.add listFromString(s, i)
            
            If i >= Len(s) Then GoTo endFunc
            i = i - 1
        Case "]"
            Set stack = stack.dropLast(1)
        
            i = i + 1
            If stack.length = 0 Then GoTo endFunc
        Case ","
            i = i + 1
        Case Else
            res.add fromString(s, i)
    End Select
    
    Loop
    
endFunc:
    Set listFromString = res
    
End Function

Private Function skipSpace(ByRef s As String, Optional ByRef i As Long = 1)
    
    Do
        Select Case Mid$(s, i, 1)
        Case " ", VBA.vbCr, VBA.vbTab
            i = i + 1
        Case Else
            Exit Function
        End Select
    Loop
End Function

Public Function elementFromString(ByRef s As String, Optional ByRef i As Long = 1) As Variant
    If Mid$(s, i, 4) = "true" Then
        elementFromString = True
        i = i + 4
    ElseIf Mid$(s, i, 5) = "false" Then
        elementFromString = False
        i = i + 5
    ElseIf Mid$(s, i, 4) = "null" Then
        elementFromString = Null
        i = i + 4
    Else
        elementFromString = numericFromString(s, i)
    End If
End Function


Public Function strFromString(ByRef s As String, ByRef i As Long) As String
    Dim start As Long
    start = i + 1
    
    Dim quotation As String
    quotation = Mid$(s, i, 1)
    
    i = i + 1
    
    Do
        If Mid$(s, i, 1) = quotation Then
            strFromString = Mid$(s, start, i - start)
            i = i + 1
            Exit Function
        End If
        
        i = i + 1
    Loop
End Function

Public Function numericFromString(ByRef s As String, ByRef i As Long) As Double
    Dim start As Long
    start = i
    
    Do
        If Not InStr("-.0123456789 ", Mid$(s, i, 1)) > 0 Then
            numericFromString = CDbl(Replace(Trim(Mid$(s, start, i - start)), ".", Application.DecimalSeparator))
            Exit Function
        End If
        
        i = i + 1
    Loop
End Function

