 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class Dicts
'@author                                   Qiou Yang
'@lastUpdate                               29.08.2018
'                                          code refactor
'                                          integrate load/reduce/map/filter into single function
'                                          new feature: load horizontally: loadH
'                                          new feature: groupBy
'@TODO                                     add comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' declaration compulsory
Option Explicit

'___________private variables_____________
'scripting.Dictionary Object
Private pDict As Object

' has column label
Private pIsLabeled As Boolean

' column label as Dicts, label -> index
Private pLabeledArray As Dicts

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

Private Sub Class_Initialize()
    ini
    Set pList = New Lists
End Sub

Private Sub Class_Terminate()
    Set pWb = Nothing
    Set pDict = Nothing
    Set pLabeledArray = Nothing
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
                 d.dict(Trim(CStr(c.Value))) = cnt
                 cnt = cnt + 1
             Next c
             
             Me.setLabel d
         Else
             If isDict(rng) Then
                 Set pLabeledArray = rng
             ElseIf IsArray(rng) Then
                 Dim k
                 
                 For k = 0 To UBound(rng) - LBound(rng)
                     d.dict(rng(k)) = k
                 Next k
                 
                 Me.setLabel d
             End If
         End If
         
        pIsLabeled = True
   End If
   
   Set setLabel = Me
   
End Function

' get length of the key-value pairs
' if recursive set to true, count the keys of all child-dicts
' allLevels only relevant in recursive-mode, count all the keys in the structure
Public Function count(Optional ByVal recursive As Boolean = False, Optional ByVal allLevels As Boolean = True) As Long
    
    If Not recursive Then
        count = pDict.count
    Else
        If isDicted_(Me) Then
            Dim k
            Dim res As Long
            
            For Each k In Me.keys
                res = IIf(allLevels, 1, 0) + res + Me.dict(k).count(True)
            Next k
            
            count = res
        Else
            count = pDict.count
        End If
    End If
End Function

' get keys as Array, if no element return null-Array
Public Property Get keysArr() As Variant
    
    Dim res()
    
    If Me.count > 0 Then
        ReDim res(0 To Me.count - 1)
        
        Dim k
        Dim cnt As Long
        cnt = 0
        
        For Each k In Me.keys
            res(cnt) = k
            cnt = cnt + 1
        Next k
    End If
    
    keysArr = res
    Erase res
    
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
    Erase res
End Property

' get keys as iterable-object
Public Property Get keys() As Variant
    keys = pDict.keys
End Property

' set underlying scripting.Dictionary Object
Public Property Let dict(ByRef dict As Object)
    Set pDict = dict
End Property

Public Property Get sign() As String
    sign = "Dicts"
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
                targRowEnd = .Cells(.Rows.count, targKeyCol).End(xlUp).row
            Else
                targRowEnd = .Cells(targKeyCol, .Columns.count).End(xlToLeft).Column
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
    Set lst = fst.Offset(rng.Rows.count - 1, rng.Columns.count - 1)
    
    Dim shtName As String
    Dim wbName As String
    shtName = "'" & fst.Worksheet.Name & "'!"
    wbName = "'[" & fst.Worksheet.parent.Name & "]" & fst.Worksheet.Name & "'!"
        
    Dim i As Long
    Dim j As Long
    
    Dim res()
    ReDim res(1 To rng.Rows.count, 1 To rng.Columns.count)
   
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
Function arrToDict(keyArr, valArr, Optional ByVal isReversed As Boolean = False, Optional ByVal keyCstr As Boolean = False) As Object
    
    Dim res As Object
    Set res = CreateObject("scripting.dictionary")
    res.compareMode = vbTextCompare
    
    ' combine the key-value pair in a zipped mode
    If arrLen(keyArr) = 0 Then
    
    ElseIf arrLen(keyArr) = 1 And arrLen(valArr) > 1 Then
        res(keyArr(LBound(keyArr))) = valArr
    Else
        If arrLen(keyArr) <> arrLen(valArr) Then
            Err.Raise 8888, "", "Arrays with different length can not be combined"
        End If
        
        Dim i
        
        If isReversed Then
            For i = UBound(keyArr) To LBound(keyArr) Step -1
                If Len(Trim(CStr(keyArr(i)))) > 0 Then
                    res.add IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i)), valArr(UBound(valArr) + i - UBound(valArr))
                End If
            Next i
        Else
            For i = LBound(keyArr) To UBound(keyArr)
                If Len(Trim(CStr(keyArr(i)))) > 0 Then
                   res.add IIf(keyCstr, Trim(CStr(keyArr(i))), keyArr(i)), valArr(LBound(valArr) + i - LBound(valArr))
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
Function rngToDict(ByRef keyRng As Range, ByRef valRng As Range, Optional ByVal isReversed As Boolean = False, Optional ByVal asAddress As Boolean = False) As Object
    
    ' if the keyRng contains only one column, it is vertical
    Dim isVertical As Boolean
    isVertical = keyRng.Columns.count = 1
    
    Set rngToDict = arrToDict(rngToArr(keyRng, Not isVertical), IIf(IIf(isVertical, valRng.Columns.count, valRng.Rows.count) = 1, rngToArr(valRng, Not isVertical, asAddress), rngToArr(valRng, isVertical, asAddress)), isReversed)
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
Public Function load(Optional ByVal Sht As String = "", Optional ByVal KeyCol As Long = 1, Optional ByVal valCol = 1, Optional RowBegine As Variant = 1, Optional ByVal RowEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False, Optional ByRef asAddress As Boolean = False, Optional appendMode As Boolean = False, Optional ByVal isVertical As Boolean = True) As Dicts
    Dim keyRng As Range
    Set keyRng = getRange(Sht, KeyCol, KeyCol, RowBegine, RowEnd, wb, isVertical)
    
    Dim valRng As Range
    Set valRng = getRange(Sht, KeyCol, valCol, RowBegine, RowEnd, wb, isVertical)
    
    If pDict.count = 0 Or Not appendMode Then
        Set pDict = rngToDict(keyRng, valRng, Reversed, asAddress)
    Else
       Set pDict = Me.union(createInstance(Me.rngToDict(keyRng, valRng, Reversed, asAddress))).dict
    End If
    
    Set load = Me
    
    Set keyRng = Nothing
    Set valRng = Nothing
End Function

Public Function loadH(Optional ByVal Sht As String = "", Optional ByVal KeyRow As Long = 1, Optional ByVal ValRow = 1, Optional ColBegine As Variant = 1, Optional ByVal ColEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False, Optional ByRef asAddress As Boolean = False, Optional appendMode As Boolean = False) As Dicts
    Set loadH = load(Sht:=Sht, KeyCol:=KeyRow, valCol:=ValRow, RowBegine:=ColBegine, RowEnd:=ColEnd, wb:=wb, Reversed:=Reversed, asAddress:=asAddress, appendMode:=appendMode, isVertical:=False)
End Function


'@desc update self with new dictionary obj
Public Function of(ByRef dictObj As Object) As Dicts
    Set pDict = dictObj
    Set of = Me
End Function

'@desc create a new instance with the dictionary obj
Public Function createInstance(ByRef dictObj As Object) As Dicts
    Dim res As New Dicts
    Set createInstance = res.of(dictObj)
    Set res = Nothing
End Function

Public Function emptyInstance() As Dicts
    Dim res As New Dicts
    Set emptyInstance = res
    Set res = Nothing
End Function


Public Function loadStruct(ByVal Sht As String, ByVal KeyCol1 As Long, ByVal KeyCol2 As Long, ByVal valCol, Optional RowBegine As Variant, Optional ByVal RowEnd As Variant, Optional ByRef wb As Workbook, Optional ByRef Reversed As Boolean = False) As Dicts

    With getTargetSht(Sht, wb)
    
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.compareMode = vbTextCompare
        
        If IsMissing(RowBegine) Then
            RowBegine = 1
        End If
        
        If IsMissing(RowEnd) Then
            RowEnd = .Cells(Rows.count, KeyCol2).End(xlUp).row
        End If
        
        Dim tmpPreviousRow As Long
        Dim tmpCurrentRow As Long
        Dim tmpDict As New Dicts
        
        tmpPreviousRow = RowEnd
        tmpCurrentRow = tmpPreviousRow
        
        Do While tmpCurrentRow > RowBegine
            tmpCurrentRow = .Cells(tmpCurrentRow, KeyCol1).End(xlUp).row
            
            Set dict(.Cells(tmpCurrentRow, KeyCol1).Value) = tmpDict.load(Sht, KeyCol2, valCol, tmpCurrentRow + 1, tmpPreviousRow, wb, Reversed)
            Set tmpDict = Nothing
            
            tmpPreviousRow = tmpCurrentRow - 1
        Loop
        
        Set loadStruct = Me.of(dict)
    End With

End Function

' rng can be Range Object or an array
Public Function frequencyCount(ByRef rng) As Dicts

    Dim res As New Dicts
    Dim k

    If Not IsArray(rng) Then
        For Each k In rng.Cells
            If Len(k.Value) > 0 Then
                If res.exists(k.Value) Then
                    res.dict(k.Value) = res.dict(k.Value) + 1
                Else
                    res.dict(k.Value) = 1
                End If
            End If
        Next k
    Else
         For Each k In rng
            If Len(k) > 0 Then
                If res.exists(k) Then
                    res.dict(k) = res.dict(k) + 1
                Else
                    res.dict(k) = 1
                End If
            End If
        Next k
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
               endRow = .Cells(.Rows.count, keyPos).End(xlUp).row
            End If
            
            For Each c In .Cells(startingRow, keyPos).Resize(endRow - startingRow + 1, 1).Cells
                If pDict.exists(c.Value) Then
                    tmp = pDict(c.Value)
                    If IsArray(tmp) Then
                        If IsMissing(endCol) Or endCol = 0 Then
                            .Cells(c.row, startingCol).Resize(1, arrLen(tmp)).Value = tmp
                        Else
                             l = pList.fromArray(tmp, False).take(endCol - startingCol + 1).toArray
                            .Cells(c.row, startingCol).Resize(1, arrLen(l)).Value = l
                        End If
                    Else
                        .Cells(c.row, startingCol).Value = tmp
                    End If
                End If
            Next c
            
        Else
            If IsMissing(endCol) Or endCol = 0 Then
               endCol = .Cells(keyPos, .Columns.count).End(xlToLeft).Column
            End If
            
            For Each c In .Cells(keyPos, startingCol).Resize(1, endCol - startingCol + 1).Cells
                If pDict.exists(c.Value) Then
                    tmp = pDict(c.Value)
                    If IsArray(tmp) Then
                        If IsMissing(endRow) Or endRow = 0 Then
                            .Cells(startingRow, c.Column).Resize(arrLen(tmp), 1).Value = Application.WorksheetFunction.Transpose(tmp)
                        Else
                            l = pList.fromArray(tmp, False).take(endCol - startingCol + 1).toArray
                            .Cells(startingRow, c.Column).Resize(arrLen(l), 1).Value = Application.WorksheetFunction.Transpose(l)
                        End If
                    Else
                        .Cells(startingRow, c.Column).Value = tmp
                    End If
                End If
            Next c
        End If
    End With

End Sub

Public Sub dump(ByVal shtName As String, Optional ByVal keyPos As Long = 1, Optional ByVal startingRow As Long = 1, Optional ByVal startingCol As Long = 2, Optional ByVal endRow As Long, Optional ByVal endCol As Long, Optional ByRef wb As Workbook, Optional ByVal isVertical As Boolean = True, Optional ByVal trailingRows As Long = 0)

    With getTargetSht(shtName, wb)

        If isDicted_(Me) Then
            Dim k
            Dim cnt As Long
            
            For Each k In Me.keys
                If isVertical Then
                    .Cells(startingRow + cnt, keyPos) = k
                Else
                    .Cells(keyPos, startingCol + cnt) = k
                End If
            
                Me.dict(k).dump shtName, keyPos + 1, startingRow + cnt + 1, startingCol + 1, startingRow + cnt + Me.dict(k).count(True), endCol, wb, isVertical, trailingRows
                cnt = cnt + Me.dict(k).count(True) + 1
    
            Next k

        Else
             'unload the key
            If isVertical Then
                .Cells(startingRow, keyPos).Resize(Me.count, 1) = Application.WorksheetFunction.Transpose(Me.keysArr)
            Else
                .Cells(keyPos, startingCol).Resize(1, Me.count) = Me.keysArr
            End If
        
            Me.unload shtName, keyPos, startingRow, startingCol, endRow, endCol, wb, isVertical
        End If
        
    End With
    
    
End Sub

Public Function exists(ByVal k) As Boolean
    exists = pDict.exists(k)
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

' if containing arrays as element
Private Function isRanged_(ByRef obj As Dicts) As Boolean
    
    Dim k
    For Each k In obj.keys
        isRanged_ = IsArray(obj.dict(k))
        Exit For
    Next k
    
End Function

' if containing Dicts as elements
Private Function isDicted_(ByRef obj As Dicts) As Boolean
    
    Dim k
    For Each k In obj.keys
        isDicted_ = isDict(obj.dict(k))
        Exit For
    Next k
    
End Function

Public Function sliceWithLabel(ByVal nm As String) As Dicts
    If pIsLabeled Then
        Dim i As Long
        Dim res As Dicts
        i = pLabeledArray.dict(nm)
        
        Set res = Me.ranged("{i}=" & i, aggregate:=AggregateMethod.Aggfilter)
        res.setLabel Array(nm)
        
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
    
    For Each k In pDict.keys
        If Not dict2.dict.exists(k) Then
            res.dict(k) = pDict(k)
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
    res.dict = pDict
    
    For Each k In dict2.dict.keys
        If Not pDict.exists(k) Then
            res.dict(k) = dict2.dict(k)
        ElseIf Not keepOriginalVal Then
            res.dict(k) = dict2.dict(k)
        End If
    Next k
    
    copyLabel Me, res
    
    Set union = res
    Set res = Nothing
End Function

Public Function intersect(ByRef dict2 As Dicts, Optional ByVal keepOriginalVal As Boolean = True) As Dicts
    Dim k
    
    Dim res As New Dicts
    
    For Each k In dict2.dict.keys
        If pDict.exists(k) Then
            If Not keepOriginalVal Then
                res.dict(k) = dict2.dict(k)
            Else
                res.dict(k) = pDict(k)
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

Public Function reduce(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal reduceWith As Long = ProcessWith.Value) As Variant
     Dim l As New Lists
     
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

Public Function reduceKey(ByVal operation As String, ByVal initialVal As Variant, Optional ByVal placeholder As String = "_", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal reduceWith As Long = ProcessWith.Value) As Variant
     reduceKey = reduce(operation, initialVal, placeholder, placeholderInitialVal, replaceDecimalPoint, ProcessWith.Key)
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
Public Function map(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal mapWith As Long = ProcessWith.Value) As Dicts
     
     If (Not IsReg(operation)) And TypeName(operation) <> "String" Then
        Err.Raise 8889, , "ParameterTypeError: 'operation' should be either String or RegExp!"
     End If
     
     Dim l As New Lists
     
     If mapWith = ProcessWith.Value Then
        If IsReg(operation) Then
            Set map = Me.updateFromArray(l.addAll(Me.valsArr).mapReg(operation).toArray, mapWith)
        Else
            Set map = Me.updateFromArray(l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo).toArray, mapWith)
        End If
     ElseIf mapWith = ProcessWith.Key Then
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
     Set mapKey = map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo, ProcessWith.Key)
End Function

Public Function ranged(ByVal operation As String, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal placeholderInitialVal As String = "?", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo = 0, Optional ByVal initialVal As Variant = 0, Optional ByVal aggregate As Long = AggregateMethod.AggReduce) As Dicts
    
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

Public Function filter(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo As Variant = 0, Optional ByVal filterWith As Long = ProcessWith.Value) As Dicts
     
     If (Not IsReg(operation)) And TypeName(operation) <> "String" Then
        Err.Raise 8889, , "ParameterTypeError: 'operation' should be either String or RegExp!"
     End If
     
     Dim l As New Lists
     Dim tmp As New Lists
     Dim res As New Dicts
     
     
     If filterWith = ProcessWith.Value Then
     
        If IsReg(operation) Then
            Set tmp = l.addAll(Me.valsArr).judgeReg(operation)
        Else
            Set tmp = l.addAll(Me.valsArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        End If
        
        res.dict = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
        
     ElseIf filterWith = ProcessWith.Key Then
        
        If IsReg(operation) Then
            Set tmp = l.addAll(Me.keysArr).judgeReg(operation)
        Else
            Set tmp = l.addAll(Me.keysArr).map(operation, placeholder, idx, replaceDecimalPoint, setNullValTo)
        End If
     
        res.dict = Me.arrToDict(l.addAll(Me.keysArr, False).filterWith(tmp).toArray, l.addAll(Me.valsArr, False).filterWith(tmp).toArray)
     ElseIf filterWith = ProcessWith.RangedValue Then
        Err.Raise 8889, , "to process RangedValue please refer to ranged"
     Else
        Err.Raise 8889, , "unknown aggregate parameter"
     End If
     
     Set filter = res
     Set l = Nothing
     Set tmp = Nothing
End Function

Public Function filterKey(ByVal operation, Optional ByVal placeholder As String = "_", Optional ByVal idx As String = "{i}", Optional ByVal replaceDecimalPoint As Boolean = True, Optional ByVal setNullValTo As Variant = 0) As Dicts
    Set filterKey = filter(operation, placeholder, idx, replaceDecimalPoint, setNullValTo, ProcessWith.Key)
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
    For Each k In Me.keys
    
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
                    parent.dict(e) = parent.dict(e) + IIf(aggregateBy = xlSum, valArr(cnt), IIf(aggregateBy = xlCount, 1, 0))
                End If
            Else
                If i = ub Then
                    parent.dict(e) = IIf(aggregateBy = xlSum, valArr(cnt), IIf(aggregateBy = xlCount, 1, 0))
                Else
                    parent.dict.add e, emptyInstance
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
    
    col = pLabeledArray.dict(valCol)
    
    Dim attrCol()
    ReDim attrCol(0 To arrLen(attr) - 1)
    
    For k = 0 To arrLen(attr) - 1
        attrCol(k) = pLabeledArray.dict(attr(k + LBound(attr)))
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
       Set tmp = tmp.dict(e)
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
Public Function updateFromArray(ByVal arr, Optional ByVal updateWith As Long = ProcessWith.Value) As Dicts
    Dim keyArr
    Dim valArr
    Dim res As New Dicts
    
    keyArr = Me.keysArr
    valArr = Me.valsArr
    
    If arrLen(arr) <> arrLen(keyArr) Then
        Err.Raise 8888, , "Input Array should be the same length with the Dict"
    End If
    
    Dim l As New Lists
    
    If updateWith = ProcessWith.Key Then
        res.dict = arrToDict(arr, valArr, , False)
    Else
        res.dict = arrToDict(keyArr, arr, , False)
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
            If isDict(Me.dict(l.getVal(i))) Then
                res.dict.add l.getVal(i), Me.dict(l.getVal(i)).sort(isAscending, sortRecursively)
            Else
                res.dict.add l.getVal(i), Me.dict(l.getVal(i))
            End If
        Else
            res.dict.add l.getVal(i), Me.dict(l.getVal(i))
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

Private Function getLabeledSubDict(k) As Dicts
    
    If pDict.exists(k) Then
        If Me.hasLabel Then
        
            Dim res As Dicts
            Set res = pDict(k)
            res.label = pList.addAll(Me.label.keysArr, False).drop(1).toDict
            
            Set getLabeledSubDict = res
            Set res = Nothing
            pList.clear
        Else
            Set getLabeledSubDict = pDict(k)
        End If
    Else
        Set getLabeledSubDict = Nothing
    End If
    
End Function


Public Function toJSON(Optional ByVal label As String = "", Optional ByVal lvl As Long = 0) As String
    
    If Len(label) = 0 Then
        If pIsLabeled Then
            label = Me.label.keysArr(0)
        Else
            label = "root"
        End If
    End If

    Dim res As String
    res = String(lvl, Chr(9)) & "{""name"":""" & label & """," & Chr(13)
    res = res & String(lvl, Chr(9)) & """children"":[" & Chr(13)
    
    Dim ky
    For Each ky In pDict.keys
        If isDict(pDict(ky)) Then
            If pIsLabeled Then
                res = res & getLabeledSubDict(ky).toJSON(ky, lvl + 1) & "," & Chr(13)
            Else
                res = res & pDict(ky).toJSON(ky, lvl + 1) & "," & Chr(13)
            End If
        Else
            res = res & String(lvl + 1, Chr(9)) & "{""name"":""" & Replace(CStr(ky), """", "") & """, " & """value"": " & Replace(CStr(pDict(ky)), ",", ".") & "}," & Chr(13)
        End If
    Next ky
    
    toJSON = Left(res, Len(res) - 2) & Chr(13) & String(lvl, Chr(9)) & "]}"
    
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

Public Function y(Optional ByVal Sht As String = "", Optional ByVal col As Long = 1, Optional ByVal wb As Workbook) As Long
    y = getTargetSht(Sht, wb).Cells(Rows.count, col).End(xlUp).row
End Function

Public Function x(Optional ByVal Sht As String = "", Optional ByVal row As Long = 1, Optional ByVal wb As Workbook) As Long
    x = getTargetSht(Sht, wb).Cells(row, Columns.count).End(xlToLeft).Column
End Function

Private Function IsReg(testObj) As Boolean
    IsReg = TypeName(testObj) = "IRegExp2"
End Function

Public Function isDict(testObj As Variant) As Boolean
   isDict = TypeName(testObj) = "Dicts"
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
