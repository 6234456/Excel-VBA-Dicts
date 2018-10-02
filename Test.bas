Sub Test()
    Dim d As New Dicts
    Dim l As New Lists

    Debug.Assert d.count = 0
    Debug.Assert l.length = 0
    
    l.remove 1
    l.removeAt 0
    
    Debug.Assert d.wb.Name = ActiveWorkbook.Name
    
    With d.fromArray(l.fromSerial(1, 10))
        Debug.Assert .dict(10) = 9
        Debug.Assert .keysArr(9) = 10
        Debug.Assert l.fromSerial(1, 10).contains(10)
        
        Debug.Assert l.addAll(.valsArr, False).map("_*2").getVal(9) = 18
        Debug.Assert l.addAll(.keysArr, False).contains(10)
        
        Debug.Assert .count = 10
        Debug.Assert .reduceKey("_+?", -1) = 54
        Debug.Assert .reduce("_+?", 10) = 55
        
        Debug.Assert .filterKey("_>8").count = 2
        Debug.Assert .filter("_>8").count = 1
        
        Debug.Assert .mapKey("_+6").keysArr(0) = 7
        Debug.Assert .map("_+6").valsArr(0) = 6
        
        Debug.Assert .diff(l.fromSerial(-10, 8).toDict).reduceKey("_+?", 0) = 19
    End With
    
    Dim l2 As New Lists
    l.clear
    l.add l2
    l2.add d
    Debug.Assert TypeName(l.getVal(0, 0)) = "Dicts"
    
    l.clear
    Debug.Assert l.of(1, 2, 3, 4, 5, Array(1, 2, 3)).length = 6
    
    Debug.Assert l.fromSerial(10, 15).mapX("Test.callback").slice(-1).getVal(0) = "225_"
    
    Debug.Assert l.fromSerial(1, 10).subgroupBy(2, 2).mapX("Test.m").reduce("_+?", 0) = 10
    
    Debug.Assert l.fromSerial(10, 15).filterX("Test.f").length = 0
    
    Debug.Assert l.fromSerial(10, 15).reduceX("Test.r", New Dicts).count = 6
    
    With l.of(1, 2, 3, 4).permutation
     Debug.Assert .length = 4 * 3 * 2 * 1
    End With
    
    Debug.Print "All tests passed!"
    
End Sub


Private Sub callback(ByRef l As Lists, e, Optional ByVal i As Long)
    l.callback = e ^ 2 & "_"
End Sub

Private Sub f(ByRef l As Lists, e, Optional ByVal i As Long)
    l.callback = i > 15
End Sub

Private Sub r(ByRef l As Lists, e, Optional ByVal i As Long)
    l.callback.dict.add e, 1
End Sub

Private Sub m(ByRef l As Lists, e, Optional ByVal i As Long)
    l.callback = e.length
End Sub
