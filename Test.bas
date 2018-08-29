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
    
    Debug.Print "All tests passed!"
    
End Sub


Sub TestShtIO()
    
    Dim d As New Dicts
    
    With d.load("1", 1, d.rng("B", "J"), 2).setLabel(Worksheets("1").Range("B1:J1"))
        .groupByLabel(Array("location", "resp", "cate")).dump "3"
    End With
    
End Sub
