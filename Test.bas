Sub Test()
    Dim d As New Dicts
    Dim l As New Lists

    Debug.Assert d.Count = 0
    Debug.Assert l.length = 0
    
    l.Remove 1
    l.removeAt 0
    
    Dim d2 As New Dicts
    d2.add 2, 1
    Debug.Assert l.fromSerial(1, 10).toMap.diff(d2).Count = 9
    
    d2.load("", 1, 2, appendMode:=True).p
    
    Debug.Assert d.wb.Name = ActiveWorkbook.Name
    
    With d.fromArray(l.fromSerial(1, 10))
        Debug.Assert .dict.Item(10) = 9
        Debug.Assert .keysArr(9) = 10
        Debug.Assert l.fromSerial(1, 10).contains(10)
        
        Debug.Assert l.addAll(.valsArr, False).map("_*2").getVal(9) = 18
        Debug.Assert l.addAll(.keysArr, False).contains(10)
        
        Debug.Assert .Count = 10
        Debug.Assert .reduceKey("_+?", -1) = 54
        Debug.Assert .reduce("_+?", 10) = 55
        
        Debug.Assert .filterKey("_>8").Count = 2
        Debug.Assert .filter("_>8").Count = 1
        
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
    
    Debug.Assert l.fromSerial(10, 15).reduceX("Test.r", New Dicts).Count = 6
    
    With l.of(1, 2, 3, 4).permutation
     Debug.Assert .length = 4 * 3 * 2 * 1
    End With
    
    d.clear
    l.clear
    
    
    d.add 1, Array(l.of(1, 2, 3), 2, 3, 4)
    d.add 2, Array(l.of(2, 2, 3), 2, 3, 4)
    d.add 3, Array(l.of(3, 2, 3), 2, 3, 4)
    d.add 4, "Qiou"
    d.add 5, Array(l.of("qiou", Null, Date), 1, 23.4, Empty, True)
    d.add 6, Nothing
    
    Dim d1 As New Dicts
    Set d1 = d1.fromArray(Array(1, 2, 3))
    d1.add 4, Array(1, 2, l.of(1, 2, 3, True, Empty, "qiou"), Null, d1.fromArray(Array("qiou", "yang")))
    
    d.add 7, d1
    
    d.add 8, 2
    
    d.label = Array("index", "a1", "a2", "a3")
    
    Debug.Assert d.getByLabel(3, "index").getVal(0) = 3

    Debug.Assert d.x_toString(Null) = "null"
    Debug.Assert d.x_toString(Nothing) = "null"
    Debug.Assert d.x_toString(Empty) = "null"
    Debug.Assert d.x_toString(True) = "true"
    Debug.Assert d.x_toString(Array(Null, Nothing)) = "[ null, null ]"
    Debug.Assert l.of(Null, Null).toString = "[ null, null ]"

    Debug.Assert Not d.isNothing("123")
    
    d.p
    
   ' d.toJSON ("src.json")
    
    
    Dim t As New TreeSets
    
    t.add 12
    t.add 3
    t.addAll 1, 2, 3, 4, 5

    Debug.Assert t.contains(3)
    
    Debug.Assert IsNull(t.ceiling(14))


    Dim m As New TreeMaps
    
    m.add 1, 2
    m.add 2, 3
    m.add 2, 4
    m.add "2", 5
    
    Debug.Assert m.Count = 3
    
    Debug.Assert Not m.exists(8)
    m.Item(8) = 0
    
    Debug.Assert m.Item(8) = 0
    
    m.Item("2") = 10
    
    m.Remove 8
    
    Debug.Assert m.Item("2") = 10
    Debug.Assert m.Count = 3
    
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


Function solutions(ByRef l As Lists) As Lists
    
    Dim res As New Lists
    
    If l.length = 0 Then
        res.add 0
    ElseIf l.length = 1 Then
        res.add l.getVal(0)
    ElseIf l.length = 2 Then
    
        Set res = res.of(l.getVal(0) + l.getVal(1), l.getVal(0) * l.getVal(1), l.getVal(1) - l.getVal(0), l.getVal(0) - l.getVal(1))
        
        If l.getVal(1) <> 0 Then
            res.add l.getVal(0) / l.getVal(1)
        End If
        
        If l.getVal(0) <> 0 Then
            res.add l.getVal(1) / l.getVal(0)
        End If
    Else
        Dim j, k

        Dim tmp1 As Lists
        Dim tmp2 As Lists

        Set tmp1 = l
        For j = 0 To tmp1.length - 1
            Set tmp2 = solutions(tmp1.copy.removeAt(j))
            For k = 0 To tmp2.length - 1
                res.addAll solutions(l.of(tmp1.getVal(j), tmp2.getVal(k)))
            Next k
        Next j
        
        Set tmp1 = Nothing
        Set tmp2 = Nothing

    End If
    
    Set solutions = res.unique
    Set res = Nothing
    
End Function

Sub demoSolutions()
    
    
    Dim l As New Lists
    
    With solutions(l.of(5, 8, 1, 4, 9))
        Debug.Print .contains(24)
        .p
    End With
    
End Sub



Function solutions1(ByRef l As Lists) As Dicts
    
    Dim res As New Dicts
    
    If l.length = 0 Then
        res.add 0, 0
    ElseIf l.length = 1 Then
        res.add l.getVal(0), l.getVal(0)
    ElseIf l.length = 2 Then
    
        res.add l.getVal(0) + l.getVal(1), l.getVal(0) & " + " & l.getVal(1)
        res.add l.getVal(0) * l.getVal(1), l.getVal(0) & " * " & l.getVal(1)
        res.add l.getVal(0) - l.getVal(1), l.getVal(0) & " - " & l.getVal(1)
        res.add l.getVal(1) - l.getVal(0), l.getVal(1) & " - " & l.getVal(0)
        
        If l.getVal(1) <> 0 Then
            res.add l.getVal(0) / l.getVal(1), l.getVal(0) & " / " & l.getVal(1)
        End If
        
        If l.getVal(0) <> 0 Then
            res.add l.getVal(1) / l.getVal(0), l.getVal(1) & " / " & l.getVal(0)
        End If
    Else
        Dim j, k, i

        Dim tmp1 As Lists
        Dim tmp2 As New Lists
        Dim tmp As Dicts
        Dim d2 As Dicts
        Dim tmp3

        Set tmp1 = l
        For j = 0 To tmp1.length - 1
            Set d2 = solutions1(tmp1.copy.removeAt(j))
            tmp2.addAll d2.Keys, False
            For k = 0 To tmp2.length - 1
                Set tmp = solutions1(l.of(tmp1.getVal(j), tmp2.getVal(k)))
                For Each i In tmp.Keys
                    tmp3 = Split(tmp.dict(i), " ")
                    
                    If Trim(tmp3(0)) = CStr(tmp1.getVal(j)) Then
                        res.add i, tmp1.getVal(j) & " " & tmp3(1) & " ( " & d2.dict(tmp2.getVal(k)) & " ) "
                    Else
                        res.add i, " ( " & d2.dict(tmp2.getVal(k)) & " ) " & " " & tmp3(1) & tmp1.getVal(j)
                    End If
                Next i
            Next k
        Next j
        
        Set tmp1 = Nothing
        Set tmp2 = Nothing

    End If
    
    Set solutions1 = res
    Set res = Nothing
    
End Function

Sub demoSolutions1()
    
    
    Dim l As New Lists
    
    With solutions1(l.of(5, 8, 3, 4))
        .p
        Debug.Print .dict(-5)
    End With
    
End Sub
