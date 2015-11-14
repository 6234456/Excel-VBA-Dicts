Option Explicit

Private Sub testLoadStructAndClone()
    
    Dim d As New Dicts
    Dim k
    Call d.loadStruct("", 1, 2, d.rng(3, 4))
    
    Dim c As Dicts
    Set c = d.clone()
    
    c.dict("aaa").dict("1") = Array(0, 0)
    
    
    For Each k In c.Keys
        c.dict(k).p
    Next k
    
    Debug.Print String(20, "_")
    
     For Each k In c.Keys
        d.dict(k).p
    Next k
End Sub
