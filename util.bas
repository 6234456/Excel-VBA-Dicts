Function shtExists(ByVal name As String, Optional ByRef wb) As Boolean
    On Error GoTo errhandler
    If IsMissing(wb) Then
        Set wb = ActiveWorkbook
    End If
    
    shtExists = Not (wb.Worksheets(name) Is Nothing)
errhandler:
    If Err.Number <> 0 Then
        shtExists = False
    End If
End Function


'return false if the sheet with that name already exists thus not created by the program
'true  a new sheet created
Function createShtIfNotExists(ByVal shtName As String, Optional ByRef wb) As Boolean
    
    Dim res As Boolean
    res = False
    
    If IsMissing(wb) Then
        Set wb = ActiveWorkbook
    End If
    
    With wb
        If Not shtExists(shtName, wb) Then
            .Worksheets.add after:=.Worksheets(.Worksheets.Count)
            .Worksheets(.Worksheets.Count).name = shtName
            
            res = True
        End If
    End With
    
    createShtIfNotExists = res
End Function
