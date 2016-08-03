' loop through the file system
' define the interface of
' sub interface_processWorkbook(byref wb as workbook, byref this as workbook)

Public Sub processWorkbooksInthePath(Optional ByVal path As String = "src", Optional ByVal readOnly As Boolean = True)
    
    On Error GoTo handler

    Application.ScreenUpdating = False
    
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    
    
    Dim targPath As String
    targPath = Trim(ActiveWorkbook.path & "\" & path)
    
    If Right(targPath, 1) = "\" Then
        targPath = Left(targPath, Len(targPath) - 1)
    End If
    
    
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    
    Dim this As Workbook
    Set this = ThisWorkbook
    
    Dim that As Workbook
    
    With re
        .Pattern = "\.xls(m|x)?$"
    End With
    
    
    Dim i As Object
    Dim p As Object
    Dim fName As String
    
    Set p = fso.getfolder(targPath)
    
    For Each i In p.Files
        fName = i.name
        If Left(fName, 1) <> "~" And re.test(fName) And fName <> this.name Then
            Application.Workbooks.Open fName, 0, readOnly
            
            Set that = ActiveWorkbook
            
            Call interface_processWorkbook(that, this)

            that.Close Not readOnly
        End If
    Next i
    
handler:
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "error"
    End If

End Sub


Sub interface_processWorkbook(ByRef wb As Workbook, ByRef this As Workbook)
    Debug.Print wb.Worksheets(1).name
End Sub

' one row or one column
' mergeCells with the same content
Private Sub mergeCells(ByRef rng As Range, Optional ByVal orient As String = "v")
    
    If rng.Cells.Count > 1 Then

        For i = rng.Cells.Count To 1 Step -1
        
            If orient = "v" Then
            
                Set thisC = rng.Cells(i, 1)
                Set nextC = rng.Cells(i - 1, 1)
                
                If i < rng.Cells.Count Then
                   Set prevC = rng.Cells(i + 1, 1)
                End If
             
            Else
                
                Set thisC = rng.Cells(1, i)
                Set nextC = rng.Cells(1, i - 1)
             
                If i < rng.Cells.Count Then
                   Set prevC = rng.Cells(1, i + 1)
                End If
            End If
        
            If i = rng.Cells.Count Then
                Set start = thisC
            ElseIf thisC.Value <> prevC.Value Then
                Set start = thisC
            End If
                
                
            If thisC.Value = nextC.Value Then
                If i = 1 Then
                    Set ende = thisC
                    tmpVal = thisC.Value
                    
                    With Range(start, ende)
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With

                End If
            Else
                Set ende = thisC
                 With Range(start, ende)
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                End With
            End If
        Next i
        
    End If
    

End Sub

Function groupAndSum(ByVal targKeyCol1 As Integer, ByVal targKeyCol2 As Integer, Optional ByVal targValCol, Optional ByVal targRowBegine, Optional ByVal targRowEnd)
    
    If IsMissing(targRowBegine) Then
        targRowBegine = 1
    End If
    
    If IsMissing(targRowEnd) Then
        targRowEnd = Cells(Rows.Count, targKeyCol2).End(xlUp).Row
    End If
    
     If IsMissing(targValCol) Then
        targValCol = targKeyCol2 + 1
    End If
    
    Dim tmpPreviousRow As Integer
    Dim tmpCurrentRow As Integer
    
    tmpPreviousRow = targRowEnd
    tmpCurrentRow = tmpPreviousRow

    
     Do While tmpCurrentRow > targRowBegine
        
        
        tmpCurrentRow = Cells(tmpCurrentRow, targKeyCol1).End(xlUp).Row
        

        Range(Cells(tmpCurrentRow + 1, 1), Cells(tmpPreviousRow, 1)).Rows.Group
        
        If targValCol <> 0 Then
            Cells(tmpCurrentRow, targValCol).Formula = "=SUM(" & Cells(tmpCurrentRow + 1, targValCol).Address(0, 0) & ":" & Cells(tmpPreviousRow, targValCol).Address(0, 0) & ")"
        End If
        
        tmpPreviousRow = tmpCurrentRow - 1
    Loop

End Function

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
