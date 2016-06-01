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
