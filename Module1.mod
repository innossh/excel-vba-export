Sub Export()

    Dim line As String, i As Long, j As Long
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then
                Dim moduleName As String, outputFilePath As String
                moduleName = .VBComponents(i).Name & ".mod"
                outputFilePath = ActiveWorkbook.Path & ":" & moduleName
                Dim outputFile
                outputFile = FreeFile
                Open outputFilePath For Output As #outputFile
                With .VBComponents(i).CodeModule
                    For j = 1 To .CountOfLines
                        line = .Lines(j, 1)
                        Print #outputFile, line
                    Next j
                End With
                Close #outputFile
            End If
        Next i
    End With

End Sub
