Sub Export()

    Dim targetModule, outputFilePath, outputFile
    targetModule = "Module1"
    outputFilePath = ActiveWorkbook.Path & ":" & targetModule & ".mod"
    outputFile = FreeFile
    Open outputFilePath For Output As #outputFile
    Dim line As String, i As Long
    With ThisWorkbook.VBProject.VBComponents(targetModule).CodeModule
        For i = 1 To .CountOfLines
            line = .Lines(i, 1)
            Print #outputFile, line
        Next i
    End With
    Close #outputFile

End Sub
