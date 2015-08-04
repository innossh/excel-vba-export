Sub Load()

    Dim Menu As Variant, Control As Variant
    Set Menu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup)
    Menu.Caption = "モジュールエクスポート"

    Dim i As Long
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then
                Dim ModuleName As String, outputFilePath As String
                ModuleName = .VBComponents(i).Name & ".mod"
                Set Control = Menu.Controls.Add
                With Control
                    .Caption = ModuleName
                    .OnAction = "Export"
                    .BeginGroup = False
                End With
            End If
        Next i
    End With

End Sub

Sub Export()

    Dim line As String, i As Long, j As Long
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then
                Dim ModuleName As String, outputFilePath As String
                ModuleName = .VBComponents(i).Name & ".mod"
                outputFilePath = ActiveWorkbook.Path & ":" & ModuleName
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
