Private Function setPathDirectory() As String

    'Establecer el dírectorio de inicio mediante selección
    '
    'Returns:
    '    String: dirección del directorio
    '

    Dim path As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True: .Show
        path = .SelectedItems(1)
    End With
    
    setPathDirectory = path
    
End Function