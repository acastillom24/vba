Private Function renameFile()
    
    Const DOWNLOAD_DIRECTORY As String = "C:\Temp"
    Const FILE_NAME As String = "myNewCsv.csv"
    
    Dim fso As Object, myFolder As Object, filename As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set myFolder = fso.GetFolder(DOWNLOAD_DIRECTORY)

    Dim objFile As Object, dteFile As Date

    'dteFile = DateSerial(2022, 7, 4)
    dteFile = Date
    For Each objFile In myFolder.Files
        Debug.Print DateValue(objFile.DateLastModified) & objFile.Name
        Debug.Print dteFile
        If objFile.DateLastModified > dteFile And fso.GetExtensionName(objFile.path) = "pdf" Then
            dteFile = objFile.DateLastModified
            filename = objFile.Name
        End If
    Next objFile
    
    Debug.Print fso.FileExists(DOWNLOAD_DIRECTORY & "\" & FILE_NAME)
    
    If filename <> vbNullString And Not fso.FileExists(DOWNLOAD_DIRECTORY & "\" & FILE_NAME) Then
       fso.MoveFile DOWNLOAD_DIRECTORY & "\" & filename, DOWNLOAD_DIRECTORY & "\" & FILE_NAME
    End If
    
End Function