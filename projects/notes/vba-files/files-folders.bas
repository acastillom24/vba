Public  Sub Macro()
    dim fso as Scripting.FileSystemObject
    dim path as String, nameFolder As String
    
    set fso = new Scripting.FileSystemObject
    path = "D:\Proyectos\GIT\vba\notes\"
    nameFolder = "Nueva Carpeta"

    if fso.FolderExists(nameFolder) Then 
    end if
    fso.CreateFolder path & nameFolder
End Sub