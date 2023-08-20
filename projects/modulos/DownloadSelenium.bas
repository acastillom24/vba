Attribute VB_Name = "DownloadSelenium"
Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function googleChromeLabs(Optional user$ = "")

    Dim solicitud As New MSXML2.XMLHTTP60
    Dim respuesta$
    Dim proveedor As Variant
    Dim Json As Dictionary
    Dim milestones As Object
    Dim version As Object
    Dim chromeVersion$
    Dim url$
    
    solicitud.Open "GET", "https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone.json", False
    solicitud.Send ("")
    respuesta = solicitud.responseText
    
    Set Json = JsonConverter.ParseJson(respuesta)
    Set milestones = Json("milestones")
    Set version = milestones(getChromeVersion())
    
    chromeVersion = version("version")
    url = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/" & chromeVersion & "/win32/chromedriver-win32.zip"
    Call downloadZip(url)
    
End Function

Private Function getChromeVersion() As String
    
    Dim regKey$
    
    On Error Resume Next
    regKey = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\Version")
    On Error GoTo 0
    
    If Not regKey = "" Then
        getChromeVersion = Mid(regKey, 1, 3)
    Else
        getChromeVersion = "Error"
    End If
    
End Function

Private Function downloadZip(url$)
    Dim zipPath$
    Dim folderPath$
    Dim resultado As Long
    Dim userName$

    userName = Environ("USERNAME")
    
    ' Ruta donde guardar el archivo descargado
    zipPath = "C:\Users\" & userName & "\AppData\Local\SeleniumBasic\chromedriver-win32.zip"
    folderPath = dirname(zipPath)
    
    ' Descargar el archivo
    resultado = URLDownloadToFile(0, url, zipPath, 0, 0)

    If resultado = 0 Then
        Call unzippingZipFiles(zipPath, folderPath)
    Else
        MsgBox "Error al descargar el archivo. Código: " & resultado
    End If
    
End Function

Private Function dirname(path$) As String
    Dim posicionUltimaBarra As Long
    
    ' Encontrar la posición de la última barra
    posicionUltimaBarra = InStrRev(path, "\")
    
    ' Obtener la subcadena que contiene la ruta sin el nombre del archivo
    If posicionUltimaBarra > 0 Then
        dirname = Left(path, posicionUltimaBarra - 1)
    Else
        dirname = ""
    End If
End Function

Private Function unzippingZipFiles(zipFilePath$, folderPath$)
    Dim objShell As Object
    Dim objFolder As Object
    Dim objItem As Object
    Dim objItems As Object
    
    Call deleteFile(folderPath & "\LICENSE.chromedriver")
    Call deleteFile(folderPath & "\chromedriver.exe")
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(folderPath & "\")
    
    ' Obtener referencia a la carpeta dentro del archivo ZIP
    Set objItems = objShell.Namespace(zipFilePath & "\chromedriver-win32").Items

    ' Extraer los elementos dentro de la carpeta ZIP en la carpeta de destino
    For Each objItem In objItems
        objFolder.CopyHere objItem, 4
    Next objItem
    
    ' Liberar los objetos
    Set objItems = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing
    
    Call deleteFile(zipFilePath)
End Function

Private Function deleteFile(filePath$) As Boolean
    ' Verificar si el archivo existe antes de eliminarlo
    If Dir(filePath) <> "" Then
        Kill filePath
        deleteFile = True
    Else
        deleteFile = False
    End If
End Function
