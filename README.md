# Visual Basic for Applications

Proyectos realizados con el lenguaje vba, el cual viene integrado en la aplicaciones de office.

## Recursos

- [wiseowl.co.uk](https://www.wiseowl.co.uk/)
- [excelcute.com](https://excelcute.com/)
- [guru99.com](https://www.guru99.com/)

## Selenium

### Instalación

- [Selenium Basic](https://github.com/florentbr/SeleniumBasic/releases)
- [ChromeDriver](https://sites.google.com/chromium.org/driver/)
- [Install the .NET Framework 3.5](https://docs.microsoft.com/en-us/dotnet/framework/install/dotnet-35-windows)

### Información

🌍

- [Introducción a Selenium Testing](https://qalified.com/introduccion-a-selenium-testing/)
- [Chromium Code Search](https://source.chromium.org/)
- [List of Chromium Command Line Switches](https://peter.sh/experiments/chromium-command-line-switches/)
- [chrome_switches.cc](https://chromium.googlesource.com/chromium/src/+/master/chrome/common/chrome_switches.cc)
- [headless_shell_switches.cc](https://chromium.googlesource.com/chromium/src/+/master/headless/app/headless_shell_switches.cc)
- [pref_names.cc](https://chromium.googlesource.com/chromium/src/+/master/chrome/common/pref_names.cc)
- [Cómo hacer data scraping con VBA y Selenium](https://excelcute.com/vba-data-scraping-selenium/)
- [Using Excel VBA and Selenium](https://www.guru99.com/excel-vba-selenium.html)
- [XPath axes](https://jrebecchi.github.io/xpath-helper/xpath-axes.html)
- [Window.scrollBy()](https://developer.mozilla.org/es/docs/Web/API/Window/scrollBy)
- [Window.scroll()](https://developer.mozilla.org/es/docs/Web/API/Window/scroll)
- [Window.scrollTo()](https://developer.mozilla.org/es/docs/Web/API/Window/scrollTo)

🎬

- [Excel VBA Introduction Part 57.1 - Getting Started with Selenium Basic and Google Chrome](https://www.youtube.com/watch?v=FoxWcvZzYVk)
- [Excel VBA Introduction Part 57.2 - Basic Web Scraping with Selenium and Google Chrome](https://www.youtube.com/watch?v=y7yWL0oCB3k)
- [Excel VBA Introduction Part 57.3 - Using Different Web Browsers with Selenium](https://www.youtube.com/watch?v=qxNx12RWihU)
- [Excel VBA Introduction Part 57.4 - Finding Web Elements in Selenium](https://www.youtube.com/watch?v=lr7CFZEI2YA&t=825s)
- [Excel VBA Introduction Part 57.5 - Implicit and Explicit Waits in Selenium](https://www.youtube.com/watch?v=ii1LxfEfY44)
- [Excel VBA Introduction Part 57.6 - Working with Multiple Tabs in Selenium](https://www.youtube.com/watch?v=_IlkdRwgIwg)
- [Excel VBA Introduction Part 57.7 - Using Select Drop Down Lists in Selenium](https://www.youtube.com/watch?v=-kjq_8i9buM)
- [Excel VBA Introduction Part 57.8 - Printing in Google Chrome using Selenium](https://www.youtube.com/watch?v=jEYvgU46gmE)
- [Scroll down a web page in Chrome with Selenium for VBA](https://www.youtube.com/watch?v=s3Bxb0wthqI)

## Ribbon

## Ribbon y Backstage

### Información

🎬

- [Cómo programar Excel Ribbon y Backstage con XML y VBA](https://www.youtube.com/watch?v=vKH13g4Xmb4)

## Proyectos

- [👇Ripley Puntos](https://github.com/acastillom24/vba/raw/main/web-scraping/ripley-puntos.xlsm): Sirve obtener la información de los productos que puedes obtener con tus puntos ripley.

## Funciones

### Guardar texto

```vb
Private Function saveString(textToSave$)
    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "[RUTA_DEL_ARCHIVO]"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, textToSave
    Close #fileNumber

    MsgBox "Archivo guardado correctamente."
End Function
```

### Generar un Status Bar

```vb
Private Function BarraDeProgreso()
    Dim i As Long
    Dim max As Long

    max = 100

    For i = 1 To max
        Application.StatusBar = "Progreso: [" & _
        String(i, ChrW(9608)) & String(max - i, " ") & "] " & _
        Format(i / max, "0%")
        Application.Wait Now + TimeValue("0:00:01")
    Next i

    Application.StatusBar = False
End Function
```

### Mover el archivo mas reciente según su extensión y opcionalmente renombrarlo.

```vb
' Referencia: Microsoft Scripting Runtime
' Info From: https://www.automateexcel.com/vba/move-files/

Private Function MoveFiles( _
    FromPath As String _
    , ToPath As String _
    , FileExt As String _
    , Optional NewName As String = "") As Boolean

    Dim fso As Scripting.FileSystemObject
    Dim FileInFromFolder As Object
    Dim fechaMasReciente As Date
    Dim FromFile As String

    Set fso = New Scripting.FileSystemObject
    fechaMasReciente = DateValue("01/01/1900")

    If ToPath = "" Then
        MsgBox "Debe ingresar una dirección de una carpeta valida"
        MoveFiles = False
        Exit Function
    End If

    ' Crea una carpeta en caso de que esta no exista
    If Not fso.FolderExists(ToPath) Then
        fso.CreateFolder ToPath
    End If

    ' Determina el archivo mas reciente
    For Each FileInFromFolder In fso.GetFolder(FromPath).Files
        If LCase(Right(FileInFromFolder.Name, 4)) = FileExt Then
            If fechaMasReciente < FileInFromFolder.DateLastModified Then
                FromFile = FileInFromFolder.path
                fechaMasReciente = FileInFromFolder.DateLastModified
            End If
        End If
    Next FileInFromFolder

    ' Mueve el archivo en caso este exista
    If fso.FileExists(FromFile) Then
        If NewName <> "" Then
            Name FromFile As ToPath & "\" & NewName & FileExt
        Else:
            fso.MoveFile _
                Source:=FromFile, _
                Destination:=ToPath
        End If
    End If

    Set fso = Nothing
    MoveFiles = True

End Function
```

### Cargar variables de archivos DotEnv

```vb
' Referencia: Microsoft Scripting Runtime

Dim FSO As FileSystemObject
Dim envDict As Scripting.Dictionary

Public Sub LoadEnv()

    'Initialize FileSystemobject and Dictionary objects
    Set FSO = New FileSystemObject
    Set envDict = New Scripting.Dictionary

    'Define path to .env file
    Dim envPath As String
    envPath = ThisWorkbook.Path & "\.env"

    'Check if .env file exists
    If Not FSO.FileExists(envPath) Then
        MsgBox "Could not find .env file"
        Exit Sub
    End If

    'Read .env file and add variables to dictionary
    Dim envFile As TextStream
    Set envFile = FSO.OpenTextFile(envPath, ForReading)
    Do Until envFile.AtEndOfStream
        Dim line As String
        line = envFile.ReadLine
        If InStr(line, "=") > 0 Then
            Dim parts() As String
            parts = Split(line, "=")
            'Delete single and doubles quotes
            envDict(parts(0)) = Replace(Replace(parts(1), "'", ""), """", "")
        End If
    Loop
    envFile.Close
End Sub

'Get value of environment variable
Private Function GetEnv(key As String) As Variant
    If envDict.Exists(key) Then
        GetEnv = envDict(key)
    Else
        GetEnv = Null
    End If
End Function
```

## Ejecuciones por consola

- Crea un libro de trabajo habilitado para macros

```batch
start excel /m
```

[🖱Más comandos](https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6#Category=Excel)
