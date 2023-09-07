# Download-ChromeDriver

El modulo **DownloadSelenium** permite descargar el chrome driver compatible con la versión actual de su google chrome.

## Ejemplos

```vb
Public Sub initSelenium()
    Dim url As String
    Dim driver As Selenium.ChromeDriver
    
    url = "https://www.google.com.pe"
    Set driver = configDriver(url)
    driver.get "/?hl=es"
    
    driver.Quit
        
End Sub
```

```vb
Private Function configDriver(url As String) As Selenium.ChromeDriver
    Dim driver As Selenium.ChromeDriver
    
    'Establecemos el driver
    Set driver = New Selenium.ChromeDriver

    'Configuraciones del driver
    driver.AddArgument "--start-maximized" 'Personalizar el tamaño de la ventana del navegador
    driver.AddArgument "--disable-gpu" 'Deshabilitar la aceleración del hardware GPU
    driver.AddArgument "--no-sandbox" 'Deshabilita la ejecución segura de los procesos
    
    On Error GoTo ErrorVersionChromeDriver
    driver.Start "chrome", baseURL:=url 'Pagina de inicio del driver
    Set configDriver = driver
    Exit Function

ErrorVersionChromeDriver:
    If Err.Number = 33 Then
        driver.Quit
        Call DownloadSelenium.googleChromeLabs
        driver.Start "chrome", baseURL:=url 'Pagina de inicio del driver
        Set configDriver = driver
        Else:
            MsgBox "No se ha podido determinar el error.", , "alincastillo1995@gmail.com"
    End If
    
End Function
```

## Opciones

Download-ChromeDriver incluye opciones para evitar errores inesperados:
- __user__ (Default = `""`) Asignar el nombre usuario actual.


## Instalación

1. Importe `DownloadSelenium.bas` a su proyecto (Abra el editor de VBA, `Alt + F11`; Archivo > Importar Archivo)
2. Añadir la referencia o clase `Dictionary`
   - Incluir la referencia "Microsoft Scripting Runtime"
3. Añadir la referencia o clase `MSXML2`
   - Incluir la referencia "Microsoft XML, v6.0"
4. Incluir el modulo ["JsonConverter"](https://github.com/VBA-tools/VBA-JSON/blob/master/README.md)

## Recursos

Puedes descargar un archivo [selenium.xlsm](https://github.com/acastillom24/vba/raw/develop/projects/selenium.xlsm?download=), el cual ya contiene todas las configuraciones.