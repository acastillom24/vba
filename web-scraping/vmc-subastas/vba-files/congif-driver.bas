
Private Function configDriver(url As String, path As String) As ChromeDriver

    'Configuración del driver con el punto de partida y el directorio
    '
    'Args:
    '   url (String): dirección web para establecer el punto de partida del driver
    '   path (String): dirección del directorio donde se quiere descargar los archivos
    '
    'Returns:
    '    ChromeDriver: objeto chrome driver con las configuraciones
    '

    Dim driver As Selenium.ChromeDriver
    Set driver = New Selenium.ChromeDriver 'Establecemos el driver

    'Configuraciones del driver
    driver.AddArgument "--start-maximized"
    driver.AddArgument "--disable-gpu"
    driver.AddArgument "no-sandbox"
    'driver.AddArgument "--headless"

    If path <> "" then
        driver.SetPreference "download.default_directory", path
        driver.SetPreference "download.directory_upgrade", True
        driver.SetPreference "download.prompt_for_download", False
    End if

    driver.Start "chrome", url 'Pagina de inicio del driver

    Set configDriver = driver
End Function