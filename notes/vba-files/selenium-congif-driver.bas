
Private Function configDriver(url As String) As ChromeDriver

    ''Configuración del driver con la dirección web de partida
    '
    'Args:
    '   url (String): dirección web para establecer el punto de partida del driver
    '
    'Returns:
    '    ChromeDriver: objeto chrome driver con las configuraciones
    '

    Dim driver As Selenium.ChromeDriver
    Set driver = New Selenium.ChromeDriver 'Establecemos el driver

    'Configuraciones del driver
    driver.AddArgument "--incognito" 'Abrir el navegador en modo incognito
    driver.AddArgument "--start-maximized" 'Maximizar la ventana del navegador
    driver.AddArgument "--window-size=1280,720" 'Personalizar el tamaño de la ventana del navegador
    driver.AddArgument "--disable-gpu" 'Desabilitar la aceleración del hardware GPU
    driver.AddArgument "--no-sandbox" 'Desabilita la ejecución segura de los procesos
    driver.AddArgument "--headless" 'Ejecutar el navegador sin ninguna interfaz de usuario

    driver.Start "chrome", baseUrl:=url 'Pagina de inicio del driver

    Set configDriver = driver
End Function

Private Function configDriver(url As String, path As String) As ChromeDriver

    'Configuración del driver con la dirección web de partida y la dirección del directorio local
    '
    'Args:
    '   url (String): dirección web para establecer el punto de partida del driver
    '
    'Returns:
    '    ChromeDriver: objeto chrome driver con las configuraciones
    '

    Dim driver As Selenium.ChromeDriver
    Set driver = New Selenium.ChromeDriver 'Establecemos el driver

    'Configuraciones del driver
    driver.AddArgument "--incognito" 'Abrir el navegador en modo incognito
    driver.AddArgument "--start-maximized" 'Maximizar la ventana del navegador
    driver.AddArgument "--window-size=1280,720" 'Personalizar el tamaño de la ventana del navegador
    driver.AddArgument "--disable-gpu" 'Desabilitar la aceleración del hardware GPU
    driver.AddArgument "--no-sandbox" 'Desabilita la ejecución segura de los procesos
    driver.AddArgument "--headless" 'Ejecutar el navegador sin ninguna interfaz de usuario

    If path <> "" then
        driver.SetPreference "download.default_directory", path
        driver.SetPreference "download.directory_upgrade", True
        driver.SetPreference "download.prompt_for_download", False
    End if

    driver.Start "chrome", baseUrl:=url 'Pagina de inicio del driver

    Set configDriver = driver
End Function