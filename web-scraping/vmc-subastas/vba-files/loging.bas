Private Function loging(driver As Selenium.ChromeDriver) As Boolean

    'logearnos en la web a la que se quiere acceder
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '
    'Returns:
    '    Boolean: verifica si se accedio a la web
    '

    Dim By As New Selenium.By, connection As Boolean
    
    driver.Get "/4panel/login.php" 'web de logeo
    
    driver.FindElementById("username").SendKeys "mapfreperu" 'User
    driver.FindElementById("password").SendKeys "crgper7i" 'Pass
    driver.FindElementById("btn_ingresar_sist").Click
    Application.Wait (Now + TimeValue("0:00:05")) 'pausa de 5 segundos
    
    'Validar el ingreso
    connection = TRUE
    If Not driver.IsElementPresent(By.Class("banner-index")) Then
        connection = FALSE 
    End If
    loging = connection

End Function