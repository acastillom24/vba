Private Function getUrl(driver As Selenium.ChromeDriver, tag_categoria As String) As String
    
    'devuelve la url de la busqueda a realizar
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   tag_categoria (String): nombre de la busqueda
    '
    'Returns:
    '    String: url de la busqueda a realizar
    '

    Dim xpath_categoria As String

    xpath_categoria = "//strong[contains(text(), '" & tag_categoria & "')]/ancestor::a"
    getUrl = driver.FindElementByXPath(xpath_categoria).Attribute("href")
End Function