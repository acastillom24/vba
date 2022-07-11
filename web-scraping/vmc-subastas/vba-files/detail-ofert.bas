Private Sub detalleOferta(driver As Selenium.ChromeDriver)

    'detalle especifico de cada proceso de subasta
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '
    'Returns:
    '    EMPTY
    '

    Dim arr_ofertas As Variant, lr As Long, i As Long, elemento As Selenium.WebElement, id As String
    
    lr = shHistorialOfertas.Range("A" & Rows.Count).End(xlUp).Row
    arr_ofertas = shHistorialOfertas.Range("A2:L" & lr)
    
    For i = LBound(arr_ofertas) To UBound(arr_ofertas)
        If arr_ofertas(i, 12) <> "ok" Then
            driver.Get arr_ofertas(i, 3)
            Application.Wait (Now + TimeValue("0:00:04")) 'Pausa de 4 segundos
            'Obtener datos de las ofertas vendidas
            If findElementDriver(driver, "//section[@id='block-sale']/descendant::tbody/tr[1]/td[2]") Then
                For Each elemento In driver.FindElementsByXPath("//section[@id='block-sale']/descendant::tbody/tr")
                    shOfertasVendidas.Rows(2).Insert , xlFormatFromRightOrBelow 'Insertamos una fila con el formato inferior
                    id = arr_ofertas(i, 1)
                    shOfertasVendidas.Range("A2:L2").Value = extraerOfertasVendidas(elemento, id)
                Next elemento
            End If
            
            'Obtener datos de las ofertas desiertas
            If findElementDriver(driver, "//section[@id='block-desiertos']/descendant::tbody/tr[1]/td[2]") Then
                For Each elemento In driver.FindElementsByXPath("//section[@id='block-desiertos']/descendant::tbody/tr")
                    shOfertasDesiertas.Rows(2).Insert , xlFormatFromRightOrBelow 'Insertamos una fila con el formato inferior
                    id = arr_ofertas(i, 1)
                    shOfertasDesiertas.Range("A2:I2").Value = extraerOfertasDesiertas(elemento, id)
                Next elemento
            End If
            
            shHistorialOfertas.Range("L" & (i + 1)) = "ok"
            
        End If
    Next i
End Sub

Private Function extraerOfertasVendidas(elemento As Selenium.WebElement, id As String) As Variant

    'extrae el detalle de las ofertas que fueron concretadas
    '
    'Args:
    '   elemento (webElement): contiene elementos html de fila (tr)
    '   id (String): identificador del detalle general de las subastas
    '
    'Returns:
    '    Variant: un arreglo con el detalle de las ofertas concretadas
    '

    Dim arr(1, 12)
    Dim exp As Object
    
    Set exp = New RegExp
    exp.Global = True
    exp.Pattern = "\s"
    
    arr(1, 1) = id
    arr(1, 2) = elemento.FindElementByXPath("./td[2]/a").Attribute("href") 'url del detalle del bien
    arr(1, 3) = elemento.FindElementByXPath("./td[2]/descendant::img").Attribute("src") 'url de la imagen del bien
    arr(1, 4) = elemento.FindElementByXPath("./td[3]").Text 'item (placa / marca / modelo / año)
    arr(1, 5) = elemento.FindElementByXPath("./td[4]").Text 'precio reserva (PR)
    
    If findElement(elemento, "./td[5]/font") Then
        arr(1, 6) = exp.Replace(Replace(elemento.FindElementByXPath("./td[5]").Text, elemento.FindElementByXPath("./td[5]/font").Text, ""), "") 'levantamiento reserva
        Else:
            arr(1, 6) = elemento.FindElementByXPath("./td[5]").Text
    End If

    arr(1, 7) = elemento.FindElementByXPath("./td[6]/descendant::b").Text 'propuesta ganadora
    
    If findElement(elemento, "./td[6]/descendant::span[2]") Then
        arr(1, 8) = elemento.FindElementByXPath("./td[6]/descendant::img/parent::span").Text 'status propuesta
        Else:
            arr(1, 7) = ""
    End If
    arr(1, 9) = elemento.FindElementByXPath("./td[7]").Text 'puesto
    arr(1, 10) = elemento.FindElementByXPath("./td[8]/descendant::b").Text 'estado
    arr(1, 11) = exp.Replace(elemento.FindElementByXPath("./td[9]").Text, " ") 'miembro [Pendiente realizar el split para dividir el tip docum + cod docum]
    If findElement(elemento, "./td[11]/a") Then
        arr(1, 12) = elemento.FindElementByXPath("./td[11]/a").Attribute("data") 'link pdf del miembro
        Else:
            arr(1, 12) = ""
    End If
    extraerOfertasVendidas = arr
End Function

Private Function extraerOfertasDesiertas(elemento As Selenium.WebElement, id As String) As Variant

    'extrae el detalle de las ofertas que no fueron concretadas
    '
    'Args:
    '   elemento (webElement): contiene elementos html de fila (tr)
    '   id (String): identificador del detalle general de las subastas
    '
    'Returns:
    '    Variant: un arreglo con el detalle de las ofertas desiertas
    '

    Dim arr(1, 9)
    Dim exp As Object
    
    Set exp = New RegExp
    exp.Global = True
    exp.Pattern = "\s"
    
    arr(1, 1) = id
    arr(1, 2) = elemento.FindElementByXPath("./td[2]/a").Attribute("href") 'url del detalle del bien
    arr(1, 3) = elemento.FindElementByXPath("./td[2]/descendant::img").Attribute("src") 'url de la imagen del bien
    arr(1, 4) = elemento.FindElementByXPath("./td[3]").Text 'item (placa / marca / modelo / año)
    arr(1, 5) = elemento.FindElementByXPath("./td[4]").Text 'precio reserva (PR)
    arr(1, 6) = exp.Replace(Replace(elemento.FindElementByXPath("./td[5]").Text, elemento.FindElementByXPath("./td[5]/font").Text, ""), "") 'levantamiento reserva
    arr(1, 7) = elemento.FindElementByXPath("./td[6]/descendant::b").Text 'propuesta ganadora
    If findElement(elemento, "./td[6]/descendant::span[2]") Then
        arr(1, 8) = elemento.FindElementByXPath("./td[6]/descendant::img/parent::span").Text 'status propuesta
        Else:
            arr(1, 8) = ""
    End If
    arr(1, 9) = elemento.FindElementByXPath("./td[7]/descendant::b").Text 'estado
    extraerOfertasDesiertas = arr
End Function

Private Function findElementDriver(driver As Selenium.ChromeDriver, xpath As String) As Boolean

    'retorna la existencia o no de un elemento html en función del driver
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   xpath (String): ubicación de la dirección del elemento del html
    '
    'Returns:
    '    Boolean: TRUE si el elemento es encontrado
    '

    Dim By As New Selenium.By
    findElementDriver = driver.IsElementPresent(By.xpath(xpath))
End Function

Private Function findElement(elemento As Selenium.WebElement, xpath As String) As Boolean

    'retorna la existencia o no de un elemento html en función de un elemento html
    '
    'Args:
    '   elemento (WebElement): elemto html
    '   xpath (String): ubicación de la dirección del elemento del html
    '
    'Returns:
    '    Boolean: TRUE si el elemento es encontrado
    '
    Dim By As New Selenium.By
    findElement = elemento.IsElementPresent(By.xpath(xpath))
End Function