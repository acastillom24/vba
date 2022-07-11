Private Sub detalleSiniestro(driver As Selenium.ChromeDriver)

    'extraer la información del detalle del siniestro y validar su ejecución
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '
    'Returns:
    '    EMPTY
    '

    Dim arr_ofertas As Variant, lr As Long, i As Long, elemento As Selenium.WebElement
    Dim element_tr As Selenium.WebElement, img As Selenium.WebElement
    Dim xpath_detalle As String, xpath_img As String, placa As String
    
    lr = shOfertasVendidas.Range("A" & Rows.Count).End(xlUp).Row
    arr_ofertas = shOfertasVendidas.Range("A2:M" & lr)
    
    For i = LBound(arr_ofertas) To UBound(arr_ofertas)
    
        If arr_ofertas(i, 13) <> "ok" Then
            driver.Get arr_ofertas(i, 2)
            Application.Wait (Now + TimeValue("0:00:04")) 'Pausa de 4 segundos
            call detalleSiniestroInfo(driver, "Oferta Vendida")
        End If
        shOfertasVendidas.Range("M" & (i + 1)) = shDetalle.Range("C2") 'Placa
        shOfertasVendidas.Range("N" & (i + 1)) = "ok"
    Next i
    
    lr = shOfertasDesiertas.Range("A" & Rows.Count).End(xlUp).Row
    arr_ofertas = shOfertasDesiertas.Range("A2:J" & lr)
    
    For i = LBound(arr_ofertas) To UBound(arr_ofertas)
        If arr_ofertas(i, 10) <> "ok" Then
            driver.Get arr_ofertas(i, 2)
            Application.Wait (Now + TimeValue("0:00:05")) 'Pausa de 5 segundos
            call detalleSiniestroInfo(driver, "Oferta Desierta")
        End If
        shOfertasDesiertas.Range("J" & (i + 1)) = shDetalle.Range("C2") 'Placa
        shOfertasDesiertas.Range("K" & (i + 1)) = "ok"
    Next i
End Sub

Private Sub detalleSiniestroInfo(driver As Selenium.ChromeDriver, tipoOferta As String)

    'extrae la información del detalle del siniestro
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   tipoOferta (String): identifica el tipo de oferta
    '
    'Returns:
    '    EMPTY
    '

    shDetalle.Rows(2).Insert , xlFormatFromRightOrBelow 'Insertamos una fila con el formato inferior
    shDetalle.Range("A2") = driver.FindElementByXPath("//tbody/tr[1]/td[2]").Text 'Siniestro
    shDetalle.Range("B2") = driver.FindElementByXPath("//tbody/tr[1]/td[4]").Text 'Poliza
    placa = driver.FindElementByXPath("//tbody/tr[1]/td[6]").Text 'Placa
    If Not placa = "" Then
        shDetalle.Range("C2") = placa
        Else:
            shDetalle.Range("C2") = arr_ofertas(i, 4)
    End If
    shDetalle.Range("D2") = driver.FindElementByXPath("//tbody/tr[2]/td[2]").Text 'Marca
    shDetalle.Range("E2") = driver.FindElementByXPath("//tbody/tr[2]/td[4]").Text 'Modelo
    shDetalle.Range("F2") = driver.FindElementByXPath("//tbody/tr[2]/td[6]").Text 'Año
    shDetalle.Range("G2") = driver.FindElementByXPath("//tbody/tr[3]/td[2]").Text 'Taller
    shDetalle.Range("H2") = tipoOferta 'Tipo de oferta
    
    xpath_img = "//ul/descendant::img[contains(@src,'.jpg') or contains(@src,'.jpeg')]"
    For Each img In driver.FindElementsByXPath(xpath_img)
        shUrlImg.Rows(2).Insert , xlFormatFromRightOrBelow 'Insertamos una fila con el formato inferior
        shUrlImg.Range("A2") = shDetalle.Range("C2") 'Placa
        shUrlImg.Range("B2") = img.Attribute("src") 'Url de la imagen
    Next img
End Sub