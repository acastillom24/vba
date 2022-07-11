Private Sub historialOfertas(driver As Selenium.ChromeDriver, fecha_inicial As Date, fecha_final As Date)

    'valida y extrae los datos del detalle general de las subastas
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   fecha_inicial (Date): fecha incial
    '   fecha_final (Date): fecha final
    '
    'Returns:
    '    NULL
    '

    Dim page As Integer, xpath_siguiente As String, siguiente As Boolean, arr_historial() As Variant
    Dim i As Long, n As Integer

    xpath_siguiente = "//div[@class='paginacion']/a[@title='Siguiente']" 'xpath para la siguiente tabla
    newOferta = 0 'Reiniciar la variable global
    page = 1 'Variable utilizada para la navegación de las paginas

    Do While page = 1 Or siguiente = True
        page = page + 1
        arr_historial = extraerDatosHistoricos(driver, fecha_inicial, fecha_final)

        on On Error Resume Next
        'Set Error Source Macro/Function name 
        n = UBound(arr_historial)
        
        If Err.Number Then
        Else:
            i = UBound(arr_historial)
            if not searchId(arr_historial(i, 1)) Then
                shHistorialOfertas.Rows("2:" & (i + 1)).Insert , xlFormatFromRightOrBelow 'Insertamos una fila con el formato inferior
                shHistorialOfertas.Ramge("A2:K" & (i + 1)) = arr_historial
            end if
        End If
        
        If newOferta = 1 Then
            Exit Do
            Else:
                siguiente = findElementDriver(driver, xpath_siguiente)
                If siguiente Then
                    driver.FindElementByXPath(xpath_siguiente).Click
                End If
        End If
    Loop
End Sub

Private Function extraerDatosHistoricos(driver As Selenium.ChromeDriver, fecha_inicial As Date, fecha_final As Date) As Variant

    'devuelve una matriz con los datos generales del detalle de la subasta
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   fecha_inicial (Date): fecha incial
    '   fecha_final (Date): fecha final
    '
    'Returns:
    '    Variant: matriz con los datos generales
    '

    Dim xpath_tr As String, arr_historico() As Variant, nrow As Long, fecha As Variant
    Dim element_tr As Selenium.WebElement, exp As Object
    
    xpath_tr = "//tbody/tr"
    Set exp = New RegExp
    exp.Pattern = "\D"
    exp.Global = True
    nrow = 1
    
    For Each element_tr In driver.FindElementsByXPath(xpath_tr)
        fecha = element_tr.FindElementByXPath("./td[2]").Text 'Constante para comparar la ejecución del bucle
        
        If fecha < fecha_inicial Then
            newOferta = 1
            Exit For
            ElseIf fecha > fecha_inicial And fecha < fecha_final Then
                If element_tr.FindElementByXPath("./td[9]").Text = "Finalizado" Then
                    
                    If nrow > 1 Then
                        arr_historico = Application.Transpose(arr_historico)
                        ReDim Preserve arr_historico(11, nrow)
                        arr_historico = Application.Transpose(arr_historico)
                        Else:
                            ReDim Preserve arr_historico(nrow, 11)
                    End If
                    
                    arr_historico(nrow, 2) = element_tr.FindElementByXPath("./td[1]").Text 'Grupo
                    arr_historico(nrow, 3) = element_tr.FindElementByXPath("./td[1]/a").Attribute("href") 'Link del grupo
                    arr_historico(nrow, 1) = "VMC_" & exp.Replace(arr_historico(nrow, 3), "") 'Creación de la llave primaría
                    arr_historico(nrow, 4) = element_tr.FindElementByXPath("./td[2]").Text 'Fecha del proceso
                    arr_historico(nrow, 5) = element_tr.FindElementByXPath("./td[3]").Text 'Totoal de ofertas publicadas
                    arr_historico(nrow, 6) = element_tr.FindElementByXPath("./td[4]").Text 'Ofertas concretadas
                    arr_historico(nrow, 7) = element_tr.FindElementByXPath("./td[5]").Text 'Ofertas desiertas
                    arr_historico(nrow, 8) = element_tr.FindElementByXPath("./td[6]").Text 'Ofertas vendidas
                    arr_historico(nrow, 9) = element_tr.FindElementByXPath("./td[7]").Text 'Expectativa de venta
                    arr_historico(nrow, 10) = element_tr.FindElementByXPath("./td[8]").Text 'Recaudación de venta
                    arr_historico(nrow, 11) = element_tr.FindElementByXPath("./td[9]").Text 'Estado
                    
                    nrow = nrow + 1
                End If
        End If
    Next element_tr
    extraerDatosHistoricos = arr_historico
End Function

Private Function searchId(id As As Variant) As Boolean

    'valida si se encuentra una subasta, en caso de que no, se registra
    '
    'Args:
    '   id (String): identificador a buscar
    '
    'Returns:
    '    Boolean: si TRUE el identificador ya se encuentra registrado
    '

    Dim lr As Long, historialOfertas As Variant, i As Long
    
    lr = shHistorialOfertas.Range("A" & Rows.Count).End(xlUp).Row
    historialOfertas = shHistorialOfertas.Range("A2:A" & lr)
    
    For i = LBound(historialOfertas) To UBound(historialOfertas)
        If historialOfertas(i, 1) = id Then
            searchId = True
            Else:
                searchId = False
        End If
    Next i
End Function