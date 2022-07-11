Public Sub mainVMC(nOption As Integer, Optional fecha_inicial As Date, Optional fecha_final As Date, Optional idFichero As String)

    'llama a las distintas funciones de acuerdo al procedimiento requerido
    '
    'Args:
    '   nOption (Integer): número que identifica la opción a ejecutar dentro del procedimiento
    '   fecha_inicial (String): fecha de partida de la ejecución del scrapeo
    '   fecha_final (String): fecha fin de la ejecución del scrapeo
    '   idFichero (String): el identificador para la descarga de ficheros
    '
    'Returns:
    '   EMPTY
    '

    Dim webVMC As Selenium.ChromeDriver, path As String

    Select Case nOption
     Case Is = 1
        Set webVMC = configDriver("https://4panel.vmcsubastas.com", "")

        if loging(webVMC) then
            webVMC.Get getUrl(webVMC, "Historial")
            call stopRun(3)
            Call historialOfertas(webVMC, fecha_inicial, fecha_final)
            Call detalleOferta(webVMC)
            Call detalleSiniestro(webVMC)

        else:
            MsgBox "No se ha podido establecer la conexción de logeo",,"Warning" 
            exit sub
        end if

     Case Is = 2
        path = setPathDirectory
        nOption = findID(idFichero)
        Select Case nOption
         Case Is = 0
            MsgBox "¡Código de la placa no encontrado!"
            exit sub
         Case Is = 1 'Descarga ofertas vendidas
            Set webVMC = configDriver("https://4panel.vmcsubastas.com", path)
            call stopRun(1)
            
            if loging(webVMC) then
                Call downloadedFilePDF(webVMC, idFichero)
                Call downloadedFileIMG(idFichero, path)
            else:
                MsgBox "No se ha podido establecer la conexción de logeo",,"Warning" 
                exit sub
            end if            

         Case Is = 2 'Descarga ofertas desiertas
            Set webVMC = configDriver("https://4panel.vmcsubastas.com", path)
            call stopRun(1)

            if loging(webVMC) then
                Call downloadedFileIMG(idFichero, path)
            else:
                MsgBox "No se ha podido establecer la conexción de logeo",,"Warning" 
                exit sub
            end if 
        End Select
    End Select

    webVMC.Close 'Cierre del driver
End Sub

Private Function findID(id As String) As Integer

    'valida si se encuentra el id para la descarga de ficheros
    '
    'Args:
    '   id (String): identificador de la placa
    '
    'Returns:
    '    Integer: valor que identica el tipo de oferta
    '

    Dim arr() As Variant, lr As Long, i As Long, tipoOferta As String
    
    lr = shDetalle.Range("C" & Rows.Count).End(xlUp).Row
    arr = shDetalle.Range("C2:H" & lr)
    
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = id Then
            tipoOferta = arr(i, 6)
        End If
    Next i
    
    If tipoOferta <> "" Then
        If tipoOferta = "Oferta Vendida" Then
            findID = 1
            ElseIf tipoOferta = "Oferta Desierta" Then
                findID = 2
        End If
        Else:
            findID = 0
    End If
    
End Function