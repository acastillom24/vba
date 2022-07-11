Private Sub downloadedFilePDF(driver As Selenium.ChromeDriver, idFichero As String)
    
    'descarga los pdfs a partir de las urls
    '
    'Args:
    '   driver (ChromeDriver): driver con la configuración inicial
    '   idFichero (String): identificador para las urls de los pdfs
    '
    'Returns:
    '    EMPTY
    '
    
    Dim lr As Long, arr() As Variant, i As Long, j As Long, urlFichero As String
    
    'Descarga del pdf
    lr = shOfertasVendidas.Range("A" & Rows.Count).End(xlUp).Row
    arr = shOfertasVendidas.Range("L2:M" & lr)
    
    j = 1
    For i = LBound(arr) To UBound(arr)
        If arr(i, 2) = idFichero And arr(i, 1) <> "" Then
            driver.Get arr(i, 1)
            j = j + 1
        End If
    Next i
    
    If j < 2 Then
        MsgBox "La placa ingresada no cuenta con archivo de adjudicación"
    End If
    
End Sub