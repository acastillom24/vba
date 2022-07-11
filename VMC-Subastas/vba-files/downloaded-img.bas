Private Sub downloadedFileIMG(idFichero As String, pathFichero As String)

    'descarga las imagenes a partir de las urls
    '
    'Args:
    '   idFichero (String): identificador para las urls
    '   pathFichero (String): directorio local para guardar las imagenes
    '
    'Returns:
    '    EMPTY
    '

    Dim lr As Long, arr() As Variant, i As Long, j As Long, urlFichero As String, img As Variant, ext As String
    
    Dim exp As Object
    Set exp = New RegExp
    exp.Global = True
    exp.Pattern = "(.jpg$)|(.jpeg$)"
    
    lr = shUrlImg.Range("A" & Rows.Count).End(xlUp).Row
    arr = shUrlImg.Range("A2:B" & lr)
    
    j = 1
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = idFichero Then
            urlFichero = arr(i, 2)
            For Each img In exp.Execute(urlFichero)
                ext = img.Value
            Next img
            
            Application.Wait (Now + TimeValue("0:00:01")) 'pausa de 1 segundos
            If j < 10 Then
                URLDownloadToFile 0, urlFichero, pathFichero & "\" & idFichero & "_0" & j & ext, 0, 0
                Else:
                    URLDownloadToFile 0, urlFichero, pathFichero & "\" & idFichero & "_" & j & ext, 0, 0
            End If
            j = j + 1
        End If
    Next i
    
    If j < 2 Then
        MsgBox "La placa ingresada no cuenta con imagenes"
    End If
End Sub