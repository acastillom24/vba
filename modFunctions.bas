Attribute VB_Name = "modFunctions"
Option Explicit
Option Base 1

Public Function establishConnection() As Boolean

    Dim Solicitud As New MSXML2.XMLHTTP60
    Dim url$
    
    url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRazonSoc&razSoc=Alin"
    
    Solicitud.Open "POST", url, False
    Solicitud.setrequestheader "Content-type", "application/x-www-form-urlencoded"
    Solicitud.send ""
    Set Solicitud = Nothing
    
    establishConnection = True
    
End Function

Private Function getRandom() As String

    Dim Solicitud As New MSXML2.XMLHTTP60
    Dim url$
    Dim respuesta$
    Dim HTML As Object
    Dim numRndInput As Object
    
    url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRazonSoc&razSoc=Alin"
    
    Solicitud.Open "POST", url, False
    Solicitud.setrequestheader "Content-type", "application/x-www-form-urlencoded"
    Solicitud.send ""
    respuesta = Solicitud.ResponseText
    
    Set HTML = CreateObject("HTMLFILE")
    HTML.body.innerHTML = respuesta
    
    Set numRndInput = HTML.getElementsByName("numRnd")(0)
    
    If Not numRndInput Is Nothing Then
        getRandom = numRndInput.value
    Else
        getRandom = ""
    End If
    
    Set HTML = Nothing
    Set numRndInput = Nothing
    Set Solicitud = Nothing

End Function

Sub GuardarStringEnTXT(textToSave$)
    Dim filePath As String
    Dim fileNumber As Integer
    
    ' Ruta del archivo
    filePath = "D:\Alin-Castillo\GitHub\vba\projects\main.html" ' Cambia esta ruta a la ubicación que desees
    
    ' Abrir el archivo para escribir
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    
    ' Escribir el contenido en el archivo
    Print #fileNumber, textToSave
    
    ' Cerrar el archivo
    Close #fileNumber
    
    MsgBox "Archivo guardado correctamente."
End Sub

Public Function consultaRuc(codDocum$) As Boolean

    Dim Solicitud As New MSXML2.XMLHTTP60
    Dim url$, respuesta$
    Dim HTML As Object
    Dim random$
    
    If validateRuc(codDocum) Then
    
        random = getRandom
        
        If random <> "" Then
            Call getDatos(codDocum, random)

        End If
    
    End If

End Function

Private Function getDatos(codDocum$, random$)
    Dim Solicitud As New MSXML2.XMLHTTP60
    Dim url$, respuesta$
    Dim HTML As New HTMLDocument
    Dim divElementos As Object
    Dim divElemento As Object
    Dim Json As Object
    Dim arr As Variant
    Dim i As Long
    
    url = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&actReturn=1&modo=1&nroRuc="
    
    Solicitud.Open "POST", url & codDocum & "&numRnd=" & random, False
    Solicitud.setrequestheader "Content-type", "application/x-www-form-urlencoded"
    Solicitud.send ""
    respuesta = Solicitud.ResponseText
    respuesta = Trim(respuesta)
    Set Solicitud = Nothing
    
    HTML.body.innerHTML = Trim(respuesta)

    Set divElementos = HTML.querySelectorAll(".list-group-item-heading")
    
    ReDim arr(divElementos.Length)
    i = 1
    
    For Each divElemento In divElementos
        arr(i) = divElemento.innerText
        i = i + 1
        
        If i > divElementos.Length Then
            Exit For
        End If
        
    Next divElemento
    
    Set HTML = Nothing
    Set Solicitud = Nothing
        

End Function

Private Function validateRuc(codRuc$) As Boolean
    
    Dim list As Variant
    list = Array(10, 15, 17, 20)
    validateRuc = True
    If Len(codRuc) <> 11 Or Not inList(Val(Left(codRuc, 2)), list) Or Not algoritmoValidarRuc(codRuc) Then
        validateRuc = False
    End If
    
End Function

Private Function inList(value As Variant, list As Variant) As Boolean
    
    Dim el As Variant
    inList = False
    For Each el In list
        If value = el Then
            inList = True
            Exit Function
        End If
    Next el

End Function

Private Function algoritmoValidarRuc(codRuc) As Boolean
    Dim suma As Integer
    Dim resto As Integer
    Dim complemento As Byte
    
    algoritmoValidarRuc = False
    
    suma = Val(Mid(codRuc, 1, 1)) * 5 + Val(Mid(codRuc, 2, 1)) * 4 + _
        Val(Mid(codRuc, 3, 1)) * 3 + Val(Mid(codRuc, 4, 1)) * 2 + _
        Val(Mid(codRuc, 5, 1)) * 7 + Val(Mid(codRuc, 6, 1)) * 6 + _
        Val(Mid(codRuc, 7, 1)) * 5 + Val(Mid(codRuc, 8, 1)) * 4 + _
        Val(Mid(codRuc, 9, 1)) * 3 + Val(Mid(codRuc, 10, 1)) * 2
        
    resto = suma Mod 11
    
    complemento = IIf(resto = 1, 0, Val(Left(11 - resto, 1)))
    
    If Val(Mid(codRuc, 11, 1)) = complemento Then
        algoritmoValidarRuc = True
        Exit Function
    End If
    
End Function

Private Function validateDni(codDni$) As Variant
    Dim codRuc$
    Dim suma As Integer
    Dim resto As Integer
    Dim complemento As Byte
    
    If Len(codRuc) <> 8 Then
        validateDni = False
        Exit Function
    End If
    
    codRuc = 10 & codDni
    suma = Val(Mid(codRuc, 1, 1)) * 5 + Val(Mid(codRuc, 2, 1)) * 4 + _
    Val(Mid(codRuc, 3, 1)) * 3 + Val(Mid(codRuc, 4, 1)) * 2 + _
    Val(Mid(codRuc, 5, 1)) * 7 + Val(Mid(codRuc, 6, 1)) * 6 + _
    Val(Mid(codRuc, 7, 1)) * 5 + Val(Mid(codRuc, 8, 1)) * 4 + _
    Val(Mid(codRuc, 9, 1)) * 3 + Val(Mid(codRuc, 10, 1)) * 2
    resto = suma Mod 11
    complemento = IIf(resto = 1, 0, Val(Left(11 - resto, 1)))
    
    validateDni = "10" & codDni & complemento
    
End Function
