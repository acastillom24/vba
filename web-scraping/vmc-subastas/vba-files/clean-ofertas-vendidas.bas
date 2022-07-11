Option Explicit

Private lr As Long, i As Long, Precio As Long, k As Long
Private x As Integer, x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, t As Integer
Private y As Integer, numSheet_tblSales As Integer
Private Item As String, condicion As String, marca As String, codDocumento As String
Private year As Variant
Private Porcentaje As Single

Public Sub limpiar_datos()
    
    numSheet_tblSales = 4
    Call cleanData

    'Determinamos el número de iteraciones
    lr = Hoja2.Range("L" & Rows.Count).End(xlUp).Row

    k = 2
    
    For i = 2 To lr
        condicion = Hoja2.Range("A" & i).Value
        If condicion = "" Then
            'Columna: (Placa / Marca / Modelo / Año)
            Item = Hoja2.Range("C" & i)
            Item = Replace(Item, Chr(160), "")
            Item = Replace(Item, "  ", " ")
            
            t = Len(Item)
            year = Mid(Item, (t - 3))
            
            If IsNumeric(year) Then
                x1 = InStr(1, Item, " ") 'Placa
                
                x2 = InStr(InStr(1, Item, " ") + 1, Item, " ") 'Marca
                marca = Mid(Item, (x1 + 1), ((x2 - x1) - 1))
                If marca = "Mercedes" Or marca = "Alfa" Or marca = "Aston" Or marca = "Land" Then
                    x2 = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Marca
                    x = Len(Item) - Len(Replace(Item, " ", ""))
                    If x = 3 Then
                        x3 = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        ElseIf x = 4 Then
                            x3 = InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        ElseIf x = 5 Then
                            x3 = InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        ElseIf x = 6 Then
                            x3 = InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        ElseIf x = 7 Then
                            x3 = InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        ElseIf x = 8 Then
                            x3 = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                    End If
                    Else:
                        x = Len(Item) - Len(Replace(Item, " ", ""))
                        If x = 3 Then
                            x3 = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                            ElseIf x = 4 Then
                                x3 = InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                            ElseIf x = 5 Then
                                x3 = InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                            ElseIf x = 6 Then
                                x3 = InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                            ElseIf x = 7 Then
                                x3 = InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                            ElseIf x = 8 Then
                                x3 = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Modelo
                        End If
                End If
                
                Call extraer_item(numSheet_tblSales, k, Item, x1, x2, x3)
                Else: Sheets(numSheet_tblSales).Range("A" & k) = Hoja2.Range("C" & i).Value 'Item
            End If
            
            'Columna: Precio de reserva
            Sheets(numSheet_tblSales).Range("E" & k) = Hoja2.Range("D" & i).Value 'Precio
            
            'Columna: Lavantamiento
            Item = Hoja2.Range("E" & i)
            x = Len(Item) - Len(Replace(Item, " ", ""))
            If x > 0 Then
                x1 = InStr(1, Item, " ") 'Precio Levantamiento
                If x = 2 Then
                    x2 = InStr(InStr(1, Item, " ") + 1, Item, " ") 'Porcentaje Levantamiento
                End If
                Call extraer_levantamiento(numSheet_tblSales, k, Item, x1, x2)
                Else:
                    Sheets(numSheet_tblSales).Range("F" & k) = Item 'Extraemos el precio
                    Sheets(numSheet_tblSales).Range("G" & k) = Item 'Extraemos el porcentaje
                    Sheets(numSheet_tblSales).Range("H" & k) = Item 'Extraemos el id
            End If
            
            'Columna: propuesta Ganadora
            Item = Hoja2.Range("F" & i)
            Item = Replace(Item, "                        ", " ")
            x = Len(Item) - Len(Replace(Item, " ", ""))
            x1 = InStr(1, Item, " ") 'Modena Propuesta Ganadora
            x2 = InStr(InStr(1, Item, " ") + 1, Item, " ") 'Precio Propuesta Ganadora
            x3 = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'id Propuesta Ganadora
            If x = 4 Then
                x4 = InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Item Propuesta Ganadora
                Call extraer_propuesta_ganadora(numSheet_tblSales, k, Item, x1, x2, x3, x4)
            End If
            
            'Columna: puesto
            Sheets(numSheet_tblSales).Range("N" & k) = Hoja2.Range("G" & i).Value 'Puesto
    
            'Columna: estado
            Sheets(numSheet_tblSales).Range("O" & k) = Hoja2.Range("H" & i).Value 'Estado
            
            'Columna: miembro
            Item = Hoja2.Range("I" & i)
            x = Len(Item) - Len(Replace(Item, " ", ""))
            Call extraer_miembro(x, Item, k, numSheet_tblSales)
            
            'Columna: grupo
            Sheets(numSheet_tblSales).Range("S" & k) = Hoja2.Range("J" & i) 'Fecha
            
            'Columna: fecha_proceso
            Sheets(numSheet_tblSales).Range("T" & k) = Format(CDate(Hoja2.Range("K" & i).Text), "YYYY/MM/DD") 'Fecha
            
            'Columna: id
            Sheets(numSheet_tblSales).Range("U" & k) = Hoja2.Range("L" & i) 'ID
            
            k = k + 1
        End If
        
        Porcentaje = Round(i / lr, 4) * 100
        Application.StatusBar = "Porcentaje: " & Porcentaje & "% completado"
    Next i
    
    Application.StatusBar = False
    
    MsgBox "Limpieza termina con Exito!"
End Sub

Private Function extraer_item(numSheet_tblSales As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer, x3 As Integer)
    Sheets(numSheet_tblSales).Range("A" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos la placa
    Sheets(numSheet_tblSales).Range("B" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos la marca
    Sheets(numSheet_tblSales).Range("C" & k) = Mid(Item, (x2 + 1), ((x3 - x2) - 1)) 'Extraemos el modelo
    Sheets(numSheet_tblSales).Range("D" & k) = Mid(Item, (x3 + 1)) 'Extraemos el año
End Function

Private Function extraer_levantamiento(numSheet_tblSales As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer)
    Sheets(numSheet_tblSales).Range("F" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos el precio
    Sheets(numSheet_tblSales).Range("G" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos el porcentaje
    Sheets(numSheet_tblSales).Range("H" & k) = Mid(Item, (x2 + 1)) 'Extraemos el id
End Function

Private Function extraer_propuesta_ganadora(numSheet_tblSales As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer)
    Sheets(numSheet_tblSales).Range("I" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos la moneda
    Sheets(numSheet_tblSales).Range("J" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos el precio
    Sheets(numSheet_tblSales).Range("K" & k) = Replace(Mid(Item, (x2 + 1), ((x3 - x2) - 1)), "(", "") 'Extraemos procentaje
    Sheets(numSheet_tblSales).Range("L" & k) = Replace(Mid(Item, (x3 + 1), (x4 - ((x3 + 3) - 1))), ")", "") 'Extraemos el id
    Sheets(numSheet_tblSales).Range("M" & k) = Replace(Mid(Item, (x4 + 1)), ")", "") 'Extraemos el item
End Function

Private Function extraer_miembro(x As Integer, Item As String, k As Long, numSheet_tblSales As Integer)
    y = 0
    If x = 1 Then
            y = InStr(1, Item, " ")
        ElseIf x = 2 Then
            y = InStr(InStr(1, Item, " ") + 1, Item, " ")
        ElseIf x = 3 Then
            y = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 4 Then
            y = InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 5 Then
            y = InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 6 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 7 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 8 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 9 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 10 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 11 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 12 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 13 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 14 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 15 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, _
            Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 16 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, _
            Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
        ElseIf x = 17 Then
            y = InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(InStr(1, Item, " ") + 1, _
            Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ")
    End If
    
    codDocumento = Mid(Item, (y - 3), 3)
    Select Case codDocumento
        Case "DNI", "RUC"
            Sheets(numSheet_tblSales).Range("P" & k) = Mid(Item, 1, (y - 4)) 'Nombre y Apellidos
            Sheets(numSheet_tblSales).Range("Q" & k) = codDocumento 'Cod Documento
        Case Else:
            Sheets(numSheet_tblSales).Range("P" & k) = Mid(Item, 1, (y - 3)) 'Nombre y Apellidos
            codDocumento = Mid(Item, (y - 2), 2)
            Sheets(numSheet_tblSales).Range("Q" & k) = codDocumento 'Cod Documento
    End Select
    
    Sheets(numSheet_tblSales).Range("R" & k).NumberFormat = "@"
    Sheets(numSheet_tblSales).Range("R" & k) = Mid(Item, (y + 1)) 'Num Documento
End Function

Private Function cleanData()

    lr = Sheets(numSheet_tblSales).Range("U" & Rows.Count).End(xlUp).Row
    If lr > 2 Then
        Sheets(numSheet_tblSales).Select
            Rows("2:" & lr).Select
            Selection.Delete Shift:=xlUp
            Range("tblSales['Placa]").Select
    End If

End Function
