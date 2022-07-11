Option Explicit

Private lr As Long, i As Long, Precio As Long, k As Long
Private x As Integer, x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, t As Integer
Private y As Integer, numSheet_tblDesiertos As Integer
Private Item As String, condicion As String, marca As String
Private year As Variant
Private Porcentaje As Single

Public Sub limpiar_datos()
    
    numSheet_tblDesiertos = 5
    Call cleanData

    'Determinamos el número de iteraciones
    lr = Hoja3.Range("J" & Rows.Count).End(xlUp).Row

    k = 2
    
    For i = 2 To lr
        condicion = Hoja3.Range("A" & i).Value
        If condicion = "" Then
            'Columna: (Placa / Marca / Modelo / Año)
            Item = Hoja3.Range("C" & i)
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
                
                Call extraer_item(numSheet_tblDesiertos, k, Item, x1, x2, x3)
                Else: Sheets(numSheet_tblDesiertos).Range("A" & k) = Hoja3.Range("C" & i).Value 'Item
            End If
            
            'Columna: Precio de reserva
            Sheets(numSheet_tblDesiertos).Range("E" & k) = Hoja3.Range("D" & i).Value 'Precio
            
            'Columna: Lavantamiento
            Item = Hoja3.Range("E" & i)
            x = Len(Item) - Len(Replace(Item, " ", ""))
            If x > 0 Then
                x1 = InStr(1, Item, " ") 'Precio Levantamiento
                If x = 2 Then
                    x2 = InStr(InStr(1, Item, " ") + 1, Item, " ") 'Porcentaje Levantamiento
                End If
                Call extraer_levantamiento(numSheet_tblDesiertos, k, Item, x1, x2)
                Else:
                    Sheets(numSheet_tblDesiertos).Range("F" & k) = Item 'Extraemos el precio
                    Sheets(numSheet_tblDesiertos).Range("G" & k) = Item 'Extraemos el porcentaje
                    Sheets(numSheet_tblDesiertos).Range("H" & k) = Item 'Extraemos el id
            End If
            
            'Columna: propuesta Ganadora
            Item = Hoja3.Range("F" & i)
            Item = Replace(Item, "                        ", " ")
            x = Len(Item) - Len(Replace(Item, " ", ""))
            x1 = InStr(1, Item, " ") 'Modena Propuesta Ganadora
            x2 = InStr(InStr(1, Item, " ") + 1, Item, " ") 'Precio Propuesta Ganadora
            x3 = InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'id Propuesta Ganadora
            If x = 4 Then
                x4 = InStr(InStr(InStr(InStr(1, Item, " ") + 1, Item, " ") + 1, Item, " ") + 1, Item, " ") 'Item Propuesta Ganadora
                Call extraer_propuesta_ganadora(numSheet_tblDesiertos, k, Item, x1, x2, x3, x4)
            End If
            
            'Columna: Estado
            Item = Hoja3.Range("G" & i)
            x = Len(Item) - Len(Replace(Item, "-", ""))
            
            If x > 0 Then
                x1 = InStr(1, Item, "-")
                Sheets(numSheet_tblDesiertos).Range("N" & k) = Mid(Item, 1, (x1 - 2)) 'Estado
                Sheets(numSheet_tblDesiertos).Range("O" & k) = Mid(Item, (x1 + 2)) 'Comentario
            End If
            
            'Columna: grupo
            Sheets(numSheet_tblDesiertos).Range("P" & k) = Hoja3.Range("H" & i) 'Fecha
            
            'Columna: fecha_proceso
            Sheets(numSheet_tblDesiertos).Range("Q" & k) = Format(CDate(Hoja3.Range("I" & i).Text), "YYYY/MM/DD") 'Fecha
            
            'Columna: id
            Sheets(numSheet_tblDesiertos).Range("R" & k) = Hoja3.Range("J" & i) 'ID
            
            k = k + 1
            
        End If
        
        Porcentaje = Round(i / lr, 4) * 100
        Application.StatusBar = "Porcentaje: " & Porcentaje & "% completado"
    Next i
    
    Application.StatusBar = False
    
    MsgBox "Limpieza termina con Exito!"
End Sub

Private Function extraer_item(numSheet_tblDesiertos As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer, x3 As Integer)
    Sheets(numSheet_tblDesiertos).Range("A" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos la placa
    Sheets(numSheet_tblDesiertos).Range("B" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos la marca
    Sheets(numSheet_tblDesiertos).Range("C" & k) = Mid(Item, (x2 + 1), ((x3 - x2) - 1)) 'Extraemos el modelo
    Sheets(numSheet_tblDesiertos).Range("D" & k) = Mid(Item, (x3 + 1)) 'Extraemos el año
End Function

Private Function extraer_levantamiento(numSheet_tblDesiertos As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer)
    Sheets(numSheet_tblDesiertos).Range("F" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos el precio
    Sheets(numSheet_tblDesiertos).Range("G" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos el porcentaje
    Sheets(numSheet_tblDesiertos).Range("H" & k) = Mid(Item, (x2 + 1)) 'Extraemos el id
End Function

Private Function extraer_propuesta_ganadora(numSheet_tblDesiertos As Integer, k As Long, Item As String, x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer)
    Sheets(numSheet_tblDesiertos).Range("I" & k) = Mid(Item, 1, (x1 - 1)) 'Extraemos la moneda
    Sheets(numSheet_tblDesiertos).Range("J" & k) = Mid(Item, (x1 + 1), ((x2 - x1) - 1)) 'Extraemos el precio
    Sheets(numSheet_tblDesiertos).Range("K" & k) = Replace(Mid(Item, (x2 + 1), ((x3 - x2) - 1)), "(", "") 'Extraemos procentaje
    Sheets(numSheet_tblDesiertos).Range("L" & k) = Replace(Mid(Item, (x3 + 1), (x4 - ((x3 + 3) - 1))), ")", "") 'Extraemos el id
    Sheets(numSheet_tblDesiertos).Range("M" & k) = Replace(Mid(Item, (x4 + 1)), ")", "") 'Extraemos el item
End Function

Private Function cleanData()

    lr = Sheets(numSheet_tblDesiertos).Range("R" & Rows.Count).End(xlUp).Row
    If lr > 2 Then
        Sheets(numSheet_tblDesiertos).Select
            Rows("2:" & lr).Select
            Selection.Delete Shift:=xlUp
            Range("Tabla5['Placa]").Select
    End If

End Function

