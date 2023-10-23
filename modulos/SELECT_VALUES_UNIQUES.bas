Private Sub Workbook_Open()
    Call main.run_with_open_workbook
End Sub

Public Sub run_with_open_workbook()
    'UDF
    Dim name_function As String
    Dim description_function As String
    Dim category_function As Variant
    Dim arr_descrip_arg(1 To 3) As String

    name_function = "SELECT_VALUES_UNIQUES"
    description_function = "Devuelve valores únicos que cumplen un criterio; por ejemplo, todos los ramos de cierto cliente que aparece en una lista de ramos."
    category_function = "Funciones matriciales"
    arr_descrip_arg(1) = "Obligatorio. Rango en el que se evalúan los criterios asociados."
    arr_descrip_arg(2) = "Obligatorio. Los criterios en forma texto que determinan las celdas que se van a contar."
    arr_descrip_arg(3) = "Obligatorio. Rango de críterios a devolver tras la coincidencia."

    Application.MacroOptions macro:=name_function, _
    Description:=description_function, _
    Category:=category_function, _
    ArgumentDescriptions:=arr_descrip_arg
End Sub

Public Sub run_with_open_workbook()

    Dim name_function As String
    Dim description_function As String
    Dim category_function As Variant
    Dim arr_descrip_arg(1 To 3) As String
    
    name_function = "{Nombre de la función UDF}"
    description_function = "{Descripción de la función UDF}"
    category_function = "{Categoría de la función UDF}"
    arr_descrip_arg(1) = "{Descripción del primer parámetro la función UDF}"
    arr_descrip_arg(2) = "{Descripción del segundo parámetro la función UDF}"
    arr_descrip_arg(3) = "{Descripción del tercer parámetro la función UDF}"
    
    Application.MacroOptions macro:=name_function, _
                            Description:=description_function, _
                            Category:=category_function, _
                            ArgumentDescriptions:=arr_descrip_arg
                            
End Sub

Private Function SELECT_VALUES_UNIQUES(ByRef rango_criterio1 As Range _
    , criterio1 As String _
    , ByRef rango_valores As Range) As Variant

    Dim i As Long, j As Long
    Dim arrCriterios() As Variant, arrValores() As Variant
    Dim arrContent() As Variant, arrValoresUnicos As Variant

    arrCriterios = rango_criterio1.Value
    arrValores = rango_valores.Value
    j = 1

    For i = LBound(arrCriterios) To UBound(arrCriterios)

        If IsEmpty(arrValores(i, 1)) Then Exit For

            If arrCriterios(i, 1) = criterio1 Then
                ReDim Preserve arrContent(1 To j)
                arrContent(j) = arrValores(i, 1)
                j = j + 1
            End If
        Next i

        arrValoresUnicos = ArrayRemoveDups(arrContent)
        arrValoresUnicos = Application.Transpose(arrValoresUnicos)
        SELECT_VALUES_UNIQUES = arrValoresUnicos
End Function

Private Function ArrayRemoveDups(MyArray As Variant) As Variant

    'https://www.automateexcel.com/vba/remove-duplicates-array
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String

    Dim arrTemp() As String
    Dim Coll As New Collection

    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)

    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i

    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0

    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)

    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i

    'Output Array
    ArrayRemoveDups = arrTemp

End Function
