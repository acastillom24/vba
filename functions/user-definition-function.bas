
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
