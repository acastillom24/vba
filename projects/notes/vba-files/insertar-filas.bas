Attribute VB_Name = "alincastillo"
Public Sub insertarFilas()
    
    'Info from: https://www.automateexcel.com/vba/insert-row-column/
    
    'Insertar una fila
    Rows(1).Insert
    
    'Insertar varias filas
    Rows("1:4").Insert
    
    'Insertar una fila, con el formato de la fila anterior
    Rows(2).Insert , xlFormatFromLeftOrAbove

    'Insertar una fila, con el formato de la fila de abajo
    Rows(5).Insert , xlFormatFromRightOrBelow

End Sub
