Private Function redimArray2D(arr As Variant, nRow As Long) As Variant
    
    'Cambia la dimensión de un array, manteniendo su contenido
    '
    'Args:
    '   arr (Variant): Array a dimensionar
    '   nRow (Long): Número de filas a dimensionar
    '
    'Returns:
    '    Variant: Array con la nueva dimensión
    '
    
    arr = Application.Transpose(arr)
    ReDim Preserve arr(nCol, UBound(arr))
    arr = Application.Transpose(arr)
    redimArray2D = arr
End Function