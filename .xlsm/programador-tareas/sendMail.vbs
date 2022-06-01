Dim Excel
Dim rutaLibro
Dim LibroExcel

rutaLibro = "D:\Alin-Castillo\GitHub\Excel\.xlsm\sendMail.xlsm"

Set Excel = CreateObject("Excel.Application")
Set LibroExcel = Excel.Workbooks.Open(rutaLibro)

Excel.Application.Visible = True
Excel.Application.Run "Main"

LibroExcel.Save
LibroExcel.Close
Excel.Quit