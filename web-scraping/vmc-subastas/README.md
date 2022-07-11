# VMC SUBASTAS

Instrucciones de funcionamiento de los modulos para el scrapeo.

## Modulos
- Modulo: Main

Modulo que contiene la llamada a todas las funciones implementadas.

## Hojas
- Hoja: shHistorialOfertas
Esta hoja va a contener el detalle general del proceso de subastas.

- Hoja: shOfertasVendidas
Esta hoja va a contener el detalle especifico de las ofertas vendidas.

- Hoja: shOfertasDesiertas
Esta hoja va a contener el detalle especifico de las ofertas desiertas.

- Hoja: shDetalle
Esta hoja va a contener el detalle de cada siniestro por tipo (vendidas o desiertas).

- Hoja: shUrlImg
Esta hoja va a contener los links de las imagenes a descargar.

## Variables globales
```vba
Option Explicit
Option Base 1

Private content As Integer, newOferta As Integer

'Para la descarga de urls
Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwreserved As Long, ByVal lpfnCB As Long) As Long
```
