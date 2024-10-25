Sub Reporte_Socios_Mes()
    Dim wsFlujo As Worksheet
    Dim wsSocios As Worksheet
    Dim lastRow As Long, colMesInicio As Long, colMesFin As Long
    Dim mesSeleccionado As String
    Dim nombreHoja As String
    Dim headerRow As Long
    Dim i As Long

    ' Asignar la hoja "FLUJO"
    Set wsFlujo = ThisWorkbook.Sheets("FLUJO")

    ' Solicitar al usuario que ingrese el mes deseado
    mesSeleccionado = InputBox("Ingrese el nombre del mes (por ejemplo, 'ENERO', 'FEBRERO'):", "Seleccionar Mes")
End Sub
