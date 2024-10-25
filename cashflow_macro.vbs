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

    ' Verificar si el usuario ingresó un mes
    If mesSeleccionado = "" Then
        MsgBox "Debe ingresar un mes para continuar.", vbExclamation
        Exit Sub
    End If

    ' Crear el nombre de la nueva hoja basado en el mes
    nombreHoja = "SOCIOS_" & UCase(mesSeleccionado)

    ' Verificar si ya existe una hoja con ese nombre y eliminarla si existe
    On Error Resume Next
    Set wsSocios = ThisWorkbook.Sheets(nombreHoja)
    If Not wsSocios Is Nothing Then
        Application.DisplayAlerts = False
        wsSocios.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub
