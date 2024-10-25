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

    ' Crear una nueva hoja "SOCIOS" para el mes seleccionado
    Set wsSocios = ThisWorkbook.Sheets.Add(After:=wsFlujo)
    wsSocios.Name = nombreHoja

    ' Encontrar la primera columna del mes seleccionado
    headerRow = 1 ' Supone que la fila 1 tiene los nombres de los meses
    For colMesInicio = 1 To wsFlujo.Cells(headerRow, Columns.Count).End(xlToLeft).Column
        If UCase(wsFlujo.Cells(headerRow, colMesInicio).Value) = UCase(mesSeleccionado) Then
            Exit For
        End If
    Next colMesInicio

    ' Verificar si se encontró el mes en la hoja "FLUJO"
    If colMesInicio > wsFlujo.Cells(headerRow, Columns.Count).End(xlToLeft).Column Then
        MsgBox "No se encontró el mes especificado en la hoja 'FLUJO'.", vbExclamation
        Exit Sub
    End If

    ' Determinar el rango fusionado del mes seleccionado
    colMesFin = wsFlujo.Cells(headerRow, colMesInicio).MergeArea.Columns.Count + colMesInicio - 1

    ' Obtener el último rango usado en la hoja "FLUJO"
    lastRow = wsFlujo.Cells(Rows.Count, 1).End(xlUp).Row

    ' Eliminar todos los bordes en la hoja "SOCIOS"
    wsSocios.Cells.Borders.LineStyle = xlNone

    ' Copiar la columna A (descripciones) a la hoja "SOCIOS"
    For i = 1 To lastRow
        wsSocios.Cells(i, 1).Value = wsFlujo.Cells(i, 1).Value
        wsSocios.Cells(i, 1).Interior.Color = wsFlujo.Cells(i, 1).Interior.Color
    Next i

    ' Copiar los encabezados del mes y semanas específicas a "SOCIOS"
    For j = colMesInicio To colMesFin
        wsSocios.Cells(2, j - colMesInicio + 2).Value = wsFlujo.Cells(headerRow + 1, j).Value
        wsSocios.Cells(2, j - colMesInicio + 2).Font.Bold = True
        wsSocios.Cells(2, j - colMesInicio + 2).Interior.Color = wsFlujo.Cells(headerRow + 1, j).Interior.Color
    Next j

    ' Combinar celdas para el nombre del mes en B1, abarcando todas las columnas del mes
    wsSocios.Range(wsSocios.Cells(1, 2), wsSocios.Cells(1, colMesFin - colMesInicio + 2)).Merge
    wsSocios.Cells(1, 2).Value = mesSeleccionado
    wsSocios.Cells(1, 2).HorizontalAlignment = xlCenter
    wsSocios.Cells(1, 2).VerticalAlignment = xlCenter
    wsSocios.Cells(1, 2).Font.Bold = True

    ' Copiar el contenido del mes seleccionado de la hoja "FLUJO" a "SOCIOS"
    For i = 3 To lastRow
        For j = colMesInicio To colMesFin
            wsSocios.Cells(i, j - colMesInicio + 2).Value = wsFlujo.Cells(i, j).Value
            wsSocios.Cells(i, j - colMesInicio + 2).Interior.Color = wsFlujo.Cells(i, j).Interior.Color
        Next j
    Next i

    ' Aplicar negritas y subrayado a filas específicas con nombres exactos
    For i = 3 To lastRow
        If wsSocios.Cells(i, 1).Value = "Saldo inicial bancos" Or _
           wsSocios.Cells(i, 1).Value = "TOTAL INGRESOS" Or _
           wsSocios.Cells(i, 1).Value = "TOTAL EGRESOS" Or _
           wsSocios.Cells(i, 1).Value = "SALDO BANCOS" Then
            wsSocios.Rows(i).Font.Bold = True
            If wsSocios.Cells(i, 1).Value = "TOTAL INGRESOS" Or _
               wsSocios.Cells(i, 1).Value = "TOTAL EGRESOS" Or _
               wsSocios.Cells(i, 1).Value = "SALDO BANCOS" Then
                wsSocios.Rows(i).Font.Underline = xlUnderlineStyleSingle
            End If
        End If
    Next i

    ' Combinar y centrar la fila "INGRESOS", con fondo verde
    For i = 3 To lastRow
        If wsSocios.Cells(i, 1).Value = "INGRESOS" Then
            wsSocios.Range(wsSocios.Cells(i, 1), wsSocios.Cells(i, colMesFin - colMesInicio + 2)).Merge
            wsSocios.Cells(i, 1).HorizontalAlignment = xlCenter
            wsSocios.Cells(i, 1).VerticalAlignment = xlCenter
            wsSocios.Cells(i, 1).Interior.Color = RGB(198, 239, 206) ' Verde claro
        End If
    Next i

    ' Combinar y centrar la fila "EGRESOS", con fondo rojo
    For i = 3 To lastRow
        If wsSocios.Cells(i, 1).Value = "EGRESOS" Then
            wsSocios.Range(wsSocios.Cells(i, 1), wsSocios.Cells(i, colMesFin - colMesInicio + 2)).Merge
            wsSocios.Cells(i, 1).HorizontalAlignment = xlCenter
            wsSocios.Cells(i, 1).VerticalAlignment = xlCenter
            wsSocios.Cells(i, 1).Interior.Color = RGB(255, 199, 206) ' Rojo claro
        End If
    Next i

    ' Filtrar y eliminar filas sin valores numéricos
    For i = lastRow To 3 Step -1
        If Application.WorksheetFunction.Count(wsSocios.Range(wsSocios.Cells(i, 2), wsSocios.Cells(i, colMesFin - colMesInicio + 2))) = 0 _
           And wsSocios.Cells(i, 1).Value <> "INGRESOS" And wsSocios.Cells(i, 1).Value <> "EGRESOS" Then
            wsSocios.Rows(i).Delete
        End If
    Next i

    ' Agregar bordes solo a la derecha de la columna A
    For i = 1 To lastRow
        With wsSocios.Cells(i, 1).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next i

    ' Ajustar columnas y filas
    wsSocios.Columns.AutoFit
    wsSocios.Rows.AutoFit

    ' Mensaje de finalización
    MsgBox "La hoja '" & nombreHoja & "' ha sido creada.", vbInformation
End Sub