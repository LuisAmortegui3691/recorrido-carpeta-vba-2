Attribute VB_Name = "Módulo1"
Sub macroPruebas()

    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim duracion As Double

    Dim diasGuardar As Integer, mesGuardar As Integer, yearGuardar As Integer
    Dim carpetaEntrada As String, carpetaSalida As String, datosEmpleados As String
    Dim archivosDatosEmpleados As String
    ' Parte 2
    Dim ultimaFila As Integer, ultimaFilaPlantilla As Integer
    Dim plantilla As String, rutaPlantilla As String

    ' Registra el tiempo de inicio
    tiempoInicio = Timer
    
    diasGuardar = Day(Date)
    mesGuardar = Month(Date)
    yearGuardar = Year(Date)
    
    carpetaEntrada = ThisWorkbook.Sheets("Main").Range("C2").Value
    carpetaSalida = ThisWorkbook.Sheets("Main").Range("C3").Value
    
    If carpetaEntrada = "" And carpetaSalida = "" Then
        MsgBox "Las carpetas de entrada y salida deben estar especificadas", vbExclamation
        Exit Sub
    ElseIf Right(carpetaEntrada, 1) <> "\" And Right(carpetaSalida, 1) <> "\" Then
        carpetaEntrada = carpetaEntrada & "\"
        carpetaSalida = carpetaSalida & "\"
    End If
    
    datosEmpleados = carpetaEntrada & "Datos Empleados\"
    archivosDatosEmpleados = Dir(datosEmpleados & "*.*")
    
    ' Parte 2
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=carpetaEntrada & "plantilla\plantilla.xlsx"
    Application.DisplayAlerts = True
    
    Do While Len(archivosDatosEmpleados) > 0
    
        Application.DisplayAlerts = False
        Workbooks.OpenText Filename:=datosEmpleados & archivosDatosEmpleados
        Application.DisplayAlerts = True
        
        ' Parte 2
        ultimaFila = Workbooks(archivosDatosEmpleados).Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        ultimaFilaPlantilla = Workbooks("plantilla.xlsx").Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        
        Workbooks(archivosDatosEmpleados).Sheets(1).Range("A2:" & "B" & ultimaFila).Copy
        Workbooks("plantilla.xlsx").Sheets(1).Range("A" & ultimaFilaPlantilla + 1).PasteSpecial xlPasteAll
        
        
        Windows(archivosDatosEmpleados).Activate
        ActiveWorkbook.Close SaveChanges:=False
        
        archivosDatosEmpleados = Dir()
    
    Loop

    ' Registra el tiempo de finalización
    tiempoFin = Timer

    ' Calcula la duración en segundos
    duracion = tiempoFin - tiempoInicio

    ' Muestra la duración en la ventana inmediata (puedes adaptarlo según tus necesidades)
    Debug.Print "La duración de la ejecución fue de: " & duracion & " segundos"

End Sub
