Attribute VB_Name = "ModuloF8126"
Option Explicit

' =====================================================================
' F.8126 - Generador de presentaciones SIN MOVIMIENTOS
'
' Como usar este modulo (una sola vez, despues de importarlo):
'   Alt+F8 > seleccionar "InstalarBotones" > Ejecutar.
'   Esto dibuja el boton sobre el recuadro azul de la hoja "Comandos"
'   y lo deja apuntando a la macro GenerarTXT_F8126_SinMovimientos.
'   No hace falta tocar la pestana "Programador" ni dibujar nada a mano.
'
' Uso mensual: ver hoja "Guia de Uso" del archivo.
' =====================================================================

Sub InstalarBotones()

    Dim wsCmd As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim b As Button

    On Error Resume Next
    Set wsCmd = ThisWorkbook.Sheets("Comandos")
    On Error GoTo 0

    If wsCmd Is Nothing Then
        MsgBox "No se encontro la hoja ""Comandos"". Verificar el nombre de la solapa.", vbCritical
        Exit Sub
    End If

    ' Evitar duplicados si se corre mas de una vez
    For Each b In wsCmd.Buttons
        If b.Name = "btnGenerarF8126" Then b.Delete
    Next b

    Set rng = wsCmd.Range("B6:C7")
    Set btn = wsCmd.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)

    With btn
        .Name = "btnGenerarF8126"
        .Caption = "GENERAR TXT DEL PERIODO ACTUAL"
        .OnAction = "GenerarTXT_F8126_SinMovimientos"
        .Font.Bold = True
        .Font.Size = 11
    End With

    wsCmd.Activate
    MsgBox "Boton instalado en la hoja ""Comandos"". Ya se puede usar la planilla normalmente " & _
           "(ver hoja ""Guia de Uso"").", vbInformation

End Sub

' =====================================================================
' Procesa UN periodo por vez, leido de la hoja "Periodo a generar":
'   B8  = Registro Cabecera (259 caracteres)
'   B9  = Longitud (solo controla el largo total, ver nota abajo)
'   B10 = Nombre de archivo sugerido
'   B12 = Estado general (combina TODAS las validaciones de campo, de
'         esta hoja y de "Datos generales"; debe decir exactamente
'         "LISTO PARA GENERAR")
'
' IMPORTANTE: B9 (Longitud) puede dar 259 aunque algun dato de origen
' este vacio o mal cargado, porque las formulas rellenan solas con
' ceros/espacios (por ejemplo un CUIT de mas de 11 digitos se trunca
' solo). Por eso el control real antes de escribir el archivo es B12,
' no B9.
'
' Escribe el .txt en la carpeta que elija el usuario, con codificacion
' ISO-8859-1 (la que exige el manual F.8126 V300, punto 2.3.2), sin
' salto de linea final. Para el mes siguiente: cambiar el periodo en
' la hoja "Periodo a generar" y volver a correr esta misma macro
' (apretando el boton de "Comandos").
' =====================================================================

Sub GenerarTXT_F8126_SinMovimientos()

    Dim wsPeriodo As Worksheet
    Dim carpeta As String
    Dim registro As String
    Dim nombreArchivo As String
    Dim longitud As Variant
    Dim estadoGeneral As String
    Dim fd As FileDialog

    On Error Resume Next
    Set wsPeriodo = ThisWorkbook.Sheets("Periodo a generar")
    On Error GoTo 0

    If wsPeriodo Is Nothing Then
        MsgBox "No se encontro la hoja ""Periodo a generar"". Verificar el nombre de la solapa.", vbCritical
        Exit Sub
    End If

    registro = wsPeriodo.Range("B8").Value
    longitud = wsPeriodo.Range("B9").Value
    estadoGeneral = Trim(wsPeriodo.Range("B12").Value)
    nombreArchivo = Trim(wsPeriodo.Range("B10").Value)

    ' Control principal: Estado general debe decir exactamente esto.
    ' Cubre CUIT, Denominacion, Hora, Numero verificador, Periodo,
    ' Secuencia y Secuencia de archivo, uno por uno (no solo el largo).
    If estadoGeneral <> "LISTO PARA GENERAR" Then
        MsgBox "No se genero el archivo. La hoja ""Periodo a generar"" dice en ""Estado general"":" & _
               vbCrLf & vbCrLf & Chr(34) & estadoGeneral & Chr(34) & vbCrLf & vbCrLf & _
               "Revisar las celdas en rojo (columna ""Validacion"") en esta hoja y en " & _
               """Datos generales"" antes de volver a intentar.", vbCritical
        Exit Sub
    End If

    ' Control secundario (defensivo): el largo total siempre debe ser 259.
    If longitud <> 259 Then
        MsgBox "La ""Longitud"" en B9 da " & longitud & " y deberia dar 259." & vbCrLf & _
               "No se genero nada. Revisar los datos cargados.", vbCritical
        Exit Sub
    End If

    If nombreArchivo = "" Then
        MsgBox "No se pudo calcular el nombre de archivo. Revisar la hoja ""Datos generales"" y " & _
               """Periodo a generar"".", vbCritical
        Exit Sub
    End If

    ' Elegir carpeta de destino
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Elegir carpeta donde guardar el archivo .txt del F.8126"
    If fd.Show <> -1 Then
        MsgBox "Operacion cancelada por el usuario.", vbExclamation
        Exit Sub
    End If
    carpeta = fd.SelectedItems(1)
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    EscribirArchivoISO88591 carpeta & nombreArchivo, registro

    MsgBox "Listo. Se genero el archivo:" & vbCrLf & nombreArchivo & vbCrLf & "en:" & vbCrLf & carpeta & _
           vbCrLf & vbCrLf & "Para el proximo mes: cambiar el periodo en ""Periodo a generar"" y volver " & _
           "a apretar este mismo boton.", vbInformation

End Sub

' ---------------------------------------------------------------------
' Escribe "contenido" en "ruta" con codificacion ISO-8859-1 exacta,
' sin BOM y sin salto de linea final (usa ADODB.Stream en modo texto).
' ---------------------------------------------------------------------
Private Sub EscribirArchivoISO88591(ByVal ruta As String, ByVal contenido As String)

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")

    st.Type = 2                 ' adTypeText
    st.Charset = "iso-8859-1"
    st.Open
    st.WriteText contenido      ' sin salto de linea agregado
    st.SaveToFile ruta, 2       ' adSaveCreateOverWrite
    st.Close

    Set st = Nothing

End Sub
