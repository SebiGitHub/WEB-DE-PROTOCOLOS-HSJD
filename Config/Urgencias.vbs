Option Explicit

'Duracion del proceso
Dim startTime	
startTime = Timer


' ******** DECLARACIÓN DE VARIABLES ********
' Variables para manejo de Excel
Dim objExcel, objWorkbook, objSheet           ' Objetos de aplicación Excel, libro y hoja
Dim strFilePath                               ' Ruta del archivo Excel de entrada
Dim lastRow, i, j, stdCol                     ' Variables para manipulación de filas
Dim fechaControl                              ' Fecha de control obtenida de la celda J4
Dim controlPanel

' Variables para conexión a Access
Dim conn, connStr



' ******** CONFIGURACIÓN INICIAL ********
' Configuración de rutas de archivos
strFilePath = "C:\temp\20250519-RRHH-Empleados-Alta-Dia-GFH-Todos_PNET-Empleados-Alta-Dia.xlsx"
connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sexposito66n\Desktop\Archivos-Proyectos\Enrique\AccessUrgencias\DatosCombinadosFinal.accdb"

' Inicialización de aplicación Excel en segundo plano
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

' Abrimos primero el workbook, así ya hay un libro activo
Set objWorkbook = objExcel.Workbooks.Open(strFilePath)
Set objSheet = objWorkbook.Sheets(1)

objExcel.ScreenUpdating = False          ' Evita repintados
objExcel.DisplayAlerts  = False          ' Sin mensajes molestos
objExcel.Calculation    = -4135          ' xlCalculationManual

' Conexión a base de datos Access
Set conn = CreateObject("ADODB.Connection")
conn.Open connStr


' Diccionarios para evitar duplicados
Dim dictContratosExcel, dictEmpresas, dictLugares, dictCategorias, dictServicios, dictEmpleados, rs
Set dictContratosExcel = CreateObject("Scripting.Dictionary")
Set dictEmpresas = CreateObject("Scripting.Dictionary")
Set dictLugares = CreateObject("Scripting.Dictionary")
Set dictCategorias = CreateObject("Scripting.Dictionary")
Set dictServicios = CreateObject("Scripting.Dictionary")
Set dictEmpleados = CreateObject("Scripting.Dictionary")

' =========================
' Tabla Empresas
' =========================
Set rs = conn.Execute("SELECT IdEmpresa, [Id] FROM T_Empresa")
Do While Not rs.EOF
    dictEmpresas.Add rs.Fields(0).Value, rs.Fields(1).Value
    rs.MoveNext
Loop: rs.Close

' =========================
' Tabla Lugares de Trabajo
' =========================
Set rs = conn.Execute("SELECT LugarTrabajo, [Id] FROM T_Lugar_Trabajo")
Do While Not rs.EOF
    dictLugares.Add rs.Fields(0).Value, rs.Fields(1).Value
    rs.MoveNext
Loop: rs.Close

' =========================
' Tabla Categorías
' =========================
Set rs = conn.Execute("SELECT Categoria, [Id] FROM T_Categoria")
Do While Not rs.EOF
    dictCategorias.Add rs.Fields(0).Value, rs.Fields(1).Value
    rs.MoveNext
Loop: rs.Close

' =========================
' Tabla Empleados
' =========================
Set rs = conn.Execute("SELECT IdEmpleado, [Id] FROM T_Empleado")
Do While Not rs.EOF
    dictEmpleados.Add rs.Fields(0).Value, rs.Fields(1).Value
    rs.MoveNext
Loop: rs.Close

' =========================
' Tabla Servicios
' =========================
Set rs = conn.Execute("SELECT Servicio, [Id] FROM T_Servicio")
Do While Not rs.EOF
    dictServicios.Add rs.Fields(0).Value, rs.Fields(1).Value
    rs.MoveNext
Loop: rs.Close

'---------------------------
' Iniciar transacción global
conn.BeginTrans
'---------------------------


' Comprobar contratos activos sin duplicados
Dim dictContratosActivos, claveSQL, arrInfo
Set dictContratosActivos = CreateObject("Scripting.Dictionary")
Set rs = conn.Execute( _
  "SELECT Id, IdEmpresa, IdLugarTrabajo, IdEmpleado, IdCategoria, IdGFH, F_Inicio, F_Confirmacion " & _
  "FROM T_Contratacion WHERE Inactivo=False")
Do While Not rs.EOF
    ' → Capturamos el valor Date de F_Inicio
    Dim dtIni, fechaKey
    dtIni = rs.Fields("F_Inicio").Value

    ' → Construimos manualmente "yyyy-mm-dd"
    fechaKey = Year(dtIni) & "-" & _
               Right("00" & Month(dtIni), 2) & "-" & _
               Right("00" & Day(dtIni), 2)

    claveSQL = rs.Fields("IdEmpresa") & "|" & rs.Fields("IdLugarTrabajo") & "|" & _
               rs.Fields("IdEmpleado") & "|" & rs.Fields("IdCategoria") & "|" & _
               rs.Fields("IdGFH") & "|" & fechaKey

    dictContratosActivos.Add claveSQL, Array(rs.Fields("Id").Value, rs.Fields("F_Confirmacion").Value)
    rs.MoveNext
Loop: rs.Close



' Obtener fecha control desde celda J4
fechaControl = objSheet.Range("J4").Value

' Encontrar última fila con datos (xlUp)
lastRow = objSheet.Cells(objSheet.Rows.Count, "A").End(-4162).Row

' Crear la hoja "Datos sin duplicados" y copiar datos de Hoja1
Dim tempSheet
On Error Resume Next
Set tempSheet = objWorkbook.Sheets("Datos sin duplicados")
If Not tempSheet Is Nothing Then
	objExcel.DisplayAlerts = False
	tempSheet.Delete
	objExcel.DisplayAlerts = True
End If
On Error GoTo 0


' Crear la nueva hoja DESPUÉS de Hoja 1
objWorkbook.Sheets.Add , objWorkbook.Sheets(1)
Set tempSheet = objWorkbook.Sheets(objWorkbook.Sheets.Count)
tempSheet.Name = "Datos sin duplicados"

' Copiar datos desde Hoja 1
objWorkbook.Sheets(1).UsedRange.Copy tempSheet.Range("A1")

' Usar la nueva hoja para todas las modificaciones
Set objSheet = tempSheet

' Recalcula la última fila en la nueva hoja
lastRow = objSheet.Cells(objSheet.Rows.Count, "A").End(-4162).Row

' Descombina todas las celdas
objSheet.UsedRange.UnMerge



' ******** PREPARACIÓN DE DATOS EN EXCEL ********
Dim telefonoData, usuarioData
Dim headerRows: headerRows = 10
Dim numRows, numCols, data
numRows  = lastRow - headerRows
numCols  = objSheet.UsedRange.Columns.Count

objSheet.Range("E" & (headerRows + 1) & ":E" & lastRow).NumberFormat = "@"

data     = objSheet.Range(objSheet.Cells(headerRows+1,1), _
                          objSheet.Cells(lastRow, numCols))


' Descombinar columnas que contienen celdas combinadas y procesar cada fila de datos (desde fila 11 en adelante)
Dim r
For r = 1 To numRows
    ' Col F (6): validar DNI
    If Trim(data(r, 6) & "") <> "" Then
        data(r, 6) = ValidarDNI(data(r, 6))
    Else
        data(r, 6) = ""
    End If

    ' Col L (12): usuario portal
    data(r, 12) = GenerarUsuario(data(r, 11), data(r, 9), data(r, 6))

    ' Col W (23): teléfono
    If Trim(data(r, 23) & "") <> "" Then
        data(r, 23) = FormatearTelefono(LimpiarTelefono(data(r, 23)))
    Else
        data(r, 23) = ""
    End If
Next

' Y al final de todo el procesamiento EXCEL, vuelcas:
objSheet.Range(objSheet.Cells(headerRows + 1, 1), _
               objSheet.Cells(lastRow, numCols)).Value = data




' ******** ELIMINACIÓN DE DUPLICADOS CON MEJOR EMAIL Y TELÉFONO ********
Dim groups, key, parts
Set groups = CreateObject("Scripting.Dictionary")

For i = headerRows + 1 To lastRow       ' de 11 a lastRow
    Dim ai : ai = i - headerRows        ' ai va de 1 a numRows
    key = Trim(data(ai,2)) & "|" & _
          Trim(data(ai,3)) & "|" & _
          Trim(data(ai,5)) & "|" & _
          Trim(data(ai,17)) & "|" & _
          Trim(data(ai,14)) & "|" & _
          Trim(data(ai,21))
		  
    If Not groups.Exists(key) Then groups.Add key, Array()
    parts = groups(key)
    ReDim Preserve parts(UBound(parts)+1)
    parts(UBound(parts)) = i
    groups(key) = parts
Next

Dim totalDuplicates, duplicatesMsg
totalDuplicates = 0
duplicatesMsg = ""

Dim rowsArray, bestRow, bestEmail, bestOriginalEmail, bestPhone
Dim currentEmail, currentPhone, rowIndex


' Diccionario para marcar las filas que se van a eliminar
Dim rowsToDeleteDict
Set rowsToDeleteDict = CreateObject("Scripting.Dictionary")


' Procesar cada grupo de duplicados
For Each key In groups.Keys
    rowsArray = groups(key)
	
    ' Si hay más de un registro en el grupo, hay duplicados
    If UBound(rowsArray) >= 1 Then
	
        ' Inicialmente, tomar la primera fila del grupo
        bestRow = CInt(rowsArray(0))
        bestOriginalEmail = objSheet.Cells(bestRow, 22).Value ' Email original
        bestEmail = LimpiarEmail(bestOriginalEmail)
        bestPhone = LimpiarTelefono(objSheet.Cells(bestRow, 23).Value)
        
        ' Recorrer cada fila del grupo para determinar cuál tiene el mejor email y teléfono
        For Each r In rowsArray
            rowIndex = CInt(r)
			
            If rowIndex <> bestRow Then
                currentEmail = LimpiarEmail(objSheet.Cells(rowIndex, 22).Value)
				
                ' Usamos la función MejorEmail para comparar el email actual con el "mejor" almacenado
                If MejorEmail(objSheet.Cells(rowIndex, 22).Value, bestOriginalEmail) Then
                    bestOriginalEmail = objSheet.Cells(rowIndex, 22).Value
                    bestEmail = currentEmail
                    bestRow = rowIndex   ' Actualizamos la fila ganadora
                End If
				
                currentPhone = LimpiarTelefono(objSheet.Cells(rowIndex, 23).Value)
				
                ' Se compara el teléfono actual con el mejor almacenado
                If MejorTelefono(currentPhone, bestPhone) Then
                    bestPhone = currentPhone
                End If
            End If
        Next
        
		
        ' Actualizar el registro ganador con el mejor email y teléfono obtenidos
        objSheet.Cells(bestRow, 22).Value = bestEmail
		
        If bestPhone <> "" Then
            objSheet.Cells(bestRow, 23).Value = FormatearTelefono(bestPhone)
        Else
            objSheet.Cells(bestRow, 23).Value = ""
        End If
        
        ' Registrar todas las filas del grupo, excepto la "ganadora", para eliminarlas
        For Each r In rowsArray
            rowIndex = CInt(r)
			
            If rowIndex <> bestRow Then
			
                If Not rowsToDeleteDict.Exists(rowIndex) Then
                    rowsToDeleteDict.Add rowIndex, rowIndex
                    totalDuplicates = totalDuplicates + 1
					
                    ' Se asume que el ID de la persona se encuentra en la columna 1
                    duplicatesMsg = duplicatesMsg & "ID: " & objSheet.Cells(rowIndex, 1).Value & vbCrLf
                End If
            End If
        Next
    End If
Next


' Convertir el diccionario de filas a eliminar en un arreglo y ordenarlo en forma descendente
Dim rowIndexes(), idx, tmp
ReDim rowIndexes(rowsToDeleteDict.Count - 1)
idx = 0

For Each r In rowsToDeleteDict.Keys
    rowIndexes(idx) = r
    idx = idx + 1
Next

' Ordenar el arreglo de mayor a menor para evitar problemas al eliminar filas
For i = 0 To UBound(rowIndexes) - 1
    For j = i + 1 To UBound(rowIndexes)
        If rowIndexes(i) < rowIndexes(j) Then
            tmp = rowIndexes(i)
            rowIndexes(i) = rowIndexes(j)
            rowIndexes(j) = tmp
        End If
    Next
Next

' Eliminar las filas duplicadas
Dim delRange, hasRange
hasRange = False

For i = 0 To UBound(rowIndexes)
    If Not hasRange Then
        ' Primera fila
        Set delRange = objSheet.Rows(rowIndexes(i))
        hasRange = True
    Else
        ' Añadimos al rango existente
        Set delRange = objExcel.Application.Union(delRange, objSheet.Rows(rowIndexes(i)))
    End If
Next

' Solo borramos si llegamos a crear el rango una vez
If hasRange Then delRange.Delete



' ******** ACTUALIZAR COLUMNA STD_DT_END ********
stdCol = 20
lastRow = objSheet.Cells(objSheet.Rows.Count, "A").End(-4162).Row
objSheet.Range(objSheet.Cells(headerRows+1, stdCol), _
               objSheet.Cells(lastRow, stdCol)).Value = fechaControl
			   

' ******** INSERCIÓN EN ACCESS ********
For i = headerRows + 1 To lastRow

    ' Obtener valores del Excel
    Dim idEmpresa, lugarTrabajo, categoria, idEmpleado, dni, portal, nombre, apellidos, idLugarTrabajo, idCategoria, idGFH, servicio, rawFConfirmacion, rawFInicio
    idEmpresa = objSheet.Cells(i, 2).Value
	idLugarTrabajo = objSheet.Cells(i, 3).Value
    lugarTrabajo = objSheet.Cells(i, 4).Value
	idEmpleado = objSheet.Cells(i, 5).Value
	dni = objSheet.Cells(i, 6).Value
	apellidos = objSheet.Cells(i, 9).Value & " " & objSheet.Cells(i, 8).Value
	nombre = objSheet.Cells(i, 11).Value
    portal = objSheet.Cells(i, 12).Value
	idGFH = objSheet.Cells(i, 14).Value
	idCategoria = objSheet.Cells(i, 17).Value
    categoria = objSheet.Cells(i, 18).Value
	rawFConfirmacion = objSheet.Cells(i, 20).Value
	rawFInicio = objSheet.Cells(i, 21).Value


	' --- Obtener o crear registros y capturar los IDs autonuméricos ---
	Dim autoID_Empresa, autoID_Lugar, autoID_Categoria, autoID_Empleado, autoID_GFh


    '============================================
    ' 1. TABLA T_EMPRESA
    '============================================
    If dictEmpresas.Exists(idEmpresa) Then
        autoID_Empresa = dictEmpresas(idEmpresa)
    Else
        conn.Execute "INSERT INTO T_Empresa (IdEmpresa, F_Creacion, Inactivo) VALUES ('" & Replace(idEmpresa, "'", "''") & "', Now(), False)"
        autoID_Empresa = conn.Execute("SELECT @@IDENTITY").Fields(0).Value
        dictEmpresas.Add idEmpresa, autoID_Empresa
    End If
				
					
	'============================================
	' 2. TABLA T_LUGAR_TRABAJO
	'============================================
	If dictLugares.Exists(lugarTrabajo) Then
		autoID_Lugar = dictLugares(lugarTrabajo)
	Else
		conn.Execute "INSERT INTO T_Lugar_Trabajo " & _
					 "(F_Creacion, Inactivo, LugarTrabajo, IdEmpresa) " & _
					 "VALUES (Now(), False, '" & Replace(lugarTrabajo, "'", "''") & "', " & autoID_Empresa & ")"
		autoID_Lugar = conn.Execute("SELECT @@IDENTITY").Fields(0).Value
		dictLugares.Add lugarTrabajo, autoID_Lugar
	End If


    '============================================
    ' 3. TABLA T_CATEGORIA
    '============================================
	If dictCategorias.Exists(categoria) Then
		autoID_Categoria = dictCategorias(categoria)
	Else
		conn.Execute "INSERT INTO T_Categoria (F_Creacion, Inactivo, Categoria) " & _
					 "VALUES (Now(), False, '" & Replace(categoria, "'", "''") & "')"
		autoID_Categoria = conn.Execute("SELECT @@IDENTITY").Fields(0).Value
		dictCategorias.Add categoria, autoID_Categoria
	End If


    '============================================
    ' 4. TABLA T_EMPLEADO 
    '============================================
	If dictEmpleados.Exists(idEmpleado) Then
		autoID_Empleado = dictEmpleados(idEmpleado)
	Else
		conn.Execute "INSERT INTO T_Empleado (F_Creacion, Inactivo, IdEmpleado, DNI, Portal_Empleado, Nombre, Apellidos) " & _
					 "VALUES (Now(), False, '" & Replace(idEmpleado, "'", "''") & "', '" & Replace(dni, "'", "''") & "', " & _
					 "'" & Replace(portal, "'", "''") & "', '" & Replace(nombre, "'", "''") & "', '" & Replace(apellidos, "'", "''") & "')"
		autoID_Empleado = conn.Execute("SELECT @@IDENTITY").Fields(0).Value
		dictEmpleados.Add idEmpleado, autoID_Empleado
	End If


	'============================================
	' 5. TABLA T_Servicio
	'============================================
	If dictServicios.Exists(idGFH) Then
		autoID_GFh = dictServicios(idGFH)
	Else
		conn.Execute "INSERT INTO T_Servicio " & _
					 "(F_Creacion, Inactivo, Servicio, IdLugarTrabajo) " & _
					 "VALUES (Now(), False, '" & Replace(idGFH, "'", "''") & "', " & autoID_Lugar & ")"
		autoID_GFh = conn.Execute("SELECT @@IDENTITY").Fields(0).Value
		dictServicios.Add idGFH, autoID_GFh
	End If


	' ============================================
	' 6. TABLA T_Contratacion
	' ============================================
	Dim fechaInicioStr, fechaConfirmacionStr, fechaInicio, fechaConfirmacion
	Dim fechaInicioClave, clave
	

	' ------------------- FORMATEO DE FECHAS PARA SQL/ACCESS -------------------
	' Parsear fechas en formato personalizado desde datos crudos
    fechaInicioStr = ParseCustomDate(rawFInicio)
	fechaConfirmacionStr = ParseCustomDate(rawFConfirmacion)
	
	
	If IsNull(fechaInicioStr) Then
		fechaInicio = "Null"
		fechaInicioClave = "NULL"
	Else
		Dim dtInicio
		dtInicio = CDate(fechaInicioStr)
		fechaInicioClave = Year(dtInicio) & "-" & _
						 Right("00" & Month(dtInicio), 2) & "-" & _
						 Right("00" & Day(dtInicio), 2)
						 
		fechaInicio = "#" & fechaInicioClave & "#"
	End If
	

	' Manejar fecha de confirmación para consultas SIN AMBIGÜEDADES
	If IsNull(fechaConfirmacionStr) Then
		fechaConfirmacion = "Null"
	Else
		Dim dtConfirm
		dtConfirm = CDate(fechaConfirmacionStr)
		' Formatear fecha como #YYYY-MM-DD# (formato Access)
		fechaConfirmacion = "#" & Year(dtConfirm) & "-" & _
						   Right("00" & Month(dtConfirm), 2) & "-" & _
						   Right("00" & Day(dtConfirm),   2) & "#"

	End If


    ' Construcción final de la clave
    clave = autoID_Empresa & "|" & autoID_Lugar & "|" & _
            autoID_Empleado & "|" & autoID_Categoria & "|" & _
            autoID_GFh & "|" & fechaInicioClave

	' Registrar clave única en diccionario para evitar duplicados
	If Not dictContratosExcel.Exists(clave) Then
		dictContratosExcel.Add clave, True  ' El valor booleano es irrelevante, solo necesitamos la clave
	End If


	' ------------------- COMPROBACIONES CONSULTAS SQL -------------------
	Dim autoID_Contratacion
    If dictContratosActivos.Exists(clave) Then
        ' 1) YA EXISTE: solo actualizar f_confirmacion si es más reciente
        arrInfo = dictContratosActivos(clave)
        autoID_Contratacion = arrInfo(0)
        If fechaConfirmacion > arrInfo(1) Then
				'MsgBox "fechaConfirmacion = " & fechaConfirmacion

            conn.Execute "UPDATE T_Contratacion " & _
                         "SET F_Confirmacion=" & fechaConfirmacion & " " & _
                         "WHERE Id=" & autoID_Contratacion
        End If
    Else
        ' 2) NO EXISTE: insertar nuevo contrato
        conn.Execute "INSERT INTO T_Contratacion " & _
                     "(F_Creacion, Inactivo, IdEmpresa, IdLugarTrabajo, " & _
                     "IdEmpleado, IdCategoria, IdGFH, F_Inicio, F_Confirmacion, Motivo_Baja) " & _
                     "VALUES (Now(), False, " & _
                     autoID_Empresa & ", " & autoID_Lugar & ", " & _
                     autoID_Empleado & ", " & autoID_Categoria & ", " & autoID_GFh & ", " & _
                     fechaInicio & ", " & fechaConfirmacion & ", '')"
    End If
Next


' —————————————————————————————————————————————————————————————
' 7. MARCAR INACTIVOS Y ASIGNAR MOTIVO_BAJA (paso 3)
' —————————————————————————————————————————————————————————————
' Primero: desactivar todos los que en BD NO estén en el Excel
For Each claveSQL In dictContratosActivos.Keys
    If Not dictContratosExcel.Exists(claveSQL) Then
        autoID_Contratacion = dictContratosActivos(claveSQL)(0)
        conn.Execute "UPDATE T_Contratacion " & _
                     "SET Inactivo=True " & _
                     "WHERE Id=" & autoID_Contratacion
    End If
Next

' Después: para cada contrato inactivo sin motivo, calculo el motivo
Dim rsBaja, newCat, newSrv, idEmp, motivo, idContrato, oldCat, oldSrv
Set rsBaja = conn.Execute( _
  "SELECT Id, IdEmpleado, IdCategoria, IdGFH " & _
  "FROM T_Contratacion " & _
  "WHERE Inactivo=True AND (Motivo_Baja IS NULL OR Motivo_Baja='')")

Do While Not rsBaja.EOF
    idContrato = rsBaja("Id")
    idEmp       = rsBaja("IdEmpleado")
    oldCat      = rsBaja("IdCategoria")
    oldSrv      = rsBaja("IdGFH")

    ' Intento leer el contrato activo más reciente para ese empleado
    Dim rsNew
    Set rsNew = conn.Execute( _
      "SELECT TOP 1 IdCategoria, IdGFH " & _
      "FROM T_Contratacion " & _
      "WHERE IdEmpleado=" & idEmp & " AND Inactivo=False " & _
      "ORDER BY F_Inicio DESC")

    If Not rsNew.EOF Then
        newCat = rsNew("IdCategoria")
        newSrv = rsNew("IdGFH")

        If oldCat <> newCat And oldSrv <> newSrv Then
            motivo = "Cambio de servicio y categoria"
        ElseIf oldSrv <> newSrv Then
            motivo = "Cambio de servicio"
        ElseIf oldCat <> newCat Then
            motivo = "Cambio de categoria"
        Else
            motivo = "Fin de Contrato"
        End If
    Else
        ' No hay contrato activo: contrato finalizado
        motivo = "Contrato finalizado"
    End If
    rsNew.Close

    ' Guardo el motivo
    conn.Execute "UPDATE T_Contratacion " & _
                 "SET Motivo_Baja='" & Replace(motivo, "'", "''") & "' " & _
                 "WHERE Id=" & idContrato

    rsBaja.MoveNext
Loop
rsBaja.Close
				 

' Optimizacion del tiempo
objExcel.ScreenUpdating = True
objExcel.Calculation = -4105   ' xlCalculationAutomatic
objExcel.EnableEvents = True		


' ↓↓↓ Cierro la transacción global ↓↓↓
conn.CommitTrans		 


' ******** FINALIZACIÓN ********
' Guardar cambios y cerrar Excel
objWorkbook.Save
objWorkbook.Close
objExcel.Quit


' Liberar recursos
Set conn = Nothing
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing


' ———————— Eliminar fichero de entrada ————————
On Error Resume Next
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(strFilePath) Then
    fso.DeleteFile strFilePath, True
End If

Set fso = Nothing

On Error GoTo 0
' ————————————————————————————————————————


' Calculamos y mostramos el tiempo
Dim elapsedTime
elapsedTime = Round(Timer - startTime, 2)  ' segundos, con dos decimales


' Mensaje de exito
MsgBox "Datos insertados correctamente en Access y guardados en 'Datos sin duplicados'." & vbCrLf & _
	   "El archivo origen ha sido eliminado." & vbCrLf & _
	   "Tiempo total: " & elapsedTime & " segundos", vbInformation


' ==================================================================================
' ****************************** FUNCIONES AUXILIARES ******************************
' ==================================================================================
' =====================
' FUNCIÓN PARA contar cuantos excels hay abierto en un momento determinado
' =====================
	' Función para obtener el numero de excels abiertos en un momento determinado
Function CountOpenExcels(exApp)
    CountOpenExcels = exApp.Workbooks.Count
End Function


' =====================
' FUNCIÓN PARA ACTUALIZAR LOS INACTIVOS Y OBTENER LOS IDS AUTONUMERICOS DE CADA TABLAA
' =====================
	' Función para obtener (o crear) el ID autonumérico de un registro en una tabla determinada.
	' tableName: nombre de la tabla (por ejemplo, "T_Empresa")
	' uniqueField: nombre del campo único (por ejemplo, "IdEmpresa")
	' autoField: nombre del campo autonumérico (por ejemplo, "AutoID")
	' value: valor proveniente de Excel (que usas para buscar)
	' insertSQL: sentencia SQL para insertar el registro en caso de no existir
	' conn: conexión a la base de datos
Function GetOrCreateAutoID(tableName, uniqueField, autoField, value, insertSQL, conn)
    Dim rs, rsID, selectSQL, quotedValue, autoID
    On Error Resume Next: Err.Clear

    ' 1) Preparar quotedValue:
    '    Si estamos en T_Empleado y uniqueField="IdEmpleado", forzamos siempre texto
    If LCase(tableName) = "t_empleado" And LCase(uniqueField) = "idempleado" Then
        quotedValue = "'" & Replace(CStr(value), "'", "''") & "'"
    ElseIf IsNumeric(value) Then
        quotedValue = CStr(value)
    Else
        quotedValue = "'" & Replace(CStr(value), "'", "''") & "'"
    End If

	
    If Err.Number <> 0 Then Err.Clear: GetOrCreateAutoID = -1: Exit Function

    ' 2) Intentar SELECT
    selectSQL = "SELECT TOP 1 [" & autoField & "] FROM [" & tableName & "] " & _
                "WHERE [" & uniqueField & "] = " & quotedValue
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open selectSQL, conn, 1, 3
	
	If Err.Number <> 0 Then
		MsgBox "Error " & Err.Number & " en SELECT: " & selectSQL
		Err.Clear
		GetOrCreateAutoID = -1
		Exit Function
	End If


    ' 3) Si existe, devolvemos ese Id
    If Not rs.EOF Then
        autoID = rs.Fields(autoField).Value
        rs.Close: Err.Clear
        GetOrCreateAutoID = autoID
        Exit Function
    End If

    ' 4) Si no existe, insertamos
    conn.Execute insertSQL
    If Err.Number <> 0 Then
        If Err.Number = 3022 Then
            ' Clave duplicada: volvemos a SELECT para capturar
            Err.Clear
            rs.Open selectSQL, conn, 1, 3
            If Err.Number <> 0 Or rs.EOF Then Err.Clear: GetOrCreateAutoID = -1: Exit Function
            autoID = rs.Fields(autoField).Value
            rs.Close: Err.Clear
            GetOrCreateAutoID = autoID
            Exit Function
        Else
            ' Otro error
            Err.Clear: rs.Close: GetOrCreateAutoID = -1: Exit Function
        End If
    End If

    ' 5) Insert OK: recuperamos @@IDENTITY
    Set rsID = conn.Execute("SELECT @@IDENTITY")
    If Err.Number <> 0 Then Err.Clear: rs.Close: GetOrCreateAutoID = -1: Exit Function
    autoID = rsID.Fields(0).Value
    rsID.Close

    ' 6) Cerrar y devolver
    rs.Close: Err.Clear
    GetOrCreateAutoID = autoID
End Function


' =====================
' FUNCIÓN PARA PARSEAR FECHAS
' =====================
    ' Convierte diferentes formatos de fecha a formato estándar YYYY-MM-DD
    ' Maneja tanto fechas seriales de Excel como textos con formatos variados
Function ParseCustomDate(value)
    Dim parts, dd, mm, yyyy, strDate
    On Error Resume Next

    ' --- Manejar números (fechas seriales de Excel) ---
	If IsNumeric(value) Then
		Dim dts: dts = CDate(value)
		ParseCustomDate = Year(dts) & "-" & _
						  Right("00" & Month(dts), 2) & "-" & _
						  Right("00" & Day(dts),   2)
		Exit Function
	End If

	
	    ' Manejar fechas con formato "4000-01-16" como nulas (fechas inválidas)
    If IsDate(value) Then
        Dim dt
        dt = CDate(value)
        If Year(dt) >= 4000 Then
            ParseCustomDate = Null
            Exit Function
        End If
    End If

    ' --- Manejar fechas como texto con formato incorrecto ---
    value = Replace(Replace(Trim(value), "/", "-"), ".", "-")
    parts = Split(value, "-")

    If UBound(parts) <> 2 Then
        ParseCustomDate = Null
        Exit Function
    End If

    dd = parts(0)
    mm = parts(1)
    yyyy = parts(2)

    If Not (IsNumeric(dd) And IsNumeric(mm) And IsNumeric(yyyy)) Then
        ParseCustomDate = Null
        Exit Function
    End If

	' --- Ajustar año de 2 dígitos (1900-2100) ---
	If Len(yyyy) = 2 Then
		Dim currentYear
		currentYear = Year(Date) ' Obtiene el año actual del sistema
		
		Dim pivot
		pivot = currentYear - 2000 ' Calcula el pivote dinámico
		
		If CInt(yyyy) <= pivot Then
			yyyy = "20" & yyyy ' Años futuros o recientes
		Else
			yyyy = "19" & yyyy ' Años pasados del siglo XX
		End If
	End If

    ParseCustomDate = yyyy & "-" & Right("00" & mm, 2) & "-" & Right("00" & dd, 2)
End Function


' =====================
' FUNCIÓN PARA ENCONTRAR EL MEJOR EMAIL
' =====================
    ' Conprueba los emails y analiza cual es se adecua mas a los criterios que necesito
    ' Los criterios son, extension, nombre compuesto y caracteres eespeciales.
Function MejorEmail(newEmail, currentBest)
    If currentBest = "" Then
        MejorEmail = True
        Exit Function
    ElseIf newEmail = "" Then
        MejorEmail = False
        Exit Function
    End If

    Dim newTieneCompuesto, currentTieneCompuesto
    newTieneCompuesto = TieneNombreCompuesto(newEmail)
    currentTieneCompuesto = TieneNombreCompuesto(currentBest)

    If newTieneCompuesto And Not currentTieneCompuesto Then
        MejorEmail = True
        Exit Function
    ElseIf currentTieneCompuesto And Not newTieneCompuesto Then
        MejorEmail = False
        Exit Function
    End If

    If Right(newEmail, 7) = "@sjd.es" And Right(currentBest, 7) <> "@sjd.es" Then
        MejorEmail = True
    Else
        MejorEmail = False
    End If
End Function


' =====================
' FUNCIÓN PARA COMPROBAR SI EL EXPLEADO TIENE UN NOMBRE COMPUESTO
' =====================
    ' Conprueba la columna nombre y asigna cuantos nombres tiene si tiene mas de uno o no
Function TieneNombreCompuesto(email)
    Dim partes, localPart
    partes = Split(email, "@")
    If UBound(partes) >= 0 Then
        localPart = partes(0)
        TieneNombreCompuesto = InStr(localPart, " ") > 0
    Else
        TieneNombreCompuesto = False
    End If
End Function


' =====================
' FUNCIÓN PARA VERIFICAR EL EMAIL
' =====================
    ' Consiste en eliminar caracteres especiales o espacios, para que el email este correctamente escrito.
Function LimpiarEmail(email)
    Dim i, emailLimpio, caracterActual, emailSinEspacios
    Dim countArroba, posAt, dominio

    ' Remover espacios y construir el email sin espacios
    emailLimpio = ""
    For i = 1 To Len(email)
        caracterActual = Mid(email, i, 1)
        If caracterActual <> " " Then
            emailLimpio = emailLimpio & caracterActual
        End If
    Next
    ' Convertir a minúsculas
    emailSinEspacios = LCase(emailLimpio)
    
    ' Corregir errores comunes en el dominio
    emailSinEspacios = Replace(emailSinEspacios, "gmial", "gmail")
    emailSinEspacios = Replace(emailSinEspacios, "hotmial", "hotmail")
    
    ' Contar cuántos "@" hay en el email
    countArroba = Len(emailSinEspacios) - Len(Replace(emailSinEspacios, "@", ""))
    If countArroba <> 1 Then
        ' Si hay cero o más de uno, retornamos cadena vacía (o podrías manejarlo como error)
        LimpiarEmail = ""
        Exit Function
    End If

    ' Comprobar que después del "@" exista algún punto (para identificar una extensión)
    posAt = InStr(emailSinEspacios, "@")
    If posAt > 0 Then
        dominio = Mid(emailSinEspacios, posAt + 1)
        If InStr(dominio, ".") = 0 Then
            ' Si no hay punto en el dominio, se asume que falta la extensión, por defecto se agrega ".com"
            emailSinEspacios = emailSinEspacios & ".com"
        End If
    End If
    
    LimpiarEmail = emailSinEspacios
End Function


' =====================
' FUNCIÓN PARA ENCONTRAR EL MEJOR TELEFONO
' =====================
    ' Conprueba los telefonos y analiza cual es se adecua mas a los criterios que necesito.
    ' Los criterios son, nºde dígitos, por que número empieza y si tienen caracteres especiales.
Function MejorTelefono(newPhone, currentBest)
    Dim newDigit, currentDigit
    If currentBest = "" Then
        MejorTelefono = True
        Exit Function
    End If

    newDigit = Left(newPhone, 1)
    currentDigit = Left(currentBest, 1)

    If (newDigit = "6" Or newDigit = "7") And (currentDigit <> "6" And currentDigit <> "7") Then
        MejorTelefono = True
    Else
        MejorTelefono = False
    End If
End Function


' =====================
' FUNCIÓN PARA VERIFICAR EL TELEFONO
' =====================
    ' Consiste en eliminar caracteres especiales o espacios, para que el telefono este correctamente escrito.
Function LimpiarTelefono(numero)
    Dim i, ch, resultado
    resultado = ""
    For i = 1 To Len(numero)
        ch = Mid(numero, i, 1)
        If ch >= "0" And ch <= "9" Then
            resultado = resultado & ch
        End If
    Next
    If Len(resultado) > 9 Then
        resultado = Left(resultado, 9)
    End If
    LimpiarTelefono = resultado
End Function


' =====================
' FUNCIÓN PARA FORMATEAR EL TELEFONO
' =====================
    ' Consiste en darle un patron especifico al telefono.
Function FormatearTelefono(numero)
    If Len(numero) = 9 Then
        FormatearTelefono = Left(numero, 3) & " " & Mid(numero, 4, 2) & " " & Mid(numero, 6, 2) & " " & Right(numero, 2)
    Else
        FormatearTelefono = numero
    End If
End Function


' =====================
' FUNCIÓN PARA VERIFICAR EL DNI
' =====================
    ' Valida y formatea un DNI español:
    ' - Elimina caracteres no numéricos
    ' - Completa con ceros si es necesario
    ' - Calcula letra de verificación
Function ValidarDNI(dni)
    Dim i, ch, limpio, numPart, letterPart
    limpio = ""
    letterPart = ""

    For i = 1 To Len(dni)
        ch = Mid(dni, i, 1)
        If ch >= "0" And ch <= "9" Then
            limpio = limpio & ch
        ElseIf i = Len(dni) And ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then
            letterPart = UCase(ch)
        End If
    Next

    If Len(limpio) = 0 Then
        ValidarDNI = dni
        Exit Function
    End If

    If Len(limpio) > 8 Then
        numPart = Left(limpio, 8)
    ElseIf Len(limpio) < 8 Then
        numPart = Right("00000000" & limpio, 8)
    Else
        numPart = limpio
    End If

    On Error Resume Next
    Dim dniNumber
    dniNumber = CLng(numPart)
    If Err.Number <> 0 Then
        ValidarDNI = dni
        Exit Function
    End If
    On Error GoTo 0

    Dim resto, letras, expectedLetter
    resto = dniNumber Mod 23
    letras = Array("T", "R", "W", "A", "G", "M", "Y", "F", "P", "D", "X", "B", "N", "J", "Z", "S", "Q", "V", "H", "L", "C", "K", "E")
    expectedLetter = letras(resto)

    If letterPart = "" Then
        letterPart = expectedLetter
    ElseIf letterPart <> expectedLetter Then
        letterPart = expectedLetter
    End If

    ValidarDNI = numPart & letterPart
End Function


' =====================
' FUNCIÓN PARA GENERAR EL USUARIO DEL PORTAL DE EMPLEADOS
' =====================
    ' Genera nombre de usuario combinando:
    ' - Iniciales del nombre (excluyendo preposiciones)
    ' - Parte del apellido
    ' - Dígitos del DNI
Function GenerarUsuario(nombreCompleto, primerApellido, dni)
    Dim palabrasExcluidas, partesNombre, iniciales, i, dniDigitos, letraDNI
	
	    ' 1) Normalizar la ñ en ambos inputs
    nombreCompleto = Replace(nombreCompleto, "Ñ", "N")
    nombreCompleto = Replace(nombreCompleto, "ñ", "n")
    primerApellido  = Replace(primerApellido,  "Ñ", "N")
    primerApellido  = Replace(primerApellido,  "ñ", "n")
	
    palabrasExcluidas = Array("De", "La", "Del", "Los", "Las", "El", "Y", "A", "En", "Con")
    partesNombre = Split(nombreCompleto, " ")
    iniciales = ""

    For i = 0 To UBound(partesNombre)
        If Not EstaEnLista(partesNombre(i), palabrasExcluidas) Then
            iniciales = iniciales & LCase(Left(partesNombre(i), 1))
        End If
    Next

    ' Procesar el apellido para excluir palabras excluidas
    Dim apellidoPartes, parte, apellidoFiltrado
    apellidoFiltrado = ""
    apellidoPartes = Split(Trim(primerApellido), " ")
    For Each parte In apellidoPartes
        If Not EstaEnLista(parte, palabrasExcluidas) Then
            apellidoFiltrado = apellidoFiltrado & parte & " "
        End If
    Next
    apellidoFiltrado = Trim(apellidoFiltrado)
	
    ' Si después de filtrar está vacío, usar el apellido original
    If apellidoFiltrado = "" Then
        apellidoFiltrado = Trim(primerApellido)
    End If
    ' Truncar a 8 caracteres
    Dim apellidoTruncado
    apellidoTruncado = LCase(Left(apellidoFiltrado, 8))

    letraDNI = Right(dni, 3)
    GenerarUsuario = iniciales & apellidoTruncado & letraDNI
End Function


' =====================
' FUNCIÓN PARA COMPROBAR QUE ELEMENTOS ESTAN EN LA LISTA PARA POSTERIORMENETE ELIMINARLOS O QUEDARNOSLOS
' =====================
    ' Genera nombre de usuario combinando:
Function EstaEnLista(palabra, lista)
    Dim item
    For Each item In lista
        If LCase(palabra) = LCase(item) Then
            EstaEnLista = True
            Exit Function
        End If
    Next
    EstaEnLista = False
End Function