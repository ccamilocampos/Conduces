Attribute VB_Name = "Module1"
Global Dbs As ADODB.Connection
Global Tbl1 As ADODB.Recordset
Global Tbl2 As ADODB.Recordset
Global Tbl10 As ADODB.Recordset
Global proveedor As String
Global nombrebasedatos  As String
Global servidor As String
Global transacte As Workspace
Global transacco As Workspace
Global respuesta
'Global dbs As Database     'Z:TERMINAL
'Global dbs1 As Database    'Z:SEGURIDAD
'Global dbsX As Database     'X:TERMINAL SERVIDOR
'Global dbs1X As Database    'X:SEGURIDAD SERVIDOR
Global Ftb01  As ADODB.Recordset
Global Ftb02  As ADODB.Recordset
Global Ftb03  As ADODB.Recordset
Global Ftb04  As ADODB.Recordset
Global Ftb07  As ADODB.Recordset
Dim fqry01 As String
Dim fqry02 As String
Dim fqry03 As String
Dim fqry04 As String
Global DbPath               ' Ruta de la base de datos
Global DbSec                ' Ruta de la base de seguridades
Global Dbpath1               ' Ruta de la base de datos EN X
Global Dbsec1                ' Ruta de la base de seguridades EN X
Global G_usu_con_roda As Boolean 'Esta variable es para determinar si el usuario desea continuar sin rodamiento (el mismo decide el aviRoda.frm)
Global G_valida_credito
Global RpPath               ' Ruta de los reportes
Global forma        ' Variables de busqueda
Global forma1       ' Variables de busqueda
Global ruvalpas     'Valor del Pasaje Corriente
Global porcli       'Porcentaje Climatizado
Global poraer       'Porcentaje Aeroban
Global pordir       'Porcentaje Directo
Global porluj       'porcentaje lujo
Global poroid       'porcentaje oid
Global porond       'porcentaje ond
Global poronn       'porcentaje onn
Global valtax       'valor conduce taxi
Global valtra       'valor conduce transito
Global oritra       'origen transito
Global pasoco       'paso
Global Login        'login
Global Pc           'pc
Global G_Turno      'turno operador cajas
Global G_poresp     'porcentaje esp
Global G_poresc     'porcentaje esc
Global G_portax     'porcentaje tax
Global G_maxtur     'Maximo turno
Global G_cajacerrada 'Caja cerrada Conduces
Global clave_usu
Global n_clave
Global c_clave
Global G_rolusu
Global G_modulo
Global hoy
Global G_NomEmpresa  'Nombre de la empresa que se encuentra en la BD
Global CIUDAD        'Nombre de la ciudad donde se encuentra el TT
Global NUMAVISOS
Global cierreturno 'controla quien debe de cerrar turno
Global G_ValalcOri
Global G_ValalcTra
Global G_ValalcInf
Global G_NumSal    'NUmero de salidas de vehiculos
Global DIRECCION
Global G_valpendiente    'wecc 05/29/2002 valor pendiente
Global G_valotros        'wecc 05/29/2002 valor otros
Global G_Manual
Public Passwd As String   'password de la base de datos
Global G_cierredia
Global ContAlco
Global G_DirImpBol
Global Tabla_Catalogo
Global G_valida_rodamiento As String 'valor para validar si se permite la venta de TU a vehiculos SIN rodamiento.
Global G_BaseServidor As Boolean
Global G_PrinterName
Global G_PrinterDriver
Global G_PrinterPort
Global G_Imprimir
Global G_ImpCon 'impresion por una epson "E" o la laser "Z"
Global G_ImpEpsonZebra 'IMPRESION RECIBO PERMANENCIA O CANCELACION MULTA P o Z
Global G_NumImpTasa 'Numero de Impresiones Tasa de Uso
Global G_IpUnix
Global G_Conexion_reportes
Global G_Bloqueo_Tasa As String
Global G_Validar_Ser As String
Global G_tipocon As String
Global G_Tipo_Terminal As String
Global G_CantRegist
Global G_BloqTasaUso As Boolean
Global G_BloqSin_Aviso As Boolean
Global G_bloqTUxMulta As String
Global G_impresion As String
Global G_Valida_Horario As String
Global G_MinHorTU 'Minutos por encima o debaho para el control de Horario en venta de TU
Global L_HoraInicial
Global L_HoraFinal
Global G_multas_vacias As Boolean
Global G_parqueo_vacio As Boolean
Global G_Cantidad_Parqueaderos As Double
Global G_Avisos_Parqueadero As Double
Global G_Numero_Parqueadero As Double
Global G_Tipo_Parqueadero As String
Global G_Minutos_Minimos_Parqueadero As Double
Global G_Carga As String
Global filasSeleccionadasMultas As Collection
Global filasSeleccionadasParqueadero As Collection
Global G_ValidaServicio As String
Global G_Imprime_Tasa_Uso As String
Global G_Codigo_Empresa As String

Global G_Codigo_Terminal As String
Global G_Url_Web_Services As String
Global G_habLecCedula As String
Global G_Clase_Vehiculo As String
Global G_Cantidad_Puestos_Vehiculo As Double

Global G_Numero_Comparendo_Antiguo As Double

Global G_PrinterName1
Global G_PrinterDriver1
Global G_PrinterPort1

Global G_Exe_Net_Impresion As String    'ruta y nombre de exe para la impresión código de barras pdf417


'Variables pantalla visualizacion
Global ColReloj As Long
Global ColMarquesina As Long

Global TamTitulo As String
Global ColTitulo As Long
Global ColTextoGrd As String
Global ColFondo1 As Long
Global ColFondo2 As Long
Global ColTextoTitulo As Long
Global TamLetraGrilla As String

'Variables Huellero
Global G_ManLector As String
Global DbPathHuella As String
Global dbs4 As Database
Global Tbl20 As Recordset
Global Tbl21 As Recordset
Public HuellaAlc() As Byte
Public HuellaAlc1() As Byte

Global sCedulaBorrarHuella As String

Global G_Aut_Vender As String    'autorizacion para verder la tasa
Global G_Clase_Autorizacion

Global G_Ruta_Archivos_Temporales As String  'VARIABLE RUTA DE ARCHIVOS TEMPORALES RPT A ELIMINAR
Global G_PermiteInfluencia As String  'variable permite venta de tasa de uso en influencia
Global G_PermiteLeerPeaje As String    'Variable para leer peajes

Global G_Saldo_Cupo As Double
Global G_Tipo_Novedad As String

'Sebastián Rondón / Abril 12 - 2023
Global G_CierreAutomatico

'Sebastián Rondón / Agosto 09 - 2023
Global G_AvisoProtech As String

'Sebastián Rondón - Marzo 20 de 2024
Global L_Bandera_Rodamiento As Boolean

'Sebastián Rondon - Mayo 30 de 2024
Global G_AvisoEmgaf As String

'Sebastián Rondón - Febrero 27 de 2025
'Variables para mensaje de multas
Global G_Mensaje_Pago As String
Global G_Mensaje_Tipo As String
Global G_Mensaje_Opc As Boolean


'Se hace de manera global esta definición puesto a que necesitamos pasar la información de un formulario a otro
Public Type MultaSeleccionada   'esta es la estructura en la cual se guardara las multas seleccionadas
    id As String
    fechaMulta As String
    valor As String
    codinf As String
End Type

Public multasGuardadas() As MultaSeleccionada



'Se hace de manera global esta definición puesto a que necesitamos pasar la información de un formulario a otro
Public Type ParqueSeleccionado   'esta es la estructura en la cual se guardara los parqueaderos seleccionados
    id As String
    fecing As String
    fecsal As String
    valor As String
    tipo As String 'tipo de multa vigencia/permanencia
End Type

Public parqueGuardado() As ParqueSeleccionado

Public Function TieneMultasSeleccionadas() As Boolean
    On Error GoTo ArregloNoInicializado

    ' Intentar acceder al primer elemento
    Dim dummy As String
    dummy = multasGuardadas(LBound(multasGuardadas)).id

    TieneMultasSeleccionadas = True
    Exit Function

ArregloNoInicializado:
    TieneMultasSeleccionadas = False
End Function


Public Function TieneParqueaderosSeleccionados() As Boolean
    On Error GoTo ArregloNoInicializado

    ' Intentar acceder al primer elemento
    Dim dummy As String
    dummy = parqueGuardado(LBound(parqueGuardado)).id

    TieneParqueaderosSeleccionados = True
    Exit Function

ArregloNoInicializado:
    TieneParqueaderosSeleccionados = False
End Function

Public Sub PintarFilaVerde(ByVal fila As Integer, ByVal GridFlex As MSFlexGrid)
    Dim col As Integer
    For col = 0 To GridFlex.Cols - 1
        GridFlex.Row = fila
        GridFlex.col = col
        GridFlex.CellBackColor = vbGreen
        GridFlex.CellForeColor = vbWhite
        DoEvents
    Next col
End Sub

Public Sub RestaurarColorFila(ByVal fila As Integer, ByVal GridFlex As MSFlexGrid)
    Dim col As Integer
    For col = 0 To GridFlex.Cols - 1
        GridFlex.Row = fila
        GridFlex.col = col
        GridFlex.CellBackColor = vbWhite
        GridFlex.CellForeColor = vbBlack
        DoEvents
    Next col
End Sub
