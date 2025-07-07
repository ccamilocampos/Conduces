VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form main 
   Caption         =   "MODULO TASAS DE USO"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   1050
   ClientWidth     =   11805
   ClipControls    =   0   'False
   DrawMode        =   1  'Blackness
   FillColor       =   &H00400000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Login"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Vehiculos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Empresas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Cierre de Turno"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Cierre de Día"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Venta de Tasas de Uso"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Anulación de Tasas de Uso"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Registro y Negociación de Multas"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Cancelación de Multas"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Anulación de Multas"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Proceso de Consolidación"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Logoff"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2880
      Top             =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   3840
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Objetos"
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8460
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "Usuario :"
            TextSave        =   "Usuario :"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "PC : "
            TextSave        =   "PC : "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Fecha Actual :"
            TextSave        =   "Fecha Actual :"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Turno :"
            TextSave        =   "Turno :"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3598
            MinWidth        =   3598
            Text            =   "PRN :"
            TextSave        =   "PRN :"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DE9964&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   15015
   End
   Begin VB.Image ImgFondo 
      Height          =   1695
      Left            =   0
      Picture         =   "main.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2CDF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2D10C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2D426
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2D740
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2D91A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2DAF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2DE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2E3A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2E6BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2E9D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2ECEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":2F008
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnuconexion 
      Caption         =   "&Conexion"
      Begin VB.Menu Mnulogin 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu Mnucambiopassword 
         Caption         =   "Cambio de Password"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnulogoff 
         Caption         =   "Logoff"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Mnuadmiopera 
      Caption         =   "&Administracion Operativa"
      Enabled         =   0   'False
      Begin VB.Menu Mnuhojvidveh 
         Caption         =   "Hoja de Vida Vehiculos"
      End
      Begin VB.Menu Mnuhojvidemp 
         Caption         =   "Hoja de Vida Empresas"
      End
      Begin VB.Menu Mnuajuste 
         Caption         =   "Ajustar Tarifas de Rutas"
      End
      Begin VB.Menu MnuConvenios 
         Caption         =   "Valor Tasa de Uso Central"
      End
      Begin VB.Menu MnuTasSur 
         Caption         =   "Valor Tasas de Uso Sur"
      End
      Begin VB.Menu mnuconfpant 
         Caption         =   "Configuración Pantallas"
      End
      Begin VB.Menu MnuConductor 
         Caption         =   "Conductor"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuline11 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConveniosEmpresariales 
         Caption         =   "Convenios"
         Begin VB.Menu MnuEmpConvenios 
            Caption         =   "Empresas Convenios"
         End
         Begin VB.Menu mnuconvemp 
            Caption         =   "Convenios de Empresariales"
         End
      End
      Begin VB.Menu Mnuentfic 
         Caption         =   "Entrega de Fichas"
      End
      Begin VB.Menu MnuDesVenTut 
         Caption         =   "Desbloqueo Venta TUT"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuVehRes 
         Caption         =   "Vehículos en Reserva"
      End
      Begin VB.Menu MnuParqueadero 
         Caption         =   "Parqueadero"
         Begin VB.Menu MnuConDeuParqueadero 
            Caption         =   "Condonar Deuda Parqueadero"
         End
         Begin VB.Menu MnuDesAviParq 
            Caption         =   "Descontar Avisos Parqueadero"
         End
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuactrut 
         Caption         =   "Actualizacion de Rutas"
         Begin VB.Menu Mnucrerut 
            Caption         =   "Creacion de Rutas"
         End
         Begin VB.Menu MnuOperadorEmpresa 
            Caption         =   "Operadores por Empresa"
         End
         Begin VB.Menu mnurutemp 
            Caption         =   "Creación de Horarios Por Rutas"
         End
         Begin VB.Menu MnuRutAdicionales 
            Caption         =   "Rutas Adicionales"
         End
         Begin VB.Menu Mnuasihorrut 
            Caption         =   "Asignar Horarios a Rutas"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu MnuCambioServicio 
            Caption         =   "Cambio Servicio Venta Tasa Uso"
         End
      End
      Begin VB.Menu MnuIntegra 
         Caption         =   "Integra"
         Begin VB.Menu MnuEquivalencias 
            Caption         =   "Equivalencias"
         End
      End
      Begin VB.Menu mnuline3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConfVisualizacion 
         Caption         =   "Configuración Visualización"
         Begin VB.Menu MnuConfiVideos 
            Caption         =   "Configuración Videos"
         End
         Begin VB.Menu MnuConfiPresentacion 
            Caption         =   "Configuración Presentación"
         End
         Begin VB.Menu MnuConfiFormatos 
            Caption         =   "Configuración Formatos"
         End
      End
      Begin VB.Menu MnuRodamiento 
         Caption         =   "Rodamiento"
         Begin VB.Menu MnuUsuarioEmpresa 
            Caption         =   "Usuarios por Empresa"
         End
      End
      Begin VB.Menu MnuPLataformas 
         Caption         =   "Plataformas"
      End
      Begin VB.Menu MnuPicoyPlaca 
         Caption         =   "Pico y Placa"
      End
      Begin VB.Menu mnuline13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconsultas 
         Caption         =   "Consultas Generales"
         Begin VB.Menu Mnuempresas 
            Caption         =   "Empresas"
         End
         Begin VB.Menu MnuReportesRut 
            Caption         =   "Tarifas por Rutas"
         End
         Begin VB.Menu mnutagcriterio 
            Caption         =   "Tags por criterio"
         End
         Begin VB.Menu Mnuruttipveh 
            Caption         =   "Rutas por tipo de Vehiculo"
         End
         Begin VB.Menu Mnuingvehiculos 
            Caption         =   "Ingreso de Vehiculos"
         End
         Begin VB.Menu Mnuprecon 
            Caption         =   "Precio Tasas de Uso"
         End
         Begin VB.Menu Mnuconaumtar 
            Caption         =   "Aumento de Tarifas"
         End
         Begin VB.Menu mnuentsinfic 
            Caption         =   "Entrada de Vehículos sin ficha"
         End
         Begin VB.Menu mnuconempcol 
            Caption         =   "Convenios Empresariales y Colaboración"
         End
         Begin VB.Menu mnucontraemp 
            Caption         =   "Contrato de empresas"
         End
         Begin VB.Menu MnuConReserva 
            Caption         =   "Consulta de reservas"
         End
         Begin VB.Menu MnuConAudAviPar 
            Caption         =   "Auditoria Descuento de Avisos Parqueadero"
         End
         Begin VB.Menu MnuRptRodamiento 
            Caption         =   "Consulta de Rodamiento"
         End
         Begin VB.Menu mnuEvaTransito 
            Caption         =   "Evasiones Transito"
         End
         Begin VB.Menu MnuRepNovedades 
            Caption         =   "Novedades"
         End
         Begin VB.Menu MnuConVehiculos 
            Caption         =   "Consulta de Vehículos"
         End
         Begin VB.Menu MnuConConductores 
            Caption         =   "Consulta de Conductores"
         End
      End
   End
   Begin VB.Menu Mnuconduces 
      Caption         =   "Tasa de Us&o"
      Enabled         =   0   'False
      Begin VB.Menu Mnumantur 
         Caption         =   "Manejo de Turnos"
         Begin VB.Menu Mnuapetur 
            Caption         =   "Cuadre de Caja por Turno"
         End
         Begin VB.Menu Mnucietur 
            Caption         =   "Cuadre de Caja del Dia"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuline9 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuvencon 
         Caption         =   "Venta de Tasas de Uso"
      End
      Begin VB.Menu Mnucancon 
         Caption         =   "Anulacion de Tasas de Uso"
      End
      Begin VB.Menu MnuReImprimirTU 
         Caption         =   "ReImprimir Tasa de Uso"
      End
      Begin VB.Menu Mnuregpas 
         Caption         =   "Registro de Pasajeros"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuline10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuconpen 
         Caption         =   "Rangos Pendientes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuingconpen 
         Caption         =   "Conduces Pendientes "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuline6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuValTasUso 
         Caption         =   "Validación Tasa de Uso"
      End
      Begin VB.Menu mnuline12 
         Caption         =   "-"
      End
      Begin VB.Menu mnucongen 
         Caption         =   "Consultas Generales"
         Begin VB.Menu Mnuconconduces 
            Caption         =   "Tasas de Uso"
         End
         Begin VB.Menu Mnuconconcri 
            Caption         =   "Tasas de Uso por Criterio"
         End
         Begin VB.Menu Mnurecemp 
            Caption         =   "Recaudo por Empresa"
         End
         Begin VB.Menu Mnuturnos 
            Caption         =   "Turnos"
         End
         Begin VB.Menu mnurutcri 
            Caption         =   "Rutas por Criterio"
         End
         Begin VB.Menu mnuconemprut 
            Caption         =   "Tasas de Uso por Criterio Totalizada"
         End
         Begin VB.Menu mnuconvehdia 
            Caption         =   "Control Vehicular"
         End
         Begin VB.Menu mnuconpasc 
            Caption         =   "Pasajeros"
         End
         Begin VB.Menu mnuconpassal 
            Caption         =   "Pasajeros y Salidas"
         End
         Begin VB.Menu mnuconsignaciones 
            Caption         =   "Consignaciones"
         End
         Begin VB.Menu mnutasdirecc 
            Caption         =   "Tasas Por Corredor"
         End
         Begin VB.Menu mnuconvehto 
            Caption         =   "Vehículo - Totales"
         End
         Begin VB.Menu mnuinfoemp 
            Caption         =   "Información de Empresas"
         End
         Begin VB.Menu mnuvensal 
            Caption         =   "Venta y Salida de Tasas de Uso"
         End
         Begin VB.Menu mnusalvehpas 
            Caption         =   "Salida de Vehículos, Pasajeros y Totales"
         End
         Begin VB.Menu MnuGenPla 
            Caption         =   "Generar Plano"
         End
         Begin VB.Menu MnuConTasUso 
            Caption         =   "Validación Tasa de Uso"
         End
         Begin VB.Menu MnuConParqueadero 
            Caption         =   "Permanencias y Vigencias"
         End
         Begin VB.Menu MnuConDesPar 
            Caption         =   "Descuento Parqueaderos"
         End
         Begin VB.Menu MnuConEmpConvenios 
            Caption         =   "Consulta Empresas con Convenio"
         End
         Begin VB.Menu MnuConLisVehCenSur 
            Caption         =   "Consulta Listado Vehiculos Terminal Central - Terminal Sur"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu MnuRptTasAutNorte 
            Caption         =   "Tasas Automaticas Terminal Norte"
         End
         Begin VB.Menu MnuConDesOriNorte 
            Caption         =   "Despachos del Oriente por el Norte"
         End
         Begin VB.Menu Mnutucondicional 
            Caption         =   "Tasas de Uso Condicional"
         End
         Begin VB.Menu MnuCoconconcep 
            Caption         =   "Tasas de Uso y sus Conceptos"
         End
      End
   End
   Begin VB.Menu Mnumultas 
      Caption         =   "&Registro de Evasiones y Elusiones"
      Enabled         =   0   'False
      Begin VB.Menu Mnuregmul 
         Caption         =   "Registro y Negociacion de Evasiones y Elusiones"
      End
      Begin VB.Menu Mnucanmul 
         Caption         =   "Cancelacion de Evasiones y Elusiones"
      End
      Begin VB.Menu Mnuregtag 
         Caption         =   "Registro de TAGS"
      End
      Begin VB.Menu Mnuanumultas 
         Caption         =   "Anulacion  / Educativo"
      End
      Begin VB.Menu MnuDesEdu 
         Caption         =   "Descarga Educativos"
      End
      Begin VB.Menu mnuline7 
         Caption         =   "-"
      End
      Begin VB.Menu Mnumulcongen 
         Caption         =   "Consultas Generales"
         Begin VB.Menu Mnuconmul 
            Caption         =   "Evasiones y Elusiones Generales"
         End
         Begin VB.Menu Mnumulcri 
            Caption         =   "Evasiones y Elusiones por Criterio"
         End
         Begin VB.Menu mnumulcan 
            Caption         =   "Evasiones y Elusiones Canceladas"
         End
         Begin VB.Menu MnuRepSan 
            Caption         =   "Reporte de Sanciones"
         End
         Begin VB.Menu MnuRptCompElectronico 
            Caption         =   "Comparendos Electrónicos"
         End
      End
   End
   Begin VB.Menu Mnualcolemia 
      Caption         =   "A&lcoholimetria"
      Enabled         =   0   'False
      Begin VB.Menu mnuconsolidado 
         Caption         =   "Proceso de Consolidado"
      End
      Begin VB.Menu mnuline8 
         Caption         =   "-"
      End
      Begin VB.Menu mnucongenal 
         Caption         =   "Consultas Generales"
         Begin VB.Menu Mnuresultados 
            Caption         =   "Resultados de Alcoholemia"
         End
         Begin VB.Menu mnuvalalc 
            Caption         =   "Valor Alcoholemia"
         End
         Begin VB.Menu mnuconaditt 
            Caption         =   "Consorcio ADITT"
         End
      End
   End
   Begin VB.Menu MnuCupos 
      Caption         =   "C&upos"
      Enabled         =   0   'False
      Begin VB.Menu MnuIngresoCupos 
         Caption         =   "Ingreso Cupos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuIngresoFactura 
         Caption         =   "Consignación Cupos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuAnuCupos 
         Caption         =   "Anulación Cupos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuAnuFactura 
         Caption         =   "Anulación Consignación"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnurecalsaldos 
         Caption         =   "Recalculo de Saldos por Empresa"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuConGralCupos 
         Caption         =   "Consultas Generales"
         Enabled         =   0   'False
         Begin VB.Menu MnuConsultaCupos 
            Caption         =   "Consignación y Cupos"
            Enabled         =   0   'False
         End
         Begin VB.Menu MnuConMovCupos 
            Caption         =   "Movimiento de Cupos"
         End
      End
   End
   Begin VB.Menu Mnurepemgaf 
      Caption         =   "&Reportes EMGAF y Otros"
      Enabled         =   0   'False
      Begin VB.Menu Mnuemgaf 
         Caption         =   "Reporte &Emgaf"
      End
      Begin VB.Menu MnuRepAlcoPruebas 
         Caption         =   "Reporte AlcoPruebas"
      End
      Begin VB.Menu mnufectu 
         Caption         =   "Reporte Facturación Tasas de Uso"
      End
      Begin VB.Menu mnuentsalpro 
         Caption         =   "Reporte Entradas y Salidas - Protech"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Mnuayudas 
      Caption         =   "Ay&udas"
      Enabled         =   0   'False
      Begin VB.Menu mnuacercade 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Bandera_CierreAutomatico As Integer

'**************************************************************************
'**************************************************************************

Private Enum tipoConversion
    deCaF
    deFaC
End Enum
Private tipo As tipoConversion

' Estas definiciones están tomadas de lo mostrado en el explorador
' al seleccionar cada una de las funciones del servicio Web
Private Const cSOAPCaF = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<ValidaConductor xmlns=""http://localhost/"">" & _
                "<vrValidaConductor>1</vrValidaConductor>" & _
            "</ValidaConductor>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"

Private Const cSOAPFaC = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<Tiempo_Conductor xmlns=""http://localhost/"">" & _
                "<vrTiempo_Conductor>1</vrTiempo_Conductor>" & _
            "</Tiempo_Conductor>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"



Private Sub cmdC2F_Click()
    tipo = deCaF
    '
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para CaF
    parser.loadXML cSOAPCaF
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/ValidaConductor/vrValidaConductor").text = txtC.text
    '
    ' Mostrar el código XML enviado al servicio Web
    Text1.text = parser.xml
    '
    ' Usar el control Inet para realizar la operación HTTP POST
    'Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://elGuille/WebServices/ValidaConductor"
    '
    enviarComando parser.xml, "http://localhost/ValidaConductor"
    '
'    enviarComando
    '
End Sub

Private Sub cmdF2C_Click()
    tipo = deFaC
    '
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para FaC
    parser.loadXML cSOAPFaC
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/Tiempo_Conductor/vrTiempo_Conductor").text = txtF.text
    '
    ' Mostrar el código XML enviado al servicio Web
    Text1.text = parser.xml
    '
    ' Usar el control Inet para realizar la operación HTTP POST
    'Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://elGuille/WebServices/Tiempo_Conductor"
    '
    enviarComando parser.xml, "http://localhost/Tiempo_Conductor"
    '
'    enviarComando
    '
End Sub

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTPRequest
    Set oHttReq = New XMLHTTPRequest
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
    oHttReq.Open "POST", txtURL.text, False
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", sSoapAction
    ' enviar el comando
    oHttReq.send sXml
    '
    ' este será el texto recibido del servicio Web
    procesarRespuesta oHttReq.responseText
    '
End Sub

Private Sub procesarRespuesta(ByVal s As String)
    ' procesar la respuesta recibida del servicio Web
    Text1.text = s
    '
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.loadXML s
    '
    On Error Resume Next
    '
    If tipo = deCaF Then
        txtF.text = "Error"
        txtF.text = parser.selectSingleNode("/soap:Envelope/soap:Body/ValidaConductorResponse/ValidaConductorResult").text
    Else
        txtC.text = "Error"
        txtC.text = parser.selectSingleNode("/soap:Envelope/soap:Body/Tiempo_ConductorResponse/Tiempo_ConductorResult").text
    End If
    '
    If Err.Number > 0 Then
        Text1.SetFocus
    End If
End Sub




'al pelo con Microsoft Internet control
'Private Sub cmdC2F_Click()
'    tipo = deCaF
'    Dim parser As DOMDocument
'    Set parser = New DOMDocument
'    ' cargar el código SOAP para CaF
'    parser.loadXML cSOAPCaF
'    '
'    ' Indicar el parámetro a enviar
'    parser.selectSingleNode("/soap:Envelope/soap:Body/ValidaConductor/vrValidaConductor").Text = txtC.Text
'
'    ' Mostrar el código XML enviado al servicio Web
'    Text1.Text = parser.xml
'    ' Usar el control Inet para realizar la operación HTTP POST
'    Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://localhost/ValidaConductor"
'
'End Sub
'
'Private Sub cmdF2C_Click()
'    tipo = deFaC
'    Dim parser As DOMDocument
'    Set parser = New DOMDocument
'    ' cargar el código SOAP para FaC
'    parser.loadXML cSOAPFaC
'    '
'    ' Indicar el parámetro a enviar
'    parser.selectSingleNode("/soap:Envelope/soap:Body/Tiempo_Conductor/vrTiempo_Conductor").Text = txtF.Text
'    '
'    ' Mostrar el código XML enviado al servicio Web
'    Text1.Text = parser.xml
'    ' Usar el control Inet para realizar la operación HTTP POST
'    Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://localhost/Tiempo_Conductor"
'End Sub
'
'
'Private Sub Inet1_StateChanged(ByVal State As Integer)
'    '
'    If (State = icResponseCompleted) Then ' icResponseCompleted = 12
'        Dim s As String
'        '
'        ' Leer los datos devueltos por el servidor
'        s = Inet1.GetChunk(4096)
'        Text1.Text = s
'        '
'        ' Poner los datos en el analizador de XML
'        Dim parser As DOMDocument
'        Set parser = New DOMDocument
'        parser.loadXML s
'        '
'        On Error Resume Next
'        '
'        If tipo = deCaF Then
'            txtF.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/ValidaConductorResponse/ValidaConductorResult").Text
'        Else
'            txtC.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/Tiempo_ConductorResponse/Tiempo_ConductorResult").Text
'        End If
'        '
'        If Err.Number > 0 Then
'            Text1.SetFocus
'        End If
'    End If
'End Sub




'original
'Private Sub cmdC2F_Click()
'    tipo = deCaF
'    Dim parser As DOMDocument
'    Set parser = New DOMDocument
'    ' cargar el código SOAP para CaF
'    parser.loadXML cSOAPCaF
'    '
'    ' Indicar el parámetro a enviar
'    parser.selectSingleNode("/soap:Envelope/soap:Body/CaF/valor").Text = txtC.Text
'
'    ' Mostrar el código XML enviado al servicio Web
'    Text1.Text = parser.xml
'    ' Usar el control Inet para realizar la operación HTTP POST
'    Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://elGuille/WebServices/CaF"
'
'End Sub
'
'Private Sub cmdF2C_Click()
'    tipo = deFaC
'    Dim parser As DOMDocument
'    Set parser = New DOMDocument
'    ' cargar el código SOAP para FaC
'    parser.loadXML cSOAPFaC
'    '
'    ' Indicar el parámetro a enviar
'    parser.selectSingleNode("/soap:Envelope/soap:Body/FaC/valor").Text = txtF.Text
'    '
'    ' Mostrar el código XML enviado al servicio Web
'    Text1.Text = parser.xml
'    ' Usar el control Inet para realizar la operación HTTP POST
'    Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://elGuille/WebServices/FaC"
'End Sub
'
'
'Private Sub Inet1_StateChanged(ByVal State As Integer)
'    '
'    If (State = icResponseCompleted) Then ' icResponseCompleted = 12
'        Dim s As String
'        '
'        ' Leer los datos devueltos por el servidor
'        s = Inet1.GetChunk(4096)
'        Text1.Text = s
'        '
'        ' Poner los datos en el analizador de XML
'        Dim parser As DOMDocument
'        Set parser = New DOMDocument
'        parser.loadXML s
'        '
'        On Error Resume Next
'        '
'        If tipo = deCaF Then
'            txtF.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/CaFResponse/CaFResult").Text
'        Else
'            txtC.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/FaCResponse/FaCResult").Text
'        End If
'        '
'        If Err.Number > 0 Then
'            Text1.SetFocus
'        End If
'    End If
'End Sub
'**************************************************************************
'**************************************************************************
Private Sub Command1_Click()
Dim f As Form, CTR As Control, i As Double, j As Double
    Dim paso As Double, sNombre As String, x As Double
    
    G_modulo = "CND"
    paso = 0
    'MANFPAGO =  CO  -  CR   -   TO
    For j = 0 To Forms.Count - 1
        For i = 0 To Forms(j).Controls.Count - 1
            If TypeOf Forms(j).Controls(i) Is CommandButton Or TypeOf Forms(j).Controls(i) Is Menu Then
                 
                paso = 1
                
                If Forms(j).Controls(i).Caption = "-" Then paso = 0
                If Forms(j).Controls(i).Caption = "&Salir" Then paso = 0
                If Forms(j).Controls(i).Caption = "&Limpiar" Then paso = 0
                
                x = InStr(1, Forms(j).Controls(i).Caption, "&", vbTextCompare)
                    
                If x > 0 Then
                    sNombre = Mid(Forms(j).Controls(i).Caption, 1, x - 1) & Mid(Forms(j).Controls(i).Caption, x + 1)
                Else
                    sNombre = Forms(j).Controls(i).Caption
                End If
                
                If TypeOf Forms(j).Controls(i) Is CommandButton Then sNombre = "Botón " & sNombre
                If TypeOf Forms(j).Controls(i) Is Menu Then sNombre = "Menú " & sNombre
                
                If paso = 1 Then
                    QRY = "INSERT INTO seobjmod (" & _
                                            "omcodmod, omnomfrm, omdesfrm, omnomobj, " & _
                                            "omdesobj) " & _
                               "VALUES ('" & _
                                        G_modulo & "','" & Forms(j).Name & "','" & "PANTALLA PRINCIPAL" & "','" & Forms(j).Controls(i).Name & "','" & _
                                        sNombre & "')"
                    On Error Resume Next
                    Dbs.Execute QRY
                End If
                
            End If
        Next i
    Next j
    
    MsgBox "Proceso Terminado ", 64, "OK"
    Screen.MousePointer = 0
End Sub



Private Sub Form_Load()
   On Error Resume Next
 ' Lblciudad.Caption = "De" & " " & CIUDAD & " " & "S.A."
  main.Caption = "TERMINAL DE TRANSPORTE S.A. - MODULO TASAS DE USO"
  G_Carga = "0"
  
'  Text1.Text = "Cliente VB6" & vbCrLf & "del Servicio Web Conversor de ºC a ºF"
''  txtURL.Text = "http://integraocl.fitcloud.com.co/swinsertupdate/integra.asmx"
'  'http://localhost/Integra/integra.asmx"
'  'http://190.85.249.35/webservices/integra.asmx"
'  'http://localhost:56256/Integra.asmx"
'  'http://localhost/Integra/integra.asmx"
'  'http://localhost/Conversor.asmx"
'  'http://www.elguille.info/Net/WebServices/conversor.asmx
L_Bandera_CierreAutomatico = 0
End Sub

Private Sub Form_Paint()
 If G_Carga = "0" Then
        main.Visible = False
        ImgFondo.Top = 0
        ImgFondo.Left = 0
        ImgFondo.Width = main.Width
        ImgFondo.Height = main.Height - 1000
        Label1.BorderStyle = 0
        Label1.Left = (main.Width / 2) - (Label1.Width / 2)
        Label1.Top = main.Height - 3000
        
        main.Visible = True
        G_Carga = "1"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
  'wecc 08/06/2002
  'controla que el sistemas se cierre sin haber
  'cerrado dia
  On Error Resume Next
  If G_cierredia = "S" Then
     'If G_turno = G_maxtur Then
     '  If Mnuapetur.Enabled = False And Mnucietur.Enabled = True Then
     '     MsgBox "      No se ha cerrado día      ", 48, "Advertencia"
     '     Cancel = 1
     '     Exit Sub
     '  End If
     'End If
  End If
  If Login <> "" Then
     sFecha = Format(Now, "mm/dd/yyyy hh:mm")
     Dbs.Execute " UPDATE gelogusu " _
          & "SET lufecsal = to_date('" & sFecha & "','mm/dd/yyyy hh24:mi')," _
          & "    luterm = '" & Pc & "' " _
          & "WHERE lucodusu = '" & Login & "' " _
          & "AND lufecsal is null"
          
   End If
   
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   Unload inicio
   Dbs.Close
   dbsX.Close
   Unload Me
   End
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error Resume Next
   frmmodulos.Visible = True
End Sub

Private Sub mnuacercade_Click()
  On Error GoTo Errores
  frmAbout.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
End Sub

Private Sub Mnuajuste_Click()
  On Error GoTo Errores
  Aumtarifas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuAnuCupos_Click()
   On Error GoTo Errores
   Anucupos.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuAnuFactura_Click()
   On Error GoTo Errores
   anufactu.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuanumultas_Click()
  On Error GoTo Errores
  Anumultas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuapetur_Click()
   On Error GoTo Errores
   Cuaturno.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuasihorrut_Click()
   On Error GoTo Errores
   Crehoemp.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Mnucambiopassword_Click()
   On Error GoTo Errores
   password.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuCambioServicio_Click()
   On Error GoTo Errores
   
   camservicio.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnucancon_Click()
   On Error GoTo Errores
   
   G_Clase_Autorizacion = "ANUTASA"
   Relogin.Show 1
   If G_Aut_Vender <> "SI" Then
      Exit Sub
   End If
      
   Anucondu.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnucanmul_Click()
   On Error GoTo Errores
   Canmultas.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnucietur_Click()
   On Error GoTo Errores
   Cuacajdia.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnucolaemp_Click()
   On Error GoTo Errores
   
   colaempr.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
End Sub

Private Sub mnuconaditt_Click()
    On Error GoTo Errores
    conaditt.Show 1
    
    Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConAudAviPar_Click()
  On Error GoTo Errores
  conaudpa.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number

End Sub

Private Sub Mnuconaumtar_Click()
  On Error GoTo Errores
  Conaumtarifas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuconconcri_Click()
   On Error GoTo Errores
   Conconcriterio.Show 1
  
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuconconduces_Click()
   On Error GoTo Errores
   Conconduces.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConDesOriNorte_Click()
On Error GoTo Errores
   
   conrutnorte.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConDesPar_Click()
   On Error GoTo Errores
   
   condespar.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConDeuParqueadero_Click()
   On Error GoTo Errores
   
   ParCondo.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConductor_Click()
   On Error GoTo Errores
   
   conductor.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconempcol_Click()
   On Error GoTo Errores
   
   conempco.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConEmpConvenios_Click()
   On Error GoTo Errores
   conconve.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconemprut_Click()
   On Error GoTo Errores
   Conempto.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConfiFormatos_Click()
   On Error GoTo Errores
   Formatos.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConfiPresentacion_Click()
   On Error GoTo Errores
   confpres.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub MnuConfiVideos_Click()
   On Error GoTo Errores
   confvide.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub mnuconfpant_Click()
   On Error GoTo Errores
   pantalla.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConLisVehCenSur_Click()
   On Error GoTo Errores
   concensur.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConMovCupos_Click()
   On Error GoTo Errores
   rptmovimientos.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconmul_Click()
   On Error GoTo Errores
   Conmultas.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConParqueadero_Click()
   On Error GoTo Errores
   Conparqu.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub
Private Sub MnuEvaTransito_Click()
    On Error GoTo Errores
    frmEvasionesTerminal.Show 1
    
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconpasc_Click()
   On Error GoTo Errores
   conpasaj.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconpassal_Click()
   On Error GoTo Errores
   cosalpas.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconpen_Click()
   On Error GoTo Errores
   coconpen.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConReserva_Click()
   On Error GoTo Errores
   conreser.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number

End Sub

Private Sub mnuconsignaciones_Click()
   On Error GoTo Errores
   conconsi.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconsolidado_Click()
   On Error GoTo Errores
   consolidado.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub



Private Sub MnuConsultaCupos_Click()
   On Error GoTo Errores
   confactu.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConTasUso_Click()
   On Error Resume Next
   convalid.Show 1
End Sub

Private Sub mnucontraemp_Click()
cocontra.Show 1
End Sub

Private Sub mnuconvehdia_Click()
   On Error GoTo Errores
   convehdi.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconvehto_Click()
   On Error GoTo Errores
   convehto.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconvemp_Click()
   On Error GoTo Errores
    convempr.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuconvenios_Click()
  On Error GoTo Errores
   convenios.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnucrerut_Click()
  On Error GoTo Errores
  Crerutas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuDesAviParq_Click()
   On Error GoTo Errores
   paraviso.Show 1
   Exit Sub
Errores:
   Exit Sub
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub MnuDesEdu_Click()
   On Error GoTo Errores
   Deseduca.Show 1
  
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuDesVenTut_Click()
   On Error GoTo Errores
   cosalveh.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuEmpConvenios_Click()
   On Error GoTo Errores
   EmprConv.Show 1
  
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuempresas_Click()
  On Error GoTo Errores
  Conempresas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuentfic_Click()
   On Error GoTo Errores
   entficha.Show 1
  
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuentsinfic_Click()
   On Error GoTo Errores
   
   conlogfi.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub



Private Sub MnuEquivalencias_Click()
   On Error GoTo Errores
   FrEquiva.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub MnuGenPla_Click()
   On Error Resume Next
   Genplano.Show 1
End Sub

Private Sub Mnuhojvidemp_Click()
  On Error GoTo Errores
  Empresas.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuhojvidveh_Click()
  On Error GoTo Errores
  Vehiculo.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuinfoemp_Click()
   On Error GoTo Errores
   conconemp.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuingconpen_Click()
   On Error GoTo Errores
   pendient.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuIngresoCupos_Click()
   On Error GoTo Errores
   cocupos.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuIngresoFactura_Click()
   On Error GoTo Errores
   cofaccup.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub


Private Sub Mnuingvehiculos_Click()
  On Error GoTo Errores
   coningveh.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnulogin_Click()
   On Error GoTo Errores
   flogin.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnulogoff_Click()
   'wecc 08/06/2002
  'controla que el sistemas se cierre sin haber
  'cerrado dia
   On Error Resume Next
   If G_cierredia = "S" Then
      'If CDbl(G_turno) = CDbl(G_maxtur) Then
      '   If Mnuapetur.Enabled = False And Mnucietur.Enabled = True Then
      '      MsgBox "      No se ha cerrado día      ", 48, "Advertencia"
      '      Exit Sub
      '   End If
      'End If
   End If
   
'   If Login <> "" Then
'     Dbs.Execute " UPDATE gelogusu " _
'                   & "SET lufecsal = '" & Now & "'," _
'                   & "    luterm = '" & Pc & "' " _
'                 & "WHERE lucodusu = '" & Login & "' " _
'                   & "AND lufecsal is null"
'   End If
  
   Unload Me
End Sub

Private Sub mnumulcan_Click()
   On Error GoTo Errores
   comulcan.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnumulcri_Click()
   On Error GoTo Errores
   Conmulcri.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuOperadorEmpresa_Click()
   On Error GoTo Errores
   operempr.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub

End Sub

Private Sub MnuPicoyPlaca_Click()
   On Error GoTo Errores
   picoyplaca.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuPLataformas_Click()
   On Error GoTo Errores
   Platafor.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuprecon_Click()
   On Error GoTo Errores
   conprecond.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnurecemp_Click()
   On Error GoTo Errores
   Concontotal.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuregmul_Click()
   On Error GoTo Errores
   Regmultas.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuregtag_Click()
    On Error GoTo Errores
        Regtag.Show 1
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub Mnuregpas_Click()
   On Error GoTo Errores
   Ingpasaj.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuReImprimirTU_Click()
   On Error GoTo Errores
   reimprimir.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRepNovedades_Click()
   On Error GoTo Errores
   repnoved.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRepSan_Click()
   cosanmul.Show 1
End Sub

Private Sub Mnuresultados_Click()
   On Error GoTo Errores
   Conalcolemia.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRptCompElectronico_Click()
   On Error GoTo Errores
   rptcomelec.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRptRodamiento_Click()
   On Error GoTo Errores
   RptRodamiento.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRptTasAutNorte_Click()
   On Error GoTo Errores
   Rpttasanorte.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuRutAdicionales_Click()
   On Error GoTo Errores
   RutAdicionales.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnurutcri_Click()
   On Error GoTo Errores
   concrirut.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnurutemp_Click()
   On Error GoTo Errores
   
   Creruthor.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Mnuruttipveh_Click()
   On Error GoTo Errores
   Ruttipveh.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuReportesRut_Click()
    On Error GoTo Errores
    rptrutas.Show 1
    
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub


Private Sub mnusalvehpas_Click()
   On Error GoTo Errores
   coemruse.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnutasdirecc_Click()
   On Error GoTo Errores
   contasdi.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuTasSur_Click()
   On Error GoTo Errores
   opconsur.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number

End Sub

Private Sub Mnuturnos_Click()
   On Error GoTo Errores
   Concietur.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuUsuarioEmpresa_Click()
   On Error GoTo Errores
   FrUsuEmpresa.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuvalalc_Click()
   On Error GoTo Errores
   convalal.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuValTasUso_Click()
   On Error Resume Next
   covalida.Show 1
End Sub

Private Sub MnuVehRes_Click()
   On Error GoTo Errores
   covehres.Show 1
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Mnuvencon_Click()
  On Error GoTo Errores
  Capconduce.Show 1
  
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuvensal_Click()
   On Error GoTo Errores
   covensal.Show 1
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnutagcriterio_Click()
    
    On Error GoTo Errores
    Contagcriterio.Show 1
    
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuemgaf_Click()
    On Error GoTo Errores
    coemgaf.Show 1
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnurepalcopruebas_Click()
    On Error GoTo Errores
    coalcopruebas.Show 1
    Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnufectu_Click()
    On Error GoTo Errores
    coconfec.Show 1
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub mnuentsalpro_Click()
    On Error GoTo Errores
    coentsalprot.Show 1
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub mnurecalsaldos_Click()
    On Error GoTo Errores
        corecasaldo.Show 1
    Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub Mnutucondicional_Click()
    On Error GoTo Errores
        cocondicional.Show 1
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub MnuConVehiculos_Click()
    On Error GoTo Errores
        ConVehiculos.Show 1
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuConConductores_Click()
    On Error GoTo Errores
        ConConductores.Show 1
    Exit Sub
        
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub MnuCoconconcep_Click()
    On Error GoTo Errores
        Coconconcep.Show 1
    Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Timer1_Timer()
    If G_Turno = "4" Then
        If L_Bandera_CierreAutomatico = 0 Then
            QRY = "SELECT psval FROM geparsis WHERE pscod = 'HORASIS' "
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                horven = Format(Tbl1!psval, "hh:mm")
                If (CDate(horven) >= CDate("00:00") And CDate(horven) <= CDate("00:05")) Then
                    'Btnsalir_Click
                    G_CierreAutomatico = 1
                    Cuaturno.Visible = False
                    Cuaturno.Show 1
                    Capconduce.Visible = False
                    G_CierreAutomatico = 0
                    L_Bandera_CierreAutomatico = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   On Error Resume Next
   If Button.index = 1 Then
      flogin.Show 1
   End If
     
   If Button.index = 3 Then
     Vehiculo.Show 1
   End If
     
   If Button.index = 4 Then
     Empresas.Show 1
   End If
     
   If Button.index = 6 Then
      'cierre turno
      Cuaturno.Show 1
   End If
   
   If Button.index = 7 Then
       'cierre dia
       Cuacajdia.Show 1
   End If
     
   If Button.index = 8 Then
      'venta conduces
      On Error GoTo Errores
      Capconduce.Show 1
   End If
     
   If Button.index = 9 Then
      'anulacion de conduces
      G_Clase_Autorizacion = "ANUTASA"
      Relogin.Show 1
      If G_Aut_Vender <> "SI" Then
         Exit Sub
      End If
      Anucondu.Show 1
   End If
   If Button.index = 11 Then
   'registro de multas
      Regmultas.Show 1
   End If
   If Button.index = 12 Then
   'cancelacion de multas
      Canmultas.Show 1
   End If
   If Button.index = 13 Then
   'anulacion de multas
      Anumultas.Show 1
   End If
   If Button.index = 15 Then
   'consolidacion
      consolidado.Show 1
   End If
   
   If Button.index = 17 Then
      On Error Resume Next
      If G_cierredia = "S" Then
         'If CDbl(G_turno) = CDbl(G_maxtur) Then
         '   If Mnuapetur.Enabled = False And Mnucietur.Enabled = True Then
         '      MsgBox "      No se ha cerrado día      ", 48, "Advertencia"
         '      Exit Sub
         '   End If
         'End If
      End If
      sFecha = Format(Now, "MM/dd/yyyy hhmm:")
      If Login <> "" Then
         Dbs.Execute " UPDATE gelogusu " _
                       & "SET lufecsal = to_date('" & sFecha & "','mm/dd/yyyy hh24:mi')," _
                       & "    luterm = '" & Pc & "' " _
                     & "WHERE lucodusu = '" & Login & "' " _
                     & "AND lufecsal is null"
          
      End If
  
      Unload Me
      
      End
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub
