VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Concontotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Tasas de Uso Totalizada"
   ClientHeight    =   7305
   ClientLeft      =   2730
   ClientTop       =   1725
   ClientWidth     =   7290
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Tipo Tasa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   4200
      TabIndex        =   37
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton OptTipoTasa 
         Caption         =   "Vendidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptTipoTasa 
         Caption         =   "Cupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Terminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   2160
      TabIndex        =   36
      Top             =   2040
      Width           =   1935
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "Conconto.frx":0000
         Left            =   120
         List            =   "Conconto.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Tipos de terminales"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Width           =   1935
      Begin VB.TextBox txtcaja 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Nombre de la caja"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Datos Adicionales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Width           =   7095
      Begin MSMask.MaskEdBox Msktotcon 
         Height          =   300
         Left            =   5280
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskalcoholemia 
         Height          =   300
         Left            =   1320
         TabIndex        =   25
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msktasas 
         Height          =   300
         Left            =   5280
         TabIndex        =   26
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskparqueadero 
         Height          =   300
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Caption         =   "Tasas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Alcoholemia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Tasas Vendidas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Total General :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Lblconven 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Parqueadero :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   6135
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   6
         ToolTipText     =   "Empresa"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lblcodemp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   4812
      End
   End
   Begin VB.CommandButton Btnsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "Conconto.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Btnimprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "Conconto.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Btnbuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "Conconto.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Btnlimpiar 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "Conconto.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   2412
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   6135
      Begin Crystal.CrystalReport rptterminal 
         Left            =   240
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdcontot 
         Height          =   2052
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   5652
         _ExtentX        =   9975
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Opttipo 
         Caption         =   "Personalizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Opttipo 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton OptFecha 
         Caption         =   "Turno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         ToolTipText     =   "Fecha final de búsqueda "
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Fecha inicial de búsqueda"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   35
      Top             =   7080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "Concontotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim origen As Integer
Dim totori As Double
Dim transito As Integer
Dim tottra As Double
Dim TITULO As String
Dim LSaldoMinimoConsignacion As Double
Dim LValorMinimoConsignacion As Double
Dim LSaldoMaximoConsignacion As Double
Dim LValorMaximoConsignacion As Double
Dim sCodigo As String
'Sebastián Rondón - 30 Sept - 2022
'Conceptos multas
Dim L_CiTElusion As String
Dim L_CiTEvasion As String
Dim L_CiTPerdida As String
Dim L_CiTExcreta As String
Dim L_CiTEvasTra As String
Dim L_Todas_Multas As String

'Variables para las multas
Dim Elusion As Double
Dim Evasion As Double
Dim EvasionTransito As Double
Dim PerDocumento As Double
Dim Permanencias As Double
Dim Vigencias As Double
Dim TotalMultas As Double



Private Sub Valida_Posicion_Consignacion(ByVal sFecha1 As String, ByVal sFecha2 As String, ByVal sCodemp As String)
   
   'SI LA CONSIGNACION FUE EL PRIMER MOVIMIENTO QUE SE HIZO DE LA EMPRESA
   QRY = "SELECT mosaldo, movalor " _
         & "FROM opmovimientos " _
        & "WHERE monumero = (SELECT MIN(monumero) " _
                         & "FROM opmovimientos " _
                         & "WHERE mofecha BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy') AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy') " _
                         & "AND mocodemp = '" & sCodemp & "') " _
         & "AND motipmov = 'DB' AND motipo = 'FACT' AND motipcupo = 'E'"
         
    
   Set Tbl2 = Dbs.Execute(QRY)
   LSaldoMinimoConsignacion = 0
   LValorMinimoConsignacion = 0
   LSaldoMaximoConsignacion = 0
   LValorMaximoConsignacion = 0

   If Not Tbl2.EOF Then
      LSaldoMinimoConsignacion = Tbl2!mosaldo
      LValorMinimoConsignacion = Tbl2!movalor
   End If
   Tbl2.Close
   
   'CONSIGNACIONES ANULADAS PARA RESTAR DE LAS CONSIGNACIONES HECHAS
   QRY = "SELECT mosaldo, movalor " _
         & "FROM opmovimientos " _
        & "WHERE monumero = (SELECT MIN(monumero) " _
                         & "FROM opmovimientos " _
                         & "WHERE mofecha BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy') AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy') " _
                         & "AND mocodemp = '" & sCodemp & "') " _
         & "AND motipmov = 'CR' AND motipo = 'AFAC' AND motipcupo = 'E'"
         
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      LValorMinimoConsignacion = LValorMinimoConsignacion - Tbl2!movalor
   End If
   Tbl2.Close
   
   'SI LA CONSIGNACION FUE EL ULTIMO MOVIMIENTO QUE SE HIZO DE LA EMPRESA
   QRY = "SELECT mosaldo, movalor " _
         & "FROM opmovimientos " _
        & "WHERE monumero = (SELECT MAX(monumero) " _
                         & "FROM opmovimientos " _
                         & "WHERE mofecha BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy') AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy') " _
                         & "AND mocodemp = '" & sCodemp & "') " _
         & "AND motipmov = 'DB' AND motipo = 'FACT' AND motipcupo = 'E'"
         
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      LSaldoMaximoConsignacion = Tbl2!mosaldo
      LValorMaximoConsignacion = Tbl2!movalor
   End If
   Tbl2.Close

   'CONSIGNACIONES ANULADAS PARA RESTAR DE LAS CONSIGNACIONES HECHAS
   QRY = "SELECT mosaldo, movalor " _
         & "FROM opmovimientos " _
        & "WHERE monumero = (SELECT MAX(monumero) " _
                         & "FROM opmovimientos " _
                         & "WHERE mofecha BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy') AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy') " _
                         & "AND mocodemp = '" & sCodemp & "') " _
         & "AND motipmov = 'CR' AND motipo = 'AFAC' AND motipcupo = 'E'"
         
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      LValorMaximoConsignacion = LValorMaximoConsignacion - Tbl2!movalor
   End If
   Tbl2.Close




End Sub




Private Sub X_Turno(ByVal Fecha As String, ByVal Fecha1 As String)
   If OptTipoTasa(0).Value = True Then
      'cupo
      QRY1 = "P"
   ElseIf OptTipoTasa(1).Value = True Then
      'vendido
      QRY1 = "D"
   End If
   
   QRY = "SELECT DISTINCT(cinomcaj) AS cinomcaj, MIN(TO_NUMBER(cinumini)) AS cinumini, MAX(TO_NUMBER(cinumfin)) AS cinumfin " _
         & "FROM cocietur " _
         & "WHERE cifecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
           & "AND cinumini <> '0' "
           
   If txtcaja.text <> "" Then
      QRY = QRY & "AND cinomcaj = '" & Trim(txtcaja.text) & "'"
   End If
   QRY = QRY & " GROUP BY cinomcaj " _
             & " ORDER BY cinomcaj "
   
   Set Ftb02 = Dbs.Execute(QRY)
   Concontotal.Enabled = False

   Do While Not Ftb02.EOF
      ProBar.Value = ProBar.Min
      If Not IsNull(Ftb02!cinumini) Then
         vmin = Ftb02!cinumini
         vmax = Ftb02!cinumfin
      Else
         vmin = ""
         vmax = ""
      End If
      If vmin <> "" Then
         QRY = "SELECT DISTINCT(cocodemp) AS cocodemp, henombre, SUM(covalor) AS covalor, count(conumero) AS CantTasas " _
               & "FROM coconduc INNER JOIN ophojemp ON cocodemp = hecodigo " _
              & "WHERE conumero BETWEEN " & vmin & " " _
                & "AND " & vmax & " " _
                & "AND conomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                & "AND coestado = 'Activo' "
         
         If OptTipo(1).Value = True Then
            QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
         End If
         If CmbTerminal.ListIndex <> 0 Then
            'ha escogido un terminal
            QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
         End If
         
         QRY = QRY & "AND comodo = '" & QRY1 & "' GROUP BY cocodemp, henombre "
         
         Set Tbl1 = Dbs.Execute(QRY)
         If Not Tbl1.EOF Then
            Tbl1.MoveLast
            ProBar.Value = ProBar.Min
            ProBar.Max = Tbl1.RecordCount + 2
            Tbl1.MoveFirst
         End If
         
         
         
         Do While Not Tbl1.EOF
            DoEvents
            
            '********************************************* Tasas de Uso Origen *******************************************
            fqry01 = "SELECT count(cooritra) as ori, sum(covalor) as totori, sum(covalalc) as totalcori " _
                  & "FROM coconduc " _
                 & "WHERE conumero BETWEEN " & vmin & " " _
                   & "AND " & vmax & " " _
                   & "AND cooritra = 'O' " _
                   & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                   & "AND conomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                   & "AND coestado = 'Activo' " _
                   & "AND comodo = '" & QRY1 & "' "
         
            If CmbTerminal.ListIndex <> 0 Then
               'ha escogido un terminal
               fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
            End If
            
            Set Ftb01 = Dbs.Execute(fqry01)
              
            If Not Ftb01.EOF Then
               If Not Ftb01!ori And Not Ftb01!totori Then
                  origen = Ftb01!ori
                  totori = Ftb01!totori
                  totalcori = Ftb01!totalcori
               Else
                  origen = 0
                  totori = 0
                  totalcori = 0
               End If
            End If
            Ftb01.Close
            '**************************************************************************************************************
            
            '******************************************** Tasas de Uso Transito *******************************************
            fqry01 = "SELECT count(cooritra) as tra, sum(covalor) as tottra, sum(covalalc) as totalctra " _
                     & "FROM coconduc " _
                    & "WHERE conumero BETWEEN " & vmin & " " _
                      & "AND " & vmax & " " _
                      & "AND cooritra = 'T' " _
                      & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                      & "AND conomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                      & "AND coestado = 'Activo' " _
                      & "AND comodo = '" & QRY1 & "' "
                      
            If CmbTerminal.ListIndex <> 0 Then
               'ha escogido un terminal
               fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
            End If
                  
            Set Ftb01 = Dbs.Execute(fqry01)
            
            If Not Ftb01.EOF Then
               If Not Ftb01!tra And Not Ftb01!tottra Then
                  transito = Ftb01!tra
                  tottra = Ftb01!tottra
                  totalctra = Ftb01!totalctra
               Else
                  transito = 0
                  tottra = 0
                  totalctra = 0
               End If
            End If
            Ftb01.Close
            '**************************************************************************************************************
            
            '************************************************* Parqueadero ************************************************
            fqry01 = "SELECT sum(coparque) as totpar " _
                     & "FROM coconduc " _
                    & "WHERE conumero BETWEEN " & vmin & " " _
                      & "AND " & vmax & " " _
                      & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                      & "AND conomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                      & "AND coestado = 'Activo' " _
                      & "AND comodo = '" & QRY1 & "' "
            
            If CmbTerminal.ListIndex <> 0 Then
               'ha escogido un terminal
               fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
            End If
            
            Set Ftb01 = Dbs.Execute(fqry01)
                    
            If Not Ftb01.EOF Then
               If Not Ftb01!totpar Then
                  parque = Ftb01!totpar
               Else
                  parque = 0
               End If
            End If
            
            Ftb01.Close
            '**************************************************************************************************************
            
            '********************************************************** Multas **************************************************
      
            'Verificamos si hay registros de ingresos
            fqry01 = "SELECT DISTINCT(abnumero) as nummul " _
                        & "FROM opabonos " _
                       & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy')  " _
                         & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                         & "AND abcerrado = 'S' " _
                         & "AND abnomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                         & "AND abestado = 'Activo' "
              
               Set Ftb03 = Dbs.Execute(fqry01)
               
               'Si hay registros buscamos concepto por concepto
               If Ftb03.EOF = False Then
                  For k = 1 To 4
                      Select Case k
                          Case 1
                              sCodigo = L_CiTElusion
                          Case 2
                              sCodigo = L_CiTEvasion
                          Case 3
                              sCodigo = L_CiTPerdida
                          Case 4
                              sCodigo = L_CiTEvasTra
                      End Select
                      
                      fqry01 = "SELECT SUM(abvalor) as valt " _
                              & "FROM opabonos, opmultas " _
                             & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                               & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                               & "AND munumero = abnumero  " _
                               & "AND abcerrado = 'S'  " _
                               & "AND mucodinf IN (" & sCodigo & ") " _
                               & "AND mucodemp = '" & Tbl1!cocodemp & "' " _
                               & "AND abnomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                               & "AND abestado = 'Activo' "
                               
                      
                      
                     Set Ftb03 = Dbs.Execute(fqry01)
                     
                     If Not Ftb03.EOF Then
                        If Not IsNull(Ftb03!valt) Then
                           Select Case k
                              Case 1
                                 Elusion = Ftb03!valt
                              Case 2
                                 Evasion = Ftb03!valt
                              Case 3
                                 PerDocumento = Ftb03!valt
                              Case 4
                                 EvasionTransito = Ftb03!valt
                           End Select
                        End If
                     End If
                     Ftb03.Close
                  Next k
                  
                  'Ahora calculamos el total de multas
                  fqry01 = "SELECT SUM(abvalor) as valt " _
                           & "FROM opabonos, opmultas " _
                            & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                            & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                            & "AND munumero = abnumero  " _
                            & "AND abcerrado = 'S'  " _
                            & "AND mucodinf IN (" & L_Todas_Multas & ") " _
                            & "AND mucodemp = '" & Tbl1!cocodemp & "' " _
                            & "AND abnomcaj = '" & Trim(Ftb02!cinomcaj) & "' " _
                            & "AND abestado = 'Activo' "
                  
                  Set Ftb03 = Dbs.Execute(fqry01)
                      
                  If Not Ftb03.EOF Then
                     If Not IsNull(Ftb03!valt) Then
                        TotalMultas = Ftb03!valt
                     End If
                  End If
               Else
                  Elusion = 0
                  Evasion = 0
                  PerDocumento = 0
                  EvasionTransito = 0
                  TotalMultas = 0
               End If
               Ftb03.Close
      
      '********************************************************************************************************************
      
       '************************************************ VIGENCIAS *********************************************************
      
      fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueop " _
              & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
              & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
              & "  AND paestado = 'C' " _
              & "  AND pacerrado = 'S' " _
              & "  AND patipo = 'V' " _
              & "  AND pacodemp = '" & Tbl1!cocodemp & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
            vigencia = Ftb01!valpar
         Else
            vigencia = 0
         End If
      End If
    Ftb01.Close
    
    '**********************************************************************************************************************
      
            
            sTotal = CDbl(totori) + CDbl(totalcori) + CDbl(tottra) + CDbl(totalctra) + CDbl(parque) + CDbl(TotalMultas) + CDbl(vigencia)
                     
            QRY = "UPDATE optempo1 " _
                   & "SET temnumero1 = temnumero1 + " & origen & ", " _
                      & " temnumdob3 = temnumdob3 + " & totori & ", " _
                      & " temnumero3 = temnumero3 + " & totalcori & ", " _
                      & " temnumero2 = temnumero2 + " & transito & ", " _
                      & " temnumdob4 = temnumdob4 + " & tottra & ", " _
                      & " temnumdob5 = temnumdob5 + " & totalctra & ", " _
                      & " temnumdob1 = temnumdob1 + " & Tbl1!covalor & ", " _
                      & " temnumdob2 = temnumdob2 + " & sTotal & ", " _
                      & " temnumero6 = temnumero6 + " & parque & ", " _
                      & " temnumero4 = temnumero4 + " & Tbl1!CantTasas & ", " _
                      & " temnumero5 = temnumero5 + " & Elusion & ", " _
                      & " temnumero11 = temnumero11 + " & Evasion & ", " _
                      & " temnumero12 = temnumero12 + " & PerDocumento & ", " _
                      & " temnumero13 = temnumero13 + " & EvasionTransito & ", " _
                      & " temnumero14 = temnumero14 + " & TotalMultas & ", " _
                    & " temnumero15 = temnumero15 + " & vigencia & " " _
                 & "WHERE temtext1 = '" & Tbl1!cocodemp & "' " _
                   & "AND tempc ='" & Pc & "' "
            
            Dbs.Execute QRY, sRecordAfecta
'            entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Str(origen) + Chr(9) + _
'                 Str(totori) + Chr(9) + Str(totalcori) + Chr(9) + Str(transito) + Chr(9) + Str(tottra) + Chr(9) + _
'                 Str(totalctra) + Chr(9) + Str(Tbl1(2)) + Chr(9) + Str(sTotal) + Chr(9) + _
'                  Str(parque)
            If sRecordAfecta = 0 Then
               QRY = "INSERT INTO optempo1 " _
               & "(temtext1, temtext2, temnumero1, temnumdob3, " _
               & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
               & " temnumdob1, temnumdob2, temnumero6, temnumero4, temnumero5, temnumero11, temnumero12, " _
               & " temnumero13, temnumero14, temnumero15, tempc) VALUES ('" _
               & Tbl1!cocodemp & "','" & Tbl1!henombre & "'," & origen & "," & totori & "," _
               & totalcori & "," & transito & "," & tottra & "," & totalctra & "," _
               & Tbl1!covalor & "," & sTotal & "," & parque & ",'" & Tbl1!CantTasas & "'," & Elusion & "," _
               & Evasion & "," & PerDocumento & "," & EvasionTransito & "," & TotalMultas & "," & vigencia & ",'" & Pc & "')"
   
               Dbs.Execute QRY
            
            End If
            
       
            
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
         Loop
         Tbl1.Close
      End If
      
      Ftb02.MoveNext
   Loop
   Ftb02.Close
End Sub

Private Sub X_Dia(ByVal Fecha As String, ByVal Fecha1 As String)

    Dim origen As Double
    Dim totori As Double
    Dim totalcori As Double
    
    Dim transito As Double
    Dim tottra As Double
    Dim totalctra As Double
    
    On Error GoTo Err_XDia
    
    no = "N"
    
   If OptTipoTasa(0).Value = True Then
      'cupo
      QRY1 = "P"
   ElseIf OptTipoTasa(1).Value = True Then
      'vendido
      QRY1 = "D"
   End If
   
   'Sebastian Rondón - 19 Octubre del 2022
   'Por petición del ingeniero Nelson, si se selecciona que el tipo de tasa sea cupo
   'Se debe mostrar el reporte anterior sin ningún cambio
   
   If OptTipoTasa(0).Value = True Then
'        QRY = "SELECT DISTINCT(mocodemp) AS cocodemp, henombre, henit, SUM(movalor) AS covalor, count(motasauso) AS CantTasas " _
'            & "FROM opmovimientos " _
'            & "INNER JOIN ophojemp " _
'            & "ON mocodemp = hecodigo " _
'            & "WHERE mofecha BETWEEN TO_DATE('" & Fecha & "','MM/DD/YYYY') " _
'            & "AND TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
'            & "AND motipo = 'VTTU' " _
'            & "GROUP BY mocodemp, henombre, henit "
   
        'Buscamos las empresas
        QRY = "SELECT DISTINCT(cocodemp) AS cocodemp, henombre, henit, SUM(covalor) AS covalor, count(conumero) AS CantTasas " _
         & "FROM coconduc INNER JOIN ophojemp ON cocodemp = hecodigo " _
         & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
         & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') "

         If txtcaja.text <> "" Then
            QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
         End If

         If OptTipo(1).Value = True Then
             QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
         End If
         If CmbTerminal.ListIndex <> 0 Then
             QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
         End If

         QRY = QRY & "AND coestado = 'Activo' AND comodo = '" & QRY1 & "' GROUP BY cocodemp, henombre, henit "
         
   Else
   
        '*** PARTE COCONDUC ***
        QRY = "SELECT DISTINCT(cocodemp) AS cocodemp, SUM(covalor) AS covalor, SUM(CantTasas) AS CantTasas FROM( " _
         & "SELECT DISTINCT(cocodemp) AS cocodemp, SUM(covalor) AS covalor, count(conumero) AS CantTasas " _
         & "FROM coconduc INNER JOIN ophojemp ON cocodemp = hecodigo " _
         & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
         & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') "
    
       If txtcaja.text <> "" Then
          QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
       End If
       
       If OptTipo(1).Value = True Then
          QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
       End If
       If CmbTerminal.ListIndex <> 0 Then
          QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
       End If
       
       QRY = QRY & "AND coestado = 'Activo' GROUP BY cocodemp, henombre "
       
       '*** PARTE MULTAS ***
       QRY = QRY & "UNION SELECT DISTINCT(mucodemp) AS cocodemp, 0 AS covalor, 0 AS CantTasas " _
         & "FROM opabonos, opmultas " _
         & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
         & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
         & "AND munumero = abnumero " _
         & "AND abcerrado = 'S' " _
         & "AND mucodinf IN (" & L_Todas_Multas & ") " _
         & "AND abestado = 'Activo' "
         
        If txtcaja.text <> "" Then
            QRY = QRY & "AND abnomcaj = '" & txtcaja.text & "' "
        End If
        
        If OptTipo(1).Value = True Then
          QRY = QRY & "AND mucodemp = '" & Txtcodemp.text & "' "
       End If
       
       '*** PARTE PERMANENCIAS ***
       QRY = QRY & "UNION SELECT DISTINCT pacodemp AS cocodemp, 0 AS covalor, 0 AS CantTasas " _
         & "FROM coparqueop " _
         & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
         & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
         & "AND patipo = 'P' "
         
        If txtcaja.text <> "" Then
            QRY = QRY & "AND pacaja = '" & txtcaja.text & "' "
        End If
        
        If OptTipo(1).Value = True Then
          QRY = QRY & "AND pacodemp = '" & Txtcodemp.text & "' "
       End If
       
       If CmbTerminal.ListIndex <> 0 Then
          QRY = QRY & "AND paterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
       End If
       
       '*** PARTE TAG ***
       QRY = QRY & "UNION SELECT DISTINCT tacodemp as cocodemp, 0 AS covalor, 0 AS CantTasas " _
         & "FROM optag " _
         & "WHERE tafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
         & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') "
         
        If txtcaja.text <> "" Then
            QRY = QRY & "AND tanomcaj = '" & txtcaja.text & "' "
        End If
        If OptTipo(1).Value = True Then
            QRY = QRY & "AND tacodemp = '" & Txtcodemp.text & "' "
        End If
         
        QRY = QRY & ") GROUP BY cocodemp "
   End If
   

    
   
   Set Tbl1 = Dbs.Execute(QRY)
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
   End If
   
   'Buscamos si hay multas para este día y realizar todo el proceso
    fqry01 = "SELECT DISTINCT(abnumero) as nummul " _
                & "FROM opabonos " _
               & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy')  " _
                  & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                 & "AND abcerrado = 'S' " _
                 & "AND abestado = 'Activo' "

          If txtcaja.text <> "" Then
           'Ha escogido caja
           fqry01 = fqry01 & " AND abnomcaj = '" & Trim(txtcaja.text) & "' "
          End If

       Set Ftb02Multas = Dbs.Execute(fqry01)
   
   
   
   Do While Not Tbl1.EOF
      DoEvents
      sTotal = 0
      
      If OptTipoTasa(0).Value = True Then
      
               'Sebastian Rondon - Enero 15 de 2024
               'Se cambia la consulta por cambios en los descuentos de saldos de cupos
               'Buscamos todas las tasas primero
'               QRY = "SELECT motasauso " _
'                        & "FROM opmovimientos " _
'                        & "WHERE mocodemp =  " _
'                        & "AND mofecha BETWEEN TO_DATE('" & Fecha & "','MM/DD/YYYY') " _
'                        & "AND TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
'                        & "AND motipo = 'VTTU' " _
'                        & "ORDER BY motasauso "
'
'               Set Tbl2 = Dbs.Execute(QRY)
'
'               If Not Tbl2.EOF Then
'                    Do While Not Tbl2.EOF
'                        fqry01 = "SELECT covalor, covalalc, coparque, cooritra " _
'                                    & "FROM coconduc " _
'                                    & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/DD/YYYY') " _
'                                    & "AND TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
'                                    & "AND coestado = 'Activo' " _
'                                    & "AND comodo = '" & QRY1 & "' " _
'                                    & "AND cocerrado = 'S' " _
'                                    & "AND conumero = " & Tbl2!motasauso & " "
'
'                        If txtcaja.text <> "" Then
'                            fqry01 = fqry01 & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
'                        End If
'
'                        If cmbTerminal.ListIndex <> 0 Then
'                            fqry01 = fqry01 & "AND coterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' "
'                        End If
'
'                        Set Ftb01 = Dbs.Execute(fqry01)
'
'                        If Not Ftb01.EOF Then
'                            If Ftb01!Cooritra = "O" Then
'                                origen = origen + 1
'                                totori = totori + Ftb01!covalor
'                                totalcori = totalcori + Ftb01!covalalc
'                            ElseIf Ftb01!Cooritra = "T" Then
'                                transito = transito + 1
'                                tottra = tottra + Ftb01!covalor
'                                totalctra = totalctra + Ftb01!covalalc
'                            End If
'
'                            parque = parque + Ftb01!coparque
'                        End If
'                    Loop
'               End If
               
               
              '********************************************** Tasas de Uso Origen *******************************************************
              fqry01 = "SELECT count(cooritra) as ori, sum(covalor) as totori, sum(covalalc) as totalcori " _
                       & "FROM coconduc " _
                      & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                        & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND cooritra = 'O' " _
                        & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                        & "AND coestado = 'Activo' " _
                        & "AND comodo = '" & QRY1 & "' " _
                        & "AND cocerrado = 'S' "

              If txtcaja.text <> "" Then
                 QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If

              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If

              Set Ftb01 = Dbs.Execute(fqry01)

              If Not Ftb01.EOF Then
                 If Not Ftb01!ori And Not Ftb01!totori Then
                    origen = Ftb01!ori
                    totori = Ftb01!totori
                    totalcori = Ftb01!totalcori
                 Else
                    origen = 0
                    totori = 0
                    totalcori = 0
                 End If
              End If
              Ftb01.Close
              '****************************************************************************************************************************

              '************************************************* Tasas Uso Transito ******************************************************
              fqry01 = "SELECT count(cooritra) as tra, sum(covalor) as tottra, sum(covalalc) as totalctra " _
                       & "FROM coconduc " _
                      & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                        & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND cooritra = 'T' " _
                        & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                        & "AND coestado = 'Activo' " _
                        & "AND comodo = '" & QRY1 & "' " _
                        & "AND cocerrado = 'S' "

              If txtcaja.text <> "" Then
                 QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If


              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If

              Set Ftb01 = Dbs.Execute(fqry01)

              If Not Ftb01.EOF Then
                 If Not Ftb01!tra And Not Ftb01!tottra Then
                    transito = Ftb01!tra
                    tottra = Ftb01!tottra
                    totalctra = Ftb01!totalctra
                 Else
                    transito = 0
                    tottra = 0
                    totalctra = 0
                 End If
              End If
              Ftb01.Close
              '****************************************************************************************************************************

              '******************************************************** Parqueadero *******************************************************

                fqry01 = "SELECT sum(coparque) as totpar " _
                     & "FROM coconduc " _
                    & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                      & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                      & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon is null) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                      & "AND coestado = 'Activo' " _
                      & "AND comodo = '" & QRY1 & "' " _
                      & "AND cocerrado = 'S' "

                  If txtcaja.text <> "" Then
                      QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
                  End If

                  If CmbTerminal.ListIndex <> 0 Then
                     'ha escogido un terminal
                     fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
                  End If


              Set Ftb01 = Dbs.Execute(fqry01)

              If Not Ftb01.EOF Then
                 If Not Ftb01!totpar Then
                    parque = Ftb01!totpar
                 Else
                    parque = 0
                 End If
              End If
              Ftb01.Close
              
              Elusion = 0
              Evasion = 0
              PerDocumento = 0
              EvasionTransito = 0
              TotalMultas = 0
              vigencia = 0
              Tag = 0
              '********************************************************************************************************************
              
              sTotal = CDbl(totori) + CDbl(totalcori) + CDbl(tottra) + CDbl(totalctra) + CDbl(parque) + CDbl(TotalMultas) + CDbl(vigencia) + CDbl(Tag)
              
              QRY = "UPDATE optempo1 " _
                     & "SET temnumero1 = temnumero1 + " & origen & ", " _
                        & " temnumdob3 = temnumdob3 + " & totori & ", " _
                        & " temnumero3 = temnumero3 + " & totalcori & ", " _
                        & " temnumero2 = temnumero2 + " & transito & ", " _
                        & " temnumdob4 = temnumdob4 + " & tottra & ", " _
                        & " temnumdob5 = temnumdob5 + " & totalctra & ", " _
                        & " temnumdob1 = temnumdob1 + " & Tbl1!covalor & ", " _
                        & " temnumdob2 = temnumdob2 + " & sTotal & ", " _
                        & " temnumero6 = temnumero6 + " & parque & ", " _
                        & " temnumero4 = temnumero4 + " & Tbl1!CantTasas & ", " _
                        & " temnumero5 = temnumero5 + " & Elusion & ", " _
                        & " temnumero11 = temnumero11 + " & Evasion & ", " _
                        & " temnumero12 = temnumero12 + " & PerDocumento & ", " _
                        & " temnumero13 = temnumero13 + " & EvasionTransito & ", " _
                        & " temnumero14 = temnumero14 + " & TotalMultas & ", " _
                        & " temnumero15 = temnumero15 + " & vigencia & ", " _
                        & " temnumero16 = temnumero16 + " & Tag & " " _
                   & "WHERE temtext1 = '" & Tbl1!cocodemp & "' " _
                     & "AND tempc ='" & Pc & "' "
              
              Dbs.Execute QRY, sRecordAfecta
            '            entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Str(origen) + Chr(9) + _
            '                 Str(totori) + Chr(9) + Str(totalcori) + Chr(9) + Str(transito) + Chr(9) + Str(tottra) + Chr(9) + _
            '                 Str(totalctra) + Chr(9) + Str(Tbl1(2)) + Chr(9) + Str(sTotal) + Chr(9) + _
            '                  Str(parque)
              If sRecordAfecta = 0 Then
                 QRY = "INSERT INTO optempo1 " _
                       & "(temtext1, temtext2, temtext3, temnumero1, temnumdob3, " _
                       & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
                       & " temnumdob1, temnumdob2, temnumero6, temnumero4, temnumero5, temnumero11, temnumero12, " _
                       & " temnumero13, temnumero14, temnumero15, temnumero16, tempc) VALUES ('" _
                       & Tbl1!cocodemp & "','" & Tbl1!henombre & "','" & Tbl1!henit & "'," & origen & "," & totori & "," _
                       & totalcori & "," & transito & "," & tottra & "," & totalctra & "," _
                       & Tbl1!covalor & "," & sTotal & "," & parque & ",'" & Tbl1!CantTasas & "'," & Elusion & "," _
                       & Evasion & "," & PerDocumento & "," & EvasionTransito & "," & TotalMultas & "," & vigencia & "," & Tag & ", '" & Pc & "')"
           
                 Dbs.Execute QRY
              
              End If
      Else
        'Sebastián Rondón - Octubre 14 - 2022
      'Se cambia la logica ya que si una empresa se encuentra en la tabla OPHOJEMP es una empresa vinculada directamente a la terminal
      'Si no, se entiende que es una empresa de convenio, la cual solo contará con multas y permanencias
      
      fqry01 = "SELECT hecodigo, henit FROM ophojemp WHERE hecodigo = '" & Tbl1!cocodemp & "' "
      Set FtbEmpresas = Dbs.Execute(fqry01)
        
        If FtbEmpresas.EOF = False Then
                    '********************************************** Tasas de Uso Origen *******************************************************
              fqry01 = "SELECT count(cooritra) as ori, sum(covalor) as totori, sum(covalalc) as totalcori " _
                       & "FROM coconduc " _
                      & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                        & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND cooritra = 'O' " _
                        & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon IS NULL) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                        & "AND coestado = 'Activo' " _
                        & "AND cocerrado = 'S' "
                        '& "AND comodo = '" & Qry1 & "' "
                        '& "AND cocodemp = '" & Tbl1!cocodemp & "' " _

              If txtcaja.text <> "" Then
                 QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If
           
              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
              
              Set Ftb01 = Dbs.Execute(fqry01)
                
              If Not Ftb01.EOF Then
                 If Not Ftb01!ori And Not Ftb01!totori Then
                    origen = Ftb01!ori
                    totori = Ftb01!totori
                    totalcori = Ftb01!totalcori
                 Else
                    origen = 0
                    totori = 0
                    totalcori = 0
                 End If
              End If
              Ftb01.Close
              
              'Buscamos los totales por empresa de convenio y los sumamos
              QRY = "SELECT count(cooritra) as ori, sum(covalor) as totori, sum(covalalc) as totalcori, cocodempcon " _
                            & "FROM coconduc " _
                            & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                            & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                            & "AND cooritra = 'O' " _
                            & "AND cocodemp = '" & Tbl1!cocodemp & "' " _
                            & "AND cocodempcon IS NOT NULL " _
                            & "AND coestado = 'Activo' " _
                            & "AND cocerrado = 'S' "
                            
                            
              If txtcaja.text <> "" Then
                QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If
           
              If CmbTerminal.ListIndex <> 0 Then
                'ha escogido un terminal
                QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
              
                QRY = QRY & " GROUP BY cocodempcon "
              Set Ftb01 = Dbs.Execute(QRY)
              
              If Not Ftb01.EOF Then
                Do While Not Ftb01.EOF
                    'Consultamos si la empresa esta afiliada a la terminal, si no, sumamos los valores recaudados
                    QRY = "SELECT hecodigo FROM ophojemp WHERE hecodigo = '" & Ftb01!cocodempcon & "' "
                    
                    Set Tbl3 = Dbs.Execute(QRY)
                    
                    If Tbl3.EOF Then
                        'Si no se encuentra nada entonces se entiende que no está afiliada a la terminal
                        If Not Ftb01!ori And Not Ftb01!totori Then
                           origen = origen + Ftb01!ori
                           totori = totori + Ftb01!totori
                           totalcori = totalcori + Ftb01!totalcori
                        End If
                    End If
                    
                    Ftb01.MoveNext
                Loop
              End If
              Ftb01.Close
              '****************************************************************************************************************************
              
              '************************************************* Tasas Uso Transito ******************************************************
              fqry01 = "SELECT count(cooritra) as tra, sum(covalor) as tottra, sum(covalalc) as totalctra " _
                       & "FROM coconduc " _
                      & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                        & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND cooritra = 'T' " _
                        & "AND ((cocodemp = '" & Tbl1!cocodemp & "' AND cocodempcon IS NULL) OR cocodempcon = '" & Tbl1!cocodemp & "') " _
                        & "AND coestado = 'Activo' " _
                        & "AND cocerrado = 'S' "
                        '& "AND comodo = '" & Qry1 & "' "
                        '& "AND cocodemp = '" & Tbl1!cocodemp & "' " _

              If txtcaja.text <> "" Then
                 QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If
                        
                        
              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
                    
              Set Ftb01 = Dbs.Execute(fqry01)
              
              If Not Ftb01.EOF Then
                 If Not Ftb01!tra And Not Ftb01!tottra Then
                    transito = Ftb01!tra
                    tottra = Ftb01!tottra
                    totalctra = Ftb01!totalctra
                 Else
                    transito = 0
                    tottra = 0
                    totalctra = 0
                 End If
              End If
              Ftb01.Close
              
              'Buscamos los totales por empresa de convenio y los sumamos
              QRY = "SELECT count(cooritra) as tra, sum(covalor) as tottra, sum(covalalc) as totalctra, cocodempcon " _
                            & "FROM coconduc " _
                            & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                            & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                            & "AND cooritra = 'T' " _
                            & "AND cocodemp = '" & Tbl1!cocodemp & "' " _
                            & "AND cocodempcon IS NOT NULL " _
                            & "AND coestado = 'Activo' " _
                            & "AND cocerrado = 'S' "
                            
              If txtcaja.text <> "" Then
                QRY = QRY & "AND conomcaj = '" & Trim(txtcaja.text) & "' "
              End If
           
              If CmbTerminal.ListIndex <> 0 Then
                'ha escogido un terminal
                QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
              
              QRY = QRY & " GROUP BY cocodempcon "
              
              Set Ftb01 = Dbs.Execute(QRY)
              
              If Not Ftb01.EOF Then
                Do While Not Ftb01.EOF
                    'Consultamos si la empresa esta afiliada a la terminal, si no, sumamos los valores recaudados
                    QRY = "SELECT hecodigo FROM ophojemp WHERE hecodigo = '" & Ftb01!cocodempcon & "' "
                    
                    Set Tbl3 = Dbs.Execute(QRY)
                    
                    If Tbl3.EOF Then
                        'Si no se encuentra nada entonces se entiende que no está afiliada a la terminal
                        If Not Ftb01!tra And Not Ftb01!tottra Then
                           transito = transito + Ftb01!tra
                           tottra = tottra + Ftb01!tottra
                           totalctra = totalctra + Ftb01!totalctra
                        End If
                    End If
                    
                    Ftb01.MoveNext
                Loop
              End If
              Ftb01.Close
              '****************************************************************************************************************************
              
              '******************************************************** Parqueadero *******************************************************
              
              'Sebastián Rondón - 19 de Octubre del 2022
              'Nelson comenta que el reporte de cupos lo utilizan como se estableció desde un inicio
              'Por ello cuándo elijan la opción de consultar por cupos la consulta seguirá igual
              
                fqry01 = "SELECT nvl(sum(pavalor), 0) as totpar " _
                            & "FROM coparqueop " _
                            & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/dd/yyyy') " _
                            & "AND TO_DATE('" & Fecha1 & "', 'MM/dd/yyyy') " _
                            & "AND pacodemp = '" & Tbl1!cocodemp & "' " _
                            & "AND patipo = 'P' " _
                            & "AND paestado = 'C' " _
                            & "AND pacerrado = 'S' "
              
              If txtcaja.text <> "" Then
                 QRY = QRY & "AND pacaja = '" & Trim(txtcaja.text) & "' "
              End If
        
              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND paterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
            
              
              Set Ftb01 = Dbs.Execute(fqry01)
                      
              If Not Ftb01.EOF Then
                 If Not Ftb01!totpar Then
                    parque = Ftb01!totpar
                 Else
                    parque = 0
                 End If
              End If
              Ftb01.Close
              '********************************************************************************************************************
              
              '*********************************************************** TAG ***************************************************
              
                'Sebastián Rondon - 21 Octubre de 2022
                'Se agrega al reporte el nuevo concepto de TAG
                
                fqry01 = "SELECT nvl(sum(tavalor), 0) as tottag " _
                            & "FROM optag " _
                            & "WHERE tafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
                            & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                            & "AND tacodemp = '" & Tbl1!cocodemp & "' " _
                            & "AND taestado = 'A' "
                      
                  If txtcaja.text <> "" Then
                      fqry01 = fqry01 & "AND tanomcaj = '" & Trim(txtcaja.text) & "' "
                  End If
                      
                  'If CmbTerminal.ListIndex <> 0 Then
                     'ha escogido un terminal
                     'fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.Text, 1, 1) & "'"
                  'End If
                  
            
              Set Ftb01 = Dbs.Execute(fqry01)
              
              If Not Ftb01.EOF Then
                 If Not Ftb01!tottag Then
                    Tag = Ftb01!tottag
                 Else
                    Tag = 0
                 End If
              End If
              Ftb01.Close
              
              '*******************************************************************************************************************
              
              '********************************************************** Multas **************************************************
              
        
                 'Si hay registros buscamos concepto por concepto
                 If Ftb02Multas.EOF = False Then
                    For k = 1 To 4
                        Select Case k
                            Case 1
                                sCodigo = L_CiTElusion
                            Case 2
                                sCodigo = L_CiTEvasion
                            Case 3
                                sCodigo = L_CiTPerdida
                            Case 4
                                sCodigo = L_CiTEvasTra
                        End Select
        
                        fqry01 = "SELECT nvl(SUM(abvalor), 0) as valt " _
                                & "FROM opabonos, opmultas " _
                               & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                                 & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                                 & "AND munumero = abnumero  " _
                                 & "AND abcerrado = 'S'  " _
                                 & "AND mucodinf IN (" & sCodigo & ") " _
                                 & "AND mucodemp = '" & Tbl1!cocodemp & "' " _
                                 & "AND abestado = 'Activo' "
        
                        If txtcaja.text <> "" Then
                         'Ha escogido caja
                         fqry01 = fqry01 & " AND abnomcaj = '" & Trim(txtcaja.text) & "' "
                        End If
        
                       Set Ftb01 = Dbs.Execute(fqry01)
        
                       If Not Ftb01.EOF Then
                          If Not IsNull(Ftb01!valt) Then
                             Select Case k
                                Case 1
                                   Elusion = Ftb01!valt
                                Case 2
                                   Evasion = Ftb01!valt
                                Case 3
                                   PerDocumento = Ftb01!valt
                                Case 4
                                   EvasionTransito = Ftb01!valt
                             End Select
                          End If
                       End If
                       Ftb01.Close
                    Next k
        
                    'Ahora calculamos el total de multas
                    fqry01 = "SELECT nvl(SUM(abvalor), 0) as valt " _
                             & "FROM opabonos, opmultas " _
                              & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                              & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                              & "AND munumero = abnumero  " _
                              & "AND abcerrado = 'S'  " _
                              & "AND mucodinf IN (" & L_Todas_Multas & ") " _
                              & "AND mucodemp = '" & Tbl1!cocodemp & "' " _
                              & "AND abestado = 'Activo' "
        
                    If txtcaja.text <> "" Then
                        'Ha escogido caja
                        fqry01 = fqry01 & " AND abnomcaj = '" & Trim(txtcaja.text) & "' "
                    End If
        
                    Set Ftb01 = Dbs.Execute(fqry01)
        
                    If Not Ftb01.EOF Then
                       If Not IsNull(Ftb01!valt) Then
                          TotalMultas = Ftb01!valt
                       End If
                    End If
                    Ftb01.Close
                 Else
                    Elusion = 0
                    Evasion = 0
                    PerDocumento = 0
                    EvasionTransito = 0
                    TotalMultas = 0
                 End If
                 
              
              '********************************************************************************************************************
              
              '************************************************ VIGENCIAS *********************************************************
              
              fqry01 = "SELECT sum(pavalor) as valpar " _
                       & "FROM coparqueop " _
                      & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                      & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                      & "  AND paestado = 'C' " _
                      & "  AND pacerrado = 'S' " _
                      & "  AND patipo = 'V' " _
                      & "  AND pacodemp = '" & Tbl1!cocodemp & "' "
                      
              Set Ftb01 = Dbs.Execute(fqry01)
              If Ftb01.EOF = False Then
                 If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
                    vigencia = Ftb01!valpar
                 Else
                    vigencia = 0
                 End If
              End If
            Ftb01.Close
            
            '**********************************************************************************************************************
              
              sTotal = CDbl(totori) + CDbl(totalcori) + CDbl(tottra) + CDbl(totalctra) + CDbl(parque) + CDbl(TotalMultas) + CDbl(vigencia) + CDbl(Tag)
                
            'Buscamos el nombre de la empresa para guardar
            fqry01 = "SELECT henombre FROM ophojemp WHERE hecodigo = '" & Tbl1!cocodemp & "' "
            Set Ftb01 = Dbs.Execute(fqry01)
            
            If Ftb01.EOF = False Then
                NomEmpresa = Ftb01!henombre
            Else
                fqry01 = "SELECT henombre FROM ophojempcon WHERE hecodigo = '" & Tbl1!cocodemp & "' "
                Set Ftb01 = Dbs.Execute(fqry01)
                
                If Ftb01.EOF = False Then
                    NomEmpresa = Ftb01!henombre
                Else
                    NomEmpresa = "N/A"
                End If
            End If
            
            
              QRY = "UPDATE optempo1 " _
                     & "SET temnumero1 = temnumero1 + " & origen & ", " _
                        & " temnumdob3 = temnumdob3 + " & totori & ", " _
                        & " temnumero3 = temnumero3 + " & totalcori & ", " _
                        & " temnumero2 = temnumero2 + " & transito & ", " _
                        & " temnumdob4 = temnumdob4 + " & tottra & ", " _
                        & " temnumdob5 = temnumdob5 + " & totalctra & ", " _
                        & " temnumdob1 = temnumdob1 + " & Tbl1!covalor & ", " _
                        & " temnumdob2 = temnumdob2 + " & sTotal & ", " _
                        & " temnumero6 = temnumero6 + " & parque & ", " _
                        & " temnumero4 = temnumero4 + " & Tbl1!CantTasas & ", " _
                        & " temnumero5 = temnumero5 + " & Elusion & ", " _
                        & " temnumero11 = temnumero11 + " & Evasion & ", " _
                        & " temnumero12 = temnumero12 + " & PerDocumento & ", " _
                        & " temnumero13 = temnumero13 + " & EvasionTransito & ", " _
                        & " temnumero14 = temnumero14 + " & TotalMultas & ", " _
                        & " temnumero15 = temnumero15 + " & vigencia & ", " _
                        & " temnumero16 = temnumero16 + " & Tag & " " _
                   & "WHERE temtext1 = '" & Tbl1!cocodemp & "' " _
                     & "AND tempc ='" & Pc & "' "
              
              Dbs.Execute QRY, sRecordAfecta
            '            entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Str(origen) + Chr(9) + _
            '                 Str(totori) + Chr(9) + Str(totalcori) + Chr(9) + Str(transito) + Chr(9) + Str(tottra) + Chr(9) + _
            '                 Str(totalctra) + Chr(9) + Str(Tbl1(2)) + Chr(9) + Str(sTotal) + Chr(9) + _
            '                  Str(parque)
              If sRecordAfecta = 0 Then
                 QRY = "INSERT INTO optempo1 " _
                       & "(temtext1, temtext2, temnumero1, temnumdob3, " _
                       & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
                       & " temnumdob1, temnumdob2, temnumero6, temnumero4, temnumero5, temnumero11, temnumero12, " _
                       & " temnumero13, temnumero14, temnumero15, temnumero16, temtext3, tempc) VALUES ('" _
                       & Tbl1!cocodemp & "','" & NomEmpresa & "'," & origen & "," & totori & "," _
                       & totalcori & "," & transito & "," & tottra & "," & totalctra & "," _
                       & Tbl1!covalor & "," & sTotal & "," & parque & ",'" & Tbl1!CantTasas & "'," & Elusion & "," _
                       & Evasion & "," & PerDocumento & "," & EvasionTransito & "," & TotalMultas & "," & vigencia & "," & Tag & ",'" & FtbEmpresas!henit & "','" & Pc & "')"
           
                 Dbs.Execute QRY
              
              End If
            
        Ftb01.Close
        
        Else
            'Cuando es una empresa de convenio, buscamos que vehículos se encontraron pagos
            'Para después cargar estos valos a la empresa contratante
            
            QRY = "SELECT DISTINCT(cocodemp) AS cocodemp, placa AS placa FROM( " _
                        & "SELECT DISTINCT(mucodemp) AS cocodemp, munumpla AS placa FROM opabonos, opmultas " _
                        & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
                        & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                        & "AND munumero = abnumero " _
                        & "AND abcerrado = 'S' AND mucodinf IN (" & L_Todas_Multas & ") " _
                        & "AND mucodemp = '" & Tbl1!cocodemp & "' " _
                        & "AND abestado = 'Activo' " _
                        & "UNION " _
                        & "SELECT DISTINCT(pacodemp) AS cocodemp, paplaca as placa FROM coparqueop " _
                        & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY')" _
                        & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                        & "AND patipo = 'P' AND pacodemp = '" & Tbl1!cocodemp & "' " _
                        & "UNION " _
                        & "SELECT DISTINCT(tacodemp) AS cocodemp, taplaca as placa FROM optag " _
                        & "WHERE tafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
                        & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                        & "AND tacodemp = '" & Tbl1!cocodemp & "' " _
                & ") GROUP BY cocodemp, placa "
                
            Set Tbl2 = Dbs.Execute(QRY)
            
            Do While Not Tbl2.EOF
                '******************************************************** Parqueadero *******************************************************
                        
                fqry01 = "SELECT nvl(sum(pavalor), 0) as totpar " _
                            & "FROM coparqueop " _
                            & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/dd/yyyy') " _
                            & "AND TO_DATE('" & Fecha1 & "', 'MM/dd/yyyy') " _
                            & "AND pacodemp = '" & Tbl2!cocodemp & "' " _
                            & "AND paplaca = '" & Tbl2!Placa & "' " _
                            & "AND patipo = 'P' " _
                            & "AND pacerrado = 'S' " _
                            & "AND paestado = 'C' "
              
              If txtcaja.text <> "" Then
                 QRY = QRY & "AND pacaja = '" & Trim(txtcaja.text) & "' "
              End If
        
              If CmbTerminal.ListIndex <> 0 Then
                 'ha escogido un terminal
                 fqry01 = fqry01 & " AND paterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
              End If
              
              Set Ftb01 = Dbs.Execute(fqry01)
                      
              If Not Ftb01.EOF Then
                 If Not Ftb01!totpar Then
                    parque = Ftb01!totpar
                 Else
                    parque = 0
                 End If
              End If
              Ftb01.Close
              '********************************************************************************************************************
              
              '********************************************************** Multas **************************************************
              
        
                 'Si hay registros buscamos concepto por concepto
                 If Ftb02Multas.EOF = False Then
                    For k = 1 To 4
                        Select Case k
                            Case 1
                                sCodigo = L_CiTElusion
                            Case 2
                                sCodigo = L_CiTEvasion
                            Case 3
                                sCodigo = L_CiTPerdida
                            Case 4
                                sCodigo = L_CiTEvasTra
                        End Select
        
                        fqry01 = "SELECT nvl(SUM(abvalor), 0) as valt " _
                                & "FROM opabonos, opmultas " _
                               & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                                 & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                                 & "AND munumero = abnumero  " _
                                 & "AND abcerrado = 'S'  " _
                                 & "AND mucodinf IN (" & sCodigo & ") " _
                                 & "AND mucodemp = '" & Tbl2!cocodemp & "' " _
                                 & "AND munumpla = '" & Tbl2!Placa & "' " _
                                 & "AND abestado = 'Activo' "
        
                        If txtcaja.text <> "" Then
                         'Ha escogido caja
                         fqry01 = fqry01 & " AND abnomcaj = '" & Trim(txtcaja.text) & "' "
                        End If
        
                       Set Ftb01 = Dbs.Execute(fqry01)
        
                       If Not Ftb01.EOF Then
                          If Not IsNull(Ftb01!valt) Then
                             Select Case k
                                Case 1
                                   Elusion = Ftb01!valt
                                Case 2
                                   Evasion = Ftb01!valt
                                Case 3
                                   PerDocumento = Ftb01!valt
                                Case 4
                                   EvasionTransito = Ftb01!valt
                             End Select
                          End If
                       End If
                       Ftb01.Close
                    Next k
        
                    'Ahora calculamos el total de multas
                    fqry01 = "SELECT nvl(SUM(abvalor), 0) as valt " _
                             & "FROM opabonos, opmultas " _
                              & "WHERE abfecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                              & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                              & "AND munumero = abnumero  " _
                              & "AND abcerrado = 'S'  " _
                              & "AND mucodinf IN (" & L_Todas_Multas & ") " _
                              & "AND mucodemp = '" & Tbl2!cocodemp & "' " _
                              & "AND munumpla = '" & Tbl2!Placa & "' " _
                              & "AND abestado = 'Activo' "
        
                    If txtcaja.text <> "" Then
                        'Ha escogido caja
                        fqry01 = fqry01 & " AND abnomcaj = '" & Trim(txtcaja.text) & "' "
                    End If
        
                    Set Ftb01 = Dbs.Execute(fqry01)
        
                    If Not Ftb01.EOF Then
                       If Not IsNull(Ftb01!valt) Then
                          TotalMultas = Ftb01!valt
                       End If
                    End If
                    Ftb01.Close
                 Else
                    Elusion = 0
                    Evasion = 0
                    PerDocumento = 0
                    EvasionTransito = 0
                    TotalMultas = 0
                 End If
              
              '********************************************************************************************************************
              
              '************************************************ VIGENCIAS *********************************************************
              
              fqry01 = "SELECT sum(pavalor) as valpar " _
                       & "FROM coparqueop " _
                      & "WHERE pafecpag BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') " _
                      & "AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                      & "  AND paestado = 'C' " _
                      & "  AND pacerrado = 'S' " _
                      & "  AND patipo = 'V' " _
                      & "  AND pacodemp = '" & Tbl2!cocodemp & "' " _
                      & "  AND paplaca = '" & Tbl2!Placa & "' "
                      
              Set Ftb01 = Dbs.Execute(fqry01)
              If Ftb01.EOF = False Then
                 If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
                    vigencia = Ftb01!valpar
                 Else
                    vigencia = 0
                 End If
              End If
              Ftb01.Close
            
                '**********************************************************************************************************************
               
               '*********************************************************** TAG ***************************************************
              
                'Sebastián Rondon - 21 Octubre de 2022
                'Se agrega al reporte el nuevo concepto de TAG
                
                fqry01 = "SELECT nvl(sum(tavalor), 0) as tottag " _
                            & "FROM optag " _
                            & "WHERE tafecpag BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
                            & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                            & "AND tacodemp = '" & Tbl2!cocodemp & "' " _
                            & "AND taplaca = '" & Tbl2!Placa & "' " _
                            & "AND taestado = 'A' "
                      
                  If txtcaja.text <> "" Then
                      fqry01 = fqry01 & "AND tanomcaj = '" & Trim(txtcaja.text) & "' "
                  End If
                      
                  'If CmbTerminal.ListIndex <> 0 Then
                     'ha escogido un terminal
                     'fqry01 = fqry01 & " AND coterminal = '" & Mid(CmbTerminal.Text, 1, 1) & "'"
                  'End If
                  
            
              Set Ftb01 = Dbs.Execute(fqry01)
              
              If Not Ftb01.EOF Then
                 If Not Ftb01!tottag Then
                    Tag = Ftb01!tottag
                 Else
                    Tag = 0
                 End If
              End If
              Ftb01.Close
              
              '*******************************************************************************************************************
              
               
                'Una vez teniendo todos los valores, buscamos por la placa del vehículo, la empresa contratante
                'Para sumarle estos valores
                
                QRY = "SELECT hvplaca, hvcodemp, hvempcon FROM ophojveh " _
                        & "WHERE hvplaca = '" & Tbl2!Placa & "' " _
                        & "AND hvcodemp = '" & Tbl2!cocodemp & "' "
                        
                Set Ftb01 = Dbs.Execute(QRY)
                
                If Ftb01.EOF = False Then
                      
                  sTotal = CDbl(parque) + CDbl(TotalMultas) + CDbl(vigencia) + CDbl(Tag)
                    
                 'Buscamos el nombre de la empresa para guardar
                 fqry01 = "SELECT henombre, henit FROM ophojemp WHERE hecodigo = '" & Ftb01!hvempcon & "' "
                 Set Ftb02 = Dbs.Execute(fqry01)
                 
                 If Ftb02.EOF = False Then
                     NomEmpresa = Ftb02!henombre
                     NitEmpresa = Ftb02!henit
                 Else
                     NomEmpresa = "N/A"
                     NitEmpresa = "N/A"
                 End If
                 
                 
                   QRY = "UPDATE optempo1 " _
                          & "SET temnumdob2 = temnumdob2 + " & sTotal & ", " _
                             & " temnumero6 = temnumero6 + " & parque & ", " _
                             & " temnumero5 = temnumero5 + " & Elusion & ", " _
                             & " temnumero11 = temnumero11 + " & Evasion & ", " _
                             & " temnumero12 = temnumero12 + " & PerDocumento & ", " _
                             & " temnumero13 = temnumero13 + " & EvasionTransito & ", " _
                             & " temnumero14 = temnumero14 + " & TotalMultas & ", " _
                             & " temnumero15 = temnumero15 + " & vigencia & ", " _
                             & " temnumero16 = temnumero16 + " & Tag & " " _
                        & "WHERE temtext1 = '" & Ftb01!hvempcon & "' " _
                          & "AND tempc ='" & Pc & "' "
                   
                   Dbs.Execute QRY, sRecordAfecta
                   
                   If sRecordAfecta = 0 Then
                      QRY = "INSERT INTO optempo1 " _
                            & "(temtext1, temtext2, temnumero1, temnumdob3, " _
                            & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
                            & " temnumdob1, temnumdob2, temnumero6, temnumero4, temnumero5, temnumero11, temnumero12, " _
                            & " temnumero13, temnumero14, temnumero15, temnumero16, temtext3, tempc) VALUES ('" _
                            & Ftb01!hvempcon & "','" & NomEmpresa & "', 0, 0, " _
                            & "0, 0, 0, 0, 0," & sTotal & "," & parque & ",'0'," & Elusion & "," _
                            & Evasion & "," & PerDocumento & "," & EvasionTransito & "," & TotalMultas & "," & vigencia & "," & Tag & ", '" & NitEmpresa & "','" & Pc & "')"
                
                      Dbs.Execute QRY
                    End If
                    Ftb01.Close
                    Ftb02.Close
                    
                    
                  Else
                    'Camilo Campos 07/01/2025 se vienen presentando diferencias entre el recaudo por empresas y cierres de turno.
                    'Actualmente se debe a que: un vehiculo tiene convenio, paga parqueadero con el convenio y se realiza cierre de turno
                    'pero da la casualidad de que cambian el convenio del vehiculo y entonces al intentar ligar la información (en el presente if) no encuentra el convenio y no suma lo que deberia
                     
                    'Por petición del ING Nelson, desea sumarselo a la ultima empresa que tuvo convenio
                    
                    'Para ello solo vamos a buscar a que empresa pertenece el vehiculo actualmente y sumarlo a dicha
                    
                      QRY = "SELECT hvplaca, hvcodemp, hvempcon FROM ophojveh " _
                            & "WHERE hvplaca = '" & Tbl2!Placa & "' "
                    
                      Set Ftb01 = Dbs.Execute(QRY)
                    
                     If Ftb01.EOF = False Then
                    
                         sTotal = CDbl(parque) + CDbl(TotalMultas) + CDbl(vigencia) + CDbl(Tag)
                        
                          fqry01 = "SELECT henombre, henit FROM ophojemp WHERE hecodigo = '" & Ftb01!hvempcon & "' "
                          
                          Set Ftb02 = Dbs.Execute(fqry01)
                         
                         If Ftb02.EOF = False Then
                             NomEmpresa = Ftb02!henombre
                             NitEmpresa = Ftb02!henit
                         Else
                             NomEmpresa = "N/A"
                             NitEmpresa = "N/A"
                         End If
                        
                         QRY = "UPDATE optempo1 " _
                                  & "SET temnumdob2 = temnumdob2 + " & sTotal & ", " _
                                     & " temnumero6 = temnumero6 + " & parque & ", " _
                                     & " temnumero5 = temnumero5 + " & Elusion & ", " _
                                     & " temnumero11 = temnumero11 + " & Evasion & ", " _
                                     & " temnumero12 = temnumero12 + " & PerDocumento & ", " _
                                     & " temnumero13 = temnumero13 + " & EvasionTransito & ", " _
                                     & " temnumero14 = temnumero14 + " & TotalMultas & ", " _
                                     & " temnumero15 = temnumero15 + " & vigencia & ", " _
                                     & " temnumero16 = temnumero16 + " & Tag & " " _
                                & "WHERE temtext1 = '" & Ftb01!hvempcon & "' " _
                                  & "AND tempc ='" & Pc & "' "
                           
                           Dbs.Execute QRY, sRecordAfecta
                           
                           If sRecordAfecta = 0 Then
                              QRY = "INSERT INTO optempo1 " _
                                    & "(temtext1, temtext2, temnumero1, temnumdob3, " _
                                    & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
                                    & " temnumdob1, temnumdob2, temnumero6, temnumero4, temnumero5, temnumero11, temnumero12, " _
                                    & " temnumero13, temnumero14, temnumero15, temnumero16, temtext3, tempc) VALUES ('" _
                                    & Ftb01!hvempcon & "','" & NomEmpresa & "', 0, 0, " _
                                    & "0, 0, 0, 0, 0," & sTotal & "," & parque & ",'0'," & Elusion & "," _
                                    & Evasion & "," & PerDocumento & "," & EvasionTransito & "," & TotalMultas & "," & vigencia & "," & Tag & ", '" & NitEmpresa & "','" & Pc & "')"
                        
                              Dbs.Execute QRY
                           End If
                     End If
                    
                End If
                Tbl2.MoveNext
            Loop
        
            
        End If
      
      End If
      
        
        
      ProBar.Value = ProBar.Value + 1
      Tbl1.MoveNext
   Loop
   Ftb02Multas.Close
   Tbl1.Close
   
   Exit Sub
   
Err_XDia:
    MsgBox "Ocurrió un error consultando la información. " & Err.Description, 48, "Advertencia"
    Exit Sub

End Sub

Private Sub Cupos_Sin_Salida(ByVal fFecha As String, ByVal fFecha1 As String)
   QRY = "SELECT cpempresa, HENOMBRE, henit, cpsaldo " _
         & "FROM cocupos LEFT JOIN optempo1 ON cpempresa = temtext1 AND tempc = '" & Pc & "' " _
                     & " INNER JOIN  ophojemp ON cpempresa = hecodigo " _
        & "WHERE temtext1 IS NULL  " _
          & "AND cpempresa <> '099'  " _
     & "ORDER BY cpempresa "
     
   'se deja quemado la empresa 099 porque esta es de pruebas para la terminal
   Set Tbl1 = Dbs.Execute(QRY)
   scero = 0
   
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
   End If
   Do While Not Tbl1.EOF
      
    QRY1 = "SELECT CASE " _
           & "WHEN motipmov = 'DB' AND motipo = 'ANTU' THEN mosaldo + movalor " _
           & "ELSE mosaldo - movalor " _
           & "END AS mosaldo, 'MAX' AS tipo " _
       & "FROM opmovimientos " _
       & "WHERE monumero = (SELECT MAX(monumero) " _
                           & "FROM opmovimientos " _
                           & "WHERE mofecha < TO_DATE('" & fFecha1 & "','MM/dd/yyyy') " _
                           & "AND mocodemp = '" & Tbl1!cpempresa & "' " _
                           & "AND ((motipmov = 'CR' AND motipo = 'VTTU') OR (motipmov = 'DB' AND motipo = 'ANTU')) " _
                           & "AND motipcupo = 'E') " _
       & "UNION ALL " _
       & "SELECT SUM(movalor) AS mosaldo, 'CON' AS tipo " _
                           & "FROM opmovimientos " _
                           & "WHERE mofecha BETWEEN TO_DATE('" & fFecha & "','MM/dd/yyyy') AND TO_DATE('" & fFecha1 & "','MM/dd/yyyy') " _
                           & "AND mocodemp = '" & Tbl1!cpempresa & "' " _
                           & "AND motipmov = 'DB' " _
                           & "AND motipo = 'FACT' " _
                           & "AND motipcupo = 'E' "
      
      'CAU2 4997
      'SE CAMBIAN LA VALIDACION DE BUSQUEDA PARA EXTRAER EL ULTIMO MOVIMIENTO DE LA VENTA DE LA TU
      'WECC 12/30/2021
      '& "WHERE mofecha BETWEEN TO_DATE('" & fFecha & "','MM/dd/yyyy') AND TO_DATE('" & fFecha1 & "','MM/dd/yyyy') " _

      Set Tbl2 = Dbs.Execute(QRY1)
      sSaldoInicial = 0
      sSaldoFinal = 0
      sConsignacion = 0
      
      Do While Not Tbl2.EOF
         If Trim(Tbl2!tipo) = "MAX" Then
            sSaldoFinal = Tbl2!mosaldo
            sSaldoInicial = Tbl2!mosaldo
         ElseIf Trim(Tbl2!tipo) = "CON" Then
            If Not IsNull(Tbl2!mosaldo) Then
               sConsignacion = Tbl2!mosaldo
            
            End If
            
            
         End If
         Tbl2.MoveNext
      Loop
      Tbl2.Close
      If sSaldoInicial = 0 Then
         If sConsignacion = 0 Then
            'sSaldoInicial = Tbl1!cpsaldo
            'sSaldoFinal = Tbl1!cpsaldo
         End If
      Else
         sSaldoFinal = sConsignacion + sSaldoInicial
      End If
      
      
      entry = Tbl1!cpempresa + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + CStr(scero) + Chr(9) + _
                 CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + _
                 CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + _
                 CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + _
                 CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + _
                 CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(scero) + Chr(9) + CStr(sSaldoInicial) + Chr(9) + CStr(sSaldoFinal) + Chr(9) + CStr(sConsignacion)
         
      Grdcontot.AddItem entry
      Tbl1.MoveNext
      ProBar.Value = ProBar.Value + 1
   Loop
   Tbl1.Close
     
   
End Sub


Private Sub BtnBuscar_Click()
   On Error GoTo Errores
   bandera = 0
   
   mskalcoholemia.text = ""
   MskTasas.text = ""
   MskParqueadero.text = ""
   Msktotcon.text = ""
   Lblconven.Caption = ""
   
   If Mskfecini.ClipText = "" Then
      MsgBox "Falta Ingresar La Fecha Inicial", 48, "Advertencia"
      Mskfecini.SetFocus
      Exit Sub
   End If
     
   If Mskfecfin.ClipText = "" Then
      MsgBox "Falta Ingresar La Fecha Final", 48, "Advertencia"
      Mskfecfin.SetFocus
      Exit Sub
   End If
          
   If Val(Left(Mskfecini.text, 2)) < 1 Or Val(Left(Mskfecini.text, 2)) > 12 Then
      MsgBox "No es un MES VALIDO", 48, "Advertencia"
      Mskfecini.SetFocus
      Exit Sub
   End If
  
   If Not IsDate(Mskfecini.text) Then
      MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
      Mskfecini.SetFocus
      Exit Sub
   End If
          
   If Val(Left(Mskfecfin.text, 2)) < 1 Or Val(Left(Mskfecfin.text, 2)) > 12 Then
      MsgBox "No es un MES VALIDO", 48, "Advertencia"
      Mskfecfin.SetFocus
      Exit Sub
   End If
  
   If Not IsDate(Mskfecfin.text) Then
      MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
      Mskfecfin.SetFocus
      Exit Sub
   End If
           
   If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
      MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
      Mskfecini.SetFocus
      Exit Sub
   End If
   
'   If Trim(txtcaja.Text) = "" Then
'      MsgBox "Falta Ingresar El Nombre de La Caja", 48, "Advertencia"
'      txtcaja.SetFocus
'      Exit Sub
'   End If
   
   If OptTipo(1).Value = True Then
      If Txtcodemp.text = "" Then
         MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   
   Fecha = Mskfecini.text
   Fecha1 = Mskfecfin.text
   
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   
   bandera = 0
   
   DIB_GRILLA
        
   '08/31/2001
   'en la tabla cocietur busca los valores minimos y maximos
   'del numero de conduce segun las fechas que se indican en
   'la consulta
   
   If OptTipoTasa(0).Value = True Then
      'cupo
      TITULO = "POR CUPOS"
   ElseIf OptTipoTasa(1).Value = True Then
      'vendido
      TITULO = "POR VENDIDOS"
   End If
   TITULO = "REPORTE TOTALIZADO DE TASAS DE USO " & TITULO
   dato1 = 0
   dato2 = 0
   dato3 = 0
   dato4 = 0
   dato5 = 0
   Dato6 = 0
   Concontotal.Enabled = False
   If OptFecha(1).Value = True Then
      Call X_Turno(Fecha, Fecha1)
   Else
      Call X_Dia(Fecha, Fecha1)
   End If
   
   QRY = "SELECT temtext1 AS cocodemp, temtext2 AS henombre, NVL(temtext3, 'N/A') AS henit, temnumero1 AS origen, temnumdob3 AS totori, " _
             & " temnumero3 AS totalcori, temnumero2 AS transito, temnumdob4 AS tottra, temnumdob5 AS totalctra, " _
             & " temnumdob1 AS sTotalTasas, temnumdob2 AS sTotal, temnumero6 AS parque, temnumero4 AS CantTasas, " _
             & " temnumero5 AS elusion, temnumero11 AS evasion, temnumero12 AS PerDocumento, temnumero13 AS evasionTransito, " _
             & " temnumero14 AS TotalMultas, temnumero15 AS vigencias, temnumero16 AS Tag " _
         & "FROM optempo1 " _
        & "WHERE tempc = '" & Pc & "' ORDER BY sTotal"
         
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
   End If
   
   
   Do While Not Tbl1.EOF
      
        
    If OptTipoTasa(0).Value = True Then
    
      Call Valida_Posicion_Consignacion(Fecha, Fecha1, Tbl1!cocodemp)


        'IMPORTANTE
        'Camilo Campos 06/10/2025 el siguiente codigo de consulta fue modificado de tal forma que:
        'Anteriormente se tomaba el saldo minimo del dia en OPMOVIMIENTOS
        'El problema de esto radica en que si en dado caso el ultimo registro es una anulación, en lugar de sumar el saldo
        'lo seguia restando, como se puede ver en la segunda union agregamos una validación que en caso tal lo sume en lugar de restar.


      QRY1 = "SELECT mosaldo AS mosaldo, 'MIN' AS tipo " _
                & "FROM opmovimientos " _
                & "WHERE monumero = (SELECT MIN(monumero) " _
                & "FROM opmovimientos " _
                        & "WHERE mofecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
                        & "AND motipmov IN ('CR') " _
                        & "AND motipo IN ('VTTU') " _
                        & "AND motipcupo = 'E') " _
         & "UNION ALL " _
                & "SELECT CASE " _
                & "WHEN motipmov = 'DB' AND motipo = 'ANTU' THEN mosaldo + movalor " _
                & "ELSE mosaldo - movalor END AS mosaldo, 'MAX' AS tipo " _
                & "FROM opmovimientos " _
                        & "WHERE monumero = (SELECT MAX(monumero) " _
                        & "FROM opmovimientos " _
                        & "WHERE mofecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                        & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
                        & "AND ((motipmov = 'CR' AND motipo = 'VTTU') OR (motipmov = 'DB' AND motipo = 'ANTU')) " _
                        & "AND motipcupo = 'E') " _

      QRY1 = QRY1 & "UNION ALL " _
           & "SELECT SUM(movalor) AS mosaldo, 'CON' AS tipo " _
                                & "FROM opmovimientos " _
                               & "WHERE mofecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                                 & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
                                 & "AND motipmov = 'DB' " _
                                 & "AND motipo = 'FACT' " _
                                 & "AND motipcupo = 'E' "

      QRY1 = QRY1 & "UNION ALL " _
           & "SELECT SUM(movalor) AS mosaldo, 'ACON' AS tipo " _
                                & "FROM opmovimientos " _
                               & "WHERE mofecha BETWEEN TO_DATE('" & Fecha & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                                 & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
                                 & "AND motipmov = 'CR' " _
                                 & "AND motipo = 'AFAC' " _
                                 & "AND motipcupo = 'E' "

      Set Tbl2 = Dbs.Execute(QRY1)
      sSaldoInicial = 0
      sSaldoFinal = 0
      sConsignacion = 0
      sConsignacionAnulada = 0
      Do While Not Tbl2.EOF
         If Trim(Tbl2!tipo) = "MIN" Then
            sSaldoInicial = Tbl2!mosaldo - LValorMinimoConsignacion
         ElseIf Trim(Tbl2!tipo) = "MAX" Then
            sSaldoFinal = Tbl2!mosaldo + LValorMaximoConsignacion
         ElseIf Trim(Tbl2!tipo) = "CON" Then
            If Not IsNull(Tbl2!mosaldo) Then
               sConsignacion = Tbl2!mosaldo
            End If
         ElseIf Trim(Tbl2!tipo) = "ACON" Then
            If Not IsNull(Tbl2!mosaldo) Then
               sConsignacionAnulada = Tbl2!mosaldo
            End If

         End If
         Tbl2.MoveNext
      Loop
      Tbl2.Close
      sFechaMax = DateAdd("d", -30, Fecha)
'      If sSaldoInicial = 0 Then
'         Qry1 = "SELECT mosaldo - movalor AS mosaldo, 'MIN' AS tipo " _
'                & "FROM opmovimientos " _
'              & " WHERE monumero = (SELECT MAX(monumero) " _
'                                & "FROM opmovimientos " _
'                               & "WHERE mofecha BETWEEN TO_DATE('" & sFechaMax & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
'                                 & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
'                                 & "AND motipmov = 'CR' " _
'                                 & "AND motipo = 'VTTU' " _
'                                 & "AND motipcupo = 'E') "
'
'         Set Tbl2 = Dbs.Execute(Qry1)
'
'         If Not Tbl2.EOF Then
'            sSaldoInicial = Tbl2!mosaldo
'         End If
'         Tbl2.Close
'      End If


'      If sSaldoFinal = 0 Then
'         Qry1 = "SELECT mosaldo - movalor AS mosaldo, 'MAX' AS tipo " _
'                & "FROM opmovimientos " _
'               & "WHERE monumero = (SELECT MAX(monumero) " _
'                                & "FROM opmovimientos " _
'                               & "WHERE mofecha BETWEEN TO_DATE('" & sFechaMax & "','MM/dd/yyyy') AND TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
'                                 & "AND mocodemp = '" & Tbl1!cocodemp & "' " _
'                                 & "AND motipmov = 'CR' " _
'                                 & "AND motipo = 'VTTU' " _
'                                 & "AND motipcupo = 'E') "
'
'         Set Tbl2 = Dbs.Execute(Qry1)
'
'         If Not Tbl2.EOF Then
'            sSaldoFinal = Tbl2!mosaldo
'         End If
'         Tbl2.Close
'
'      End If
    End If
      sTotal = CDbl(Tbl1!totori) + CDbl(Tbl1!tottra) + CDbl(totalcori) + CDbl(totalctra) + CDbl(parque) + CDbl(Tbl1!TotalMultas) + CDbl(Tbl1!Vigencias)
      sValorTotalTasas = CDbl(Tbl1!totori) + CDbl(Tbl1!tottra)
      
      If sSaldoInicial = "" Then
        sSaldoInicial = 0
      End If
      If sSaldoFinal = "" Then
        sSaldoFinal = 0
      End If
      
'      entry = Tbl1!cocodemp + Chr(9) + Tbl1!henombre + Chr(9) + str(Tbl1!origen) + Chr(9) + _
'                 str(Tbl1!totori) + Chr(9) + str(Tbl1!totalcori) + Chr(9) + str(Tbl1!transito) + Chr(9) + str(Tbl1!tottra) + Chr(9) + _
'                 str(Tbl1!totalctra) + Chr(9) + CStr(sValorTotalTasas) + Chr(9) + str(Tbl1!sTotal) + Chr(9) + _
'                  str(Tbl1!parque) + Chr(9) + CStr(sSaldoInicial) + Chr(9) + CStr(sSaldoFinal) + Chr(9) + CStr(sConsignacion - sConsignacionAnulada)
'
          
      'Sebastián Rondón - 03 de Octubre - 2022
      'Se modifica la traida y visualización de datos
      entry = Tbl1!cocodemp + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + str(Tbl1!origen) + Chr(9) + _
                str(Tbl1!totori) + Chr(9) + str(Tbl1!transito) + Chr(9) + str(Tbl1!tottra) + Chr(9) + _
                str(Tbl1!totalcori) + Chr(9) + str(Tbl1!totalctra) + Chr(9) + _
                str(Tbl1!parque) + Chr(9) + str(Tbl1!Elusion) + Chr(9) + _
                str(Tbl1!Evasion) + Chr(9) + str(Tbl1!EvasionTransito) + Chr(9) + str(Tbl1!PerDocumento) + Chr(9) + _
                str(Tbl1!TotalMultas) + Chr(9) + str(Tbl1!Vigencias) + Chr(9) + _
                str(Tbl1!Tag) + Chr(9) + _
                CStr(sValorTotalTasas) + Chr(9) + str(Tbl1!sTotal) + Chr(9) + CStr(sSaldoInicial) + Chr(9) + _
                CStr(sSaldoFinal) + Chr(9) + CStr(sConsignacion - sConsignacionAnulada)
                
      Grdcontot.AddItem entry
      bandera = 1
      
      dato1 = dato1 + CDbl(Tbl1!totalcori)
      dato3 = dato3 + CDbl(Tbl1!totalctra)
      dato4 = dato4 + CDbl(Tbl1!totori) + CDbl(Tbl1!tottra)
      dato5 = dato5 + CDbl(Tbl1!parque)
      Dato6 = Dato6 + CDbl(Tbl1!sTotalTasas) + CDbl(Tbl1!totalcori) + CDbl(Tbl1!totalctra) + CDbl(Tbl1!parque) + CDbl(Tbl1!TotalMultas) + CDbl(Tbl1!Vigencias) + CDbl(Tbl1!Tag)
      
      dato2 = dato2 + CDbl(Tbl1!CantTasas)
      ProBar.Value = ProBar.Value + 1
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   
   If OptTipoTasa(0).Value = True Then
   'WECC MAY 05/2021 CAU No 4160
   'VISUALIZAR EMPRESAS DE CUPOS QUE NO TENGAN SALIDA EN LA FECHA DE BUSQUEDA Y SOLO MOSTRANDO EL SALDO COMO INICIAL Y FINAL
      'Call Cupos_Sin_Salida(Fecha, Fecha)
      'Call Cupos_Sin_Salida(DateAdd("d", -30, Fecha), Fecha) '01/17/2022 CAU 5049 se cambia esta consulta por que se debe mostrar la informacion actual de la fecha solicitada
      Call Cupos_Sin_Salida(Fecha, Fecha)
   End If
   
   
   
   Msktotcon.text = Dato6
   mskalcoholemia.text = dato1 + dato3
   MskTasas.text = dato4
   MskParqueadero.text = dato5
   Lblconven.Caption = dato2
   
'   If vmin <> "" Then
'
'
'
'
'      Qry = "INSERT INTO optempo1 " _
'                      & "(temtext1, temnumdob1, temnumdob2, tempc) " _
'               & "SELECT cocodemp, Sum(covalor), count(conumero), '" & Pc & "' " _
'                 & "FROM coconduc " _
'                & "WHERE conumero BETWEEN " & vmin & " " _
'                  & "AND " & vmax & " " _
'                  & "AND conomcaj = '" & Trim(txtcaja.Text) & "' " _
'                  & "AND coestado = 'Activo' "
'
'      If OptTipo(1).Value = True Then
'         Qry = Qry & "AND cocodemp = '" & Txtcodemp.Text & "' "
'      End If
'      If CmbTerminal.ListIndex <> 0 Then
'         'ha escogido un terminal
'         Qry = Qry & " AND coterminal = '" & Mid(CmbTerminal.Text, 1, 1) & "'"
'      End If
'
'      Qry = Qry & "AND comodo = '" & Qry1 & "' GROUP BY cocodemp "
'
'      Dbs.Execute Qry
'
'      Qry = "SELECT temtext1, henombre, temnumdob2, temnumdob1 " _
'            & "FROM optempo1, ophojemp " _
'           & "WHERE temtext1 = hecodigo " _
'             & "AND tempc = '" & Pc & "'" _
'        & "ORDER BY temnumdob1 DESC "
'
'
'      Set Tbl1 = Dbs.Execute(Qry)
'      If Not Tbl1.EOF Then
'         Tbl1.MoveLast
'         ProBar.Value = ProBar.Min
'         ProBar.Max = Tbl1.RecordCount + 2
'         Tbl1.MoveFirst
'      End If
'
'
'      Concontotal.Enabled = False
'      Do While Tbl1.EOF = False
'         DoEvents
'
'
'         'If IsNull(alcoh) Then
'         '   alcoh = "0"
'         'End If
'
'
'         ProBar.Value = ProBar.Value + 1
'         Tbl1.MoveNext
'      Loop
'
'   End If
   
   
   
   ProBar.Value = ProBar.Max
   Concontotal.Enabled = True
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 48, "Advertencia"
      Exit Sub
   End If
    
   Call Aplicar_Perfil(Me, "Concontotal")
   Exit Sub
Errores:
   Concontotal.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo err_mensaje
   If Grdcontot.rows > 1 Then
      lineas = Grdcontot.rows - 1
      Grdcontot.Row = 1
      Grdcontot.Visible = False
      ProBar.Value = ProBar.Min
      ProBar.Max = Grdcontot.rows + 2
      Concontotal.Enabled = False
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
            
     
      Dbs.BeginTrans
      On Error GoTo Err_Trans
      Do
      
         ReDim dato(1, 21)
                
         For i = 0 To 21
            Grdcontot.col = i
            dato(1, i) = Trim(Grdcontot.text)
         Next i
            
         'insertamos en la tabla
                     
'        QRY = "INSERT INTO optempo1 " _
'                    & "(temtext1, temtext2, temnumero1, temnumdob3, " _
'                    & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
'                    & " temnumdob1, temnumdob2, temnumero6, tempc,temnumdob6,temnumdob7, temnumdob8, " _
'                    & " temnumero5, temnumero11, temnumero12, temnumero13, temnumero14, " _
'                    & " temnumero15 ) VALUES ('" _
'                    & dato(1, 0) & "','" & dato(1, 1) & "'," & dato(1, 2) & "," & dato(1, 3) & "," _
'                    & dato(1, 6) & "," & dato(1, 4) & "," & dato(1, 5) & "," & dato(1, 7) & "," _
'                    & dato(1, 15) & "," & dato(1, 16) & "," & dato(1, 8) & ",'" & Pc & "'," _
'                    & dato(1, 17) & "," & dato(1, 18) & "," & dato(1, 19) & "," & dato(1, 9) & "," _
'                    & dato(1, 10) & "," & dato(1, 12) & "," & dato(1, 11) & "," & dato(1, 13) & "," & dato(1, 14) & ")"

        QRY = "INSERT INTO optempo1 " _
                            & "(temtext1, temtext2, temtext3, temnumero1, temnumdob3, " _
                            & " temnumero3, temnumero2, temnumdob4, temnumdob5, " _
                            & " temnumdob1, temnumdob2, temnumero6, tempc,temnumdob6,temnumdob7, temnumdob8, " _
                            & " temnumero5, temnumero11, temnumero12, temnumero13, temnumero14, " _
                            & " temnumero15, temnumero16 ) VALUES ('" _
                            & dato(1, 0) & "','" & dato(1, 1) & "','" & dato(1, 2) & "'," & dato(1, 3) & "," & dato(1, 4) & "," _
                            & dato(1, 7) & "," & dato(1, 5) & "," & dato(1, 6) & "," & dato(1, 8) & "," _
                            & dato(1, 17) & "," & dato(1, 18) & "," & dato(1, 9) & ",'" & Pc & "'," _
                            & dato(1, 19) & "," & dato(1, 20) & "," & dato(1, 21) & "," & dato(1, 10) & "," _
                            & dato(1, 11) & "," & dato(1, 13) & "," & dato(1, 12) & "," & dato(1, 14) & "," & dato(1, 15) & "," & dato(1, 16) & " )"
         Dbs.Execute QRY
       
         'fin insertar
             
         lineas = lineas - 1
         ProBar.Value = ProBar.Value + 1
         If lineas <> 0 Then
            Grdcontot.Row = Grdcontot.Row + 1
         End If
      Loop Until lineas = 0  'do_while que se mueve x la grilla

      Dbs.CommitTrans
      Grdcontot.Visible = True
      On Error GoTo err_mensaje
       respuesta = MsgBox("Desea escoger la impresora (SI) o impresora predeterminada (NO)", vbYesNo + 32, "Impresoras")
     
      If respuesta = 6 Then
         imprimir.Show 1
         If G_Imprimir = "0" Then
            Exit Sub
         End If
      End If
      If OptTipoTasa(0).Value = True Then
         reporte = RpPath + "\concontocupo.rpt"
      Else
         reporte = RpPath + "\conconto.rpt"
      End If
      rptterminal.ReportFileName = reporte
      rptterminal.Connect = G_Conexion_reportes
      rptterminal.Formulas(0) = "fecini = '" & Mskfecini.FormattedText & "'"
      rptterminal.Formulas(1) = "fecfin = '" & Mskfecfin.FormattedText & "'"
      rptterminal.Formulas(2) = "empresa = '" & G_NomEmpresa & "'"
      rptterminal.Formulas(3) = "titulo = '" & TITULO & "'"
      
      Concontotal.Enabled = False
      rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
      If respuesta = 6 Then
         rptterminal.PrinterName = G_PrinterName
         rptterminal.PrinterDriver = G_PrinterDriver
         rptterminal.PrinterPort = G_PrinterPort
      Else
         rptterminal.PrinterName = ""
         rptterminal.PrinterDriver = ""
         rptterminal.PrinterPort = ""
      End If
      rptterminal.Destination = 0
      rptterminal.WindowState = crptMaximized
      rptterminal.Action = 1
      Concontotal.Enabled = True
     
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      ProBar.Value = ProBar.Max
   End If
  
   Exit Sub
Err_Trans:
  
   Dbs.RollbackTrans
   
  
err_mensaje:
   Grdcontot.Visible = True
   Concontotal.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub

End Sub

Private Sub Btnlimpiar_Click()
  OptTipo(0).Value = True
  OptTipo(0).SetFocus
  Mskfecini.Mask = ""
  Mskfecini.text = ""
  Mskfecini.Mask = "##/##/####"
  Mskfecfin.Mask = ""
  Mskfecfin.text = ""
  Mskfecfin.Mask = "##/##/####"
  Txtcodemp.text = ""
  Lblcodemp.Caption = ""
  Msktotcon.text = ""
  Lblconven.Caption = ""
  mskalcoholemia.text = ""
  MskTasas.text = ""
  MskParqueadero.text = ""
  txtcaja.text = ""
  CmbTerminal.ListIndex = 0
  OptTipoTasa(1).Value = True
  OptFecha(0).Value = True
  DIB_GRILLA
  ProBar.Value = ProBar.Min
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub





Private Sub Form_Activate()
   Call Aplicar_Perfil(Me, "Concontotal")
End Sub

Private Sub Form_Load()
  DIB_GRILLA
   QRY = "SELECT dsdes " _
         & "FROM gesuptip, gedetsuptip " _
        & "WHERE dscodtip = stcodtip " _
          & "AND stdes = 'TIPOS DE TERMINAL' " _
          & "and dsest = 'A' "
    
   Set Tbl1 = Dbs.Execute(QRY)
   
   CmbTerminal.Clear
   CmbTerminal.AddItem "GENERAL"
   Do While Not Tbl1.EOF
      CmbTerminal.AddItem Tbl1(0)
      
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   CmbTerminal.ListIndex = 0
   
   'Sebastián Rondón - 30 Sept - 2022
   'Se empieza a traer conceptos de multas
   'COBRO POR ELUSION DE TUT: 51
   'COBRO POR EVASION DE TUT: 52
   'COBRO POR PERDIDA DE DOCUMENTOS: 53
   'COBRO POR EVASION DE TUT EN TRANSITO: 54
   
   L_CiTElusion = Fn_Parametros("CITELUSION", 1)
   L_CiTEvasion = Fn_Parametros("CITEVASION", 1)
   L_CiTPerdida = Fn_Parametros("CITPERDIDA", 1)
   L_CiTEvasTra = Fn_Parametros("CITEVASTRA", 1)
   L_CiTExcreta = Fn_Parametros("CITEXCRETA", 1)
   
   
   L_Todas_Multas = Trim(L_CiTElusion & ", " & L_CiTEvasion & ", " & L_CiTPerdida & ", " & L_CiTEvasTra)
  
   
   If G_rolusu = "REPFIN" Then
      OptFecha(1).Visible = False
   End If
End Sub

Private Sub Mskfecfin_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskfecfin_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Or (KeyAscii% = 47) Then
      KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecfin_LostFocus()
  If Mskfecfin.ClipText <> "" Then
     If Val(Left(Mskfecfin.text, 2)) < 1 Or Val(Left(Mskfecfin.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecfin.SetFocus
        Exit Sub
     End If
  
     If Mskfecfin.ClipText <> "" Then
        If Not IsDate(Mskfecfin.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Mskfecfin.SetFocus
           Exit Sub
        End If
     End If
  End If
End Sub

Private Sub Mskfecini_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskfecini_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecini_LostFocus()
  If Mskfecini.ClipText <> "" Then
     If Val(Left(Mskfecini.text, 2)) < 1 Or Val(Left(Mskfecini.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecini.SetFocus
        Exit Sub
     End If
  
     If Mskfecini.ClipText <> "" Then
        If Not IsDate(Mskfecini.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Mskfecini.SetFocus
           Exit Sub
        End If
     End If
  End If
End Sub

Private Sub OptTipo_Click(index As Integer)
  If OptTipo(0).Value = True Then
     Txtcodemp.Enabled = False
     Txtcodemp.text = ""
     Lblcodemp.Caption = ""
  End If
  
  If OptTipo(1).Value = True Then
     Txtcodemp.Enabled = True
  End If
  
End Sub

Private Sub Txtcaja_DblClick()
    forma = "12"
    forma1 = "3"
    txtcaja.text = ""
    buscar.Show 1
End Sub

Private Sub txtcaja_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
    
   If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
      KeyAscii% = 0
   End If

End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "16"
  Txtcodemp.text = ""
  Lblcodemp.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
  If Txtcodemp.text <> "" Then
     fqry01 = "SELECT * " _
              & "FROM ophojemp " _
             & "WHERE hecodigo = '" & Txtcodemp.text & "' "
              
     Set Ftb01 = Dbs.Execute(fqry01)
      
     If Not Ftb01.EOF Then
        Lblcodemp.Caption = Ftb01!henombre
     Else
        MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
        forma = "16"
        forma1 = "16"
        buscar.Show 1
        Txtcodemp.SetFocus
     End If
  End If
End Sub

Private Function DIB_GRILLA()
   Grdcontot.rows = 1
   Grdcontot.Cols = 15
   Grdcontot.Clear
   
   's$ = "Cód. Empresa  |Nombre   |Cant.Origen |Tot. Origen |Cant.Transito |Tot. Transito |Vr.Alc. Origen |Vr.Alc. Transito |Parqueadero |Elusion |Evasion |Evasion Transito |Per. Documento |Total Multas |Vigencias |Tasas de Uso Vendidos  |Total General |Saldo Inicial | Saldo Final| Consignacion|"
   s$ = "Cód. Empresa  |Nombre   |NIT |Cant.Origen |Tot. Origen |Cant.Transito |Tot. Transito |Vr.Alc. Origen |Vr.Alc. Transito |Parqueadero |Elusion |Evasion |Evasion Transito |Per. Documento |Total Multas |Vigencias |TAG |Tasas de Uso Vendidos  |Total General |Saldo Inicial | Saldo Final| Consignacion "
   Grdcontot.FormatString = s$
   Grdcontot.Row = 0
   Grdcontot.col = 0
   Grdcontot.ColWidth(0) = 1400
   Grdcontot.ColWidth(1) = 2200
   Grdcontot.ColWidth(2) = 1400
   Grdcontot.ColWidth(3) = 1800
   Grdcontot.ColWidth(4) = 1800
   Grdcontot.ColWidth(5) = 1600
   Grdcontot.ColWidth(6) = 1600
   Grdcontot.ColWidth(7) = 1600
   Grdcontot.ColWidth(8) = 1600
   Grdcontot.ColWidth(9) = 1600
   Grdcontot.ColWidth(10) = 1600
   Grdcontot.ColWidth(11) = 1600
   Grdcontot.ColWidth(12) = 1600
   Grdcontot.ColWidth(13) = 1600
   Grdcontot.ColWidth(14) = 1600
   Grdcontot.ColWidth(15) = 1600
   Grdcontot.ColWidth(16) = 1600
   Grdcontot.ColWidth(17) = 1600
   Grdcontot.ColWidth(18) = 1600
   Grdcontot.ColWidth(19) = 1600
   Grdcontot.ColWidth(20) = 1600
   Grdcontot.ColWidth(21) = 1600
   
   For i = 0 To Grdcontot.Cols - 1
       Grdcontot.col = i
       Grdcontot.CellFontBold = 1
   Next i
End Function
