VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Conparqu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Permanencias y Vigencias"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnExportar 
      Caption         =   "&Exportar"
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
      Left            =   8040
      Picture         =   "Conparqu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame4 
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
      Height          =   1215
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   7695
      Begin MSMask.MaskEdBox Msktotal 
         Height          =   300
         Left            =   5400
         TabIndex        =   40
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MskCanceladas 
         Height          =   300
         Left            =   5400
         TabIndex        =   41
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MskActivas 
         Height          =   300
         Left            =   1320
         TabIndex        =   42
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MskDescuento 
         Height          =   300
         Left            =   1320
         TabIndex        =   46
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
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
      Begin VB.Label Label2 
         Caption         =   "Descuentos:"
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
         TabIndex        =   47
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Canceladas:"
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
         Left            =   3720
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Total General:"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Activas:"
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
         TabIndex        =   43
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   37
      Top             =   3960
      Width           =   7695
      Begin Crystal.CrystalReport rptterminal 
         Left            =   1920
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid GrdParqueo 
         Height          =   2055
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
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
      Left            =   8040
      Picture         =   "Conparqu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Buscar informacion en la base de datos"
      Top             =   3840
      Width           =   855
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
      Left            =   8040
      Picture         =   "Conparqu.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7200
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
      Left            =   8040
      Picture         =   "Conparqu.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Limpia la forma actual"
      Top             =   6360
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
      Left            =   8040
      Picture         =   "Conparqu.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   5520
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterio de Búsqueda:"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   2295
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Tipo"
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
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Estado"
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
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Cód. Empresa"
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
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Placa"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Turno"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Chkcriterio 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valores de Búsqueda:"
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
      Height          =   2415
      Left            =   2520
      TabIndex        =   26
      Top             =   1320
      Width           =   5295
      Begin VB.ComboBox CmbTipo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Conparqu.frx":1412
         Left            =   1560
         List            =   "Conparqu.frx":141C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Estado del registro"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.ComboBox CmbCaja 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Conparqu.frx":1437
         Left            =   1560
         List            =   "Conparqu.frx":1441
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Nmobre de la caja"
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox Cmbestado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Conparqu.frx":1458
         Left            =   1560
         List            =   "Conparqu.frx":1465
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Estado del registro"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   14
         ToolTipText     =   "Código de la empresa (Precione doble click para ayuda)"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   13
         ToolTipText     =   "Número de la placa del vehículo (Precione doble click para ayuda)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtTurno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   16
         ToolTipText     =   "Número del turno"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo:"
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
         TabIndex        =   48
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Turno:"
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
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Cód. Empresa:"
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
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Placa:"
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
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Caja:"
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
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Estado:"
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
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
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
         Height          =   285
         Left            =   2400
         TabIndex        =   29
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label LblPlaca 
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
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label LblTurno 
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
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Top             =   1320
         Width           =   2775
      End
   End
   Begin VB.Frame Frmfechas 
      Caption         =   "Fechas:"
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
      TabIndex        =   23
      Top             =   120
      Width           =   7695
      Begin VB.OptionButton OptFecha 
         Caption         =   "Activa"
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
         Index           =   5
         Left            =   6600
         TabIndex        =   51
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Turno"
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
         Index           =   4
         Left            =   6600
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkValor 
         Caption         =   "Valor Mayor 0"
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
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Desc."
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
         Index           =   3
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Pago"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Salida"
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
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Ingreso"
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
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Fecha en la que empiece a buscar"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         ToolTipText     =   "Fecha hasta donde quiere que busque"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
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
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   1800
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   36
      Top             =   7800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   7920
      X2              =   7920
      Y1              =   7920
      Y2              =   0
   End
End
Attribute VB_Name = "Conparqu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DIB_GRILLA()
   GrdParqueo.rows = 1
   GrdParqueo.Cols = 15
   GrdParqueo.Clear
     
     
   If CmbTipo.text = "Vigencia" Then
      s$ = "Número|Placa  |Empresa|Fecha Limite |Fecha Salida |Fecha Pago |Tiempo Excedido|" _
           & "Turno|Caja|Tasa|Vr Descuento| Valor a Pagar|Estado|Observaciones|Tipo|Usuario"
   Else
      s$ = "Número|Placa  |Empresa|Fecha Ingreso |Fecha Salida |Fecha Pago |Tiempo Permanencia|" _
        & "Turno|Caja|Tasa|Vr Descuento| Valor a Pagar|Estado|Observaciones|Tipo|Usuario"
   End If
   GrdParqueo.FormatString = s$
   GrdParqueo.Row = 0
   GrdParqueo.Col = 0
   GrdParqueo.ColWidth(0) = 1000
   GrdParqueo.ColWidth(1) = 1000
   GrdParqueo.ColWidth(2) = 1800
   GrdParqueo.ColWidth(3) = 1600
   GrdParqueo.ColWidth(4) = 1600
   GrdParqueo.ColWidth(5) = 1600
   GrdParqueo.ColWidth(6) = 1900
   GrdParqueo.ColWidth(7) = 1000
   GrdParqueo.ColWidth(8) = 1000
   GrdParqueo.ColWidth(9) = 1000
   GrdParqueo.ColWidth(10) = 1300
   GrdParqueo.ColWidth(11) = 1300
   GrdParqueo.ColWidth(12) = 1000
   GrdParqueo.ColWidth(13) = 2000
   GrdParqueo.ColWidth(14) = 1000
   GrdParqueo.ColWidth(15) = 1000
     
   For i = 0 To GrdParqueo.Cols - 1
       GrdParqueo.Col = i
       GrdParqueo.CellFontBold = 1
   Next i
End Sub

Private Function ImpReporte()
     
   reporte = RpPath + "\conparqu.rpt"
   rptterminal.ReportFileName = reporte
   rptterminal.Connect = G_Conexion_reportes
   rptterminal.Formulas(0) = "empresa = '" & G_NomEmpresa & "'"
   rptterminal.Formulas(1) = "inicio = '" & Mskfecini.FormattedText & "'"
   rptterminal.Formulas(2) = "final = '" & Mskfecfin.FormattedText & "'"
   rptterminal.Formulas(3) = "total = '" & Msktotal.FormattedText & "'"
   rptterminal.Formulas(4) = "totactivas = '" & MskActivas.FormattedText & "'"
   rptterminal.Formulas(5) = "totcanceladas = '" & MskCanceladas.FormattedText & "'"
   rptterminal.Formulas(6) = "totdescuento = '" & MskDescuento.FormattedText & "'"
   Conparqu.Enabled = False
   rptterminal.WindowState = crptMaximized
   rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
   rptterminal.Destination = 0
   rptterminal.Action = 1
   Conparqu.Enabled = True
   
End Function

Private Sub Cargar_Combo_Caja()
   On Error GoTo Errores
   
   QRY = "SELECT canomcaj " _
         & "FROM cocaja "
   
   Set Tbl1 = Dbs.Execute(QRY)
   CmbCaja.Clear
   Do While Not Tbl1.EOF
      CmbCaja.AddItem Tbl1!canomcaj
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub BtnBuscar_Click()
   On Error GoTo Errores
   Dim L_DirCristal As String
   
   MskActivas.text = "0"
   MskCanceladas.text = "0"
   Msktotal.text = "0"
   MskDescuento.text = "0"
   L_DirCristal = "N"
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
        
   If Chkcriterio(0).Value = 1 Then
      If TxtPlaca.text = "" Then
         MsgBox "Falta ingresar el código de la placa", 48, "Advertencia"
         TxtPlaca.SetFocus
         Exit Sub
      End If
   End If
        
   If Chkcriterio(1).Value = 1 Then
      If Txtcodemp.text = "" Then
         MsgBox "Falta ingresar el código de la empresa", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(2).Value = 1 Then
      If CmbCaja.ListIndex = -1 Then
         MsgBox "Falta seleccionar la caja", 48, "Advertencia"
         CmbCaja.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(3).Value = 1 Then
      If TxtTurno.text = "" Then
         MsgBox "Falta ingresar el número de turno", 48, "Advertencia"
         TxtTurno.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(4).Value = 1 Then
      If Cmbestado.ListIndex = -1 Then
         MsgBox "Falta seleccionar el estado", 48, "Advertencia"
         Cmbestado.SetFocus
         Exit Sub
      End If
   End If
        
   If Chkcriterio(5).Value = 1 Then
      If CmbTipo.ListIndex = -1 Then
         MsgBox "Falta seleccionar el tipo", 48, "Advertencia"
         CmbTipo.SetFocus
         Exit Sub
      End If
   End If
   
   DIB_GRILLA
   
   QRY = "SELECT paplaca, henombre, pafecini, pafecsal, pafecpag, " _
              & "patiempo, paturno, pacaja, NVL(patasapag, 0) as patasa, NVL(padescuento, 0) as padescuento, pavalor, paestado, paobservacion, panumero, patipo, pafundes " _
         & "FROM coparqueop INNER JOIN ophojemp " _
           & "ON coparqueop.pacodemp = ophojemp.hecodigo " _
        
   If OptFecha(0).Value = True Then
      QRY = QRY & "WHERE pafecini "
   ElseIf OptFecha(1).Value = True Then
      QRY = QRY & "WHERE pafecsal "
   ElseIf OptFecha(2).Value = True Then
      QRY = QRY & "WHERE pafecpag "
   ElseIf OptFecha(3).Value = True Then
      QRY = QRY & "WHERE pafecdes "
   ElseIf OptFecha(4).Value = True Then
      QRY = QRY & "WHERE pafechadia "
   ElseIf OptFecha(5).Value = True Then
      'Buscar las activas
      QRY = QRY & "WHERE pafecpag IS NULL AND pafecdes IS NULL AND pafecini "
   End If
   
    QRY = QRY & " BETWEEN TO_DATE('" & Mskfecini.text & " 00:00', 'MM/dd/yyyy HH24:MI') " _
            & "AND TO_DATE('" & Mskfecfin.text & " 23:59', 'MM/dd/yyyy HH24:MI') "

   
   If Trim(TxtPlaca.text) <> "" Then
      QRY = QRY & " AND paplaca = '" & TxtPlaca.text & "' "
   End If
   
   If Trim(Txtcodemp.text) <> "" Then
      QRY = QRY & " AND pacodemp = '" & Txtcodemp.text & "' "
   End If
   If Trim(CmbCaja.text) <> "" Then
      QRY = QRY & " AND pacaja = '" & CmbCaja.text & "' "
   End If
   If Trim(TxtTurno.text) <> "" Then
      QRY = QRY & " AND paturno = " & TxtTurno.text & " "
   End If
   If Trim(Cmbestado.text) <> "" Then
      QRY = QRY & " AND paestado = '" & Mid(Cmbestado.text, 1, 1) & "' "
   End If
   If Trim(CmbTipo.text) <> "" Then
      QRY = QRY & " AND patipo = '" & Mid(CmbTipo.text, 1, 1) & "' "
   End If
   If chkValor.Value = 1 Then
      QRY = QRY & " AND pavalor > 0 "
   End If
   
   QRY = QRY & " ORDER BY henombre, panumero "
   Set Tbl1 = Dbs.Execute(QRY)
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      NumRegistros = Tbl1.RecordCount
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
      If CDbl(G_CantRegist) < CDbl(NumRegistros) Then
         MsgBox "El número de registros encontrados excede el limite para mostrar " & _
         Chr(13) & "             en la grilla, se generará directo a pantalla", 48, "Advertencia"
         L_DirCristal = "S"
         
      End If
   End If
   bandera = 0
   sValActiva = 0
   sValCancelada = 0
   sValDescuento = 0
   ProBar.Value = ProBar.Min
   Conparqu.Enabled = False
   Do While Not Tbl1.EOF
      DoEvents
      sFecPag = ""
      If Not IsNull(Tbl1!pafecpag) Then
         sFecPag = Format(Tbl1!pafecpag, "MMM/dd/yyyy")
      Else
        sFecPag = "N/A"
      End If
      sturno = ""
      If Not IsNull(Tbl1!paturno) Then
         sturno = Tbl1!paturno
      Else
        'sturno = "N/A"
        sturno = 0
      End If
      sCaja = ""
      If Not IsNull(Tbl1!pacaja) Then
         sCaja = Tbl1!pacaja
      End If
      
      If Tbl1!paestado = "A" Then
         sEstado = "Activa"
         sValActiva = sValActiva + Tbl1!pavalor
      ElseIf Tbl1!paestado = "C" Then
         sEstado = "Cancelada"
         sValCancelada = sValCancelada + Tbl1!pavalor
      End If
      
      sObservacion = ""
      If Not IsNull(Tbl1!paobservacion) Then
         sObservacion = Tbl1!paobservacion
      End If
      sObservacion = Replace(sObservacion, "'", "")
      'tiempo de permanencia en el terminal
      horest = DateDiff("s", Tbl1!pafecini, Tbl1!pafecsal)
      
      horest1 = horest / 86400
      If horest >= 86400 Then
         horest = horest / 86400
         perma1 = Fix(horest)
         perma = horest - perma1
         sTiempo = perma1 & " día " & Format(perma, "hh:mm") 'para revisar
      Else
         horest = horest / 86400
         sTiempo = Format(horest, "hh:mm")
      End If
      
      sDescuento = 0
      If Not IsNull(Tbl1!padescuento) Then
         sDescuento = Tbl1!padescuento
      End If
      
      Dim funcionario
      If Not IsNull(Tbl1!PAFUNDES) Then
        funcionario = Tbl1!PAFUNDES
      Else
        funcionario = ""
      End If
          
          
      If L_DirCristal = "N" Then
         entry = CStr(Tbl1!panumero) + Chr(9) + Tbl1!paplaca + Chr(9) + Tbl1!henombre + Chr(9) + Format(Tbl1!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
                 Format(Tbl1!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + sFecPag + Chr(9) + sTiempo + Chr(9) + CStr(sturno) + Chr(9) + _
                 sCaja + Chr(9) + CStr(Tbl1!patasa) + Chr(9) + CStr(sDescuento) + Chr(9) + CStr(Tbl1!pavalor) + Chr(9) + _
                 sEstado + Chr(9) + sObservacion + Chr(9) + Tbl1!patipo + Chr(9) + funcionario
         GrdParqueo.AddItem entry
      End If
      sPlaca = Tbl1!paplaca
      sEmpresa = Tbl1!henombre
      sFecIni = Format(Tbl1!pafecini, "MMM/dd/yyyy hh:mm")
      sFecSal = Format(Tbl1!pafecsal, "MMM/dd/yyyy hh:mm")
      'sFecPag = sFecPag
      'sTiempo = sTiempo
      'sturno = sturno
      'sCaja = sCaja
      sTasa = Tbl1!patasa
      sDescuento = sDescuento
      sValor = Tbl1!pavalor
      'sEstado = sEstado
      'sObservacion = sObservacion
      
      QRY = "INSERT INTO optempo1 " _
         & "(temtext1,temtext2, temtext3, temtext4, " _
         & "temtext5, temtext6, temnumero1, temtext7, " _
         & "temnumdob1, temtext8, tempc, temnumdob2, temmemo1, temnumdob3, temnumero2, temtext9, temtext10) VALUES ('" _
         & sPlaca & "' ,'" & sEmpresa & "','" & sFecIni & "','" & sFecSal & "','" _
         & sFecPag & "','" & sTiempo & "' ," & sturno & ",'" & sCaja & "'," & sValor _
         & ",'" & sEstado & "','" & Pc & "'," & sTasa & ",'" & sObservacion & "'," & sDescuento & "," & Tbl1!panumero & ",'" & Tbl1!patipo & "', '" & funcionario & "')"
      
      Dbs.Execute QRY

      sValDescuento = sValDescuento + sDescuento
      
      bandera = 1
      ProBar.Value = ProBar.Value + 1
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   MskActivas.text = sValActiva
   MskCanceladas.text = sValCancelada
   MskDescuento.text = sValDescuento
   Msktotal.text = sValActiva + sValCancelada + sValDescuento
   Conparqu.Enabled = True
   If bandera = 0 Then
      MsgBox "No se encontro ningún registro activo", 64, "Busqueda"
      Exit Sub
   End If
   If L_DirCristal = "S" Then
      ImpReporte
   End If
   ProBar.Value = ProBar.Max
   Call Aplicar_Perfil(Me, "Conparqu")
   Exit Sub
Errores:
   Conparqu.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub BtnExportar_Click()
    On Error GoTo Errores
   Dim L_DirCristal As String
   
   MskActivas.text = "0"
   MskCanceladas.text = "0"
   Msktotal.text = "0"
   MskDescuento.text = "0"
   L_DirCristal = "N"
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
        
   If Chkcriterio(0).Value = 1 Then
      If TxtPlaca.text = "" Then
         MsgBox "Falta ingresar el código de la placa", 48, "Advertencia"
         TxtPlaca.SetFocus
         Exit Sub
      End If
   End If
        
   If Chkcriterio(1).Value = 1 Then
      If Txtcodemp.text = "" Then
         MsgBox "Falta ingresar el código de la empresa", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(2).Value = 1 Then
      If CmbCaja.ListIndex = -1 Then
         MsgBox "Falta seleccionar la caja", 48, "Advertencia"
         CmbCaja.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(3).Value = 1 Then
      If TxtTurno.text = "" Then
         MsgBox "Falta ingresar el número de turno", 48, "Advertencia"
         TxtTurno.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(4).Value = 1 Then
      If Cmbestado.ListIndex = -1 Then
         MsgBox "Falta seleccionar el estado", 48, "Advertencia"
         Cmbestado.SetFocus
         Exit Sub
      End If
   End If
        
   If Chkcriterio(5).Value = 1 Then
      If CmbTipo.ListIndex = -1 Then
         MsgBox "Falta seleccionar el tipo", 48, "Advertencia"
         CmbTipo.SetFocus
         Exit Sub
      End If
   End If
   
   'Limpiamos la Grilla
   DIB_GRILLA
   
   QRY = "SELECT paplaca, henombre, pafecini, pafecsal, pafecpag, " _
              & "patiempo, paturno, pacaja, NVL(patasapag, 0) as patasa, NVL(padescuento, 0) as padescuento, pavalor, paestado, paobservacion, panumero, patipo, pafundes " _
         & "FROM coparqueop INNER JOIN ophojemp " _
           & "ON coparqueop.pacodemp = ophojemp.hecodigo " _
        
   If OptFecha(0).Value = True Then
      QRY = QRY & "WHERE pafecini "
   ElseIf OptFecha(1).Value = True Then
      QRY = QRY & "WHERE pafecsal "
   ElseIf OptFecha(2).Value = True Then
      QRY = QRY & "WHERE pafecpag "
   ElseIf OptFecha(3).Value = True Then
      QRY = QRY & "WHERE pafecdes "
   ElseIf OptFecha(4).Value = True Then
      QRY = QRY & "WHERE pafechadia "
   ElseIf OptFecha(5).Value = True Then
      'Buscar las activas
      QRY = QRY & "WHERE pafecpag IS NULL AND pafecdes IS NULL AND pafecini "
   End If
   
   QRY = QRY & " BETWEEN TO_DATE('" & Mskfecini.text & " 00:00', 'MM/dd/yyyy HH24:MI') " _
         & "AND TO_DATE('" & Mskfecfin.text & " 23:59', 'MM/dd/yyyy HH24:MI') "
   
   If Trim(TxtPlaca.text) <> "" Then
      QRY = QRY & " AND paplaca = '" & TxtPlaca.text & "' "
   End If
   
   If Trim(Txtcodemp.text) <> "" Then
      QRY = QRY & " AND pacodemp = '" & Txtcodemp.text & "' "
   End If
   If Trim(CmbCaja.text) <> "" Then
      QRY = QRY & " AND pacaja = '" & CmbCaja.text & "' "
   End If
   If Trim(TxtTurno.text) <> "" Then
      QRY = QRY & " AND paturno = " & TxtTurno.text & " "
   End If
   If Trim(Cmbestado.text) <> "" Then
      QRY = QRY & " AND paestado = '" & Mid(Cmbestado.text, 1, 1) & "' "
   End If
   If Trim(CmbTipo.text) <> "" Then
      QRY = QRY & " AND patipo = '" & Mid(CmbTipo.text, 1, 1) & "' "
   End If
   If chkValor.Value = 1 Then
      QRY = QRY & " AND pavalor > 0 "
   End If
   
   QRY = QRY & " ORDER BY henombre, panumero "
   
   ExportarExcel (QRY)
   
Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
  
End Sub

Private Sub ExportarExcel(QRY As String)
    On Error GoTo Errores
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "NUMERO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "PLACA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "EMPRESA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "FEC. ENTRADA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "FEC. SALIDA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "FEC. PAGO"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "TIEMPO PERMANENCIA"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "TURNO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "CAJA"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "TASA"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "VR DESCUENTO"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "VALOR A PAGAR"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "ESTADO"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "OBSERVACIONES"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "TIPO"
    ApExcel.Cells(fila, 15).Font.Size = 10
    ApExcel.Cells(fila, 16).formula = "USUARIO"
    ApExcel.Cells(fila, 16).Font.Size = 10
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el Progress Bar
    If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
    End If
    
    If Not Tbl1.EOF Then
        While Not Tbl1.EOF
            sFecPag = ""
            If Not IsNull(Tbl1!pafecpag) Then
               sFecPag = Format(Tbl1!pafecpag, "MMM/dd/yyyy")
            Else
              sFecPag = "N/A"
            End If
            sturno = ""
            If Not IsNull(Tbl1!paturno) Then
               sturno = Tbl1!paturno
            Else
              sturno = "0"
            End If
            sCaja = ""
            If Not IsNull(Tbl1!pacaja) Then
               sCaja = Tbl1!pacaja
            End If
            
            If Tbl1!paestado = "A" Then
               sEstado = "Activa"
               sValActiva = sValActiva + Tbl1!pavalor
            ElseIf Tbl1!paestado = "C" Then
               sEstado = "Cancelada"
               sValCancelada = sValCancelada + Tbl1!pavalor
            End If
            
            sObservacion = ""
            If Not IsNull(Tbl1!paobservacion) Then
               sObservacion = Tbl1!paobservacion
            End If
            sObservacion = Replace(sObservacion, "'", "")
            
            'tiempo de permanencia en el terminal
            horest = DateDiff("s", Tbl1!pafecini, Tbl1!pafecsal)
            
            horest1 = horest / 86400
            If horest >= 86400 Then
               horest = horest / 86400
               perma1 = Fix(horest)
               perma = horest - perma1
               sTiempo = perma1 & " día " & Format(perma, "hh:mm") 'para revisar
            Else
               horest = horest / 86400
               sTiempo = Format(horest, "hh:mm")
            End If
            
            sDescuento = 0
            If Not IsNull(Tbl1!padescuento) Then
               sDescuento = Tbl1!padescuento
            End If
            
            Dim funcionario
            If Not IsNull(Tbl1!PAFUNDES) Then
              funcionario = Tbl1!PAFUNDES
            Else
              funcionario = ""
            End If
    
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!panumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!paplaca & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!pafecini & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!pafecsal & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & sFecPag & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & sTiempo & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & sturno & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & sCaja & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!patasa & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & sDescuento & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!pavalor & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & sEstado & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & sObservacion & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & Tbl1!patipo & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            ApExcel.Cells(fila, 16).formula = "" & funcionario & ""
            ApExcel.Cells(fila, 16).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
            
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA" & Month(Mskfecini.text) & Day(Mskfecini.text) & Year(Mskfecini.text)
        ApExcel.Visible = True
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub
Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo Errores
   
'   ProBar.Value = ProBar.Min
'   ProBar.Max = GrdParqueo.Rows + 2
'
'   If GrdParqueo.Rows > 1 Then
'      Dbs.Execute "DELETE FROM optempo1 WHERE  tempc = '" & Pc & "' "
'
'      Conparqu.Enabled = False
'      Dbs.BeginTrans
'      On Error GoTo Err_Trans
'      For i = 1 To GrdParqueo.Rows - 1
'
'         sPlaca = GrdParqueo.TextMatrix(i, 0)
'         sEmpresa = GrdParqueo.TextMatrix(i, 1)
'         sFecIni = GrdParqueo.TextMatrix(i, 2)
'         sFecSal = GrdParqueo.TextMatrix(i, 3)
'         sFecPag = GrdParqueo.TextMatrix(i, 4)
'         sTiempo = GrdParqueo.TextMatrix(i, 5)
'         sturno = GrdParqueo.TextMatrix(i, 6)
'         sCaja = GrdParqueo.TextMatrix(i, 7)
'         sTasa = GrdParqueo.TextMatrix(i, 8)
'         sDescuento = GrdParqueo.TextMatrix(i, 9)
'         sValor = GrdParqueo.TextMatrix(i, 10)
'         sEstado = GrdParqueo.TextMatrix(i, 11)
'         sObservacion = GrdParqueo.TextMatrix(i, 12)
'
'         ' insertamos en la tabla
'
'         qry = "INSERT INTO optempo1 " _
'         & "(temtext1,temtext2, temtext3, temtext4, " _
'         & "temtext5, temtext6, temnumero1, temtext7, " _
'         & "temnumdob1, temtext8, tempc, temnumdob2, temmemo1, temnumdob3) VALUES ('" _
'         & sPlaca & "' ,'" & sEmpresa & "','" & sFecIni & "','" & sFecSal & "','" _
'         & sFecPag & "','" & sTiempo & "' ," & sturno & ",'" & sCaja & "'," & sValor _
'         & ",'" & sEstado & "','" & Pc & "'," & sTasa & ",'" & sObservacion & "'," & sDescuento & ")"
'
'         Dbs.Execute qry
'         ' fin insertar
'
'         ProBar.Value = ProBar.Value + 1
'
'      Next i
'
'      Dbs.CommitTrans
'      On Error GoTo Errores
'      ProBar.Value = ProBar.Max
       
      Call ImpReporte
      Conparqu.Enabled = True
'  End If
   
   
   
   Exit Sub
Err_Trans:
   Dbs.RollbackTrans
Errores:
   Conparqu.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   
   Mskfecini.Mask = ""
   Mskfecini.text = ""
   Mskfecini.Mask = "##/##/####"
   Mskfecfin.Mask = ""
   Mskfecfin.text = ""
   Mskfecfin.Mask = "##/##/####"
   chkValor.Value = 0
   OptFecha(4).Value = True
   For i = 0 To 4
      Chkcriterio(i).Value = 0
   Next i
   MskActivas.text = ""
   MskCanceladas.text = ""
   MskDescuento.text = ""
   Msktotal.text = ""
   ProBar.Value = ProBar.Min
   Mskfecini.SetFocus
   DIB_GRILLA

   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Chkcriterio_Click(index As Integer)

   
   If Chkcriterio(index).Value = 1 Then
      If index = 0 Then
         TxtPlaca.Enabled = True
      ElseIf index = 1 Then
         Txtcodemp.Enabled = True
      ElseIf index = 2 Then
         CmbCaja.Enabled = True
      ElseIf index = 3 Then
         TxtTurno.Enabled = True
      ElseIf index = 4 Then
         Cmbestado.Enabled = True
      ElseIf index = 5 Then
         CmbTipo.Enabled = True
      
      End If
      
   ElseIf Chkcriterio(index).Value = 0 Then
      
      If index = 0 Then
         TxtPlaca.Enabled = False
         TxtPlaca.text = ""
         LblPlaca.Caption = ""
      ElseIf index = 1 Then
         Txtcodemp.Enabled = False
         Txtcodemp.text = ""
         Lblcodemp.Caption = ""
      ElseIf index = 2 Then
         CmbCaja.Enabled = False
         CmbCaja.ListIndex = -1
      ElseIf index = 3 Then
         TxtTurno.Enabled = False
         TxtTurno.text = ""
         LblTurno.Caption = ""
      ElseIf index = 4 Then
         Cmbestado.Enabled = False
         Cmbestado.ListIndex = -1
      ElseIf index = 5 Then
         CmbTipo.Enabled = False
         CmbTipo.ListIndex = -1
      End If
      
   End If
  

End Sub

Private Sub Form_Activate()
   Call Cargar_Combo_Caja
   Call Aplicar_Perfil(Me, "Conparqu")
End Sub

Private Sub Form_Load()
   DIB_GRILLA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
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
      
      If Mskfecini.ClipText <> "" Then
         If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
            MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
            Mskfecini.SetFocus
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

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "43"
  Txtcodemp.text = ""
  Lblcodemp.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   On Error GoTo Errores
   If Txtcodemp.text <> "" Then
      fqry01 = "SELECT henombre " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & Txtcodemp.text & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub TxtPlaca_DblClick()
   forma = "2"
   forma1 = "15"
   TxtPlaca.text = ""
   LblPlaca.Caption = ""
   buscar.Show 1
   TxtPlaca_LostFocus
End Sub

Private Sub TxtPlaca_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub TxtPlaca_LostFocus()
   On Error GoTo Errores
   
   If Trim(TxtPlaca.text) <> "" Then
      fqry01 = "SELECT hvconact " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & TxtPlaca.text & "' "
       
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         LblPlaca.Caption = Ftb01!hvconact
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Placa De Vehiculo No Existe", 48, "Advertencia"
         TxtPlaca.SetFocus
         Exit Sub
      End If
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub TxtTurno_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtTurno_KeyPress(KeyAscii As Integer)
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub
