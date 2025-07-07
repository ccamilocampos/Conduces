VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Canmultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de Evasiones Y Elusiones"
   ClientHeight    =   7305
   ClientLeft      =   1260
   ClientTop       =   1200
   ClientWidth     =   9285
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rpt 
      Left            =   8640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton btnconduces 
      Caption         =   "&T. U"
      Enabled         =   0   'False
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
      Left            =   8400
      Picture         =   "Canmulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Venta de conduces"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton btnabonos 
      Caption         =   "&Abonos"
      Enabled         =   0   'False
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
      Left            =   8400
      Picture         =   "Canmulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Abonos"
      Top             =   1080
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
      Left            =   8400
      Picture         =   "Canmulta.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   3840
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
      Left            =   8400
      Picture         =   "Canmulta.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   4680
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
      Left            =   8400
      Picture         =   "Canmulta.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "limpia la forma actual"
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Btnsalir 
      Caption         =   "&Salir"
      Enabled         =   0   'False
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
      Left            =   8400
      Picture         =   "Canmulta.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir de la forma actual"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Btngrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
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
      Left            =   8400
      Picture         =   "Canmulta.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ingresa la empresa a la base de datos"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frmplanilla 
         Caption         =   "Datos Planilla"
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
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   3120
         TabIndex        =   73
         Top             =   1800
         Width           =   2775
         Begin VB.TextBox Txttiquete 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   76
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Txtplanilla 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   75
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Txtdespla 
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
            Left            =   960
            MaxLength       =   3
            TabIndex        =   74
            ToolTipText     =   "Código de la Ruta, Presione DblClick para Ayuda"
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Mskvalpla 
            Height          =   285
            Left            =   1320
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Valor conduce"
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483644
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0;(#,##0)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Caption         =   "Tiquete:"
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
            TabIndex        =   82
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Planilla:"
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
            TabIndex        =   81
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "V/r Tasa Uso"
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
            TabIndex        =   80
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Destino:"
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
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lbldespla 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame Frmdestino 
         Caption         =   "Datos Viaje"
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
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   2775
         Begin VB.TextBox Txtnumcon 
            Height          =   285
            Left            =   840
            TabIndex        =   65
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Txtdestino 
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
            Left            =   960
            MaxLength       =   3
            TabIndex        =   64
            ToolTipText     =   "Código de la ruta (presione DblClick para ayuda)"
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Mskvalor 
            Height          =   285
            Left            =   840
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "Valor conduce"
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483644
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0;(#,##0)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Mskfecsal 
            Height          =   285
            Left            =   840
            TabIndex        =   67
            ToolTipText     =   "Fecha hasta donde quiere que busque"
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   16
            Format          =   "mmm/dd/yyyy HH:mm"
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha:"
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
            TabIndex        =   72
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "T.U"
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
            TabIndex        =   71
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Valor:"
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
            Index           =   9
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Destino:"
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
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lbldestino 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.TextBox Txtauxalt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   60
         ToolTipText     =   "Código del auxiliar de control que hizo el parte(doble Click)"
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtplaca 
         Height          =   285
         Left            =   120
         MaxLength       =   7
         TabIndex        =   0
         ToolTipText     =   "Número de placa del vehículo a quien se le hizo el parte(doble Click para ayuda)"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtcodinf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Txttipveh 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtsitio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtnomaux 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox Txtnomtes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   4200
         Width           =   2715
      End
      Begin VB.TextBox Txtnomnot 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   2715
      End
      Begin VB.TextBox Txtnumint 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Control 2 :"
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
         Left            =   3120
         TabIndex        =   62
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblauxalt 
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
         Left            =   4080
         TabIndex        =   61
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Lbltipveh 
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
         Left            =   1080
         TabIndex        =   59
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Lblsitio 
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
         Left            =   3960
         TabIndex        =   58
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa:"
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
         Left            =   2520
         TabIndex        =   50
         Top             =   240
         Width           =   1095
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
         Left            =   3720
         TabIndex        =   49
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Nro. Placa:"
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
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Interno:"
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
         Left            =   1320
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cód. Infracción:"
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
         TabIndex        =   46
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Lblcodinf 
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
         Left            =   2880
         TabIndex        =   45
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Tip. Vehículo:"
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
         TabIndex        =   44
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Sitio de la Infracción:"
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
         TabIndex        =   43
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Control 1:"
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
         TabIndex        =   42
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Nom. Testigo:"
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
         Left            =   3120
         TabIndex        =   41
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Nom. Notificado:"
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
         TabIndex        =   40
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Lblauxcon 
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
         Left            =   1080
         TabIndex        =   39
         Top             =   4800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Sistema"
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
      Height          =   4095
      Left            =   6240
      TabIndex        =   34
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Txtnumero 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Número de la multa(doble Click)"
         Top             =   600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskhora 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm "
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecsis 
         Height          =   285
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         Format          =   "mmm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecgra 
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         Format          =   "mmm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         Caption         =   "Fec Grabacion:"
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
         TabIndex        =   54
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Fec del Sistema:"
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
         TabIndex        =   52
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Evasion"
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
         TabIndex        =   37
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Número:"
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
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hora:"
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
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Observaciones"
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
      Height          =   1695
      Left            =   2520
      TabIndex        =   33
      Top             =   5400
      Width           =   3615
      Begin VB.TextBox Txtobservaciones 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
      Begin Crystal.CrystalReport rptterminal 
         Left            =   4200
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Estado de Evasion"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   32
      Top             =   5400
      Width           =   2295
      Begin VB.OptionButton Optestado 
         Caption         =   "Abonada"
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Optestado 
         Caption         =   "Pendiente"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Optestado 
         Caption         =   "Pagada"
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
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Optestado 
         Caption         =   "Anulado"
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Datos"
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
      Height          =   2775
      Left            =   6240
      TabIndex        =   24
      Top             =   4320
      Width           =   1935
      Begin VB.TextBox Txtrecpag 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Recibo de pago"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Txtdescuento 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin MSMask.MaskEdBox Mskvalmul 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msknetpag 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msksaldo 
         Height          =   285
         Left            =   840
         TabIndex        =   55
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo"
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
         Left            =   960
         TabIndex        =   56
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Vr. Evasion:"
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
         Left            =   480
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "% Des:"
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
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Recibo de Pago:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Neto a Pagar"
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
         Left            =   360
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      X1              =   8280
      X2              =   9360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   8280
      X2              =   8280
      Y1              =   0
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   9480
      Y1              =   7200
      Y2              =   7200
   End
End
Attribute VB_Name = "Canmultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim estado As String
Dim saldo As Double
Dim L_DescontarAviso  As String

Private Sub btnabonos_Click()
   'If Trim(Txtnumero.Text) <> "" And Txtnumero.Enabled = False Then
      'abonos.Txtnumero.Text = Txtnumero.Text
      abonos.Show 1
      Txtnumero.Enabled = True
      Txtnumero_LostFocus
   'End If
End Sub

Private Sub BtnBuscar_Click()
  forma = "7"
  forma1 = "2"
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
End Sub

Private Sub btnconduces_Click()
   On Error Resume Next

    On Error Resume Next
    If main.Mnuvencon.Enabled = True Then
      Unload Me
      Capconduce.Show 1
    End If

End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
   If txtPlaca.text = "" Then
      MsgBox "Falta Ingresar El Codigo De La Placa", 48, "Advertencia"
      txtPlaca.SetFocus
      Exit Sub
   End If
  
   If Txtnumero.text = "" Then
      MsgBox "Falta Ingresar El Numero De La Evasion", 48, "Advertencia"
      Txtnumero.SetFocus
      Exit Sub
   End If
   
   If Txtrecpag.text = "" Or Txtrecpag.text = "0" Then
      MsgBox "Falta Ingresar El Recibo De Paga", 48, "Advertencia"
      Txtrecpag.SetFocus
      Exit Sub
   End If
   i = 1
salto:
   QRY = "SELECT MAX(abrecibo) " _
         & "FROM opabonos "
          
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      If IsNull(Tbl2(0)) Then
         Txtrecpag.text = 1
      Else
         Txtrecpag.text = Tbl2(0) + 1
      End If
   End If
   Tbl2.Close
   QRY = "SELECT * " _
         & "FROM opabonos " _
        & "WHERE abrecibo = " & Txtrecpag.text & ""
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
      Tbl1.Close
      i = i + 1
      If i = 5 Then
         MsgBox "El Número de Recibo de pago Ya Existe", 48, "Advertencia"
         Txtrecpag.SetFocus
         Exit Sub
      End If
      GoTo salto
   End If
   Tbl1.Close
   
   estado = "Pagada"
     
   respuesta = MsgBox("Desea Grabar La Evasion", 32 + vbYesNo + vbQuestion + vbDefaultButton2, "GRABAR")
  
   If respuesta = vbNo Then
      Exit Sub
   End If
  
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
           
   Set Ftb01 = Dbs.Execute(fqry01)
    
   If Not Ftb01.EOF Then
      hora = Mid(Ftb01!psval, 12, 5)
      hora = Format(hora, "hh:mm")
   End If
   Ftb01.Close
   
   no = "N"
   
   descue = Mskvalmul.text - Msknetpag.text
   cero = 0
   
   If Trim(Txtdescuento.text) = "" Then
      Txtdescuento.text = "0"
   End If
   
   'Sebastián Rondón - Febrero 27 de 2025
   'Si el vehículo no cuenta con rodamiento, preguntamos al operario si desea continuar con la operación
'   If Not L_Bandera_Rodamiento Then
'        G_Mensaje_Pago = "VEHÍCULO SIN RODAMIENTO PROGRAMADO, DESEA CONTINUAR EL PAGO?"
'        G_Mensaje_Tipo = "E"
'        G_Mensaje_Opc = False
'        AviRecMult.Show 1
'
'        If Not G_Mensaje_Opc Then
'            Exit Sub
'        End If
'   End If

    'Sebastián Rondón - Abril 02 de 2025
    'Mensaje solamente informativo para confirmar si se desea realizar el proceso
'    G_Mensaje_Pago = "¿DESEA REALIZAR EL PROCESO PARA LA PLACA " & txtPlaca.text & "? "
'    G_Mensaje_Tipo = "E"
'    G_Mensaje_Opc = False
'
'    AviRecMult.Show 1
'
'    If Not G_Mensaje_Opc Then
'        Exit Sub
'    End If
   
   
   Dbs.BeginTrans
   On Error GoTo transac
   
   QRY = "UPDATE opmultas " _
          & "SET muestado = '" & estado & "'," _
          & "    murecpag =  " & Txtrecpag.text & "," _
          & "    muturno  =  " & G_Turno & "," _
          & "    muoperar = '" & Login & "'," _
          & "    muhora   = to_Date('" & hora & "','hh24:mi:ss')," _
          & "    mucerrado = '" & no & "'," _
          & "    muciedia  = '" & no & "'," _
          & "    mufecsis = TO_DATE('" & Mskfecsis.text & "','MM/DD/YYYY') ," _
          & "    munetpag = " & Msknetpag.text & " ," _
          & "    muvaldes = " & descue & " ," _
          & "    musaldo = " & cero & ", " _
          & "    munomcaj = '" & Pc & "', " _
          & "    mudescue = " & Txtdescuento.text & " " _
        & "WHERE munumero = " & Txtnumero.text & " "
    
   Dbs.Execute QRY
    
   'se inserta en lqa tabla de abonos
   'con el fin de hacer mas facil la busqueda
   'en los cuadres de caja y de dia
   'wecc 08/05/2002
   
   QRY = "INSERT INTO opabonos " _
                 & "(abnumero, abrecibo, abfecha, " _
                 & " abvalor, abfuncio, abcerrado, " _
                 & " abciedia, abturno, abnomcaj) VALUES (" _
                 & Txtnumero.text & "," & Txtrecpag.text & ",TO_DATE('" & Mskfecsis.text & "','MM/DD/YYYY') ," _
                 & MskSaldo.text & ",'" & Login & "','N', 'N'," _
                 & G_Turno & ", '" & Pc & "')"
   
   Dbs.Execute QRY
   
   QRY = "UPDATE geparsis " _
          & "SET psval = '" & Txtrecpag.text & "' " _
        & "WHERE pscod = 'NUMRECPAG' "
   
   Dbs.Execute QRY
   
   If L_DescontarAviso = "S" Then
      QRY = "UPDATE opmultas " _
             & "SET muaviso = muaviso - 1 " _
           & "WHERE munumpla = '" & txtPlaca.text & "' " _
             & "AND muestado = 'Pendiente' "
      Dbs.Execute QRY
   End If
   
   'WECC 0813/2019 CAU No 1254
   '**********************************************************************
'   QRY = "UPDATE optempo1 " _
'          & "SET temnumero2 = temnumero2 + " & Msksaldo.text & " " _
'        & "WHERE temtext4 = '" & Trim(Txtplaca.text) & "' " _
'          & "AND temnumero3 = " & Txtnumero.text & " " _
'          & "AND tempc = '" & Trim(Pc) & "'"
'
'   Dbs.Execute QRY, sRecordAffected
'
'   If sRecordAffected = 0 Then
'      QRY = "INSERT INTO optempo1(temtext4, temnumero2, temnumero3, tempc)" _
'               & "VALUES('" & Trim(Txtplaca.text) & "'," & Msksaldo.text & "," & Txtnumero.text & ", '" & Trim(Pc) & "' )"
'
'      Dbs.Execute QRY
'   End If
   '**********************************************************************
        
   'Sebastián Rondón - Marzo 20 de 2024
   'Por orden del jefe Álvaro se cambia el pago de las permanencias y multas
   'Si una infracción es pagada en la terminal Norte y Sur, se crea automaticamente una tasa sin viaje
   'Una tasa sin viaje es una tasa con valor 0, la cual solo se crea, para asociar el pago de la multa, y esta se pueda enviar a la DIAN
   'Si es un pago en la terminal Central, valida si tiene rodamiento, en cuyo caso no, crea de igual manera la tasa sin viaje
   
   'Sebastián Rondón - Febrero 26 de 2025
    'Se empieza a validar rodamientos para todas las terminales
   If Mid(G_Tipo_Terminal, 1, 1) = "N" Or Mid(G_Tipo_Terminal, 1, 1) = "S" Or Mid(G_Tipo_Terminal, 1, 1) = "C" Then
   
        'Sebastián Rondón - Abril 01 de 2025
        'Desde ahora, cada pago genera Tasa de Uso Sin Viaje siempre
        L_Bandera_Rodamiento = False
        If L_Bandera_Rodamiento Then
            'Felipe Mendoza - Marzo 12/2024
            'Se hace el insert de los pagos en una nueva tabla NO TEMPORAL
            QRY = "UPDATE opreccaja " _
                   & "SET rcvalmul = rcvalmul + " & MskSaldo.text & " " _
                 & "WHERE rcplaca = '" & Trim(txtPlaca.text) & "' " _
                   & "AND rcnumero = " & Txtnumero.text & " " _
                   & "AND rcnomcaj = '" & Trim(Pc) & "'"
            
            Dbs.Execute QRY, sRecordAffected
            
            If sRecordAffected = 0 Then
               QRY = "INSERT INTO opreccaja(rcplaca, rcvalmul, rcnumero, rcnomcaj, rcfecha, rchora)" _
                        & "VALUES('" & Trim(txtPlaca.text) & "'," & MskSaldo.text & ", " & Txtnumero.text & ", '" _
                        & Trim(Pc) & "', TO_DATE('" & Mskfecsis.text & "','MM/DD/YYYY'), TO_DATE('" & hora & "','HH24:MI') )"
            
               Dbs.Execute QRY
            End If
        Else
            sFechaCompleta = Mskfecsis.text & " " & hora
        
            'Consultamos el último consecutivo de conduces para la caja
            QRY = "SELECT caconcaj FROM cocaja WHERE canomcaj = '" & Pc & "' "
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                numconduce = Tbl1!caconcaj + 1
                
                'Creamos la Tasa de uso sin viaje
                QRY = "INSERT INTO coconduc (conumero, cocodrut, cocodemp, cofecsal, cohorven, cohorsal, coplaca, coclase, " _
                            & "coconduc, cooritra, cofuncio, coturno, coestado, comodo, conumint, coobser, conomcaj , " _
                            & "cocredit, coterminal, cohorasig) " _
                            & "VALUES(" & numconduce & ", '000', '" & Txtcodemp.text & "', TO_DATE('" & Mskfecsis.text & "','MM/DD/YYYY'), " _
                            & "TO_DATE('" & hora & "','HH24:MI'), TO_DATE('" & hora & "','HH24:MI'), '" & txtPlaca.text & "', '" & Txttipveh.text & "', " _
                            & "'0', 'SV', '" & Login & "', " & G_Turno & ", 'Activo', 'D', " & TxtNumInt.text & ", " _
                            & "'TASA DE USO PARA PAGO DE MULTA', '" & Pc & "', 'N', '" & Mid(G_Tipo_Terminal, 1, 1) & "', " _
                            & "TO_DATE('" & sFechaCompleta & "','MM/DD/YYYY HH24:MI')) "
                Dbs.Execute (QRY)
                
                'Asignamos la tasa de uso en la multa pagada
                QRY = "UPDATE opmultas " _
                        & "SET mutasapag = " & numconduce & " " _
                        & "WHERE munumero = " & Txtnumero.text & " "
                Dbs.Execute (QRY)
                
                'Aumentamos el número consecutivo del conduce
                QRY = "UPDATE cocaja " _
                        & "SET caconcaj = " & numconduce & " " _
                        & "WHERE canomcaj = '" & Pc & "' "
                Dbs.Execute (QRY)
                
                'Sebastián Rondón - Abril 02 de 2025
                'Registro en OPRECCAJA por si van a pagar la tasa de uso de al instante y se vea la multa en la calculadora

                '06/17/2025 se parcha por razones operativas
'                QRY = "UPDATE opreccaja " _
'                    & "SET rcvalmul = rcvalmul + " & MskSaldo.text & " " _
'                    & "WHERE rcplaca = '" & Trim(txtPlaca.text) & "' " _
'                    & "AND rcnumero = " & Txtnumero.text & " " _
'                    & "AND rcnomcaj = '" & Trim(Pc) & "'"
            
                Dbs.Execute QRY, sRecordAffected
                
                If sRecordAffected = 0 Then
                   QRY = "INSERT INTO opreccaja(rcplaca, rcvalmul, rcnumero, rcnomcaj, rcfecha, rchora)" _
                            & "VALUES('" & Trim(txtPlaca.text) & "'," & MskSaldo.text & ", " & Txtnumero.text & ", '" _
                            & Trim(Pc) & "', TO_DATE('" & Mskfecsis.text & "','MM/DD/YYYY'), TO_DATE('" & hora & "','HH24:MI') )"
                
                   Dbs.Execute QRY
                End If
                
                'Mostramos el mensaje de que el recaudo no se reflejará en la venta de la tasa
'                G_Mensaje_Pago = "POR FAVOR RECAUDE EL VALOR DE LA MULTA/PERMANENCIA, YA QUE NO SE VERÁ REFLEJADO EN LA VENTA DE LA TASA DE USO"
'                G_Mensaje_Tipo = "M"
'                AviRecMult.Show 1
            End If
        End If
   End If
   
   Dbs.CommitTrans
   On Error GoTo Errores
   MsgBox "Evasion " & estado, 64, "O.K."
          
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMRECPAG' "
           
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Txtrecpag.text = Ftb01!psval + 1
   End If
   Ftb01.Close
   
   Btngrabar.Enabled = False
   G_Numero_Comparendo_Antiguo = 0
   Btnimprimir_Click
   Btnlimpiar_Click
   Exit Sub
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
   Exit Sub
   
transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub
Private Sub Btnimprimir_Click()
   On Error GoTo Errores
      '-----------------------------------------------------------------------'
      '  DEFINICION DE VARIABLES PARA LA IMRESION DEL TIQUETE                 '
      '-----------------------------------------------------------------------'
      Dim sComparendo As String
      Dim sFecha As String
      Dim sHora As String
      Dim sNumeroInterno As String
      Dim SInfraccion As String
      Dim sPlaca As String
      Dim sNotificado As String
      Dim sFechaInf As String
      Dim sDestinoTUT As String
      Dim sDestinoPLA As String
      Dim sDestinoTIQ As String
      Dim sCajero As String
      Dim sCaja As String
      Dim sturno As String
      Dim sNivelServicio As String
      Dim sValorMulta As String
      Dim sDescuento As String
      Dim sValorTotal As String
   
   If Btngrabar.Enabled = False And Txtnumero.Enabled = False Then
  
      fqry01 = "SELECT psval " _
             & "FROM geparsis " _
             & "WHERE pscod = 'HORASIS' "
   
      Set Ftb01 = Dbs.Execute(fqry01)
    
      If Not Ftb01.EOF Then
         sHora = Format(Mid(Ftb01!psval, 12, 8), "HH:mm")
      End If
      Ftb01.Close
   
      sDescuento = CDbl(Mskvalmul.text) - CDbl(Msknetpag.text)
  
      
      '-----------------------------------------------------------------------'
      '  CARGA DE LAS VARIABLES PARA LA IMRESION DEL TIQUETE                  '
      '-----------------------------------------------------------------------'
      sComparendo = Txtnumero.text
      sFecha = Format(Mskfecsis.text, "MMM/dd/yyyy")
      sHora = Format(Mskhora.text, "HH:MM")
      SInfraccion = LblCodInf.Caption
      sNumeroInterno = TxtNumInt.text
      sPlaca = txtPlaca.text
      sNotificado = Txtnomnot.text
      sFechaInf = Format(Mskfecha.text, "MMM/dd/yyyy")
      sDestinoTUT = Lbldestino.Caption
      sDestinoPLA = Lbldespla.Caption
      sDestinoTIQ = Txttiquete.text
      sCajero = Login
      sCaja = Pc
      sturno = G_Turno
      sNivelServicio = Lbltipveh.Caption
      sValorMulta = Mskvalmul.FormattedText
      sDescuento = MskSaldo.FormattedText
      sValorTotal = Msknetpag.FormattedText
      sCodBarras = "*" & Txtnumero.text & "*"
      
      '-----------------------------------------------------------------------'
      '  IMPRESION DEL TIQUETE NUEVO FORMATO                                  '
      '-----------------------------------------------------------------------'
      If G_ImpCon = "Z" Then
         Open "LPT1" For Output As #1
         
         Print #1, "^XA"
         Print #1, "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ"
         Print #1, "~DG000.GRF,01536,016,"
         Print #1, ",::::::::X020,,U0IF80,S0540H01C,S0E0J03A0,R0380K070,Q01E0M08,gG06,Q0620M0380,P0180O040,P030P020,P040P030,P080P028,O010R0C,O020R06,O040R01,O0C0R01A0,N010T0C0,N030T0C0,N050T060,N060T0H2,N040T010,N080T010,N080U08,M0181FF20L0HFI08,M0101FF0M07F0H04,M0301FF0M0HFI0C,M0101DD0M0HDI06,M0203FF0M0HFI02,M0401770M0F70H02,M0603FE0M0HFI02,M040155DPD5,L02601FSF02,M040177FPFE,M0400FRFE01,M0400F55DODE01,M0E00FRFE01,M04001557F7F7F7F7F601,M0C00BFQFE01,M040055DPD601,M0E01FRFE03,M0401FF7FF77FKFE02,M0C03FRFE02,M0401DD7DODC,M0403FF9FOFE02,M0403FE0M0H7I02,M0603FE0M0HFI02,M0401DC0M05D0H06,L02303FE0M0FE0H02,M0303FE0M07C0H0C,M0303FE0M0FE0H0C,M0101C0N05C0H08,M0380U02,gJ010,N0C0T030,N040,N060T060,N050T040,N030T080,O0C0R01,O0E0R02,O060R06,O020R08,O01,O01A0P030,P050P040,P0380O080,P0140N01,Q060N0E,Q0180L010,R0E0L0E0,S0C0J07,S03E0H038,T0755DE0,,::::::::::::~DG001.GRF,01536,016,"
         Print #1, ",::::::::V020,,S0IF80,Q0540H01C,Q0E0J03A0,P0380K070,O01E0M08,Y06,O0620M0380,N0180O040,N030P020,N040P030,N080P028,M010R0C,M020R06,M040R01,M0C0R01A0,L010T0C0,L030T0C0,L050T060,L060T0H2,L040T010,L080T010,L080U08,K0181FF20L0HFI08,K0101FF0M07F0H04,K0301FF0M0HFI0C,K0101DD0M0HDI06,K0203FF0M0HFI02,K0401770M0F70H02,K0603FE0M0HFI02,K040155DPD5,J02601FSF02,K040177FPFE,K0400FRFE01,K0400F55DODE01,K0E00FRFE01,K04001557F7F7F7F7F601,K0C00BFQFE01,K040055DPD601,K0E01FRFE03,K0401FF7FF77FKFE02,K0C03FRFE02,K0401DD7DODC,K0403FF9FOFE02,K0403FE0M0H7I02,K0603FE0M0HFI02,K0401DC0M05D0H06,J02303FE0M0FE0H02,K0303FE0M07C0H0C,K0303FE0M0FE0H0C,K0101C0N05C0H08,K0380U02,gH010,L0C0T030,L040,L060T060,L050T040,L030T080,M0C0R01,M0E0R02,M060R06,M020R08,M01,M01A0P030,N050P040,N0380O080,N0140N01,O060N0E,O0180L010,P0E0L0E0,Q0C0J07,Q03E0H038,R0755DE0,,::::::::::::^XA^LL0743"
         Print #1, "^PW815"
         Print #1, "^FT288,128^XG000.GRF,1,1^FS"
         Print #1, "^FT704,128^XG001.GRF,1,1^FS"
         
         '-----------------------------------------------------------------------'
         '  IMPRESION DEL PRIMER TIQUETE                                         '
         '-----------------------------------------------------------------------'
         Print #1, "^FT29,499^A0R,20,19^FH\^FDVALOR A PAGAR:^FS"
         Print #1, "^FT29,661^A0R,20,19^FH\^FD" & sValorTotal & "^FS"
         Print #1, "^FT61,718^A0R,20,19^FH\^FD" & sDescuento & "^FS"
         Print #1, "^FT93,661^A0R,20,19^FH\^FD" & sValorMulta & "^FS"
         Print #1, "^FT59,499^A0R,20,19^FH\^FDDESCUENTO:^FS"
         Print #1, "^FT319,218^A0R,20,19^FH\^FDNIT. 860.052.155-6^FS"
         Print #1, "^FT91,500^A0R,20,19^FH\^FDVALOR EVASION:^FS"
         Print #1, "^FT193,660^A0R,20,19^FH\^FD" & sturno & "^FS"
         Print #1, "^FT195,525^A0R,20,19^FH\^FD" & sCaja & "^FS"
         Print #1, "^FT195,411^A0R,20,19^FH\^FD" & sCajero & "^FS"
         Print #1, "^FT162,441^A0R,20,19^FH\^FDNIVEL DE SERVICIO^FS"
         Print #1, "^FT346,165^A0R,25,24^FH\^FDTerminal de Transporte S.A.^FS"
         Print #1, "^FT93,26^A0R,20,19^FH\^FDDESTINO TIQUETE:^FS"
         Print #1, "^FT119,26^A0R,20,19^FH\^FDDESTINO PLANILLA:^FS"
         Print #1, "^FT145,26^A0R,20,19^FH\^FDDESTINO TUT:^FS"
         Print #1, "^FT171,26^A0R,20,19^FH\^FDFECHA INFRACCION:^FS"
         Print #1, "^FT198,27^A0R,20,19^FH\^FDNOMBRE NOTIFICADO:^FS"
         Print #1, "^FT223,15^A0R,20,19^FH\^FDINFRACCION:^FS"
         Print #1, "^FT135,443^A0R,20,19^FH\^FD" & sNivelServicio & "^FS"
         Print #1, "^FT262,572^A0R,20,19^FH\^FD" & sPlaca & "^FS"
         Print #1, "^FT93,212^A0R,20,19^FH\^FD" & sDestinoTIQ & "^FS"
         Print #1, "^FT342,605^A0R,34,33^FH\^FD" & sComparendo & "^FS"
         Print #1, "^FT118,212^A0R,20,19^FH\^FD" & sDestinoPLA & "^FS"
         Print #1, "^FT287,574^A0R,20,19^FH\^FDPlaca:^FS"
         Print #1, "^FT146,212^A0R,20,19^FH\^FD" & sDestinoTUT & "^FS"
         Print #1, "^FT260,427^A0R,20,19^FH\^FD" & sNumeroInterno & "^FS"
         Print #1, "^FT170,212^A0R,20,19^FH\^FD" & sFechaInf & "^FS"
         Print #1, "^FT198,211^A0R,20,19^FH\^FD" & sNotificado & "^FS"
         Print #1, "^FT257,240^A0R,20,19^FH\^FD" & sHora & "^FS"
         Print #1, "^FT224,128^A0R,20,19^FH\^FD" & SInfraccion & "^FS"
         Print #1, "^FT258,58^A0R,20,19^FH\^FD" & sFecha & "^FS"
         Print #1, "^FT223,638^A0R,20,19^FH\^FDTURNO^FS"
         Print #1, "^FT286,414^A0R,20,19^FH\^FDNo. INT:^FS"
         Print #1, "^FT224,529^A0R,20,19^FH\^FDCAJA^FS"
         Print #1, "^FT225,408^A0R,20,19^FH\^FDCAJERO^FS"
         Print #1, "^FT285,210^A0R,20,19^FH\^FDHora Salida:^FS"
         Print #1, "^FT285,80^A0R,20,19^FH\^FDFecha:^FS"
         Print #1, "^FT378,209^A0R,14,14^FH\^FDAlcaldia mayor de bogota D.C.^FS"
         Print #1, "^FT358,504^A0R,14,14^FH\^FDNumero^FS"
         Print #1, "^FT340,505^A0R,14,14^FH\^FDComparendo^FS"
         Print #1, "^BY2,2.5,61^FT21,259^B3R,N,,Y,N"
         Print #1, "^FD" & sComparendo & "^FS"
         
         '-----------------------------------------------------------------------'
         '  IMPRESION DEL SEGUNDO TIQUETE                                        '
         '-----------------------------------------------------------------------'
         Print #1, "^FT437,499^A0R,20,19^FH\^FDVALOR A PAGAR:^FS"
         Print #1, "^FT437,661^A0R,20,19^FH\^FD" & sValorTotal & "^FS"
         Print #1, "^FT469,718^A0R,20,19^FH\^FD" & sDescuento & "^FS"
         Print #1, "^FT501,661^A0R,20,19^FH\^FD" & sValorMulta & "^FS"
         Print #1, "^FT467,499^A0R,20,19^FH\^FDDESCUENTO:^FS"
         Print #1, "^FT727,218^A0R,20,19^FH\^FDNIT. 860.052.155-6^FS"
         Print #1, "^FT499,500^A0R,20,19^FH\^FDVALOR EVASION:^FS"
         Print #1, "^FT601,660^A0R,20,19^FH\^FD" & sturno & "^FS"
         Print #1, "^FT603,525^A0R,20,19^FH\^FD" & sCaja & "^FS"
         Print #1, "^FT603,411^A0R,20,19^FH\^FD" & sCajero & "^FS"
         Print #1, "^FT570,441^A0R,20,19^FH\^FDNIVEL DE SERVICIO^FS"
         Print #1, "^FT754,165^A0R,25,24^FH\^FDTerminal de Transporte S.A.^FS"
         Print #1, "^FT501,26^A0R,20,19^FH\^FDDESTINO TIQUETE:^FS"
         Print #1, "^FT527,26^A0R,20,19^FH\^FDDESTINO PLANILLA:^FS"
         Print #1, "^FT553,26^A0R,20,19^FH\^FDDESTINO TUT:^FS"
         Print #1, "^FT579,26^A0R,20,19^FH\^FDFECHA INFRACCION:^FS"
         Print #1, "^FT606,27^A0R,20,19^FH\^FDNOMBRE NOTIFICADO:^FS"
         Print #1, "^FT631,15^A0R,20,19^FH\^FDINFRACCION:^FS"
         Print #1, "^FT543,443^A0R,20,19^FH\^FD" & sNivelServicio & "^FS"
         Print #1, "^FT670,572^A0R,20,19^FH\^FD" & sPlaca & "^FS"
         Print #1, "^FT501,212^A0R,20,19^FH\^FD" & sDestinoTIQ & "^FS"
         Print #1, "^FT750,605^A0R,34,33^FH\^FD" & sComparendo & "^FS"
         Print #1, "^FT526,212^A0R,20,19^FH\^FD" & sDestinoPLA & "^FS"
         Print #1, "^FT695,574^A0R,20,19^FH\^FDPlaca:^FS"
         Print #1, "^FT554,212^A0R,20,19^FH\^FD" & sDestinoTUT & "^FS"
         Print #1, "^FT668,427^A0R,20,19^FH\^FD" & sNumeroInterno & "^FS"
         Print #1, "^FT578,212^A0R,20,19^FH\^FD" & sFechaInf & "^FS"
         Print #1, "^FT606,211^A0R,20,19^FH\^FD" & sNotificado & "^FS"
         Print #1, "^FT665,240^A0R,20,19^FH\^FD" & sHora & "^FS"
         Print #1, "^FT632,128^A0R,20,19^FH\^FD" & SInfraccion & "^FS"
         Print #1, "^FT666,58^A0R,20,19^FH\^FD" & sFecha & "^FS"
         Print #1, "^FT631,638^A0R,20,19^FH\^FDTURNO^FS"
         Print #1, "^FT694,414^A0R,20,19^FH\^FDNo. INT:^FS"
         Print #1, "^FT632,529^A0R,20,19^FH\^FDCAJA^FS"
         Print #1, "^FT633,408^A0R,20,19^FH\^FDCAJERO^FS"
         Print #1, "^FT693,210^A0R,20,19^FH\^FDHora Salida:^FS"
         Print #1, "^FT693,80^A0R,20,19^FH\^FDFecha:^FS"
         Print #1, "^FT786,209^A0R,14,14^FH\^FDAlcaldia mayor de bogota D.C.^FS"
         Print #1, "^FT766,504^A0R,14,14^FH\^FDNumero^FS"
         Print #1, "^FT748,505^A0R,14,14^FH\^FDComparendo^FS"
         Print #1, "^BY2,2.5,61^FT429,259^B3R,N,,Y,N"
         Print #1, "^FD" & sComparendo & "^FS"
         
         Print #1, "^PQ1,0,1,Y^XZ"
         Print #1, "^XA^ID000.GRF^FS^XZ"
         Print #1, "^XA^ID001.GRF^FS^XZ"
    
         Close #1
   
      End If
      If G_ImpCon = "C" Or G_ImpCon = "N" Then
'         imprimir.Show 1
'         If G_Imprimir = "0" Then
'            Exit Sub
'         End If


         fqry01 = "SELECT dsdes,dsval " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txtcodinf.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'INFRACCIONES PARA COBRAR'  "
               
         Set Ftb01 = Dbs.Execute(fqry01)
     
         If Not Ftb01.EOF Then
            sTipoInfraccion = Ftb01!dsval
         End If
         
         If sTipoInfraccion = 1 Then
            'multas de las 51 a la 54
            
            If G_ImpEpsonZebra = "P" Then
               reporte = RpPath + "\recmulta.rpt"
               sCantidad = 2
            ElseIf G_ImpEpsonZebra = "Z" Then
               reporte = RpPath + "\recmultazebra.rpt"
               sCantidad = 1
            End If
            
            
            Rpt.ReportFileName = reporte
            Rpt.Connect = G_Conexion_reportes
            For j = 1 To sCantidad
            
               Rpt.Formulas(0) = "nombreempresa = '" & G_NomEmpresa & "'"
               Rpt.Formulas(1) = "numcomparendo = '" & sComparendo & "'"
               Rpt.Formulas(2) = "hora = '" & sHora & "'"
               Rpt.Formulas(3) = "numint = '" & sNumeroInterno & "'"
               Rpt.Formulas(4) = "numplaca = '" & sPlaca & "'"
               Rpt.Formulas(5) = "infraccion = '" & SInfraccion & "'"
               Rpt.Formulas(6) = "notificado = '" & sNotificado & "'"
               Rpt.Formulas(7) = "fechainfraccion = '" & sFechaInf & "'"
               Rpt.Formulas(8) = "destinotut = '" & sDestinoTUT & "'"
               Rpt.Formulas(9) = "destinoplanilla = '" & sDestinoPLA & "'"
               Rpt.Formulas(10) = "destinotiquete = '" & sDestinoTIQ & "'"
               Rpt.Formulas(11) = "usuario = '" & sCajero & "'"
               Rpt.Formulas(12) = "pc = '" & sCaja & "'"
               Rpt.Formulas(13) = "turno = '" & sturno & "'"
               Rpt.Formulas(14) = "servicio = '" & sNivelServicio & "'"
               Rpt.Formulas(15) = "valmulta = '" & Format(sValorMulta, "$##,##0") & "'"
               Rpt.Formulas(16) = "descuento = '" & Format(sDescuento, "$##,##0") & "'"
               Rpt.Formulas(17) = "totalpagar = '" & Format(sValorTotal, "$##,##0") & "'"
               Rpt.Formulas(18) = "codigobarras = '" & sCodBarras & "'"
               Rpt.Formulas(19) = "fecha = '" & sFecha & "'"
               If G_ImpEpsonZebra = "P" Then
                  Rpt.PrinterName = G_PrinterName1
                  Rpt.PrinterDriver = G_PrinterDriver1
                  Rpt.PrinterPort = G_PrinterPort1
                  'Rpt.ProgressDialog = False
               End If
               '
               Rpt.Destination = 1
               Rpt.Action = 1
            Next j
         Else
            reporte = RpPath + "\recexcreta.rpt"
            Rpt.ReportFileName = reporte
            Rpt.Connect = G_Conexion_reportes
            For j = 1 To 2
            
               Rpt.Formulas(0) = "nombreempresa = '" & G_NomEmpresa & "'"
               Rpt.Formulas(1) = "numcomparendo = '" & sComparendo & "'"
               Rpt.Formulas(2) = "hora = '" & sHora & "'"
               Rpt.Formulas(3) = "numint = '" & sNumeroInterno & "'"
               Rpt.Formulas(4) = "numplaca = '" & sPlaca & "'"
               Rpt.Formulas(5) = "notificado = '" & sNotificado & "'"
               Rpt.Formulas(6) = "usuario = '" & sCajero & "'"
               Rpt.Formulas(7) = "pc = '" & sCaja & "'"
               Rpt.Formulas(8) = "turno = '" & sturno & "'"
               Rpt.Formulas(9) = "servicio = '" & sNivelServicio & "'"
               Rpt.Formulas(10) = "totalpagar = '" & Format(sValorTotal, "$##,##0") & "'"
               Rpt.Formulas(11) = "codigobarras = '" & sCodBarras & "'"
               Rpt.Formulas(12) = "fecha = '" & sFecha & "'"
               
               Rpt.Formulas(13) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(14) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(15) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(16) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(17) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(18) = "fecha = '" & sFecha & "'"
               Rpt.Formulas(19) = "fecha = '" & sFecha & "'"
               
               
               Rpt.PrinterName = G_PrinterName1
               Rpt.PrinterDriver = G_PrinterDriver1
               Rpt.PrinterPort = G_PrinterPort1
               'rpt.ProgressDialog = False
               Rpt.Destination = 1
               Rpt.Action = 1
            Next j
         End If
      End If
  End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

Private Sub Btnimprimir_original_Click()
   If Btngrabar.Enabled = False And Txtnumero.Enabled = False Then
     
     
     'qry = "INSERT INTO optempo1 " _
                            & "(temnumero1, temtext1, temnumero2, temtext2, temdate1, temtext3, temtext4, temtext5, temtext6, temtext7, temtext8, temnumdob1,  temtext9, temtext10, temnumdob2, temnumero3, temnumdob3, temdate2, tempc) " _
                    & "SELECT munumero,   mucodemp, mucodinf,   munomnot, mufecha, muhora,   munumpla,  musitio,  munomtes, muauxcon, mutipveh, mudescue,  muobserv, muestado,  munetpag,  murecpag, muvaldes, mufecsis, '" & pc & "' " _
                      & "FROM opmultas " _
                     & "WHERE munumero = " & Txtnumero.Text & ""
      
     '   dbs.Execute qry
     
      
     'abono = CDbl(Msknetpag.Text) - CDbl(Msksaldo.Text)
     'reporte = rppath + "\regmulta.rpt"
     'rptterminal.DataFiles(0) = dbpath
     'rptterminal.password = Passwd
     'rptterminal.ReportFileName = reporte
     'rptterminal.Formulas(0) = "empresa = '" & G_nomempresa & "'"
     'rptterminal.Formulas(1) = "emptra = '" & Lblcodemp & "'"
     'rptterminal.Formulas(2) = "infrac = '" & Lblcodinf & "'"
     'rptterminal.Formulas(3) = "abono = '" & Format(abono, "##,##0") & "'"
     'rptterminal.Formulas(4) = "saldo = '" & Msksaldo.FormattedText & "'"
    '
     'rptterminal.SelectionFormula = "{optempo1.tempc}= '" & pc & "' "
     'Canmultas.Enabled = False
     'rptterminal.Destination = 0
     'rptterminal.Action = 1
     'Canmultas.Enabled = True
     '   dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & pc & "' "
     
     'Exit Sub
  
      fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
   
      Set Ftb01 = Dbs.Execute(fqry01)
    
      If Not Ftb01.EOF Then
         sHora = Format(Mid(Ftb01!psval, 12, 8), "HH:mm")
      End If
      Ftb01.Close
   
      sDescuento = CDbl(Mskvalmul.text) - CDbl(Msknetpag.text)
  
      If G_ImpCon = "Z" Then
         Open "LPT1" For Output As #1
         
         Print #1, "^XA"
         Print #1, "^DFPRUEBA^FS"
         Print #1, "^PRC"
         Print #1, "^LH0,0^FS"
         Print #1, "^LL727"
         Print #1, "^MD0"
         Print #1, "^MNY"
         Print #1, "^LH0,0^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO102,362^AIB,25,0^CI0^FR^FDHORA^FS"
         Print #1, "^CWJ,110O01JW.FNT^FS"
         Print #1, "^FO25,248^AJB,34,0^CI0^FR^FDTerminal de Transporte S.A.^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO102,498^AIB,25,0^CI0^FR^FDFECHA^FS"
         Print #1, "^CWK,X00O01JW.FNT^FS"
         Print #1, "^FO53,325^AKB,31,0^CI0^FR^FDNIT.860.052.155-6^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO104,214^AIB,25,0^CI0^FR^FDNo. INT.^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO102,84^AIB,25,0^CI0^FR^FDPLACA^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO157,622^AIB,25,0^CI0^FR^FDINFRACCION^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO185,533^AIB,25,0^CI0^FR^FDNOMBRE NOTIFICADO^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO188,107^AIB,25,0^CI0^FR^FDNIVEL DE SERVICIO^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO245,636^AIB,25,0^CI0^FR^FDCAJERO^FS"
         Print #1, "^FO278,655^A0B,23,19^FD" & Login & "^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO245,544^AIB,25,0^CI0^FR^FDCAJA^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO245,406^AIB,25,0^CI0^FR^FDTURNO^FS"
         Print #1, "^CWL,T00S8_S2.FNT^FS"
         Print #1, "^FO8,302^ALB,22,0^CI0^FR^FDAlcaldía Mayor de Bogotá D.C.^FS"
         Print #1, "^FO135,569^A0B,23,19^CI13^FR^FD" & Format(Mskfecsis.text, "MMM/dd/yyyy") & "^FS"
         Print #1, "^FO135,369^A0B,23,19^CI13^FR^FD" & sHora & "^FS"
         Print #1, "^FO135,236^A0B,23,19^CI13^FR^FD" & TxtNumInt.text & "^FS"
         Print #1, "^FO135,84^A0B,23,19^CI13^FR^FD" & txtPlaca.text & "^FS"
         Print #1, "^FO165,398^A0B,23,19^CI13^FR^FD" & Txtcodinf.text & "-" & LblCodInf.Caption & " ^FS"
         Print #1, "^FO216,398^A0B,23,19^FD" & Txtnomnot.text & "^FS"
         Print #1, "^FO216,151^A0B,23,19^CI13^FR^FD" & Lbltipveh.Caption & "^FS"
         Print #1, "^FO275,541^A0B,23,19^CI13^FR^FD" & Pc & "^FS"
         Print #1, "^FO275,436^A0B,23,19^CI13^FR^FD" & G_Turno & "^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO246,147^AIB,25,0^CI0^FR^FDVALOR EVASION   $^FS"
         Print #1, "^FO254,89^A0B,23,19^CI13^FR^FD" & Mskvalmul.FormattedText & "^FS"
         Print #1, "^FO290,93^A0B,22,18^CI13^FR^FD" & sDescuento & "^FS"
         'Print #1, "^FO318,93^A0B,22,18^CI13^FR^FD" & Mskparqueadero.FormattedText & "^FS"
         Print #1, "^FO346,93^A0B,22,18^CI13^FR^FD" & Msknetpag.FormattedText & "^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO284,155^AMB,23,0^CI0^FR^FDDESCUENTO^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         'Print #1, "^FO312,177^AMB,23,0^CI0^FR^FDPARQUEADERO^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO338,169^AMB,23,0^CI0^FR^FDVALOR A PAGAR^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         'Print #1, "^FO300,480^AMB,23,0^CI0^FR^FDPRUEBA DE ALCOHOLIMETRIA^FS"
         Print #1, "^BY2,3.0^FO330,410^BCB,56,N,Y,N^FR^FD>:" & Txtnumero.text & "^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO25,163^AMB,23,0^CI0^FR^FDNúmero ^FS"
         Print #1, "^FO45,163^AMB,23,0^CI0^FR^FD Comparendo^FS"
         Print #1, "^FO33,21^A0B,37,31^CI13^FR^FD" & Txtnumero.text & "^FS"
         'Print #1, "^FO306,442^A0B,23,19^CI13^FR^FD" & presenprueba & "^FS"
         Print #1, "^FO8,631^FR^XGWJR6CTT,1,1^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO510,362^AIB,25,0^CI0^FR^FDHORA^FS"
         Print #1, "^CWJ,110O01JW.FNT^FS"
         Print #1, "^FO433,248^AJB,34,0^CI0^FR^FDTerminal de Transporte S.A.^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO510,498^AIB,25,0^CI0^FR^FDFECHA^FS"
         Print #1, "^CWK,X00O01JW.FNT^FS"
         Print #1, "^FO461,325^AKB,31,0^CI0^FR^FDNIT.860.052.155-6^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO512,214^AIB,25,0^CI0^FR^FDNo. INT.^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO510,84^AIB,25,0^CI0^FR^FDPLACA^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO565,622^AIB,25,0^CI0^FR^FDINFRACCION^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO593,533^AIB,25,0^CI0^FR^FDNOMBRE NOTIFICADO^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO595,107^AIB,25,0^CI0^FR^FDNIVEL DE SERVICIO^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO653,636^AIB,25,0^CI0^FR^FDCAJERO^FS"
         Print #1, "^FO685,655^A0B,23,19^FD" & Login & "^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO653,544^AIB,25,0^CI0^FR^FDCAJA^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO653,406^AIB,25,0^CI0^FR^FDTURNO^FS"
         Print #1, "^CWL,T00S8_S2.FNT^FS"
         Print #1, "^FO416,302^ALB,22,0^CI0^FR^FDAlcaldía Mayor de Bogotá D.C.^FS"
         Print #1, "^FO543,569^A0B,23,19^CI13^FR^FD" & Format(Mskfecsis.text, "MMM/dd/yyyy") & "^FS"
         Print #1, "^FO543,369^A0B,23,19^CI13^FR^FD" & sHora & "^FS"
         Print #1, "^FO543,236^A0B,23,19^CI13^FR^FD" & TxtNumInt.text & "^FS"
         Print #1, "^FO543,84^A0B,23,19^CI13^FR^FD" & txtPlaca.text & "^FS"
         Print #1, "^FO572,398^A0B,23,19^CI13^FR^FD" & Txtcodinf.text & "-" & LblCodInf.Caption & " ^FS"
         Print #1, "^FO624,398^A0B,23,19^FD" & Txtnomnot.text & "^FS"
         Print #1, "^FO624,151^A0B,23,19^CI13^FR^FD" & Lbltipveh.Caption & "^FS"
         Print #1, "^FO683,541^A0B,23,19^CI13^FR^FD" & Pc & "^FS"
         Print #1, "^FO683,436^A0B,23,19^CI13^FR^FD" & G_Turno & "^FS"
         Print #1, "^CWI,P00O01JW.FNT^FS"
         Print #1, "^FO654,147^AIB,25,0^CI0^FR^FDVALOR EVASION   $^FS"
         Print #1, "^FO661,89^A0B,23,19^CI13^FR^FD" & Mskvalmul.FormattedText & "^FS"
         Print #1, "^FO697,93^A0B,22,18^CI13^FR^FD" & sDescuento & "^FS"
         'Print #1, "^FO725,93^A0B,22,18^CI13^FR^FD" & Mskparqueadero.FormattedText & "^FS"
         Print #1, "^FO753,93^A0B,22,18^CI13^FR^FD" & Msknetpag.FormattedText & "^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO691,155^AMB,23,0^CI0^FR^FDDESCUENTO^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         'Print #1, "^FO719,177^AMB,23,0^CI0^FR^FDPARQUEADERO^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO746,169^AMB,23,0^CI0^FR^FDVALOR A PAGAR^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         'Print #1, "^FO707,480^AMB,23,0^CI0^FR^FDPRUEBA DE ALCOHOLIMETRIA^FS"
         Print #1, "^BY2,3.0^FO738,410^BCB,56,N,Y,N^FR^FD>:" & Txtnumero.text & "^FS"
         Print #1, "^CWM,L00O01JW.FNT^FS"
         Print #1, "^FO433,163^AMB,23,0^CI0^FR^FDNúmero ^FS"
         Print #1, "^FO453,163^AMB,23,0^CI0^FR^FD Comparendo^FS"
         Print #1, "^FO441,21^A0B,37,31^CI13^FR^FD" & Txtnumero.text & "^FS"
         'Print #1, "^FO713,442^A0B,23,19^CI13^FR^FD" & presenprueba & "^FS"
         Print #1, "^FO416,631^FR^XGWJR6CTT,1,1^FS"
         Print #1, "^XZ"
         Print #1, "^FX End of job"
         Print #1, "^XA"
         Print #1, "^XFPRUEBA"
         Print #1, "^XZ"
    
         
         Close #1
   
      End If
  End If
End Sub

Private Sub Btnlimpiar_Click()
   G_Numero_Comparendo_Antiguo = 0
   txtPlaca.text = ""
   txtPlaca.Enabled = True
   TxtNumInt.text = ""
   Txtcodemp.text = ""
   Lblcodemp.Caption = ""
   Txttipveh.text = ""
   Lbltipveh.Caption = ""
   Txtcodinf.text = ""
   LblCodInf.Caption = ""
   Txtsitio.text = ""
   Txtnomnot.text = ""
   Txtnomtes.text = ""
   Txtnomaux.text = ""
   Lblauxcon.Caption = ""
   TxtObservaciones.text = ""
   Txtnumero.Enabled = True
   Txtnumero.text = ""
   Mskfecha.Mask = ""
   Mskfecha.text = ""
   Mskfecha.Mask = "##/##/####"
   Mskhora.text = ""
   Txtrecpag.text = ""
   Mskvalmul.text = ""
   Txtdescuento.text = ""
   Msknetpag.text = ""
   Btngrabar.Enabled = False
   OptEstado(1).Enabled = True
   txtPlaca.Enabled = True
   txtPlaca.SetFocus
   'Txtrecpag.Enabled = True
   Mskfecsis.text = ""
   MskSaldo.text = ""

   Lblsitio.Caption = ""
   Txtauxalt.text = ""
   lblauxalt.Caption = ""
   Frmdestino.Enabled = False
   Frmplanilla.Enabled = False
   Txtdespla.text = ""
   Txtdestino.text = ""
   Mskfecsal.Mask = ""
   Mskfecsal.text = ""
   Mskfecsal.Mask = "##/##/#### ##:##"
   Mskvalor.text = ""
   Mskvalpla.text = ""
   Txtnumcon.text = ""
   Lbldestino.Caption = ""
   Lbldespla.Caption = ""
   Txtplanilla.text = ""
   Txttiquete.text = ""
   
   btnabonos.Enabled = False
   OptEstado(1).Value = True
   OptEstado(3).Enabled = False
   OptEstado(0).Enabled = False
   'fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'NUMRECPAG' "
          
   'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
   
   'If Not ftb01.EOF Then
   '   Txtrecpag.Text = ftb01!psval + 1
   'End If
   'ftb01.Close
   
   QRY = "SELECT MAX(abrecibo) " _
         & "FROM opabonos "
          
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      If IsNull(Tbl2(0)) Then
         Txtrecpag.text = 1
      Else
         Txtrecpag.text = Tbl2(0) + 1
      End If
   End If
   Tbl2.Close
   
   fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'HORASIS' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
   End If
   Ftb01.Close
   If main.Mnuvencon.Enabled = True Then
      btnconduces.Enabled = True
   Else
      btnconduces.Enabled = False
   End If
End Sub

Private Sub btnSalir_Click()
   G_Numero_Comparendo_Antiguo = 0
   
   
   Unload Me
End Sub


Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Canmultas")
  fqry01 = "SELECT psval " _
          & "FROM geparsis " _
         & "WHERE pscod = 'CAJCERCON' "
         
  Set Ftb01 = Dbs.Execute(fqry01)
  
  If Not Ftb01.EOF Then
     If Ftb01!psval = "S" Then
        Ftb01.Close
        MsgBox "LA CAJA DE TASAS DE USO ESTA CERRADA, NO SE PUEDE ABRIR LA FORMA", 48, "Advertencia"
        Unload Me
     End If
  End If
  Ftb01.Close
  
   If G_Numero_Comparendo_Antiguo <> 0 Then
      Txtnumero.text = G_Numero_Comparendo_Antiguo
      Txtnumero_LostFocus
   End If
   
   
   
 
   
End Sub


Private Sub Form_Load()
   
  
   'fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMRECPAG' " _
             & "AND psdes = 'NUMERO DE RECIBO DE PAGO'"
  
   'If Not ftb01.EOF Then
   '   Txtrecpag.Text = ftb01!psval + 1
   'End If
   'ftb01.Close
   
   QRY = "SELECT MAX(abrecibo) " _
         & "FROM opabonos "
          
   Set Tbl2 = Dbs.Execute(QRY)
   
   If Not Tbl2.EOF Then
      If IsNull(Tbl2(0)) Then
         Txtrecpag.text = 1
      Else
         Txtrecpag.text = Tbl2(0) + 1
      End If
   End If
   Tbl2.Close
   
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
   
   
   Set Ftb01 = Dbs.Execute(fqry01)
    
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
      
   End If
   Ftb01.Close
   
   L_DescontarAviso = Fn_Parametros("DESAVIPAGA", 0)
   If L_DescontarAviso = "" Then
      L_DescontarAviso = "N"
   End If
   
   If main.Mnuvencon.Enabled = True Then
      btnconduces.Enabled = True
   Else
      btnconduces.Enabled = False
   End If
   
End Sub

Private Sub Txtauxalt_DblClick()
  forma = "15"
  forma1 = "25"
  Tabla_Catalogo = "AUXILIARES DE CONTROL"
  Txtauxalt.text = ""
  buscar.Show 1
  Txtauxalt_LostFocus

End Sub

Private Sub Txtauxalt_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(LCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub Txtauxalt_LostFocus()
   On Error GoTo Err_Txtauxalt_LostFocus
   If Txtauxalt.text <> "" Then
   
       
      fqry01 = "SELECT * " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtauxalt.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'AUXILIARES DE CONTROL' "
               
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         lblauxalt.Caption = Ftb01!dsdes
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Auxiliar De Control No Existe", 48, "Advertencia"
         Txtauxalt_DblClick
      End If
   End If
   Exit Sub
Err_Txtauxalt_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Txtdespla_DblClick()
  forma = "3"
  forma1 = "20"
  Txtdespla.text = ""
  buscar.Show 1
  Txtdespla.SetFocus
  Txtdespla_LostFocus
End Sub

Private Sub Txtdespla_KeyPress(KeyAscii As Integer)
   If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If

End Sub

Private Sub Txtdespla_LostFocus()
   On Error GoTo ERR_Txtdespla_LostFocus
   If Txtdespla.text <> "" And Txtdespla.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Txtdespla.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb01!ruestado = "Inactiva" Then
            Ftb01.Close
            On Error Resume Next
            MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
            Txtdespla.SetFocus
            Exit Sub
         End If
        
         Lbldespla.Caption = Ftb01!rudestino
         Ftb01.Close
         
         QRY = "SELECT * " _
               & "FROM opconvenios " _
              & "WHERE cvcodrut = '" & Txtdespla.text & "' " _
                & "AND cvcodser = '" & Txttipveh.text & "' "
                
         
         Set Ftb04 = Dbs.Execute(QRY)
      
         If Not Ftb04.EOF Then
            Mskvalpla.text = Ftb04!cvvalor
            Ftb04.Close
            Exit Sub
             
         Else
            Ftb04.Close
            Mskvalpla.text = 0
         End If
         'fin ahorro tiempo
       
      Else
        
         Txtdespla.SetFocus
         MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
         Txtdespla.text = ""
        
      End If
        
   End If
   
   Exit Sub
   
ERR_Txtdespla_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
Resume Next
End Sub

Private Sub Txtdestino_DblClick()
 
  forma = "3"
  forma1 = "19"
  Txtdestino.text = ""
  buscar.Show 1
  Txtdestino.SetFocus
  Txtdestino_LostFocus
End Sub

Private Sub Txtdestino_KeyPress(KeyAscii As Integer)
  
   If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If

End Sub

Private Sub Txtdestino_LostFocus()
   On Error GoTo ERR_Txtdestino_LostFocus
   If Txtdestino.text <> "" And Txttipveh.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Txtdestino.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb01!ruestado = "Inactiva" Then
            Ftb01.Close
            Txtdestino.SetFocus
            MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
            Exit Sub
         End If
        
         Lbldestino.Caption = Ftb01!rudestino
         Ftb01.Close
          
         QRY = "SELECT * " _
               & "FROM opconvenios " _
              & "WHERE cvcodrut = '" & Txtdestino.text & "' " _
                & "AND cvcodser = '" & Txttipveh.text & "' "
                
         Set Ftb04 = Dbs.Execute(QRY)
      
         If Not Ftb04.EOF Then
            Mskvalor.text = Ftb04!cvvalor
            Ftb04.Close
            Exit Sub
             
         Else
            Ftb04.Close
            Mskvalor.text = 0
         End If
         'fin ahorro tiempo
       
      Else
        
         Txtdestino.SetFocus
         MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
         Txtdestino.text = ""
        
      End If
        
   End If
   
   Exit Sub
   
ERR_Txtdestino_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
   Exit Sub
End Sub

Private Sub Txtnumcon_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub Txtnumcon_LostFocus()
   On Error GoTo Err_Txtnumcon_LostFocus
   
   If Trim(Txtnumcon.text) <> "" Then
      
      QRY = "SELECT * " _
            & "FROM coconduc " _
           & "WHERE conumero = " & Trim(Txtnumcon.text) & ""
      
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         Mskfecsal.text = Format(Tbl1!cohorsal, "MM/dd/yyyy HH:mm")
         
         Tbl1.Close
      Else
         Tbl1.Close
         MsgBox "NUMERO DE TASA DE USO no existe", 48, "Advertencia"
        
         On Error Resume Next
         Txtnumcon.SetFocus
         Exit Sub
      End If
      
      
   End If
   
   Exit Sub
Err_Txtnumcon_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub


Private Sub TxtNumero_GotFocus()
   Txtnumero.SelStart = 0
   Txtnumero.SelLength = Len(Txtnumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumero_LostFocus()
   If Txtnumero.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM opmultas " _
              & "WHERE munumero = " & Txtnumero.text & " " _
                '& "AND (muestado = 'Pendiente' or muestado = 'Abonada') "
     
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         sEstadoMulta = Ftb01!muestado
         sNombreCaja = Ftb01!munomcaj
         If Ftb01!muestado = "Pagada" Then
            MsgBox "Numero de la Evasion esta pagada", 48, "Advertencia"
            'Txtnumero.SetFocus
            'Exit Sub
         ElseIf Ftb01!muestado = "Anulado" Then
            MsgBox "Numero de la Evasion esta anulada", 48, "Advertencia"
            'Txtnumero.SetFocus
            'Exit Sub
         
         ElseIf Ftb01!muestado = "Educativo" Then
            MsgBox "Numero de la Evasion es educativa", 48, "Advertencia"
            'Txtnumero.SetFocus
            'Exit Sub
         
         End If
         
         btnabonos.Enabled = True
         Txtnumero.Enabled = False
         txtPlaca.Enabled = False
         txtPlaca.text = Ftb01!munumpla
         Txtnumero.text = Ftb01!munumero
         Txtcodemp.text = Ftb01!mucodemp
         Txtcodinf.text = Ftb01!mucodinf
         Txtnomnot.text = Ftb01!munomnot
         Mskfecha.text = Ftb01!mufecha
         Mskhora.text = Ftb01!muhora
         Txtsitio.text = Ftb01!musitio
         Txtnomtes.text = Ftb01!munomtes
         Txtnomaux.text = Ftb01!muauxcon
         Txttipveh.text = Ftb01!mutipveh
         Txtdescuento.text = Ftb01!mudescue
         TxtObservaciones.text = Ftb01!muobserv
         saldo = Ftb01!musaldo    'wecc 08/06/2002
         Msknetpag.text = Ftb01!munetpag
         fecmul = Ftb01!mufecsis    ' ahh 04/03/2002
         Mskfecgra.text = Ftb01!mufecsis   'ahh   04/04/2002
         Mskvalmul.text = Ftb01!munetpag
         
         
         If Ftb01!murecpag <> "" Then
            
            If Ftb01!murecpag <> 0 Then
               Txtrecpag.text = Ftb01!murecpag
               Txtrecpag_LostFocus
            End If
         End If
        
         If Ftb01!muestado = "Abonada" Then
            OptEstado(3).Enabled = True
            OptEstado(3).Value = True
            OptEstado(1).Enabled = False
         ElseIf Ftb01!muestado = "Pendiente" Then
            OptEstado(0).Enabled = True
            OptEstado(0).Value = True
            OptEstado(1).Enabled = False
         End If
                               
         If Not IsNull(Ftb01!muauxalt) Then
            Txtauxalt.text = Ftb01!muauxalt
         End If
             
         If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            If Not IsNull(Ftb01!mudestino) Then
                Txtdestino.text = Ftb01!mudestino
            End If
            If Not IsNull(Ftb01!muconduce) Then
                Txtnumcon.text = Ftb01!muconduce
            End If
            
            If Not IsNull(Ftb01!mudespla) Then
                Txtdespla.text = Ftb01!mudespla
            End If
            
            If Not IsNull(Ftb01!muplanilla) Then
                Txtplanilla.text = Ftb01!muplanilla
            End If
            
            
            If Not IsNull(Ftb01!mutiquete) Then
                Txttiquete.text = Ftb01!mutiquete
            End If
            
            If Not IsNull(Ftb01!muvalcon) Then
                sValdes = Ftb01!muvalcon
            Else
                sValdes = 0
            End If
            
            If Not IsNull(Ftb01!muvalpla) Then
                sValPla1 = Ftb01!muvalpla
            Else
                sValPla1 = 0
            End If
            Ftb01.Close
            Txtdestino_LostFocus
            Txtnumcon_LostFocus
            Txtdespla_LostFocus
            
            Mskvalor.text = sValdes
            Mskvalpla.text = sValPla1
         Else
            Ftb01.Close
         End If
                               
         
                               
         fqry01 = "SELECT * " _
                  & "FROM ophojveh " _
                 & "WHERE hvplaca = '" & txtPlaca.text & "'"
              
   
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            TxtNumInt.text = Ftb01!hvnumint
         End If
         Ftb01.Close
                
         fqry02 = "SELECT * " _
                  & "FROM geparsis " _
                 & "WHERE pscod = '" & Txtcodinf.text & "' "
         
         Set Ftb02 = Dbs.Execute(fqry02)
         
         If Not Ftb02.EOF Then
            LblCodInf.Caption = Ftb02!psdes
         
            'fqry01 = "SELECT psval " _
                     & "FROM geparsis " _
                    & "WHERE pscod = 'SALMIN' "
                 
            'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
         
            'If Not ftb01.EOF Then
            '   Mskvalmul.Text = CLng((ftb01!psval / 30) * ftb02!psval)
            '   numero = Str(Mskvalmul.Text)
            '   digitos = Len(numero)
            '   inicio1 = digitos - 1
            '   valor = Mid(numero, inicio1, 2)
            '   If valor <> "00" Then
            '      complemento = 100 - valor
            '      valtot = Abs(complemento) + numero
            '      Mskvalmul.Text = valtot
            '   End If
            'End If
         End If
         Ftb02.Close
         
         fqry01 = "SELECT * " _
                  & "FROM ophojemp " _
                 & "WHERE hecodigo = '" & Txtcodemp.text & "'"
              
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lblcodemp.Caption = Ftb01!henombre
         Else
            Ftb01.Close
            fqry01 = "SELECT henombre " _
                    & "FROM ophojempcon " _
                   & "WHERE hecodigo = '" & Txtcodemp.text & "'"
        
            Set Ftb01 = Dbs.Execute(fqry01)
   
            If Not Ftb01.EOF Then
               Lblcodemp.Caption = Ftb01!henombre
            End If
            Ftb01.Close
         End If
         
         
         fqry01 = "SELECT dsdes " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txttipveh.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'TIPOS DE VEHICULO' "
                
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lbltipveh.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         
         fqry01 = "SELECT * " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txtnomaux.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'AUXILIARES DE CONTROL' "
                
         Set Ftb01 = Dbs.Execute(fqry01)
         
         If Not Ftb01.EOF Then
            Lblauxcon.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         fqry01 = "SELECT dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtsitio.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'SITIOS INFRACCION' "
               
         Set Ftb01 = Dbs.Execute(fqry01)
         
         If Not Ftb01.EOF Then
            Lblsitio.Caption = Ftb01!dsdes
            Ftb01.Close
         End If
      
         
         Txtauxalt_LostFocus
               
         If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            Mskvalmul.text = (Msknetpag.text * 100) / (100 - Txtdescuento.text)
         End If
        
            'NUEVA LIQUIDACION DE DESCUENTOS POR DIAS    ahh 04/03/2002
         'If Txtcodinf.Text = "13" Or Txtcodinf.Text = "9" Or Txtcodinf.Text = "16" Then
         If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            MskSaldo.text = saldo
         
         Else
              If Txtdescuento.text = 0 Then
                 'liquido dias
                  dias = DateDiff("d", fecmul, hoy)
                 'busco parametro de acuerdo a los dias
                  fqry01 = "SELECT psval " _
                           & "FROM geparsis " _
                          & "WHERE pscod = 'DIASMUL1' "
                  Set Ftb01 = Dbs.Execute(fqry01)
        
                 If Not Ftb01.EOF Then
                    dmul1 = Ftb01!psval
                 End If
                 Ftb01.Close
                 
                 fqry01 = "SELECT psval " _
                          & "FROM geparsis " _
                         & "WHERE pscod = 'DIASMUL2' "
                 Set Ftb01 = Dbs.Execute(fqry01)
           
                 If Not Ftb01.EOF Then
                    dmul2 = Ftb01!psval
                 End If
                 Ftb01.Close
                 
                 fqry01 = "SELECT psval " _
                          & "FROM geparsis " _
                         & "WHERE pscod = 'DIASMUL3' "
                 Set Ftb01 = Dbs.Execute(fqry01)
           
                 If Not Ftb01.EOF Then
                    dmul3 = Ftb01!psval
                 End If
                 Ftb01.Close
                 
                 'Busco % de acuerdo a los dias
                 If dias <= Val(dmul1) Then
                    fqry02 = "SELECT psval " _
                             & "FROM geparsis " _
                            & "WHERE pscod = 'PORMUL1' "
                    Set Ftb02 = Dbs.Execute(fqry02)
           
                    If Not Ftb02.EOF Then
                       pmul = Ftb02!psval
                    End If
                    Ftb02.Close
                 End If
                 If dias > Val(dmul1) And dias <= Val(dmul2) Then
                    fqry02 = "SELECT psval " _
                             & "FROM geparsis " _
                            & "WHERE pscod = 'PORMUL2' "
                    Set Ftb02 = Dbs.Execute(fqry02)
           
                    If Not Ftb02.EOF Then
                       pmul = Ftb02!psval
                    End If
                    Ftb02.Close
                 End If
                 If dias >= Val(dmul3) Then
                    fqry02 = "SELECT psval " _
                             & "FROM geparsis " _
                            & "WHERE pscod = 'PORMUL3' "
                    Set Ftb02 = Dbs.Execute(fqry02)
           
                    If Not Ftb02.EOF Then
                       pmul = Ftb02!psval
                    End If
                    Ftb02.Close
                 End If
                 'liquido descuento de acuerdo al %
                 Txtdescuento.text = pmul
                 If pmul <> 0 Then
                    Msknetpag.text = Val(Mskvalmul.text) - (Val(Mskvalmul.text) * (Val(pmul) / 100))
                 Else
                    Msknetpag.text = Val(Mskvalmul.text)
                 End If
                 fqry02 = "SELECT SUM(abvalor) " _
                          & "FROM opabonos " _
                         & "WHERE abnumero = " & Txtnumero.text & " "
                         
                 Set Ftb02 = Dbs.Execute(fqry02)
                 
                 If Not Ftb02.EOF Then
                    If IsNull(Ftb02(0)) Then
                       valabonado = 0
                    Else
                       valabonado = Ftb02(0)
                    End If
                 Else
                    valabonado = 0
                 End If
                 Ftb02.Close
                 
                 MskSaldo.text = Msknetpag.text - valabonado
              Else
                 MskSaldo.text = saldo
          End If
        End If
         'FIN LIQUIDACION
         
         If Msknetpag.text <> "0" Then
            numero = str(Msknetpag.text)
            digitos = Len(numero)
            inicio1 = digitos - 1
            valor = Mid(numero, inicio1, 2)
            If valor <> "00" Then
               complemento = 50 - valor
               valtot = Abs(complemento) + numero
               Msknetpag.text = valtot
               MskSaldo.text = valtot
            Else
               Msknetpag.text = Mskvalmul.text
            End If
         End If
         
         If sEstadoMulta = "Pendiente" Or sEstadoMulta = "Abonada" Then
            Btngrabar.Enabled = True
            Btngrabar.SetFocus
         ElseIf sEstadoMulta = "Anulado" Then
            Btngrabar.Enabled = False
            Btnimprimir.Enabled = False
         ElseIf sEstadoMulta = "Pagada" Then
            Mskfecsis.text = Mskfecgra.text
            Btngrabar.Enabled = False
            If sNombreCaja = Pc Then
               Btnimprimir.Enabled = True
            Else
               Btnimprimir.Enabled = False
            End If
         
         End If
          
      Else
         MsgBox "Numero De Evasion No Existe", 48, "Advertencia"
         Txtnumero.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub TxtPlaca_DblClick()
  forma = "7"
  forma1 = "2"
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
'  Char = Chr(KeyAscii%)
'  KeyAscii = Asc(UCase(Char))  ' mayuscula
'  longitud = Len(TxtPlaca.Text)
'  longitud = longitud + 1
'  If longitud <= 2 Then
'     If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'         Exit Sub
'     Else
'         'MsgBox "En El Numero De Placa las Tres Primeras Deben Ser Letras", 48, "Advertencia"
'         KeyAscii% = 0
'         Exit Sub
'     End If
'  ElseIf longitud > 3 Then
'     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
'        Exit Sub
'     Else
'        'MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
'        KeyAscii% = 0
'        Exit Sub
'     End If
'  End If
'  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
'      Exit Sub
'  Else
'     'KeyAscii% = 0
'  End If
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If

End Sub

Private Sub TxtPlaca_LostFocus()
   If txtPlaca.text <> "" Then
      
      longitud = Len(txtPlaca.text)
      
      'Sebastián Rondón - 29 de Mayo de 2023
      'Se consulta parametro para definir si el sistema permite cobrar Duplicado de documentos
      QRY = "SELECT psval FROM geparsis WHERE pscod = 'COBRODUPLI' "
      Set Tbl1 = Dbs.Execute(QRY)
   
      If longitud < 6 Then
         MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
         txtPlaca.SetFocus
         Exit Sub
      End If
        
      fqry01 = "SELECT * " _
               & "FROM opmultas " _
              & "WHERE "
              
      If G_Numero_Comparendo_Antiguo <> 0 Then
         fqry01 = fqry01 & " munumero = " & G_Numero_Comparendo_Antiguo & "  AND "
      End If
      
      If Tbl1!psval = "N" Then
        fqry01 = fqry01 & " mucodinf NOT IN (53) AND "
      End If
      
      fqry01 = fqry01 & " munumpla = '" & txtPlaca.text & "' " _
                      & " AND (muestado = 'Pendiente' or muestado = 'Abonada') order by mufecha "
                      
      Tbl1.Close
      
      Set Ftb01 = Dbs.Execute(fqry01)
         
      If Not Ftb01.EOF Then
         Txtnumero.Enabled = False
         txtPlaca.Enabled = False
         Txtnumero.text = Ftb01!munumero
         Txtcodemp.text = Ftb01!mucodemp
         Txtcodinf.text = Ftb01!mucodinf
         Txtnomnot.text = Ftb01!munomnot
         Mskfecha.text = Ftb01!mufecha
         Mskhora.text = Ftb01!muhora
         Txtsitio.text = Ftb01!musitio
         Txtnomtes.text = Ftb01!munomtes
         Txtnomaux.text = Ftb01!muauxcon
         Txttipveh.text = Ftb01!mutipveh
         Txtdescuento.text = Ftb01!mudescue
         TxtObservaciones.text = Ftb01!muobserv
         Msknetpag.text = Ftb01!munetpag
         fecmul = Ftb01!mufecsis    ' ahh 04/03/2002
         Mskfecgra.text = Ftb01!mufecsis   'ahh   04/04/2002
         saldo = Ftb01!musaldo   'wecc 08/06/2002
         If Ftb01!murecpag <> "" Then
            
            If Ftb01!murecpag <> 0 Then
               Txtrecpag.text = Ftb01!murecpag
               Txtrecpag_LostFocus
            End If
         End If
                                    
         If Ftb01!muestado = "Abonada" Then
            OptEstado(3).Enabled = True
            OptEstado(3).Value = True
            OptEstado(1).Enabled = False
         ElseIf Ftb01!muestado = "Pendiente" Then
            OptEstado(0).Enabled = True
            OptEstado(0).Value = True
            OptEstado(1).Enabled = False
         End If
         Ftb01.Close
                                    
         fqry01 = "SELECT * " _
                  & "FROM ophojveh " _
                 & "WHERE hvplaca = '" & txtPlaca.text & "'"
              
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            TxtNumInt.text = Ftb01!hvnumint
         End If
         Ftb01.Close
                  
         fqry02 = "SELECT * " _
                     & "FROM geparsis " _
                    & "WHERE pscod = '" & Txtcodinf.text & "' "
         
         Set Ftb02 = Dbs.Execute(fqry02)
      
         If Not Ftb02.EOF Then
            LblCodInf.Caption = Ftb02!psdes
            
            If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
               MskSaldo.text = saldo
            Else
            
'               fqry01 = "SELECT psval " _
'                        & "FROM geparsis " _
'                       & "WHERE pscod = 'SALMIN' "
'
'               Set Ftb01 = Dbs.Execute(fqry01)
'
'               If Not Ftb01.EOF Then
'                  Mskvalmul.Text = CLng((Ftb01!psval / 30) * Ftb02!psval)
'                  numero = Str(Mskvalmul.Text)
'                  digitos = Len(numero)
'                  inicio1 = digitos - 1
'                  valor = Mid(numero, inicio1, 2)
'                  If valor <> "00" Then
'                     complemento = 100 - valor
'                     valtot = Abs(complemento) + numero
'                     Mskvalmul.Text = valtot
'                  End If
'               End If
            End If
         End If
         
         fqry01 = "SELECT * " _
                  & "FROM ophojemp " _
                 & "WHERE hecodigo = '" & Txtcodemp.text & "'"
              
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lblcodemp.Caption = Ftb01!henombre
            Ftb01.Close
         Else
            Ftb01.Close
            fqry01 = "SELECT henombre " _
                    & "FROM ophojempcon " _
                   & "WHERE hecodigo = '" & Txtcodemp.text & "'"
        
            Set Ftb01 = Dbs.Execute(fqry01)
   
            If Not Ftb01.EOF Then
               Lblcodemp.Caption = Ftb01!henombre
            End If
            Ftb01.Close
         End If
         
         
         fqry01 = "SELECT dsdes " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txttipveh.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'TIPOS DE VEHICULO' "
                
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lbltipveh.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         
         fqry01 = "SELECT * " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txtnomaux.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'AUXILIARES DE CONTROL' "
                
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lblauxcon.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         
         If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            Mskvalmul.text = (Msknetpag.text * 100) / (100 - Txtdescuento.text)
         End If
         
         Btngrabar.Enabled = True
      Else
         MsgBox "Codigo De Placa No Tiene Ningun Registro De Evasiones y Elusiones", 48, "Advertencia"
         'txtPlaca.SetFocus
         Exit Sub
      End If
     'NUEVA LIQUIDACION DE DESCUENTOS POR DIAS    ahh 04/03/2002
        'liquido dias
      
      If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            MskSaldo.text = saldo
         
      Else
         If Txtdescuento.text = 0 Then
            dias = DateDiff("d", fecmul, hoy)
           'busco parametro de acuerdo a los dias
            fqry01 = "SELECT psval " _
                     & "FROM geparsis " _
                    & "WHERE pscod = 'DIASMUL1' "
                    
            Set Ftb01 = Dbs.Execute(fqry01)
      
            If Not Ftb01.EOF Then
               dmul1 = Ftb01!psval
            End If
            Ftb01.Close
            fqry01 = "SELECT psval " _
                     & "FROM geparsis " _
                    & "WHERE pscod = 'DIASMUL2' "
                    
            Set Ftb01 = Dbs.Execute(fqry01)
      
            If Not Ftb01.EOF Then
               dmul2 = Ftb01!psval
            End If
            Ftb01.Close
            fqry01 = "SELECT psval " _
                     & "FROM geparsis " _
                    & "WHERE pscod = 'DIASMUL3' "
                    
            Set Ftb01 = Dbs.Execute(fqry01)
      
            If Not Ftb01.EOF Then
               dmul3 = Ftb01!psval
            End If
            Ftb01.Close
            fqry01 = "SELECT psval " _
                     & "FROM geparsis " _
                    & "WHERE pscod = 'DIASMUL4' "
                    
            Set Ftb01 = Dbs.Execute(fqry01)
      
            If Not Ftb01.EOF Then
               dmul4 = Ftb01!psval
            End If
            Ftb01.Close
            'Busco % de acuerdo a los dias
            If dias <= Val(dmul1) Then
               fqry02 = "SELECT psval " _
                        & "FROM geparsis " _
                       & "WHERE pscod = 'PORMUL1' "
                       
               Set Ftb02 = Dbs.Execute(fqry02)
      
               If Not Ftb02.EOF Then
                  pmul = Ftb02!psval
               End If
               Ftb02.Close
            End If
            If dias > Val(dmul1) And dias <= Val(dmul2) Then
               fqry02 = "SELECT psval " _
                        & "FROM geparsis " _
                       & "WHERE pscod = 'PORMUL2' "
                       
               Set Ftb02 = Dbs.Execute(fqry02)
      
               If Not Ftb02.EOF Then
                  pmul = Ftb02!psval
               End If
               Ftb02.Close
            End If
            If dias > Val(dmul2) And dias <= Val(dmul3) Then
               fqry02 = "SELECT psval " _
                        & "FROM geparsis " _
                       & "WHERE pscod = 'PORMUL3' "
                       
               Set Ftb02 = Dbs.Execute(fqry02)
      
               If Not Ftb02.EOF Then
                  pmul = Ftb02!psval
               End If
               Ftb02.Close
            End If
            If dias >= Val(dmul4) Then
               fqry02 = "SELECT psval " _
                        & "FROM geparsis " _
                       & "WHERE pscod = 'PORMUL4' "
                       
               Set Ftb02 = Dbs.Execute(fqry02)
      
               If Not Ftb02.EOF Then
                  pmul = Ftb02!psval
               End If
               Ftb02.Close
            End If
           'liquido descuento de acuerdo al %
            Txtdescuento.text = pmul
            If pmul <> 0 Then
               Msknetpag.text = Val(Mskvalmul.text) - (Val(Mskvalmul.text) * (Val(pmul) / 100))
            Else
               Msknetpag.text = Val(Mskvalmul.text)
            End If
            fqry02 = "SELECT SUM(abvalor) " _
                     & "FROM opabonos " _
                    & "WHERE abnumero = " & Txtnumero.text & " "
                   
            Set Ftb02 = Dbs.Execute(fqry02)
   
            If Not Ftb02.EOF Then
               If IsNull(Ftb02(0)) Then
                  valabonado = 0
               Else
                  valabonado = Ftb02(0)
               End If
            Else
               valabonado = 0
            End If
            Ftb02.Close
            
            MskSaldo.text = Msknetpag.text - valabonado
         Else
            MskSaldo.text = saldo
           'FIN LIQUIDACION
         End If
      End If
   End If
End Sub

Private Sub Txtrecpag_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtrecpag_LostFocus()
On Error GoTo err_Txtrecpag_LostFocus
   
   If Trim(Txtrecpag.text) <> "" Then
   
      QRY = "SELECT * " _
          & "FROM opabonos " _
         & "WHERE abrecibo = " & Trim(Txtrecpag.text) & ""
      
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         QRY = "SELECT MAX(abrecibo) " _
               & "FROM opabonos "
          
         Set Tbl2 = Dbs.Execute(QRY)
         
         If Not Tbl2.EOF Then
            Txtrecpag.text = Tbl2(0) + 1
         End If
         Tbl2.Close
      End If
      Tbl1.Close
      
   End If
   Exit Sub
   
err_Txtrecpag_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub
