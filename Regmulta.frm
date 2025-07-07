VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Regmultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura y Negociación de Evasiones y Elusiones"
   ClientHeight    =   7515
   ClientLeft      =   1260
   ClientTop       =   1110
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   6240
      TabIndex        =   82
      Top             =   120
      Width           =   1935
      Begin VB.TextBox TxtNumConvenio 
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
         MaxLength       =   10
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de Transporte (presione DblClick para ayuda)"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Núm. Convenio:"
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
         Index           =   24
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnconductor 
      Caption         =   "&Cedula"
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
      Picture         =   "Regmulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Mostrar Pantalla de Conductores"
      Top             =   120
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
      Height          =   5775
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   6015
      Begin VB.TextBox TxtCedula 
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
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   14
         ToolTipText     =   "Número cédula del conductor"
         Top             =   4440
         Width           =   1335
      End
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
         Height          =   2415
         Left            =   3000
         TabIndex        =   66
         Top             =   1920
         Width           =   2775
         Begin VB.TextBox TxtPasPla 
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   13
            ToolTipText     =   "Cantidad de pasajeros en planilla"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Txttiquete 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   12
            ToolTipText     =   "Número incial por el que empiece a buscar"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Txtplanilla 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "Número incial por el que empiece a buscar"
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
            TabIndex        =   10
            ToolTipText     =   "Código de la ruta (presione DblClick para Ayuda)"
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Mskvalpla 
            Height          =   285
            Left            =   1320
            TabIndex        =   71
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
         Begin VB.Label Label1 
            Caption         =   "Pasajeros:"
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
            Index           =   7
            Left            =   120
            TabIndex        =   79
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label22 
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   70
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
            TabIndex        =   69
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
         Height          =   2415
         Left            =   120
         TabIndex        =   61
         Top             =   1920
         Width           =   2775
         Begin VB.TextBox TxtPasVia 
            Height          =   285
            Left            =   960
            MaxLength       =   2
            TabIndex        =   9
            ToolTipText     =   "Cantidad de pasajeros en viaje"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Txtnumcon 
            Height          =   285
            Left            =   960
            TabIndex        =   7
            ToolTipText     =   "Número de la tasa de uso"
            Top             =   1320
            Width           =   1695
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
            TabIndex        =   6
            ToolTipText     =   "Código de la ruta (presione DblClick para Ayuda)"
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Mskvalor 
            Height          =   285
            Left            =   960
            TabIndex        =   64
            TabStop         =   0   'False
            ToolTipText     =   "Valor conduce"
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
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
            Left            =   960
            TabIndex        =   8
            ToolTipText     =   "Fecha hasta donde quiere que busque"
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   16
            Format          =   "mmm/dd/yyyy HH:mm"
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Pasajeros:"
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
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label20 
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   65
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
            TabIndex        =   63
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
            TabIndex        =   62
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.TextBox Txtauxalt 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   18
         ToolTipText     =   "Código del auxiliar de control que hizo el parte (doble Click)"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Txtcodemp 
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   2
         ToolTipText     =   "Código de la empresa con la que esta asociada el vehículo (doble Click para ayuda)"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtplaca 
         Height          =   285
         Left            =   120
         MaxLength       =   7
         TabIndex        =   0
         ToolTipText     =   "El número de la placa del vehículo a quien se le hace el parte (doble Click para ayuda)"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtcodinf 
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   5
         ToolTipText     =   "Código de la infracción (doble Click para ayuda)"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Txttipveh 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   3
         ToolTipText     =   "Tipo de vehículo(doble Click para ayuda)"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtsitio 
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Sitio de la infracción"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtnomaux 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         ToolTipText     =   "Código del auxiliar de control que hizo el parte (doble Click)"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Txtnomtes 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         ToolTipText     =   "Nombre de un Testigo"
         Top             =   4800
         Width           =   4455
      End
      Begin VB.TextBox Txtnomnot 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Nombre del conductor activo "
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox Txtnumint 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Número interno del vehículo a quien se le hace el parte"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula"
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
         Index           =   15
         Left            =   120
         TabIndex        =   80
         Top             =   4440
         Width           =   1335
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
         Left            =   3000
         TabIndex        =   77
         Top             =   1560
         Width           =   2895
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
         TabIndex        =   76
         Top             =   5400
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
         TabIndex        =   75
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   60
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label19 
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
         TabIndex        =   59
         Top             =   5160
         Width           =   1095
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
         TabIndex        =   57
         Top             =   240
         Width           =   975
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   1560
         Width           =   1455
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Control 1 :"
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
         TabIndex        =   50
         Top             =   5160
         Width           =   1095
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
         Left            =   120
         TabIndex        =   49
         Top             =   4800
         Width           =   1455
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
      Height          =   3015
      Left            =   6240
      TabIndex        =   42
      Top             =   1200
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
         Height          =   300
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Número de la multa (d oble Click)"
         Top             =   480
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Mskfecsis 
         Height          =   300
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         Format          =   "mmm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   300
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Fecha en que se hizo el parte"
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskhora 
         Height          =   300
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Hora en que se hizo el parte"
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         MaxLength       =   5
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
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
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
         TabIndex        =   47
         Top             =   1560
         Width           =   855
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   2280
         Width           =   855
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
         TabIndex        =   44
         Top             =   960
         Width           =   1455
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
      Height          =   1335
      Left            =   2400
      TabIndex        =   41
      Top             =   6000
      Width           =   3735
      Begin VB.TextBox Txtobservaciones 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         ToolTipText     =   "Observaciones que tenga de la multa"
         Top             =   240
         Width           =   3495
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
      Caption         =   "Estado de la Evasion"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   40
      Top             =   6000
      Width           =   2175
      Begin VB.OptionButton Optestado 
         Caption         =   "Pendiente"
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
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Optestado 
         Caption         =   "Pagada"
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
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1335
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
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Datos de la Evasion"
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
      Height          =   3015
      Left            =   6240
      TabIndex        =   32
      Top             =   4320
      Width           =   1935
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
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Valor de descuento no debe exceder del 50%"
         Top             =   1800
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Mskvalmul 
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
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
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msknetpag 
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2400
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
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
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
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "% Descuento:"
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
         TabIndex        =   38
         Top             =   1560
         Width           =   1335
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
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1575
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
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Lblrecpag 
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
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CommandButton Btngrabar 
      Caption         =   "&Grabar"
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
      Picture         =   "Regmulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Ingresa la empresa a la base de datos"
      Top             =   3240
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
      Picture         =   "Regmulta.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   4080
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
      Picture         =   "Regmulta.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   4920
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
      Picture         =   "Regmulta.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Limpia la forma actual"
      Top             =   5760
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
      Left            =   8400
      Picture         =   "Regmulta.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Salir de la forma actual"
      Top             =   6600
      Width           =   855
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   8640
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   8280
      X2              =   8280
      Y1              =   0
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   9360
      Y1              =   7440
      Y2              =   7440
   End
End
Attribute VB_Name = "Regmultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numero
Dim valtiq
Dim hora
Dim valdes
Dim bandera As Integer
Dim anosistemas As String
Dim L_Cedula_Conductor As String
Dim L_Valor_Excreta As Double
Dim L_Aviso_Multa As Double
Dim L_CodigoEmpresaConvenio As String
Dim L_CodigoEmpresaConvenioEmpresarial As String
Dim L_Bandera_Multas As String   'variabe q identifica si la multas tiene el valro en gedetsuptip en 0
Dim L_CodAuxCtr As String        'variable codigo de control integra
'Dim NumUni As Long
'Dim dbsca() As Database
'Dim Exi_Multa() As String

Private Function ValidaEmpresaRuta(ByVal fdestino As String, ByVal fempresa As String) As Boolean
   ValidaEmpresaRuta = False
   If G_Validar_Ser = "N" Then
      QRY = "SELECT * " _
            & "FROM opruthor " _
           & "WHERE rhcodemp = '" & fempresa & "' " _
             & "AND rhcodrut = '" & fdestino & "' "
      
      Set Ftb04 = Dbs.Execute(QRY)
      
      If Ftb04.EOF Then
         Ftb04.Close
'         If Len(fempresa) = 4 Then
'            qry = "SELECT ceempcon " _
'                  & "FROM opdetconemp, opconemp " _
'                 & "WHERE ceempcon = '" & fempresa & "' " _
'                   & "AND ceestado = 'A' " _
'                   & "AND cenumero = cdnumero " _
'                   & "AND cdcodrut = '" & fdestino & "'"
'
'            Set Ftb04 = Dbs.Execute(qry)
'
'            If Ftb04.EOF Then
'               Ftb04.Close
'               Exit Function
'            End If
'         End If
         If Trim(L_CodigoEmpresaConvenio) <> "" Then
            QRY = "SELECT cenumero,cetipo, cefecvig, cecodemp, cefecini " _
                & "FROM opconemp " _
                & "WHERE ceempcon ='" & fempresa & "' " _
                & " AND cecodemp = '" & L_CodigoEmpresaConvenio & "' " _
                & "AND ceestado = 'A' "
        
            L_Codigo_Empresa = txtCodEmp.text
        
            Set Ftb02 = Dbs.Execute(QRY)
            
            Do While Not Ftb02.EOF
               sTipo = Ftb02!cetipo
          
               QRY = "SELECT * " _
                    & "FROM opdetconemp " _
                    & "WHERE cdnumero = " & Ftb02!cenumero & " " _
                    & "AND cdcodrut = '" & fdestino & "' "
                    
               Set Ftb03 = Dbs.Execute(QRY)
               If Not Ftb03.EOF Then
                  ValidaEmpresaRuta = True
               End If
               Ftb03.Close
               Ftb02.MoveNext
            Loop
            Ftb02.Close
         Else
            If Trim(TxtNumConvenio.text) <> "" Then
               QRY = "SELECT * " _
                     & "FROM opconemp " _
                     & "WHERE cenumero ='" & TxtNumConvenio.text & "' " _
                     & " AND ceempcon = '" & fempresa & "' "
                     
               Set Ftb02 = Dbs.Execute(QRY)
               If Not Ftb02.EOF Then
                  If Ftb02!ceestado = "I" Then
                     Ftb02.Close
                     Exit Function
                  End If
                  
                  QRY = "SELECT cdtipveh, cdestado " _
                        & "FROM opdetconemp " _
                        & "WHERE cdnumero = " & TxtNumConvenio.text & " " _
                        & "AND cdcodrut = '" & fdestino & "' " _
                  
                  Set Ftb03 = Dbs.Execute(QRY)
                  sBanderaTipoServicio = False
                  Do While Not Ftb03.EOF
                     
                     If Trim(Ftb03!cdtipveh) = Trim(Txttipveh.text) Then
                     
                        If Ftb03!cdestado = "I" Then
                           Ftb03.Close
                           Exit Function
                        End If
                        sBanderaTipoServicio = True
                        Exit Do
                     End If
                     Ftb03.MoveNext
                  Loop
                  Ftb03.Close
                  
                  If sBanderaTipoServicio = False Then
                     Ftb02.Close
                     Exit Function
                  End If
                  ValidaEmpresaRuta = True
               Else
                  Ftb02.Close
                  Exit Function
               End If
               Ftb02.Close
               Exit Function
            End If
         End If
      Else
         ValidaEmpresaRuta = True
         Exit Function
      End If
    Else
      'Ftb04.Close
      ValidaEmpresaRuta = True
    End If

End Function

Private Sub BtnBuscar_Click()
  forma = "7"
  forma1 = "1"
  Txtnumero.text = ""
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
End Sub

Private Sub btnconductor_Click()
   On Error GoTo Errores
   
   If Trim(Txtnumero.text) = "" Then
      MsgBox "Falta Ingresar El Numero De La Evasion", 48, "Advertencia"
      Txtnumero.SetFocus
      Exit Sub
   End If
   
   If txtCedula.text = "" Then
      MsgBox "Falta ingresar el número de cedula del conductor", 48, "Advertencia"
      txtCedula.SetFocus
      Exit Sub
   End If
   If Txtnomnot.text = "" Then
      MsgBox "Falta ingresar el Nombre Del Notificado", 48, "Advertencia"
      txtCedula.SetFocus
      Exit Sub
   End If
  
   If Trim(L_Cedula_Conductor) = Trim(txtCedula.text) Then
      MsgBox "El número de cedula del conductor no ha cambiado", 48, "Advertencia"
      txtCedula.SetFocus
      Exit Sub
   End If
   respuesta = MsgBox("Desea actualizar el número de cedula y nombre del infractor", 32 + vbYesNo, "GRABAR")
   
   If respuesta = vbNo Then
      Exit Sub
   End If
   
   Dbs.BeginTrans
   On Error GoTo Err_Transac
   
   QRY = "UPDATE opmultas " _
          & "SET mucedula = '" & txtCedula.text & "', " _
             & " munomnot = '" & Txtnomnot.text & "' " _
        & "WHERE munumero = " & Txtnumero.text & " "
      
      Dbs.Execute QRY
   
   sFecha = Format(Fn_Hora_Servidor, "MM/dd/yyyy hh:mm")
   sObservacion = "CAMBIO DE CEDULA, EVASION " & Txtnumero.text
   QRY = " INSERT INTO gelogaud " _
             & "(lofecact, lofecnue, lofuncio, " _
             & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
             & "VALUES ( to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),  to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),'" _
             & Login & "', to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),'" & sObservacion & "', 'MUL', '" _
             & L_Cedula_Conductor & "', '" & txtCedula.text & "')"
        
   Dbs.Execute QRY
      
   Dbs.CommitTrans
   On Error GoTo Errores
   MsgBox "El conductor se actualizó correctamente", 64, "O.K."
   
   btnconductor.Enabled = False
   
   Exit Sub
Err_Transac:
   Dbs.RollbackTrans
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub

   
End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
 
   If txtPlaca.text = "" Then
      MsgBox "Falta ingresar el Codigo De La Placa", 48, "Advertencia"
      txtPlaca.SetFocus
      Exit Sub
   End If
  
   If TxtNumInt.Enabled = True And TxtNumInt.text = "" Then
      MsgBox "Falta ingresar el Numero Interior Del Vehiculo", 48, "Advertencia"
      TxtNumInt.SetFocus
      Exit Sub
   End If
  
   If txtCodEmp.text = "" Then
      MsgBox "Falta ingresar el Codigo De La Empresa", 48, "Advertencia"
      txtCodEmp.SetFocus
      Exit Sub
   End If
  
   If Txttipveh.text = "" Then
      MsgBox "Falta ingresar el Tipo De Vehiculo", 48, "Advertencia"
      Txttipveh.SetFocus
      Exit Sub
   End If
   If Txtsitio.text = "" Then
      MsgBox "Falta ingresar el Sitio", 48, "Advertencia"
      Txtsitio.SetFocus
      Exit Sub
   End If
  
   If Txtcodinf.text = "" Then
      MsgBox "Falta ingresar el Codigo De La Infraccion", 48, "Advertencia"
      Txtcodinf.SetFocus
      Exit Sub
   End If
   
   fqry01 = "SELECT dsdes " _
            & "FROM gesuptip, gedetsuptip " _
           & "WHERE dscoddet = '" & Txtcodinf.text & "' " _
             & "AND dscodtip = stcodtip " _
             & "AND stdes = 'INFRACCIONES PARA COBRAR' AND dsval = 1 "
               
   Set Ftb01 = Dbs.Execute(fqry01)
     
   If Not Ftb01.EOF Then
   'If Trim(Txtcodinf.Text) = "51" Or Trim(Txtcodinf.Text) = "52" Then
   'If Trim(Txtcodinf.Text) = "51" Then 'No se tiene en cuenta el 52 porque es una evacion segun requerimiento 1075 con soporte 1481 del 16 de Oct de 2007
      If Txtdestino.text = "" Then
         MsgBox "Falta ingresar el DESTINO", 48, "Advertencia"
         Txtdestino.SetFocus
         Exit Sub
      End If
      
      If Trim(Txtcodinf.text) = "51" Then
        If Txtnumcon.text = "" Then
           MsgBox "Falta ingresar el NUMERO DE TASA DE USO", 48, "Advertencia"
           Txtnumcon.SetFocus
           Exit Sub
        End If
      Else
         Txtnumcon.text = "0"
      End If
      If TxtPasVia.text = "" Then
         MsgBox "Falta ingresar la cantidad de pasajeros en viaje", 48, "Advertencia"
         TxtPasVia.SetFocus
         Exit Sub
      End If
      If Txtdespla.text = "" Then
         MsgBox "Falta ingresar el DESTINO DE PLANILLA", 48, "Advertencia"
         Txtdespla.SetFocus
         Exit Sub
      End If
      If Txtplanilla.text = "" Then
         MsgBox "Falta ingresar el NUMERO DE PLANILLA", 48, "Advertencia"
         Txtplanilla.SetFocus
         Exit Sub
      End If
      If Txttiquete.text = "" Then
         MsgBox "Falta ingresar el NUMERO DE TIQUETE", 48, "Advertencia"
         Txttiquete.SetFocus
         Exit Sub
      End If
      If TxtPasPla.text = "" Then
         MsgBox "Falta ingresar la cantidad de pasajeros en planilla", 48, "Advertencia"
         TxtPasPla.SetFocus
         Exit Sub
      End If
      sDestino = Trim(Txtdestino.text)
      sValCon = Mskvalor.text
      sConduce = Txtnumcon.text
      sDesPla = Trim(Txtdespla.text)
      sValPla = Mskvalpla.text
      sPlanilla = Txtplanilla.text
      sTiquete = Txttiquete.text
      sPasVia = TxtPasVia.text
      sPasPla = TxtPasPla.text
   Else
      sDestino = ""
      sValCon = 0
      sConduce = 0
      sDesPla = ""
      sValPla = 0
      sPlanilla = ""
      sTiquete = ""
      sPasVia = 0
      sPasPla = 0
   End If
   If Trim(Txtcodinf.text) = "19" Or Trim(Txtcodinf.text) = "20" Then
      If TxtPasVia.text = "" Then
         MsgBox "Falta ingresar la cantidad de pasajeros en viaje", 48, "Advertencia"
         TxtPasVia.SetFocus
         Exit Sub
      End If
      If TxtPasPla.text = "" Then
         MsgBox "Falta ingresar la cantidad de pasajeros en planilla", 48, "Advertencia"
         TxtPasPla.SetFocus
         Exit Sub
      End If
      sPasVia = TxtPasVia.text
      sPasPla = TxtPasPla.text
   End If
   If Trim(txtCedula.text) = "" Then
      txtCedula.text = "0"
   End If
   If Trim(Txtnomnot.text) = "" Then
      Txtnomnot.text = "PENDIENTE"
   End If
  
   If Txtnomtes.text = "" Then
      MsgBox "Falta ingresar el Nombre Del Testigo", 48, "Advertencia"
      Txtnomtes.SetFocus
      Exit Sub
   End If
  
   If Txtnomaux.text = "" Then
      MsgBox "Falta Ingresar  El Codigo Del Auxiliar De Control", 48, "Advertencia"
      Txtnomaux.SetFocus
      Exit Sub
   End If
   
   If TxtObservaciones.text = "" Then
      MsgBox "Falta Ingresar Las Observaciones", 48, "Advertencia"
      TxtObservaciones.SetFocus
      Exit Sub
   End If
   
   If Txtnumero.text = "" Then
      MsgBox "Falta ingresar el Numero De La Evasion", 48, "Advertencia"
      Txtnumero.SetFocus
      Exit Sub
   End If
   
   If Mskfecha.ClipText = "" Then
      MsgBox "Falta Ingresar La Fecha", 48, "Advertencia"
      Mskfecha.SetFocus
      Exit Sub
   End If
   
   If Mskhora.ClipText = "" Then
      MsgBox "Falta Ingresar La Hora", 48, "Advertencia"
      Mskhora.SetFocus
      Exit Sub
   End If
   
   If Txtdescuento.text = "" Then
      MsgBox "Falta ingresar el Descuento", 48, "Advertencia"
      Txtdescuento.SetFocus
      Exit Sub
   End If
   
    
   respuesta = MsgBox("Desea Grabar La Evasion", 32 + vbYesNo, "GRABAR")
   
   If respuesta = vbNo Then
      Exit Sub
   End If
  
   If bandera = 1 Then
      Dbs.BeginTrans
      On Error GoTo transac
     
      QRY = "UPDATE opmultas " _
             & "SET mudescue = " & Txtdescuento.text & " ," _
             & "    munetpag = " & Msknetpag.text & " ," _
             & "    muobserv = '" & TxtObservaciones.text & "' ," _
             & "    muvaldes = " & valdes & ", " _
             & "    munomcaj = '" & Pc & "', " _
             & "    musaldo = " & Msknetpag.text & " " _
           & "WHERE munumero = " & Txtnumero.text & " "
      
      Dbs.Execute QRY
      sFecha = Format(Now, "MM/dd/yyyy hh:mm")
      QRY = " INSERT INTO gelogaud " _
                & "(lofecact, lofecnue, lofuncio, " _
                & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
                & "VALUES ( to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),  to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'" _
                & Login & "', to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'SE NEGOCIO UNA EVASION', 'MUL', '" _
                & Txtnumero.text & "', '" & Txtdescuento.text & "')"
        
      Dbs.Execute QRY
        
      Dbs.CommitTrans
      On Error GoTo Errores
  
   Else
      fqry01 = "SELECT * " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & txtPlaca.text & "'"
  
      Set Ftb01 = Dbs.Execute(fqry01)
  
      If Txtdescuento.text = "" Then
         Txtdescuento.text = "0"
      End If
      
      'Sebastian Rondón - Octubre 24 de 2023
      'Antes de insertar, actualizamos el consecutivo de la multas nuevamente
      fqry02 = "SELECT psval " _
                & "FROM geparsis " _
               & "WHERE pscod = 'NUMMUL' "
     
    
       Set Ftb02 = Dbs.Execute(fqry02)
      
       If Not Ftb02.EOF Then
          Txtnumero.text = Ftb02!psval + 1
       End If
       
      Dbs.BeginTrans
      On Error GoTo transac
      If Ftb01.EOF Then
         Dbs.Execute "INSERT INTO ophojveh " _
                   & "(hvplaca, hvnumint, hvcodemp, hvpropie, " _
                   & " hvnumpas, hvestado, hvconact, hvtipveh, " _
                   & " hvfecultpru, hvprueba) VALUES ('" _
                   & txtPlaca.text & "','" & TxtNumInt.text & "','" & txtCodEmp.text & "','" _
                   & Txtnomnot.text & "',45,   'Activo','" & Txtnomnot.text & "','" & Txttipveh.text & "',to_date ('" _
                   & Date & "','MM/DD/YYYY'),'No')"
      End If
     
      If Trim(Txtauxalt.text) = "" Then
         sAuxAlt = ""
      Else
         sAuxAlt = Trim(Txtauxalt.text)
      End If
     
      hora = Format(Mskhora.text, "hh:mm")
                  
      sAvimulta = 0
                  
      If L_Bandera_Multas = "S" Then
         sAvimulta = L_Aviso_Multa
      End If
                  
      QRY = "INSERT INTO opmultas " _
                & "(munumero, mucodemp, mucodinf, munomnot, " _
                & " mufecha, muhora, munumpla, musitio, " _
                & " munomtes, muauxcon, mutipveh, mudescue, " _
                & " muobserv, muestado, munetpag, muvaldes, " _
                & " mufecsis, musaldo, munomcaj, muauxalt, " _
                & " mudestino, muvalcon, muconduce, mudespla, " _
                & " muvalpla, muplanilla, mutiquete, mupasvia, " _
                & " mupaspla, mucedula, muaviso) VALUES (" _
                & Txtnumero.text & ",'" & txtCodEmp.text & "'," & Txtcodinf.text & ",'" & Txtnomnot.text & "',TO_DATE('" _
                & Mskfecha.text & "','MM/DD/YYYY') ,TO_DATE('" & hora & "','hh24:mi:ss'),'" & txtPlaca.text & "','" & Txtsitio.text & "','" _
                & Txtnomtes.text & "','" & Txtnomaux.text & "','" & Txttipveh.text & "'," & Txtdescuento.text & ",'" _
                & TxtObservaciones.text & "','Pendiente'," & Msknetpag.text & "," & valdes & ",to_date('" _
                & Mskfecsis.text & "','MM/DD/YYYY'), " & Msknetpag.text & ",'" & Pc & "','" & sAuxAlt & "', '" _
                & sDestino & "' , " & sValCon & "," & sConduce & ",'" & sDesPla & "', " _
                & sValPla & ",'" & sPlanilla & "','" & sTiquete & "'," & sPasVia & "," _
                & CDbl(sPasPla) & ",'" & txtCedula.text & "'," & sAvimulta & ")"
      
      
      Dbs.Execute QRY
      sFecha = Format(Now, "MM/dd/yyyy hh:mm")
      QRY = " INSERT INTO gelogaud " _
                & "(lofecact, lofecnue, lofuncio, " _
                & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
                & "VALUES ( to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),  to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'" _
                & Login & "', to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'SE CREO UNA MULTA', 'MUL', '" _
                & txtPlaca.text & "', '" & Txtnumero.text & "')"
      
      Dbs.Execute QRY
        
      QRY = "UPDATE geparsis " _
                   & "SET psval = '" & Txtnumero.text & "' " _
                 & "WHERE pscod = 'NUMMUL' " _
                   & "AND psdes = 'NUMERO CONSECUTIVO DE MULTAS' "
   
      Dbs.Execute QRY
     
      Dbs.CommitTrans
      On Error GoTo Errores

   End If
  
   MsgBox "Evasion Grabada", 64, "O.K."
         
   Btngrabar.Enabled = False
   Txtnumero.Enabled = False
  
   Exit Sub
  
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
  
   Exit Sub
Resume '
End Sub

Private Sub Btnimprimir_Click()
If Txtnumero.Enabled = False Then
     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
     
     Dbs.Execute "INSERT INTO optempo1 " _
                            & "(temnumero1, temtext1, temnumero2, temtext2, temdate1, temtext3, temtext4, temtext5, temtext6, temtext7, temtext8, temnumdob1,  temtext9, temtext10, temnumdob2, temnumero3, temnumdob3, temdate2, tempc) " _
                    & "SELECT    munumero,   mucodemp, mucodinf,   munomnot, mufecha, muhora,   munumpla,  musitio,  munomtes, muauxcon, mutipveh, mudescue,  muobserv, muestado,  munetpag,  murecpag, muvaldes, mufecsis,'" & Pc & "' " _
                      & "FROM opmultas " _
                     & "WHERE munumero = " & Txtnumero.text & ""
      
     reporte = RpPath + "\regmulta.rpt"
     rptterminal.Connect = G_Conexion_reportes
     rptterminal.ReportFileName = reporte
     rptterminal.Formulas(0) = "empresa = '" & G_NomEmpresa & "'"
     Regmultas.Enabled = False
     rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
     rptterminal.Destination = 0
     rptterminal.Action = 1
     Regmultas.Enabled = True
     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
  End If
End Sub

Private Sub Btnlimpiar_Click()
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
   End If
   Ftb01.Close
   
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMMUL' "
 

   Set Ftb02 = Dbs.Execute(fqry02)
  
   If Not Ftb02.EOF Then
      Txtnumero.text = Ftb02!psval + 1
   End If
   Ftb02.Close
   bandera = 0
   valdes = 0
   Txtcodinf.Enabled = True
   Txtnumero.Enabled = True
   btnconductor.Enabled = False
   Txtdescuento.text = "0"
   txtPlaca.Enabled = True
   txtPlaca.text = ""
   txtPlaca.Enabled = True
   TxtNumInt.text = ""
   TxtNumInt.Enabled = True
   txtCodEmp.text = ""
   Lblcodemp.Caption = ""
   Txttipveh.text = ""
   Lbltipveh.Caption = ""
   Txtcodinf.text = ""
   LblCodInf.Caption = ""
   Txtsitio.text = ""
   txtCedula.text = ""
   Txtnomnot.text = ""
   Txtnomtes.text = ""
   Txtnomaux.text = ""
   Lblauxcon.Caption = ""
   TxtObservaciones.text = ""
   Txtnumero.Enabled = True
  ' Txtnumero.Text = ""
   Mskvalmul.text = ""
   Msknetpag.text = ""
   Btngrabar.Enabled = True
   OptEstado(0).Enabled = True
   OptEstado(0).Value = True
   OptEstado(1).Enabled = False
   OptEstado(2).Enabled = False
   txtCodEmp.Enabled = True
   Txttipveh.Enabled = True
   Mskfecha.Mask = ""
   Mskfecha.text = ""
   Mskfecha.Mask = "##/##/####"
   Mskhora.Mask = ""
   Mskhora.text = ""
   Mskhora.Mask = "##:##"
      
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
   TxtPasVia.text = "0"
   TxtPasPla.text = "0"
   txtPlaca.SetFocus
   L_CodigoEmpresaConvenio = ""
   L_CodigoEmpresaConvenioEmpresarial = ""
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Regmultas")
End Sub

Private Sub Form_Load()
  
   On Error GoTo errorres
   
   TxtPasVia.text = "0"
   TxtPasPla.text = "0"
   fqry02 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'NUMMUL' "
          
   Set Ftb02 = Dbs.Execute(fqry02)
  
  If Not Ftb02.EOF Then
     Txtnumero.text = Ftb02!psval + 1
  End If
  Ftb02.Close
  fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'HORASIS' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
      anosistemas = Fecha
   End If
   Ftb01.Close
   bandera = 0
   valdes = 0
   L_Valor_Excreta = Fn_Parametros("VALEXCRETA", 1)
   L_Aviso_Multa = Fn_Parametros("AVIMULTA", 1)
   L_CodAuxCtr = Fn_Parametros("CODAUXCTR", 1)
   Exit Sub
errorres:
  MsgBox Err.Description, 16, "Error " & Err.Number
  
  Exit Sub
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Mskfecha_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskfecha_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecha_LostFocus()
  If Mskfecha.ClipText <> "" Then
     If Not IsDate(Mskfecha.text) Then
         MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
         Mskfecha.SetFocus
         Exit Sub
     End If
   
     If Val(Left(Mskfecha.text, 2)) < 1 Or Val(Left(Mskfecha.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecha.SetFocus
        Exit Sub
     End If
     
     anomulta1 = Year(Mskfecha.text)
     anosistemas1 = Year(anosistemas)
     If anosistemas1 <> anomulta1 Then
         anomulta = Mskfecha.text
         anosistemas1 = DateAdd("m", -2, anosistemas)
         If CDate(anomulta) < CDate(anosistemas1) Then
            On Error Resume Next
            MsgBox "El año es diferente al actual", 48, "Advertencia"
            Mskfecha.SetFocus
            Exit Sub
         End If
     End If
     If CDate(Mskfecsis.text) < CDate(Mskfecha.text) Then
         On Error Resume Next
         MsgBox "la fecha de la evasion es mayor a la fecha del sistema", 48, "Advertencia"
         Mskfecha.SetFocus
         Exit Sub
      End If
     
  End If
End Sub

Private Sub Mskhora_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskhora_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskhora_LostFocus()
  If Mskhora.ClipText <> "" Then
     longitud = Len(Mskhora.ClipText)
     If longitud = 4 Then
        If Mskhora.ClipText < "0000" Or Mskhora.ClipText > "2400" Then
           MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
           Mskhora.SetFocus
           Exit Sub
        End If
        Minutos = Mid(Mskhora.ClipText, 3)
        If Minutos > 60 Then
           MsgBox "El Valor De Minutos Que Coloco No Son Validos", 48, "Advertencia"
           Mskhora.SetFocus
           Exit Sub
        End If
     Else
        MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
        Mskhora.SetFocus
        Exit Sub
     End If
  End If
End Sub

Private Sub Msknetpag_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
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

Private Sub Txtauxalt_GotFocus()
   Senalar_Texto Me
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
      If bandera = 0 Then
         If Trim(Txtauxalt.text) = Trim(L_CodAuxCtr) Then
            MsgBox "No se puede crear una evasion con este control", 48, "Advertencia"
            Txtauxalt.SetFocus
            Exit Sub
         End If
      End If
      
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

Private Sub txtcedula_DblClick()
   forma = "20"
   forma1 = "3"
   buscar.Show 1
   txtcedula_LostFocus
End Sub

Private Sub txtcedula_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcedula_KeyPress(KeyAscii As Integer)
   
   If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
      
   End If
End Sub

Private Sub txtcedula_LostFocus()
   On Error GoTo err_txtcedula
   Txtnomnot.text = ""
   
   If Trim(txtCedula.text) <> "" Then
      
      QRY = "SELECT conombres, coapellidos " _
           & "FROM coconductor " _
          & "WHERE cocedula = '" & txtCedula.text & "'"
     
      Set Ftb04 = Dbs.Execute(QRY)
     
      If Not Ftb04.EOF Then
         Txtnomnot.text = Mid(Ftb04!conombres & " " & Ftb04!coapellidos, 1, 50)
      Else
        MsgBox "Número de cedula no existe como conductor", 48, "Advertencia"
        txtCedula.SetFocus
     End If
     Ftb04.Close
  End If
  
  
  Exit Sub
err_txtcedula:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub


Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "10"
  txtCodEmp.text = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   If txtCodEmp.text <> "" Then
       
      fqry01 = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & txtCodEmp.text & "'"
           
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
         If txtPlaca.text <> "" Then
            fqry01 = "SELECT * " _
                     & "FROM ophojveh " _
                    & "WHERE hvplaca = '" & txtPlaca.text & "' "
             
            Set Ftb01 = Dbs.Execute(fqry01)
     
            If Not Ftb01.EOF Then
               Ftb01.Close
               fqry01 = "SELECT * " _
                        & "FROM ophojveh " _
                       & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                         & "AND hvcodemp = '" & txtCodEmp.text & "' "
              
              
               Set Ftb01 = Dbs.Execute(fqry01)
              
               If Ftb01.EOF Then
                  Ftb01.Close
                  MsgBox "Este Vehiculo No Pertenece A " & txtCodEmp.text, 48, "Advertencia"
                  Exit Sub
               Else
                  Ftb01.Close
               End If
            Else
               Ftb01.Close
            End If
        
         End If
      Else
         Ftb01.Close
         fqry01 = "SELECT henombre " _
                    & "FROM ophojempcon " _
                   & "WHERE hecodigo = '" & txtCodEmp.text & "'"
        
         Set Ftb01 = Dbs.Execute(fqry01)
   
         If Not Ftb01.EOF Then
            Lblcodemp.Caption = Ftb01!henombre
           
         Else
            MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
            forma = "16"
            forma1 = "10"
            buscar.Show 1
            txtCodEmp.Enabled = True
            txtCodEmp.SetFocus
         End If
         Ftb01.Close
        
        
        
     End If
  End If
End Sub

Private Sub Txtcodinf_DblClick()
  forma = "4"
  forma1 = "1"
  Txtcodinf.text = ""
  buscar.Show 1
  Txtcodinf_LostFocus
End Sub

Private Sub Txtcodinf_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtcodinf_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtcodinf_LostFocus()
   If Txtcodinf.text <> "" And Btngrabar.Enabled = True Then
     
      If Txtdescuento.text <> "0" Then
         Txtdescuento.text = "0"
      End If
     
      Frmdestino.Enabled = False
      Frmplanilla.Enabled = False
      Txtdestino.Enabled = True
      Txtnumcon.Enabled = True
      Txtdespla.Enabled = True
      Txtplanilla.Enabled = True
      Txttiquete.Enabled = True
      TxtPasVia.Enabled = True
      TxtPasPla.Enabled = True
      TxtPasVia.text = "0"
      TxtPasPla.text = "0"
      L_Bandera_Multas = "N"
      
      fqry01 = "SELECT dsdes, dsval " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtcodinf.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'INFRACCIONES PARA COBRAR' "
               
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         
         'van quemados estos por porque son los que se cobran
         'se abrio a un catalogo porque van a crear otras infracciones para cobro manual de la infraccion
         '*****If Txtcodinf.Text = "51" Or Txtcodinf.Text = "52" Then
         'If Txtcodinf.Text = "51" Then 'No se tiene en cuenta el 52 porque es una evacion segun requerimiento 1075 con soporte 1481 de Oct 16 de 2007
         
         If Ftb01!dsval = 1 Then
            valtiq = InputBox("Coloque el valor de la tasa de uso o del cobro.", "Valor del Tiquete")
         Else
            valtiq = L_Valor_Excreta
            L_Bandera_Multas = "S"
         End If
         
         If IsNumeric(valtiq) Then
            If CDbl(valtiq) > 0 Then
               fqry02 = "SELECT * " _
                        & "FROM geparsis " _
                       & "WHERE pscod = '" & Txtcodinf.text & "' "
            
               Set Ftb02 = Dbs.Execute(fqry02)
            
               If Not Ftb02.EOF Then
                  If Ftb01!dsval = 1 Then
                     Frmdestino.Enabled = True
                     Frmplanilla.Enabled = True
                  End If
                  If Txtcodinf.text = "52" Then
                  'esta queda quemada porque con esta infracción no se solicita el numero de tasa
                  'y no vale la pena crear un catalogo para esta infraccion
                     Txtnumcon.Enabled = False
                  End If
                  
                  LblCodInf.Caption = Ftb02!psdes
                  Mskvalmul.text = valtiq '* ftb02!psval
                  Msknetpag.text = valtiq '* ftb02!psval
                  If Ftb01!dsval = 1 Then
                     Txtdestino.SetFocus
                  End If
                  Ftb02.Close
                  
                  numero = str(Msknetpag.text)
                  digitos = Len(numero)
                  inicio1 = digitos - 1
                  valor = Mid(numero, inicio1, 2)
                  If valor <> "00" Then
                     complemento = 50 - valor
                     valtot = Abs(complemento) + numero
                     Msknetpag.text = valtot
                     Mskvalmul.text = valtot
                  Else
                     Msknetpag.text = Mskvalmul.text
                  End If
                  Exit Sub
               End If
            Else
               MsgBox "El valor debe ser mayor a cero", 48, "Advertencia"
               Txtcodinf.SetFocus
               Exit Sub
            End If
         Else
            If Txtcodinf <> "8" Then
               MsgBox "Debe colocar un valor para poder cancelar la evasion", 48, "Advertencia"
               Txtcodinf.SetFocus
               Exit Sub
            Else
               Txtdestino.SetFocus
               Exit Sub
            End If
         End If
      Else
         If Txtcodinf.text = "19" Or Txtcodinf.text = "20" Then
            'van quemados porque solo ha estas dos infracciones se solicitan los pasajeros
            Frmdestino.Enabled = True
            Frmplanilla.Enabled = True
            Txtdestino.Enabled = False
            Txtnumcon.Enabled = False
            Txtdespla.Enabled = False
            Txtplanilla.Enabled = False
            Txttiquete.Enabled = False
            TxtPasVia.Enabled = True
            TxtPasPla.Enabled = True
            TxtPasVia.SetFocus
         End If
     End If
     fqry02 = "SELECT * " _
              & "FROM geparsis " _
             & "WHERE pscod = '" & Txtcodinf.text & "' "
        
     Set Ftb02 = Dbs.Execute(fqry02)
     
     If Not Ftb02.EOF Then
        LblCodInf.Caption = Ftb02!psdes
        
         fqry01 = "SELECT psval " _
                  & "FROM geparsis " _
                 & "WHERE pscod = 'SALMIN' "
                 
         Set Ftb01 = Dbs.Execute(fqry01)
         
         If Not Ftb01.EOF Then
            Mskvalmul.text = CLng((Ftb01!psval / 30) * Ftb02!psval)
            
            REDONDEAR
         End If
         Ftb02.Close
         Ftb01.Close
        
     Else
        Ftb02.Close
        MsgBox "El Codigo De Infraccion No Existe", 64, "Atencion"
        forma = "4"
        forma1 = "1"
        buscar.Show 1
        Txtcodinf.SetFocus
        Exit Sub
     End If
  End If
End Sub

Private Sub Txtdescuento_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


Private Sub Txtdescuento_LostFocus()
  If Txtdescuento.text <> "" And Mskvalmul.text <> "" And Txtdescuento.text <> "0" Then
     fqry01 = "SELECT psval " _
              & "FROM geparsis " _
             & "WHERE pscod = 'PORDESMUL' "
             
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        If CDbl(Txtdescuento.text) > CDbl(Ftb01!psval) Then
           Ftb01.Close
           MsgBox "NO PUEDE COLOCAR UN PORCENTAJE MAYOR A " & Ftb01!psval, 48, "Advertencia"
           Txtdescuento.text = ""
           Txtdescuento.SetFocus
           Exit Sub
        End If
     End If
     Ftb01.Close
     
     Msknetpag.text = Mskvalmul.text - ((CDbl(Mskvalmul.text) * CDbl(Txtdescuento.text)) / 100)
     numero = str(Msknetpag.text)
     digitos = Len(numero)
     inicio1 = digitos - 1
     valor = Mid(numero, inicio1, 2)
     If valor <> "00" Then
        complemento = 100 - valor
        valtot = Abs(complemento) + numero
        Msknetpag.text = valtot
     End If
     valdes = CDbl(Mskvalmul.text) - CDbl(Msknetpag.text)
  End If
  If Txtdescuento.text = "0" Then
     Msknetpag.text = Mskvalmul.text
  End If
  
End Sub



Private Sub Txtdespla_DblClick()
  forma = "3"
  forma1 = "20"
  Txtdespla.text = ""
  If L_CodigoEmpresaConvenio <> "" Then
      G_Codigo_Empresa = L_CodigoEmpresaConvenio
   Else
      If TxtNumConvenio.text = "" Then
         G_Codigo_Empresa = Trim(txtCodEmp.text)
      Else
         QRY = "SELECT cecodemp  " _
               & "FROM opconemp " _
               & "WHERE cenumero ='" & TxtNumConvenio.text & "' " _
               & " AND ceempcon = '" & txtCodEmp.text & "' "
                      
         Set Ftb02 = Dbs.Execute(QRY)
         If Not Ftb02.EOF Then
            G_Codigo_Empresa = Ftb02!cecodemp
         End If
         Ftb02.Close
      End If
   End If
  buscar.Show 1
  Txtdespla.SetFocus
  Txtdespla_LostFocus
End Sub

Private Sub Txtdespla_GotFocus()
   Senalar_Texto Me
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
              & "WHERE rucodigo = '" & Txtdespla.text & "' "
             
      Set Ftb04 = Dbs.Execute(fqry01)
        
      If Not Ftb04.EOF Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb04!ruestado = "Inactiva" Then
            Ftb04.Close
            Txtdespla.SetFocus
            MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
            Exit Sub
         End If
        
         Lbldespla.Caption = Ftb04!rudestino
         Ftb04.Close
         
         If ValidaEmpresaRuta(Txtdespla.text, txtCodEmp.text) = False Then
            MsgBox "La Empresa no tiene autorizada esta ruta", 48, "Advertencia"
            Txtdespla.SetFocus
            Exit Sub
         End If
         
         QRY = "SELECT * " _
               & "FROM opconvenios " _
              & "WHERE cvcodrut = '" & Txtdespla.text & "' " _
                & "AND cvcodser = '" & Txttipveh.text & "' "
                
         
         Set Ftb04 = Dbs.Execute(QRY)
      
         If Not Ftb04.EOF Then
            Mskvalpla.text = Ftb04!cvvalor
            Ftb04.Close
            
            If Txtcodinf.text = "51" And Trim(Mskvalor.text) <> "" Then
            'esta queda quemada porque con esta infracción solo se cobrar la diferencia de las dos tasas _
            y no vale la pena crear un catalogo para esta infraccion _
            y en el caso 6548 solo lo solicitaron para elusión y se le pregunto a Nelson y él dijo que _
            solo para la infracción 51 que las otras que estan parametrizadas en el catalogo _
            INFRACCIONES PARA COBRAR el valor es el de la tasa que ellos digitan.
            
               Mskvalmul.text = Abs(CDbl(Mskvalpla.text) - CDbl(Mskvalor.text))
               Msknetpag.text = Mskvalmul.text
            End If
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
Resume
End Sub

Private Sub Txtdestino_DblClick()
 
   forma = "3"
   forma1 = "19"
   Txtdestino.text = ""
  
   If L_CodigoEmpresaConvenio <> "" Then
      G_Codigo_Empresa = L_CodigoEmpresaConvenio
   Else
      If TxtNumConvenio.text = "" Then
         G_Codigo_Empresa = Trim(txtCodEmp.text)
      Else
         QRY = "SELECT cecodemp  " _
               & "FROM opconemp " _
               & "WHERE cenumero ='" & TxtNumConvenio.text & "' " _
               & " AND ceempcon = '" & txtCodEmp.text & "' "
                      
         Set Ftb02 = Dbs.Execute(QRY)
         If Not Ftb02.EOF Then
            G_Codigo_Empresa = Ftb02!cecodemp
         End If
         Ftb02.Close
      End If
   End If
   buscar.Show 1
   Txtdestino.SetFocus
   Txtdestino_LostFocus
End Sub

Private Sub Txtdestino_GotFocus()
   Senalar_Texto Me
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
             
      Set Ftb04 = Dbs.Execute(fqry01)
        
      If Not Ftb04.EOF Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb04!ruestado = "Inactiva" Then
            Ftb04.Close
            Txtdestino.SetFocus
            MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
            Exit Sub
         End If
         Lbldestino.Caption = Ftb04!rudestino
         Ftb04.Close
          
         
         If ValidaEmpresaRuta(Txtdestino.text, txtCodEmp.text) = False Then
            Txtdestino.SetFocus
            MsgBox "La Empresa no tiene autorizada esta ruta", 48, "Advertencia"
            Exit Sub
         End If
         QRY = "SELECT * " _
               & "FROM opconvenios " _
              & "WHERE cvcodrut = '" & Txtdestino.text & "' " _
                & "AND cvcodser = '" & Txttipveh.text & "' "
                
         
         Set Ftb04 = Dbs.Execute(QRY)
      
         If Not Ftb04.EOF Then
            Mskvalor.text = Ftb04!cvvalor
            Ftb04.Close
            
            If Txtcodinf.text = "51" And Trim(Mskvalpla.text) <> "" Then
            'esta queda quemada porque con esta infracción solo se cobrar la diferencia de las dos tasas _
            y no vale la pena crear un catalogo para esta infraccion _
            y en el caso 6548 solo lo solicitaron para elusión y se le pregunto a Nelson y él dijo que _
            solo para la infracción 51 que las otras que estan parametrizadas en el catalogo _
            INFRACCIONES PARA COBRAR el valor es el de la tasa que ellos digitan.
            
               Mskvalmul.text = Abs(CDbl(Mskvalpla.text) - CDbl(Mskvalor.text))
               Msknetpag.text = Mskvalmul.text
            End If
            
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
   Resume
End Sub

Private Sub Txtnomaux_DblClick()
  forma = "15"
  forma1 = "15"
  Tabla_Catalogo = "AUXILIARES DE CONTROL"
  Txtnomaux.text = ""
  buscar.Show 1
  Txtnomaux_LostFocus
End Sub

Private Sub Txtnomaux_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnomaux_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(LCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnomaux_LostFocus()
   If Txtnomaux.text <> "" Then
      If bandera = 0 Then
         If Trim(Txtnomaux.text) = Trim(L_CodAuxCtr) Then
            MsgBox "No se puede crear una evasion con este control", 48, "Advertencia"
            Txtnomaux.SetFocus
            Exit Sub
         End If
      End If
  
      
  
     fqry01 = "SELECT * " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & Txtnomaux.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'AUXILIARES DE CONTROL' "
               
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lblauxcon.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        Ftb01.Close
        On Error Resume Next
        MsgBox "Auxiliar De Control No Existe", 48, "Advertencia"
        Txtnomaux.SetFocus
        Exit Sub
     End If
  End If
End Sub

Private Sub Txtnomnot_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnomnot_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


Private Sub Txtnomtes_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnomtes_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumcon_GotFocus()
   Senalar_Texto Me
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
      If Trim(Txtcodinf.text) <> "52" Then
        QRY = "SELECT * " _
              & "FROM coconduc " _
             & "WHERE conumero = " & Trim(Txtnumcon.text) & ""
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
           Mskfecsal.text = Format(Tbl1!cohorsal, "MM/dd/yyyy HH:mm")
           
           Tbl1.Close
        Else
           Tbl1.Close
           On Error Resume Next
           MsgBox "NUMERO DE TASA DE USO no existe", 48, "Advertencia"
           Txtnumcon.SetFocus
           Exit Sub
        End If
     End If
   Else
      Txtnumcon.text = "0"
   End If
   
   Exit Sub
Err_Txtnumcon_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub
Private Sub TxtNumConvenio_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtNumConvenio_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub
Private Sub TxtNumero_GotFocus()
   Senalar_Texto Me
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
                & "AND muestado = 'Pendiente' "
     
      Set Ftb01 = Dbs.Execute(fqry01)
       
      If Not Ftb01.EOF Then
         Txtnumero.Enabled = False
         txtPlaca.Enabled = False
         Txtcodinf.Enabled = False
         btnconductor.Enabled = True
         txtPlaca.text = Ftb01!munumpla
         Txtnumero.text = Ftb01!munumero
         txtCodEmp.text = Ftb01!mucodemp
         Txtcodinf.text = Ftb01!mucodinf
         Txtnomnot.text = Ftb01!munomnot
         Mskfecha.text = Format(Ftb01!mufecha, "MM/DD/YYYY")
         Mskhora.text = Format(Ftb01!muhora, "hh:mm")
         Txtsitio.text = Ftb01!musitio
         Txtnomtes.text = Ftb01!munomtes
         Txtnomaux.text = Ftb01!muauxcon
         Txttipveh.text = Ftb01!mutipveh
         Txtdescuento.text = Ftb01!mudescue
         TxtObservaciones.text = Ftb01!muobserv
         Msknetpag.text = Ftb01!munetpag
         Mskvalmul.text = Ftb01!musaldo
         txtCedula.text = ""
         L_Cedula_Conductor = ""
         If Not IsNull(Ftb01!mucedula) Then
            txtCedula.text = Ftb01!mucedula
            L_Cedula_Conductor = Ftb01!mucedula
         End If
         If Txtdescuento.text <> "0" Then
            Btngrabar.Enabled = False
         Else
            bandera = 1
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
            If Not IsNull(Ftb01!mupasvia) Then
                TxtPasVia.text = Ftb01!mupasvia
            Else
                TxtPasVia.text = 0
            End If
            If Not IsNull(Ftb01!mupaspla) Then
                TxtPasPla.text = Ftb01!mupaspla
            Else
                TxtPasPla.text = 0
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
                 & "WHERE hvplaca = '" & txtPlaca.text & "' "
         
         Set Ftb01 = Dbs.Execute(fqry01)
         
         If Not Ftb01.EOF Then
            TxtNumInt.Enabled = False
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
               'Mskvalmul.Text = CLng((ftb01!psval / 30) * ftb02!psval)
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
         txtcodemp_LostFocus
         Txttipveh_LostFocus
         Txtnomaux_LostFocus
         Txtsitio_LostFocus
         Txtauxalt_LostFocus
            
         If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
            Mskvalmul.text = (Msknetpag.text * 100) / (100 - Txtdescuento.text)
         End If
      Else
         Ftb01.Close
         
         fqry01 = "SELECT muestado " _
                  & "FROM opmultas " _
                 & "WHERE munumero = " & Txtnumero.text & " "
                 
         Set Ftb01 = Dbs.Execute(fqry01)
                 
         If Not Ftb01.EOF Then
            
            If Ftb01(0) = "Abonada" Then
               MsgBox "La Evasion, Se Encuentra Abonada", 48, "Advertencia"
            ElseIf Ftb01(0) = "Pagada" Then
               MsgBox "Numero De Evasion Ya Existe, Debe Colocar Otro", 48, "Advertencia"
            End If
            Txtnumero.SetFocus
            Ftb01.Close
            Exit Sub
         End If
         Ftb01.Close
      End If
   Else
      Txtnumero.text = "0"
   End If
End Sub

Private Sub Txtnumint_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtobservaciones_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtPasPla_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPasPla_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub TxtPasVia_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPasVia_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub TxtPlaca_DblClick()
  forma = "7"
  forma1 = "1"
  Txtnumero.text = ""
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
End Sub

Private Sub TxtPlaca_GotFocus()
   Senalar_Texto Me
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
  
     If longitud < 6 Then
        MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
        txtPlaca.SetFocus
        Exit Sub
     End If
     
     fqry01 = "SELECT * " _
              & "FROM ophojveh " _
              & "WHERE hvplaca = '" & txtPlaca.text & "' "
              
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        TxtNumInt.text = Ftb01!hvnumint
        txtCodEmp.text = Ftb01!hvcodemp
        Txttipveh.text = Ftb01!hvtipveh
        
        L_CodigoEmpresaConvenio = ""
        If Not IsNull(Ftb01!hvempcon) Then
            L_CodigoEmpresaConvenio = Ftb01!hvempcon
        End If
        
        Ftb01.Close
        txtcodemp_LostFocus
        Txttipveh_LostFocus
        txtPlaca.Enabled = False
        TxtNumInt.Enabled = False
        txtCodEmp.Enabled = False
        Txttipveh.Enabled = False
        
        
     Else
         MsgBox "El Vehiculo No Existe En La Base de Datos", 48, "Advertencia"
         txtPlaca.SelStart = 0
         txtPlaca.SelLength = Len(txtPlaca.text)
         txtPlaca.SetFocus
     End If
   End If
End Sub

Private Sub Txtplanilla_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtplanilla_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub Txtsitio_DblClick()
   forma = "15"
   forma1 = "24"
   Tabla_Catalogo = "SITIOS INFRACCION"
   Txtsitio.text = ""
   buscar.Show 1
   Txtsitio_LostFocus
End Sub

Private Sub Txtsitio_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtsitio_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtsitio_LostFocus()
   On Error GoTo ERR_Txtsitio_LostFocus
   
   If Txtsitio.text <> "" Then
      fqry01 = "SELECT dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtsitio.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'SITIOS INFRACCION' "
               
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         Lblsitio.Caption = Ftb01!dsdes
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "SITIO DE INFRACCION no existe", 48, "Advertencia"
         Txtsitio_DblClick
      End If
   End If
   Exit Sub
ERR_Txtsitio_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Txttipveh_DblClick()
  forma = "15"
  forma1 = "1"
  Tabla_Catalogo = "TIPOS DE VEHICULO"
  Txttipveh.text = ""
  buscar.Show 1
  Txttipveh_LostFocus
End Sub

Private Sub Txttipveh_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txttipveh_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtvalcon_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtvalmul_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txttipveh_LostFocus()
  If Txttipveh.text <> "" Then
     fqry01 = "SELECT dsdes " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & Txttipveh.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'TIPOS DE VEHICULO' "
               
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lbltipveh.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        Ftb01.Close
        MsgBox "Tipo De Vehiculo No Existe", 48, "Advertencia"
        forma = "15"
        forma1 = "1"
        Tabla_Catalogo = "TIPOS DE VEHICULO"
        buscar.Show 1
        Txttipveh.SetFocus
     End If
  End If
End Sub

Private Function REDONDEAR()
  If Mskvalmul.text <> "0" Then
  
      numero = str(Mskvalmul.text)
      digitos = Len(numero)
      inicio1 = digitos - 1
      valor = Mid(numero, inicio1, 2)
      If valor <> "00" Then
         complemento = 100 - valor
         valtot = Abs(complemento) + numero
         Msknetpag.text = valtot
         Mskvalmul.text = valtot
      Else
         Msknetpag.text = Mskvalmul.text
      End If
   Else
      Msknetpag.text = Mskvalmul.text
   End If
End Function

Private Sub Txttiquete_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txttiquete_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub
