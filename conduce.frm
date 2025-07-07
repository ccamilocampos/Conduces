VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C0A57E85-C72B-44C4-A7EF-EBE05047E8C6}#1.1#0"; "CheckPeople3.ocx"
Begin VB.Form Capconduce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta de Tasas de Uso"
   ClientHeight    =   10110
   ClientLeft      =   2355
   ClientTop       =   1560
   ClientWidth     =   7605
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPasajeros 
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
      Left            =   6480
      MaxLength       =   2
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   0
   End
   Begin VB.CheckBox chkDatosPersonales 
      Caption         =   "¿Tratamiento de datos personales?"
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
      Height          =   375
      Left            =   120
      TabIndex        =   93
      Top             =   7200
      Width           =   2775
   End
   Begin VB.TextBox TxtPeajeUltimo 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Ultimo Peaje por ruta"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Txtcopias 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "1"
      ToolTipText     =   "Cantidad de Copias a imprimir"
      Top             =   9120
      Width           =   1695
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
      Left            =   6480
      Picture         =   "conduce.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salir de la forma actual"
      Top             =   9000
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
      Left            =   6480
      Picture         =   "conduce.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   8160
      Width           =   855
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
      Left            =   6480
      Picture         =   "conduce.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ingresa la tasa de uso a la base de datos"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Btnimprimir 
      Caption         =   "&Imprimir"
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
      Left            =   6480
      Picture         =   "conduce.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir recibo de conduce"
      Top             =   7320
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
      Left            =   6480
      Picture         =   "conduce.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Buscar Informacion de Conduces Ingresados"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton btnmultas 
      Caption         =   "&Cancelar"
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
      Left            =   6480
      Picture         =   "conduce.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Cancelación de multas/ Comparendos"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton BtnParqueadero 
      Caption         =   "&Parquea"
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
      Left            =   6480
      Picture         =   "conduce.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Pago de parqueadero"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton btncambio 
      Caption         =   "C&ambio"
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
      Left            =   6480
      Picture         =   "conduce.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Cambio de tasas de uso"
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame FrmTipo 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5040
      TabIndex        =   27
      Top             =   0
      Width           =   1215
      Begin VB.OptionButton OptTipo 
         Caption         =   "Cupo"
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
         Index           =   1
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Linea"
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
         Left            =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   3135
      Begin VB.OptionButton OptTipoPlaca 
         Caption         =   "C&olombiana"
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptTipoPlaca 
         Caption         =   "&Extranjera"
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
         Left            =   1560
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   3360
      TabIndex        =   37
      Top             =   840
      Width           =   2895
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de Transporte (presione DblClick para ayuda)"
         Top             =   240
         Width           =   1095
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
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resolución 315"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   60
      Top             =   4680
      Width           =   6135
      Begin VB.TextBox TxtPeaje 
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
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   15
         ToolTipText     =   "Código barras de peajes"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Txtconact 
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
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Nombre del conductor actual"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtcedula 
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
         MaxLength       =   12
         TabIndex        =   8
         ToolTipText     =   "Número cédula del conductor"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtapellido 
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
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Apellido del conductor actual"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Txtapellido1 
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
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Apellido del conductor Actual"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Txtcedula1 
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
         MaxLength       =   12
         TabIndex        =   11
         ToolTipText     =   "Número cédula del conductor"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Txtconact1 
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
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Nombre del conductor Actual"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtNumAlistamiento 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   14
         ToolTipText     =   "Númerode alistamiento del vehículo"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Peajes:"
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
         Index           =   29
         Left            =   3480
         TabIndex        =   63
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Index           =   8
         Left            =   1440
         TabIndex        =   43
         Top             =   240
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
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido"
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
         Index           =   19
         Left            =   3720
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido"
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
         Index           =   20
         Left            =   3720
         TabIndex        =   47
         Top             =   840
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
         Index           =   21
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Index           =   22
         Left            =   1440
         TabIndex        =   51
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "No. Alistamiento:"
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
         Index           =   23
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkcredito 
      Caption         =   "Crédito"
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
      Left            =   3480
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Txtcaja 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Txtorigen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   7
      ToolTipText     =   "Código de la ruta de origen (presione DblClick para ayuda)"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Txtoritra 
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
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   6
      ToolTipText     =   "Digite <T> para Transito y <O> para Origen"
      Top             =   3600
      Width           =   855
   End
   Begin Crystal.CrystalReport repcon 
      Left            =   7080
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin MSMask.MaskEdBox Mskcambio 
      Height          =   375
      Left            =   4560
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de cambio"
      Top             =   8640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskvalor 
      Height          =   375
      Left            =   4560
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskhora 
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
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
      Format          =   "hh:mm"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskfecha 
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
      Format          =   "MMM/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskrecibo 
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      ToolTipText     =   "Dinero recibido para el pago del conduce"
      Top             =   8160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Txtcodrut 
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
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   5
      ToolTipText     =   "Código de la ruta (Presione DblClick para ayuda)"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Txtcodemp 
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Código de la empresa de Transporte (presione DblClick para ayuda)"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Txtnumcon 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Número de tasa de uso"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Vehículo"
      BeginProperty Font 
         Name            =   "Times New Roman"
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
      TabIndex        =   30
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox Txttipveh 
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Clase de servicio del vehículo"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Txtnumint 
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Número interno del vehículo"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txtnumpla 
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
         Left            =   720
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Placa del vehículo"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbltipveh 
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
         Left            =   4680
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Clase"
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
         Left            =   3360
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Núm.Int."
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
         Left            =   1800
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
   End
   Begin MSMask.MaskEdBox Mskparqueadero 
      Height          =   375
      Left            =   1680
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de cambio"
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskMultas 
      Height          =   375
      Left            =   1680
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de cambio"
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskingreso 
      Height          =   375
      Left            =   1680
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msktottas 
      Height          =   375
      Left            =   4560
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de Cambio"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   6720
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin MSMask.MaskEdBox Mskdinero 
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      ToolTipText     =   "Dinero Recibido para el pago del Conduce"
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskdevolver 
      Height          =   375
      Left            =   4560
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de Cambio"
      Top             =   9600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin Biometria2.CheckPeople CheckPeople1 
      Height          =   1155
      Left            =   6480
      TabIndex        =   75
      Top             =   3960
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   2037
      DefaultPicture  =   "conduce.frx":2210
      PermitirCancelar=   0   'False
      Sensibilidad    =   0
      Seguridad       =   0
      ImagenFondo     =   "conduce.frx":222C
   End
   Begin MSMask.MaskEdBox MskValorTasaNorte 
      Height          =   375
      Left            =   1680
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   8640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskValTag 
      Height          =   375
      Left            =   4560
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de cambio"
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskhorsis 
      Height          =   255
      Left            =   5280
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm AM/PM"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox ChkTag 
      Caption         =   "Cobrar Tag"
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
      Left            =   6360
      TabIndex        =   92
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Crystal.CrystalReport rptMultas 
      Left            =   7080
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Crystal.CrystalReport rptPar 
      Left            =   6480
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label lblPasajeros 
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
      Left            =   6480
      TabIndex        =   94
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Lblcupo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   6360
      TabIndex        =   89
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Salida:"
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
      TabIndex        =   83
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Tag $"
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
      Index           =   30
      Left            =   3240
      TabIndex        =   91
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "T. Permanencia $"
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
      Index           =   16
      Left            =   120
      TabIndex        =   88
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Vr Tasa Norte $"
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
      Index           =   27
      Left            =   120
      TabIndex        =   87
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dinero Recibido $"
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
      Index           =   25
      Left            =   120
      TabIndex        =   86
      Top             =   9720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "H. de Ingreso:"
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
      Index           =   18
      Left            =   120
      TabIndex        =   85
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Comparendos $"
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
      Index           =   17
      Left            =   120
      TabIndex        =   84
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cambio $"
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
      Index           =   26
      Left            =   3240
      TabIndex        =   81
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total $"
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
      Index           =   11
      Left            =   3240
      TabIndex        =   80
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Vr Alcoh. $"
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
      Index           =   10
      Left            =   3240
      TabIndex        =   79
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Vr Tasa Uso $"
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
      Left            =   3240
      TabIndex        =   78
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Horario de Salida:"
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
      Left            =   3240
      TabIndex        =   77
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Copias:"
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
      Index           =   28
      Left            =   3240
      TabIndex        =   76
      Top             =   9240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   7560
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   10080
   End
   Begin VB.Label Lblorigen 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   74
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Lbloritra 
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
      Left            =   2880
      TabIndex        =   73
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Lblnomemp 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   72
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Lblnomrut 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   71
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Código de  Origen"
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
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   70
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Origen/ Transito"
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
      Index           =   13
      Left            =   120
      TabIndex        =   69
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Código de  Ruta:"
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
      TabIndex        =   68
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Código Empresa:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   6360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
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
      Index           =   14
      Left            =   1800
      TabIndex        =   65
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "No. Tasa Uso:"
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
      TabIndex        =   64
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Capconduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim horven
Dim Horope As String
Dim alcoholemia As String
Dim numsalveh As Integer  'numero de salidas del vehiculo
Dim prealco   'presentarse alcoholemia
Dim ALCTODOS As String  'VARIABLE PARA VER SI YA SE CUMPLIO EL
                        'TOTAL DE VEHICULOS A REVISAR
Dim ESTVEH As String    'ESTADO DEL VEHICULO EN EL LOST FOCUS DE
                        'DE PLACA LO PASA AL LOST FOCUS DE
                        'ORIGEN / TRANSITO
Dim presentaalcoh As String
Dim influencia As String
Dim RevisionOri As Double
Dim RevisionTra As Double
Dim presenprueba As String
Dim valparqueadero As String
Dim tiempoesta As String
Dim sinviaje As String
Dim ValorParqueo As String
Dim crecon As String 'credito contado para la impresion
Dim horasis
Dim codrut   '
Dim direruta As String 'direccion de la ruta
Dim plaexis As String
Dim AviMulta As Long
Dim NumCar As Long
Dim contador As Integer
Dim ConHora As String
Dim ConNombre As String
Dim Desc_Server As Boolean
Dim sNumConvenio  As Double
Dim L_MinVenTas As Double     'variable para determnar la hora permitida de venta de tasa
Dim L_HoraSist As Date
Dim L_MinValTasa As Double    'variable para determnar la horas para volver a comprar la tasa
Dim L_MinValTasaCedula As Double    'variable para determnar la horas para volver a comprar la tasa los conductores
Dim L_ValidaCentral As String 'variabla que almacena el valor del parametro VALIDACENT
Dim L_MinCenSur As String
Dim L_Clase As String
Dim L_Terminal As String      'identificar el terminal
Dim L_Distancia As Double
Dim L_Valida_Horario As String
Dim L_AvisoParqueadero As Double
Dim L_Consecutivo_Parqueadero As Double
Dim L_Codigo_Empresa As String
Dim L_Nombre_Empresa  As String


Dim L_Tipo_Venta As String
Dim L_Operador  As String

Dim L_CodigoEmpresaConvenio As String
Dim L_CodigoEmpresaConvenioEmpresarial As String
Dim LTipoConvenio As String


'WECC 08/21/2012 VARIABLES MANEJO DE INTEGRA
Dim LHoraAsignada As String

Dim LDirreccionRuta As String 'direccion de la ruta

Dim LCodigoRutaIntegra As String
Dim LCodigoEmpresaIntegra As String
Dim LClaseIntegra As String
Dim LValorPublicidadIntegra As String
Dim LValorPuertasIntegra As String
Dim LValorOtrosIntegra As String
Dim LValorParqueoIntegra As String
Dim LNumeroPuestosIntegra As String
Dim LCedulaIntegra As String
Dim LTipoServicioIntegra As String
Dim LQryTransmitir As String
Dim ExisteCedula
Dim ExisteCedula1

Dim bandera_huella
Dim L_Valida_CedHuella

Dim bandera_huella1
Dim L_Valida_CedHuella1

Dim Respuesta_huella
Dim Respuesta_huella1

'Variables Huellero
Dim RegistoHuella As Boolean
Dim errorHuellero As Boolean
Dim RegistoHuella1 As Boolean
Dim errorHuellero1 As Boolean
'---------------------------
Dim L_HorasPermitidas As String

Dim L_ValorTransitoSur As Double

Dim L_Minutos_2_Conductores As Double   'Variable para los minutos de los dos conductores
Dim L_Tiempo_Ruta As Double
Dim Aviso_2_Conductores As Boolean
Dim Aviso_Alistamiento As Boolean
Dim L_Fecha_Aviso_Alistamiento As String
Dim L_Conexion As String
Dim L_Digita_Conductor As String

Dim datosCedula1(13) As String
Dim datosCedula2(13) As String
Dim L_Transmitir_Conductor1 As String
Dim L_Transmitir_Conductor2 As String
Dim L_Empresa_Valida_Dia As String
Dim L_Fecha_salida As String
'wecc 05/20/2015 bandera para controlar el cambio de dia cuando se vende la TU antes de media noche y va a salir despues de media noche
Dim L_Bandera_Fecha As String

'*******************************************************************
'Variables para coenctar webservices
Dim L_Valida_Alistamiento As String

Dim L_Tipo_Convenio As String
Dim L_Pasajeros_Entran As Integer
'WECC 05/23/2016 Cupos - Rodamientos
'Variable para incrementar la hora en minutos y buscar en los despachos por rango de horas
Dim L_Minutos_Despacho As Double
Dim L_Numero_Despacho As Double
Dim L_Hora_Despacho As String
Dim L_HabilitaDevueltas As String


'WECC feb/02/2017 TERMINAL NORTE CAU 14098
Dim L_Numero_Terminal_Norte, L_Valor_Alcoholimetria_Norte As Double
Dim L_Ruta_Transito_Norte, LCodigoRutaNorteIntegra, L_Caja_Norte, L_Terminal_Tasa_Automatica As String
Dim L_Cedula1Norte, L_NombreConductor1Norte, L_Cedula2Norte, L_NombreConductor2Norte As String
Dim L_FechaEntradaNorte, L_HoraEntradaNorte, L_FechaSalidaNorte, L_HoraSalidaNorte, L_DireccionNorte As String


'WECC abr 27 2017 funcionario para reimpresion
Dim sFuncioImpCon As String
Dim sTurnoImpCon As Double
Dim L_Bandera_Copia As Boolean
Dim L_CodigoSuper As String


'WECC Feb 07/2019 CAU2 592 Identifica el tipo de cupo E Empresa C Conductor
Dim L_Tipo_Cupo As String
Dim L_Numero_Cupo As Double

Dim L_Valida_Conductor As String
Dim L_DescontarAviso  As String
Dim L_Placa_Permanencia As String
Dim L_TipoDespachoVehiculo As String

'WECC Nov/25/2019 Caso CAU2 No  1907
Dim L_Minutos_Controlar_Rutas, L_Maximo_Tiempo_Ruta, L_Tiempo_Ruta_Placa As Integer
Dim L_Valida_Control_Ruta, L_Tipo_Conductor1, L_Tipo_Conductor2, L_EstViajeCedula1, L_EstViajeCedula2, L_EstViajePlaca As String

Dim L_EmpresaTransitoCentral, L_BanderaVentaTasa, L_TipoRuta As String
Dim L_PicoyPlaca As String
Dim L_Bandera_Escoger As String

'Sebastian Rondon Jun/13/2022
'Variables conexion Web Services Control Salidas
Dim L_IntegracionProtech As String
Dim L_UrlApiControlSalida As String
Dim L_UrlServer As String

'Sebastián Rondón Noviembre 02 - 2022
Dim L_Bandera_CobroTag
Dim L_Bandera_NumeroTag
Dim L_Bandera_ReimpresionTag As Boolean

'Sebastián Rondon Abril 12 - 2023
Dim L_Bandera_CierreAutomatico As Integer

'Sebastián Rondón - Noviembre 14 de 2023
Dim L_Bandera_Ruta_Lector As String

Dim PasajerosOrigen

Dim s_Num_Multa_Pagada

'Venta de tasas condicionadas
Dim L_Bandera_Tasa_Condicionada As Boolean
Dim L_Bandera_Venta_Condicionada As Boolean
Private Enum tipoConversion
    deCaF
    deFaC
End Enum
Private tipo As tipoConversion

' Estas definiciones están tomadas de lo mostrado en el explorador
' al seleccionar cada una de las funciones del servicio Web
'Private Const cSOAPCaF = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<Alistamiento_Vehiculo xmlns=""http://localhost/"">" & _
                "<vrAlistamiento_Vehiculo>1</vrAlistamiento_Vehiculo>" & _
            "</Alistamiento_Vehiculo>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"

Private Const cSOAPValida = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<Valida xmlns=""http://localhost/"">" & _
                "<sql>1</sql>" & _
            "</Valida>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"
    


Private Sub Transmitir(ByVal sQry As String)
      tipo = deFaC
    '
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para FaC
    parser.loadXML cSOAPValida
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/Valida/sql").text = sQry
    '
    ' Mostrar el código XML enviado al servicio Web
    sBandera = parser.xml
    '
    ' Usar el control Inet para realizar la operación HTTP POST
    'Inet1.Execute txtURL.Text, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://elGuille/WebServices/Tiempo_Conductor"
    '
    enviarComando parser.xml, "http://localhost/Valida"
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
    oHttReq.Open "POST", G_Url_Web_Services, False
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
    sss = s
    '
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.loadXML s
    '
    On Error Resume Next
    '
    'If tipo = deCaF Then
    '    L_Valida_Alistamiento = "Error"
    '    L_Valida_Alistamiento = parser.selectSingleNode("/soap:Envelope/soap:Body/Alistamiento_VehiculoResponse/Alistamiento_VehiculoResult").Text
    'Else
        L_Valida_Alistamiento = "Error"
        L_Valida_Alistamiento = parser.selectSingleNode("/soap:Envelope/soap:Body/ValidaResponse/ValidaResult").text
    'End If
    '
    If Err.Number > 0 Then
        L_Valida_Alistamiento = "Error"
    End If
End Sub

Public Sub leerCedula(ByVal cedula As String, ByRef conductor() As String)

        Dim datoCed As String
        Dim longDatosCedula(13) As Integer
   'On Error GoTo Err_leerCedula
   'Definicion de valores
   'l_verif1 = 22   '22
   'l_verif2 = 18   '40
   'l_verif3 = 8    '48
   'l_cedula = 10   '58
   'l_prinom = 23   '81
   'l_segnom = 23   '104
   'l_priape = 23   '127
   'l_segape = 23   '150
   'l_sexo = 2      '152
   'l_fecnac = 8    '160
   'l_verif4 = 6    '166
   'l_rh = 2        '168
   'l_verif5 = 2    '170

   longDatosCedula(0) = 22 'verif1
   longDatosCedula(1) = 18 'verif2
   longDatosCedula(2) = 8  'verif3
   longDatosCedula(3) = 10 'cedula
   longDatosCedula(4) = 23 'primer apellido
   longDatosCedula(5) = 23 'segundo apellido
   longDatosCedula(6) = 23 'primer nombre
   longDatosCedula(7) = 23 'segundo nombre
   longDatosCedula(8) = 1  'sexo
   longDatosCedula(9) = 8  'fecha nacimiento
   longDatosCedula(10) = 6 'verif4
   longDatosCedula(11) = 2 'RH
   longDatosCedula(12) = 2 'verif5

   
   'Limpieza de datos
   For i = 0 To 12
       conductor(i) = ""
   Next i
   
   If Len(cedula) = 170 Then
      datoCed = Trim(cedula)
   Else
     datoCed = Trim(Mid(cedula, Len(cedula) - 169))
   End If

   conductor(0) = Trim(Mid(datoCed, 1, longDatosCedula(0)))       'verif1
   conductor(1) = Trim(Mid(datoCed, 23, longDatosCedula(1)))      'verif2
   conductor(2) = Trim(Mid(datoCed, 41, longDatosCedula(2)))      'verif3
   conductor(3) = CDbl(Trim(Mid(datoCed, 49, longDatosCedula(3))))      'cedula
   conductor(4) = Trim(Mid(datoCed, 59, longDatosCedula(4)))      'primer apellido
   conductor(5) = Trim(Mid(datoCed, 82, longDatosCedula(5)))      'segundo apellido
   conductor(6) = Trim(Mid(datoCed, 105, longDatosCedula(6)))     'primer nombre
   conductor(7) = Trim(Mid(datoCed, 128, longDatosCedula(7)))     'segundo nombre
   conductor(8) = Trim(Mid(datoCed, 152, longDatosCedula(8)))     'sexo
   conductor(9) = Trim(Mid(datoCed, 153, longDatosCedula(9)))     'fecha nacimiento
   conductor(9) = Mid(conductor(9), 5, 2) & "/" & Mid(conductor(9), 7, 2) & "/" & Mid(conductor(9), 1, 4)
   conductor(10) = Trim(Mid(datoCed, 161, longDatosCedula(10)))   'verif4
   conductor(11) = Trim(Mid(datoCed, 167, longDatosCedula(11)))   'rh - tipo sangre
   conductor(12) = Trim(Mid(datoCed, 169, longDatosCedula(12)))   'verif5

   Exit Sub
'Err_leerCedula:
'   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub


Private Function Traer_Novedades(ByVal fPlaca As String) As Boolean
   fqry02 = "SELECT dvobservaciones " _
             & "FROM opdetveh " _
            & "WHERE dvplaca = '" & fPlaca & "' " _
              & "AND dvestado = 'A' "
           
   Set Ftb02 = Dbs.Execute(fqry02)
      
   TextNove = ""
   Traer_Novedades = True
   Do While Not Ftb02.EOF
      If TextNove = "" Then
         TextNove = Ftb02!dvobservaciones
      Else
         TextNove = TextNove & ", " & Chr(13) & Ftb02!dvobservaciones
      End If
      
      Ftb02.MoveNext
      Traer_Novedades = False
   Loop
   Ftb02.Close
      
   If TextNove <> "" Then
      MsgBox TextNove, 48, "Información"
   End If
   
End Function

    Private Sub Validar_Cupo(ByVal fCodigoEmpresa As String, ByVal fPlaca As String)
   On Error GoTo Err_Validar_Cupo
   If Trim(fCodigoEmpresa) <> "-" Then
      QRY = "SELECT * " _
            & "FROM cocupos " _
           & "WHERE cpempresa = '" & fCodigoEmpresa & "' "
      
      
   End If
      
   If fPlaca <> "-" Then
      QRY = "SELECT * " _
            & "FROM cocupos INNER JOIN cocuposdet ON cpnumcup = cdnumcupo " _
           & "WHERE cdplaca = '" & fPlaca & "' " _
             & "AND cpestado = 'A' " _
             & "AND cdestado = 'A'"
            
            'cpcedula = '" & fCedula & "' AND
      L_Numero_Cupo = 0
   End If
   
   OptTipo(0).Enabled = True
   OptTipo(1).Enabled = True
   
   If Trim(Mskcambio.text) = "" Then
      Mskcambio.text = "0"
   End If
         
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
      If Tbl1!cpestado = "A" Then
         Lblcupo.Caption = "CUPO"
         If Tbl1!cpsaldo > CDbl(Mskcambio.text) Then
            
            Lblcupo.Caption = "CUPO"
            L_Numero_Cupo = Tbl1!cpnumcup
            OptTipo(0).Enabled = False
            OptTipo(1).Enabled = True
            OptTipo(1).Value = True
            
            If Tbl1!cpsaldo <= Tbl1!cpminimo Then
               MsgBox "" & Chr(13) & "EL SALDO ESTA MENOR QUE EL VALOR MINIMO DEL CUPO " & Chr(13) & "", 48, "Advertencia"
            End If
            
            If Trim(fCodigoEmpresa) <> "-" Then
               L_Tipo_Cupo = "E"
            End If
            If fPlaca <> "-" Then
               L_Tipo_Cupo = "P"
            End If
            
            If Tbl1!cpterminal <> "G" Then
               If Tbl1!cpterminal <> Mid(G_Tipo_Terminal, 1, 1) Then
                  L_Numero_Cupo = 0
                  L_Tipo_Cupo = ""
                  OptTipo(0).Value = True
                  OptTipo(1).Enabled = False
                  Lblcupo.Caption = "LINEA"
               Else
                  If Mid(G_Tipo_Terminal, 1, 1) = "N" And influencia = "No" Then
                     L_Numero_Cupo = 0
                     L_Tipo_Cupo = ""
                     OptTipo(0).Value = True
                     OptTipo(1).Enabled = False
                     Lblcupo.Caption = "LINEA"
                  Else
                     If Mid(G_Tipo_Terminal, 1, 1) = "N" And Txtoritra.text = "T" Then
                        L_Numero_Cupo = 0
                        L_Tipo_Cupo = ""
                        OptTipo(0).Value = True
                        OptTipo(1).Enabled = False
                        Lblcupo.Caption = "LINEA"
                     End If
                  End If
               End If
            End If
            
         Else
            If Lblcupo.Caption = "CUPO" Then
               Lblcupo.Caption = "LINEA"
               G_Saldo_Cupo = Tbl1!cpsaldo
               On Error Resume Next
               AviCupos.Show 1
               On Error GoTo Err_Validar_Cupo
               OptTipo(1).Enabled = False
               OptTipo(0).Value = True
               
               'MsgBox "El cupo no tiene saldo para comprar la tasa de uso," & Chr(13) & _
                      "se le vendera normalmente, el saldo es: " & Format(Tbl1!cpsaldo, "$#,###0"), 48, "Advertencia"
               
               L_Tipo_Cupo = ""
               L_Numero_Cupo = 0
            End If
         End If
         
      Else
        Lblcupo.Caption = "LINEA"
      
         If Lblcupo.Caption <> "CUPO" Then
            OptTipo(1).Enabled = False
            OptTipo(0).Value = True
            L_Tipo_Cupo = ""
            L_Numero_Cupo = 0
         Else
            OptTipo(1).Enabled = True
            OptTipo(0).Value = False
         End If
      End If
   Else
      Lblcupo.Caption = "LINEA"
   
      If Lblcupo.Caption <> "CUPO" Then
         OptTipo(1).Enabled = False
         OptTipo(0).Value = True
         L_Tipo_Cupo = ""
         L_Numero_Cupo = 0
'      Else
'         OptTipo(1).Enabled = False
'            OptTipo(0).Value = True
      End If
   End If
   
   Exit Sub
Err_Validar_Cupo:
   
   Exit Sub
   
End Sub

Private Sub Validar_Despacho_Multas(ByVal fPlaca As String, ByVal fHora1 As String)
    L_Bandera_Rodamiento = False

    sFecha = Format(Mskfecha.text, "MM/dd/yyyy")
    fHora1 = sFecha & " " & fHora1
    sHora2 = Format(DateAdd("n", L_Minutos_Despacho, fHora1), "MM/dd/yyyy HH:mm")
    sHora1 = Format(DateAdd("n", -L_Minutos_Despacho, fHora1), "MM/dd/yyyy HH:mm")
    
    QRY = "SELECT denumero, denumalistamiento, decodrut, decedula1, decedula2, TO_CHAR(dehora,'HH24:MI') AS dehora, decodemp, deoritra, TO_CHAR(defecha,'MM/dd/yyyy') AS defecha " _
          & "FROM codespachos " _
         & "WHERE deplaca = '" & fPlaca & "' " _
           & "AND defecha >= TO_DATE('" & sFecha & "','MM/dd/yyyy') " _
           & "AND dehora BETWEEN TO_DATE('" & sHora1 & "','MM/dd/yyyy HH24:MI') " _
           & "AND TO_DATE('" & sHora2 & "','MM/dd/yyyy HH24:MI') " _
           & "AND deestado = 'D' AND deoritra = 'O' " _
           & "ORDER BY defecha ASC, dehora "
        
    Set Tbl10 = Dbs.Execute(QRY)
    
    If Not Tbl10.EOF Then
        L_Bandera_Rodamiento = True
    End If
    
    Tbl10.Close
End Sub
   

'WECC 05/23/2016 BUSCA EL DESPACHO CUANDO SE VA A GRABAR SE LE CAMBIA DE ESTADO A T _
CON ESTE ESTADO EN RODAMIENTO YA NO SE VISUALIZA
Private Sub Validar_Despacho(ByVal fPlaca As String, ByVal fHora1 As String)
      On Error GoTo errHandler
      
L_Bandera_Rodamiento = False
   sFecha = Format(Mskfecha.text, "MM/dd/yyyy")
   fHora1 = sFecha & " " & fHora1
   sHora2 = Format(DateAdd("n", L_Minutos_Despacho, fHora1), "MM/dd/yyyy HH:mm")
   sHora1 = Format(DateAdd("n", -L_Minutos_Despacho, fHora1), "MM/dd/yyyy HH:mm")
   
   'qry = "SELECT denumero, denumalistamiento, decodrut, decedula1, decedula2, TO_CHAR(dehora,'HH24:MI') AS dehora " _
         & "FROM codespachos " _
        & "WHERE deplaca = '" & fPlaca & "' " _
          & "AND defecha >= TO_DATE('" & sFecha & "','MM/dd/yyyy') " _
          & "AND TO_DATE(TO_CHAR(dehora,'HH24:MI'),'HH24:MI') BETWEEN TO_DATE('" & sHora1 & "','HH24:MI') " _
          & "AND TO_DATE('" & sHora2 & "','HH24:MI') " _
          & "AND deestado = 'D' " _
          & "ORDER BY defecha ASC, dehora "
          
   'wecc
   QRY = "SELECT denumero, denumalistamiento, decodrut, decedula1, decedula2, TO_CHAR(dehora,'HH24:MI') AS dehora, decodemp, deoritra, TO_CHAR(defecha,'MM/dd/yyyy') AS defecha " _
         & "FROM codespachos " _
        & "WHERE deplaca = '" & fPlaca & "' " _
          & "AND defecha >= TO_DATE('" & sFecha & "','MM/dd/yyyy') " _
          & "AND dehora BETWEEN TO_DATE('" & sHora1 & "','MM/dd/yyyy HH24:MI') " _
          & "AND TO_DATE('" & sHora2 & "','MM/dd/yyyy HH24:MI') " _
          & "AND deestado = 'D' AND deoritra = 'O' " _
          & "ORDER BY defecha ASC, dehora "
          
          
   Set Tbl10 = Dbs.Execute(QRY)
   L_Numero_Despacho = 0
   Capconduce.Enabled = False
   
   If Not Tbl10.EOF Then
   L_Bandera_Rodamiento = True
      Txtcodrut.Enabled = False
      TxtNumAlistamiento.Enabled = False
      txtCedula.Enabled = False
      TxtCedula1.Enabled = False
      Txtoritra.Enabled = False
      
      L_Numero_Despacho = Tbl10!denumero
      Mskhora.text = Tbl10!dehora
      'Mskfecha.Text = Tbl10!defecha
      L_Hora_Despacho = Tbl10!dehora
      If Not IsNull(Tbl10!decodrut) Then
         Txtcodrut.text = Tbl10!decodrut
      End If
      If Not IsNull(Tbl10!denumalistamiento) Then
         TxtNumAlistamiento.text = Tbl10!denumalistamiento
      End If
      If Not IsNull(Tbl10!decedula1) Then
         txtCedula.text = Tbl10!decedula1
         txtCedula.Locked = True
      End If
      If Not IsNull(Tbl10!decedula2) Then
         TxtCedula1.text = Tbl10!decedula2
         TxtCedula1.Locked = True
      End If
      
      If Tbl10!decodemp = Txtcodemp.text Then
         L_CodigoEmpresaConvenio = ""
      End If
      
      Txtoritra.text = Tbl10!deoritra
      
      Txtcodrut.Locked = True
      Txtoritra.Locked = True
      TxtNumAlistamiento.Locked = True
      Txtcodrut.Locked = True
      Txtcodrut.Enabled = True
      TxtNumAlistamiento.Enabled = True
      txtCedula.Enabled = True
      TxtCedula1.Enabled = True
      Txtoritra.Enabled = True
      
      'L_Bandera_Rodamiento = True
   End If
   Capconduce.Enabled = True
   Tbl10.Close
      Set Tbl10 = Nothing
   Exit Sub

errHandler:
   MsgBox "Error en Validar_Despacho: " & Err.Description, vbCritical
End Sub

'funcion valida el vehiculo
Private Function Valida_Placa_Conductor(ByVal sPlaca As String, ByVal sFecha1 As String, _
                                        ByVal sFecha2 As String, ByVal sTipo As String, _
                                        ByVal sTipoTasa As String) As Boolean
   Valida_Placa_Conductor = False
   
   If sTipoTasa = "O" Then
      QRY = "SELECT MAX(cohorsal), rutievia " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cohorsal BETWEEN to_date('" & sFecha1 & "','mm/dd/yyyy hh24:mi') " _
             & "AND to_date('" & sFecha2 & "','mm/dd/yyyy hh24:mi') "
             
      If sTipo = "P" Then
         QRY = QRY & "AND coplaca = '" & sPlaca & "' "
      ElseIf sTipo = "C" Then
         QRY = QRY & "AND '" & sPlaca & "' IN (cocedula,cocedula1) "
      End If
             
      QRY = QRY & "AND cooritra = 'O' " _
                & "AND coestado = 'Activo' " _
           & "GROUP BY rutievia order by max(cohorsal) desc "
      
      Set Ftb010 = Dbs.Execute(QRY)
   
      If Not Ftb010.EOF Then
         sTiempo_Ruta = ((Hour(Ftb010!rutievia) * 60) + Minute(Ftb010!rutievia))
         
         sfecsal1 = Mskfecha.text & " " & Mskhora.text
         sMinutosBuscar = sTiempo_Ruta * 2
         L_Tiempo_Ruta_Placa = sMinutosBuscar
         sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
         
         QRY = "SELECT cohorsal " _
               & "FROM coconduc " _
              & "WHERE cohorsal BETWEEN TO_DATE('" & sfecsal1 & "','MM/dd/yyyy hh24:mi') " _
                & "AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy hh24:mi') "
                
         If sTipo = "P" Then
            QRY = QRY & "AND coplaca = '" & sPlaca & "' "
         ElseIf sTipo = "C" Then
            QRY = QRY & "AND '" & sPlaca & "' IN (cocedula,cocedula1) "
         End If
                
                
         QRY = QRY & "AND cooritra = 'O' AND coestado = 'Activo' " _
                & "ORDER BY cohorsal DESC "
      
         Set Ftb010 = Dbs.Execute(QRY)
      
         If Not Ftb010.EOF Then
            L_Fecha_salida = Ftb010!cohorsal
            Ftb010.Close
            Exit Function
         End If
         Ftb010.Close
      End If
   Else
      QRY = "SELECT cohorsal " _
            & "FROM coconduc " _
           & "WHERE cohorsal BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy hh24:mi') " _
             & "AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy hh24:mi') "
             
      If sTipo = "P" Then
         QRY = QRY & "AND coplaca = '" & sPlaca & "' "
      ElseIf sTipo = "C" Then
         QRY = QRY & "AND '" & sPlaca & "' IN (cocedula,cocedula1) "
      End If
             
             
      QRY = QRY & "AND cooritra = 'T' AND coestado = 'Activo' " _
             & "ORDER BY cohorsal DESC "
   
      Set Ftb010 = Dbs.Execute(QRY)
   
      If Not Ftb010.EOF Then
         L_Fecha_salida = Ftb010!cohorsal
         Ftb010.Close
         Exit Function
      End If
      Ftb010.Close
   End If
   Valida_Placa_Conductor = True
End Function



'FUNCION VALIDA COMPRA CONDUCTOR

'Private Function Valida_Conductor(ByVal sCedula As String, ByVal sFecha1 As String, ByVal sFecha2 As String) As Boolean
'
'   Valida_Conductor = False
''   Qry = "SELECT svfecha " _
''         & "FROM cosalveh " _
''        & "WHERE svfecha BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy hh24:mi') " _
''          & "AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy hh24:mi') " _
''          & "AND '" & sCedula & "' IN (svcedula,svcedula1) "
''
''   Set Ftb010 = Dbs.Execute(Qry)
''
''   If Not Ftb010.EOF Then
''      L_Fecha_salida = Ftb010!svfecha
''      Ftb010.Close
''      Exit Function
''   End If
''   Ftb010.Close
'
'   Qry = "SELECT cohorsal " _
'         & "FROM coconduc " _
'        & "WHERE cohorsal BETWEEN TO_DATE('" & sFecha1 & "','MM/dd/yyyy hh24:mi') " _
'          & "AND TO_DATE('" & sFecha2 & "','MM/dd/yyyy hh24:mi') " _
'          & "AND '" & sCedula & "' IN (cocedula,cocedula1) " _
'          & "AND cooritra = 'O' AND coestado = 'Activo' " _
'          & "ORDER BY cohorsal DESC "
'
'   Set Ftb010 = Dbs.Execute(Qry)
'
'   If Not Ftb010.EOF Then
'      L_Fecha_salida = Ftb010!cohorsal
'      Ftb010.Close
'      Exit Function
'   End If
'   Ftb010.Close
'
'
'
'   Valida_Conductor = True
'
'
'
'End Function


'*******************************************************************

'Private Sub btnalcoholemia_Click()
'   On Error GoTo err_btnalcoholemia_Click
'   If btnalcoholemia.Caption = "&Iniciar" Then
'      btnalcoholemia.Caption = "&Finalizar"
'      btnalcoholemia.Picture = LoadPicture(DIRECCION & "Trffc10c.ico")
'   ElseIf btnalcoholemia.Caption = "&Finalizar" Then
'      btnalcoholemia.Caption = "&Iniciar"
'      btnalcoholemia.Picture = LoadPicture(DIRECCION & "Trffc10a.ico")
'   End If
   
'   Exit Sub
   
'err_btnalcoholemia_Click:
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Sub
'
'End Sub

'Private Function FnValidaHorario() As Boolean
'   On Error GoTo Errores
'   FnValidaHorario = False
'
'   Qry = "SELECT rhestado  " _
'         & "FROM ophoremp " _
'        & "WHERE rhcodemp = '" & Trim(Txtcodemp.Text) & "' " _
'          & "AND rhcodrut = '" & Trim(Txtcodrut.Text) & "' " _
'          & "AND rhservic = '" & Trim(Txttipveh.Text) & "' "
'
'          '& "AND rhhora = '" & Format(Mskhora.Text, "hh:mm") & "' " _
'          & "AND rhestado = 'Activa' "
'
'   Fn_CalcularRangoHoras (Format(Mskhora.Text, "HH:mm"))
'   If Len(Trim(L_HoraInicial)) > 0 And Len(Trim(L_HoraFinal)) > 0 Then
'        Qry = Qry & "AND rhhora between '" & Format(L_HoraInicial, "hh:mm") & "' AND '" & Format(L_HoraFinal, "hh:mm") & "' "
'   Else
'        Qry = Qry & "AND rhhora = '" & Format(Mskhora.Text, "hh:mm") & "' " _
'
'   End If
'   Qry = Qry & "AND rhestado = 'Activa' "
'   Set Tbl9 = Dbs.Execute(Qry)
'
'   If Not Tbl9.EOF Then
'      FnValidaHorario = True
'   End If
'   Tbl9.Close
'   Exit Function
'
'Errores:
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Function
'
'End Function

Private Function FnValidaHorario(ByVal fTipo_Venta As String) As Boolean
   On Error GoTo Errores
   FnValidaHorario = False
   
   If Fn_ValidarRuta = False Then
        FnValidaHorario = True 'Ruta que No se de debe controlar horario
        Exit Function
   End If
   'variable para determinar la cantidad de veces en el for si es operador y empresa es una vez
   'si es mixta dos veces para validar por las dos opciones si existe en la primera se sale del for
   sVeces = 1
   If fTipo_Venta = "M" Then
      If L_Operador <> "" Then
         sVeces = 2
      End If
   End If
   
   For i = 1 To sVeces
   
      QRY = "SELECT rhimpresion, rhhora, NVL(rhvalida,'N') AS rhvalida " _
               & "FROM opruthor " _
              & "WHERE rhcodrut = '" & Trim(Txtcodrut.text) & "' "
              
      If fTipo_Venta = "E" Then
         QRY = QRY & "AND rhcodemp = '" & Trim(Txtcodemp.text) & "' "
   
      ElseIf fTipo_Venta = "O" Then
         QRY = QRY & "AND rhoperador = '" & L_Operador & "' "
      ElseIf fTipo_Venta = "M" Then
         If i = 1 Then
            QRY = QRY & "AND rhcodemp = '" & Trim(Txtcodemp.text) & "' "
         Else
            QRY = QRY & "AND rhoperador = '" & L_Operador & "' "
         End If
      End If
      
      If G_ValidaServicio = "S" Then
         QRY = QRY & "AND rhservic = '" & Trim(Txttipveh.text) & "' "
      End If
      
      QRY1 = QRY
      
     '************************************************************************'
     'Fn_CalcularRangoHoras (Mskhora.Text)                                    '
     '------------------------------------------------------------------------'
     ' Funcion que con la hora actual determina el Rango de Horas actualizando'
     ' dos variables L_HoraInicial y L_HoraFinal permitidas para vender       '
     ' las tasas de Uso                                                       '
     ' Fecha: Abril 26 de 2010                                                '
     ' Desarrollador: Miller Dorado                                           '
     '------------------------------------------------------------------------'
      
   '   Fn_CalcularRangoHoras (Format(Mskhora.Text, "HH:mm"))
   '   If Len(Trim(L_HoraInicial)) > 0 And Len(Trim(L_HoraFinal)) > 0 Then
   '        Qry = Qry & "AND rhhora between '" & Format(L_HoraInicial, "hh:mm") & "' AND '" & Format(L_HoraFinal, "hh:mm") & "' "
   '   Else
   '        Qry = Qry & "AND rhhora = '" & Format(Mskhora.Text, "hh:mm") & "' " _
   '
   '   End If
      
      'WECC 07/22/2010
      'CUANDO LA HORA DIGITADA ES MAYOR A LA HORA DEL SISTEMA SE HACE UNA IGUALDAD
      'CUANDO LA HORA DIGITADA ES MENOR A LA HORA DEL SISTEMA SE HACE UNA BETWEEN
      '**********************************
      sHora_Digitada = Format(Mskhora.text, "HH:mm")
      sHora_Sistema = Format(Mskhorsis.text, "HH:mm")
      
'      If sHora_Digitada >= sHora_Sistema Then
'         Qry = Qry & "AND rhhora = '" & Format(sHora_Digitada, "hh:mm") & "' "
'      Else
'         L_HoraInicial1 = Format(DateAdd("n", -G_MinHorTU, sHora_Sistema), "HH:mm")
'         L_HoraFinal1 = Format(sHora_Sistema, "HH:mm")
'
'         If sHora_Digitada >= L_HoraInicial1 And sHora_Digitada <= L_HoraFinal1 Then
'            L_HoraInicial = Format(DateAdd("n", -G_MinHorTU, sHora_Digitada), "HH:mm")
'            L_HoraFinal = Format(sHora_Digitada, "HH:mm")
'
'            'qry = qry & "AND rhhora between '" & Format(L_HoraInicial, "hh:mm") & "' AND '" & Format(L_HoraFinal, "hh:mm") & "' "
'            Qry = Qry & "AND rhhora = '" & Format(sHora_Digitada, "hh:mm") & "' "
'         Else
'            Exit Function
'         End If
'      End If
      '**********************************
      
      QRY = QRY & "AND rhestado = 'Activa' ORDER BY rhhora "
      
      Set tbl9 = Dbs.Execute(QRY)
      G_Imprime_Tasa_Uso = ""
      
      '---------------------------------------
      G_Imprime_Tasa_Uso = "E"
      L_HorasPermitidas = ""
      Do While Not tbl9.EOF
         If L_HorasPermitidas = "" Then
            L_HorasPermitidas = tbl9!rhhora
         Else
            L_HorasPermitidas = L_HorasPermitidas & ", " & tbl9!rhhora
         End If
         
         If tbl9!rhhora = sHora_Digitada Then
            If Not IsNull(tbl9!rhimpresion) Then
               G_Imprime_Tasa_Uso = tbl9!rhimpresion
            End If
            L_Valida_Control_Ruta = tbl9!rhvalida
            FnValidaHorario = True
            tbl9.Close
            Exit Function
         End If
         tbl9.MoveNext
      Loop
      '----------------------------------------
      
'      If Not Tbl9.EOF Then
'         G_Imprime_Tasa_Uso = "E"
'         If Not IsNull(Tbl9!rhimpresion) Then
'            G_Imprime_Tasa_Uso = Tbl9!rhimpresion
'         End If
'         FnValidaHorario = True
'         Tbl9.Close
'         Exit Function
'      Else
'         Tbl9.Close
'         Qry1 = Qry1 & "AND rhestado = 'Activa' "
'
'         Set Tbl9 = Dbs.Execute(Qry1)
'         L_HorasPermitidas = ""
'         Do While Not Tbl9.EOF
'            If L_HorasPermitidas = "" Then
'               L_HorasPermitidas = Tbl9!rhhora
'            Else
'               L_HorasPermitidas = L_HorasPermitidas & ", " & Tbl9!rhhora
'            End If
'            Tbl9.MoveNext
'         Loop
'         Tbl9.Close
'      End If
      
   Next i
   Exit Function
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function

End Function

Public Function Fn_ValidarRuta() As Boolean
  Dim sRespuesta As Boolean
  sRespuesta = False
  If Txtcodrut.text <> "" Then
     fqry01 = "SELECT RUVALHOR " _
              & "FROM opruta " _
             & "WHERE rucodigo = '" & Txtcodrut.text & "'"
             
     Set Ftb01 = Dbs.Execute(fqry01)
     L_Valida_Horario = "N"
     If Not Ftb01.EOF Then
        If Ftb01!RUVALHOR = "S" Then
           sRespuesta = True
           L_Valida_Horario = "S"
        Else
           sRespuesta = False
        End If
     End If
  End If
  Fn_ValidarRuta = sRespuesta
End Function

Public Function Fn_OperadorEmpresa(ByVal sOperador As String, ByVal sCodEmpresa As String) As Boolean
   On Error GoTo Err_Fn_OperadorEmpresa
   Fn_OperadorEmpresa = False
   
   QRY = "SELECT oeoperador " _
         & "FROM opopeemp " _
        & "WHERE oeoperador ='" & sOperador & "' " _
          & "AND oecodemp = '" & sCodEmpresa & "' " _
          & "AND oeestado = 'A' "
   
   Set tbl9 = Dbs.Execute(QRY)
   
   If Not tbl9.EOF Then
      Fn_OperadorEmpresa = True
      
   End If
   tbl9.Close
   Exit Function
Err_Fn_OperadorEmpresa:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Function


Private Sub Valida_Terminal_Central(vTerminal As String)
   On Error GoTo Errores
   
   G_encontro_despacho = False
   
   QRY = "SELECT svnumero, svcodrut " _
         & "FROM cosalveh " _
        & "WHERE svplaca = '" & Trim(Txtnumpla.text) & "' " _
        & "AND svterminal = 'C'"
   
   'wecc 03/07/2012 va quemado C porque valida la ruta de la central
          
   If vTerminal = "RUTA" Then
      QRY = QRY & "AND svcodrut = '" & Trim(Txtcodrut.text) & "'"
   End If
   Set Tbl1 = Dbs.Execute(QRY)
   If L_Numero_Despacho = 0 Then
      Txtoritra.Locked = False
   End If
   
   PasajerosOrigen = 0
   If Not Tbl1.EOF Then
      If vTerminal = "PLACA" Then
         L_Bandera_Ruta_Lector = Tbl1!svcodrut
         Txtcodrut.text = Tbl1!svcodrut
         sNumeroTasa = Tbl1(0)
         Tbl1.Close
         sFecha2 = Format(Mskfecha.text, "MM/dd/yyyy")
         sFecha1 = DateAdd("d", -1, sFecha2)
         QRY = "SELECT cocedula, cocedula1, conumalistamiento, cocodemp, copasaje " _
               & "FROM coconduc " _
              & "WHERE conumero = " & sNumeroTasa & " " _
                & "AND cofecsal BETWEEN TO_DATE('" & sFecha1 & "','MM/DD/YYYY') " _
                & "AND TO_DATE('" & sFecha2 & "','MM/DD/YYYY')  " _
                & "AND coplaca = '" & Trim(Txtnumpla.text) & "' " _
                & "AND cocodrut = '" & Txtcodrut.text & "' AND coestado = 'Activo' "
              
         Set Tbl1 = Dbs.Execute(QRY)
         
         If Not Tbl1.EOF Then
         
            G_encontro_despacho = True
            TxtNumAlistamiento.text = ""
            txtCedula.text = ""
            TxtCedula1.text = ""
            L_Pasajeros_Entran = Tbl1!copasaje
            If Not IsNull(Tbl1!conumalistamiento) Then
               TxtNumAlistamiento.text = Tbl1!conumalistamiento
            End If
            
            If Not IsNull(Tbl1!cocedula) Then
               txtCedula.text = Tbl1!cocedula
            End If
            
            If Not IsNull(Tbl1!copasaje) Then
               PasajerosOrigen = Tbl1!copasaje
            Else
               PasajerosOrigen = 0
            End If
            
            If Not IsNull(Tbl1!cocedula1) Then
               TxtCedula1.text = Tbl1!cocedula1
            End If
            If Trim(Txtcodemp.text) = Trim(Tbl1!cocodemp) Then
               L_CodigoEmpresaConvenio = ""
            End If
         End If
         Tbl1.Close
         TxtNumAlistamiento.Enabled = False
         txtcedula_LostFocus
         txtcedula1_LostFocus
      End If
      Txtoritra.text = "T"
   Else
      'Txtoritra.Locked = True  'mayo 19 2009
      Txtoritra.text = "O"
      Tbl1.Close
   End If
   
      
      
   
   
   
'   sHoraSis = Format(Fn_Parametros("HORASIS", 1), "MM/dd/yyyy HH:mm")
'   shoracentral = Format(DateAdd("n", -L_MinCenSur, sHoraSis), "MM/dd/yyyy HH:mm")
'
'   qry = "SELECT conumero " _
'         & "FROM coconduc " _
'        & "WHERE cohorsal BETWEEN TO_DATE('" & shoracentral & "','MM/dd/yyyy hh24:mi') " _
'          & "AND TO_DATE('" & sHoraSis & "','MM/dd/yyyy hh24:mi') " _
'          & "AND coplaca = '" & Txtnumpla.Text & "' " _
'          & "AND cocodrut = '" & Txtcodrut.Text & "' " _
'          & "AND coestado = 'Activo' "
'
'   Set Tbl1 = Dbs.Execute(Qry)
'   Txtoritra.Locked = True
'   If Not tbl1.EOF Then
'      Txtoritra.Text = "T"
'   Else
'      Txtoritra.Text = "O"
'   End If
'   tbl1.Close
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Traer_Tasas_Terminal_Norte(ByVal sPlaca As String)
   QRY = "SELECT tnnumero, tnvalor, tnvalalc, tncodrut, tncaja, tncedula1, tncedula2, tndireccion, " _
              & "TO_CHAR(tnfechaentradanorte,'MM/DD/YYYY') as tnfechaentradanorte, TO_CHAR(tnhoraentradanorte,'HH24:MI') as tnhoraentradanorte, TO_CHAR(tnfechasalidanorte,'MM/DD/YYYY') as tnfechasalidanorte, TO_CHAR(tnhorasalidanorte, 'HH24:MI') as tnhorasalidanorte, " _
              & "TO_CHAR(tnfechasalida,'MM/DD/YYYY') as tnfechasalida, TO_CHAR(tnhorasalida,'HH24:MI') as tnhorasalida, tnterminal " _
         & "FROM cotransitonorte " _
        & "WHERE tnplaca = '" & sPlaca & "' " _
          & "AND tnestado = 'P' "
          
   Set Tbl2 = Dbs.Execute(QRY)
   
   L_Numero_Terminal_Norte = 0
   L_Valor_Alcoholimetria_Norte = 0
   L_Cedula2Norte = ""
   L_NombreConductor2Norte = ""
   If Not Tbl2.EOF Then
      L_Numero_Terminal_Norte = Tbl2!tnnumero
      MskValorTasaNorte.text = Tbl2!tnvalor
      L_Valor_Alcoholimetria_Norte = Tbl2!tnvalalc
      L_Ruta_Transito_Norte = Tbl2!tncodrut
      L_Caja_Norte = Tbl2!tncaja
      L_DireccionNorte = Trim(Tbl2!tndireccion)
      L_Terminal_Tasa_Automatica = Tbl2!tnterminal
      L_Cedula1Norte = Trim(Tbl2!tncedula1)
      L_NombreConductor1Norte = Traer_Nombre_Conductor(L_Cedula1Norte)
      
      If Not IsNull(Tbl2!tncedula2) Then
         L_Cedula2Norte = Trim(Tbl2!tncedula2)
         L_NombreConductor2Norte = Traer_Nombre_Conductor(L_Cedula2Norte)
      End If
      
      If Not IsNull(Tbl2!tnfechaentradanorte) Then
         L_FechaEntradaNorte = Tbl2!tnfechaentradanorte
      Else
         L_FechaEntradaNorte = Tbl2!tnfechasalida
      End If
      
      If Not IsNull(Tbl2!tnhoraentradanorte) Then
         L_HoraEntradaNorte = Tbl2!tnhoraentradanorte
      Else
         L_HoraEntradaNorte = Tbl2!tnhorasalida
      End If
      
      
      If Not IsNull(Tbl2!tnfechasalidanorte) Then
         L_FechaSalidaNorte = Tbl2!tnfechasalidanorte
      Else
         L_FechaSalidaNorte = Tbl2!tnfechasalida
      End If
      
      If Not IsNull(Tbl2!tnhorasalidanorte) Then
         L_HoraSalidaNorte = Tbl2!tnhorasalidanorte
      Else
         L_HoraSalidaNorte = Tbl2!tnhorasalida
      End If
      
      
      LCodigoRutaNorteIntegra = Fn_Equivalencia_Ruta(L_Ruta_Transito_Norte)
      If Trim(LCodigoRutaNorteIntegra) = "" Then
         MsgBox "El código de la ruta transito no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   Tbl2.Close
      
End Sub

Private Function EscogerVehiculo() As Boolean
   EscogerVehiculo = False
      
   If ESTVEH = "No" Then
      
      If Val(ContAlco) = Val(NumCar) Then
         If L_Bandera_Escoger = "SI" Then
            presentar = 0
            ContAlco = 0
            L_Bandera_Escoger = "NO"
          Else
            If ContAlco <> 0 Then
               presentar = 1
            End If
         End If
      Else
         If Val(NumCar) <> 0 Then
           presentar = 1
           If L_Bandera_Escoger = "SI" Then
               ContAlco = ContAlco + 1
               L_Bandera_Escoger = "NO"
            End If
         Else
            presentar = 0
            ContAlco = 0
         End If
      End If
      alcoholemia = "S"
      If presentar = 0 Then
         
         'MsgBox "SE DEBE DE PRESENTAR AL EXAMEN DE ALCOHOLEMIA", 48, "INFORMACION"
  
         Dbs.Execute "UPDATE ophojveh " _
                      & "SET  hvestsel = 'Si' " _
                    & "WHERE hvplaca = '" & Txtnumpla & "'"
         
         ESTVEH = "Si"

         EscogerVehiculo = True
         
      Else
         EscogerVehiculo = False
      End If
   Else
      EscogerVehiculo = True
   End If
   

End Function

Private Function Valida_Hora_Cierre_Facturacion() As Boolean
    'Sebastián Rondón - Enero 14 de 2025
   'Funcion para que se visualice un mensaje indicando si es hora de cerrar el turno u obligando el cierre de turno.
   On Error GoTo Err_Valida_Hora_Cierre_Facturacion
   Valida_Hora_Cierre_Facturacion = True
   If Val(G_Turno) = Val(G_maxtur) Then
      sHora = Format(Mskhora.text, "hh:mm")
      horcierre1 = "23:59"
      horcierre2 = "23:55"
      horcierre3 = "00:01"
      horcierre4 = "05:00"
      
      If CDate(sHora) >= CDate(Format(horcierre2, "hh:mm")) And CDate(sHora) <= CDate(Format(horcierre1, "hh:mm")) Then
         MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
         Txtnumpla.text = ""
         Valida_Hora_Cierre_Facturacion = False
         Exit Function
      ElseIf CDate(sHora) >= CDate(Format(horcierre3, "hh:mm")) And CDate(sHora) <= CDate(Format(horcierre4, "hh:mm")) Then
         MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
         Txtnumpla.text = ""
         Valida_Hora_Cierre_Facturacion = False
         Exit Function
      End If
   End If
   
   Exit Function
Err_Valida_Hora_Cierre_Facturacion:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function
   Resume
End Function

Private Function Valida_Hora_Cierre_Turno() As Boolean
   'WECC Jul 24 2019   CAU2
   'Funcion para que se visualice un mensaje indicando si es hora de cerrar el turno u obligando el cierre de turno.
   On Error GoTo Err_Valida_Hora_Cierre_Turno
   Valida_Hora_Cierre_Turno = True
   QRY = "SELECT TO_CHAR(caturno1,'HH24:mi') AS caturno1, TO_CHAR(caturno2,'HH24:mi') AS caturno2, TO_CHAR(caturno3,'HH24:mi') AS caturno3, TO_CHAR(caturno4,'HH24:mi') AS caturno4 " _
        & "FROM cocaja " _
       & "WHERE canomcaj = '" & Pc & "'"
  
   Set Tbl1 = Dbs.Execute(QRY)
  
   If Not Tbl1.EOF Then
      sHora = Format(Mskhora.text, "hh:mm")
      sHoraTurno1 = Format(Tbl1!caturno1, "hh:mm")
      sHoraTurno2 = Format(Tbl1!caturno2, "hh:mm")
      sHoraTurno3 = Format(Tbl1!caturno3, "hh:mm")
      sHoraTurno4 = Format(Tbl1!caturno4, "hh:mm")
      
      sHoraTurno1_Add = Format(DateAdd("n", 120, sHoraTurno1), "hh:mm")
      sHoraTurno2_Add = Format(DateAdd("n", 120, sHoraTurno2), "hh:mm")
      sHoraTurno3_Add = Format(DateAdd("n", 120, sHoraTurno3), "hh:mm")
      sHoraTurno4_Add = Format(DateAdd("n", 120, sHoraTurno4), "hh:mm")
      
      If Val(G_Turno) = 1 Then
         horcierre1 = "15:01"
         horcierre2 = "23:59"
         
         If CDate(sHora) > CDate(sHoraTurno1) Then
            If CDate(sHora) > CDate(Format(horcierre1, "hh:mm")) And CDate(sHora) < CDate(Format(horcierre2, "hh:mm")) Then
               Exit Function
            End If
            If CDate(sHora) > CDate(sHoraTurno1_Add) Then
               MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
               Txtnumpla.text = ""
               Valida_Hora_Cierre_Turno = False
               Exit Function
            End If
            MsgBox "Es Hora de Cerrar Turno ", 48, "Advertencia"
            Valida_Hora_Cierre_Turno = False
            Exit Function
         End If
      ElseIf Val(G_Turno) = 2 Then
         If CDate(sHora) > CDate(sHoraTurno2) Then
         
            If CDate(sHora) > CDate(sHoraTurno2_Add) Then
               MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
               Txtnumpla.text = ""
               Valida_Hora_Cierre_Turno = False
               Exit Function
            End If
         
            MsgBox "Es Hora de Cerrar Turno ", 48, "Advertencia"
            Valida_Hora_Cierre_Turno = False
            Exit Function
         End If
      ElseIf Val(G_Turno) = 3 Then
                 
         If CDate(sHora) > CDate(sHoraTurno3) Then
            
            If CDate(sHora) > CDate(sHoraTurno3_Add) Then
               MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
               Txtnumpla.text = ""
               Valida_Hora_Cierre_Turno = False
               Exit Function
            End If
            MsgBox "Es Hora de Cerrar Turno ", 48, "Advertencia"
            Valida_Hora_Cierre_Turno = False
            Exit Function
         End If
      ElseIf Val(G_Turno) = 4 Then
                 
         horcierre1 = "13:01"
         horcierre2 = "00:00"
         
         If CDate(sHora) >= CDate(Format(horcierre2, "hh:mm")) And CDate(sHora) <= CDate(Format(horcierre1, "hh:mm")) Then
            MsgBox "Tiene Que Cerrar Turno ", 48, "Advertencia"
            Txtnumpla.text = ""
            Valida_Hora_Cierre_Turno = False
            Exit Function
         End If
      End If
   End If
   Tbl1.Close
   Exit Function
Err_Valida_Hora_Cierre_Turno:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function
   Resume
End Function


Private Sub TraerValorPermanenciasMultas()
'   QRY = "SELECT temnumero1 AS vrPermanencia, temnumero2 AS vrMultas, temnumero3 AS numMulta " _
'         & "FROM optempo1 " _
'        & "WHERE temtext4 = '" & Trim(Txtnumpla.text) & "' " _
'          & "AND tempc = '" & Trim(Pc) & "'"
'
'   Set Tbl10 = Dbs.Execute(QRY)
'   Mskparqueadero.text = "0"
'   MskMultas.text = "0"
'   If Not Tbl10.EOF Then
'        Do While Tbl10.EOF = False
'            Mskparqueadero.text = CDbl(Mskparqueadero.text) + CDbl(Tbl10!vrPermanencia)
'            MskMultas.text = CDbl(MskMultas.text) + CDbl(Tbl10!vrMultas)
'            Tbl10.MoveNext
'        Loop
'   End If
   
   'Felipe Mendoza
   'Marzo 12/2024
   'Se genera una nueva validación para evitar que se haga pagos en una caja distinta
   QRY = "SELECT rcvalper AS vrPermanencia, rcvalmul AS vrMultas, rcnumero AS numMulta, rcnomcaj " _
            & "FROM opreccaja " _
            & "INNER JOIN cocaja " _
            & "ON opreccaja.rcnomcaj = cocaja.canomcaj " _
            & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
            & "AND rcnomcaj <> '" & Trim(Pc) & "' " _
            & "AND rcfecha = TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY') " _
            & "AND catipo = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
      
   Set Tbl10 = Dbs.Execute(QRY)
   If Not Tbl10.EOF Then
        MsgBox "El vehículo tiene pagos en otra caja (" & Tbl10!rcnomcaj & "), no es posible la venta", 48, "Advertencia"
        Btnlimpiar_Click
        Exit Sub
   End If
   
   QRY = "SELECT rcvalper AS vrPermanencia, rcvalmul AS vrMultas, rcnumero AS numMulta " _
            & "FROM opreccaja " _
            & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
            & "AND rcnomcaj = '" & Trim(Pc) & "' " _
            & "AND rcfecha = TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY')"
      
   Set Tbl10 = Dbs.Execute(QRY)
   MskParqueadero.text = "0"
   MskMultas.text = "0"
   If Not Tbl10.EOF Then
        Do While Tbl10.EOF = False
            MskParqueadero.text = CDbl(MskParqueadero.text) + CDbl(Tbl10!vrPermanencia)
            MskMultas.text = CDbl(MskMultas.text) + CDbl(Tbl10!vrMultas)
            Tbl10.MoveNext
        Loop
   End If
   
   's_Num_Multa_Pagada = 0
   'If Not Tbl10.EOF Then
      'Mskparqueadero.text = Tbl10!vrPermanencia
      'MskMultas.text = Tbl10!vrMultas
      's_Num_Multa_Pagada = Tbl10!numMulta
   'End If
   Tbl10.Close
End Sub

Private Function ControlarCompra_TU(ByVal fCodigoRuta As String, ByVal fCodigoEmpresa As String, _
                                      ByVal fFecha1 As String, ByVal fValida As String) As Boolean
   'WECC Nov/25/2019 Caso CAU2 No  1907
   'Se busca si existe una relacion de ser correcto se realiza la busqueda para la ruta _
   dentro de un intervalo de hora según el descuento de minutos que tenga el parametro MINCONRUTA _
   con la la hora de salida digitada
   On Error GoTo Error_FnControlarCompra_TU
   ControlarCompra_TU = True
   
   If fValida = "S" Then
      QRY = "SELECT * " _
            & "FROM oprutasadicionales " _
           & "WHERE racodemp = '" & fCodigoEmpresa & "' " _
             & "AND racodrutadicional = '" & fCodigoRuta & "' " _
             & "AND raestado = 'A' "
       
      Set Tbl10 = Dbs.Execute(QRY)
      
      sRutaOriginal = ""
      If Not Tbl10.EOF Then
         sRutaOriginal = Tbl10!racodrutoriginal
      End If
      
      
      If sRutaOriginal <> "" Then
         sFechaHora1 = Format(DateAdd("n", -L_Minutos_Controlar_Rutas, fFecha1), "MM/dd/yyyy HH:mm")
         sFechaHora2 = Format(fFecha1, "MM/dd/yyyy HH:mm")
         QRY = "SELECT coplaca " _
               & "FROM coconduc " _
              & "WHERE cohorsal BETWEEN TO_DATE('" & sFechaHora1 & "','MM/dd/yyyy HH24:MI') " _
                & "AND TO_DATE('" & sFechaHora2 & "','MM/dd/yyyy HH24:MI') " _
                & "AND cocodemp = '" & fCodigoEmpresa & "' AND cocodrut = '" & sRutaOriginal & "' " _
                & "AND coestado = 'Activo'"
         
         Set Tbl10 = Dbs.Execute(QRY)
         
         If Tbl10.EOF Then
            ControlarCompra_TU = False
         End If
         Tbl10.Close
         
      End If
   End If
   Exit Function
Error_FnControlarCompra_TU:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function
End Function

Private Sub BuscacambioServicio(ByVal fCodEmp As String, ByVal fCodRut As String)
   fqry = "SELECT erstipser " _
          & "FROM opemprutser " _
         & "WHERE erscodemp = '" & fCodEmp & "' " _
           & "AND erscodrut = '" & fCodRut & "' " _
           & "AND ersestado = 'A'"
   
   
   Set Tbl10 = Dbs.Execute(fqry)
   
   If Not Tbl10.EOF Then
      Txttipveh.text = Tbl10!erstipser
      Tbl10.Close
      Txttipveh_LostFocus
   Else
      Tbl10.Close
   End If
   
   
End Sub


Private Function Fn_PicoyPlaca(ByVal fFecha As String, ByVal fDigito As Integer) As Boolean
   Fn_PicoyPlaca = False
   
   QRY = "SELECT ppconsecutivo " _
         & "FROM oppicoyplaca " _
        & "WHERE ppfecha = TO_DATE('" & fFecha & "','MM/dd/yyyy') " _
          & "AND ppplaca = " & fDigito & " " _
           & "AND ppestado = 'A'"
      
   Set Tbl10 = Dbs.Execute(QRY)
   
   If Not Tbl10.EOF Then
      Fn_PicoyPlaca = True
   End If
   Tbl10.Close

End Function

Private Sub BtnBuscar_Click()
'  forma = "c1"
'  forma1 = "c1"
'  Txtcodemp.Text = ""
'  Buscar.Show 1
'  Txtcodemp.SetFocus
'  txtcodemp_LostFocus


   G_Clase_Autorizacion = "REIMPTASA"
   Relogin.Show 1
   If G_Aut_Vender <> "SI" Then
      Exit Sub
   End If
   Btnimprimir.Enabled = False
   Btnbuscar.Enabled = False
   Txtnumcon.text = ""
   Txtnumcon.Locked = False
   Txtnumpla.Enabled = False
   Txttipveh.Enabled = False
   Txtcodemp.Enabled = False
   Txtcodrut.Enabled = False
   Txtoritra.Enabled = False
   txtCedula.Enabled = False
   TxtCedula1.Enabled = False
   Txtconact.Enabled = False
   Txtconact1.Enabled = False
   txtapellido.Enabled = False
   Txtapellido1.Enabled = False
   Mskfecha.Enabled = False
   Mskhora.Enabled = False
   Txtnumcon.SetFocus
   'Sebastián Rondón Agosto 31 - 2022
   'Sentencia para que ya haya reimprimido una vez
   'La próxima le siga solicitando el usuario administrador
   G_Aut_Vender = "NO"
End Sub

Private Sub btncambio_Click()
   On Error GoTo Errores
   
   Unload Me
   camcondu.Show 1
   
   Exit Sub
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

Private Sub BtnGrabar_Click()

   Dim sSaldoCupo As Double
   Dim NumTasa As String
   Dim Placa As String
   Dim Tterminal As String
   
   On Error GoTo err_Btngrabar
   
   
   'Camilo Campos 06/05/2025 no se permite la venta en CAJATAQ debido a que es exclusiva de rodamiento
   If Pc = "CAJATAQ" Then
        MsgBox "No se permiten las ventas con CAJATAQ", 48, "Advertencia"
        Exit Sub
   End If
   
     'Camilo Campos 05/27/2025 debido a los ajustes hechos en rodamiento las tasas de uso por crédito/cupo deben pagarse en el mismo rodamiento
     'Por lo cual, en caso de que venga cualquier vehiculo perteneciente a credito/cupo el sistema no debe permitir el ingreso en conduces
            
      If G_valida_credito = "S" Then
            If OptTipo(1).Value = True Or chkcredito.Enabled = True Then
                
                MsgBox "El vehiculo pertenece a credito o cupo, debe ir directamente a rodamiento", 48, "Advertencia"
                Btnlimpiar_Click
                Exit Sub
            End If
      End If
    
    
   If txtcaja.text = "" Then
      MsgBox "Falta El Nombre de La Caja", 48, "Advertencia"
      Exit Sub
   End If
   Txtnumpla.Enabled = True
   If Trim(Txtnumpla.text) = "" Then
      Txtnumpla.SetFocus
      Exit Sub
   End If
   
   If Txtcodemp.text = "" Then
      MsgBox "Falta ingresar el codigo de la empresa", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   If Txtcodrut.text = "" Then
      MsgBox "Falta ingresar el codigo de la ruta", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   
   If Txtnumpla.text = "" Then
      MsgBox "Falta ingresar el número de la placa", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   If TxtNumInt.text = "" Then
      MsgBox "Falta ingresar el número interno del bus", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   If Txttipveh.text = "" Then
      MsgBox "Falta ingresar la clase del vehículo", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   
   If Trim(TxtNumAlistamiento.text) = "" Then
      MsgBox "Falta ingresar el número de alistamiento", 48, "Advertencia"
      TxtNumAlistamiento.Enabled = True
      TxtNumAlistamiento.SetFocus
      Exit Sub
   End If
   
   If CDbl(TxtNumAlistamiento.text) = 0 Then
      MsgBox "Falta ingresar el número de alistamiento", 48, "Advertencia"
      TxtNumAlistamiento.Enabled = True
      TxtNumAlistamiento.SetFocus
      Exit Sub
   End If
   
   
   If Trim(txtCedula.text) = "" Then
      txtCedula.SetFocus
      MsgBox "Falta ingresar el número de cédula", 48, "Advertencia"
      Exit Sub
   End If
   If Aviso_2_Conductores = False Then
      If Len(txtCedula.text) <= 4 Then
            MsgBox "Tamaño incorrecto del número de cédula", 48, "Advertencia"
            txtCedula.SetFocus
            Exit Sub
      End If
   End If
   If Trim(Txtconact.text) = "" Then
      Txtconact.SetFocus
      MsgBox "Falta ingresar el nombre de conductor", 48, "Advertencia"
      Exit Sub
   End If
   If Trim(txtapellido.text) = "" Then
      txtapellido.SetFocus
      MsgBox "Falta ingresar el apellido del conductor", 48, "Advertencia"
      Exit Sub
   End If
   
   If Aviso_2_Conductores = False Then
      If Trim(TxtCedula1.text) = "" Then
         TxtCedula1.Enabled = True
         MsgBox "Falta ingresar el número de cédula del segundo conductor", 48, "Advertencia"
         TxtCedula1.SetFocus
         Exit Sub
      End If
      
      If Len(TxtCedula1.text) <= 4 Then
         TxtCedula1.Enabled = True
         MsgBox "Tamaño incorrecto del número de cédula", 48, "Advertencia"
         TxtCedula1.SetFocus
         Exit Sub
      End If
      
      
      If Trim(Txtconact1.text) = "" Then
         Txtconact1.Enabled = True
         Txtconact1.SetFocus
         MsgBox "Falta ingresar el nombre del segundo conductor", 48, "Advertencia"
         Exit Sub
      End If
      If Trim(Txtapellido1.text) = "" Then
         Txtapellido1.Enabled = True
         Txtapellido1.SetFocus
         MsgBox "Falta ingresar el apellido del segundo conductor", 48, "Advertencia"
         Exit Sub
      End If
   End If
   If Mskfecha.ClipText = "" Then
      MsgBox "Falta ingresar la fecha de salida", 48, "Advertencia"
      Exit Sub
   End If
   If Mskhora.ClipText = "" Then
      MsgBox "Falta ingresar la hora de salida", 48, "Advertencia"
      Exit Sub
   End If
   If Mskvalor.ClipText = "" Then
      MsgBox "Falta ingresar el VALOR de la TASA DE USO ", 48, "Advertencia"
      Exit Sub
   End If
   If Txtoritra.text = "" Then
      MsgBox "Falta ingresar Origen / Transito", 48, "Advertencia"
      Exit Sub
   End If
   If Txtoritra.text <> "O" And Txtoritra.text <> "T" Then
      MsgBox "Debe ser <O> Origen / <T> Transito", 48, "Advertencia"
      Exit Sub
   End If
   
   If chkDatosPersonales.Value = 0 Then
      MsgBox "El conductor no ha autorizado el tratamiento de los datos personales", 48, "Advertencia"
      Exit Sub
   End If
      
   'Si el campo de pasajeros está habilitado, deberán digitar la cantidad de pasajeros
   If Txtpasajeros.Enabled = True Then
        If Txtpasajeros.text = "" Then
            MsgBox "Debe ingresar la cantidad de pasajeros", 48, "Advertencia"
            Exit Sub
        End If
   End If
      
   If UCase(G_Tipo_Terminal) = "CENTRAL" Or UCase(G_Tipo_Terminal) = "NORTE" Then
      If Txtoritra.text = "O" Then
         If TxtPeaje.Enabled = True Then
            If TxtPeaje.text = "" Then
               MsgBox "Falta ingresar el código de peaje", 48, "Advertencia"
               TxtPeaje.SetFocus
               Exit Sub
            End If
         Else
            TxtPeaje.text = "0"
         End If
      Else
         TxtPeaje.text = "0"
      End If
   Else
      TxtPeaje.text = "0"
   End If
   If Mskvalor.text = "0" Then
      Txtoritra.SetFocus
      Exit Sub
   End If
   
   If OptTipo(0).Value = True Then
      If L_HabilitaDevueltas = "S" And Mskdinero.ClipText = "" Or Mskdinero.ClipText = "0" Then
         MsgBox "Falta ingresar el dinero recibido", 48, "Advertencia"
         Mskdinero.SetFocus
         Exit Sub
      End If
   End If
   
   LHoraAsignada = Format(Mskhora.text, "hh:mm")
'   If G_Valida_Horario = "S" Then
'      If L_Valida_Horario = "S" Then
'         If FnValidaHorario = False Then
'            MsgBox "La empresa no tiene autorizado para despachar " & Chr(13) _
'                  & "                    un vehículo en esta hora", 48, "Advertencia"
'            If Mskhora.Enabled = True Then
'               Mskhora.SetFocus
'            End If
'            Exit Sub
'         End If
'      End If
'   End If
   
   On Error GoTo err_Btngrabar
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      fechaVenta = Fecha
      If Horope = "N" Then
         hora = Mid(Ftb01!psval, 12, 5)
         sfecsalveh = Mid(Ftb01!psval, 1, 10)
         shorsalven = hora
         Mskhorsis.text = hora
         
         
         fqry01 = "SELECT psval " _
                  & "FROM geparsis " _
                 & "WHERE pscod = 'ADERELOJ' "
      
         Set Tbl1 = Dbs.Execute(fqry01)
         If Tbl1.EOF = True Then
            hora1 = Format(DateAdd("n", Tbl1(0), hora), "hh:mm:ss")
            horaprueba = Format(hora, "hh:mm:ss")
         End If
         Tbl1.Close
         Mskhora.Mask = ""

         Mskhora.text = CDate(hora1)
         Mskfecha.text = Fecha
      
      Else
         shorsalven = Format(Mskhora.text, "HH:mm") 'Mid(Ftb01!psval, 12, 5)
         sfecsalveh = Mid(Ftb01!psval, 1, 10)
         hora = Mskhora.text ' (ftb01!psval, 12, 5)
         Mskhora.text = hora
         horaprueba = Format(hora, "hh:mm:ss")
         hora1 = Mskhora.text
         If L_Bandera_Fecha = "N" Then
            Mskfecha.text = Format(Fecha, "MM/DD/YYYY")
         End If
      End If
   End If
   Ftb01.Close
   
   
   
   If L_PicoyPlaca = "S" Then
      If Fn_PicoyPlaca(Mskfecha.text, Mid(Txtnumpla.text, Len(Txtnumpla.text), 1)) = True Then
         MsgBox "Por Pico y Placa no se le puede vender tasa de uso", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
         
   
   If Trim(LCodigoEmpresaIntegra) = "" Then
      'L_Codigo_Empresa = Txtcodemp.Text
      LCodigoEmpresaIntegra = Fn_Equivalencia_Empresa(L_Codigo_Empresa)
      If Trim(LCodigoEmpresaIntegra) = "" Then
         MsgBox "El código de la empresa no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtnumpla.Enabled = True
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   If Trim(LCodigoRutaIntegra) = "" Then
      LCodigoRutaIntegra = Fn_Equivalencia_Ruta(Txtcodrut.text)
      If Trim(LCodigoRutaIntegra) = "" Then
         Txtcodrut.SetFocus
         MsgBox "El código de la ruta no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtcodrut.SelStart = 0
         Txtcodrut.SelLength = Len(Txtcodrut.text)
         Exit Sub
      End If
   End If
   If Trim(LTipoServicioIntegra) = "" Then
      LTipoServicioIntegra = Traer_Descripcion(Txttipveh.text, "RELACION TIPO VEHICULO-TIPO SERVICIO")
      If Trim(LClaseIntegra) = "" Then
         MsgBox "El tipo de servicio no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtnumpla.Enabled = True
         Txtnumpla.SetFocus
         Exit Sub
      End If
      If Trim(LTipoServicioIntegra) = "" Then
         MsgBox "El tipo de servicio no tiene una relación con la clase", 48, "Advertencia"
         Txtnumpla.Enabled = True
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   
   If L_Numero_Terminal_Norte <> 0 And Trim(LCodigoRutaNorteIntegra) = "" Then
      LCodigoRutaNorteIntegra = Fn_Equivalencia_Ruta(L_Ruta_Transito_Norte)
      If Trim(LCodigoRutaNorteIntegra) = "" Then
         MsgBox "El código de la ruta transito no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   estado = "Activo"
   Modo = "D"
   PENDIENTE = "PENDIENTE"
   numpas = 0
   no = "N"
   
   'WECC MAY/05/2016
   '******************************************************************************
   sSaldoCupo = 0
   If OptTipo(1).Value = True Then
      
      If L_Tipo_Cupo = "E" Then
      
      'Sebastián Rondón - Mayo 11 de 2023
      'Si la empresa o ruta cuenta con convenio
      'Se le debe ser cargada a la empresa contratante
'      If Trim(txtCodEmp.text) <> Trim(L_Codigo_Empresa) Then
'        codempresa = L_Codigo_Empresa
'      Else
'        codempresa = txtCodEmp.text
'      End If

      'Sebastián Rondón - Noviembre 07 de 2023
      'Por petición de la termina
      'Si la empresa o ruta cuenta con convenio
      'Se le debe ser cargada a la empresa que compra la tasa de uso
      codempresa = Txtcodemp.text
      
      QRY = "SELECT * " _
            & "FROM cocupos " _
           & "WHERE cpempresa = '" & codempresa & "'"
      End If
      
      If L_Tipo_Cupo = "P" Then
         QRY = "SELECT * " _
            & "FROM cocupos INNER JOIN cocuposdet ON cpnumcup = cdnumcupo " _
           & "WHERE cdplaca = '" & Txtnumpla.text & "' AND cpestado = 'A' AND cdestado = 'A' "
           
            'cpcedula = '" & L_Cedula_Cupo & "' AND
      End If
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         If Tbl1!cpestado = "A" Then
            If Tbl1!cpsaldo > CDbl(Mskcambio.text) Then
               Lblcupo.Caption = "CUPO"
               If Tbl1!cpsaldo <= Tbl1!cpminimo Then
                  MsgBox "" & Chr(13) & "EL SALDO ESTA MENOR QUE EL VALOR MININO DEL CUPO " & Chr(13) & "", 48, "Advertencia"
               End If
               L_Numero_Cupo = Tbl1!cpnumcup
               sSaldoCupo = Tbl1!cpsaldo
               Modo = "P"
               sCupos = "P"
               OptTipo(1).Value = True
               
               If Tbl1!cpterminal <> "G" Then
                  If Tbl1!cpterminal <> Mid(G_Tipo_Terminal, 1, 1) Then
                     Modo = "D"
                     L_Numero_Cupo = 0
                     L_Tipo_Cupo = ""
                     OptTipo(0).Value = True
                     OptTipo(1).Enabled = False
                     Lblcupo.Caption = "LINEA"
                  Else
                     If Mid(G_Tipo_Terminal, 1, 1) = "N" And influencia = "No" Then
                        L_Numero_Cupo = 0
                        L_Tipo_Cupo = ""
                        OptTipo(0).Value = True
                        OptTipo(1).Enabled = False
                        Lblcupo.Caption = "LINEA"
                     Else
                        If Mid(G_Tipo_Terminal, 1, 1) = "N" And Txtoritra.text = "T" Then
                           L_Numero_Cupo = 0
                           L_Tipo_Cupo = ""
                           OptTipo(0).Value = True
                           OptTipo(1).Enabled = False
                           Lblcupo.Caption = "LINEA"
                        End If
                     End If
                  End If
               End If
               
            Else
               '********************************************************************************
               '********************************************************************************
               'Tbl1.Close
               'MsgBox "No tiene saldo para comprar la tasa de uso", 48, "Advertencia"
               'Txtnumpla.SetFocus
               'Exit Sub
               If Lblcupo.Caption = "CUPO" Then
                  G_Saldo_Cupo = Tbl1!cpsaldo
                  On Error Resume Next
                  AviCupos.Show 1
                  On Error GoTo err_Btngrabar
                  OptTipo(0).Value = True
                  OptTipo(1).Enabled = False
                  Lblcupo.Caption = "LINEA"
                  'MsgBox "El cupo no tiene saldo para comprar la tasa de uso," & Chr(13) & _
                         "se le vendera normalmente, el saldo es: " & Format(Tbl1!cpsaldo, "$#,###0"), 48, "Advertencia"
                  L_Numero_Cupo = 0
                  L_Tipo_Cupo = ""
               End If
               '********************************************************************************
               '********************************************************************************
            End If
         Else
            '********************************************************************************
            '********************************************************************************
            'Tbl1.Close
            'MsgBox "El cupo esta inactivo", 48, "Advertencia"
            'Txtnumpla.SetFocus
            'Exit Sub
            If Lblcupo.Caption <> "CUPO" Then
               OptTipo(0).Value = True
               OptTipo(1).Enabled = False
               Lblcupo.Caption = "LINEA"
               MsgBox "El cupo esta inactivo" & Chr(13) & "                  se le vendera normalmente", 48, "Advertencia"
               L_Numero_Cupo = 0
            End If
            '********************************************************************************
            '********************************************************************************
         End If
      Else
         If Lblcupo.Caption <> "CUPO" Then
            Lblcupo.Caption = "LINEA"
            OptTipo(0).Value = True
            OptTipo(1).Enabled = False
            L_Numero_Cupo = 0
         End If
      End If
      Tbl1.Close
   End If
   
   
   
   'hora = DateAdd("n", -10, Mskhora.Text)
   
   'horaprueba = Format(DateAdd("n", -10, hora), "hh:mm:ss")
   
   Fecha = Mid(Mskhora.text, 7, 4)
   If Fecha < "2000" Then
      'Mskhora.Mask = ""
      'Mskhora.Text = Mskfecha.Text & " " & hora
      Mskhora.Mask = ""
      Mskhora.text = Mskfecha.text & " " & hora1
      hora1 = Mskfecha.text & " " & CDate(hora1)
   End If
   horven = Format(Mskhorsis.text, "hh:mm")
   
   'wecc 02/22/2002
   'se inserto el campo del valor de alcoholemia
   If chkcredito.Enabled = True Then
      If chkcredito.Value = 0 Then
         credito = "N"
         crecon = "Contado"
      Else
         credito = "S"
         crecon = "Crédito"
      End If
   Else
      credito = "N"
      crecon = "Contado"
   End If
   
   'Sebastián Rondón - Enero 27 de 2025
   'Si el check de crédito está activo, valido por última vez si la empresa maneja este método
   If chkcredito.Enabled = True Then
        If chkcredito.Value = 1 Then
            QRY = "SELECT hecredit FROM ophojemp WHERE hecodigo = '" & Txtcodemp.text & "' "
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                If Tbl1!hecredit = "N" Then
                    credito = "N"
                    crecon = "Contado"
                End If
            End If
        End If
   End If
   
   '*******************************************************************
   'Variable apra almacenar la fecha del dia a cual corresponde la TU
   'WECC 06/19/2012
   '*******************************************************************
   sFechaDia = Format(Mskfecha.text, "MM/dd/yyyy")
   sHora0 = "21:00"
   sHora7 = "23:59"
   If Val(G_Turno) = 1 Then
      If CDate(Mskhora.text) >= CDate(sHora0) And CDate(Mskhora.text) <= CDate(sHora7) Then
         sFechaDia = Format(DateAdd("d", 1, Mskfecha.text), "MM/dd/yyyy")
      End If
   End If
   '*******************************************************************
   '*******************************************************************
   If Trim(TxtNumAlistamiento.text) = "" Then
      sNumAlistamiento = "0"
   Else
      sNumAlistamiento = TxtNumAlistamiento.text
   End If
   
   'WECC 12/02/2016 CONVENIOS PARA INTEGRA Y LA SUPERTRANSPORTES
   sTipoConvenio = "1"
   If L_Tipo_Convenio = "" Then
      sTipoConvenio = "1"
   ElseIf L_Tipo_Convenio = "T" Then
      sTipoConvenio = "3"
   ElseIf L_Tipo_Convenio = "E" Then
      sTipoConvenio = "2"
   End If
   
   Dbs.BeginTrans
   On Error GoTo transac
      
    horven = Format(horven, "HH:mm:ss")
      
   'Camilo Campos 06/19/2025 con el fin de que se paguen las multas dentro de la misma tasa de uso sin necesidad de ir al otro formulario
   'Hacemos el guardado de las seleccionadas, estas tasas seleccionadas estan en un arreglo previo (arreglo que se llena al salir del aviso)
   'Se hace la misma sentencia que tienen en el formulario de pago de multas común, con la excepción que ligamos de una vez la tasa actual a dicha multa pagada
   If TieneMultasSeleccionadas() Then 'metodo que devuelve true en caso de tener multas seleccionadas
   
       For i = LBound(multasGuardadas) To UBound(multasGuardadas)
    
    ' Obtener el nuevo número de recibo
            QRY = "SELECT MAX(abrecibo) FROM opabonos"
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                If IsNull(Tbl2(0)) Then
                    recpag = 1
                Else
                    recpag = Tbl2(0) + 1
                End If
            End If
        
            ' Actualizar la multa correspondiente
            QRY = "UPDATE opmultas " & _
                  "SET muestado = 'Pagada'," & _
                  "    murecpag = " & recpag & "," & _
                  "    muturno = " & G_Turno & "," & _
                  "    muoperar = '" & Login & "'," & _
                  "    muhora = TO_DATE('" & horven & "','hh24:mi:ss')," & _
                  "    mucerrado = 'N'," & _
                  "    muciedia = 'N'," & _
                  "    mufecsis = TO_DATE('" & fechaVenta & "','MM/DD/YYYY'), " & _
                  "    munetpag = " & multasGuardadas(i).valor & "," & _
                  "    mutasapag = '" & Txtnumcon.text & "', " & _
                  "    munomcaj = '" & Pc & "' " & _
                  "WHERE munumero = " & multasGuardadas(i).id
        
            Dbs.Execute QRY
        
            ' Insertar el abono correspondiente
            QRY = "INSERT INTO opabonos " & _
                  "(abnumero, abrecibo, abfecha, abvalor, abfuncio, abcerrado, abciedia, abturno, abnomcaj) VALUES (" & _
                  multasGuardadas(i).id & "," & recpag & ", TO_DATE('" & fechaVenta & "','MM/DD/YYYY'), " & _
                  multasGuardadas(i).valor & ", '" & Login & "', 'N', 'N', " & G_Turno & ", '" & Pc & "')"
        
            Dbs.Execute QRY
        
        'La terminal debe generar si o si un comprobante de pago SEPARADO (una por una y no todas en una sola)
             If Not imprimirMulta(multasGuardadas(i).fechaMulta, _
                         multasGuardadas(i).codinf, _
                         multasGuardadas(i).id, _
                         multasGuardadas(i).valor, _
                         fechaVenta, _
                         horven) Then
                GoTo transac 'significa que devolvió false, por ello vamos a este error que posee el "rollback"
            End If
            DoEvents
        Next i
        
        Tbl2.Close
    End If
    
    
    'Camilo Campos 06/20/2025 se hace el metodo para dejar pagos los parqueaderos seleccionados y ligarlos a la tasa de uso actual
    If TieneParqueaderosSeleccionados() Then 'metodo para que devuelva true en caso de tener parqueaderos seleccionados
        
        For i = LBound(parqueGuardado) To UBound(parqueGuardado)
        
        QRY = "UPDATE coparqueop " _
             & "SET pafecpag = to_date('" & fechaVenta & "','MM/dd/yyyy') , " _
                & " pacaja = '" & Pc & "', " _
                & " paturno = " & G_Turno & ", " _
                & " patasapag = '" & Txtnumcon.text & "', " _
                & " pacerrado= 'N', " _
                & " paestado = 'C' " _
           & "WHERE panumero = " & parqueGuardado(i).id & ""
      
         Dbs.Execute QRY
    
    'Procedemos a replicar el código que contabilizaba el tiempo de permanencia
    'con unas ligeras mejoras en terminos de legibilidad y eficienia
    
        Dim horaIngreso As Date
        Dim horaSalida As Date
        Dim minutosTotales As Long
        Dim minutosCobro As Long
        Dim horas As Long
        Dim minutos As Long
        
        horaIngreso = parqueGuardado(i).fecing
        horaSalida = parqueGuardado(i).fecsal
        
        minutosTotales = DateDiff("n", horaIngreso, horaSalida)
        sMinutosIniciarCobro = G_Minutos_Minimos_Parqueadero
        sMinutosMaximos = 60
        
        If minutosTotales > sMinutosIniciarCobro Then
            minutosCobro = minutosTotales - sMinutosIniciarCobro
            horas = minutosCobro \ sMinutosMaximos
            minutos = minutosCobro Mod sMinutosMaximos
            
            If minutos >= sMinutosIniciarCobro Then
                horas = horas + 1
                minutos = 0
            End If
            
            sTiempoPermanencia = horas & " Horas " & minutos & " minutos"
        Else
            sTiempoPermanencia = "0 Horas 0 minutos"
        End If
         
         
    'Guardamos las variables necesarias para la impresión del comprobante de la permanencia
    
         sNombreTerminal = G_NomEmpresa
         sNit = "NIT.860.052.155-6"
         sFecha = Format(fechaVenta, "MMM/dd/yyyy")
         sHora = Format(Mskhora.text, "hh:mm")
         sNumeroInterno = TxtNumInt.text
         sPlaca = Txtnumpla.text
         sEmpresa = Lblnomemp.Caption
         sFechaIngreso = parqueGuardado(i).fecing
         sFechaSalida = parqueGuardado(i).fecsal
         sValorPagar = parqueGuardado(i).valor
         sDescuento = "0"
         sValor = CDbl(sValorPagar) + CDbl(sDescuento)
         sCajero = Login
         sCaja = Pc
         sturno = G_Turno
         sNumero = parqueGuardado(i).id
         sNumeroBarras = "*" & parqueGuardado(i).id & "*"
         If parqueGuardado(i).tipo = "P" Then
            snombrefechaingreso = "FECHA INGRESO"
            sTitulo = "PAGO POR TIEMPO DE PERMANENCIA"
            sTiempoVigencia = "TIEMPO PERMANENCIA:"
         ElseIf parqueGuardado(i).tipo = "V" Then
            snombrefechaingreso = "FECHA LIMITE"
            sTitulo = "PAGO POR VIGENCIA"
            sTiempoVigencia = "VIGENCIA:"
         End If
         
         'este metodo si nos lo copiamos tal cual como esta en la pantalla de pagos de permanencias
         'Se hace como funcion para controlar
         If Not ReImprimirPago(sTiempoPermanencia, sNombreTerminal, sNit, sFecha, _
                       sHora, sNumeroInterno, sPlaca, sEmpresa, _
                       sFechaIngreso, sFechaSalida, sValorPagar, sDescuento, _
                       sValor, sCajero, sCaja, sturno, _
                       sNumero, sNumeroBarras, snombrefechaingreso, sTitulo, _
                       sTiempoVigencia) Then
            GoTo transac 'Significa que devolvió falso y para ello vamos al error con el "rollback"
        End If
       
        DoEvents
       Next i
    End If
    
   'base de datos del servidor
   'On Error GoTo err_Btngrabar_Dbx
   Guardo_Servidor = "N"
   sNomConductor = Txtconact.text & " " & txtapellido.text
   sNomConductor1 = Txtconact1.text & " " & Txtapellido1.text
   sHora = Format(Mskhora.text, "MM/dd/yyyy HH:mm:ss")
  
   QRY = "INSERT INTO coconduc " _
             & "(conumero, cocodrut, cocodemp, cofecsal, " _
             & " cohorven, cohorsal, coplaca, conumint, " _
             & " coclase, coconduc, cooritra, covalor, " _
             & " cofuncio, coturno, coestado, comodo, " _
             & " coorigen, cocerrado, cociedia, covalalc, " _
             & " coparque, conomcaj, copastra, cocredit, codirecc, " _
             & "coterminal, cooperador"
             
   If Trim(Txtcodemp.text) <> Trim(L_Codigo_Empresa) Then
      QRY = QRY & ", cocodempcon"
   End If
   
   If Txtpasajeros.Enabled = True Then
        If Txtpasajeros.text <> "" Then
            QRY = QRY & ", copasaje"
        End If
   End If
             
   QRY = QRY & ", cohorasig, cocedula, cocedula1, coconduc1, conumalistamiento, cotipoconvenio, conumcupo, copeaje) VALUES (" _
             & Txtnumcon.text & ",'" & Txtcodrut.text & "','" & L_Codigo_Empresa & "',to_date('" & fechaVenta & "','MM/DD/YYYY'),to_date('" _
             & horven & "','hh24:mi:ss'),to_date('" & sHora & "','MM/DD/YYYY hh24:mi:ss'),'" & Txtnumpla.text & "','" & TxtNumInt.text & "','" _
             & Txttipveh.text & "','" & sNomConductor & "','" & Txtoritra.text & "','" & Mskvalor.text & "','" _
             & Login & "','" & G_Turno & "','" & estado & "','" & Modo & "','" _
             & Txtorigen.text & "','" & no & "','" & no & "','" & Mskrecibo.text & "','" _
             & MskParqueadero.text & "','" & txtcaja.text & "','" & PasajerosOrigen & "','" & credito & "','" & direruta & "','" _
             & L_Terminal & "','" & L_Operador & "' "
             
   L_EmpresaContrada = ""
   If Trim(Txtcodemp.text) <> Trim(L_Codigo_Empresa) Then
      QRY = QRY & ", '" & Txtcodemp.text & "'"
      L_EmpresaContrada = Txtcodemp.text
   End If
   
   If Txtpasajeros.Enabled = True Then
        If Txtpasajeros.text <> "" Then
            QRY = QRY & ", " & Txtpasajeros.text & ""
        End If
   End If
   
   'QRY = QRY & ",TO_DATE('" & LHoraAsignada & "','hh24:mi'), '" & txtcedula.Text & "', '" & Txtcedula1.Text & "','" & sNomConductor1 & "', '" & sNumAlistamiento & "', '" & sTipoConvenio & "'," & L_Numero_Cupo & ",'" & TxtPeaje.Text & "', 'PEN', '" & MskValTag.Text & "')"
   QRY = QRY & ",TO_DATE('" & LHoraAsignada & "','hh24:mi'), '" & txtCedula.text & "', '" & TxtCedula1.text & "','" & sNomConductor1 & "', '" & sNumAlistamiento & "', '" & sTipoConvenio & "'," & L_Numero_Cupo & ",'" & TxtPeaje.text & "')"
   
   Dbs.Execute QRY
   
   '*************************************************************************************
   'Sebastián Rondón Noviembre 08 - 2022
   'Se inserta en la tabla que luego enviará la información al Web Service de Protech
   QRYP = "INSERT INTO coconenvp (conumero, coplaca, cofecsal, cohorsal, cocaja, coterminal, coestado)" _
            & "VALUES (" & Txtnumcon.text & ", '" & Txtnumpla.text & "', to_date('" & fechaVenta & "', 'MM/DD/YYYY'), to_date('" & horven & "', 'hh24:mi:ss'), '" _
            & txtcaja.text & "', '" & L_Terminal & "', 'PEN')"
            
    Dbs.Execute QRYP
   
   '**************************************************************************************
   
   sconductor = Trim(Mid(Txtconact.text, 1, 24)) & " " & Trim(Mid(txtapellido.text, 1, 24))
   sCedula = Trim(txtCedula.text)
   
   sConductor1 = Trim(Mid(Txtconact1.text, 1, 24)) & " " & Trim(Mid(Txtapellido1.text, 1, 24))
   sConductor1 = Trim(sConductor1)
   sCedula1 = Trim(TxtCedula1.text)
   
   sModo = "L"
   sValorTasa = Mskvalor.text
   sValorAlcoholimetria = Mskrecibo.text
   sValorSello = 0
   
   
   sOperador = "N"
   sValorPublicidad = 0
   sValorPuertas = 0
   sValorOtros = 0
   sNumeroPlanilla = 0
   
   sCodRutaTerminal = Trim(Txtcodrut.text)
   sTipServicioTerminal = Trim(Txttipveh.text)
   L_Bandera_Escoger = "SI"
   If LNumeroPuestosIntegra = "" Then
      LNumeroPuestosIntegra = 0
   End If
   
   LQryTransmitir = "INSERT INTO coconduc " _
             & "(conumero, coterminal, cocodrut, cocodemp, " _
             & "cofecsal, cohorven, cohorsal, cohorasig, " _
             & "coplaca, coclase, coconduc, cooritra, " _
             & "comodo, covalor, cofuncio, coturno, " _
             & "coestado, conumint, covalalc, covalsel, " _
             & "cooperador, covalpub, covalpue, covalotr, " _
             & "conumpue, cocedula, cotipser, codirecc, " _
             & "conumpla, conomcaj, cocodrutter, cotipserter, " _
             & "cofecdia, coclaeva, covaltra, cocedula1, " _
             & "coconduc1, conumalistamiento, conumconvenio, coempcontratada, cotipoconvenio, coclaseconvenio) VALUES (" _
             & Txtnumcon.text & ",'" & G_Codigo_Terminal & "','" & LCodigoRutaIntegra & "','" & LCodigoEmpresaIntegra & "',TO_DATE('" _
             & fechaVenta & "','MM/dd/yyyy'), TO_DATE('" & horven & "','hh24:mi:ss'),TO_DATE('" & sHora & "','MM/dd/yyyy hh24:mi:ss'),TO_DATE('" & LHoraAsignada & "','hh24:mi'), '" _
             & Txtnumpla.text & "','" & LClaseIntegra & "','" & sconductor & "','" & Txtoritra.text & "','" _
             & sModo & "'," & sValorTasa & ",'" & Login & "'," & G_Turno & ",'" _
             & estado & "'," & TxtNumInt.text & "," & sValorAlcoholimetria & "," & sValorSello & ",'" _
             & sOperador & "'," & sValorPublicidad & "," & sValorPuertas & "," & sValorOtros & "," _
             & LNumeroPuestosIntegra & ", '" & sCedula & "', '" & LTipoServicioIntegra & "','" & LDirreccionRuta & "'," _
             & sNumeroPlanilla & ", '" & Pc & "','" & sCodRutaTerminal & "','" & sTipServicioTerminal & "',TO_DATE('" _
             & sFechaDia & "','MM/DD/YYYY'),'" & L_Clase & "'," & L_ValorTransitoSur & ", '" & sCedula1 & "','" _
             & sConductor1 & "', '" & sNumAlistamiento & "', " _
             & sNumConvenio & ", '" & L_EmpresaContrada & "', '" & sTipoConvenio & "','" & LTipoConvenio & "')"
               
Sigue_Proceso_Grabacion:
   'SE ALMACENA EL REGISTRO EN UNA TABLA PARA LUEGO TRANSMITIR
   lQry = Replace(LQryTransmitir, "'", "<")
   QRY = "INSERT INTO optransmitir(trsql) VALUES('" & lQry & "')"
   Dbs.Execute QRY
   
   If Trim(TxtPeaje.text) <> "" And Trim(TxtPeaje.text) <> "0" Then
      QRY = "UPDATE opruta SET ruultpeaje = '" & TxtPeaje.text & "' " _
            & "WHERE RUCODIGO = '" & Txtcodrut.text & "' "
            
      Dbs.Execute QRY
   End If
   
   'Sebastian Rondon - 28 de Marzo de 2023
   'Se guarda el registro de que el conductor acepto el tratamiento de datos personales
   If chkDatosPersonales.Value = 1 Then
      QRY = "INSERT INTO cocodatper (conumero, codatper) VALUES (" & Txtnumcon.text & ", 'S')"
      Dbs.Execute QRY
   End If
   
   'WECC FEB/22/2021 CAU 3802
   '********************************************************
   'If L_TipoRuta = "S" Then
      'qry = "DELETE FROM copasajcovid WHERE pcplaca = '" & Trim(Txtnumpla.Text) & "'"
      
      'Dbs.Execute qry
      
      'qry = "INSERT INTO copasajcovid (pcplaca, pcfecha, pcpasajori, pcpasajtra, pcterminal) " _
               '& "VALUES('" _
               '& Trim(Txtnumpla.Text) & "', TO_DATE('" & sHora & "','MM/DD/YYYY hh24:mi:ss'),0,0,'" & G_Codigo_Terminal & "')"
      
      'Dbs.Execute qry
   'End If
   
   '********************************************************
   'WECC FEB 02/2017 CAU 14098
   '********************************************************
   If L_Numero_Terminal_Norte <> 0 Then
      QRY = "UPDATE cotransitonorte " _
             & "SET TNNUMEROTASA = " & Txtnumcon.text & " , " _
                & " TNCAJATASA = '" & txtcaja.text & "', " _
                & " TNFECHATASA = to_date('" & fechaVenta & "','MM/DD/YYYY'), " _
                & " TNHORATASA = to_date('" & horven & " ','hh24:mi:ss'), " _
                & " TNTURNOTASA = " & G_Turno & ", " _
                & " TNCERRADOTASA = 'N', " _
                & " TNFUNCIONARIOTASA = '" & Login & "', " _
                & " TNESTADO = 'C', " _
                & " TNMODO = '" & Modo & "' " _
         & "WHERE tnnumero = " & L_Numero_Terminal_Norte & " " _
          & " AND tnplaca = '" & Txtnumpla.text & "'"
          
      Dbs.Execute QRY
      
   End If
   

   '********************************************************
    L_BanderaVentaTasa = "EFECTIVO"
               
   If Modo = "P" Then
      L_BanderaVentaTasa = "CUPO"
      
'      If Trim(Txtcodemp.text) <> Trim(L_Codigo_Empresa) Then
'        codempresa = L_Codigo_Empresa
'      Else
'        codempresa = Txtcodemp.text
'      End If

      'Sebastián Rondón - Enero 15 de 2024
      'Por petición de la termina
      'Si la empresa o ruta cuenta con convenio
      'Se le debe ser cargada a la empresa que compra la tasa de uso
      codempresa = Txtcodemp.text
      
      QRY = "INSERT INTO opmovimientos (MOFECHA, MOHORA, MOTIPMOV, MOVALOR, MOSALDO, MOTIPO, MOTIPCUPO, MOTASAUSO, MOCAJA, MOCODEMP " _
               
      LTotalCupo = CDbl(Mskrecibo.text) + CDbl(Mskvalor.text)
      If L_Tipo_Cupo = "E" Then
         QRY = QRY & " "
      ElseIf L_Tipo_Cupo = "P" Then
         QRY = QRY & ", MOCEDULA, MOPLACA "
      End If
      
      QRY = QRY & ") VALUES(TO_DATE('" & fechaVenta & "','MM/DD/YYYY'),TO_DATE('" & horven & "','hh24:mi:ss'),'CR'," _
                        & LTotalCupo & "," & sSaldoCupo & ",'VTTU','" & L_Tipo_Cupo & "', " & Txtnumcon.text & ",'" & txtcaja.text & "','" & codempresa & "' "
      
      If L_Tipo_Cupo = "E" Then
         QRY = QRY & ""
      Else
         QRY = QRY & ",'" & txtCedula.text & "','" & Txtnumpla.text & "'"
      End If
      
      QRY = QRY & ")"
      
      
      'L_Codigo_Empresa
      Dbs.Execute QRY
      
      QRY = "UPDATE cocupos " _
             & "SET cpsaldo = cpsaldo - " & LTotalCupo & " WHERE cpnumcup = " & L_Numero_Cupo & "" _
           
'      If L_Tipo_Cupo = "E" Then
'         Qry = Qry & "WHERE cpempresa = '" & Txtcodemp.Text & "' "
'      End If
'
'      If L_Tipo_Cupo = "C" Then
'         Qry = Qry & "WHERE cpcedula = '" & L_Cedula_Cupo & "' " _
'                     & "AND cpplaca = '" & Txtnumpla.Text & "' "
'      End If
      QRY = QRY & " AND cpestado = 'A' "
      
      Dbs.Execute QRY
   End If
   
   'Sebastián Rondón - Mayo 03 de 2024
   'Se verifica si la tasa es Crédito, para reflejarlo en la impresión de la tasa de uso
   If chkcredito.Enabled = True Then
      If chkcredito.Value = 0 Then
         L_BanderaVentaTasa = L_BanderaVentaTasa
      Else
         L_BanderaVentaTasa = "CREDITO"
      End If
   End If
   
   If L_Numero_Despacho <> 0 Then
      QRY = "UPDATE codespachos " _
             & "SET deestado = 'T', " _
                & " denumtasa = " & Txtnumcon.text & ", " _
                & " denomcaj = '" & txtcaja.text & "' " _
           & "WHERE denumero = " & L_Numero_Despacho & " "
      
      Dbs.Execute QRY
   End If
   '********************************************************
   
   If L_Transmitir_Conductor1 = "S" And datosCedula1(3) <> "" Then
       edad = 0
       estatura = 0
       licencia = 0
       peso = 0

       LQryTransmitir = " INSERT INTO coconductor " & _
                 "(cocedula, coapellidos, conombres, " & _
                 "cofecing, cofecnac, cocodemp, coobserv, " & _
                 "coedad,  coestatura, colicencia, copeso, " & _
                 "coprinom, cosegnom, copriape, cosegape, cosexo, corh) " & _
                 "VALUES ('" & txtCedula.text & "' ,'" & txtapellido.text & "','" & Txtconact.text & _
                 "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & datosCedula1(9) & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','NINGUNA'," & _
                 edad & ", '" & estatura & "','" & licencia & "','" & peso & "'," & _
                 "'" & datosCedula1(6) & "','" & datosCedula1(7) & "','" & datosCedula1(4) & "','" & datosCedula1(5) & "','" & datosCedula1(8) & "','" & datosCedula1(11) & "')"

       'SE ALMACENA EL REGISTRO EN UNA TABLA PARA LUEGO TRANSMITIR
       lQry = Replace(LQryTransmitir, "'", "<")
       QRY = "INSERT INTO optransmitir(trsql) VALUES('" & lQry & "')"
       Dbs.Execute QRY
   End If
   If L_Transmitir_Conductor2 = "S" And datosCedula2(3) <> "" Then
       edad = 0
       estatura = 0
       licencia = 0
       peso = 0
       
       LQryTransmitir = " INSERT INTO coconductor " & _
                     "(cocedula, coapellidos, conombres, " & _
                     "cofecing, cofecnac, cocodemp, coobserv, " & _
                     "coedad,  coestatura, colicencia, copeso, " & _
                     "coprinom, cosegnom, copriape, cosegape, cosexo, corh) " & _
                     "VALUES ('" & TxtCedula1.text & "' ,'" & Txtapellido1.text & "','" & Txtconact1.text & _
                     "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & datosCedula2(9) & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','NINGUNA'," & _
                     edad & ", '" & estatura & "','" & licencia & "','" & peso & "'," & _
                     "'" & datosCedula2(6) & "','" & datosCedula2(7) & "','" & datosCedula2(4) & "','" & datosCedula2(5) & "','" & datosCedula2(8) & "','" & datosCedula2(11) & "')"

       'SE ALMACENA EL REGISTRO EN UNA TABLA PARA LUEGO TRANSMITIR
       lQry = Replace(LQryTransmitir, "'", "<")
       QRY = "INSERT INTO optransmitir(trsql) VALUES('" & lQry & "')"
       Dbs.Execute QRY
   End If
   If ExisteCedula = 1 Then
      edad = 0
      estatura = 0
      peso = 0
      observacion = "NINGUNA"
      If datosCedula1(3) = "" Then
         QRY = " INSERT INTO coconductor " _
                    & "(cocedula, coapellidos, conombres, " _
                    & "cofecing, cofecnac, cocodemp, coobserv, " _
                    & "coedad,  coestatura, colicencia, copeso) " _
                    & "VALUES ('" _
                    & txtCedula.text & "' ,'" & txtapellido.text & "','" & Txtconact.text & "',TO_DATE('" _
                    & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','" & observacion & "'," _
                    & edad & ", '" & estatura & "','" & txtCedula.text & "','" & peso & "')"
      Else
          QRY = " INSERT INTO coconductor " & _
                "(cocedula, coapellidos, conombres, " & _
                "cofecing, cofecnac, cocodemp, coobserv, " & _
                "coedad,  coestatura, colicencia, copeso, " & _
                "coprinom, cosegnom, copriape, cosegape, cosexo, corh) " & _
                "VALUES ('" & txtCedula.text & "' ,'" & txtapellido.text & "','" & Txtconact.text & _
                "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & datosCedula1(9) & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','NINGUNA'," & _
                edad & ", '" & estatura & "','" & licencia & "','" & peso & "'," & _
                "'" & datosCedula1(6) & "','" & datosCedula1(7) & "','" & datosCedula1(4) & "','" & datosCedula1(5) & "','" & datosCedula1(8) & "','" & datosCedula1(11) & "')"
      End If
      Dbs.Execute QRY
   Else
      If datosCedula1(3) = "" Then
          QRY = "UPDATE coconductor " _
             & " SET conombres = '" & Txtconact.text & "', " _
               & "   coapellidos = '" & txtapellido.text & "', coestviaje ='S' " _
           & " WHERE cocedula = '" & txtCedula.text & "' "
      Else
          QRY = "UPDATE coconductor SET " & _
                "coapellidos = '" & txtapellido.text & "', " & _
                "conombres = '" & Txtconact.text & "', " & _
                "coprinom = '" & datosCedula1(6) & "', " & _
                "cosegnom = '" & datosCedula1(7) & "', " & _
                "copriape = '" & datosCedula1(4) & "', " & _
                "cosegape = '" & datosCedula1(5) & "', " & _
                "cosexo = '" & datosCedula1(8) & "', " & _
                "corh = '" & datosCedula1(11) & "', " & _
                "cofecnac = TO_DATE('" & datosCedula1(9) & "','MM/DD/YYYY'), coestviaje ='S' " & _
                "WHERE cocedula = '" & txtCedula.text & "' "
      End If
      Dbs.Execute QRY
                       
                       
   End If
   '************************************************************************************************
   'segundo conductor wecc 05/02/2013
   If ExisteCedula1 = 1 Then
      edad = 0
      estatura = 0
      peso = 0
      observacion = "NINGUNA"
      If Txtconact1.text <> "" Then
         If datosCedula2(3) = "" Then
            QRY = " INSERT INTO coconductor " _
              & "(cocedula, coapellidos, conombres, " _
              & "cofecing, cofecnac, cocodemp, coobserv, " _
              & "coedad,  coestatura, colicencia, copeso) " _
              & "VALUES ('" _
              & TxtCedula1.text & "' ,'" & Txtapellido1.text & "','" & Txtconact1.text & "',TO_DATE('" _
              & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','" & observacion & "'," _
              & edad & ", '" & estatura & "','" & TxtCedula1.text & "','" & peso & "')"

         Else
            QRY = " INSERT INTO coconductor " & _
                    "(cocedula, coapellidos, conombres, " & _
                    "cofecing, cofecnac, cocodemp, coobserv, " & _
                    "coedad,  coestatura, colicencia, copeso, " & _
                    "coprinom, cosegnom, copriape, cosegape, cosexo, corh) " & _
                    "VALUES ('" & TxtCedula1.text & "' ,'" & Txtapellido1.text & "','" & Txtconact1.text & _
                    "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & datosCedula2(9) & "','MM/DD/YYYY'),'" & Txtcodemp.text & "','NINGUNA'," & _
                    edad & ", '" & estatura & "','" & licencia & "','" & peso & "'," & _
                    "'" & datosCedula2(6) & "','" & datosCedula2(7) & "','" & datosCedula2(4) & "','" & datosCedula2(5) & "','" & datosCedula2(8) & "','" & datosCedula2(11) & "')"

         End If

         Dbs.Execute QRY
      End If
      
   Else
      If Trim(Txtconact1.text) <> "" Then
         If datosCedula2(3) = "" Then
            QRY = "UPDATE coconductor " _
                  & " SET conombres = '" & Txtconact1.text & "', " _
                      & " coapellidos = '" & Txtapellido1.text & "', coestviaje ='S' " _
                 & "WHERE cocedula = '" & TxtCedula1.text & "' "
         Else
            QRY = "UPDATE coconductor  " & _
                     "SET coapellidos = '" & Txtapellido1.text & "', " & _
                        " conombres = '" & Txtconact1.text & "', " & _
                        " coprinom = '" & datosCedula2(6) & "', " & _
                        " cosegnom = '" & datosCedula2(7) & "', " & _
                        " copriape = '" & datosCedula2(4) & "', " & _
                        " cosegape = '" & datosCedula2(5) & "', " & _
                        " cosexo = '" & datosCedula2(8) & "', " & _
                        " corh = '" & datosCedula2(11) & "', " & _
                        " cofecnac = TO_DATE('" & datosCedula2(9) & "','MM/DD/YYYY'), coestviaje ='S' " & _
                   "WHERE cocedula = '" & TxtCedula1.text & "' "
         End If

         Dbs.Execute QRY
      End If
                       
   End If
   
   If Aviso_2_Conductores = False Then
      sTipo = "CON"
      QRY = "INSERT INTO gelogconali(" _
                        & "catipo, cafecha, cafechahora, " _
                        & "cafuncionario, caplaca, cacodemp, carutadestino, carutaorigen, " _
                        & "catipser, cacedula, cacedula1, catasauso, canomcaja)" _
               & "VALUES('" _
                        & sTipo & "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'), TO_DATE('" & sHora & "','MM/DD/YYYY hh24:mi:ss'), '" _
                        & Login & "','" & Txtnumpla.text & "','" & L_Codigo_Empresa & "','" & Txtcodrut.text & "', '" & Txtorigen.text & "', '" _
                        & Txttipveh.text & "','" & sCedula & "','" & sCedula1 & "'," & Txtnumcon.text & ",'" & txtcaja.text & "')"
      
      Dbs.Execute QRY
   End If
   
   If Aviso_Alistamiento = False Then
      sTipo = "ALI"
      QRY = "INSERT INTO gelogconali(" _
                        & "catipo, cafecha, cafechahora, " _
                        & "cafuncionario, caplaca, cacodemp, carutadestino,  " _
                        & "catipser, cacedula, cacedula1, catasauso, canomcaja)" _
               & "VALUES('" _
                        & sTipo & "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'), TO_DATE('" & sHora & "','MM/DD/YYYY hh24:mi:ss'), '" _
                        & Login & "','" & Txtnumpla.text & "','" & L_Codigo_Empresa & "','" & Txtcodrut.text & "', '" _
                        & Txttipveh.text & "','" & sCedula & "','" & sCedula1 & "'," & Txtnumcon.text & ",'" & txtcaja.text & "')"
      
      Dbs.Execute QRY
   End If
   '***********************************************
   'wecc 08/26/2010    parqueadero
   'WECC SE PONE EN COMENTARIO PORQUE VAN A CANCELAR LAS PERMANENCIAS DESDE LA PANTALLA DE CAPPARQUE
'   If L_Consecutivo_Parqueadero <> 0 Then
'        If G_Tipo_Terminal = "CENTRAL" Then
'            qry = "UPDATE coparqueo " _
'                   & "SET pafecpag = to_date('" & Mskfecha.Text & "','MM/dd/yyyy') , " _
'
'            If G_Tipo_Parqueadero = "P" Then
'               'WECC 10/17/2014 CAU 8616 PARA Q NO ACTUALICE EL NUMERO DE TASA
'               qry = qry & " patasa = " & Txtnumcon.Text & ", "
'            End If
'
'            qry = qry & " pacaja = '" & txtcaja.Text & "', " _
'                      & " paturno = " & G_Turno & ", " _
'                      & " pacerrado= 'N', " _
'                      & " paestado = 'C' " _
'                 & "WHERE panumero = " & L_Consecutivo_Parqueadero & ""
'
'            Dbs.Execute qry
'        End If
'   End If
   scero = 0
   
   If G_Tipo_Terminal = "CENTRAL" Then
        QRY = "UPDATE coparqueop " _
               & "SET paaviso = paaviso + 1 " _
             & "WHERE paplaca = '" & Txtnumpla.text & "' AND pavalor <> " & scero & " " _
               & "AND paestado = 'A' "
              
        Dbs.Execute QRY
   End If
   
   
   
   '***********************************************
   
   'shorsalven = Format(shora, "MM/dd/yyyy") & " " & shorsalven
   'If Trim(Txtoritra.Text) = "O" Then
   'End If
   
   'WECC 0813/2019 CAU No 1254
   '**********************************************************************
'   If MskMultas.text <> "" Or MskMultas.text <> "0" Then
'        If s_Num_Multa_Pagada <> 0 Then
'
'            QRY = "UPDATE opmultas " _
'                    & "SET muconducep = " & Txtnumcon.text & " " _
'                    & "WHERE munumero = " & s_Num_Multa_Pagada & " "
'            Dbs.Execute (QRY)
'        End If
'   End If
   
'   QRY = "SELECT temnumero1 AS vrPermanencia, temnumero2 AS vrMultas, temnumero3 AS numMulta " _
'         & "FROM optempo1 " _
'        & "WHERE temtext4 = '" & Trim(Txtnumpla.text) & "' " _
'          & "AND tempc = '" & Trim(Pc) & "'"
'
'  Set Tbl1 = Dbs.Execute(QRY)
'
'  If Not Tbl1.EOF Then
'
'    Do While Tbl1.EOF = False
'        If Tbl1!vrPermanencia <> 0 And Tbl1!vrMultas = 0 Then
'            QRY1 = "UPDATE coparqueop " _
'                        & "SET patasapag = " & Txtnumcon.text & " " _
'                        & "WHERE panumero = " & Tbl1!numMulta & " "
'            Dbs.Execute (QRY1)
'
'        ElseIf Tbl1!vrPermanencia = 0 And Tbl1!vrMultas <> 0 Then
'            QRY1 = "UPDATE opmultas " _
'                        & "SET mutasapag = " & Txtnumcon.text & " " _
'                        & "WHERE munumero = " & Tbl1!numMulta & " "
'            Dbs.Execute (QRY1)
'        End If
'        Tbl1.MoveNext
'    Loop
'
'  End If

'   QRY = "DELETE FROM optempo1 " _
'        & "WHERE temtext4 = '" & Trim(Txtnumpla.text) & "' " _
'          & "AND tempc = '" & Pc & "'"
'   Dbs.Execute QRY
  
    'Sebastián Rondón - Marzo 13 de 2024
    'Se valida nueva forma de pago de multas y permanencias
    QRY = "SELECT rcvalper AS vrPermanencia, rcvalmul AS vrMultas, rcnumero AS numMulta, rcplaca, rcnomcaj " _
            & "FROM opreccaja " _
            & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
            & "AND rcnomcaj = '" & Trim(Pc) & "' " _
            & "AND rcfecha = TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY') "
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        Do While Tbl1.EOF = False
            If Tbl1!vrPermanencia <> 0 And Tbl1!vrMultas = 0 Then
                QRY1 = "UPDATE coparqueop " _
                            & "SET patasapag = " & Txtnumcon.text & " " _
                            & "WHERE panumero = " & Tbl1!numMulta & " " _
                            & "AND paestado = 'C' "
                Dbs.Execute (QRY1)
            ElseIf Tbl1!vrPermanencia = 0 And Tbl1!vrMultas <> 0 Then
                QRY1 = "UPDATE opmultas " _
                            & "SET mutasapag = " & Txtnumcon.text & " " _
                            & "WHERE munumero = " & Tbl1!numMulta & " " _
                            & "AND muestado = 'Pagada' "
                Dbs.Execute (QRY1)
            End If
            Tbl1.MoveNext
        Loop
    End If
   
    QRY = "DELETE FROM opreccaja " _
                & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
                & "AND rcnomcaj = '" & Trim(Pc) & "' " _
                & "AND rcfecha = TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY') "
    Dbs.Execute (QRY)
    
    'Sebastián Rondón - Marzo 14 de 2024
    'Por orden del Jefe Álvaro, si se encuentra una permanencia o multa anterior sin tasa, le asigna la tasa de uso actual
    'Para que se envié el pago a la facturación electrónica de la DIAN
'    QRY = "SELECT rcvalper AS vrPermanencia, rcvalmul AS vrMultas, rcnumero AS numMulta, rcplaca, rcnomcaj " _
'            & "FROM opreccaja " _
'            & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
'            & "AND rcfecha < TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY') "
'    Set Tbl1 = Dbs.Execute(QRY)
'
'    If Not Tbl1.EOF Then
'        Do While Tbl1.EOF = False
'            If Tbl1!vrPermanencia <> 0 And Tbl1!vrMultas = 0 Then
'                'Primero buscamos si la permanencia no tiene tasa de uso asociada
'                qry2 = "SELECT panumero, patasapag FROM coparqueop WHERE panumero = " & Tbl1!numMulta & " "
'                Set Tbl2 = Dbs.Execute(qry2)
'
'                If Not Tbl2.EOF Then
'                    If Tbl2!patasapag = 0 Then
'                        QRY1 = "UPDATE coparqueop " _
'                            & "SET patasapag = " & Txtnumcon.text & " " _
'                            & "WHERE panumero = " & Tbl1!numMulta & " "
'                        Dbs.Execute (QRY1)
'                    End If
'                End If
'            ElseIf Tbl1!vrPermanencia = 0 And Tbl1!vrMultas <> 0 Then
'                'Primero buscamos si la multa no tiene tasa de uso asociada
'                qry2 = "SELECT munumero, mutasapag FROM opmultas WHERE munumero = " & Tbl1!numMulta & " "
'                Set Tbl2 = Dbs.Execute(qry2)
'
'                If Not Tbl2.EOF Then
'                    If Tbl2!mutasapag = 0 Then
'                        QRY1 = "UPDATE opmultas " _
'                            & "SET mutasapag = " & Txtnumcon.text & " " _
'                            & "WHERE munumero = " & Tbl1!numMulta & " "
'                        Dbs.Execute (QRY1)
'                    End If
'                End If
'            End If
'            Tbl1.MoveNext
'        Loop
'    End If
    
'    QRY = "DELETE FROM opreccaja " _
'                & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
'                & "AND rcfecha < TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY') "
'    Dbs.Execute (QRY)
   '**********************************************************************
   
   If G_Tipo_Terminal = "SUR" Or G_Tipo_Terminal = "NORTE" Then
      QRY = "DELETE FROM cotmpcon WHERE tcplaca = '" & Txtnumpla.text & "' "
      Dbs.Execute QRY
   End If
   
   'On Error GoTo transac
   '********
   'wecc sep26
   'actualiza el numero de conduce en la tabla de cajas
   QRY = "UPDATE cocaja " _
               & "SET caconcaj = caconcaj + 1 "
   
   'WECC 07/15/2019
   'CAU2 1258
   
   QRY = QRY & "WHERE canomcaj =  '" & txtcaja.text & "' "
   
   Dbs.Execute QRY
   
   If pasoco = 1 Then
      numsal = 1
      Dbs.Execute "INSERT INTO ophojveh " _
                & "(hvplaca, hvnumint, hvcodemp, hvpropie, " _
                & "hvconact, hvnumpas, hvestado, hvtipveh, " _
                & "hvfecultpru, hvprueba, hvestsel, hvnumsal, hvconduct, hvconduct1) VALUES ('" _
                & Txtnumpla.text & "'," & TxtNumInt.text & ",'" & Txtcodemp.text & "','PENDIENTE','" _
                & sconductor & "'," & numpas & ",'" & estado & "','" & Txttipveh.text & "',TO_DATE('" _
                & Mskfecha.text & "','MM/dd/yyyy'),'No','No'," & numsal & "," & txtCedula.text & ",'" & TxtCedula1.text & "')"
   Else
      
      QRY = "UPDATE ophojveh " _
                   & "SET  hvnumint = '" & TxtNumInt.text & "'," _
                   & "     hvcodemp = '" & Txtcodemp.text & "'," _
                   & "     hvconact = '" & sconductor & "'," _
                   & "     hvnumsal = hvnumsal + 1, " _
                   & "     hvconduct = " & txtCedula.text & ", " _
                   & "     hvestsel = 'No', " _
                   & "     hvconduct1 = '" & TxtCedula1.text & "', " _
                   & "     hvoritra = '" & Trim(Txtoritra.text) & "', hvestviaje = 'S', " _
                   & "     hvultimatasauso = " & Txtnumcon.text & ", hvultimacaja = '" & Trim(txtcaja.text) & "' "

         '& "     hvtipveh = '" & Txttipveh.Text & "', "


      If Aviso_Alistamiento = True Then
         QRY = QRY & ", hvfecalistamiento = TO_DATE('" & Mskfecha.text & "','MM/dd/yyyy') "
      End If
                 
      QRY = QRY & "WHERE hvplaca = '" & Txtnumpla & "'"
      
      Dbs.Execute QRY
   End If
   
   shorsalven = Format(sfecsalveh, "MM/dd/yyyy") & " " & shorsalven
   QRY1 = "INSERT INTO cosalveh(" _
            & "svnumero, svplaca, svfecha, svcodrut, " _
            & "svterminal, svcedula, svcedula1) VALUES (" _
            & Txtnumcon.text & ",'" & Txtnumpla.text & "', to_date('" & shorsalven & "','MM/DD/YYYY hh24:mi:ss'), '" & Trim(Txtcodrut.text) & "', '" _
            & Mid(G_Tipo_Terminal, 1, 1) & "', '" & txtCedula.text & "', '" & TxtCedula1.text & "')"

    Dbs.Execute QRY1
       
   Dbs.CommitTrans
   
   On Error GoTo err_mensaje
     
   'WECC 06/14/2019 CAU2 No 1101 se agrega que solo cuando se compre la tasa de uso en origen y _
   'las multas dejarla con la cantidad de avisos  permitidos, en la cancelación de las multas se descuenta un aviso
   
   If Txtoritra.text = "O" Then
      fqry01 = "SELECT * " _
               & "FROM opmultas " _
              & "WHERE munumpla = '" & Txtnumpla.text & "' " _
                & "AND muestado = 'Pendiente' "
      
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         QRY = "UPDATE opmultas "
                   
         If L_DescontarAviso = "S" Then
            QRY = QRY & "SET muaviso = " & AviMulta & ""
         Else
            QRY = QRY & "SET muaviso = muaviso + 1 "
            
         End If
         QRY = QRY & "WHERE munumpla = '" & Txtnumpla.text & "' " _
                   & "AND muestado = 'Pendiente' "
         Dbs.Execute QRY
   '      Do Until Ftb01.EOF
   '         avisos = Ftb01!muaviso + 1
   '
   '         QRY = "UPDATE opmultas " _
   '                & "SET muaviso = " & avisos & " " _
   '              & "WHERE munumpla = '" & Txtnumpla.Text & "' " _
   '                & "AND muestado = 'Pendiente' " _
   '                & "AND munumero = " & Ftb01!munumero & " "
   '
   '         Dbs.Execute QRY
   '
   '         Ftb01.MoveNext
   '      Loop
         
      End If
      Ftb01.Close
   End If
   '**********************************************************************************
   '**********************************************************************************
   'WECC CAU No 1103 Grabar el número de tasa de uso en un txt
   NOMDIR = App.Path & "\CONSECUTIVO_" & Pc & ".TXT"
             
   NOMBRE = Dir(NOMDIR)
 
   If NOMBRE <> "" Then
      Kill NOMDIR
   End If
 

   On Error Resume Next
   Open NOMDIR For Append As #2
   If Err Then
      MsgBox Error$, 48
      Capconduce.Enabled = True
      Close #2
   Else
   
      Print #2, Txtnumcon.text & "-" & Txtnumpla.text & "-" & Pc
      Close #2
   End If
   '**********************************************************************************
   '**********************************************************************************
   
   
   
   'If G_turno > 1 Then
   '   fecha = Mskfecha.Text
   'ElseIf G_turno = 1 Then
   '   horcierre1 = "21:00"
   '   horcierre2 = "23:59"
   '
   '   If CDate(Format(Mskhora.Text, "hh:mm")) > CDate(Format(horcierre1, "hh:mm")) And CDate(Format(Mskhora.Text, "hh:mm")) < CDate(Format(horcierre2, "hh:mm")) Then
      
   '      fecha = Mskfecha.Text
   '   Else
   '      fecha = DateAdd("d", -1, Mskfecha.Text)
   '   End If
   'End If
   
   'fecha = "#" & fecha & "#"
   
   '04/24/2002 wecc
   
   'If prealco = 1 Then
   '   If alcoholemia = "S" Then
   '      If Txtoritra.Text = "O" Then
   '         dbs.Execute "UPDATE coconsol " _
                         & "SET csrealori = csrealori + 1 " _
                       & "WHERE csfecha = " & fecha & " " _
                         & "AND cscodemp = '" & Txtcodemp.Text & "' "
            
   '      Else
   '         dbs.Execute "UPDATE coconsol " _
                         & "SET csrealtra = csrealtra + 1 " _
                      & "WHERE csfecha = " & fecha & " " _
                        & "AND cscodemp = '" & Txtcodemp.Text & "' "
   '      End If
   '   End If
   'End If
   
   'If influencia = "Si" Then
   '   dbs.Execute "UPDATE coconsol " _
                   & "SET csinflue = csinflue + 1 " _
                 & "WHERE csfecha = " & fecha & " " _
                   & "AND cscodemp = '" & Txtcodemp.Text & "' "
   'End If
   
   
   'wecc 07/02/2002 actualiza la fecha en que se compro el conduce
   '******************************
   'qry = "UPDATE coingsal " _
          & "SET isfeccon  = #" & Mskfecha.Text & "# " _
        & "WHERE isplaca = '" & Txtnumpla.Text & "' " _
          & "AND isnull(isfeccon) "
   
   'dbs.Execute qry
   '******************************
   
   'MsgBox "Conduce Ingresado", 64, "O.K."
   
   'Sebastian Rondon / 11 Agosto de 2023
   'Registro de venta tasa condicionada
   If L_Bandera_Tasa_Condicionada = True Then
        'Se aumenta el número del aviso
        QRY = "UPDATE ophojveh SET hvnumaviprot = hvnumaviprot + 1 WHERE hvplaca = '" & Txtnumpla.text & "' "
        Dbs.Execute (QRY)
        
        'Se ingresa el registro en la nueva tabla
        QRY = "INSERT INTO VEHCONCOND (cdplaca, cdcodemp, cdfecha, cdhora, cdterminal, cdtasauso, cdcoddes) " _
                    & "VALUES ('" & Txtnumpla.text & "', '" & Txtcodemp.text & "', " _
                    & "TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'), " _
                    & "TO_DATE('" & Mskhorsis.text & "','HH24:MI:SS'), '" & Mid(G_Tipo_Terminal, 1, 1) & "', " _
                    & Txtnumcon.text & ", '" & Txtcodrut.text & "' )"
        Dbs.Execute (QRY)
   End If
   
   If L_Bandera_Venta_Condicionada = True Then
        QRY = "UPDATE ophojveh SET hvbandvencon = 0 WHERE hvplaca = '" & Txtnumpla.text & "' "
        Dbs.Execute (QRY)
   End If
   
   'Sebastián Rondon Noviembre 02 - 2022
   'Verificamos si pago Tag el vehículo
   
   If L_Bandera_CobroTag = True Then
   
    If ChkTag.Value = 1 Then
        
        QRY = "SELECT psval FROM geparsis " _
                    & "WHERE pscod = 'HORASIS' "
                    
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
            
            Fecha = Mid(Tbl1!psval, 1, 10)
            
            QRY = "UPDATE optag SET " _
                & "tafecpag = TO_DATE('" & Fecha & "', 'MM/DD/YYYY'), " _
                & "tahorpag = TO_DATE('" & horven & "', 'hh24:mi:ss'), " _
                & "taturno = '" & G_Turno & "', " _
                & "tanomcaj = '" & txtcaja.text & "', " _
                & "tanumtasa = " & Txtnumcon.text & ", " _
                & "tacerrado = 'N' " _
                & "WHERE taplaca = '" & Txtnumpla.text & "' " _
                & "AND taestado = 'A' " _
                & "AND tanumero = " & L_Bandera_NumeroTag & " "
                
            Dbs.Execute (QRY)
            
        End If
        Tbl1.Close
        
'        respuesta = MsgBox("¿Desea Imprimir El Tag?", 64 + vbYesNo, "Imprimir")
'        If respuesta = vbYes Then
'        Imprimir_Zebra_Tag
        
        End If
    Else
        
        QRY = "UPDATE optag SET " _
                & "taaviso = taaviso + 1 " _
                & "WHERE taplaca = '" & Txtnumpla.text & "' " _
                & "AND taestado = 'A' "
                
        Dbs.Execute (QRY)
   
   End If
   
   
   Btngrabar.Enabled = False
   
   
   'Camilo Campos - Mayo 05 de 2025
   'Debido a la implementación de rodamiento las tasas de uso deben imprimirse allá mismo, pero igual se deja parametrizado en geparsis
   'la posibilidad de imprimir, con el código PARIMPTASA
   
   If G_impresion = "S" Then
   
        respuesta = MsgBox("Desea Imprimir", 64 + vbYesNo, "Imprimir")
        If respuesta = vbYes Then
           Btnimprimir_Click
        End If
   
   End If
   
   '***************************
   'Sebastian Rondon Jun/13/2022
   'Mandamos información al Web Service de Control y Salidas
   
'    NumTasa = Trim(Txtnumcon.Text)
'    Placa = Trim(Txtnumpla.Text)
'
'    'Verificamos en que tipo de terminal esta
'    If G_Tipo_Terminal = "CENTRAL" Then
'        Tterminal = "C"
'    Else
'        If G_Tipo_Terminal = "NORTE" Then
'            Tterminal = "N"
'        Else
'            Tterminal = "S"
'        End If
'    End If
    
    '****************************************
    'Sebastián Rondón - Oct 06 - 2022
    'Se parcha todo el armado y envió del JSON al Web Services de Protech
    'Por motivos de constante falla de comunicación
    
'    Call CreandoJsonControlSalida(NumTasa, Placa, Tterminal)
'
'    sBandera = CadenaJSON
'    sParametro = ""
'    sUrl = L_UrlApiControlSalida
'    sEvento = "INS"
'
'    If L_IntegracionProtech = "S" Then
'      sRespuesta = Fn_Web_Sevices_Protech(L_UrlServer, sUrl, sEvento, sBandera, sParametro)
'
'      If sRespuesta <> "0" Then
'         sFechaSistemaProtech = Fn_Parametros("HORASIS", 1)
'
'         QRY = "INSERT INTO logprotech(lotipo, lovalor, lodescripcion, lofecha)" _
'         & "VALUES ('V','" & Txtnumpla.Text & "','" & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta & "',TO_DATE('" & sFechaSistemaProtech & "','MM/dd/yyyy HH24:mi:ss'))"
'
'         Dbs.Execute QRY
'
'         MsgBox "Se presento error en el webservices de Protech por: " & Chr(13) & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta, 48, "Advertencia"
'      End If
'   End If
    '****************************************
   
   
   '**************************
   
   
   'wecc 04/25/2002************
   'insertamos en la tabla de alcoholemia
   'los vehiculos que se presentaron al examen de alcoholemia
   
'   If prealco = 1 Then
'      If alcoholemia = "S" Then
'         horaprueba = CDate(horaprueba)
'         fecha = Mskfecha.Text
'
'         aprobo = "Si"
'         estado = "Revisado"
'         Guardo_Servidor = "N"
'     '    If G_BaseServidor = True Then
'      '      dbsX.Execute "INSERT INTO opalcolemia " _
'      '             & "(alhora, alplaca, " _
'      '             & "alempresa, alclase, alvalor, " _
'      '             & "alaprobo, aloritra, alestado, " _
'      '             & "alfecha) " _
'      '             & "VALUES (#" & Format(horaprueba, "hh:mm") & "#,'" & Txtnumpla.Text & "','" _
'      '             & txtcodemp.Text & "','" & Txttipveh.Text & "','" & Mskrecibo.Text & "','" _
'      '             & aprobo & "','" & Txtoritra.Text & "','" & estado & "', #" _
'      '             & fecha & "#)"
'      '
'      '        Guardo_Servidor = "S"
'      '   End If
'
'         horaprueba = Format(horaprueba, "HH:mm")
'         dbs.Execute "INSERT INTO opalcolemia " _
'                & "(alhora, alplaca, " _
'                & "alempresa, alclase, alvalor, " _
'                & "alaprobo, aloritra, alestado, " _
'                & "alfecha) " _
'                & "VALUES (to_date('" & horaprueba & "','hh24:mi:ss') ,'" & Txtnumpla.Text & "','" _
'                & Txtcodemp.Text & "','" & Txttipveh.Text & "','" & Mskrecibo.Text & "','" _
'                & aprobo & "','" & Txtoritra.Text & "','" & estado & "', to_date('" & fecha & "','MM/dd/yyyy') )" _
'                & ""
'
'         prealco = 0
'      End If
'   End If

   '***************************
   
   Btnlimpiar_Click
   
   Exit Sub
   
err_Btngrabar:
    If Err.Number = 52 Then
      If Desc_Server = True Then
      MsgBox "Se Ha Desconectado del Servidor", 48, "Desconectado"
      End If
      Desc_Server = False
      G_BaseServidor = False
      Resume Next
    End If
    
    Btnlimpiar_Click
    MsgBox Err.Description, 16, "Error " & Err.Number
    
    Exit Sub
    
err_mensaje:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
   Btnlimpiar_Click
   Exit Sub
err_Btngrabar_Dbx:
   G_BaseServidor = False
   Resume Next
transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub


Private Sub Btnimprimir_Click()

   hora = Mid(Mskhora.text, 12, 15)
   hora = Format(hora, "hh:mm")
   
   Msktottas.text = CDbl(Mskvalor.text) + CDbl(Mskrecibo.text) '+ CDbl(Mskparqueadero.Text)
   totaltasa = Msktottas.FormattedText
      
   Capconduce.Enabled = False
   
   On Error GoTo ERR_IMPRIMIR
   If G_ImpCon = "E" Then
       
      NOMDIR = G_DirImpBol & "IMPRIMIR.PRN"

      NOMBRE = Dir(NOMDIR)
      
      If NOMBRE <> "" Then
         Kill NOMDIR
      End If
      
      
      On Error Resume Next
      Open NOMDIR For Append As #2
      If Err Then
         MsgBox Error$, 48
         Close #2
         Capconduce.Enabled = True
         Exit Sub
      Else
          
          On Error GoTo ERR_IMPRIMIR
          
            
            a = ""
            Print #2, a
            
            b = crecon & "     " & Txtnumcon.text & "                                      " & Txtnumcon.text & "                                           " & Txtnumcon.text
            Print #2, b
            Print #2, a
      
            b = Format(Mskfecha.text, "MM/dd/yy") & " " & hora & " " & TxtNumInt.text & "  " & Txtnumpla.text & "        " & Mskfecha.text & "    " & hora & "    " & TxtNumInt.text & "      " & Txtnumpla.text & "        " & Mskfecha.text & "    " & hora & "    " & TxtNumInt.text & "      " & Txtnumpla.text
            Print #2, b
            Print #2, a
      
            If Len(Lbltipveh.Caption) < 6 Then
               tipservi = Lbltipveh.Caption
               For i = 0 To 6 - Len(Lbltipveh.Caption)
                  tipservi = tipservi & " "
               Next i
            Else
               tipservi = Mid(Lbltipveh.Caption, 1, 6)
            End If
            If ConNombre = "" Then
               codempresa = Trim(Txtcodemp.text) & "-" & Trim(Lblnomemp.Caption)
            Else
               codempresa = ConNombre
            End If
      
            If Len(codempresa) >= 30 Then
              codempresa = Mid(codempresa, 1, 30)
            Else
              For i = Len(codempresa) To 30
                  espacios = espacios & " "
              Next
              codempresa = codempresa & espacios
            End If
            codempresa = codempresa & "  " & crecon
            longitud = Len(codempresa)
      
            If longitud >= 40 Then
              codempresa = Mid(codempresa, 1, 40)
            Else
              For i = longitud To 40
                  espacios = espacios & " "
              Next
              codempresa = codempresa & espacios
            End If
            If ConNombre = "" Then
               b = Mid(Txtcodemp.text & "-" & Lblnomemp.Caption, 1, 14) & " " & Txtcodrut.text & " " & tipservi & "       " & codempresa & "        " & codempresa
            Else
               If Len(ConNombre) >= 14 Then
                  ConNombre = Mid(ConNombre, 1, 14)
              Else
                For i = Len(ConNombre) To 14
                    espacios = espacios & " "
                Next
                ConNombre = ConNombre & espacios
              End If
               b = ConNombre & " " & Txtcodrut.text & " " & tipservi & "       " & codempresa & "        " & codempresa
            End If
            Print #2, b
            Print #2, a
      
            If Len(Login) < 10 Then
               funcio = Login
                For i = 0 To 10 - Len(Login)
                  funcio = funcio & " "
               Next i
            Else
               funcio = Login
            End If
            ruta = Trim(Txtcodrut.text) & "-" & Trim(Lblnomrut.Caption)
      
            If Len(ruta) > 25 Then
              ruta = Mid(ruta, 1, 25)
            Else
              For i = Len(ruta) To 25
                  ruta = ruta & " "
              Next
            End If
      
            nivel = Trim(Lbltipveh.Caption)
            If Len(nivel) > 15 Then
              nivel = Mid(nivel, 1, 15)
            Else
              For i = Len(nivel) To 15
                  nivel = nivel & " "
              Next
            End If
      
            b = Pc & "        " & G_Turno & "    " & funcio & "   " & ruta & " " & nivel & "     " & ruta & " " & nivel
            Print #2, b
            Print #2, a
      
            b = "         " & presenprueba & "        " & Mskvalor.FormattedText & "       " & funcio & Pc & "  " & G_Turno & "               " & Mskvalor.FormattedText & "        " & funcio & Pc & "   " & G_Turno & "            " & Mskvalor.FormattedText
            Print #2, b
            'Print #2, a
      
            b = "                   " & Mskrecibo.FormattedText & "                     " & presenprueba & "                  " & Mskrecibo.FormattedText & "                     " & presenprueba & "                  " & Mskrecibo.FormattedText
            Print #2, b
            Print #2, a
      
            b = "                   " & MskParqueadero.FormattedText & "                                             " & MskParqueadero.FormattedText & "                                                " & MskParqueadero.FormattedText
            Print #2, b
            b = "                   " & totaltasa & "                                       " & totaltasa & "                                       " & totaltasa
            Print #2, b
      
      
            Print #2, a
            Print #2, a
            Print #2, a
            contador = contador + 1
            If contador = 2 Then
               'Print #2, a
               contador = 0
            End If
      
      
          Close #2
      
       End If
    
      NOMDIR = G_DirImpBol & "IMPRIMIR.BAT"
    
      NOMBRE = Dir(NOMDIR)
    
       If NOMBRE <> "" Then
          Kill NOMDIR
       End If
      
      
       On Error Resume Next
       Open NOMDIR For Append As #2
      
       If Err Then
          MsgBox Error$, 48
          Close #2
          Capconduce.Enabled = True
          Exit Sub
       Else
           'If G_PrinterPort = "LPT1:" Then
             ' a = "copy " & G_DirImpBol & "imprimir.prn PRN"
              a = "copy " & G_DirImpBol & "imprimir.prn PRN"
           'Else
           '   A = "copy C:\ARCHIV~1\TERMINAL\CONDUCES\imprimir.prn " & G_PrinterPort
           'End If
      
           Print #2, a
           Close #2
        End If
        Shell G_DirImpBol & "IMPRIMIR.BAT", vbHide
      
      
   ElseIf G_ImpCon = "Z" Then
      On Error GoTo ERR_IMPRIMIR
      'Imprimir en la zebra
      Imprimir_Zebra
      
      If L_Bandera_CobroTag = True Then
    
        If ChkTag.Value = 1 Then
            respuesta = MsgBox("¿Desea Imprimir El Tag?", 64 + vbYesNo, "Imprimir")
            If respuesta = vbYes Then
                Imprimir_Zebra_Tag "Z", "N"
            End If
        End If

     End If
     
     If L_Bandera_ReimpresionTag = True Then
     
        QRY = "SELECT tanumero, tacodtag, taconcep FROM optag " _
            & "WHERE tafecpag IS NOT NULL " _
            & "AND taplaca = '" & Txtnumpla.text & "' " _
            & "AND tanumtasa = " & Txtnumcon.text & " "
            
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            respuesta = MsgBox("¿Desea reimprimir el Tag?", 64 + vbYesNo, "Reimprimir")
            If respuesta = vbYes Then
                Imprimir_Zebra_Tag "Z", "S"
            End If
        End If
     End If
       
   ElseIf G_ImpCon = "C" Then
      On Error GoTo ERR_IMPRIMIR
      'Imprime por crystal
      Imprimir_Crystal
   ElseIf G_ImpCon = "N" Then
      On Error GoTo ERR_IMPRIMIR
      'Imprime por exe Net
      Imprimir_Exe_Net
      
      If L_Bandera_CobroTag = True Then
    
        If ChkTag.Value = 1 Then
            respuesta = MsgBox("¿Desea Imprimir El Tag?", 64 + vbYesNo, "Imprimir")
            If respuesta = vbYes Then
                Imprimir_Zebra_Tag "N", "N"
            End If
        End If
     End If
     
    If L_Bandera_ReimpresionTag = True Then
        QRY = "SELECT tanumero, tacodtag, taconcep FROM optag " _
            & "WHERE tafecpag IS NOT NULL " _
            & "AND taplaca = '" & Txtnumpla.text & "' " _
            & "AND tanumtasa = " & Txtnumcon.text & " "
            
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            respuesta = MsgBox("¿Desea reimprimir el Tag?", 64 + vbYesNo, "Reimprimir")
            If respuesta = vbYes Then
                Imprimir_Zebra_Tag "N", "S"
            End If
        End If
    End If
   End If
   If L_Bandera_Copia = True Then
      Btnlimpiar_Click
      On Error Resume Next
      Txtnumpla.Enabled = True
      'Txtnumpla.SetFocus
      On Error GoTo ERR_IMPRIMIR
   End If
   
   Capconduce.Enabled = True
   Exit Sub
ERR_IMPRIMIR:
   Capconduce.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Imprimir_Zebra()
   hora = Mid(Mskhora.text, 12, 15)
   hora = Format(hora, "hh:mm")
      
   totaltasa = Msktottas.FormattedText
   
   If CDbl(Mskrecibo.text) = 0 Then
      presenprueba = "No"
   End If
     
   presenprueba = presenprueba & "-" & Lbloritra.Caption
   sNit = "NIT.860.052.155-6"
   'sEmpVeh = Txtcodemp.Text & "-" & Lblnomemp.Caption & "   TERMINAL " & G_Tipo_Terminal
   
   
   sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & Mid(G_Tipo_Terminal, 1, 1)
   
   If G_Imprime_Tasa_Uso = "E" Then
      sNombreEmpresa = sEmpVeh
   Else
      If L_Operador <> "" Then
         sNombreEmpresa = Traer_Descripcion(L_Operador, "OPERADORES DE SERVICIO")
      Else
         sNombreEmpresa = sEmpVeh
      End If
   End If
   If Len(Txtnumcon.text) = 1 Then
      LNumConBarra = "00" & Txtnumcon.text
   ElseIf Len(Txtnumcon.text) = 2 Then
      LNumConBarra = "0" & Txtnumcon.text
   Else
      LNumConBarra = Txtnumcon.text
   End If
   sEmpVeh = sNombreEmpresa
   'sAlcoholimetria = "PGS_SVIAL.ALCH"
   sAlcoholimetria = "PRGSALYSEGVIAL"
   sNumero = Txtnumcon.text
   sruta = Txtcodrut.text & "-" & Lblnomrut.Caption
   sNivelServicio = Lbltipveh.Caption
   sValorParqueadero = MskParqueadero.FormattedText
   sValorTasa = Mskvalor.FormattedText
   sValorAlcoholimetria = Mskrecibo.FormattedText
   sNumInterno = TxtNumInt.text
   sFecha = Format(Mskfecha.text, "MMM/dd/yyyy")
   sPlaca = Txtnumpla.text
   
   s2Conductores = ""
   If Aviso_2_Conductores = False Then
     s2Conductores = "DESTINO NECESITA DOS CONDUCTORES"
   End If
   If G_Cantidad_Parqueaderos > 0 Then
      sPermanencia1 = "Permanencias"
      sPermanencia2 = "Pendientes " & G_Cantidad_Parqueaderos
      sPermanencia3 = "Aviso " & G_Avisos_Parqueadero & " de " & L_AvisoParqueadero
   Else
      sPermanencia1 = ""
      sPermanencia2 = ""
      sPermanencia3 = ""
   End If
   Open "LPT1" For Output As #1
   'WECC 05/13/2014 Esta linea es cuando la papeleria tiene muesca
   Print #1, "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ"
   
   'ESTA ES LA LINEA CUANDO LA PAPELERIA TIENE LA BANDA
   'Print #1, "^XA~TA000~JSN^LT43^MMT^MNM^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ"
   
   Print #1, "~DG000.GRF,02048,016,"
   Print #1, ",:::::::::::::::::::::::::T010H010,,U07FHF0,T02F0H0FC0,S01F10H0171,S0280K08,R0140L03,R0680M080,Q01D0N070,Q0180N028,Q060O01C,Q0E0P0E,P0180P03,P020Q0180,P040R040,P0C0R0A0,N0H180R030,O020S018,O060T08,O040T0C,O0C0T04,O080T02,N0180T03,N030U01,N030U0180,N020V080,N060V040,N080EE80M0EA0H080,N0C0FFC0L01FF0H040,N080FB80L01BA0H020,M0180FFC0L01FF0H020,N080EE80M0HEI020,M0100FFC010H0101FF01030,P0HB80L01FE0H020,M03007FC0L01FF0H030,M0200EAERE0,M0300FTF010,M02002BFBFBFBFBFBFBE010,M03007FSF010,M02002AAEQEH08,M07003FSF018,M02001ABBFOFB018,M060H0TF018,M020H0HAREH08,M06013FSF018,M02003FFBFBFBFBFBFBB018,M07007FSF018,M0200ESEA0,L01300FTF010,M0200FFBFIFBFKFA010,M0301FFC7FPF010,M0200EE80M0EA0H020,M0101FF80L01FF0H020,M0100FB80L01FA0H020,M0181FF0M01FE0H020,M0181AE80M0HEI020,M0181FF80L01FF0H040,N080BA80L01FE0,N0C0V040,N0E0V080,M0160U0180,N020U01,N070U01,N010U02,N0180T04,O08,O040T0C,O060T08,O070S010,O0380R020,O01C0R040,P0C0R080,P070Q01,P0380P02,P01C0P04,Q060P08,Q030O030," _
           & "Q0180N0E0,R070M03,R0180L0E,R01D0K05810,S03A0J0A0,T0H7H017,U0BFEA0,W0"
   Print #1, "1,,::::::::::::::::~DG001.GRF,02048,016,"
   Print #1, ",:::::::::::::::::::::::::R010H010,,S07FHF0,R02F0H0FC0,Q01F10H0171,Q0280K08,P0140L03,P0680M080,O01D0N070,O0180N028,O060O01C,O0E0P0E,N0180P03,N020Q0180,N040R040,N0C0R0A0,L0H180R030,M020S018,M060T08,M040T0C,M0C0T04,M080T02,L0180T03,L030U01,L030U0180,L020V080,L060V040,L080EE80M0EA0H080,L0C0FFC0L01FF0H040,L080FB80L01BA0H020,K0180FFC0L01FF0H020,L080EE80M0HEI020,K0100FFC010H0101FF01030,N0HB80L01FE0H020,K03007FC0L01FF0H030,K0200EAERE0,K0300FTF010,K02002BFBFBFBFBFBFBE010,K03007FSF010,K02002AAEQEH08,K07003FSF018,K02001ABBFOFB018,K060H0TF018,K020H0HAREH08,K06013FSF018,K02003FFBFBFBFBFBFBB018,K07007FSF018,K0200ESEA0,J01300FTF010,K0200FFBFIFBFKFA010,K0301FFC7FPF010,K0200EE80M0EA0H020,K0101FF80L01FF0H020,K0100FB80L01FA0H020,K0181FF0M01FE0H020,K0181AE80M0HEI020,K0181FF80L01FF0H040,L080BA80L01FE0,L0C0V040,L0E0V080,K0160U0180,L020U01,L070U01,L010U02,L0180T04,M08,M040T0C,M060T08,M070S010,M0380R020,M01C0R040,N0C0R080,N070Q01,N0380P02,N01C0P04,O060P08,O030O030,O0180N0E0," _
           & "P070M03,P0180L0E,P01D0K05810,Q03A0J0A0,R0H7H017,S0BFEA0,U0"
   Print #1, "1,,::::::::::::::::^XA^LL0743"
   Print #1, "^PW815"
   Print #1, "^FT288,128^XG000.GRF,1,1^FS"
   Print #1, "^FT704,128^XG001.GRF,1,1^FS"
   Print #1, "^BY2,2.5,66^FT20,135^B3R,N,,Y,N"
   Print #1, "^FD" & LNumConBarra & "^FS"
   Print #1, "^FT94,640^A0R,20,21^FH\^FD" & sValorParqueadero & "^FS"
   Print #1, "^FT146,629^A0R,20,21^FH\^FD" & sValorTasa & "^FS"
   Print #1, "^FT119,640^A0R,20,21^FH\^FD" & sValorAlcoholimetria & "^FS"
   Print #1, "^FT119,320^A0R,20,21^FH\^FD" & G_Turno & "^FS"
   Print #1, "^FT342,134^A0R,23,21^FH\^FD" & G_NomEmpresa & "^FS"
   Print #1, "^FT363,500^A0R,17,16^FH\^FDTasa de^FS"
   Print #1, "^FT342,500^A0R,17,16^FH\^FDUso No.^FS"
   Print #1, "^FT281,262^A0R,20,16^FH\^FDHORA^FS"
   Print #1, "^FT228,260^A0R,20,19^FH\^FD" & sEmpVeh & "^FS"
   Print #1, "^FT199,28^A0R,20,19^FH\^FDRUTA AUTORIZADA^FS"
   Print #1, "^FT228,28^A0R,20,19^FH\^FDEMPRESA:^FS"
   Print #1, "^FT146,188^A0R,20,19^FH\^FDCAJA^FS"
   Print #1, "^FT21,420^A0R,17,16^FH\^FD" & s2Conductores & "^FS"
   Print #1, "^FT17,28^A0R,17,16^FH\^FD" & sPermanencia3 & "^FS"     '---
   Print #1, "^FT43,27^A0R,17,16^FH\^FD" & sPermanencia2 & "^FS"     '|
   Print #1, "^FT68,27^A0R,17,16^FH\^FD" & sPermanencia1 & "^FS"   '---
   Print #1, "^FT97,28^A0R,17,16^FH\^FDPRUEBA DE ALCOHOLIMETRIA^FS"
   Print #1, "^FT199,452^A0R,20,19^FH\^FDNIVEL DE SERVICIO^FS"
   Print #1, "^FT281,566^A0R,20,16^FH\^FDPLACA^FS"
   'Print #1, "^FT256,424^A0R,20,21^FH\^FD" & sNumInterno & "^FS"
    Print #1, "^FT248,457^A0R,28,28^FH\^FD" & sNumInterno & "^FS"
   'Print #1, "^FT281,424^A0R,20,16^FH\^FDNo INT.^FS"
   Print #1, "^FT281,457^A0R,20,16^FH\^FDNo INT.^FS"
   Print #1, "^FT96,300^A0R,17,16^FH\^FD" & presenprueba & "^FS"
   Print #1, "^FT256,262^A0R,20,21^FH\^FD" & hora & "^FS"
   Print #1, "^FT174,452^A0R,20,21^FH\^FD" & sNivelServicio & "^FS"
   Print #1, "^FT57,619^A0R,25,24^FH\^FD" & totaltasa & "^FS"
   Print #1, "^FT95,451^A0R,20,19^FH\^FDT.PERMANENCIA^FS"
   Print #1, "^FT174,28^A0R,20,21^FH\^FD" & sruta & "^FS"
   Print #1, "^FT120,451^A0R,20,19^FH\^FDPRGSALYSEGVIAL^FS"
   Print #1, "^FT147,451^A0R,20,19^FH\^FDTASA DE USO $^FS"
   Print #1, "^FT54,451^A0R,25,24^FH\^FDValor Pagar^FS"
   Print #1, "^FT146,300^A0R,20,19^FH\^FDTURNO^FS"
   Print #1, "^FT118,28^A0R,20,21^FH\^FD" & Login & "^FS"
   Print #1, "^FT118,188^A0R,20,21^FH\^FD" & Pc & "^FS"
   'Print #1, "^FT256,566^A0R,20,21^FH\^FD" & sPlaca & "^FS"
   Print #1, "^FT248,566^A0R,28,28^FH\^FD" & sPlaca & "^FS"
   Print #1, "^FT146,28^A0R,20,19^FH\^FDCAJERO^FS"
   Print #1, "^FT281,28^A0R,20,16^FH\^FDFECHA DE EXPEDICION^FS"
   Print #1, "^FT256,28^A0R,20,21^FH\^FD" & sFecha & "^FS"
   Print #1, "^FT315,186^A0R,20,19^FH\^FD" & sNit & "^FS"
   Print #1, "^FT343,583^A0R,34,33^FH\^FD" & sNumero & "^FS"
   Print #1, "^FT372,159^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
   Print #1, "^BY2,2.5,66^FT428,135^B3R,N,,Y,N"
   Print #1, "^FD" & LNumConBarra & "^FS"
   Print #1, "^FT502,640^A0R,20,21^FH\^FD" & sValorParqueadero & "^FS"
   Print #1, "^FT554,629^A0R,20,21^FH\^FD" & sValorTasa & "^FS"
   Print #1, "^FT527,640^A0R,20,21^FH\^FD" & sValorAlcoholimetria & "^FS"
   Print #1, "^FT527,320^A0R,20,21^FH\^FD" & G_Turno & "^FS"
   Print #1, "^FT750,134^A0R,23,21^FH\^FD" & G_NomEmpresa & "^FS"
   Print #1, "^FT771,500^A0R,17,16^FH\^FDTasa de^FS"
   Print #1, "^FT750,500^A0R,17,16^FH\^FDUso No.^FS"
   Print #1, "^FT689,262^A0R,20,16^FH\^FDHORA^FS"
   Print #1, "^FT636,260^A0R,20,19^FH\^FD" & sEmpVeh & "^FS"
   Print #1, "^FT607,28^A0R,20,19^FH\^FDRUTA AUTORIZADA^FS"
   Print #1, "^FT636,28^A0R,20,19^FH\^FDEMPRESA:^FS"
   Print #1, "^FT554,188^A0R,20,19^FH\^FDCAJA^FS"
   Print #1, "^FT429,420^A0R,17,16^FH\^FD" & s2Conductores & "^FS"
   Print #1, "^FT425,28^A0R,17,16^FH\^FD" & sPermanencia3 & "^FS"       '---
   Print #1, "^FT451,27^A0R,17,16^FH\^FD" & sPermanencia2 & "^FS"       '|
   Print #1, "^FT476,27^A0R,17,16^FH\^FD" & sPermanencia1 & "^FS"       '---
   Print #1, "^FT505,28^A0R,17,16^FH\^FDPRUEBA DE ALCOHOLIMETRIA^FS"
   Print #1, "^FT607,452^A0R,20,19^FH\^FDNIVEL DE SERVICIO^FS"
   Print #1, "^FT689,566^A0R,20,16^FH\^FDPLACA^FS"
   'Print #1, "^FT664,424^A0R,20,21^FH\^FD" & sNumInterno & "^FS"
   Print #1, "^FT656,457^A0R,28,28^FH\^FD" & sNumInterno & "^FS"
   'Print #1, "^FT689,424^A0R,20,16^FH\^FDNo INT.^FS"
   Print #1, "^FT689,457^A0R,20,16^FH\^FDNo INT.^FS"
   Print #1, "^FT504,300^A0R,17,16^FH\^FD" & presenprueba & "^FS"
   Print #1, "^FT664,262^A0R,20,21^FH\^FD" & hora & "^FS"
   Print #1, "^FT582,452^A0R,20,21^FH\^FD" & sNivelServicio & "^FS"
   Print #1, "^FT465,619^A0R,25,24^FH\^FD" & totaltasa & "^FS"
   Print #1, "^FT503,451^A0R,20,19^FH\^FDT.PERMANENCIA^FS"
   Print #1, "^FT582,28^A0R,20,21^FH\^FD" & sruta & "^FS"
   Print #1, "^FT528,451^A0R,20,19^FH\^FDPRGSALYSEGVIAL^FS"
   Print #1, "^FT555,451^A0R,20,19^FH\^FDTASA DE USO $^FS"
   Print #1, "^FT462,451^A0R,25,24^FH\^FDValor Pagar^FS"
   Print #1, "^FT554,300^A0R,20,19^FH\^FDTURNO^FS"
   Print #1, "^FT526,28^A0R,20,21^FH\^FD" & Login & "^FS"
   Print #1, "^FT526,188^A0R,20,21^FH\^FD" & Pc & "^FS"
   'Print #1, "^FT664,566^A0R,20,21^FH\^FD" & sPlaca & "^FS"
   Print #1, "^FT656,566^A0R,28,28^FH\^FD" & sPlaca & "^FS"
   Print #1, "^FT554,28^A0R,20,19^FH\^FDCAJERO^FS"
   Print #1, "^FT689,28^A0R,20,16^FH\^FDFECHA DE EXPEDICION^FS"
   Print #1, "^FT664,28^A0R,20,21^FH\^FD" & sFecha & "^FS"
   Print #1, "^FT723,186^A0R,20,19^FH\^FD" & sNit & "^FS"
   Print #1, "^FT751,583^A0R,34,33^FH\^FD" & sNumero & "^FS"
   Print #1, "^FT780,159^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
   Print #1, "^PQ1,0,1,Y^XZ"
   Print #1, "^XA^ID000.GRF^FS^XZ"
   Print #1, "^XA^ID001.GRF^FS^XZ"
   Close #1
End Sub

Public Sub Imprimir_Zebra_Tag(ByVal f As String, ByVal d As String)

    hora = Mid(Mskhora.text, 12, 15)
    hora = Format(hora, "hh:mm")
    
   ' totaltasa = msktottas.FormattedText
     
   sNit = "NIT.860.052.155-6"
   'sEmpVeh = Txtcodemp.Text & "-" & Lblnomemp.Caption & "   TERMINAL " & G_Tipo_Terminal
   
   sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & Mid(G_Tipo_Terminal, 1, 1)
   
   If G_Imprime_Tasa_Uso = "E" Then
      sNombreEmpresa = sEmpVeh
   Else
      If L_Operador <> "" Then
         sNombreEmpresa = Traer_Descripcion(L_Operador, "OPERADORES DE SERVICIO")
      Else
         sNombreEmpresa = sEmpVeh
      End If
   End If
   
   If Len(Txtnumcon.text) = 1 Then
      LNumConBarra = "00" & Txtnumcon.text
   ElseIf Len(Txtnumcon.text) = 2 Then
      LNumConBarra = "0" & Txtnumcon.text
   Else
      LNumConBarra = Txtnumcon.text
   End If
   
   'Buscamos el número consecutivo del TAG
   QRY = "SELECT * FROM optag " _
            & "WHERE tafecpag IS NOT NULL " _
            & "AND taplaca = '" & Txtnumpla.text & "' " _
            & "AND tanumtasa = " & Txtnumcon.text & " "
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Tbl1.EOF = False Then
        ConsecutivoTag = Tbl1!tanumero
        CodigoTag = Tbl1!tacodtag
    Else
        ConsecutivoTag = 0
        CodigoTag = 0
    End If
    
    ConceptoTag = "PAGO REPOSICION TAG"
   
   If d = "S" Then
        valtag = Tbl1!tavalor
        ValIVA = Tbl1!taiva
        ValSinIVa = Tbl1!tasubtot
        sValorTag = CDbl(Tbl1!tavalor)
   Else
       'Calculamos el IVA
       QRY = "SELECT psval FROM geparsis " _
                & "WHERE pscod = 'IVA' "
                
       Set Tbl1 = Dbs.Execute(QRY)
       
       If Tbl1.EOF = False Then
            
    '        valtag = MskValTag.text
    '        IVA = "0." & Tbl1!psval
    '        ValIVA = CDbl(MskValTag.text) * CDbl(IVA)
    '        ValSinIVA = CDbl(MskValTag.text) - CDbl(ValIVA)
            
'            valtag = MskValTag.text
'            IVA = Tbl1!psval
'            ValIVA = valtag / (1 + (IVA / 100))
'            ValIVA = ValIVA * (IVA / 100)
'
'            decimales = CCur(ValIVA) - Int(ValIVA)
'            If decimales > 0.5 Then
'              ValIVA = Int(ValIVA) + 1
'            Else
'              ValIVA = Int(ValIVA)
'            End If
'
'            ValSinIva = CDbl(valtag) - CDbl(ValIVA)

            valtag = MskValTag.text
            IVA = Tbl1!psval
            Base = valtag / (1 + (IVA / 100))
            
            decimales = CCur(Base) - Int(Base)
            If decimales > 0.5 Then
                Base = Int(Base) + 1
            Else
                Base = Int(Base)
            End If
            
            ValIVA = CDbl(valtag) - CDbl(Base)
            ValSinIVa = CDbl(Base)
       End If
       
       sValorTag = CDbl(MskValTag.text)
   End If
   
   Tbl1.Close
   
   sEmpVeh = sNombreEmpresa
   sNumero = Txtnumcon.text
   sValorTasa = Mskvalor.FormattedText
   sNumInterno = TxtNumInt.text
   sFecha = Format(Mskfecha.text, "MMM/dd/yyyy")
   sPlaca = Txtnumpla.text
   
   If f = "Z" Then
    Open "LPT1" For Output As #1
   
    Print #1, "CT~~CD,~CC^~CT~"
    Print #1, "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2~SD25^JUS^LRN^CI0^XZ"
    Print #1, "^XA"
    Print #1, "^MMT"
    Print #1, "^PW831"
    Print #1, "^LL0743"
    Print #1, "^LS0"
    Print #1, "^FO0,608^GFA,01536,01536,00012,:Z64:" _
         & "eJztUjtOw0AQHccmshbkmAJBgYRlmihcIkfhCtzAovIxXFobiSuQI3AERBVRRCmoDLLxzM7sjiMkesRIWT2/vPntPoD/+DXSuvL4YtgKnLVgBScAJ4LbUd0yXo+/nOVYxHg5QMTV6QwdAO7pbOhcKv5aKY1KmFN51YUjqnx5l5arqQrGBX8D/y/5urHmUz84wJlvC5Ap/Rx5QwusZrjt4oDj9DGmZ6845v6x8vwQIb9APu7GBdIJn+1ITwPcOB7HMJvOWrvpRj5fDBQ9bKHO9i4m/BoKws+D1n+NfMr6neIPqD/9xLDNhN9C4fTvu0mdiusPnT3ipf4T8pc8D+1191aW5a2lvWjf2PF8P5jr9mXe3U8f7jPu0QdyPyY8+0ei3iXwznbyjrV6x+N3Fz+gNxrGDQTPrBWuwPsqYn9gkA+T0FYGMl44df7SFxb/F74twJXKIuez/SOZXqQyba62MzVY2Q4ezl8EwirQk6jgZ/4PxjfrALGs:D6AD"
    Print #1, "^BY3,2,47^FT377,574^B3B,N,,Y,N"
    Print #1, "^FDA1B2C3^FS"
    Print #1, "^FT348,109^A0B,20,21^FH\^FD" & ValIVA & "^FS"
    Print #1, "^FT319,109^A0B,20,21^FH\^FD" & ValSinIVa & "^FS"
    Print #1, "^FT251,138^A0B,20,55^FH\^FD" & G_Turno & "^FS"
    Print #1, "^FT60,519^A0B,23,21^FH\^FD" & G_NomEmpresa & "^FS"
    Print #1, "^FT52,217^A0B,14,12^FH\^FDNumero^FS"
    Print #1, "^FT190,511^A0B,20,16^FH\^FDHORA^FS"
    Print #1, "^FT248,511^A0B,20,19^FH\^FD" & sEmpVeh & "^FS"
    Print #1, "^FT251,672^A0B,20,19^FH\^FDEMPRESA:^FS"
    Print #1, "^FT224,217^A0B,20,19^FH\^FDCAJA^FS"
    Print #1, "^FT283,672^A0B,20,19^FH\^FDCOD. TAG:^FS"
    Print #1, "^FT190,331^A0B,20,16^FH\^FDPLACA^FS"
    Print #1, "^FT214,426^A0B,20,21^FH\^FD" & sNumeroInterno & "^FS"
    Print #1, "^FT190,427^A0B,20,16^FH\^FDNo INT.^FS"
    Print #1, "^FT214,511^A0B,20,21^FH\^FD" & hora & "^FS"
    Print #1, "^FT383,107^A0B,25,24^FH\^FD" & sValorTag & "^FS"
    Print #1, "^FT345,229^A0B,20,19^FH\^FDIVA^FS"
    Print #1, "^FT318,229^A0B,20,19^FH\^FDValor:^FS"
    Print #1, "^FT383,227^A0B,25,24^FH\^FDValor Pagar^FS"
    Print #1, "^FT251,217^A0B,20,19^FH\^FDTURNO^FS"
    Print #1, "^FT190,138^A0B,20,21^FH\^FD" & Login & "^FS"
    Print #1, "^FT224,138^A0B,20,21^FH\^FD" & Pc & "^FS"
    Print #1, "^FT190,215^A0B,20,19^FH\^FDCAJERO^FS"
    Print #1, "^FT216,331^A0B,20,21^FH\^FD" & sPlaca & "^FS"
    Print #1, "^FT190,671^A0B,20,16^FH\^FDFECHA PAGO^FS"
    Print #1, "^FT217,672^A0B,20,21^FH\^FD" & sFecha & "^FS"
    Print #1, "^FT80,460^A0B,17,16^FH\^FD" & sNit & "^FS"
    Print #1, "^FT64,147^A0B,34,33^FH\^FD" & ConsecutivoTag & "^FS"
    Print #1, "^FT33,494^A0B,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
    Print #1, "^FT137,542^A0B,31,31^FH\^FDPAGO REPOSICION TAG^FS"
    Print #1, "^FT283,511^A0B,20,19^FH\^FD" & CodigoTag & "^FS"
    Print #1, "^PQ1,0,1,Y^XZ"
    Close #1

   Else
    
    reporte = RpPath + "\rectagzebra.rpt"
    Rpt.ReportFileName = reporte
    Rpt.Connect = G_Conexion_reportes
    
    Rpt.Formulas(0) = "nombreempresa = '" & G_NomEmpresa & "'"
    Rpt.Formulas(1) = "numcontag = '" & ConsecutivoTag & "'"
    Rpt.Formulas(2) = "hora = '" & hora & "'"
    Rpt.Formulas(3) = "numint = '" & sNumeroInterno & "'"
    Rpt.Formulas(4) = "numplaca = '" & sPlaca & "'"
    Rpt.Formulas(5) = "usuario = '" & Login & "'"
    Rpt.Formulas(6) = "pc = '" & Pc & "'"
    Rpt.Formulas(7) = "conceptotag = '" & ConceptoTag & "'"
    Rpt.Formulas(8) = "empresa = '" & sEmpVeh & "'"
    Rpt.Formulas(9) = "totalpagar = '" & Format(sValorTag, "$##,##0") & "'"
    Rpt.Formulas(10) = "iva = '" & Format(ValIVA, "$##,##0") & "'"
    Rpt.Formulas(11) = "valtag = '" & Format(ValSinIVa, "$##,##0") & "'"
    
    Rpt.Formulas(12) = "codtag = '" & CodigoTag & "'"
    Rpt.Formulas(13) = "turno = '" & G_Turno & "'"
    Rpt.Formulas(14) = "fecha = '" & sFecha & "'"
    
    If G_ImpEpsonZebra = "P" Then
       Rpt.PrinterName = G_PrinterName1
       Rpt.PrinterDriver = G_PrinterDriver1
       Rpt.PrinterPort = G_PrinterPort1
       'Rpt.ProgressDialog = False
    End If
    '
    Rpt.Destination = 1
    Rpt.Action = 1
    
   End If
   
   
End Sub

Private Sub Imprimir_Crystal()
   hora = Mid(Mskhora.text, 12, 15)
   hora = Format(hora, "hh:mm")
   
   totaltasa = Msktottas.FormattedText
   
   If CDbl(Mskrecibo.text) = 0 Then
      presenprueba = "No"
   End If
     
   presenprueba = presenprueba & "-" & Lbloritra.Caption
   sNit = "NIT.860.052.155-6"
   'sEmpVeh = Txtcodemp.Text & "-" & Lblnomemp.Caption & "   TERMINAL " & G_Tipo_Terminal
   
   sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & Mid(G_Tipo_Terminal, 1, 1)
   
   If G_Imprime_Tasa_Uso = "E" Then
      sNombreEmpresa = sEmpVeh
   Else
      If L_Operador <> "" Then
         sNombreEmpresa = Traer_Descripcion(L_Operador, "OPERADORES DE SERVICIO")
      Else
         sNombreEmpresa = sEmpVeh
      End If
   End If
   If Len(Txtnumcon.text) = 1 Then
      LNumConBarra = "00" & Txtnumcon.text
   ElseIf Len(Txtnumcon.text) = 2 Then
      LNumConBarra = "0" & Txtnumcon.text
   Else
      LNumConBarra = Txtnumcon.text
   End If

      
   sEmpVeh = sNombreEmpresa
   'sAlcoholimetria = "PGS_SVIAL.ALCH"
   sAlcoholimetria = "PRGSALYSEGVIAL"
   sNumero = Txtnumcon.text
   sruta = Txtcodrut.text & "-" & Lblnomrut.Caption
   sNivelServicio = Lbltipveh.Caption
   sValorParqueadero = MskParqueadero.FormattedText
   sValorTasa = Mskvalor.FormattedText
   sValorAlcoholimetria = Mskrecibo.FormattedText
   sNumInterno = TxtNumInt.text
   sFecha = Format(Mskfecha.text, "MMM/dd/yyyy")
   sPlaca = Txtnumpla.text
   
   s2Conductores = ""
   If Aviso_2_Conductores = False Then
     s2Conductores = "DESTINO NECESITA DOS CONDUCTORES"
   End If
   If G_Cantidad_Parqueaderos > 0 Then
      sPermanencia1 = "Permanencias"
      sPermanencia2 = "Pendientes " & G_Cantidad_Parqueaderos
      sPermanencia3 = "Aviso " & G_Avisos_Parqueadero & " de " & L_AvisoParqueadero
   Else
      sPermanencia1 = ""
      sPermanencia2 = ""
      sPermanencia3 = ""
   End If
   
   sNumeroBarras = "*" & LNumConBarra & "*"
   reporte = RpPath + "\conduce.rpt"
   Rpt.ReportFileName = reporte
   Rpt.Connect = G_Conexion_reportes
   Rpt.Formulas(0) = "empveh = '" & sEmpVeh & "'"
   Rpt.Formulas(1) = "fechaexp = '" & sFecha & "'"
   Rpt.Formulas(2) = "hora = '" & hora & "'"
   Rpt.Formulas(3) = "nivelservicio = '" & sNivelServicio & "'"
   Rpt.Formulas(4) = "nombreempresa = '" & G_NomEmpresa & "'"
   Rpt.Formulas(5) = "numint = '" & sNumInterno & "'"
   Rpt.Formulas(6) = "numplaca = '" & sPlaca & "'"
   Rpt.Formulas(7) = "numtasauso = '" & sNumero & "'"
   Rpt.Formulas(8) = "pc = '" & Pc & "'"
   Rpt.Formulas(9) = "permanencias1 = '" & sPermanencia1 & "'"
   Rpt.Formulas(10) = "permanencias2 = '" & sPermanencia2 & "'"
   Rpt.Formulas(11) = "permanencias3 = '" & sPermanencia3 & "'"
   Rpt.Formulas(12) = "pruebaalcohol = '" & presenprueba & "'"
   Rpt.Formulas(13) = "ruta = '" & sruta & "'"
   Rpt.Formulas(14) = "turno = '" & G_Turno & "'"
   Rpt.Formulas(15) = "usuario = '" & Login & "'"
   Rpt.Formulas(16) = "valoralcohol = '" & sValorAlcoholimetria & "'"
   'rpt.Formulas(17) = "valorparqueo = '" & sValorParqueadero & "'"
   Rpt.Formulas(18) = "valortasa = '" & sValorTasa & "'"
   Rpt.Formulas(19) = "valortotal = '" & totaltasa & "'"
   Rpt.Formulas(20) = "codigobarras = '" & sNumeroBarras & "'"
   Rpt.Formulas(21) = "2conductores = '" & s2Conductores & "'"
   Rpt.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
   Rpt.Destination = 1
   Rpt.Action = 1
End Sub

Private Sub Imprimir_Exe_Net()
   
      If L_Numero_Terminal_Norte <> 0 Then
         sCantidadImpresiones = 2
         sNombreRutaTransito = Traer_Nombre_Ruta(L_Ruta_Transito_Norte)
      Else
         sCantidadImpresiones = 1
      End If
      
      If Txtcopias.text = "" Then
         sCopias = 1
      ElseIf Txtcopias.text = "0" Then
         sCopias = 1
      Else
         sCopias = Val(Txtcopias.text)
      End If
      
      'If sCopias = 1 Then
         For k = 1 To sCantidadImpresiones
            If k = 1 Then
               'imprime la tasa de uso normal
               'hora = Mid(Mskhora.Text, 12, 15)
               hora = Format(Mskhora.text, "hh:mm")
               
               If CDbl(Mskrecibo.text) = 0 Then
                  presenprueba = "No"
               End If
                 
               presenprueba = presenprueba & "-" & Lbloritra.Caption
               
               sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & Mid(G_Tipo_Terminal, 1, 1)
               
               If G_Imprime_Tasa_Uso = "E" Then
                  sNombreEmpresa = sEmpVeh
               Else
                  If L_Operador <> "" Then
                     sNombreEmpresa = Traer_Descripcion(L_Operador, "OPERADORES DE SERVICIO")
                  Else
                     sNombreEmpresa = sEmpVeh
                  End If
               End If
               If Len(Txtnumcon.text) = 1 Then
                  LCodBarras = "00" & Txtnumcon.text
               ElseIf Len(Txtnumcon.text) = 2 Then
                  LCodBarras = "0" & Txtnumcon.text
               Else
                  LCodBarras = Txtnumcon.text
               End If
               sNumeroBarras = "*" & LCodBarras & "*"
               s2Conductores = txtCedula.text
               If Aviso_2_Conductores = False Then
                 s2Conductores = s2Conductores & "-" & TxtCedula1.text  ' "DESTINO NECESITA DOS CONDUCTORES"
               End If
               If G_Cantidad_Parqueaderos > 0 Then
                  sPermanencia1 = "Permanencias"
                  sPermanencia2 = "Pendientes " & G_Cantidad_Parqueaderos
                  sPermanencia3 = "Aviso " & G_Avisos_Parqueadero & " de " & L_AvisoParqueadero
               Else
                  sPermanencia1 = "-"
                  sPermanencia2 = "-"
                  sPermanencia3 = "-"
               End If
               
               cero = "0"
               cedula1 = txtCedula.text
               For i = 1 To 10 - Len(txtCedula.text)
                  cedula1 = cero & cedula1
               Next i
               
               cedula2 = TxtCedula1.text
               For i = 1 To 10 - Len(TxtCedula1.text)
                  cedula2 = cero & cedula2
               Next i
               
               numcon = Txtnumcon.text
               For i = 1 To 10 - Len(Txtnumcon.text)
                  numcon = cero & numcon
               Next i
               
               sCodigoDestino = Txtcodrut.text
               For i = 1 To 4 - Len(Txtcodrut.text)
                  sCodigoDestino = cero & sCodigoDestino
               Next i
               
               'If OptTipo(0).Value = True Then
                  sForma_Pago = L_BanderaVentaTasa
               'Else
               '   sForma_Pago = "CUPO"
               'End If
               
               If Txtoritra.text = "O" Then
                    If influencia = "No" Then
                        sPagaPrueba = "S"
                    Else
                        If Mskrecibo.text = 0 Or Mskrecibo.text = "0" Then
                            sPagaPrueba = "N"
                        Else
                            sPagaPrueba = "S"
                        End If
                    End If
               Else
                    sPagaPrueba = "N"
               End If
               
               LNumConBarra = Txtnumpla.text & cedula1 & cedula2 & numcon & sCodigoDestino & sPagaPrueba
               sNumeroTasa = Txtnumcon.text
               sFechaExpedicion = Format(Mskfecha.text, "MMM/dd/yyyy")
               sDestino = Replace(Trim(Txtcodrut.text & "-" & Lblnomrut.Caption), " ", "_")
               sValorTasa = Mskvalor.FormattedText
               sValorAlcohol = Mskrecibo.FormattedText
               sValorTotal = Msktottas.FormattedText
               sCaja = Pc
            Else
               'imprime la tasa de uso automatica norte
               hora = L_HoraSalidaNorte
               If influencia = "No" Then
                  presenprueba = "Si"
               End If
               If CDbl(L_Valor_Alcoholimetria_Norte) = 0 Then
                  presenprueba = "No"
               End If
                 
               presenprueba = presenprueba & "-Transito"
               
               sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & L_Terminal_Tasa_Automatica
               
               If Len(L_Numero_Terminal_Norte) = 1 Then
                  LCodBarras = "00" & L_Numero_Terminal_Norte
               ElseIf Len(Txtnumcon.text) = 2 Then
                  LCodBarras = "0" & L_Numero_Terminal_Norte
               Else
                  LCodBarras = L_Numero_Terminal_Norte
               End If
               sNumeroBarras = "*" & LCodBarras & "*"
               s2Conductores = L_Cedula1Norte
               If L_Cedula2Norte <> "" Then
                 s2Conductores = s2Conductores & "-" & L_Cedula2Norte  ' "DESTINO NECESITA DOS CONDUCTORES"
               End If
               sPermanencia1 = "-"
               sPermanencia2 = "-"
               sPermanencia3 = "-"
            
               cero = "0"
               numcon = L_Numero_Terminal_Norte
               For i = 1 To 10 - Len(L_Numero_Terminal_Norte)
                  numcon = cero & numcon
               Next i
               
      '         cedula1 = L_Cedula1Norte
      '         For i = 1 To 10 - Len(L_Cedula1Norte)
      '            cedula1 = cero & cedula1
      '         Next i
      '
      '         cedula2 = L_Cedula2Norte
      '         For i = 1 To 10 - Len(L_Cedula2Norte)
      '            cedula2 = cero & cedula2
      '         Next i
      '
      '         sCodigoDestino = L_Ruta_Transito_Norte
      '         For i = 1 To 4 - Len(L_Ruta_Transito_Norte)
      '            sCodigoDestino = cero & sCodigoDestino
      '         Next i
                              
               sForma_Pago = L_BanderaVentaTasa
               
               LNumConBarra = "NORTE" 'numcon
               sNumeroTasa = L_Numero_Terminal_Norte
               sFechaExpedicion = Format(L_FechaSalidaNorte, "MMM/dd/yyyy")
               sDestino = Replace(Trim(L_Ruta_Transito_Norte & "-" & sNombreRutaTransito), " ", "_")
               sValorTasa = MskValorTasaNorte.FormattedText
               sValorAlcohol = L_Valor_Alcoholimetria_Norte
               sValorTotal = CDbl(MskValorTasaNorte.text) + CDbl(L_Valor_Alcoholimetria_Norte)
               sCaja = L_Caja_Norte
               Sleep 1000
            End If
            
            sNombreTerminal = G_NomEmpresa
            
            
            sHoraExpedicion = hora
            sNumeroInterno = TxtNumInt.text
            sPlaca = Txtnumpla.text
            sEmpresaVehiculo = Replace(sNombreEmpresa, " ", "_") & "_____" & sForma_Pago
            sNivelServicio = Replace(Lbltipveh.Caption, " ", "_")
            
            L_Bandera_Copia = False
            If sFuncioImpCon = "" Then
               sUsuario = Login         'Usuario que esta registro la tasa de uso actualmente
            Else
               sUsuario = sFuncioImpCon   'Usuario que registro la tasa de uso previamente
               s2Conductores = s2Conductores & "-COPIA-TU"
               L_Bandera_Copia = True
            End If
            
            If sTurnoImpCon = 0 Then
               sturno = G_Turno        'turno que esta logeado actualmente
            Else
               sturno = sTurnoImpCon   'turno que se regsitro la tasa de uso
               L_Bandera_Copia = True
            End If
            sPruebaAlcohol = presenprueba
            s2Conductores = Replace(s2Conductores, " ", "_")
            sPermanencias1 = Replace(sPermanencia1, " ", "_")
            sPermanencias2 = Replace(sPermanencia2, " ", "_")
            sPermanencias3 = Replace(sPermanencia3, " ", "_")
            sCodigoBarras = LNumConBarra
            
            'WECC CAU2 109 oct 22 2018 Cuando es influencia se imprime en una pos en el exe de NET el reporte se llama TASAUSOPOS
            'Y PARA LAS RUTAS QUE NOS SON INFLUENCIA SE LLAMA TASAUSO
            If influencia <> "Si" Then
               sTipoImpresion = "Z"
            Else
                If sCaja = "CAJA8" Then
                    sTipoImpresion = "P"
                Else
                    sTipoImpresion = "Z"
                End If
            End If
            
            sTitulo = ""
            sTitNumAuto = "No_AUTORIZACION: "
            sNumAuto = ""
            sTitFecAuto = "FECHA_AUTORIZACION: "
            sFecAuto = ""
            
            QRY = "SELECT psval FROM geparsis WHERE pscod = 'IMPFETU' "
            
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                If Tbl1!psval = "N" Then
                
                    QRY1 = "SELECT canumaut, cafecaut, caprefijo " _
                            & "FROM cocaja WHERE canomcaj = '" & sCaja & "' "
                            
                    Set Tbl2 = Dbs.Execute(QRY1)
                    
                    If Not Tbl2.EOF Then
                        sTitulo = "D.E.S/POS_" & Tbl2!caprefijo & " "
                        'sTitNumAuto = "No_AUTORIZACION: "
                        sNumAuto = Tbl2!canumaut
                        'sTitFecAuto = "FECHA_AUTORIZACION: "
                        sFecAuto = Mid(Tbl2!cafecaut, 1, 10)
                    End If
                
                Else
                    sTitulo = "Tasa_de_Uso_No.:"
                    sNumAuto = "N/A"
                    sFecAuto = "N/A"
                End If
            End If
            
         '   For i = 1 To G_NumImpTasa
            scadena = sNumeroTasa & " " & sFechaExpedicion & " " & sHoraExpedicion & " " & sNumeroInterno & " " _
                 & sPlaca & " " & sEmpresaVehiculo & " " & sDestino & " " & sNivelServicio & " " & sUsuario & " " & sCaja & " " _
                 & sturno & " " & sValorTasa & " " & sValorAlcohol & " " & sValorTotal & " " & sPruebaAlcohol & " " & s2Conductores & " " _
                 & sPermanencias1 & " " & sPermanencias2 & " " & sPermanencias3 & " " & sCodigoBarras & " " & sNumeroBarras & " " & sTipoImpresion & " " _
                 & sTitulo & " " & sTitNumAuto & " " & sNumAuto & " " & sTitFecAuto & " " & sFecAuto
                 
            s = Shell(G_Exe_Net_Impresion & " " & scadena, vbHide)
         Next k
'      Else
'         'cuando son mas de una copia, esto se da por el cierre de la via al llano y porque se esta de viaje en _
'         barranquilla y el ingeniero alvaro dijo que urgente
'         'wecc jun/27/2019
'
'            'imprime la tasa de uso normal
'            'hora = Mid(Mskhora.Text, 12, 15)
'            hora = Format(Mskhora.Text, "hh:mm")
'
'            If CDbl(Mskrecibo.Text) = 0 Then
'               presenprueba = "No"
'            End If
'
'            presenprueba = presenprueba & "-" & Lbloritra.Caption
'
'            sEmpVeh = L_Codigo_Empresa & "-" & L_Nombre_Empresa & "   T " & Mid(G_Tipo_Terminal, 1, 1)
'
'            If G_Imprime_Tasa_Uso = "E" Then
'               sNombreEmpresa = sEmpVeh
'            Else
'               If L_Operador <> "" Then
'                  sNombreEmpresa = Traer_Descripcion(L_Operador, "OPERADORES DE SERVICIO")
'               Else
'                  sNombreEmpresa = sEmpVeh
'               End If
'            End If
'            If Len(Txtnumcon.Text) = 1 Then
'               LCodBarras = "00" & Txtnumcon.Text
'            ElseIf Len(Txtnumcon.Text) = 2 Then
'               LCodBarras = "0" & Txtnumcon.Text
'            Else
'               LCodBarras = Txtnumcon.Text
'            End If
'            sNumeroBarras = "*" & LCodBarras & "*"
'            s2Conductores = txtcedula.Text
'            If Aviso_2_Conductores = False Then
'              s2Conductores = s2Conductores & "-" & Txtcedula1.Text  ' "DESTINO NECESITA DOS CONDUCTORES"
'            End If
'            If G_Cantidad_Parqueaderos > 0 Then
'               sPermanencia1 = "Permanencias"
'               sPermanencia2 = "Pendientes " & G_Cantidad_Parqueaderos
'               sPermanencia3 = "Aviso " & G_Avisos_Parqueadero & " de " & L_AvisoParqueadero
'            Else
'               sPermanencia1 = "-"
'               sPermanencia2 = "-"
'               sPermanencia3 = "-"
'            End If
'
'            cero = "0"
'            cedula1 = txtcedula.Text
'            For I = 1 To 10 - Len(txtcedula.Text)
'               cedula1 = cero & cedula1
'            Next I
'
'            cedula2 = Txtcedula1.Text
'            For I = 1 To 10 - Len(Txtcedula1.Text)
'               cedula2 = cero & cedula2
'            Next I
'
'            numcon = Txtnumcon.Text
'            For I = 1 To 10 - Len(Txtnumcon.Text)
'               numcon = cero & numcon
'            Next I
'
'            sCodigoDestino = Txtcodrut.Text
'            For I = 1 To 4 - Len(Txtcodrut.Text)
'               sCodigoDestino = cero & sCodigoDestino
'            Next I
'
'            If OptTipo(0).Value = True Then
'               sForma_Pago = "EFECTIVO"
'            Else
'               sForma_Pago = "CUPO"
'            End If
'
'            LNumConBarra = Txtnumpla.Text & cedula1 & cedula2 & numcon & sCodigoDestino
'            sNumeroTasa = Txtnumcon.Text
'            sFechaExpedicion = Format(Mskfecha.Text, "MMM/dd/yyyy")
'            sDestino = Replace(Trim(Txtcodrut.Text & "-" & Lblnomrut.Caption), " ", "_")
'            sValorTasa = Mskvalor.FormattedText
'            sValorAlcohol = Mskrecibo.FormattedText
'            sValorTotal = msktottas.FormattedText
'            sCaja = Pc
'
'
'            sNombreTerminal = G_NomEmpresa
'
'
'            sHoraExpedicion = hora
'            sNumeroInterno = Txtnumint.Text
'            sPlaca = Txtnumpla.Text
'            sEmpresaVehiculo = Replace(sNombreEmpresa, " ", "_") & "_____" & sForma_Pago
'            sNivelServicio = Replace(Lbltipveh.Caption, " ", "_")
'
'            L_Bandera_Copia = False
'            If sFuncioImpCon = "" Then
'               sUsuario = Login         'Usuario que esta registro la tasa de uso actualmente
'            Else
'               sUsuario = sFuncioImpCon   'Usuario que registro la tasa de uso previamente
'               s2Conductores = s2Conductores & "-COPIA-TU"
'               L_Bandera_Copia = True
'            End If
'
'            If sTurnoImpCon = 0 Then
'               sturno = G_Turno        'turno que esta logeado actualmente
'            Else
'               sturno = sTurnoImpCon   'turno que se regsitro la tasa de uso
'               L_Bandera_Copia = True
'            End If
'            sPruebaAlcohol = presenprueba
'            s2Conductores = Replace(s2Conductores, " ", "_")
'            sPermanencias1 = Replace(sPermanencia1, " ", "_")
'            sPermanencias2 = Replace(sPermanencia2, " ", "_")
'            sPermanencias3 = Replace(sPermanencia3, " ", "_")
'            sCodigoBarras = LNumConBarra
'
'            'WECC CAU2 109 oct 22 2018 Cuando es influencia se imprime en una pos en el exe de NET el reporte se llama TASAUSOPOS
'            'Y PARA LAS RUTAS QUE NOS SON INFLUENCIA SE LLAMA TASAUSO
'            If influencia <> "Si" Then
'               sTipoImpresion = "Z"
'            Else
'               sTipoImpresion = "P"
'            End If
'
'         '   For i = 1 To G_NumImpTasa
'            scadena = sNumeroTasa & " " & sFechaExpedicion & " " & sHoraExpedicion & " " & sNumeroInterno & " " _
'                 & sPlaca & " " & sEmpresaVehiculo & " " & sDestino & " " & sNivelServicio & " " & sUsuario & " " & sCaja & " " _
'                 & sturno & " " & sValorTasa & " " & sValorAlcohol & " " & sValorTotal & " " & sPruebaAlcohol & " " & s2Conductores & " " _
'                 & sPermanencias1 & " " & sPermanencias2 & " " & sPermanencias3 & " " & sCodigoBarras & " " & sNumeroBarras & " " & sTipoImpresion
'         For k = 1 To sCopias
'            If k > 1 Then
'               MsgBox "Imprimir la siguiente copia", vbInformation, "Ok"
'            End If
'            S = Shell(G_Exe_Net_Impresion & " " & scadena, vbHide)
'
'         Next k
'
'      End If
        ' sEliminar = G_DirImpBol & "\Eliminar.BAT"
         
         If Mid(G_DirImpBol, Len(G_DirImpBol), 1) = "\" Then
            sEliminar = G_DirImpBol & "Eliminar.BAT"
         Else
            sEliminar = G_DirImpBol & "\Eliminar.BAT"
         End If
      
         s = Shell(sEliminar, vbHide)
   '      Sleep 2000
   '   Next i
   
   
End Sub


Private Sub Btnlimpiar_Click()
   L_EstViajeCedula1 = ""
   L_EstViajeCedula2 = ""
   L_EstViajePlaca = ""
   Erase multasGuardadas 'para limpiar el arreglo con multas
   Erase parqueGuardado 'para limpiar el arreglo con parqueo
   Txtnumcon.Locked = True
   Btnimprimir.Enabled = False
   Btnbuscar.Enabled = True
   TxtNumInt.Enabled = True
   Txttipveh.Enabled = True
   Txtcodemp.Enabled = True
   Txtcodrut.Enabled = True
   Txtcodrut.Locked = False
   Txtoritra.Enabled = True
   TxtCedula1.Enabled = True
   Txtconact1.Enabled = True
   Txtapellido1.Enabled = True
   Mskfecha.Enabled = True
   sFuncioImpCon = ""
   sTurnoImpCon = 0
   LTipoConvenio = "N"
   influencia = ""
   s_Num_Multa_Pagada = 0
   L_Numero_Despacho = 0
   L_Codigo_Empresa = ""
   L_Nombre_Empresa = ""
   L_Bandera_Ruta_Lector = ""
   fqry01 = "SELECT * " _
            & "FROM cocaja " _
           & "WHERE canomcaj = '" & Pc & "' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      If Ftb01!caestado = "Inactiva" Then
         MsgBox "La Caja Se Encuentra Inactiva", 16, "Error "
         Unload Me
         Exit Sub
      End If
      Txtnumcon.text = CDbl(Ftb01!caconcaj) + 1
      
      L_Terminal = Ftb01!catipo
   Else
      Btngrabar.Enabled = False
      Btnlimpiar.Enabled = False
   End If
   Ftb01.Close
   
   'wecc may/05/2016 implementacion de cupos
   '**********************************************************
   FrmTipo.Enabled = False
   Lblcupo.Caption = "LINEA"
   OptTipo(0).Value = True
   OptTipo(1).Enabled = True
   mskdevolver.text = ""
   Mskdinero.text = ""
   Txtcodrut.Locked = False
   Txtoritra.Locked = False
   txtCedula.Locked = False
   TxtCedula1.Locked = False
   TxtNumAlistamiento.Locked = False
   chkDatosPersonales.Value = 1
   '**********************************************************
   
   'TxtNumAlistamiento.Enabled = False
   'wecc 01/29/2014  CAU7511   CASO DE USO 002
   '***********************************************************
   OptTipoPlaca(0).Value = True
   Txtnumpla.MaxLength = 6
   '***********************************************************
   L_ValorTransitoSur = 0
   L_Numero_Terminal_Norte = 0
   L_Valor_Alcoholimetria_Norte = 0
   
   G_Imprime_Tasa_Uso = "E"
   ExisteCedula1 = 0
   ExisteCedula = 0
   TxtNumAlistamiento.Enabled = True
   TxtNumAlistamiento.text = ""
   TxtNumConvenio.text = ""
   L_CodigoEmpresaConvenioEmpresarial = ""
   L_CodigoEmpresaConvenio = ""
   Txtcodemp.text = ""
   Txtcodrut.text = ""
   Txtnumpla.text = ""
   Txtnumpla.Enabled = True
   TxtNumInt.text = ""
   Txttipveh.text = ""
   txtCedula.Enabled = True
   Txtconact.Enabled = True
   txtapellido.Enabled = True
   txtCedula.text = ""
   txtapellido.text = ""
   Txtconact.text = ""
   TxtCedula1.text = ""
   Txtapellido1.text = ""
   Txtconact1.text = ""
   Txtoritra.text = ""
   Txtorigen.text = ""
   Lblnomemp.Caption = ""
   Lblnomrut.Caption = ""
   Lbltipveh.Caption = ""
   Lbloritra.Caption = ""
   Lblorigen.Caption = ""
   Txtcopias.text = "1"
   TxtPeajeUltimo.text = "0"
   TxtPeaje.text = ""
   Mskvalor.text = "0"
   MskValorTasaNorte.text = "0"
   Mskrecibo.text = "0"
   Mskcambio.text = "0"
   ChkTag.Value = False
   MskValTag.text = "0"
   sNumConvenio = 0
   'Txtnumpla.SetFocus
   Btngrabar.Enabled = True
   Txtorigen.Enabled = False
   Lblorigen.Enabled = False
   Label1(12).Enabled = False
   MskParqueadero.text = "0"
   MskMultas.text = "0"
   Mskhora.Enabled = True
   bandera_huella = "N"
   L_Valida_CedHuella = ""
   bandera_huella1 = "N"
   L_Valida_CedHuella1 = ""
   L_Bandera_Tasa_Condicionada = False
   L_Bandera_Venta_Condicionada = False
   L_Bandera_NumeroTag = 0
   L_Bandera_ReimpresionTag = False
   
   
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      hora = Mid(Ftb01!psval, 12, 5)
      Mskhorsis.text = hora
      
      Fecha = Mid(Ftb01!psval, 1, 10)
       
      fqry01 = "SELECT psval " _
               & "FROM geparsis " _
              & "WHERE pscod = 'ADERELOJ' "
   
      Set Tbl1 = Dbs.Execute(fqry01)
      If Not Tbl1.EOF Then
         hora1 = DateAdd("n", Tbl1(0), hora)
         horasis = Format(hora1, "hh:mm")
      End If
      Tbl1.Close
      Mskhora.Mask = ""
      Mskhora.text = ""
      Mskhora.text = CDate(hora1)
      Mskfecha.text = Fecha
      main.StatusBar1.Panels(3) = "Fecha Actual : " & Fecha
   End If
   Ftb01.Close
   
   If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
      Txtoritra.text = "T"
   Else
      Txtoritra.text = "O"
   End If
    
   ValorParqueo = "0"
   Mskingreso.text = ""
   chkcredito.Enabled = True
   chkcredito.Value = 0
   Btngrabar.Default = False
   direruta = ""
   ConNombre = ""
   ConHora = ""
   LblPasajeros.Visible = False
   Txtpasajeros.Enabled = False
   Txtpasajeros.Visible = False
   Txtpasajeros.text = ""
   If main.Mnucanmul.Enabled = True Then
      btnmultas.Enabled = True
   Else
      btnmultas.Enabled = False
   End If
   
   'L_Bandera_Rodamiento = False
   
   On Error Resume Next
   Txtnumpla.SetFocus
End Sub

Private Sub btnmultas_Click()
    On Error Resume Next
    If main.Mnucanmul.Enabled = True Then
       Canmultas.txtPlaca.text = Txtnumpla.text
       Unload Me
       Canmultas.Show 1
    End If

End Sub

Private Sub BtnParqueadero_Click()
    If Txtnumpla.text <> "" Then
        L_Placa_Permanencia = Txtnumpla.text
    End If
   capparqu.Txtnumpla.text = L_Placa_Permanencia
   Unload Me
   capparqu.Show 1
End Sub

Private Sub Btnsalir_Click()
 CerrarFormulario
 Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CerrarFormulario
End Sub

Private Sub CerrarFormulario()
    On Error Resume Next
    
    Close 'Para cerrar puertos abiertos y liberar recursos ya que el ejecutable esta colapsando
    Set rptMultas = Nothing 'liberar recursos
    Set rptPar = Nothing
  
    
    
    If G_ManLector = "S" Then
        If Not dbs4 Is Nothing Then
            dbs4.Close
            Set dbs4 = Nothing
        End If
    End If

    RegistoHuella = False

    On Error GoTo 0
End Sub

Private Sub chkcredito_Click()
   On Error Resume Next
   Btngrabar.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub ChkTag_Click()

    'Sebastián Rondón - 20 Octubre 2022
    'Se consulta el valor del nuevo concepto a cobrar
    QRY = "SELECT tavalor FROM optag " _
            & "WHERE taplaca = '" & Txtnumpla.text & "' " _
            & "AND tafecpag is NULL " _
            & "AND taestado = 'A' "
            

    Set Tbl1 = Dbs.Execute(QRY)

    If ChkTag.Value = 1 Then

        If Tbl1.EOF = False Then
            MskValTag.text = CDbl(Tbl1!tavalor)
            If Mskcambio.text = "" Then
                Mskcambio.text = CDbl(Tbl1!tavalor)
            Else
                Mskcambio.text = CDbl(Mskcambio.text) + CDbl(Tbl1!tavalor)
            End If
        Else
            MskValTag.text = "0"
        End If
    Else
        If Tbl1.EOF = False Then
            MskValTag.text = "0"
            If Mskcambio.text <> "" And Mskcambio.text <> "0" Then
                Mskcambio.text = CDbl(Mskcambio.text) - CDbl(Tbl1!tavalor)
            End If
        Else
            MskValTag.text = "0"
        End If
    End If
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Capconduce")
End Sub

Function EsArregloInicializado(v As Variant) As Boolean
    On Error Resume Next
    EsArregloInicializado = IsArray(v) And Not IsEmpty(v) And (UBound(v) >= LBound(v))
End Function



Private Function SumarMultas() As Double
    On Error Resume Next
    Dim i As Integer, total As Double
    total = 0
    

       
    For i = LBound(multasGuardadas) To UBound(multasGuardadas)
        If IsNumeric(multasGuardadas(i).valor) Then
         total = total + Val(Replace(multasGuardadas(i).valor, ",", ""))
        End If
    Next i

    SumarMultas = total
End Function

Private Function SumarParqueo() As Double
    On Error Resume Next
    Dim i As Integer, total As Double
    total = 0

    For i = LBound(parqueGuardado) To UBound(parqueGuardado)
      If IsNumeric(parqueGuardado(i).valor) Then
        total = total + Val(Replace(parqueGuardado(i).valor, ",", ""))
       End If
    Next i

    SumarParqueo = total
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
   Tabular KeyAscii
End Sub

Private Sub Form_Load()
   On Error GoTo Errorres_Load
   L_Minutos_2_Conductores = Val(Fn_Parametros("HORA2COND", 1)) * 60
   L_Bandera_Escoger = "SI"
   bandera_huella = "N"
   L_Valida_CedHuella = ""
   bandera_huella1 = "N"
   L_Valida_CedHuella1 = ""
   LTipoConvenio = "N"
   'WECC 05/23/2016
   'Minutos para buscar despachos
   L_Minutos_Despacho = Fn_Parametros("MINDESPACH", 1)
   L_HabilitaDevueltas = UCase(Fn_Parametros("HABDEVUELT", 1))
   L_Valida_Conductor = UCase(Fn_Parametros("VALEMPCOND", 1))
   contador = 0
   txtcaja.text = Pc
   prealco = 1
   influencia = ""
   L_ValorTransitoSur = 0
'   Txtcedula.Enabled = False
'   Txtconact.Enabled = False
'   txtapellido.Enabled = False
   Txtoritra.text = "O"
   Desc_Server = True
   L_Bandera_CierreAutomatico = 0
   L_Bandera_ReimpresionTag = False
   
   'Sebastián Rondón 28 de Marzo de 2023
   'La terminal empieza a exigir la solicitud de tratamiento de datos personales
   chkDatosPersonales.Value = 1
   
   'SEP26 WILLIAM
   'SACAMOS EL NUMERO DEL CONDUCE DE LA TABLA DE CAJAS
   fqry01 = "SELECT caconcaj, caestado, catipo  " _
            & "FROM cocaja " _
           & "WHERE canomcaj = '" & Pc & "' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      If Ftb01!caestado = "Inactiva" Then
         MsgBox "La Caja Se Encuentra Inactiva", 16, "Error "
         Unload Me
         Exit Sub
      End If
      Txtnumcon.text = CDbl(Ftb01!caconcaj) + 1
      
      L_Terminal = Ftb01!catipo
   Else
      Btngrabar.Enabled = False
      Btnlimpiar.Enabled = False
   End If
   Ftb01.Close
  
   'hora actual del sistema
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      hora = Mid(Ftb01!psval, 12, 5)
      Mskhorsis.text = hora
      Fecha = Mid(Ftb01!psval, 1, 10)
      
      fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'ADERELOJ' "
      
      Set Tbl1 = Dbs.Execute(fqry01)
      If Not Tbl1.EOF Then
         hora1 = DateAdd("n", Tbl1(0), hora)
         horasis = Format(hora1, "hh:mm")
      End If
      Tbl1.Close
      Mskhora.Mask = ""

      Mskhora.text = CDate(hora1)
      Mskfecha.text = Fecha
   End If
   Ftb01.Close
   
   valparqueadero = Fn_Parametros("VALPAR", 1)
   tiempoesta = Fn_Parametros("MAXTIEPAR", 1)
   AviMulta = Fn_Parametros("AVIMULTA", 1)
   L_AvisoParqueadero = Fn_Parametros("AVIPARQUE", 1)
   NumCar = Fn_Parametros("NUMCAR", 1)
   L_MinVenTas = Fn_Parametros("MINVENTAS", 1)
   'wecc 07/24/2007
   '******************
   L_MinValTasa = Fn_Parametros("MINVALTASA", 1)
   L_MinValTasaCedula = Fn_Parametros("MINVALCEDU", 1)
   L_Conexion = Fn_Parametros("CONINTER", 1)
   L_Digita_Conductor = Fn_Parametros("DIGCONDUCT", 1)
   
   If UCase(G_Tipo_Terminal) = "SUR" Then
      'wecc 06/26/2008 parametro para hacer validacion con el terminal central _
      si compro TU
      L_MinCenSur = Fn_Parametros("MINCENSUR", 1)
      L_ValidaCentral = Fn_Parametros("VALIDACENT", 1)
      Txtoritra.text = "T"
      
   ElseIf UCase(G_Tipo_Terminal) = "NORTE" Then
      L_MinCenSur = Fn_Parametros("MINCENNOR", 1)
      L_ValidaCentral = Fn_Parametros("VALIDACENT", 1)
      Txtoritra.text = "T"
      
   Else
     Txtoritra.text = "O"
     
   End If
      
   fqry01 = "SELECT * " _
             & "FROM geparsis " _
            & "WHERE pscod = 'HORAOPE'" _
              & "AND psval = 'S'"
  
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      Horope = "S"
      Mskhora.Enabled = True
      'Mskhora.Mask = ""
      hora1 = CDate(hora1)
      hora1 = Format(hora1, "hh:mm")
      Mskhora.text = Format(hora1, "hh:mm")
      
   Else
      Horope = "N"
      Mskhora.Enabled = False
   End If
   Ftb01.Close
   direruta = ""
   Btnimprimir.Enabled = False
   
   '***
   
   'WECC Nov/25/2019 Caso CAU2 No  1907
   sBandera = Fn_Parametros("MINCONRUTA", 1)
   
   If sBandera = "" Then
      L_Minutos_Controlar_Rutas = 10
   Else
      L_Minutos_Controlar_Rutas = CInt(sBandera)
   End If
   
   
   
   L_DescontarAviso = Fn_Parametros("DESAVIPAGA", 1)
   If L_DescontarAviso = "" Then
      L_DescontarAviso = "N"
   End If
   
   If main.Mnucanmul.Enabled = True Then
      btnmultas.Enabled = True
   Else
      btnmultas.Enabled = False
   End If
   
   
   If G_habLecCedula = "S" Then
      txtCedula.MaxLength = 250
      TxtCedula1.MaxLength = 250
   Else
      txtCedula.MaxLength = 12
      TxtCedula1.MaxLength = 12
   End If
   
   
   QRY = "SELECT MAX(rutievia) AS rutievia " _
         & "FROM opruta " _
        & "WHERE ruestado = 'Activa'"
   
   Set Ftb01 = Dbs.Execute(QRY)
   
   If Not Ftb01.EOF Then
      L_Maximo_Tiempo_Ruta = (Hour(Ftb01!rutievia) * 60) + Minute(Ftb01!rutievia)
   End If
   
   If G_ManLector = "S" Then
      connectdb = ";DATABASE=" & DbPathHuella & ";PWD=cta2000"
      
      Set dbs4 = OpenDatabase("", False, False, connectdb)
      
   End If
   
   RegistoHuella = False
   
   'Sebastian Rondon Jun/13/2022
   L_IntegracionProtech = Fn_Parametros("INTPROTECH", 1)
  
   If L_IntegracionProtech = "" Then
      L_IntegracionProtech = "N"
   End If
   If L_IntegracionProtech = "S" Then
      L_UrlServer = Fn_Parametros("URLWSCOSAL", 1)
      L_UrlApiControlSalida = Fn_Parametros("URLAPICSTU", 1)
   End If
   
   MskValTag.text = "0"
   L_Bandera_Ruta_Lector = ""
   
   If Mid(G_Tipo_Terminal, 1, 1) = "N" Or Mid(G_Tipo_Terminal, 1, 1) = "S" Then
        BtnParqueadero.Enabled = False
        'btnmultas.Enabled = False
   End If
   
   'L_Bandera_Rodamiento = False
   
   'Sebastián Rondón - Abril 02 de 2025
   'Elimino los registros de OPRECCAJA de la caja con fechas anteriores
   'Para eliminar registros basura de la tabla
   Dbs.Execute "DELETE FROM OPRECCAJA " _
            & "WHERE RCNOMCAJ = '" & Pc & "' " _
            & "AND RCFECHA < TO_DATE('" & Mskfecha.text & "', 'MM/DD/YYYY') "
   
   Exit Sub
Errorres_Load:
   MsgBox Err.Description, vbCritical, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Mskdinero_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskdinero_KeyPress(KeyAscii As Integer)
    If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskdinero_LostFocus()
  
  If Mskcambio.text <> "0" And Mskcambio.text <> "" And Val(Mskdinero.text) <> 0 And Mskdinero.text <> "" Then
     If CDbl(Mskcambio.text) > CDbl(Mskdinero.text) Then
         MsgBox "Valor RECIBIDO es Menor al VALOR de LA TASA DE USO", 48, "Advertencia"
         Mskdinero.SetFocus
         Exit Sub
     Else
        mskdevolver.text = CDbl(Mskdinero.text) - CDbl(Mskcambio.text)
     End If
  End If
End Sub
Private Sub Mskfecha_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskhora_GotFocus()
   fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'HORASIS' "
        
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      
      L_HoraSist = Format(Mid(Ftb01!psval, 12, 5), "HH:mm")
   End If
   Ftb01.Close
   Mskhora.text = Format(Mskhora.text, "hh:mm")
   'horasis = Mskhora.Text

   Mskhora.Mask = "##:##"
   Mskhora.SelStart = 0
   Mskhora.SelLength = Len(Mskhora.text)
   'If Txtcodrut.Locked = False Then
      Fecha = Mid(main.StatusBar1.Panels(3), 16, 11)
      Mskfecha.text = Format(Fecha, "MM/dd/yyyy")
   'End If
End Sub

Private Sub Mskhora_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     'If Len(Mskhora.Text) = 5 Then
     '   Btngrabar.Default = True
     'End If
     
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskhora_LostFocus()
   On Error GoTo Errores
   Dim sBanderaOperador As Boolean
   L_Bandera_Fecha = "N"
   If Trim(L_CodigoEmpresaConvenio) <> "" Then
      L_Codigo_Empresa = L_CodigoEmpresaConvenio
   Else
      If Trim(TxtNumConvenio.text) = "" Then
         L_Codigo_Empresa = Txtcodemp.text
      Else
         L_Codigo_Empresa = L_CodigoEmpresaConvenioEmpresarial
      End If
   End If
   If LCodigoEmpresaIntegra = "" Or Trim(TxtNumConvenio.text) <> "" Then
      LCodigoEmpresaIntegra = Fn_Equivalencia_Empresa(L_Codigo_Empresa)
      If Trim(LCodigoEmpresaIntegra) = "" Then
         MsgBox "El código de la empresa no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
         Txtnumpla.Enabled = True
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   
   L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20)
   LTipoConvenio = "N"
   If Mskhora.ClipText <> "" Then
      longitud = Len(Mskhora.ClipText)
      If longitud = 5 Or longitud = 4 Then
         If Mskhora.ClipText < "0000" Or Mskhora.ClipText > "2400" Then
            MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
            Mskhora.SetFocus
            Mskhora.Mask = ""
            Mskhora.text = ""
            Mskhora.Mask = "##:##"
            Mskhora.text = horasis
            Exit Sub
         End If
         If longitud = 5 Then
            minutos = Mid(Mskhora.ClipText, 4)
         Else
            minutos = Mid(Mskhora.ClipText, 3)
         End If
         If minutos >= 60 Then
            MsgBox "El Valor De Minutos Que Coloco No Son Validos", 48, "Advertencia"
            Mskhora.SetFocus
            Mskhora.Mask = ""
            Mskhora.text = ""
            Mskhora.Mask = "##:##"
            Mskhora.text = horasis
            Exit Sub
         End If
         
         If L_Numero_Despacho <> 0 Then
            If L_Hora_Despacho <> Mskhora.text Then
               MsgBox "No se puede modificar la hora programada por la empresa" & Chr(13) & "                                      transportadora", 48, "Advertencia"
               Mskhora.text = L_Hora_Despacho
               Mskhora.SetFocus
               Exit Sub
            End If
         End If
         
         
         If Mskhora.text < horasis Then
            'wecc 11/14/2002 se parcha por el manejo de los turnos
            'If G_turno <= 2 Then
            If G_Turno < 4 Then
               MsgBox "La hora No Puede Ser Menor A La Hora Del Sistema", 48, "Advertencia"
               Mskhora.SetFocus
               'Mskhora.Mask = ""
               Mskhora.text = horasis
               Mskhora.SelStart = 0
               Mskhora.SelLength = Len(Mskhora.text)
            ElseIf G_Turno = 4 Then
                
               fqry01 = "SELECT psval " _
                        & "FROM geparsis " _
                       & "WHERE pscod = 'HORASIS' "
                      
               Set Ftb01 = Dbs.Execute(fqry01)
               
               If Ftb01.EOF = False Then
                  Fecha = Mid(Ftb01!psval, 1, 10)
                  hora = Mid(Ftb01!psval, 12, 5)
                  hora1 = "23:59"
                  hora2 = "22:00"
                  If (CDate(hora) <= CDate(hora1) And CDate(hora) > CDate(hora2)) Then
                     If CDate(Mskhora.text) < "02:00" Then
                        Mskfecha.text = DateAdd("d", 1, Mskfecha.text)
                        L_Bandera_Fecha = "S"
                     Else
                        If CDate(Mskhora.text) < horasis Then
                           MsgBox "La hora No Puede Ser Menor A La Hora Del Sistema", 48, "Advertencia"
                           Mskhora.SetFocus
                           'Mskhora.Mask = ""
                           Mskhora.text = horasis
                           Mskhora.SelStart = 0
                           Mskhora.SelLength = Len(Mskhora.text)
                        End If
                     End If
                  Else
                     If CDate(Mskhora.text) < hora Then
                        MsgBox "La hora No Puede Ser Menor A La Hora Del Sistema", 48, "Advertencia"
                        Mskhora.SetFocus
                        'Mskhora.Mask = ""
                        Mskhora.text = horasis
                        Mskhora.SelStart = 0
                        Mskhora.SelLength = Len(Mskhora.text)
                     End If
                  End If
               End If
               Ftb01.Close
            End If
         Else
            Fecha = Mid(main.StatusBar1.Panels(3), 16, 11)
            Mskfecha.text = Format(Fecha, "MM/dd/yyyy")
         End If
         
         If ConHora <> "" Then
            sHoraSalida = Format(Mskhora.text, "HH:mm")
            'qry = "SELECT * " _
                 & "FROM opdetconemp " _
                & "WHERE cdnumero = " & sNumConvenio & " " _
                  & "AND cdcodrut = '" & Txtcodrut.Text & "' " _
                  & "AND cdtipveh = '" & Txttipveh.Text & "' " _
                  & "AND cdhora = to_date('" & sHoraSalida & "','HH24:mi:ss')"
              
            QRY = "SELECT * " _
                  & "FROM opdetconemp, opconemp " _
                 & "WHERE cdnumero = " & sNumConvenio & " " _
                   & "AND cdcodrut = '" & Txtcodrut.text & "' " _
                   & "AND cdtipveh = '" & Txttipveh.text & "' " _
                   & "AND cdhora = '" & sHoraSalida & "' AND cenumero = cdnumero "
              
            Set Ftb01 = Dbs.Execute(QRY)
           
            If Ftb01.EOF Then
               Ftb01.Close
              
               If Txtoritra <> "T" Then
                  QRY = "SELECT * " _
                        & "FROM opdetconemp " _
                       & "WHERE cdnumero = " & sNumConvenio & " " _
                         & "AND cdcodrut = '" & Txtcodrut.text & "' " _
                         & "AND cdtipveh = '" & Txttipveh.text & "' ORDER BY cdhora " _
                     
                  sHorAuto = ""
                  Set Ftb01 = Dbs.Execute(QRY)
                  Do While Not Ftb01.EOF
                     If sHorAuto = "" Then
                        sHorAuto = Format(Ftb01!cdhora, "HH:mm")
                     Else
                        sHorAuto = sHorAuto & ", " & Format(Ftb01!cdhora, "HH:mm")
                     End If
                     Ftb01.MoveNext
                  Loop
                  Ftb01.Close
                  'If Format(Mskhora.Text, "HH:mm") <> Format(ConHora, "HH:mm") Then
                     MsgBox "      No se le Puede Vender Tasas de uso porque " & Chr(13) & "la hora del número convenio: " & sNumConvenio & ", no es la autorizada   " & Chr(13) & "Las Horas Autorizadas son: " & sHorAuto, 48, "Advertencia"
                     Txtnumpla.SetFocus
                     Txtnumpla.SelStart = 0
                     Txtnumpla.SelLength = Len(Txtnumpla.text)
                     Btnlimpiar_Click
                     Exit Sub
                  'End If
               End If
            Else
               L_Codigo_Empresa = Ftb01!cecodemp
               LTipoConvenio = Ftb01!cetipconv
               Ftb01.Close
               LCodigoEmpresaIntegra = Fn_Equivalencia_Empresa(L_Codigo_Empresa)
               If Trim(LCodigoEmpresaIntegra) = "" Then
                  MsgBox "El código de la empresa no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
                  Txtnumpla.Enabled = True
                  Txtnumpla.SetFocus
                  Exit Sub
               End If
               If L_CodigoEmpresaConvenio <> "" Then
                  fqry01 = "SELECT henombre FROM ophojemp WHERE hecodigo = '" & L_CodigoEmpresaConvenio & "' "
                 
                  Set Ftb01 = Dbs.Execute(fqry01)
   
                  If Not Ftb01.EOF Then
                     L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20) & "-" & Mid(Ftb01!henombre, 1, 20)
                  End If
                  Ftb01.Close
                 
               Else
                  L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20)
               End If
              
'              fqry01 = "SELECT henombre FROM ophojemp WHERE hecodigo = '" & L_Codigo_Empresa & "' "
'
'              Set Ftb01 = Dbs.Execute(fqry01)
'
'              If Not Ftb01.EOF Then
'                L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20) & "-" & Mid(Ftb01!henombre, 1, 20)
'              End If
'              Ftb01.Close
               sFecSal = Format(Mskfecha.text, "MM/dd/yyyy") & " " & Format(sHoraSalida, "HH:mm")
              
               QRY = "SELECT COUNT(conumero) " _
                     & "FROM coconduc " _
                    & "WHERE cofecsal = TO_DATE ('" & Mskfecha.text & "','MM/dd/yyyy') " _
                      & "AND cohorsal = TO_DATE('" & sFecSal & "','MM/dd/yyyy hh24:mi') " _
                      & "AND cocodemp = '" & Txtcodemp.text & "' " _
                      & "AND cocodrut = '" & Txtcodrut.text & "' " _
                      & "AND coclase = '" & Txttipveh.text & "' " _
                      & "AND coestado = 'Activo' "
              
               Set Ftb01 = Dbs.Execute(QRY)
              
               If Not Ftb01.EOF Then
                  If Not IsNull(Ftb01(0)) Then
                     If Ftb01(0) >= 1 Then
                        Ftb01.Close
                        MsgBox "Ya se vendió la tasa de uso al convenio " & sNumConvenio & " a esta hora", 48, "Advertencia"
                        'Mskhora.SetFocus
                        Btnlimpiar_Click
                        Exit Sub
                     End If
                  End If
               End If
               Ftb01.Close
            End If
         Else
            If L_CodigoEmpresaConvenioEmpresarial <> "" Then
               fqry01 = "SELECT henombre FROM ophojemp WHERE hecodigo = '" & L_CodigoEmpresaConvenioEmpresarial & "' "
              
               Set Ftb01 = Dbs.Execute(fqry01)

               If Not Ftb01.EOF Then
                  L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20) & "-" & Mid(Ftb01!henombre, 1, 20)
               End If
               Ftb01.Close
              
            Else
               L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20)
            End If
         
         End If
        
         '----------------------------------------------------------------------------------------
         'wecc se parcho porq se va a controlar la salida para las dos terminales CENTRAL Y SUR
         'MARZO 07 2012 VISITA EFECTUADA
        
        'If Trim(Txtoritra.Text) = "O" And UCase(G_Tipo_Terminal) = "CENTRAL" Then
        'If Trim(Txtoritra.Text) = "O" Then
        '----------------------------------------------------------------------------------------
           'valida la hora segun parametro MINVENTAS
            shorventas = DateAdd("n", L_MinVenTas, L_HoraSist)
   
            If CDate(Mskhora.text) > CDate(shorventas) Then
               MsgBox "No se le puede vender la tasa de uso", 16, "Venta Tasa de Uso"
               'Mskhora.Text = Format(L_HoraSist, "HH:mm")
               Btnlimpiar_Click
               'Mskhora.Enabled = True
               'Mskhora.SetFocus
               Exit Sub
            End If
   '
            
            '*****If Trim(Txtoritra.Text) = "O" Then
            'WECC CAU No 12706 solo valida cuando sea origen
               sfecsal1 = Mskfecha.text & " " & Mskhora.text
               sfecsal1 = Format(DateAdd("n", -L_MinValTasa, sfecsal1), "MM/dd/yyyy HH:mm")
               sfecsal2 = Mskfecha.text & " " & Mskhora.text
               sfecsal2 = Format(DateAdd("n", L_MinValTasa, sfecsal2), "MM/dd/yyyy HH:mm")
               If influencia <> "Si" And Txtoritra.text = "O" Then
'                  Qry = "SELECT svfecha " _
'                        & "FROM cosalveh " _
'                       & "WHERE svfecha BETWEEN TO_DATE('" & sfecsal1 & "','MM/dd/yyyy hh24:mi') " _
'                         & "AND TO_DATE('" & sfecsal2 & "','MM/dd/yyyy hh24:mi') " _
'                         & "AND svplaca = '" & Trim(Txtnumpla.Text) & "' " _
'                         & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
'
'                  Set Ftb01 = Dbs.Execute(Qry)
'
'                  If Not Ftb01.EOF Then
'                     MsgBox "            No se le puede vender tasa de uso, " & Chr(13) & "la última vez que la compro fue: " & Format(Ftb01!svfecha, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
'                     Ftb01.Close
'                     Mskhora.Enabled = True
'                     Mskhora.SetFocus
'                     Exit Sub
'                  End If
'                  Ftb01.Close
                  If L_EstViajePlaca = "S" Then
                     If Trim(Txtnumpla.text) <> "" Then
                        sfecsal1 = Mskfecha.text & " " & Mskhora.text
                        sMinutosBuscar = L_Maximo_Tiempo_Ruta * 2
                        sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
                        sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
                        L_Tiempo_Ruta_Placa = 0
                        If Valida_Placa_Conductor(Txtnumpla.text, sfecsal1, sfecsal2, "P", "O") = False Then
                           sfecsal3 = Format(DateAdd("n", L_Tiempo_Ruta_Placa, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                           MsgBox "No se le puede vender tasa de uso al vehículo, " & Chr(13) _
                           & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                           & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                           Btnlimpiar_Click
                           Exit Sub
                        End If
                     End If
                  End If
                  
                  sfecsal1 = Mskfecha.text & " " & Mskhora.text
                  sfecsal1 = Format(DateAdd("n", -L_MinValTasaCedula, sfecsal1), "MM/dd/yyyy HH:mm")
                  sfecsal2 = Mskfecha.text & " " & Mskhora.text
                  sfecsal2 = Format(DateAdd("n", L_MinValTasa, sfecsal2), "MM/dd/yyyy HH:mm")
                  
                  
                  'If Val(L_Tiempo_Ruta) >= Val(L_Minutos_2_Conductores) Then
                  If Txtoritra.text = "O" Then
                     
                     If L_EstViajeCedula1 = "S" Then
                        If Trim(txtCedula.text) <> "" Then
                           
                           sfecsal1 = Mskfecha.text & " " & Mskhora.text
                           If L_Tipo_Conductor1 = "R" Then
                              sMinutosBuscar = L_Tiempo_Ruta
                           Else
                              sMinutosBuscar = L_Tiempo_Ruta * 2
                           End If
                           
                           sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
                           sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
                           
                           
                           
   '                        Qry = "SELECT svfecha " _
   '                              & "FROM cosalveh " _
   '                             & "WHERE svfecha BETWEEN TO_DATE('" & sfecsal1 & "','MM/dd/yyyy hh24:mi') " _
   '                               & "AND TO_DATE('" & sfecsal2 & "','MM/dd/yyyy hh24:mi') " _
   '                               & "AND (svcedula = '" & Trim(txtcedula.Text) & "' OR svcedula1 = '" & Trim(txtcedula.Text) & "') " _
   '                               & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
   '
   '                        Set Ftb01 = Dbs.Execute(Qry)
   '
   '                        If Not Ftb01.EOF Then
   '                           MsgBox "No se le puede vender tasa de uso a la cedula " & txtcedula.Text & ", " & Chr(13) & "    la última vez que la compro fue: " & Format(Ftb01!svfecha, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
   '                           Ftb01.Close
   '                           Mskhora.Enabled = True
   '                           Mskhora.SetFocus
   '                           Exit Sub
   '                        End If
   '                        Ftb01.Close
                           
                           
                           If Valida_Placa_Conductor(txtCedula.text, sfecsal1, sfecsal2, "C", "O") = False Then
                              sfecsal3 = Format(DateAdd("n", sMinutosBuscar, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                              MsgBox "No se le puede vender tasa de uso al conductor, " & Chr(13) _
                              & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                              & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                              Btnlimpiar_Click
                              Exit Sub
                           End If
                           
                        End If
                     End If
                     If L_EstViajeCedula2 = "S" Then
                        If Trim(TxtCedula1.text) <> "" Then
                           sfecsal1 = Mskfecha.text & " " & Mskhora.text
                           If L_Tipo_Conductor2 = "R" Then
                              sMinutosBuscar = L_Tiempo_Ruta
                           Else
                              sMinutosBuscar = L_Tiempo_Ruta * 2
                           End If
                           
                           sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
                           sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
                           
   '                        Qry = "SELECT svfecha " _
   '                              & "FROM cosalveh " _
   '                             & "WHERE svfecha BETWEEN TO_DATE('" & sfecsal1 & "','MM/dd/yyyy hh24:mi') " _
   '                               & "AND TO_DATE('" & sfecsal2 & "','MM/dd/yyyy hh24:mi') " _
   '                               & "AND (svcedula = '" & Trim(Txtcedula1.Text) & "' OR svcedula1 = '" & Trim(Txtcedula1.Text) & "') " _
   '                               & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
   '
   '                        Set Ftb01 = Dbs.Execute(Qry)
   '
   '                        If Not Ftb01.EOF Then
   '                           MsgBox "No se le puede vender tasa de uso a la cedula " & Txtcedula1.Text & ", " & Chr(13) & "    la última vez que la compro fue: " & Format(Ftb01!svfecha, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
   '                           Ftb01.Close
   '                           Mskhora.Enabled = True
   '                           Mskhora.SetFocus
   '                           Exit Sub
   '                        End If
   '                        Ftb01.Close
                           If Valida_Placa_Conductor(TxtCedula1.text, sfecsal1, sfecsal2, "C", "O") = False Then
                              sfecsal3 = Format(DateAdd("n", sMinutosBuscar, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                              MsgBox "No se le puede vender tasa de uso al conductor, " & Chr(13) _
                              & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                              & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                              Btnlimpiar_Click
                              Exit Sub
                           End If
                           
                        End If
                     End If
                  End If
               Else
                  If Trim(Txtoritra.text) = "T" Then
                     If L_EstViajePlaca = "S" Then
                        If Trim(Txtnumpla.text) <> "" Then
                           sfecsal1 = Mskfecha.text & " " & Mskhora.text
                           sMinutosBuscar = 60
                           sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
                           sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
                           L_Tiempo_Ruta_Placa = 0
                           If Valida_Placa_Conductor(Txtnumpla.text, sfecsal1, sfecsal2, "P", "T") = False Then
                              sfecsal3 = Format(DateAdd("n", L_Tiempo_Ruta_Placa, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                              MsgBox "No se le puede vender tasa de uso al vehículo, " & Chr(13) _
                              & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                              & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                              Btnlimpiar_Click
                              Exit Sub
                           End If
                        End If
                     End If
                  End If
               End If
            '*****End If
            
         'End If
         '********************************************************************
         'funcion para validar que la empresa este asociada a un operador
         'wecc 06/08/2011
         '********************************************************************
         If L_Tipo_Venta = "O" Or L_Tipo_Venta = "M" Then
            If L_Operador <> "" Then
               If Fn_OperadorEmpresa(L_Operador, Txtcodemp.text) = False Then
                  
                  MsgBox "La empresa no esta asociada al operador", 48, "Advertencia"
                  'Mskhora.SetFocus
                  'Exit Sub
                  stipo_venta = "E"
               Else
                  stipo_venta = L_Tipo_Venta
               End If
            Else
               stipo_venta = L_Tipo_Venta
            End If
         Else
            stipo_venta = L_Tipo_Venta
         End If
         
         If Trim(Txtoritra.text) = "O" Then
            If Fn_ValidarRuta = False Then
                 sValida = "N"
            Else
               If G_Valida_Horario = "S" Then
                  If FnValidaHorario(stipo_venta) = True Then
                     sValida = L_Valida_Control_Ruta
                  Else
                     sValida = "N"
                  End If
               End If
            End If
            
            
            If ControlarCompra_TU(Txtcodrut.text, Txtcodemp.text, Mskfecha.text & " " & Mskhora.text, sValida) = False Then
               MsgBox "No se ha comprado tasa de uso para la ruta principal", 48, "Advertencia"
               Mskhora.SetFocus
               Exit Sub
            End If
         End If
         
         
         G_Imprime_Tasa_Uso = "E"
         If Trim(Txtoritra.text) = "O" Then
            If G_Valida_Horario = "S" Then
               If FnValidaHorario(stipo_venta) = False Then
                  
                  If L_HorasPermitidas = "" Then
                     MsgBox "   La empresa no tiene autorizado para despachar " & Chr(13) & "                            este tipo de servicio", 48, "Advertencia"
                  Else
                     MsgBox "   La empresa no tiene autorizado para despachar " & Chr(13) _
                           & "un vehículo en esta hora, sus horas permitidas son:  " & Chr(13) & L_HorasPermitidas, 48, "Advertencia"
                  End If
                  Mskhora.SetFocus
                  Exit Sub
                  
               End If
               
               If L_Valida_Horario = "S" Then
                  sHoraSalida = Format(Mskhora.text, "HH:mm")
                  sFecSal = Format(Mskfecha.text, "MM/dd/yyyy") & " " & Format(sHoraSalida, "HH:mm")
              
                  QRY = "SELECT COUNT(conumero) " _
                        & "FROM coconduc " _
                       & "WHERE cofecsal = TO_DATE ('" & Mskfecha.text & "','MM/dd/yyyy') " _
                         & "AND cohorsal = TO_DATE('" & sFecSal & "','MM/dd/yyyy hh24:mi') " _
                         & "AND cocodemp = '" & Txtcodemp.text & "' " _
                         & "AND cocodrut = '" & Txtcodrut.text & "' " _
                         & "AND coestado = 'Activo' "
                  
                   Set Ftb01 = Dbs.Execute(QRY)
                  
                   If Not Ftb01.EOF Then
                      If Not IsNull(Ftb01(0)) Then
                         If Ftb01(0) >= 1 Then
                            Ftb01.Close
                            MsgBox "Ya se compro una tasa de uso para las " & sHoraSalida & " para esta ruta", 48, "Advertencia"
                            Mskhora.SetFocus
                            Exit Sub
                         End If
                      End If
                   End If
                   Ftb01.Close
               End If
            End If
         End If
         '********************************************************************
         '********************************************************************
      Else
         'MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
         If Trim(Txtnumpla.text) = "" Then
            Mskhora.Mask = ""
            Mskhora.text = ""
            Mskhora.Mask = "##:##"
            Mskhora.text = horasis
            Txtnumpla.SetFocus
         Else
            Mskhora.SetFocus
            Mskhora.Mask = ""
            Mskhora.text = ""
            Mskhora.Mask = "##:##"
            Mskhora.text = horasis
         End If
         Exit Sub
      End If
   Else
   
      If L_Numero_Despacho <> 0 Then
         Mskhora.text = L_Hora_Despacho
         Mskhora.SetFocus
         Exit Sub
      Else
      
         If horasis <> "" Then
            Mskhora.text = horasis
         End If
      End If
      
   End If
  
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub


Private Sub Mskrecibo_LostFocus()
  If Mskvalor.text <> "0" And Mskvalor.text <> "" And Val(Mskrecibo.text) <> 0 And Mskrecibo.text <> "" Then
     If Val(Mskvalor.text) > Val(Mskrecibo.text) Then
         MsgBox "Valor RECIBIDO es Menor al VALOR del Tasa de Uso", 48, "Advertencia"
     Else
        Mskcambio = Mskvalor - Mskrecibo
     End If
  End If
End Sub

Private Sub OptTipo_Click(index As Integer)
   If index = 0 Then
      Lblcupo.Caption = "LINEA"
   Else
      Lblcupo.Caption = "CUPO"
   End If
End Sub


Private Sub OptTipoPlaca_Click(index As Integer)
   If index = 0 Then
      Txtnumpla.MaxLength = 6
   ElseIf index = 1 Then
      Txtnumpla.MaxLength = 7
   End If
End Sub

Private Sub Text1_Change()

End Sub

'Private Sub Timer1_Timer()
'    If Pc = "CAJA8" Then
'        If G_Turno = "4" Then
'            horven = Format(Mskhorsis.text, "hh:mm")
'            If (CDate(horven) >= "01:00" And CDate(horven) <= "01:10") Then
'                Btnsalir_Click
'                G_CierreAutomatico = 1
'                Cuaturno.Show 1
'                G_CierreAutomatico = 0
'            Else
'                If L_Bandera_CierreAutomatico = 0 Then
'                    If (CDate(horven) >= "00:50" And CDate(horven) <= "01:00") Then
'                        MsgBox "¡En 10 minutos se realizará el cierre de caja automático!", 48, "Advertencia"
'                        L_Bandera_CierreAutomatico = 1
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'    End If
'End Sub

Private Sub txtapellido_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtapellido_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Or (KeyAscii% = 209) Then
      Exit Sub
   Else
      If KeyAscii% = 13 Then
         Txtcodrut.SetFocus
      Else
         KeyAscii% = 0
      End If
   End If
End Sub


Private Sub txtapellido1_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtapellido1_KeyPress(KeyAscii As Integer)
Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Or (KeyAscii% = 209) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


Private Sub txtcedula_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcedula_KeyPress(KeyAscii As Integer)
   If G_habLecCedula <> "S" Then
      If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Then
         Exit Sub
      Else
         If KeyAscii% = 13 Then
            Txtconact.SetFocus
         Else
            KeyAscii% = 0
         End If
      End If
   Else
      Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))
      If KeyAscii = 124 Or KeyAscii = 39 Or KeyAscii = 34 Then
         KeyAscii% = 0
      Else
         Exit Sub
      End If
   
   End If

End Sub

Private Sub txtcedula_LostFocus()
   On Error GoTo err_txtcedula
   Txtconact.Enabled = True
   txtapellido.Enabled = True
   Txtconact.text = ""
   txtapellido.text = ""
   
   For i = 0 To 12
       datosCedula1(i) = ""
   Next
   If Trim(txtCedula.text) <> "" Then
        
      If G_habLecCedula = "S" And Len(txtCedula.text) > 10 Then
          Call leerCedula(txtCedula.text, datosCedula1)
          Txtconact.Enabled = False
          txtapellido.Enabled = False
          Txtconact.text = datosCedula1(6) & " " & datosCedula1(7)
          txtapellido.text = datosCedula1(4) & " " & datosCedula1(5)
          txtCedula.text = datosCedula1(3)
      End If
     
     
     
      QRY = "SELECT conombres, coapellidos, coprinom, coestado, cocodemp, cotipoconductor, NVL(coestviaje,'S') AS coestviaje " _
           & "FROM coconductor " _
          & "WHERE cocedula = '" & txtCedula.text & "'"
     
      Set Ftb04 = Dbs.Execute(QRY)
      L_EstViajeCedula1 = ""
      
      If Not Ftb04.EOF Then
         If Trim(Ftb04!coestado) = "Inactivo" Then
            MsgBox "El conductor se encuentra inactivo", 48, "Advertencia"
            Ftb04.Close
            txtCedula.SetFocus
            Exit Sub
         End If
         
         If L_Valida_Conductor = "S" Then
            If Trim(Ftb04!cocodemp) <> Trim(Txtcodemp.text) Then
               MsgBox "El conductor no estas asignado a la empresa", 48, "Advertencia"
               Ftb04.Close
               txtCedula.SetFocus
               Exit Sub
            End If
         End If
         L_EstViajeCedula1 = Ftb04!coestviaje
         If Ftb04!coestviaje = "S" Then
            If Trim(Txtoritra.text) = "O" Then
               
               sfecsal1 = Mskfecha.text & " " & Mskhora.text
               L_Tipo_Conductor1 = Ftb04!cotipoconductor
               If Ftb04!cotipoconductor = "R" Then
                  sMinutosBuscar = L_Tiempo_Ruta
               Else
                  sMinutosBuscar = L_Tiempo_Ruta * 2
               End If
               sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
               sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
               
               L_Fecha_salida = ""
               If influencia <> "Si" Then
                  
               
                  If Valida_Placa_Conductor(txtCedula.text, sfecsal1, sfecsal2, "C", "O") = False Then
                     sfecsal3 = Format(DateAdd("n", sMinutosBuscar, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                     MsgBox "No se le puede vender tasa de uso al conductor, " & Chr(13) _
                     & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                     & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                     Ftb04.Close
                     Btnlimpiar_Click
                     Exit Sub
                  End If
               End If
            End If
         End If
         
         
        'Sebastián Rondón - Abril 11 de 2023
        'Se empieza a validar que el conductor cuente con certificado EMGAF vigente
        QRY = "SELECT cocedula, NVL(TO_CHAR(cofecinicer, 'MM/DD/YYYY'), 'N/A') AS cofecinicer, NVL(TO_CHAR(cofecfincer, 'MM/DD/YYYY'), 'N/A') AS cofecfincer " _
              & "FROM coconductor " _
              & "WHERE cocedula = '" & txtCedula.text & "' "

        Set Tbl1 = Dbs.Execute(QRY)

        If Not Tbl1.EOF Then
            If Tbl1!cofecinicer <> "N/A" And Tbl1!cofecfincer <> "N/A" Then
                If CDate(Mskfecha.text) > CDate(Tbl1!cofecfincer) Then
                    G_AvisoEmgaf = "SEÑOR CONDUCTOR: EL CERTIFICADO MEDICO DE EMGAF SE ENCUENTRA VENCIDO O PROXIMO A VENCER." & Chr(13) & "ACERQUESE AL CONSULTORIO MEDICO DE ALCOHOLIMETRIA (ADITT-ASOTRANS) " & Chr(9) & "PARA RENOVARLO O ACTUALIZAR LA INFORMACIÓN EN CASO QUE LO TENGA DE OTRA TERMINAL." & Chr(13) & "FECHA DE VENCIMIENTO: " & Tbl1!cofecfincer
                    AviEmgaf.Show 1
                Else
                    sDias = DateDiff("d", CDate(Mskfecha.text), CDate(Tbl1!cofecfincer))
                    If sDias <= 15 Then
                        G_AvisoEmgaf = "SEÑOR CONDUCTOR: EL CERTIFICADO MEDICO DE EMGAF SE ENCUENTRA VENCIDO O PROXIMO A VENCER." & Chr(13) & "ACERQUESE AL CONSULTORIO MEDICO DE ALCOHOLIMETRIA (ADITT-ASOTRANS) " & Chr(9) & "PARA RENOVARLO O ACTUALIZAR LA INFORMACIÓN EN CASO QUE LO TENGA DE OTRA TERMINAL." & Chr(13) & "FECHA DE VENCIMIENTO: " & Tbl1!cofecfincer
                        AviEmgaf.Show 1
                    End If
                End If
            End If
        End If
    
         sCedulaBorrarHuella = txtCedula.text
         If Trim(Txtconact.text) = "" Then
            Txtconact.text = Ftb04!conombres
            txtapellido.text = Ftb04!coapellidos
            Txtconact.Enabled = False
            txtapellido.Enabled = False
         End If
         L_Transmitir_Conductor1 = "N"
         If IsNull(Ftb04!coprinom) Then
            L_Transmitir_Conductor1 = "S"
         Else
            Txtconact.Enabled = False
            txtapellido.Enabled = False
         End If
         ExisteCedula = 0
         If Txtconact.Enabled = True Then
             Txtconact.SetFocus
         End If
         
         If bandera_huella = "N" Or txtCedula.text <> L_Valida_CedHuella Then
         
            If G_ManLector = "S" And Trim(Txtoritra.text) = "O" Then
                  
               qry2 = "SELECT * " & _
                     "FROM coconductor " & _
                     "WHERE cocedula = '" & txtCedula.text & "' "
               
               Set Tbl20 = dbs4.OpenRecordset(qry2, dbOpenDynaset)
           
               If Not Tbl20.EOF Then
                   Respuesta_huella = MsgBox("Desea Usar La Huella Principal(Si) o La Huella Alternativa(No)", 64 + vbYesNo, "Advertencia")
                   
                   If Respuesta_huella = vbYes Then
                      If Not IsNull(Tbl20!cohuella) Then
                          'Tiene huella registrada y se comprueba
                          INICIALIZAR_LECTOR_HUELLA
                          RegistoHuella = False
                          VERIFICAR_HUELLA
                          If RegistoHuella = False Then
                              'WECC AGO04/2021 SE DESACTIVA POR SOLICITUD CAU No 4389
                              '*****************************************************************
                             Respuesta_elimina = MsgBox("La huella del conductor no corresponde, Desea Eliminar La Huella", vbExclamation + vbYesNo)
                             If Respuesta_elimina = vbYes Then
                                 QRY = " UPDATE coconductor SET cohuella = NULL, " & _
                                          " cohuealt = NULL,  " & _
                                          " codedo = 0, " & _
                                          " codedalt = 0 " & _
                                          " WHERE cocedula = '" & txtCedula.text & "' "

                                 Dbs.Execute (QRY)

                                 MsgBox "Huellas del conductor borrada", 64, "O.K."

                             End If
                              '*****************************************************************
                              MsgBox "La huella del conductor no corresponde, por favor dirigirse a la oficina del transportador", 48, "Advertencia"
                              txtCedula.SetFocus
                              Exit Sub
                           End If
                           bandera_huella = "S"
                           L_Valida_CedHuella = txtCedula.text
      '                    Txtcodrut.SetFocus
                          
                      Else
                          errorHuellero = False
                          INICIALIZAR_LECTOR_HUELLA
                          CAPTURAR_HUELLA
                          FINALIZAR_LECTOR_HUELLA
                          
                          If errorHuellero = True Then
                              MsgBox "Revise la conexion del huellero!!! Si no utiliza huellero desactivelo en el archivo de configuracion.", vbExclamation, "Huellero"
                              txtCedula.text = ""
                              Txtnumpla.text = ""
                              'Txtnumpla.SetFocus
                              Btnlimpiar_Click
                              Exit Sub
                          End If
                          
                      End If
                  ElseIf Respuesta_huella = vbNo Then
                      If Not IsNull(Tbl20!cohuealt) Then
                          'Tiene huella registrada y se comprueba
                          INICIALIZAR_LECTOR_HUELLA
                          RegistoHuella = False
                          VERIFICAR_HUELLA
                          If RegistoHuella = False Then
                              'WECC AGO04/2021 SE DESACTIVA POR SOLICITUD CAU No 4389
                              '*****************************************************************
                             Respuesta_elimina = MsgBox("La huella del conductor no corresponde, Desea Eliminar La Huella", vbExclamation + vbYesNo)
                             If Respuesta_elimina = vbYes Then
                                 'frmborrarhuella.Show 1
                                 QRY = " UPDATE coconductor SET cohuella = NULL, " & _
                                          " cohuealt = NULL,  " & _
                                          " codedo = 0, " & _
                                          " codedalt = 0 " & _
                                          " WHERE cocedula = '" & txtCedula.text & "' "

                                 Dbs.Execute (QRY)

                                 MsgBox "Huellas del conductor borrada", 64, "O.K."
                             End If
                              '*****************************************************************
                              MsgBox "La huella del conductor no corresponde, por favor dirigirse a la oficina del transportador", 48, "Advertencia"
                              txtCedula.SetFocus
                              Exit Sub
                           End If
                          bandera_huella = "S"
                          L_Valida_CedHuella = txtCedula.text
      '                    Txtcodrut.SetFocus
                          
                      Else
                          errorHuellero = False
                          INICIALIZAR_LECTOR_HUELLA
                          CAPTURAR_HUELLA
                          FINALIZAR_LECTOR_HUELLA
                          
                          If errorHuellero = True Then
                              MsgBox "Revise la conexion del huellero!!! Si no utiliza huellero desactivelo en el archivo de configuracion.", vbExclamation, "Huellero"
                              txtCedula.text = ""
                              Txtnumpla.text = ""
                              'Txtnumpla.SetFocus
                              Btnlimpiar_Click
                              Exit Sub
                          End If
                          
                      End If
                  End If
               'Else
               '    INICIALIZAR_LECTOR_HUELLA
               '    CAPTURAR_HUELLA
               '    FINALIZAR_LECTOR_HUELLA
               
               End If
               'Tbl20.Close
           
            End If
         End If
      Else
'         Txtconact.Enabled = True
'         txtapellido.Enabled = True
'         Txtconact.Text = ""
'         txtapellido.Text = ""
         ExisteCedula = 1
         L_Transmitir_Conductor1 = "S"
         If Trim(Txtconact.text) = "" Then
             On Error Resume Next
             Txtconact.Enabled = True
             txtapellido.Enabled = True
             Txtconact.text = ""
             txtapellido.text = ""
             Txtconact.SetFocus
         End If
        'On Error Resume Next
        'Txtconact.SetFocus
     End If
     Ftb04.Close
  Else
      Txtconact.text = ""
      txtapellido.text = ""
     
  End If
  
  
  Exit Sub
err_txtcedula:
   If Err.Number <> 13 Then
      MsgBox Err.Description, 16, "Error " & Err.Number
   Else
      txtCedula.text = ""
      Txtconact.text = ""
      txtapellido.text = ""
      MsgBox "Longitud incorrecta vuelva a leer la cédula " & Err.Description, 16, "Error Lectura " & Err.Number
      txtCedula.SetFocus
   End If
   
   Exit Sub
   Resume
End Sub


Private Sub txtcedula1_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcedula1_KeyPress(KeyAscii As Integer)
'If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
'     Exit Sub
'  Else
'     KeyAscii% = 0
'  End If
On Error GoTo Errores
   If G_habLecCedula <> "S" Then
      If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Then
         Exit Sub
      Else
         If KeyAscii% = 13 Then
            Txtconact1.Enabled = True
            Txtconact1.SetFocus
         Else
            KeyAscii% = 0
         End If
      End If
   Else
      Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))
      If KeyAscii = 124 Or KeyAscii = 39 Or KeyAscii = 34 Then
         KeyAscii% = 0
      Else
         Exit Sub
      End If
   
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

Private Sub txtcedula1_LostFocus()
   On Error GoTo err_txtcedula
   Txtconact1.Enabled = True
   Txtapellido1.Enabled = True
   Txtconact1.text = ""
   Txtapellido1.text = ""
   For i = 0 To 12
       datosCedula2(i) = ""
   Next
   If Trim(TxtCedula1.text) <> "" Then
   sMensaje = "1"
      If G_habLecCedula = "S" And Len(TxtCedula1.text) > 10 Then
      sMensaje = "2"
          Call leerCedula(TxtCedula1.text, datosCedula2)
          sMensaje = "3"
          Txtconact1.Enabled = False
          Txtapellido1.Enabled = False
          TxtCedula1.text = datosCedula2(3)
          Txtconact1.text = datosCedula2(6) & " " & datosCedula2(7)
          Txtapellido1.text = datosCedula2(4) & " " & datosCedula2(5)
          sMensaje = "4"
      End If
            
      If Trim(txtCedula.text) <> "" Then
      sMensaje = "5"
         If Trim(txtCedula.text) = Trim(TxtCedula1.text) Then
         sMensaje = "6"
            Txtconact1.text = ""
            Txtapellido1.text = ""
            Txtconact1.Enabled = True
            Txtapellido1.Enabled = True
            MsgBox "El número de cedulas son iguales", 16, "Error"
            TxtCedula1.SetFocus
            Exit Sub
         End If
      End If
      
      QRY = "SELECT conombres, coapellidos, coprinom, coestado, cocodemp, cotipoconductor, NVL(coestviaje,'S') AS coestviaje " _
            & "FROM coconductor " _
           & "WHERE cocedula = '" & TxtCedula1.text & "'"
      sMensaje = "7"
      Set Ftb04 = Dbs.Execute(QRY)
      sMensaje = "8"
      L_EstViajeCedula2 = ""
      
      If Not Ftb04.EOF Then
      sMensaje = "9"
         If Trim(Ftb04!coestado) = "Inactivo" Then
            MsgBox "El conductor se encuentra inactivo", 48, "Advertencia"
            Ftb04.Close
            TxtCedula1.SetFocus
            Exit Sub
         End If
         
         If L_Valida_Conductor = "S" Then
            If Trim(Ftb04!cocodemp) <> Trim(Txtcodemp.text) Then
               MsgBox "El conductor no estas asignado a la empresa", 48, "Advertencia"
               Ftb04.Close
               TxtCedula1.SetFocus
               Exit Sub
            End If
        End If
         sMensaje = "10"
         L_EstViajeCedula2 = Ftb04!coestviaje
         If Ftb04!coestviaje = "S" Then
            If Trim(Txtoritra.text) = "O" Then
               sMensaje = "11"
               sfecsal1 = Mskfecha.text & " " & Mskhora.text
               L_Tipo_Conductor2 = Ftb04!cotipoconductor
               If Ftb04!cotipoconductor = "R" Then
                  sMinutosBuscar = L_Tiempo_Ruta
               Else
                  sMinutosBuscar = L_Tiempo_Ruta * 2
               End If
               sfecsal1 = Format(DateAdd("n", -sMinutosBuscar, sfecsal1), "MM/dd/yyyy HH:mm")
               sfecsal2 = Format(Mskfecha.text & " " & Mskhora.text, "MM/dd/yyyy HH:mm")
               
               sMensaje = "12"
               If influencia <> "Si" Then
                  sMensaje = "13"
                  If Valida_Placa_Conductor(TxtCedula1.text, sfecsal1, sfecsal2, "C", "O") = False Then
                     sMensaje = "14"
                     sfecsal3 = Format(DateAdd("n", sMinutosBuscar, L_Fecha_salida), "MM/dd/yyyy HH:mm")
                     MsgBox "No se le puede vender tasa de uso al conductor, " & Chr(13) _
                     & "la última vez que la compro fue: " & Format(L_Fecha_salida, "MMM dd/yyyy HH:mm") & "," & Chr(13) _
                     & "se le puede vender tasa de uso apartir de: " & Format(sfecsal3, "MMM dd/yyyy HH:mm"), 48, "Advertencia"
                     Ftb04.Close
                     Btnlimpiar_Click
                     Exit Sub
                  End If
               End If
            End If
         End If
         
         'Sebastián Rondón - Abril 11 de 2023
            'Se empieza a validar que el conductor cuente con certificado EMGAF vigente
            QRY = "SELECT cocedula, NVL(TO_CHAR(cofecinicer, 'MM/DD/YYYY'), 'N/A') AS cofecinicer, NVL(TO_CHAR(cofecfincer, 'MM/DD/YYYY'), 'N/A') AS cofecfincer " _
                  & "FROM coconductor " _
                  & "WHERE cocedula = '" & TxtCedula1.text & "' "

            Set Tbl1 = Dbs.Execute(QRY)

            If Not Tbl1.EOF Then
                If Tbl1!cofecinicer <> "N/A" And Tbl1!cofecfincer <> "N/A" Then
                    If CDate(Mskfecha.text) > CDate(Tbl1!cofecfincer) Then
                        G_AvisoEmgaf = "SEÑOR CONDUCTOR: EL CERTIFICADO MEDICO DE EMGAF SE ENCUENTRA VENCIDO O PROXIMO A VENCER." & Chr(13) & "ACERQUESE AL CONSULTORIO MEDICO DE ALCOHOLIMETRIA (ADITT-ASOTRANS) " & Chr(9) & "PARA RENOVARLO O ACTUALIZAR LA INFORMACIÓN EN CASO QUE LO TENGA DE OTRA TERMINAL." & Chr(13) & "FECHA DE VENCIMIENTO: " & Tbl1!cofecfincer
                        AviEmgaf.Show 1
                    Else
                        sDias = DateDiff("d", CDate(Mskfecha.text), CDate(Tbl1!cofecfincer))
                        If sDias <= 15 Then
                            G_AvisoEmgaf = "SEÑOR CONDUCTOR: EL CERTIFICADO MEDICO DE EMGAF SE ENCUENTRA VENCIDO O PROXIMO A VENCER." & Chr(13) & "ACERQUESE AL CONSULTORIO MEDICO DE ALCOHOLIMETRIA (ADITT-ASOTRANS) " & Chr(9) & "PARA RENOVARLO O ACTUALIZAR LA INFORMACIÓN EN CASO QUE LO TENGA DE OTRA TERMINAL." & Chr(13) & "FECHA DE VENCIMIENTO: " & Tbl1!cofecfincer
                            AviEmgaf.Show 1
                        End If
                    End If
                End If
            End If
         
         sMensaje = "15"
         sCedulaBorrarHuella = TxtCedula1.text
         If Trim(Txtconact1.text) = "" Then
         sMensaje = "16"
            Txtconact1.text = Ftb04!conombres
            Txtapellido1.text = Ftb04!coapellidos
            Txtconact1.Enabled = False
            Txtapellido1.Enabled = False
            sMensaje = "17"
         End If
         L_Transmitir_Conductor2 = "N"
         If IsNull(Ftb04!coprinom) Then
            L_Transmitir_Conductor2 = "S"
         Else
            Txtconact1.Enabled = False
            Txtapellido1.Enabled = False
         End If
         
         If Txtconact1.Enabled = True Then
             Txtconact1.SetFocus
         End If
         sMensaje = "18"
         
         
         ExisteCedula1 = 0
         'Huellero YR Dic/11/2015
         If bandera_huella1 = "N" Or TxtCedula1.text <> L_Valida_CedHuella1 Then
         sMensaje = "20"
            If G_ManLector = "S" And Trim(Txtoritra.text) = "O" Then
            sMensaje = "21"
               qry2 = "SELECT * " & _
                      "FROM coconductor " & _
                      "WHERE cocedula = '" & TxtCedula1.text & "' "
                
                Set Tbl20 = dbs4.OpenRecordset(qry2, dbOpenDynaset)
                sMensaje = "22"
                If Not Tbl20.EOF Then
                    Respuesta_huella1 = MsgBox("Desea Usar La Huella Principal(Si) o La Huella Alternativa(No)", 64 + vbYesNo, "Advertencia")
                    If Respuesta_huella1 = vbYes Then
                       If Not IsNull(Tbl20!cohuella) Then
                           'Tiene huella registrada y se comprueba
                           INICIALIZAR_LECTOR_HUELLA
                           RegistoHuella1 = False
                           VERIFICAR_HUELLA1
                           If RegistoHuella1 = False Then
                              'WECC AGO04/2021 SE DESACTIVA POR SOLICITUD CAU No 4389
                              '*****************************************************************
                              Respuesta_elimina = MsgBox("La huella del conductor no corresponde, Desea Eliminar La Huella", vbExclamation + vbYesNo)
                              If Respuesta_elimina = vbYes Then
                                 'frmborrarhuella.Show 1
                                 QRY = " UPDATE coconductor SET cohuella = NULL, " & _
                                          " cohuealt = NULL,  " & _
                                          " codedo = 0, " & _
                                          " codedalt = 0 " & _
                                          " WHERE cocedula = '" & TxtCedula1.text & "' "

                                 Dbs.Execute (QRY)

                                 MsgBox "Huellas del conductor borrada", 64, "O.K."
                              End If
                              '*****************************************************************
                              MsgBox "La huella del conductor no corresponde, por favor dirigirse a la oficina del transportador", 48, "Advertencia"
                              TxtCedula1.SetFocus
                              Exit Sub
                           End If
                                    
                           bandera_huella1 = "S"
                           L_Valida_CedHuella1 = TxtCedula1.text
      '                     Txtcodrut.SetFocus
                           
                       Else
                           errorHuellero = False
                           INICIALIZAR_LECTOR_HUELLA
                           CAPTURAR_HUELLA1
                           FINALIZAR_LECTOR_HUELLA
                           
                           If errorHuellero = True Then
                               MsgBox "Revise la conexion del huellero!!! Si no utiliza huellero desactivelo en el archivo de configuracion.", vbExclamation, "Huellero"
                               TxtCedula1.text = ""
                               Txtnumpla.text = ""
                               'Txtnumpla.SetFocus
                               Btnlimpiar_Click
                               Exit Sub
                           End If
                           
                       End If
                     ElseIf Respuesta_huella1 = vbNo Then
                        If Not IsNull(Tbl20!cohuealt) Then
                           'Tiene huella registrada y se comprueba
                           INICIALIZAR_LECTOR_HUELLA
                           RegistoHuella1 = False
                           VERIFICAR_HUELLA1
                           If RegistoHuella1 = False Then
                              'WECC AGO04/2021 SE DESACTIVA POR SOLICITUD CAU No 4389
                              '*****************************************************************
                              Respuesta_elimina = MsgBox("La huella del conductor no corresponde, Desea Eliminar La Huella", vbExclamation + vbYesNo)
                              If Respuesta_elimina = vbYes Then
                                 'frmborrarhuella.Show 1
                                 QRY = " UPDATE coconductor SET cohuella = NULL, " & _
                                          " cohuealt = NULL,  " & _
                                          " codedo = 0, " & _
                                          " codedalt = 0 " & _
                                          " WHERE cocedula = '" & TxtCedula1.text & "' "

                                 Dbs.Execute (QRY)

                                 MsgBox "Huellas del conductor borrada", 64, "O.K."
                              End If
                              '*****************************************************************
                              MsgBox "La huella del conductor no corresponde, por favor dirigirse a la oficina del transportador", 48, "Advertencia"
                              TxtCedula1.SetFocus
                              Exit Sub
                           End If
                                    
                           bandera_huella1 = "S"
                           L_Valida_CedHuella1 = TxtCedula1.text
      '                     Txtcodrut.SetFocus
                           
                        Else
                           errorHuellero = False
                           INICIALIZAR_LECTOR_HUELLA
                           CAPTURAR_HUELLA1
                           FINALIZAR_LECTOR_HUELLA
                           
                           If errorHuellero = True Then
                               MsgBox "Revise la conexion del huellero!!! Si no utiliza huellero desactivelo en el archivo de configuracion.", vbExclamation, "Huellero"
                               TxtCedula1.text = ""
                               Txtnumpla.text = ""
                               'Txtnumpla.SetFocus
                               Btnlimpiar_Click
                               Exit Sub
                           End If
                           
                        End If
                     End If
                'Else
                '    INICIALIZAR_LECTOR_HUELLA
                '    CAPTURAR_HUELLA
                '    FINALIZAR_LECTOR_HUELLA
                
                End If
                'Tbl20.Close
            End If
         End If
         'fin huellero YR Dic/11/2015
      Else
         Ftb04.Close
         ExisteCedula1 = 1
         L_Transmitir_Conductor2 = "S"
         If Trim(Txtconact1.text) = "" Then
            Txtconact1.Enabled = True
            Txtapellido1.Enabled = True
            Txtconact1.text = ""
            Txtapellido1.text = ""
            Txtconact1.SetFocus
         End If
         Exit Sub
                
      End If
      Ftb04.Close
   Else
      ExisteCedula1 = 0
      Txtconact1.text = ""
      Txtapellido1.text = ""
   End If
   Exit Sub
err_txtcedula:
   If Err.Number <> 13 Then
      MsgBox Err.Description, 16, "Error " & Err.Number
   Else
      TxtCedula1.text = ""
      Txtconact1.text = ""
      Txtapellido1.text = ""
      MsgBox "Longitud incorrecta vuelva a leer la cédula " & Err.Description & "-" & sMensaje, 16, "Error Lectura " & Err.Number
      TxtCedula1.SetFocus
   End If
   
   Exit Sub
   Resume
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "c2"
  Txtcodemp.text = ""
  buscar.Show 1
  Txtcodemp.SetFocus
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Txtcodemp.SelStart = 0
   Txtcodemp.SelLength = Len(Txtcodemp.text)
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
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
     If Ftb01.EOF = False Then
        If Ftb01!heestado = "Inactiva" Then
            Ftb01.Close
            MsgBox "La empresa esta inactiva", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
         End If
         If Not IsNull(Ftb01!hevencon) Then
            sFecVigCon = Ftb01!hevencon
            If sFecVigCon <> "" Then
                If CDate(sFecVigCon) < CDate(Mskfecha.text) Then
                   MsgBox "El Convenio de la empresa ha caducado", 48, "Advertencia"
                   Btnlimpiar_Click
                   Exit Sub
                End If
             End If
         End If
         LCodigoEmpresaIntegra = Fn_Equivalencia_Empresa(Txtcodemp.text)
         If Trim(LCodigoEmpresaIntegra) = "" Then
            MsgBox "El código de la empresa no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
         End If
        Lblnomemp.Caption = Ftb01!henombre
        L_Empresa_Valida_Dia = Ftb01!hevalidadia
        L_EmpresaTransitoCentral = Ftb01!hetransitocentral
        L_Codigo_Empresa = Txtcodemp.text
        L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20)
        If Ftb01!hecredit = "S" Then
           chkcredito.Enabled = True
           chkcredito.Value = 1
        Else
           chkcredito.Enabled = False
           chkcredito.Value = 0
        End If
        Ftb01.Close
        'wecc MAY 05 2016 BUSCA EN FUNCION DE CUPOS
         Call Validar_Cupo(Txtcodemp.text, "-")
     Else
        MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
        Ftb01.Close
        Txtcodemp.text = ""
        Txtcodemp.SetFocus
     End If
  End If
End Sub

Private Sub txtcodrut_DblClick()
   
   If L_Numero_Despacho = 0 Then
      forma = "3"
      forma1 = "c3"
      Txtcodrut.text = ""
      
      
      If L_CodigoEmpresaConvenio <> "" Then
         G_Codigo_Empresa = L_CodigoEmpresaConvenio
      Else
         If TxtNumConvenio.text = "" Then
            G_Codigo_Empresa = Trim(Txtcodemp.text)
         Else
            QRY = "SELECT cecodemp,cetipconv " _
                  & "FROM opconemp " _
                 & "WHERE cenumero ='" & TxtNumConvenio.text & "' " _
                  & " AND ceempcon = '" & Txtcodemp.text & "' "
                         
            Set Ftb02 = Dbs.Execute(QRY)
            If Not Ftb02.EOF Then
               G_Codigo_Empresa = Ftb02!cecodemp
               LTipoConvenio = Ftb02!cetipconv
            End If
            Ftb02.Close
         End If
      End If
      
      buscar.Show 1
      'Txtcodrut.SetFocus
      'Txtcodrut_LostFocus
      Txtoritra.SetFocus
   End If
End Sub

Private Sub Txtcodrut_GotFocus()
   
   If pasoco = 1 Then
      MsgBox "El Vehiculo No Existe En La Base de Datos", 48, "Advertencia"
      Txtnumpla.SelStart = 0
      Txtnumpla.SelLength = Len(Txtnumpla.text)
      Txtnumpla.SetFocus
      Exit Sub
   Else
      If Trim(Txtnumpla.text) <> "" Then
         Txtcodrut.SelStart = 0
         Txtcodrut.SelLength = Len(Txtcodrut.text)
      Else
         Txtnumpla.SetFocus
      End If
      
   End If
End Sub

Private Sub txtcodrut_KeyPress(KeyAscii As Integer)
  
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     If KeyAscii% <> 8 Then
      Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))  ' mayuscula
      longitud = Len(Txtcodrut.text)
      longitud = longitud + 1
      If longitud = 3 Then
         codrut = "1"
         Txtcodrut_LostFocus
         Txtoritra.SetFocus
      End If
     End If
     codrut = "0"
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtcodrut_LostFocus()
   On Error GoTo ERR_Txtcodrut_LostFocus
   If Txtcodrut.text <> "" And Txttipveh.text <> "" Then
   
      If L_Bandera_Ruta_Lector <> "" Then
        Txtcodrut.text = L_Bandera_Ruta_Lector
      End If
        
      fqry01 = "SELECT * " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Txtcodrut.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb01!ruestado = "Inactiva" Then
             Txtcodrut.SetFocus
             MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
             Exit Sub
         End If
         
         L_TipoRuta = Ftb01!rutipo
         Call BuscacambioServicio(Txtcodemp.text, Txtcodrut.text)
                  
         If Not IsNull(Ftb01!rufecvig) Then
             sFecVig = Format(Ftb01!rufecvig, "MM/DD/YYYY")
             If CDate(sFecVig) < CDate(Mskfecha.text) Then
                 Txtcodrut.SetFocus
                 MsgBox "La Ruta tiene la Fecha de Vigencia Menor A La Actual", 48, "Advertencia"
                 Exit Sub
             End If
         End If
         '**************************************************************************
         'wecc 09/15/2014 CAU No 8658, vender tasas por dias segun el calendario
         If L_Empresa_Valida_Dia = "S" Then
            If Ftb01!ruvalidadia = "S" Then
               sMes = Month(Mskfecha.text)
               sDias = Day(Mskfecha.text)
               QRY = "SELECT dfmes " _
                   & "FROM badiafes " _
                   & "WHERE dfnumero =" & sMes & " " _
                   & "AND dfdias = " & sDias & " "
                 
               Set Tbl10 = Dbs.Execute(QRY)
            
               If Tbl10.EOF Then
                   Tbl10.Close
                   Txtcodrut.SetFocus
                   MsgBox "La empresa no tiene autorizada esta ruta para el día de hoy", 48, "Advertencia"
                   Exit Sub
               End If
            
            End If
         End If
         '**************************************************************************
         'wecc 06/04/2021 CAU No 4194, pico y placa
         L_PicoyPlaca = Ftb01!rupicoyplaca
         If Ftb01!rupicoyplaca = "S" Then
            If Fn_PicoyPlaca(Mskfecha.text, Mid(Txtnumpla.text, Len(Txtnumpla.text), 1)) = True Then
               MsgBox "Por Pico y Placa no se le puede vender tasa de uso", 48, "Advertencia"
               Txtnumpla.SetFocus
               Exit Sub
            End If
         End If
         
         '**************************************************************************
         
         L_Tiempo_Ruta = (Hour(Ftb01!rutievia) * 60) + Minute(Ftb01!rutievia)
         'wecc 05/03/2013 validacion cuando es origen y la ruta tiene mas de 8 horas o segun paramero HORA2COND
         Aviso_2_Conductores = True
         TxtCedula1.Enabled = False
         Txtconact1.Enabled = False
         Txtapellido1.Enabled = False
         
         If Val(L_Tiempo_Ruta) >= Val(L_Minutos_2_Conductores) Then
            'If Trim(Txtcedula1.Text) = "" Then
               MsgBox "En la ruta para " & Ftb01!rudestino & " deben de ir 2 conductores ", 48, "Advertencia"
               Aviso_2_Conductores = False
            'End If
            txtCedula.Enabled = True
            Txtconact.Enabled = True
            txtapellido.Enabled = True
            TxtCedula1.Enabled = True
            Txtconact1.Enabled = True
            Txtapellido1.Enabled = True
         Else
            TxtCedula1.text = ""
            Txtconact1.text = ""
            Txtapellido1.text = ""
            If L_Digita_Conductor = "N" And ExisteCedula = 0 Then
               If txtCedula.text <> "" Then
                   QRY = "SELECT conombres, coapellidos, coprinom " _
                       & "FROM coconductor " _
                       & "WHERE cocedula = '" & txtCedula.text & "'"

                   Set Ftb04 = Dbs.Execute(QRY)

                   If Not Ftb04.EOF Then
                       If Trim(Txtconact.text) = "" Then
                           Txtconact.text = Ftb04!conombres
                           txtapellido.text = Ftb04!coapellidos
                       End If
                   End If
               End If
               txtCedula.Enabled = False
               Txtconact.Enabled = False
               txtapellido.Enabled = False
            End If
         End If
         Txtcopias.text = Ftb01!rucopias
         TxtPeajeUltimo.text = "0"
         If Txtcopias.text = "1" Then
            TxtPeaje.text = "0"
            TxtPeaje.Enabled = False
            
         Else
            If G_PermiteLeerPeaje = "S" Then
               TxtPeaje.Enabled = True
               TxtPeaje.text = ""
               If Not IsNull(Ftb01!ruultpeaje) Then
                  TxtPeajeUltimo.text = Ftb01!ruultpeaje
               End If
            Else
               TxtPeaje.text = "0"
               TxtPeaje.Enabled = False
            End If
         End If
             Lblnomrut.Caption = Ftb01!rudestino
         L_Distancia = Ftb01!rudistan
         
         '***********************************************************************************
         'WECC 04/22/2013 Cuando el terminal es el central se almacena el valor de la tasa _
         para cuando es en transito en la terminal del sur o la q aplique _
         se transmite para consulta en el integra
         L_ValorTransitoSur = 0
         If G_Tipo_Terminal = "CENTRAL" Then
            If Val(L_Distancia) >= 0 And Val(L_Distancia) <= 150 Then
                sdistancia = 150
            ElseIf Val(L_Distancia) > 150 And Val(L_Distancia) <= 300 Then
                sdistancia = 300
            ElseIf Val(L_Distancia) >= 301 Then
                sdistancia = 9999
            End If
         
            QRY = "SELECT csdistan, csvalor " _
                    & "FROM opconsur " _
                    & "WHERE csclase = '" & L_Clase & "' " _
                    & "AND csdistan = " & sdistancia & " " _
                    & "AND csestado = 'A' " 'ORDER BY csdistan "
                   
            Set Ftb04 = Dbs.Execute(QRY)
            
            If Not Ftb04.EOF Then
                L_ValorTransitoSur = Ftb04!CSVALOR
            End If
            Ftb04.Close
             End If
            '***********************************************************************************
         LCodigoRutaIntegra = Fn_Equivalencia_Ruta(Txtcodrut.text)
         If Trim(LCodigoRutaIntegra) = "" Then
            Txtcodrut.SetFocus
            MsgBox "El código de la ruta no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
            Txtcodrut.SelStart = 0
            Txtcodrut.SelLength = Len(Txtcodrut.text)
            Exit Sub
         End If
         If IsNull(Ftb01!rudirecc) Then
             direruta = ""
         Else
             direruta = Ftb01!rudirecc
         End If
         LDirreccionRuta = direruta
         If IsNull(Ftb01!ruinflue) Then
             influencia = "No"
         Else
             influencia = Ftb01!ruinflue
         End If
         
         If L_TipoDespachoVehiculo = "I" And influencia = "No" Then
            'solo espacho influencia
            MsgBox "El vehículo solo se puede despachar en rutas de influencia", 48, "Advertencia"
            Btnlimpiar_Click
            Exit Sub
         ElseIf L_TipoDespachoVehiculo = "N" And influencia = "Si" Then
            'solo rutas nacionales
            MsgBox "El vehículo solo se puede despachar en rutas de intermunicipales", 48, "Advertencia"
            Btnlimpiar_Click
            Exit Sub
         End If
                 
         
         If G_PermiteInfluencia = "N" And influencia = "Si" Then
            MsgBox Pc & " NO permite venta de influencia", 48, "Advertencia"
            Txtcodrut.SetFocus
            Exit Sub
         End If
         
         'Sebastián Rondón - Agosto 08 de 2023
         'Consultamos si la ruta pertenece al catalogo de rutas por digitar pasajeros
         QRY = "SELECT dscoddet, dsdes FROM gesuptip " _
                    & "INNER JOIN gedetsuptip ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                    & "WHERE STDES = 'RUTAS PARA INGRESO PASAJEROS' " _
                    & "AND dscoddet = '" & Txtcodrut.text & "' "
         Set Tbl2 = Dbs.Execute(QRY)
         
         If Not Tbl2.EOF Then
            LblPasajeros.Visible = True
            Txtpasajeros.Visible = True
            Txtpasajeros.Enabled = True
         Else
            LblPasajeros.Visible = False
            Txtpasajeros.Visible = False
            Txtpasajeros.Enabled = False
            Txtpasajeros.text = ""
         End If
         
         Tbl2.Close
                 
         L_Operador = ""
         L_Tipo_Venta = ""
         If IsNull(Ftb01!rutipventa) Then
             L_Tipo_Venta = "E"
         Else
             L_Tipo_Venta = Ftb01!rutipventa
             If Not IsNull(Ftb01!ruoperador) Then
                 L_Operador = Ftb01!ruoperador
             End If
         End If
'         If IsNull(ftb01!ruvalhor) Then
'            L_Valida_Horario = "N"
'         Else
'            L_Valida_Horario = ftb01!ruvalhor
'         En    d If
         ConHora = ""
         ConNombre = ""
         'para ahorrar tiempo
          
         'si tiene convenio con otra empresa
         L_Tipo_Convenio = ""
         sNumConvenio = 0
         If Trim(L_CodigoEmpresaConvenio) <> "" Then
            QRY = "SELECT cenumero,cetipo, cefecvig, cecodemp, cefecini, cetipconv " _
                  & "FROM opconemp " _
                 & "WHERE ceempcon ='" & Txtcodemp.text & "' " _
                  & " AND cecodemp = '" & L_CodigoEmpresaConvenio & "' " _
                   & "AND ceestado = 'A' "
        
            L_Codigo_Empresa = Txtcodemp.text
        
            Set Ftb02 = Dbs.Execute(QRY)
            'If ftb02.EOF = False Then
            Do While Not Ftb02.EOF
            'ConNombre = ftb02!cenombre
                sTipo = Ftb02!cetipo
                L_Tipo_Convenio = sTipo
                LTipoConvenio = Ftb02!cetipconv
                QRY = "SELECT * " _
                    & "FROM opdetconemp " _
                    & "WHERE cdnumero = " & Ftb02!cenumero & " " _
                    & "AND cdcodrut = '" & Txtcodrut.text & "' " _
                    '& "AND cdtipveh = '" & Txttipveh.Text & "' "
          
                'sNumConvenio = ftb02(0)
                If Not IsNull(Ftb02(2)) Then
                    sFecVigCon = Ftb02(2)
                Else
                    sFecVigCon = ""
                End If
           
                If Not IsNull(Ftb02(2)) Then
                    sFecIniCon = Ftb02!cefecini
                Else
                    sFecIniCon = ""
                End If
          
                '*ftb02.Close
                sNumConvenio = Ftb02!cenumero
                Set Ftb03 = Dbs.Execute(QRY)
                If Ftb03.EOF = False Then
                    sNumConvenio = Ftb02(0)
                    If Not IsNull(Ftb03!cdhora) Then
                        If sTipo = "T" Then
                            'G_Validar_Ser = "S"
                            If Ftb03!cdhora <> "00:00" Then
                                If CDate(Ftb03!cdhora) < CDate(Mskfecha.text) Then
                                    MsgBox "El Número de Convenio: " & sNumConvenio & ", ha caducado", 48, "Advertencia"
                                    Btnlimpiar_Click
                                    Exit Sub
                                End If
                            End If
                    
                            If sFecIniCon <> "" Then
                                If CDate(sFecIniCon) > CDate(Mskfecha.text) Then
                                    MsgBox "El Número de Convenio: " & sNumConvenio & ", NO se encuentra Vigente", 48, "Advertencia"
                                    Btnlimpiar_Click
                                    Exit Sub
                                End If
                            End If
                    
                        '******************************************************
                            If sFecVigCon <> "" Then
                                If CDate(sFecVigCon) < CDate(Mskfecha.text) Then
                                    MsgBox "El Número de Convenio: " & sNumConvenio & ", ha caducado", 48, "Advertencia"
                                    Btnlimpiar_Click
                                    Exit Sub
                                End If
                            End If
                        '******************************************************
                        Else
                        ConHora = Format(Ftb03!cdhora, "HH:mm")
                    
                        If sFecIniCon <> "" Then
                            If CDate(sFecIniCon) > CDate(Mskfecha.text) Then
                                MsgBox "El Número de Convenio: " & sNumConvenio & ", NO se encuentra Vigente", 48, "Advertencia"
                                Btnlimpiar_Click
                                Exit Sub
                            End If
                        End If
                       
                        If sFecVigCon <> "" Then
                            If CDate(sFecVigCon) < CDate(Mskfecha.text) Then
                                MsgBox "El Número de Convenio: " & sNumConvenio & ", ha caducado", 48, "Advertencia"
                                Btnlimpiar_Click
                                Exit Sub
                                End If
                            End If
                        End If
                    End If
                Else
                    If sTipo = "E" Or sTipo = "T" Then
                        
                        respuesta = MsgBox("La ruta no esta autorizada en número de convenio: " & sNumConvenio & "," & Chr(13) & Chr(13) _
                              & "Desea despacharlo para la ruta autorizada de la empresa", 48 + vbYesNo, "Advertencia")
                        If respuesta = vbNo Then
                           On Error Resume Next
                           Txtcodrut.SetFocus
                           Txtcodrut.text = ""
                           Exit Sub
                        Else
                           LCodigoEmpresaIntegra = Txtcodemp.text
                           L_Codigo_Empresa = Txtcodemp.text
                           L_Tipo_Convenio = ""
                           L_CodigoEmpresaConvenio = ""
                           Exit Do
                        End If
                        
                        
                    End If
                End If
                Ftb03.Close
                L_Codigo_Empresa = Ftb02!cecodemp
                Ftb02.MoveNext
            Loop
            Ftb02.Close
         Else
            'wecc 05/07/2014 Caso CAU No 8172 referente a os convenios entre empresas de la terminal
            '***************************************************************************************
            '***************************************************************************************
            If Trim(TxtNumConvenio.text) <> "" Then
               QRY = "SELECT cenumero,cetipo, cefecvig, cecodemp, ceestado, cefecini, cetipconv  " _
                     & "FROM opconemp " _
                    & "WHERE cenumero ='" & TxtNumConvenio.text & "' " _
                     & " AND ceempcon = '" & Txtcodemp.text & "' "
                      
               L_Codigo_Empresa = Txtcodemp.text
               sNumConvenio = TxtNumConvenio.text
               Set Ftb02 = Dbs.Execute(QRY)
               If Not Ftb02.EOF Then
                  If Ftb02!ceestado = "I" Then
                     Ftb02.Close
                     MsgBox "El estado del Convenio: " & sNumConvenio & ", es inactivo", 48, "Advertencia"
                     TxtNumConvenio.SetFocus
                     Exit Sub
                  End If
                  If Ftb02!cetipo = "T" Then
                     Ftb02.Close
                     MsgBox "El Número de Convenio: " & sNumConvenio & ", es un convenio de CONTRATOS", 48, "Advertencia"
                     TxtNumConvenio.SetFocus
                     Exit Sub
                  End If
                  L_Tipo_Convenio = Ftb02!cetipo
                  LTipoConvenio = Ftb02!cetipconv
                  sFecVigCon = ""
                  If Not IsNull(Ftb02(2)) Then
                     sFecVigCon = Ftb02(2)
                  End If
                  sFecIniCon = ""
                  If Not IsNull(Ftb02(2)) Then
                     sFecIniCon = Ftb02!cefecini
                  End If
                  If sFecIniCon <> "" Then
                     If CDate(sFecIniCon) > CDate(Mskfecha.text) Then
                        MsgBox "El Número de Convenio: " & sNumConvenio & ", NO se encuentra Vigente", 48, "Advertencia"
                        Btnlimpiar_Click
                        Exit Sub
                     End If
                  End If
                  
                  If sFecVigCon <> "" Then
                     If CDate(sFecVigCon) < CDate(Mskfecha.text) Then
                        MsgBox "El Número de Convenio: " & sNumConvenio & ", ha caducado", 48, "Advertencia"
                        TxtNumConvenio.SetFocus
                        Exit Sub
                     End If
                  End If
                  
                  L_Codigo_Empresa = Ftb02!cecodemp
                  G_Codigo_Empresa = Ftb02!cecodemp
                  L_CodigoEmpresaConvenioEmpresarial = Ftb02!cecodemp
                  Ftb02.Close
                  
                  QRY = "SELECT cdtipveh, cdestado " _
                        & "FROM opdetconemp " _
                       & "WHERE cdnumero = " & TxtNumConvenio.text & " " _
                         & "AND cdcodrut = '" & Txtcodrut.text & "' " _
                  
                  Set Ftb03 = Dbs.Execute(QRY)
                  sBanderaTipoServicio = False
                  Do While Not Ftb03.EOF
                     
                     If Trim(Ftb03!cdtipveh) = Trim(Txttipveh.text) Then
                     
                        If Ftb03!cdestado = "I" Then
                           Ftb03.Close
                           MsgBox "El estado para el número de convenio: " & sNumConvenio & "," & Chr(13) & "para esta ruta esta inactivo", 48, "Advertencia"
                           TxtNumConvenio.SetFocus
                           Exit Sub
                        End If
                        sBanderaTipoServicio = True
                        Exit Do
                     End If
                     Ftb03.MoveNext
                  Loop
                  Ftb03.Close
                  
                  If sBanderaTipoServicio = False Then
                     Txtcodrut.SetFocus
                     MsgBox "El Número de Convenio: " & sNumConvenio & "," & Chr(13) & "para esta ruta no esta asignado en el tipo de servicio", 48, "Advertencia"
                     
                     Exit Sub
                  End If
               Else
                  Ftb02.Close
                  MsgBox "Número de convenio no pertencece a la empresa contratada", 48, "Advertencia"
                  TxtNumConvenio.SetFocus
                  Exit Sub
               End If
               '***************************************************************************************
               '***************************************************************************************
            End If
         End If
         '*End If
          
         'wecc 06/26/2008
         'verifica que el vehiculo haya comprado tasa de uso en la terminal central
         
         If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
            If L_ValidaCentral = "S" Then
               Valida_Terminal_Central ("RUTA")
            End If
         Else
            'Txtoritra.Locked = True 'mayo 19 2009
            
         End If
         '*********************
         If G_Validar_Ser = "S" Then
            QRY = "SELECT * " _
                  & "FROM opconvenios " _
                 & "WHERE cvcodrut = '" & Txtcodrut.text & "' " _
                   & "AND cvcodser = '" & Txttipveh.text & "' " _
                   & "AND cvcodemp = '" & L_Codigo_Empresa & "' "
            
            If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
               If L_ValidaCentral = "S" Then
                  QRY = QRY & " AND cvoritra = '" & Txtoritra.text & "' "
               End If
            End If
            
            
            
         ElseIf G_Validar_Ser = "N" Then
            QRY = "SELECT * " _
                  & "FROM opruthor " _
                 & "WHERE rhcodemp = '" & L_Codigo_Empresa & "' " _
                   & "AND rhcodrut = '" & Txtcodrut.text & "' "
            
            Set Ftb04 = Dbs.Execute(QRY)
            
            If Ftb04.EOF Then
               Ftb04.Close
               Txtcodrut.SetFocus
               MsgBox "La Empresa no tiene autorizada esta ruta", 48, "Advertencia"
               On Error Resume Next
               Txtcodrut.text = ""
               Exit Sub
            Else
               If Ftb04!rhestado = "Inactiva" Then
                  Ftb04.Close
                  sCodigoruta = Txtcodrut.text
                  Txtcodrut.SetFocus
                  Txtcodrut.text = ""
                  MsgBox "La ruta " & sCodigoruta & " esta inactiva para esta empresa", 48, "Advertencia"
                  
                  Exit Sub
               End If
            End If
            Ftb04.Close
            
            QRY = "SELECT * " _
                  & "FROM opconvenios " _
                 & "WHERE cvcodrut = '" & Txtcodrut.text & "' " _
                   & "AND cvcodser = '" & Txttipveh.text & "' "
            
            If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
               If L_ValidaCentral = "S" And Txtoritra.text = "O" Then
                  QRY = QRY & " AND cvoritra = '" & Txtoritra.text & "' "
               End If
            Else
               QRY = QRY & " AND cvoritra = 'O' "
            End If
            
         Else
            Txtcodrut.SetFocus
            MsgBox "El valor del parámetro VALIDARSER es " & G_Validar_Ser & ", debe ser S o N", 48, "Advertencia"
            Exit Sub
         End If
         
         If influencia = "Si" And (Mid(G_Tipo_Terminal, 1, 1) = "N" Or Mid(G_Tipo_Terminal, 1, 1) = "S") Then
            Txtoritra.text = "O"
            Txtoritra.Locked = True
         End If
         
         If Trim(Txtoritra.text) = "O" Then
            Set Ftb04 = Dbs.Execute(QRY)
            If Ftb04.EOF = False Then
               If Ftb04!cvestado = "I" Then
                  Ftb04.Close
                  Txtcodrut.SetFocus
                  
                  MsgBox "El valor de la tasa esta Inactivo", 48, "Advertencia"
                  Txtcodrut.SelStart = 0
                  Txtcodrut.SelLength = Len(Txtcodrut.text)
                  Exit Sub
               End If
                
               If UCase(G_Tipo_Terminal) = "CENTRAL" Then
                  Txtoritra.text = "O"
               Else
                  Txtoritra.text = Ftb04!cvoritra
               End If
               If IsNull(Ftb01!ruinflue) Then
                  influencia = "No"
               Else
                  influencia = Ftb01!ruinflue
               End If
               
               If Ftb04!cvoritra = "O" Then
                  If alcoholemia = "S" Then
                     If influencia = "Si" Then
                        Mskrecibo.text = G_ValalcInf
                     ElseIf influencia = "No" Then
                        Mskrecibo.text = G_ValalcOri
                     ElseIf influencia = "Ex" Then
                        Mskrecibo.text = "0"
                     End If
                  Else
                     If influencia = "Si" Then
                        Mskrecibo.text = G_ValalcInf
                     ElseIf influencia = "No" Then
                        Mskrecibo.text = G_ValalcOri
                     ElseIf influencia = "Ex" Then
                        Mskrecibo.text = "0"
                     End If
                  End If
               Else
                   
                  If influencia = "Si" Or influencia = "No" Then
                     Mskrecibo.text = G_ValalcTra
                  ElseIf influencia = "Ex" Then
                     Mskrecibo.text = "0"
                  End If
                  
               End If
               
                   QRY = "SELECT dscoddet, dsdes FROM gesuptip " _
                           & "INNER JOIN gedetsuptip ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                           & "WHERE STDES = 'RUTAS PARA INGRESO PASAJEROS' " _
                           & "AND dscoddet = '" & Txtcodrut.text & "' "
                Set Tbl2 = Dbs.Execute(QRY)
                
                If Not Tbl2.EOF Then
                   LblPasajeros.Visible = True
                   Txtpasajeros.Visible = True
                   Txtpasajeros.Enabled = True
                Else
                   LblPasajeros.Visible = False
                   Txtpasajeros.Visible = False
                   Txtpasajeros.Enabled = False
                   Txtpasajeros.text = ""
                End If
                
                Tbl2.Close
                       
               Mskvalor.text = Ftb04!cvvalor
               
               If MskParqueadero.text = "" Then
                  MskParqueadero.text = 0
               End If
               
               If MskValorTasaNorte.text = "" Then
                  MskValorTasaNorte.text = 0
               End If
               If MskMultas.text = "" Then
                  MskMultas.text = 0
               End If
               
               Mskcambio.text = CDbl(Mskvalor.text) + CDbl(Mskrecibo.text) + CDbl(MskParqueadero.text) + CDbl(MskValorTasaNorte.text) + CDbl(MskMultas.text) + CDbl(MskValTag.text)
               Ftb01.Close
               Ftb04.Close
               'Mskhora.Enabled = False
               FrmTipo.Enabled = True
               
               Call Validar_Cupo("-", Txtnumpla.text)
               
               If L_Numero_Cupo = 0 Then
                  Call Validar_Cupo(Txtcodemp.text, "-")
                       End If
               Exit Sub
                
            Else
               Ftb04.Close
               On Error Resume Next
               Txtcodrut.SetFocus
               MsgBox "La Empresa no tiene autorizada esta ruta ", 48, "Advertencia"
               Txtcodrut.text = ""
               Exit Sub
                  
               'wecc 12/18/2003
               'cuando la ruta tiene convenio con otra empresa
                  
                  
               'qry = "SELECT * " _
                     & "FROM opconemp " _
                    & "WHERE cecodrut = '" & txtcodrut.Text & "'" _
                      & "AND cecodser ='" & Txttipveh.Text & "' " _
                      & "AND ceempcon ='" & txtcodemp.Text & "' " _
                      & "AND ceestado = 'A' "
                         
                  
               'Set Ftb02 = Dbs.Execute(Qry)
               'If ftb02.EOF = False Then
               '  ConHora = Format(ftb02!cehora, "HH:mm")
               '  ConNombre = ftb02!cenombre
               '  If IsNull(ftb01!ruinflue) Then
               '     influencia = "No"
               '  Else
               '     influencia = ftb01!ruinflue
               '  End If
               '
               '
               '  If alcoholemia = "S" Then
               '     If influencia = "Si" Then
               '        Mskrecibo.Text = G_ValalcTra
               '     ElseIf influencia = "No" Then
               '        Mskrecibo.Text = G_ValalcOri
               '     ElseIf influencia = "Ex" Then
               '        Mskrecibo.Text = "0"
               '     End If
               '
               '  Else
               '     If influencia = "Si" Then
               '        Mskrecibo.Text = G_ValalcTra
               '     ElseIf influencia = "No" Then
               '        Mskrecibo.Text = G_ValalcOri
               '     ElseIf influencia = "Ex" Then
               '        Mskrecibo.Text = "0"
               '     End If
               '  End If
               '
               '
               '  Mskvalor.Text = ftb02!cevalor
               '  Mskcambio.Text = Val(Mskvalor.Text) + Val(Mskrecibo.Text)
               '  ftb01.Close
               '  ftb02.Close
                  
               'Else
               '  ftb01.Close
               '  txtcodrut.SetFocus
               '  txtcodrut.SelStart = 0
               '  txtcodrut.SelLength = Len(txtcodrut.Text)
               '  MsgBox "La Ruta No esta asignada a esta empresa ", 48, "Advertencia"
               '  Exit Sub
               'End If
            End If
         Else
'            Call Validar_Cupo("-", Txtnumpla.text)
'
'            If L_Numero_Cupo = 0 Then
'               Call Validar_Cupo(Txtcodemp.text, "-")
'            End If
         End If
         
         'fin ahorro tiempo
      Else
         If codrut = "0" Then
            Txtcodrut.SetFocus
            MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
            Txtcodrut.text = ""
         End If
      End If
   End If
   
   Exit Sub
   
ERR_Txtcodrut_LostFocus:
   If Err.Number <> 5 Then
      MsgBox Err.Description, 16, "Error " & Err.Number

   End If
   Exit Sub
   Resume
End Sub

Private Sub Txtconact_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtconact_KeyPress(KeyAscii As Integer)
  
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Or (KeyAscii% = 209) Then
      Exit Sub
   Else
      If KeyAscii% = 13 Then
         txtapellido.SetFocus
      Else
         KeyAscii% = 0
      End If
   End If
End Sub

Private Sub Txtconact_LostFocus()
'  If Txtnumpla.Enabled = False Then
'     On Error Resume Next
'     Txtcodrut.SetFocus
'  End If
End Sub


Private Sub Txtconact1_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtconact1_KeyPress(KeyAscii As Integer)
Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Or (KeyAscii% = 209) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub



Private Sub Txtcopias_KeyPress(KeyAscii As Integer)
If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtNumAlistamiento_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtNumAlistamiento_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
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
   On Error GoTo Errores
   
   Mskvalor.text = "0"
   MskValorTasaNorte.text = "0"
   Mskrecibo.text = "0"
   Mskcambio.text = "0"
   MskParqueadero.text = "0"
   sFuncioImpCon = ""
   sTurnoImpCon = 0
   If Trim(Txtnumcon.text) <> "" And Btnbuscar.Enabled = False Then
      fqry01 = "SELECT cocodrut, cocodemp, cofecsal, cohorsal, coplaca, " _
                   & " conumint, coclase, cocedula, coconduc, cocedula1, " _
                   & " coconduc1, cooritra, covalor, cofuncio, coturno, " _
                   & " coestado, comodo, covalalc, coparque, conomcaj, " _
                   & " cocredit, codirecc, coterminal, conumalistamiento, comodo " _
               & "FROM coconduc " _
              & "WHERE conumero = " & Txtnumcon.text & " AND conomcaj = '" & Pc & "'"
              
             
      Set Ftb01 = Dbs.Execute(fqry01)
          
      If Not Ftb01.EOF Then
      
         If Ftb01!coestado = "Anulado" Then
            MsgBox "La tasa de uso esta anulada", 48, "Advertencia"
            Btnlimpiar_Click
            Exit Sub
         End If
         L_Numero_Terminal_Norte = 0
         Txtnumpla.Enabled = False
         TxtNumInt.Enabled = False
         Txttipveh.Enabled = False
         Txtcodemp.Enabled = False
         Txtcodrut.Enabled = False
         Txtoritra.Enabled = False
         txtCedula.Enabled = False
         TxtCedula1.Enabled = False
         Txtconact.Enabled = False
         txtapellido.Enabled = False
         Txtconact1.Enabled = False
         Txtapellido1.Enabled = False
         TxtNumAlistamiento.Enabled = False
         Mskfecha.Enabled = False
         Mskhora.Enabled = False
         
         Btngrabar.Enabled = False
         
         Txtnumpla.text = Ftb01!coplaca
         TxtNumInt.text = Ftb01!conumint
         Txttipveh.text = Ftb01!coclase
         Txtcodemp.text = Ftb01!cocodemp
         Txtcodrut.text = Ftb01!cocodrut
         Txtoritra.text = Ftb01!Cooritra
         Mskfecha.text = Ftb01!cofecsal
         Mskhora.Mask = ""
         Mskhora.text = ""
         'Mskhora.Mask = "##:##"
         Mskhora.text = Format(Ftb01!cohorsal, "hh:mm")
         
         txtCedula.text = Ftb01!cocedula
         If Not IsNull(Ftb01!cocedula1) Then
            TxtCedula1.text = Ftb01!cocedula1
         End If
         TxtNumAlistamiento.text = Ftb01!conumalistamiento
         If Ftb01!comodo = "D" Then
            Lblcupo.Caption = "LINEA"
            OptTipo(0).Value = True
            OptTipo(1).Enabled = False
         Else
            Lblcupo.Caption = "CUPO"
            OptTipo(1).Value = True
            OptTipo(0).Enabled = False
         End If
         
         Mskvalor.text = Ftb01!covalor
         Mskrecibo.text = Ftb01!covalalc
         Mskcambio.text = Ftb01!covalor + Ftb01!covalalc
         sFuncioImpCon = Ftb01!cofuncio
         sTurnoImpCon = Ftb01!coturno
         Ftb01.Close
         Txttipveh_LostFocus
         'txtcodemp_LostFocus
         
         fqry01 = "SELECT henombre " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & Txtcodemp.text & "'"
             
         Set Ftb01 = Dbs.Execute(fqry01)
                 
         If Not Ftb01.EOF Then
            Lblnomemp.Caption = Ftb01!henombre
         End If
         Ftb01.Close
         L_Codigo_Empresa = Txtcodemp.text
         L_Nombre_Empresa = Lblnomemp.Caption
         
         fqry01 = "SELECT rudestino " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Txtcodrut.text & "'"
             
         Set Ftb01 = Dbs.Execute(fqry01)
                 
         If Not Ftb01.EOF Then
            Lblnomrut.Caption = Ftb01!rudestino
         End If
         Ftb01.Close
         
         fqry01 = "SELECT dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtoritra.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'TIPOS DE CONDUCE' "
             
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
            Lbloritra.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         
         If Trim(txtCedula.text) <> "" Then
             fqry01 = "SELECT conombres, coapellidos " _
                   & "FROM coconductor " _
                  & "WHERE cocedula = '" & txtCedula.text & "'"
         
             Set Ftb01 = Dbs.Execute(fqry01)
                     
             If Not Ftb01.EOF Then
                Txtconact.text = Ftb01!conombres
                txtapellido.text = Ftb01!coapellidos
             End If
             Ftb01.Close
         End If
         If Trim(TxtCedula1.text) <> "" Then
            fqry01 = "SELECT conombres, coapellidos " _
                  & "FROM coconductor " _
                 & "WHERE cocedula = '" & TxtCedula1.text & "'"
            
            Set Ftb01 = Dbs.Execute(fqry01)
                    
            If Not Ftb01.EOF Then
               Txtconact1.text = Ftb01!conombres
               Txtapellido1.text = Ftb01!coapellidos
            End If
            Ftb01.Close
         End If
         L_Bandera_ReimpresionTag = True
         Btnimprimir.Enabled = True
         Btnimprimir.SetFocus
      Else
         sFuncioImpCon = ""
         sTurnoImpCon = 0
         L_Bandera_ReimpresionTag = False
         MsgBox "Número de tasa de uso no existe", 48, "Advertencia"
         Btnlimpiar_Click
      End If
   
   End If
   
   Exit Sub
Errores:
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

Private Sub Txtnumint_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumpla_GotFocus()
   Txtnumpla.SelStart = 0
   Txtnumpla.SelLength = Len(Txtnumpla.text)
End Sub

Private Sub Txtnumpla_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   longitud = Len(Txtnumpla.text)
   longitud = longitud + 1
   L_Bandera_NumeroTag = 0
   If OptTipoPlaca(0).Value = True Then
      If longitud <= 2 Then
      
         If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
            
            Exit Sub
         Else
            'MsgBox "En El Numero De Placa las Tres Primeras Deben Ser Letras", 48, "Advertencia"
            'KeyAscii% = 0
            
            Exit Sub
         End If
      ElseIf longitud > 3 Then
         If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
            If longitud = 6 Then
               If KeyAscii% <> 8 Then
                  Txtnumpla_LostFocus
                  Txtcodrut.SetFocus
               End If
            End If
            Exit Sub
         Else
            If longitud >= 7 Then
               Txtnumpla.text = ""
               Exit Sub
            End If
            'MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
            KeyAscii% = 0
            Exit Sub
         End If
      End If
      If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
            Exit Sub
      Else
         'KeyAscii% = 0
      End If
   ElseIf OptTipoPlaca(1).Value = True Then
      
      If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
         If longitud = 7 Then
            If KeyAscii% <> 8 Then
               Txtnumpla_LostFocus
               Txtcodrut.SetFocus
            End If
         End If
         Exit Sub
      Else
         KeyAscii% = 0
      End If

      
   End If
End Sub

Private Sub Txtnumpla_LostFocus()
 
   On Error GoTo Message_Next
   Erase multasGuardadas 'para limpiar el arreglo con multas
   Erase parqueGuardado 'para limpiar el arreglo con parqueo
   Dim L_Entrada_Salida As String
   L_Bandera_Tasa_Condicionada = False
   L_Bandera_Venta_Condicionada = False
   pasoco = 0
   L_Entrada_Salida = ""
   L_Transmitir_Conductor1 = "N"
   L_Transmitir_Conductor2 = "N"
   MskMultas.text = "0"
   MskParqueadero.text = "0"
   Txtcodrut.Locked = False
   Txtoritra.Locked = False
   txtCedula.Locked = False
   TxtCedula1.Locked = False
   TxtNumAlistamiento.Locked = False
   L_Bandera_Ruta_Lector = ""
   G_multas_vacias = True
   G_parqueo_vacio = True
   
   
   
   'Sebastián Rondón - Marzo 06 2023
   'Setencia únicamente para presentanción Facturación Electrónica de Tasas Usos
   '********************************************************************************
   
'   QRY = "SELECT * FROM bolsa " _
'            & "WHERE btnit = '860052155-6' "
'
'    Set Tbl1 = Dbs.Execute(QRY)
'
'    If Not Tbl1.EOF Then
'        If Tbl1!btnumfac >= Tbl1!btcantidad Then
'            MsgBox "Se alcanzo el limite permitido de la bolsa. No se puede realizar la venta de Tasa de Uso", 48, "Advertencia"
'            Btnlimpiar_Click
'            Exit Sub
'        End If
'        fqry01 = "SELECT psval " _
'            & "FROM geparsis " _
'           & "WHERE pscod = 'HORASIS' "
'
'        Set Ftb01 = Dbs.Execute(fqry01)
'        If Ftb01.EOF = False Then
'            Fecha = Mid(Ftb01!psval, 1, 10)
'            If CDate(Fecha) > CDate(Tbl1!btfecha) Then
'                MsgBox "Se venció la fecha de resolución de la bolsa. No se puede realizar la venta de Tasa de Uso", 48, "Advertencia"
'                Btnlimpiar_Click
'                Exit Sub
'            End If
'        Else
'            MsgBox "No se pudo validar la fecha de resolución de la bolsa. No se puede realizar la venta de Tasa de Uso", 48, "Advertencia"
'            Btnlimpiar_Click
'            Exit Sub
'        End If
'    Else
'        MsgBox "No se encontró bolsa de facturación. No se puede realizar la venta de Tasa de Uso", 48, "Advertencia"
'        Btnlimpiar_Click
'        Exit Sub
'    End If
   
   '********************************************************************************
   
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Ftb01.EOF = False Then
      hora = Mid(Ftb01!psval, 12, 5)
      sHoraControlarCierre = Mid(hora, 4, 2)
      Mskhorsis.text = hora
      
      Fecha = Mid(Ftb01!psval, 1, 10)
       
      fqry01 = "SELECT psval " _
               & "FROM geparsis " _
              & "WHERE pscod = 'ADERELOJ' "
   
      Set Tbl1 = Dbs.Execute(fqry01)
      If Not Tbl1.EOF Then
         hora1 = DateAdd("n", Tbl1(0), hora)
         horasis = Format(hora1, "hh:mm")
      End If
      Tbl1.Close
      Mskhora.Mask = ""
      Mskhora.text = ""
      Mskhora.text = CDate(hora1)
      Mskfecha.text = Fecha
      
   End If
   Ftb01.Close
   
   
   If Txtnumpla.text <> "" Then
        
      Placa = Len(Txtnumpla.text)
       
      If Placa < Txtnumpla.MaxLength Then
         Exit Sub
      End If
      
      
      If Val(sHoraControlarCierre) Mod 5 = 0 Then
         If Valida_Hora_Cierre_Turno = False Then
            If Trim(Txtnumpla.text) = "" Then
               Btnlimpiar_Click
               Exit Sub
            End If

         End If
      Else
         If G_Turno = 4 Then
            If Valida_Hora_Cierre_Turno = False Then
               If Trim(Txtnumpla.text) = "" Then
                  Btnlimpiar_Click
                  Exit Sub
               End If
   
            End If
         End If
      End If
      
         'Sebastián Rondón - Enero 14 de 2025
      'Se valida que no se vendan tasas faltando 10 minutos de acabar el día
      If G_Turno = 4 Then
        If Valida_Hora_Cierre_Facturacion = False Then
            If Trim(Txtnumpla.text) = "" Then
               Btnlimpiar_Click
               Exit Sub
            End If
         End If
      End If
      
         If G_valida_rodamiento = "S" Then
          Call Validar_Despacho(Txtnumpla.text, hora)
    
            If L_Bandera_Rodamiento = False Then
    
                aviRoda.Show 1
    
                If G_usu_con_roda = False Then
    
                Txtnumpla.text = ""
                Exit Sub
    
                End If
            End If
       End If
      
      
   
      
       'Sebastian Rondón - Abril 08 de 2024
      'Validamos a primera mano si el vehículo cuenta con Rodamiento programado
       Call Validar_Despacho_Multas(Txtnumpla.text, hora)
      
      If UCase(G_Tipo_Terminal) = "CENTRAL" Then
      
        'Sebastián Rondón - 29 de Mayo de 2023
        'Se consulta parametro para definir si el sistema permite cobrar Duplicado de documentos
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'COBRODUPLI' "
        Set Tbl1 = Dbs.Execute(QRY)
      
         s_Numero_Comparendo_Antiguo = 0
         fqry01 = "SELECT * " _
                  & "FROM opmultas " _
                 & "WHERE munumpla = '" & Txtnumpla.text & "' " _
                   & "AND (muestado = 'Pendiente' " _
                    & "OR muestado = 'Abonada') "
                    
        If Tbl1!psval = "N" Then
            fqry01 = fqry01 & "AND mucodinf NOT IN (53) "
        End If
                   
         Set Ftb01 = Dbs.Execute(fqry01)
         If Ftb01.EOF = False Then
             
            Ftb01.Close
            G_BloqTasaUso = False
            Avimul.Show 1 'cuando le dan salir a este aviso se toman las multas seleccionadas, las sumamos más adelante
             
          '****************************************************************'
          ' Se bloquea la Tasa se uso si la multa es 13 (Ingreso a rampas  '
          ' sin autorizacion) segun requerimiento ingresada en el crentro  '
          ' de atencion a usuarios CAU con numero de incidencia 1128.      '
          ' Desarrollador: Miller Dorado                                   '
          ' Fecha: Mayo 11 de 2009                                         '
          '================================================================'
            If G_BloqTasaUso = True Then
   '            MsgBox "NO SE LE PUEDE VENDER TASA DE USO..." & Chr(13) _
   '                 & "INGRESO A RAMPLA SIN AUTORIZACION..." & Chr(13) _
   '                 & "DIRIJASE LA OFICINA DEL TRANSPORTADOR", 16, "AVISO."
               
               MsgBox "NO SE LE PUEDE VENDER TASA DE USO..." & Chr(13) _
                    & UCase(G_bloqTUxMulta) & "..." & Chr(13) _
                    & "DIRIJASE LA OFICINA DEL TRANSPORTADOR", 16, "AVISO."
               
               Btngrabar.Enabled = False
               Txtnumpla.SetFocus
               Txtnumpla.SelStart = 0
               Txtnumpla.SelLength = Len(Txtnumpla.text)
               Exit Sub
            End If
          '----------------------------------------------------------------'
            If G_BloqSin_Aviso = True Then
               sAvisosParametro = 0
            Else
               sAvisosParametro = AviMulta
            End If
            
            
             If Not TieneMultasSeleccionadas() Then 'significa que van a pagar multa, por lo que los avisos no cuetnan
            
                If NUMAVISOS >= Val(sAvisosParametro) Then
                   MsgBox "NO SE LE PUEDE VENDER TASA DE USO, AVISOS EXCEDIDOS..", 16, "AVISO."
                   Btngrabar.Enabled = False
                   Txtnumpla.SetFocus
                   Txtnumpla.SelStart = 0
                   Txtnumpla.SelLength = Len(Txtnumpla.text)
                   Exit Sub
                End If
                
            End If
            
         Else
            Ftb01.Close
            NUMAVISOS = 0
         End If
      End If
      
      LCodigoEmpresaIntegra = ""
      LCodigoRutaIntegra = ""
      LTipoServicioIntegra = ""
      L_Tipo_Conductor1 = ""
      L_Tipo_Conductor2 = ""
      L_EstViajePlaca = ""
      fqry01 = "SELECT * " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & Txtnumpla.text & "' "
             
  
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
          
         If Ftb01!hvestado = "Inactivo" Then
            MsgBox "EL VEHICULO SE ENCUENTRA INACTIVO," & vbCrLf & "  NO SE LE PUEDE VENDER TASAS DE USO", 16, "AVISO"
            Call Traer_Novedades(Txtnumpla.text)
            Txtnumpla.SetFocus
            Exit Sub
         End If
         L_EstViajePlaca = Ftb01!hvestviaje
         Call Traer_Novedades(Txtnumpla.text)
         ESTVEH = "No"
         If Not IsNull(Ftb01!hvestsel) Then
            ESTVEH = Ftb01!hvestsel
         End If
         L_TipoDespachoVehiculo = Ftb01!hvtipodespacho
         'alistamiento
         'wecc 05/03/2013
         '********************************************************************
         L_Fecha_Aviso_Alistamiento = ""
         If Not IsNull(Ftb01!hvfecalistamiento) Then
            L_Fecha_Aviso_Alistamiento = Format(Ftb01!hvfecalistamiento, "MM/DD/YYYY")
         End If
         
         '********************************************************************
         
         If Not IsNull(Ftb01!hvcarnet) Then
            If Ftb01!hvcarnet = "N" Then
               MsgBox "El vehículo no tiene FICHA", 48, "Advertencia"
            End If
         End If
         If Not IsNull(Ftb01!hvclase) Then
            L_Clase = Ftb01!hvclase
         Else
            MsgBox "El vehículo no tiene clase asignada", 48, "Advertencia"
            Txtnumpla.SetFocus
            Exit Sub
         End If
         
         'verifica si se no se ha cumplido el total de vehiculos
         'wecc 04/24/2002******************
         numsalveh = Ftb01!hvnumsal
         If G_NumSal <> "0" Then
            presentaalcoh = (Val(G_NumSal)) - numsalveh
         End If
              

               presenprueba = "Si"
               alcoholemia = "S"
               prealco = 1

        
         TxtNumInt.text = Ftb01!hvnumint
         Txtcodemp.text = Ftb01!hvcodemp
         Txttipveh.text = Ftb01!hvtipveh
         txtCedula.text = ""
         Txtconact.text = ""
         txtapellido.text = ""
         
         If IsNumeric(Ftb01!HVCONDUCT) Or Ftb01!HVCONDUCT <> 0 Then
            txtCedula.text = Ftb01!HVCONDUCT

         Else
            txtCedula.text = "0"
            Txtconact.text = "TERMINAL"
            txtapellido.text = "TRANSPORTES"
         End If
         
         TxtCedula1.text = ""
         Txtconact1.text = ""
         Txtapellido1.text = ""
         If IsNumeric(Ftb01!hvconduct1) And Ftb01!hvconduct1 <> "0" Then
             TxtCedula1.text = Ftb01!hvconduct1

         End If
        
         LClaseIntegra = Fn_Equivalencia_Servicio(Txttipveh.text)
         LTipoServicioIntegra = Traer_Descripcion(Txttipveh.text, "RELACION TIPO VEHICULO-TIPO SERVICIO")
         If Trim(LClaseIntegra) = "" Then
            MsgBox "El tipo de servicio no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
            Txtnumpla.Enabled = True
            Txtnumpla.SetFocus
            Exit Sub
         End If
         If Trim(LTipoServicioIntegra) = "" Then
            MsgBox "El tipo de servicio no tiene una relación con la clase", 48, "Advertencia"
            Txtnumpla.Enabled = True
            Txtnumpla.SetFocus
            Exit Sub
         End If
        
         LNumeroPuestosIntegra = 0
         If Not IsNull(Ftb01!hvnumpas) Then
            LNumeroPuestosIntegra = Ftb01!hvnumpas
         End If
        
         
         If IsNull(Ftb01!hventsal) Then
             L_Entrada_Salida = "S"
         Else
             L_Entrada_Salida = Ftb01!hventsal
         End If
        
         L_CodigoEmpresaConvenioEmpresarial = ""
         L_CodigoEmpresaConvenio = ""
         
         If Not IsNull(Ftb01!hvempcon) Then
             L_CodigoEmpresaConvenio = Ftb01!hvempcon
         End If
         
         fqry01 = "SELECT * " _
                  & "FROM ophojemp " _
                 & "WHERE hecodigo = '" & Txtcodemp.text & "' "
         
         Set Ftb01 = Dbs.Execute(fqry01)
         If Ftb01.EOF = False Then
            If Ftb01!heestado = "Inactiva" Then
               Ftb01.Close
               MsgBox "La empresa esta inactiva", 48, "Advertencia"
               Txtnumpla.SetFocus
               Exit Sub
            End If
            
            If Not IsNull(Ftb01!hevencon) Then
               sFecVigCon = Ftb01!hevencon
               If sFecVigCon <> "" Then
                  If CDate(sFecVigCon) < CDate(Mskfecha.text) Then
                     MsgBox "El Convenio de la empresa ha caducado", 48, "Advertencia"
                     Btnlimpiar_Click
                     Exit Sub
                  End If
               End If
               
            End If
              
             Lblnomemp.Caption = Ftb01!henombre
            L_Empresa_Valida_Dia = Ftb01!hevalidadia
            L_EmpresaTransitoCentral = Ftb01!hetransitocentral
            L_Codigo_Empresa = Txtcodemp.text
            L_Nombre_Empresa = Mid(Lblnomemp.Caption, 1, 20)
             
            If UCase(G_Tipo_Terminal) = "CENTRAL" Then
               If Ftb01!hecredit = "S" Then
                  chkcredito.Enabled = True
                  chkcredito.Value = 1
               Else
                  chkcredito.Enabled = False
                  chkcredito.Value = 0
               End If
            Else
               chkcredito.Enabled = False
               chkcredito.Value = 0
            End If
            
            If Txtcodrut.text <> "" Then
                Txtcodrut_LostFocus
            Else
                'wecc MAY 05 2016 BUSCA EN FUNCION DE CUPOS
                Call Validar_Cupo("-", Txtnumpla.text)
                
                If L_Numero_Cupo = 0 Then
                   Call Validar_Cupo(Txtcodemp.text, "-")
                End If
            End If
            
           
             
         Else
            Ftb01.Close
            fqry01 = "SELECT henombre " _
                     & "FROM ophojempcon " _
                    & "WHERE hecodigo = '" & Txtcodemp.text & "'"
             
            Set Ftb01 = Dbs.Execute(fqry01)
        
            If Not Ftb01.EOF Then
               Lblnomemp.Caption = Ftb01!henombre
              
            End If
         End If
         
         fqry01 = "SELECT dsdes " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txttipveh.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'TIPOS DE VEHICULO' "
               
         Set Ftb01 = Dbs.Execute(fqry01)
         If Ftb01.EOF = False Then
            Lbltipveh.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
    
          Call Validar_Despacho(Txtnumpla.text, hora)
      
         
         
         If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
            L_Pasajeros_Entran = 0
            If L_ValidaCentral = "S" Then
               Valida_Terminal_Central ("PLACA")
            End If
         Else
            'Txtoritra.Locked = True  'mayo 19 2009
         End If
         
         
            
        '***********************************************
        'Sebastián Rondón - 02 Noviembre - 2022
        'Validar si el vehículo cuenta con Tag para pagar
            
        'If UCase(G_Tipo_Terminal) = "CENTRAL" Then
            
        QRY = "SELECT * FROM optag " _
            & "WHERE taplaca = '" & Txtnumpla.text & "' " _
            & "AND tafecpag is NULL " _
            & "AND taestado = 'A' " _
            & "ORDER BY tanumero DESC"
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
        
            L_Bandera_CobroTag = True
            
            fqry01 = "SELECT psval FROM geparsis " _
                        & "WHERE pscod = 'NUMAVITAG' "
                        
            Set Ftb01 = Dbs.Execute(fqry01)
            
                If Ftb01.EOF = False Then
                
                If CDbl(Tbl1!taaviso) < CDbl(Ftb01!psval) Then
                
                    respuestaTag = MsgBox("El vehículo cuenta con pago de TAG pendiente." & Chr(13) & "              ¿Desea cobrarlo en esta Tasa?" & Chr(13) & "                     Avisos restantes: " & CDbl(Ftb01!psval) - CDbl(Tbl1!taaviso) & ". ", 64 + vbYesNo, "Advertencia")
                    
                    If respuestaTag = vbYes Then
                        L_Bandera_NumeroTag = Tbl1!tanumero
                        ChkTag.Value = 1
                    End If
                ElseIf CDbl(Tbl1!taaviso) >= CDbl(Ftb01!psval) Then
                    MsgBox "El vehículo cuenta con pago de TAG pendiente." & Chr(13) & "¡Último aviso!, se debe cobrar obligatoriamente.", 48, "Advertencia"
                    L_Bandera_NumeroTag = Tbl1!tanumero
                    ChkTag.Value = 1
                End If
                
            End If
            Ftb01.Close
        Else
            L_Bandera_CobroTag = False
        End If
        Tbl1.Close
            
        'End If
            
        '************************************************
         
         If G_Bloqueo_Tasa = "S" Then
            'wecc 04/07/2011, se parcho porque se modifico el manejo de validar si el vehiculo entro a al terminal _
            por el campo hvensal en al tabla ophojveh, esto se hizo para agilizar el proceso de venta de tasa de uso.
            
            'Qry = "SELECT * " _
                  & "FROM coingsal " _
                 & "WHERE isplaca = '" & Txtnumpla.Text & "' " _
                   & "AND isfecsal is null "
                    
            'Set Tbl1 = Dbs.Execute(Qry)
            
            'If Tbl1.EOF Then
            If L_Entrada_Salida = "S" Then
               'Tbl1.Close
               
               Avialc = "No"
               If UCase(G_Tipo_Terminal) = "SUR" Then
                  L_Entrada = "TS"
                  'respuesta = MsgBox("No Se Ha Registrado El Ingresado Del Vehículo Al Terminal." & Chr(13) & "                            Desea crear el ingreso", 64 + vbYesNo, "Advertencia")
                  respuesta = vbYes
               ElseIf UCase(G_Tipo_Terminal) = "NORTE" Then
                  L_Entrada = "TN"
                  'respuesta = MsgBox("No Se Ha Registrado El Ingresado Del Vehículo Al Terminal." & Chr(13) & "                            Desea crear el ingreso", 64 + vbYesNo, "Advertencia")
                  respuesta = vbYes
               Else
                  L_Entrada = "EC"
                  'MsgBox "No se ha registrado ingreso a las instalaciones del terminal," & Chr(13) & " por favor dirigirse a la oficina de servicio al transportador", 48, "Advertencia"
                  'respuesta = vbNo
                  'MsgBox "No se registro el ingreso del vehículo" & Chr(13) & "Se hara el ingreso automático a partir de la venta de Tasa de Uso", 48, "Advertencia"
                  respuesta = vbYes
               End If
               
               sobserv = "R" & Mid(Pc, Len(Pc))
               'sfecsistema = Fn_Parametros("HORASIS", 1)
               sFechaVenta = Format(Mskfecha.text, "MM/DD/YYYY")
               sHoraVenta = Format(Mskhorsis.text, "hh:mm")
               
               If respuesta = vbYes Then
                  If Txtoritra.text = "T" Then
                     stipollegada = "T"
                  Else
                     stipollegada = "I"
                  End If
                  
                  QRY = "INSERT INTO coingsal " _
                     & "(isplaca, isfecing, ishoring, " _
                     & " isfuncio, isturno, isingoritra, " _
                     & " isprualc, iscasent, isobserv, isruttra, istipollegada, isrutdes, ispasent) VALUES ('" _
                     & Txtnumpla.text & "',TO_DATE('" & Mskfecha.text & "','MM/DD/YYYY'),TO_DATE('" & sHoraVenta & "','hh24:mi'),'" _
                     & Login & "','" & G_Turno & "','O','" & Avialc & "','" & L_Entrada & "','" & sobserv & "','00','" & stipollegada & "','" & Trim(Txtcodrut.text) & "'," & L_Pasajeros_Entran & ")"
                  
                  Dbs.Execute QRY, AfecRegistro
                  
                  '******************************************
                  'ESTO ES LO NUEVO PARA DETECTAR EL QRY QUE SE ENVIA Y PORQUE NO SE GRABA
                  '******************************************
                  If AfecRegistro = 0 Then
                     MsgBox QRY, 48, "Advertencia"
                  End If
                  
                  '******************************************
                  sFechaIngreso = Format(Mskfecha.text, "MM/DD/YYYY") & " " & Format(sHoraVenta, "hh:mm")
                  
                  
                 QRY = "UPDATE ophojveh " _
                               & "SET hventsal = 'E', " _
                               & " hvfecing = TO_DATE('" & sFechaIngreso & "','MM/DD/YYYY HH24:MI'), hvtasa = 0 " _
                             & "WHERE hvplaca = '" & Txtnumpla.text & "' "
                             
                  Dbs.Execute QRY, AfecRegistro
                  
                  '******************************************
                  'ESTO ES LO NUEVO PARA DETECTAR EL QRY QUE SE ENVIA Y PORQUE NO SE GRABA
                  '******************************************
                  If AfecRegistro = 0 Then
                     MsgBox QRY, 48, "Advertencia"
                  End If
                  '******************************************
               Else
                  Btnlimpiar_Click
                  Exit Sub
               End If
               
            End If
            
            '**********************************************
            'Sebastián Rondón - 01 - Noviembre - 2022
            'Conocer si el vehículo se le registro entrada por parte de Protech
            If UCase(G_Tipo_Terminal) = "CENTRAL" Then
            
                'Consultamos primero la última entrada
                fqry01 = "SELECT * FROM ( " _
                                    & "SELECT isfecsal, TO_CHAR(ishorsal, 'HH24:MI:SS') AS ishorsal FROM coingsalp " _
                                    & "WHERE isfecsal IS NOT NULL " _
                                    & "AND isterminal = '" & Mid(UCase(G_Tipo_Terminal), 1, 1) & "' " _
                                    & "AND isplaca = '" & Txtnumpla.text & "' " _
                                    & "ORDER BY isfecsal DESC, ishorsal DESC " _
                                & ") WHERE ROWNUM <= 1 "
                Set Tbl1 = Dbs.Execute(fqry01)
                
                If Not Tbl1.EOF Then
'                    fqry01 = "SELECT isplaca FROM coingsalp " _
'                        & "WHERE isfecing >= TO_DATE('" & Tbl1!isfecsal & "','MM/DD/YYYY') " _
'                        & "AND ishoring >= TO_DATE('" & Tbl1!ishorsal & "','HH24:MI:SS') " _
'                        & "AND isfecsal is NULL " _
'                        & "AND isterminal = 'C' " _
'                        & "AND isplaca = '" & Txtnumpla.text & "' "

                     fqry01 = "SELECT isplaca FROM coingsalp " _
                        & "WHERE isfecing >= TO_DATE('" & Tbl1!isfecsal & "','MM/DD/YYYY') " _
                        & "AND isfecsal is NULL " _
                        & "AND isterminal = '" & Mid(UCase(G_Tipo_Terminal), 1, 1) & "' " _
                        & "AND isplaca = '" & Txtnumpla.text & "' "
                    
                    Set Tbl1 = Dbs.Execute(fqry01)
                
                    If Tbl1.EOF = True Then
                        'QRY1 = "INSERT INTO VEINGPROT (ptplaca, ptfecha, pthora) " _
                        '        & "VALUES('" & Txtnumpla.text & "', TO_DATE('" & Fecha & "', 'MM/DD/YYYY'), TO_DATE('" & horven1 & "', 'hh24:mi:ss') )"
                        'Dbs.Execute (QRY1)
                        'MsgBox "No se ha registrado el ingreso del vehículo por parte de Protech.", 48, "Advertencia"
                        
                        'Consultamos primero cuantos avisos tiene el vehículo
                        fqry02 = "SELECT hvnumaviprot, hvbandvencon FROM ophojveh " _
                                    & "WHERE hvplaca = '" & Txtnumpla.text & "' "
                        Set Tbl2 = Dbs.Execute(fqry02)
                        
                        If Not Tbl2.EOF Then
                            If Tbl2!hvnumaviprot > 0 Then
                                G_AvisoProtech = "EL VEHICULO YA HA ALCANZADO EL MAXIMO DE TASAS DE USO CONDICIONADA, SE BLOQUEA LA VENTA DE TASA DE USO, POR FAVOR ACERCARSE A LA OFICINA DE SERVICIO AL TRANSPORTADOR"
                                AviProtech.Show 1
                                Btnlimpiar_Click
                                Exit Sub
                            Else
                                If Tbl2!hvbandvencon < 1 Then
                                    G_AvisoProtech = "NO SE HA REGISTRADO EL INGRESO DEL VEHICULO A LAS INSTALACIONES DE LA TERMINAL, SE LE VENDERA TASA DE USO CONDICIONADA POR UNICA VEZ, POR FAVOR ACERCARSE A LA OFICINA DE SERVICIO AL TRANSPORTADOR"
                                    AviProtech.Show 1
                                    L_Bandera_Tasa_Condicionada = True
                                Else
                                    L_Bandera_Venta_Condicionada = True
                                End If
                            End If
                        End If
                        'Tbl2.Close
                    End If
                    'Tbl1.Close
                End If
            
                
            End If
            
         End If
         
         
                 
         
         '*********************
'         If Txtcedula.Enabled = True Then
'            Txtcedula.SetFocus
'         Else
            On Error Resume Next
            Txtcodrut.SetFocus
             On Error GoTo Message_Next
'         End If
      Else
         pasoco = 1
      End If
      L_Numero_Terminal_Norte = 0
      L_Valor_Alcoholimetria_Norte = 0
      MskValorTasaNorte.text = 0
      L_Placa_Permanencia = ""
      If UCase(G_Tipo_Terminal) = "CENTRAL" Then
         G_Avisos_Parqueadero = 0
         G_Cantidad_Parqueaderos = 0
         
         If Parqueadero = False Then
            L_Placa_Permanencia = Trim(Txtnumpla.text)
            'Btnlimpiar_Click
            Btngrabar.Enabled = False
            BtnParqueadero.Enabled = True
         End If
         Traer_Tasas_Terminal_Norte (Txtnumpla.text)
         
      End If
      
      
      'Recordemos que este sumar multas se saca del arreglo de multas, arreglo que se genera cuando se seleccionan en avisos y luego le dan a "salir"
    

      'WECC AGO/13/2019 Trae los valores que cancelo de permanencias y multas.
      Call TraerValorPermanenciasMultas
      
     'Mskparqueadero.Text = 0
     ' Mskotros.Text = 0
     
     
     'Camilo Campos 06/20/2025 el metodo sumarMultas y sumarParqueo quedan aparte debido a que el tipo de dato son objetos distintos
    'sumado a esto validamos que ambos tipos no esten vacios, ya que esto provoca un error EXCLUSIVO del ejecutable (funciona en pruebas, pero no en el EXE)
         If Not G_multas_vacias Then
            MskMultas.text = CDbl(MskMultas.text) + SumarMultas()
         End If
             
        If Not G_parqueo_vacio Then
           MskParqueadero.text = CDbl(MskParqueadero.text) + SumarParqueo()
        End If
  End If
salir:
   Exit Sub
   
Message_Next:
   If Err.Number = 400 Then
      Exit Sub
   ElseIf Err.Number = 52 Then
      
      If Desc_Server = True Then
         MsgBox "Se Ha Desconectado del Servidor", 48, "Desconectado"
      End If
      Desc_Server = False
      G_BaseServidor = False
      Resume Next

   
   Else
      MsgBox Err.Description, 16, "Error " & Err.Number
   End If
   
'err_txtplaca:
'   If Err.Number = 11 Then
'      presentar = 0
'      Resume Next
'   End If
 Exit Sub
 Resume '
End Sub

Private Sub Txtorigen_DblClick()
  forma = "3"
  forma1 = "c4"
  Txtorigen.text = ""
  buscar.Show 1
  Txtorigen.SetFocus
  Txtorigen_LostFocus
End Sub

Private Sub Txtorigen_KeyPress(KeyAscii As Integer)
If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtorigen_LostFocus()
   If Txtorigen.text <> "" And Txttipveh.text <> "" Then
      fqry01 = "SELECT rudestino, rutievia " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Txtorigen.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         Lblorigen.Caption = Ftb01!rudestino
         
      Else
         MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
         Txtorigen.text = ""
         Txtorigen.SetFocus
      End If
   End If
End Sub

Private Sub Txtoritra_GotFocus()
   On Error Resume Next
   Senalar_Texto Me
  ' Mskhora.Enabled = False
End Sub

Private Sub Txtoritra_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))
  
   If G_Tipo_Terminal <> "CENTRAL" Then
      If KeyAscii% = 79 Or KeyAscii% = 84 Or KeyAscii% = 8 Then
          Exit Sub
      Else
         KeyAscii% = 0
      End If
   Else
      If L_EmpresaTransitoCentral = "S" Then
         If KeyAscii% = 79 Or KeyAscii% = 84 Or KeyAscii% = 8 Then
             Exit Sub
         Else
            KeyAscii% = 0
         End If
      Else
         If KeyAscii% = 79 Or KeyAscii% = 8 Then
             Exit Sub
         Else
            KeyAscii% = 0
         End If
      End If
   End If
End Sub

Private Sub Txtoritra_LostFocus()
   On Error GoTo err_txtoritra
  
   If Txtoritra.text <> "" And Txttipveh.text <> "" Then
      
      
      If G_Tipo_Terminal = "CENTRAL" And Trim(Txtoritra.text) <> "O" And L_EmpresaTransitoCentral = "N" Then
         MsgBox "No se puede vender tasa de uso en transito en el terminal Central", 48, "Advertencia"
         Txtoritra.text = "O"
         Txtoritra.SetFocus
         Exit Sub
      End If
      
      If Trim(Txtoritra.text) = "O" Then
         Txtcopias.Enabled = False
      Else
         Txtcopias.Enabled = False
         Txtcopias.text = "1"
      End If
       
      If UCase(G_Tipo_Terminal) = "SUR" Or UCase(G_Tipo_Terminal) = "NORTE" Then
         If Txtoritra.text = "O" Then
            s_Numero_Comparendo_Antiguo = 0
            fqry01 = "SELECT * " _
                     & "FROM opmultas " _
                    & "WHERE munumpla = '" & Txtnumpla.text & "' " _
                      & "AND (muestado = 'Pendiente' " _
                       & "OR muestado = 'Abonada') "
                      
            Set Ftb01 = Dbs.Execute(fqry01)
            If Ftb01.EOF = False Then
                
               Ftb01.Close
               G_BloqTasaUso = False
               Avimul.Show 1
               
             '****************************************************************'
             ' Se bloquea la Tasa se uso si la multa es 13 (Ingreso a rampas  '
             ' sin autorizacion) segun requerimiento ingresada en el crentro  '
             ' de atencion a usuarios CAU con numero de incidencia 1128.      '
             ' Desarrollador: Miller Dorado                                   '
             ' Fecha: Mayo 11 de 2009                                         '
             '================================================================'
               If G_BloqTasaUso = True Then
      '            MsgBox "NO SE LE PUEDE VENDER TASA DE USO..." & Chr(13) _
      '                 & "INGRESO A RAMPLA SIN AUTORIZACION..." & Chr(13) _
      '                 & "DIRIJASE LA OFICINA DEL TRANSPORTADOR", 16, "AVISO."
                  
                  MsgBox "NO SE LE PUEDE VENDER TASA DE USO..." & Chr(13) _
                       & UCase(G_bloqTUxMulta) & "..." & Chr(13) _
                       & "DIRIJASE LA OFICINA DEL TRANSPORTADOR", 16, "AVISO."
                  
                  Btnlimpiar_Click
                  Exit Sub
               End If
             '----------------------------------------------------------------'
               If G_BloqSin_Aviso = True Then
                  sAvisosParametro = 0
               Else
                  sAvisosParametro = AviMulta
               End If
               
               If NUMAVISOS >= Val(sAvisosParametro) Then
                  MsgBox "NO SE LE PUEDE VENDER TASA DE USO..", 16, "AVISO."
                  Btnlimpiar_Click
                  Exit Sub
               End If
            Else
               Ftb01.Close
               NUMAVISOS = 0
            End If
         End If
      End If
      'Mskhora.Enabled = False
      fqry01 = "SELECT dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & Txtoritra.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'TIPOS DE CONDUCE' "
             
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         Lbloritra.Caption = Ftb01!dsdes
      Else
         Ftb01.Close
         MsgBox "Codigo De Origen/Transito No Existe", 48, "Advertencia"
         Txtoritra.SetFocus
         Exit Sub
      End If
      Ftb01.Close
      If Txtoritra.text = "T" Then
         'Txtorigen.Enabled = True
         'Label1(12).Enabled = True
         'Lblorigen.Enabled = True
         'Txtorigen.SetFocus
      Else
         Txtorigen.text = ""
         Lblorigen.Caption = ""
         Txtorigen.Enabled = False
         Label1(12).Enabled = False
         Lblorigen.Enabled = False
      End If
      
      If ConNombre = "" Then
         Aviso_Alistamiento = True
         
         If Txtoritra.text = "O" Then
            
            If Trim(Txtcodrut.text) <> "" Then
               
               If G_Validar_Ser = "S" Then
                
                 QRY = "SELECT * " _
                      & "FROM opconvenios " _
                     & "WHERE cvcodrut = '" & Txtcodrut.text & "' " _
                       & "AND cvoritra = '" & Txtoritra.text & "' " _
                       & "AND cvcodser = '" & Txttipveh.text & "' " _
                       & "AND cvcodemp = '" & L_Codigo_Empresa & "' "
               ElseIf G_Validar_Ser = "N" Then
                  QRY = "SELECT * " _
                        & "FROM opruthor " _
                       & "WHERE rhcodemp = '" & L_Codigo_Empresa & "' " _
                         & "AND rhcodrut = '" & Txtcodrut.text & "' "
                  
                  Set Ftb04 = Dbs.Execute(QRY)
                  
                  If Ftb04.EOF Then
                     On Error Resume Next
                     Ftb04.Close
                     Txtcodrut.SetFocus
                     MsgBox "La Empresa no tiene autorizada esta ruta ", 48, "Advertencia"
                     Txtcodrut.text = ""
                     Exit Sub
                  End If
                  Ftb04.Close
                  
                  QRY = "SELECT * " _
                        & "FROM opconvenios " _
                       & "WHERE cvcodrut = '" & Txtcodrut.text & "' " _
                         & "AND cvoritra = '" & Txtoritra.text & "' " _
                         & "AND cvcodser = '" & Txttipveh.text & "' "
               Else
                  On Error Resume Next
                  Txtcodrut.SetFocus
                  MsgBox "El valor del parámetro VALIDARSER es " & G_Validar_Ser & ", debe ser S o N", 48, "Advertencia"
                  Exit Sub
               End If
                       
               Set Ftb04 = Dbs.Execute(QRY)
               If Ftb04.EOF = False Then
                  If Ftb04!cvestado = "I" Then
                      MsgBox "El valor de tasa de uso esta inactivo ", 48, "Advertencia"
                      Txtoritra.SetFocus
                      Exit Sub
                  End If
                   
                  Mskvalor.text = Ftb04!cvvalor
                  Ftb04.Close
                  
                  
                  'If alcoholemia = "S" Then
                     If influencia = "Si" Then
                        Mskrecibo.text = G_ValalcInf
                     ElseIf influencia = "No" Then
                        Mskrecibo.text = G_ValalcOri
                     ElseIf influencia = "Ex" Then
                        Mskrecibo.text = "0"
                     End If
                   
'                  Else
'                     If influencia = "Si" Then
'                        Mskrecibo.Text = G_ValalcInf
'                     ElseIf influencia = "No" Then
'                        Mskrecibo.Text = G_ValalcOri
'                     ElseIf influencia = "Ex" Then
'                        Mskrecibo.Text = "0"
'                     End If
'                  End If
                     
                  If Mskrecibo.text = "" Then
                     Mskrecibo.text = 0
                  End If
                  If MskParqueadero.text = "" Then
                     MskParqueadero.text = 0
                  End If
                  If MskValorTasaNorte.text = "" Then
                     MskValorTasaNorte.text = 0
                  End If
                  If MskMultas.text = "" Then
                     MskMultas.text = 0
                  End If
                  'nuevo
                  'WECC OCT/05/2018 Proceso de escoger cuando la ruta es de influencia CAU2 109
                  If influencia = "Si" Then
                     If EscogerVehiculo = True Then
                        presenprueba = "Si"
                     Else
                        presenprueba = "No"
                        Mskrecibo.text = 0
                     End If
                  End If
                  
                  Mskcambio.text = CDbl(Mskvalor.text) + CDbl(Mskrecibo.text) + CDbl(MskParqueadero.text) + CDbl(MskValorTasaNorte.text) + CDbl(MskMultas.text) + CDbl(MskValTag.text)
                  
                  
                  
                  
                  'alistamiento
                  'wecc 05/03/2013
                  '********************************************************************
                  
                  'qry = "SELECT CONUMALISTAMIENTO " _
                         & "FROM COCONDUC  " _
                        & "WHERE COFECSAL = TO_DATE('" & Mskfecha.Text & "','MM/DD/YYYY')  " _
                          & "AND COPLACA ='" & Trim(Txtnumpla.Text) & "'  " _
                          & "AND COESTADO = 'Activo'  " _
                          & "AND CONUMALISTAMIENTO <> '0'  "
                            
                  QRY = "SELECT conumalistamiento " _
                         & "FROM vAlistamiento   " _
                        & "WHERE coplaca ='" & Trim(Txtnumpla.text) & "'  "
                            
                  
                  If L_Fecha_Aviso_Alistamiento = "" And (Trim(TxtNumAlistamiento.text) = "" Or Trim(TxtNumAlistamiento.text) = "0") Then
                     
                     If L_Conexion = "S" Then
                        Call Transmitir(QRY)
                     Else
                        L_Valida_Alistamiento = "False"
                     End If
                     If L_Valida_Alistamiento = "False" Then
                        'MsgBox "Falta el número del alistamiento", 48, "Advertencia"
                        'TxtNumAlistamiento.Enabled = False
                        'respuesta = MsgBox("Falta el número del alistamiento, desea ingresar el número?", 48 + vbYesNo, "Advertencia")
                        
                        Aviso_Alistamiento = False
                    '     Debug.Print Aviso_Alistamiento & " 1"
                        'If respuesta = vbYes Then
                           'TxtNumAlistamiento.Enabled = True
                           On Error Resume Next
'                           TxtNumAlistamiento.SetFocus
                           'Mskhora.Enabled = True
                           Exit Sub
                        'Else
                           
                        'End If
      
                     End If
                     
                     
                  Else
                     If L_Fecha_Aviso_Alistamiento <> "" Then
                        
                        If CDate(L_Fecha_Aviso_Alistamiento) < CDate(Mskfecha.text) Then
                           If (Trim(TxtNumAlistamiento.text) = "" Or Trim(TxtNumAlistamiento.text) = "0") Then
                              
                              If L_Conexion = "S" Then
                                 Call Transmitir(QRY)
                              Else
                                 L_Valida_Alistamiento = "False"
                              End If
                              
                              If L_Valida_Alistamiento = "False" Then
                                 'MsgBox "Falta el número del alistamiento", 48, "Advertencia"
                                 'TxtNumAlistamiento.Enabled = False
                                 'respuesta = MsgBox("Falta el número del alistamiento, desea ingresar el número?", 48 + vbYesNo, "Advertencia")
                                 Aviso_Alistamiento = False
                   '              Debug.Print Aviso_Alistamiento & " 2"
                                 'If respuesta = vbYes Then
      '                              Mskhora.Text = ""
      '                              Mskhora.Mask = "##:##"
                                    'TxtNumAlistamiento.Enabled = True
                                    On Error Resume Next
'                                    TxtNumAlistamiento.SetFocus
                                    'Mskhora.Enabled = True
                                    Exit Sub
                                 'End If
      
                              End If
                              
                           End If
      '                  ElseIf CDate(L_Fecha_Aviso_Alistamiento) = CDate(Mskfecha.Text) Then
      '                     'ya se hizo el alistamiento localmente
      ''                     'WECC 05/06/2013 SE LLAMA WEB SERVICES
      ''                     'PARA VERIFICAR QUE SE LO ALLA HEHCO EN OTRO TERMINAL QUE ALLA SALIDO EN ORIGEN _
      ''                     CUANDO ES UNA RUTA CORTA
      ''                     MsgBox "webservices"
      '                     If (Trim(TxtNumAlistamiento.Text) = "" Or Trim(TxtNumAlistamiento.Text) = "0") Then
      '                        Call Transmitir(qry)
      '
      '                        If L_Valida_Alistamiento = "False" Then
      '                           MsgBox "Falta el número del alistamiento", 48, "Advertencia"
      '                           Aviso_Alistamiento = False
      '                           MsgBox L_Valida_Alistamiento & " 3"
      '                        End If
      '                     End If
                        End If
                        
                     End If
                  End If
                  'Debug.Print Aviso_Alistamiento
                  'Mskhora.Enabled = True
                  '********************************************************************
                 
                  
                  
'                  Mskhora.Enabled = True
'                  Mskhora.SetFocus
                  Exit Sub
               Else
                  On Error Resume Next
                  MsgBox "La Ruta No esta asignada a esta empresa ", 48, "Advertencia"
                  Ftb04.Close
                  Txtcodrut.SetFocus
                  On Error GoTo 0
                  Exit Sub
               End If
               Ftb04.Close
            End If
         ElseIf Txtoritra.text = "T" Then
         
            'Sebastián Rondón - Febrero 19 de 2025
            'La terminal solicita que el sistema no permita venta de tasas de uso transito seguidos
            'Por lo cual determinamos si ya se vendio tasa de uso reciente para el vehículo
            If Valida_Ruta_Transito = False Then
                MsgBox "El vehículo compro recientemente tasa de uso en transito. " & vbNewLine & "Se cancela la venta.", 48, "Advertencia"
                Txtnumpla.text = ""
                Btnlimpiar_Click
                Exit Sub
            End If
            
            If Trim(Txtcodrut.text) <> "" Then
             'se busca en la tabla de opconsur
             'el valor se maneja por clase
                
                
                If Val(L_Distancia) >= 0 And Val(L_Distancia) <= 150 Then
                   sdistancia = 150
                ElseIf Val(L_Distancia) > 150 And Val(L_Distancia) <= 300 Then
                   sdistancia = 300
                ElseIf Val(L_Distancia) >= 301 Then
                   sdistancia = 9999
                End If
                
                QRY = "SELECT csdistan, csvalor " _
                      & "FROM opconsur " _
                     & "WHERE csclase = '" & L_Clase & "' " _
                       & "AND csdistan = " & sdistancia & " " _
                       & "AND csestado = 'A' " 'ORDER BY csdistan "
                       
                Set Ftb04 = Dbs.Execute(QRY)
                
                If Not Ftb04.EOF Then
                   Mskvalor.text = Ftb04!CSVALOR
                End If
                Ftb04.Close
                If influencia = "Si" Or influencia = "No" Then
                   Mskrecibo.text = G_ValalcTra
                   If presenprueba = "No" Then
                     Mskrecibo.text = 0
                   End If
                ElseIf influencia = "Ex" Then
                   Mskrecibo.text = "0"
                End If
                If Mskrecibo.text = "" Then
                   Mskrecibo.text = 0
                End If
                
                If MskParqueadero.text = "" Then
                   MskParqueadero.text = 0
                End If
                If MskValorTasaNorte.text = "" Then
                   MskValorTasaNorte.text = 0
                End If
                If MskMultas.text = "" Then
                  MskMultas.text = 0
               End If
               Mskcambio.text = CDbl(Mskvalor.text) + CDbl(Mskrecibo.text) + CDbl(MskParqueadero.text) + CDbl(MskValorTasaNorte.text) + CDbl(MskMultas.text) + CDbl(MskValTag.text)
'              Mskhora.Enabled = True
'              Mskhora.SetFocus

                Call Validar_Cupo("-", Txtnumpla.text)
               
                If L_Numero_Cupo = 0 Then
                   Call Validar_Cupo(Txtcodemp.text, "-")
                End If
               Exit Sub
                
            End If
         End If
     End If
  End If
  
   Exit Sub
  
err_txtoritra:
   If Err.Number = 11 Then
      presentar = 0
      Resume Next
   Else
      MsgBox Err.Description, 16, "Error 1" & Err.Number

      Exit Sub
      Resume
   End If
   
End Sub

Private Sub Txtpasajeros_KeyPress(KeyAscii As Integer)
    If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     If KeyAscii% <> 8 Then
      Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))  ' mayuscula
      longitud = Len(Txtpasajeros.text)
      'longitud = longitud + 1
      If longitud = 2 Then
        KeyAscii% = 0
      End If
     End If

     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtPasajeros_LostFocus()
    If Txtpasajeros.text <> "" Then
        If Txtnumpla.text <> "" Then
            'Buscamos el vehículo para verficar que la cantidad de pasajeros no supere a la capacidad máxima del vehículo
            QRY = "SELECT hvnumpas FROM ophojveh WHERE hvplaca = '" & Txtnumpla.text & "' "
            Set Tbl1 = Dbs.Execute(QRY)

            If Not Tbl1.EOF Then
                If CDbl(Txtpasajeros.text) > CDbl(Tbl1!hvnumpas) Then
                    MsgBox "El número de pasajeros digitados es mayor al número de pasajeros permitidos del vehículo", 48, "Advertencia"
                    Txtpasajeros.text = ""
                    Exit Sub
                End If
            Else
                MsgBox "No se encontro el vehículo digitado", 48, "Advertencia"
                Txtpasajeros.text = ""
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtPeaje_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPeaje_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub Txttipveh_DblClick()
  forma = "15"
  forma1 = "c2"
  Tabla_Catalogo = "TIPOS DE VEHICULO"
  Txttipveh.text = ""
  buscar.Show 1
  Txttipveh.SetFocus
  Txttipveh_LostFocus
End Sub


Private Sub Txttipveh_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
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
               
      Set Ftb02 = Dbs.Execute(fqry01)
      If Ftb02.EOF = False Then
         Lbltipveh.Caption = Ftb02!dsdes
         Ftb02.Close
         LClaseIntegra = Fn_Equivalencia_Servicio(Txttipveh.text)
         
         LTipoServicioIntegra = Traer_Descripcion(Txttipveh.text, "RELACION TIPO VEHICULO-TIPO SERVICIO")
         If Trim(LClaseIntegra) = "" Then
            MsgBox "El tipo de servicio no tiene un código de equivalencia de INTEGRA", 48, "Advertencia"
            Txttipveh.SetFocus
            Exit Sub
         End If
         If Trim(LTipoServicioIntegra) = "" Then
            MsgBox "El tipo de servicio no tiene una relación con la clase", 48, "Advertencia"
            Txttipveh.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "Tipo De Vehiculo No Existe", 48, "Advertencia"
         Txttipveh.text = ""
         Ftb02.Close
         Txttipveh.SetFocus
      End If
  
   End If
End Sub

Private Function Parqueadero() As Boolean
   Dim tiempo As Double
   Dim tiempar As Date
   
   On Error GoTo err_Parqueadero
   If Txtnumpla.text <> "" Then
      
   'wecc 08/25 2010
   'cau 3747 y segun seguimiento 20016
      L_Consecutivo_Parqueadero = 0
      Parqueadero = True
      
    'Sebastián Rondón - Enero 24 de 2023
    'Se consulta parámetro para la consulta de coparqueop
    'Mientras se decide que se va hacer con la data de esta tabla
    QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'FECPARPRO' "

    Set Tbl1 = Dbs.Execute(QRY)

    If Tbl1.EOF = False Then
        fecpar = Tbl1!psval
    End If
'
      fqry01 = "SELECT panumero, NVL(paaviso, 0) AS paaviso " _
               & "FROM coparqueop " _
              & "WHERE paplaca = '" & Txtnumpla.text & "' " _
                & "AND paestado = 'A' AND pavalor <> 0 " _
                & "AND pafecini >= TO_DATE('" & fecpar & "', 'MM/DD/YYYY') "

      Set Tbl1 = Dbs.Execute(fqry01)
      If Tbl1.EOF = False Then
          
          Do While Not Tbl1.EOF
            If Tbl1!paaviso >= Val(L_AvisoParqueadero) Then
                sAviso = Tbl1!paaviso
            End If
            Tbl1.MoveNext
          Loop
         'Ftb01.Close
         
         AviParqu.Show 1
         
         If Not TieneParqueaderosSeleccionados() Then
             If sAviso >= Val(L_AvisoParqueadero) Then
                'cuando llega al limite de avisos obliga a cancelar el valor
                MsgBox "Debe de cancelar el valor del parqueadero " & Chr(13) & "ya llegó al limite de avisos de parqueadero", 48, "Advertencia"
                Parqueadero = False
    '            respuesta = 6
    '         Else
    '            respuesta = MsgBox("Desea cancelar el parqueadero", 64 + vbYesNo, "Imprimir")
                
             End If
         End If
'         If Txtnumpla.text <> "" Then
'             L_Placa_Permanencia = Txtnumpla.text
'         End If
'         capparqu.Txtnumpla.text = L_Placa_Permanencia
'         Unload Me
'         capparqu.Show 1


'         If respuesta = 6 Then
'            L_Consecutivo_Parqueadero = G_Numero_Parqueadero
'            qry = "SELECT pavalor " _
'                  & "FROM coparqueo " _
'                 & "WHERE panumero = " & G_Numero_Parqueadero & " "
'
'            Set Tbl1 = Dbs.Execute(qry)
'
'            If Not Tbl1.BOF Then
'               Mskparqueadero.Text = Tbl1!pavalor
'            End If
'            Tbl1.Close
'         Else
            MskParqueadero.text = 0
            L_Consecutivo_Parqueadero = 0
'         End If
         
         
         
         
         
      End If
      
      
      
      
      
'      Qry = "SELECT isfecing, ishoring " _
'            & "FROM coingsal " _
'           & "WHERE isplaca = '" & Txtnumpla.Text & "' " _
'             & "AND isfeccon is null AND isfecsal is null "
'
'      Set Tbl1 = Dbs.Execute(Qry)
'      If Not tbl1.EOF Then
'         feching = tbl1!isfecing
'         horing = tbl1!ishoring
'         Mskingreso.Text = Format(tbl1!ishoring, "HH:mm")
'
'         sFecEntrada = feching & " " & Mskingreso.Text
'
'         sFecSistema = Fn_Parametros("HORASIS", 1)
'
'         sminutosDif = (DateDiff("n", sFecEntrada, sFecSistema))
'         sdiferencia = (DateDiff("n", sFecEntrada, sFecSistema))
'
'         sMinutosIniciarCobro = 120       'minutos para iniciar cobro de parqueadero
'         sMinutosMaximos = 720            'el cobro es de 12 horas o fraccion es decir 720 minutos
'
'         'Controlamos que solocobre parqueadero  contando 12 horas a partir de los minutos inial de cobro
'         If sminutosDif >= sMinutosMaximos Then
'           sdiferencia = sminutosDif - sMinutosIniciarCobro
'         Else
'           sdiferencia = sminutosDif
'         End If
'
'         If sdiferencia > sMinutosIniciarCobro Then
'            cociente = sdiferencia \ sMinutosMaximos
'            residuo = sdiferencia Mod sMinutosMaximos
'            Mskparqueadero.Text = "0"
'            If residuo < sMinutosIniciarCobro Then
'               If residuo <> 0 Then
'                  cociente = cociente + 1
'               End If
'
'            Else
'               cociente = cociente + 1
'            End If
'            Mskparqueadero.Text = (valparqueadero * cociente) '+ CDbl(Mskparqueadero.Text)
'         Else
'            Mskparqueadero.Text = 0
'            ValorParqueo = Mskparqueadero.Text
'         End If
'
'
'
'
'
'      Else
'        MsgBox "NO se ha registrado el Ingresado del vehiculo al Terminal ", 48, "Advertencia"
'        Txtnumpla.SetFocus
'        Mskparqueadero.Text = 0
'      End If
'      tbl1.Close
   End If
   'Tbl1.Close
   Exit Function
   
err_Parqueadero:
   MsgBox Err.Description, 16, "Error " & Err.Number

End Function

'Funciones Huellero
Public Sub INICIALIZAR_LECTOR_HUELLA()
    CheckPeople1.Inicializar
End Sub

Public Sub FINALIZAR_LECTOR_HUELLA()
    CheckPeople1.Finalizar
End Sub

Public Sub VERIFICAR_HUELLA()
   LEER_DATOS_HUELLA
   
   If CheckPeople1.Verificar(HuellaAlc(), lblCodigo.Caption) Then
      RegistoHuella = True
   Else
      RegistoHuella = False
   End If
   
   FINALIZAR_LECTOR_HUELLA
   
End Sub

Public Sub LIMPIAR_LECTOR_HUELLA()
   'Desocupa campos
    lblCodigo.Caption = ""
    txtCedula.SetFocus
    CheckPeople1.Limpiar
End Sub

Public Sub LEER_DATOS_HUELLA()
   'Valida que sea un dato numerico
   If Not IsNumeric(txtCedula.text) Then txtCedula.text = ""

   If Not Trim(txtCedula.text) = "" Then
        
        QRY = "SELECT * " & _
              "FROM coconductor  " & _
              "WHERE cocedula = '" & txtCedula.text & "' "
        
        Set Tbl20 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
        'Set Tbl20 = dbs4.OpenRecordset(qry, dbOpenSnapshot, dbReadOnly)
            
        If Not Tbl20.EOF Then
            If Respuesta_huella = vbYes Then
               l = LenB(Tbl20.Fields("cohuella"))
               If l > 0 Then
                   lblCodigo = Tbl20.Fields("codedo")
                   HuellaAlc() = Tbl20.Fields("cohuella").GetChunk(0, l)
               End If
            ElseIf Respuesta_huella = vbNo Then
               l = LenB(Tbl20.Fields("cohuealt"))
               If l > 0 Then
                   lblCodigo = Tbl20.Fields("codedalt")
                   HuellaAlc() = Tbl20.Fields("cohuealt").GetChunk(0, l)
               End If
            End If
        End If
        Tbl20.Close
    End If
End Sub

Public Sub CAPTURAR_HUELLA()
   'Captura la huella
   CheckPeople1.Registrar
   lblCodigo.Caption = CheckPeople1.Dedo
   'lblNombreDedo.Caption = CheckPeople1.NombreDedo
      
   'Inserta el registro con la huella
   If Not IsEmpty(CheckPeople1.HuellaByte) And CheckPeople1.Dedo > 0 Then
         
         errorHuellero = False
         
         QRY = " SELECT * FROM COCONDUCTOR " & _
               " WHERE COCEDULA = '" & txtCedula.text & "' "
         
         Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
         
         If Not Tbl21.EOF Then
            
            Tbl21.Close
            If Respuesta_huella = vbYes Then
               QRY = "UPDATE coconductor " _
                         & "SET CODEDO = " & lblCodigo.Caption & " " _
                       & "WHERE cocedula = '" & txtCedula.text & "' "
            ElseIf Respuesta_huella = vbNo Then
               QRY = "UPDATE coconductor " _
                         & "SET CODEDALT = " & lblCodigo.Caption & " " _
                       & "WHERE cocedula = '" & txtCedula.text & "' "
            End If
            
            dbs4.Execute QRY
            If Respuesta_huella = vbYes Then
               QRY = " SELECT cohuella " & _
                    " FROM coconductor " & _
                   " WHERE cocedula = '" & txtCedula.text & "' "
            ElseIf Respuesta_huella = vbNo Then
               QRY = " SELECT cohuealt " & _
                    " FROM coconductor " & _
                   " WHERE cocedula = '" & txtCedula.text & "' "
            End If
            
            Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
            
            Tbl21.Edit
            Tbl21(0).AppendChunk (CheckPeople1.HuellaByte)
            Tbl21.Update
         End If
         Tbl21.Close
   Else
        errorHuellero = True
   End If
    
End Sub
Public Sub VERIFICAR_HUELLA1()
   LEER_DATOS_HUELLA1
   
   If CheckPeople1.Verificar(HuellaAlc1(), lblCodigo.Caption) Then
      RegistoHuella1 = True
   Else
      RegistoHuella1 = False
   End If
   
   FINALIZAR_LECTOR_HUELLA
   
End Sub

Public Sub LIMPIAR_LECTOR_HUELLA1()
   'Desocupa campos
    lblCodigo.Caption = ""
    TxtCedula1.SetFocus
    CheckPeople1.Limpiar
End Sub

Public Sub LEER_DATOS_HUELLA1()
   'Valida que sea un dato numerico
   If Not IsNumeric(TxtCedula1.text) Then TxtCedula1.text = ""

   If Not Trim(TxtCedula1.text) = "" Then
        
        QRY = "SELECT * " & _
              "FROM coconductor  " & _
              "WHERE cocedula = '" & TxtCedula1.text & "' "
        
        Set Tbl20 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
        'Set Tbl20 = dbs4.OpenRecordset(qry, dbOpenSnapshot, dbReadOnly)
            
        If Not Tbl20.EOF Then
            If Respuesta_huella1 = vbYes Then
               l = LenB(Tbl20.Fields("cohuella"))
               If l > 0 Then
                   lblCodigo = Tbl20.Fields("codedo")
                   HuellaAlc1() = Tbl20.Fields("cohuella").GetChunk(0, l)
               End If
            ElseIf Respuesta_huella1 = vbNo Then
               l = LenB(Tbl20.Fields("cohuealt"))
               If l > 0 Then
                   lblCodigo = Tbl20.Fields("codedalt")
                   HuellaAlc1() = Tbl20.Fields("cohuealt").GetChunk(0, l)
               End If
            End If
        End If
        Tbl20.Close
    End If
End Sub

Public Sub CAPTURAR_HUELLA1()
   'Captura la huella
   CheckPeople1.Registrar
   lblCodigo.Caption = CheckPeople1.Dedo
   'lblNombreDedo.Caption = CheckPeople1.NombreDedo
      
   'Inserta el registro con la huella
   If Not IsEmpty(CheckPeople1.HuellaByte) And CheckPeople1.Dedo > 0 Then
         
         errorHuellero1 = False
         
         QRY = " SELECT * FROM COCONDUCTOR " & _
               " WHERE COCEDULA = '" & TxtCedula1.text & "' "
         
         Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
         
         If Not Tbl21.EOF Then
            
            Tbl21.Close
            If Respuesta_huella1 = vbYes Then
               QRY = "UPDATE coconductor " _
                         & "SET CODEDO = " & lblCodigo.Caption & " " _
                       & "WHERE cocedula = '" & TxtCedula1.text & "' "
            ElseIf Respuesta_huella1 = vbNo Then
               QRY = "UPDATE coconductor " _
                         & "SET CODEDALT = " & lblCodigo.Caption & " " _
                       & "WHERE cocedula = '" & TxtCedula1.text & "' "
            End If
            
            dbs4.Execute QRY
            
            If Respuesta_huella1 = vbYes Then
               QRY = " SELECT cohuella " & _
                    " FROM coconductor " & _
                   " WHERE cocedula = '" & TxtCedula1.text & "' "
            ElseIf Respuesta_huella1 = vbNo Then
               QRY = " SELECT cohuealt " & _
                    " FROM coconductor " & _
                   " WHERE cocedula = '" & TxtCedula1.text & "' "
            End If
            
            Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
            
            Tbl21.Edit
            Tbl21(0).AppendChunk (CheckPeople1.HuellaByte)
            Tbl21.Update
         End If
         Tbl21.Close
   Else
        errorHuellero1 = True
   End If
    
End Sub

Private Function Valida_Ruta_Transito() As Boolean
    Dim sFechaActual As String
    Dim sFechaVenta  As String

    'Consultamos la fecha y hora actual
    QRY = "SELECT * FROM " _
                & "geparsis " _
                & "WHERE pscod = 'HORASIS' "
    
    Set Tbl4 = Dbs.Execute(QRY)
    
    If Not Tbl4.EOF Then
        sFechaActual = Tbl4!psval
    End If
                
    'Consultamos la última tasa de uso en transito vendida para el vehículo
    QRY = "SELECT * FROM (" _
                & "SELECT coplaca, TO_CHAR(cofecsal, 'MM/DD/YYYY') AS cofecsal, TO_CHAR(cohorven, 'HH24:MI') AS cohorven " _
                            & "FROM coconduc " _
                            & "WHERE coplaca = '" & Txtnumpla.text & "' " _
                            & "AND cooritra = 'T' " _
                            & "AND coestado = 'Activo' " _
                            & "AND cofecsal >= TO_DATE('" & Mskfecha.text & "', 'MM/DD/YYYY') " _
                            & "ORDER BY cofecsal DESC, cohorven DESC " _
            & ") WHERE ROWNUM <= 1 "
            
    Set Tbl4 = Dbs.Execute(QRY)
    
    If Not Tbl4.EOF Then
        sFechaActual = Format(sFechaActual, "MM/dd/yyyy HH:mm")
        sFechaVenta = CDate(Tbl4!cofecsal & " " & Tbl4!cohorven)
        sFechaVenta = Format(sFechaVenta, "MM/dd/yyyy HH:mm")
        
        'Miramos la diferencia en horas
        sDiferenciaHoras = DateDiff("h", sFechaActual, sFechaVenta)
        
        If sDiferenciaHoras < 0 Then
            sDiferenciaHoras = DateDiff("h", sFechaVenta, sFechaActual)
        End If
        
        If sDiferenciaHoras <= 0 Then
            Valida_Ruta_Transito = False
        Else
            Valida_Ruta_Transito = True
        End If
    Else
        'Si no trae datos la consulta, se entiende que no ha comprado tasa de uso para el presente día
        Valida_Ruta_Transito = True
    End If
    
    
End Function



Function imprimirMulta(ByVal fechaMulta As String, ByVal codinf As String, ByVal numeroMulta As String, ByVal valorMulta As String, ByVal fechaVenta As String, ByVal horven As String) As Boolean
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
   

  
         fqry05 = "SELECT * " _
                     & "FROM geparsis " _
                    & "WHERE pscod = '" & codinf & "' "
         
         Set Ftb07 = Dbs.Execute(fqry05)
      
         If Not Ftb07.EOF Then
            descodinf = Ftb07!psdes
        End If
  
     
       Ftb07.Close
       Set Ftb07 = Nothing 'para la liberacion de recursos y evitar errores en el ejecutable (se bloquea)
     
      'Camilo Campos 07/04/2025
      'Ahora vamos a hacer la consulta para traer la INFO faltante de la multa
        fqry05 = "SELECT mudestino, rudestino, muplanilla " _
                     & "FROM opmultas JOIN opruta on mudestino = rucodigo " _
                    & "WHERE munumero = '" & numeroMulta & "' "
         
         Set Ftb07 = Dbs.Execute(fqry05)
      
         If Not Ftb07.EOF Then
                sDestinoTUT = Ftb07!rudestino 'trae el nombre de la ruta
                sDestinoPLA = Ftb07!muplanilla 'numero de planilla con la que quedo la multa
                sDestinoTIQ = Ftb07!mudestino 'código de la ruta
        Else
        'Como contigencia en caso de fallar la consulta de arriba, dejamos la INFO de la TU actual
                sDestinoTUT = Lblnomrut.Caption
                sDestinoPLA = TxtNumAlistamiento.text
                sDestinoTIQ = Txtcodrut.text
        End If
   
   
        
      sDescuento = "0"
      
      '-----------------------------------------------------------------------'
      '  CARGA DE LAS VARIABLES PARA LA IMRESION DEL TIQUETE                  '
      '-----------------------------------------------------------------------'
      sComparendo = numeroMulta
      sFecha = Format(fechaVenta, "MMM/dd/yyyy")
      sHora = Format(horven, "HH:MM")
      SInfraccion = descodinf
      sNumeroInterno = TxtNumInt.text
      sPlaca = Txtnumpla.text
      sNotificado = "PENDIENTE"
      sFechaInf = Format(fechaMulta, "MMM/dd/yyyy")
      sCajero = Login
      sCaja = Pc
      sturno = G_Turno
      sNivelServicio = Lbltipveh.Caption
      sValorMulta = valorMulta
      sDescuento = "0"
      sValorTotal = valorMulta
      sCodBarras = "*" & numeroMulta & "*"
      
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
         
   
      End If
      If G_ImpCon = "C" Or G_ImpCon = "N" Then


         fqry05 = "SELECT dsdes,dsval " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & codinf & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'INFRACCIONES PARA COBRAR'  "
               
         Set Ftb07 = Dbs.Execute(fqry05)
     
         If Not Ftb07.EOF Then
            sTipoInfraccion = Ftb07!dsval
         End If
         
      
         
         If sTipoInfraccion = 1 Then
            'multas de las 51 a la 54
            
            If G_ImpEpsonZebra = "P" Then
               reporte = RpPath + "\recmulta.rpt"
            ElseIf G_ImpEpsonZebra = "Z" Then
               reporte = RpPath + "\recmultazebra.rpt"
            End If
            
            
            rptMultas.ReportFileName = reporte
            rptMultas.Connect = G_Conexion_reportes
            
            
               rptMultas.Formulas(0) = "nombreempresa = '" & G_NomEmpresa & "'"
               rptMultas.Formulas(1) = "numcomparendo = '" & sComparendo & "'"
               rptMultas.Formulas(2) = "hora = '" & sHora & "'"
               rptMultas.Formulas(3) = "numint = '" & sNumeroInterno & "'"
               rptMultas.Formulas(4) = "numplaca = '" & sPlaca & "'"
               rptMultas.Formulas(5) = "infraccion = '" & SInfraccion & "'"
               rptMultas.Formulas(6) = "notificado = '" & sNotificado & "'"
               rptMultas.Formulas(7) = "fechainfraccion = '" & sFechaInf & "'"
               rptMultas.Formulas(8) = "destinotut = '" & sDestinoTUT & "'"
               rptMultas.Formulas(9) = "destinoplanilla = '" & sDestinoPLA & "'"
               rptMultas.Formulas(10) = "destinotiquete = '" & sDestinoTIQ & "'"
               rptMultas.Formulas(11) = "usuario = '" & sCajero & "'"
               rptMultas.Formulas(12) = "pc = '" & sCaja & "'"
               rptMultas.Formulas(13) = "turno = '" & sturno & "'"
               rptMultas.Formulas(14) = "servicio = '" & sNivelServicio & "'"
               rptMultas.Formulas(15) = "valmulta = '" & Format(sValorMulta, "$##,##0") & "'"
               rptMultas.Formulas(16) = "descuento = '" & Format(sDescuento, "$##,##0") & "'"
               rptMultas.Formulas(17) = "totalpagar = '" & Format(sValorTotal, "$##,##0") & "'"
               rptMultas.Formulas(18) = "codigobarras = '" & sCodBarras & "'"
               rptMultas.Formulas(19) = "fecha = '" & sFecha & "'"
               If G_ImpEpsonZebra = "P" Then
                  rptMultas.PrinterName = G_PrinterName1
                  rptMultas.PrinterDriver = G_PrinterDriver1
                  rptMultas.PrinterPort = G_PrinterPort1
               End If
               
               rptMultas.Destination = 1
               rptMultas.Action = 1
        
         End If
         
      Else
        MsgBox "Verifique el tipo de impresión que tiene definida", 48, "Advertencia"
        'Si entra aqui significa que tienen mal definida la impresión
        imprimirMulta = False
      
        Exit Function
      End If
    
 
   
   imprimirMulta = True
    GoTo Finalizar  'queda de esta forma debido a que en el ejecutable se esta colapsando si realiza más de 2 impresiones de una multa/parqueadero
   
Errores:
   MsgBox "No se ha podido generar la impresión, vuelva a intentar " & Err.Description, 16, "Error " & Err.Number
   imprimirMulta = False  'MUY IMPORTANTE definimos esto como una función para poder validar en el metodo principal si lanza error
   Exit Function
   
   
Finalizar:
  On Error Resume Next
    ' Liberar siempre
    If Not Ftb07 Is Nothing Then
        If Ftb07.State <> 0 Then Ftb07.Close
        Set Ftb07 = Nothing
    End If
    Close #1
    Set rptMultas = Nothing
    On Error GoTo 0
    Exit Function
End Function




Function ReImprimirPago(ByVal sTiempoPermanencia As String, ByVal sNombreTerminal As String, ByVal sNit As String, ByVal sFecha As String, _
                           ByVal sHora As String, ByVal sNumeroInterno As String, ByVal sPlaca As String, ByVal sEmpresa As String, _
                           ByVal sFechaIngreso As String, ByVal sFechaSalida As String, ByVal sValorPagar As String, ByVal sDescuento As String, _
                           ByVal sValor As String, ByVal sCajero As String, ByVal sCaja As String, ByVal sturno As String, _
                           ByVal sNumero As String, ByVal sNumeroBarras As String, ByVal snombrefechaingreso As String, ByVal sTitulo As String, _
                           ByVal sTiempoVigencia As String) As Boolean
   On Error GoTo Errores
   
   
   On Error Resume Next
   Close #1
   On Error GoTo Errores
   
   If G_ImpCon = "Z" Then
      
      Open "LPT1" For Output As #1

      Print #1, "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ"
      Print #1, "~DG000.GRF,02048,016,"
      Print #1, ",:::::::::::::::::::::::::T010H010,,U07FHF0,T02F0H0FC0,S01F10H0171,S0280K08,R0140L03,R0680M080,Q01D0N070,Q0180N028,Q060O01C,Q0E0P0E,P0180P03,P020Q0180,P040R040,P0C0R0A0,N0H180R030,O020S018,O060T08,O040T0C,O0C0T04,O080T02,N0180T03,N030U01,N030U0180,N020V080,N060V040,N080EE80M0EA0H080,N0C0FFC0L01FF0H040,N080FB80L01BA0H020,M0180FFC0L01FF0H020,N080EE80M0HEI020,M0100FFC010H0101FF01030,P0HB80L01FE0H020,M03007FC0L01FF0H030,M0200EAERE0,M0300FTF010,M02002BFBFBFBFBFBFBE010,M03007FSF010,M02002AAEQEH08,M07003FSF018,M02001ABBFOFB018,M060H0TF018,M020H0HAREH08,M06013FSF018,M02003FFBFBFBFBFBFBB018,M07007FSF018,M0200ESEA0,L01300FTF010,M0200FFBFIFBFKFA010,M0301FFC7FPF010,M0200EE80M0EA0H020,M0101FF80L01FF0H020,M0100FB80L01FA0H020,M0181FF0M01FE0H020,M0181AE80M0HEI020,M0181FF80L01FF0H040,N080BA80L01FE0,N0C0V040,N0E0V080,M0160U0180,N020U01,N070U01,N010U02,N0180T04,O08,O040T0C,O060T08,O070S010,O0380R020,O01C0R040,P0C0R080,P070Q01,P0380P02,P01C0P04,Q060P08,Q030O030,Q0180N0E0,R070M03," _
               & "R0180L0E,R01D0K05810,S03A0J0A0,T0H7H017,U0BFEA0,W"
      Print #1, "01,,::::::::::::::::~DG001.GRF,02048,016,"
      Print #1, ",:::::::::::::::::::::::::R010H010,,S07FHF0,R02F0H0FC0,Q01F10H0171,Q0280K08,P0140L03,P0680M080,O01D0N070,O0180N028,O060O01C,O0E0P0E,N0180P03,N020Q0180,N040R040,N0C0R0A0,L0H180R030,M020S018,M060T08,M040T0C,M0C0T04,M080T02,L0180T03,L030U01,L030U0180,L020V080,L060V040,L080EE80M0EA0H080,L0C0FFC0L01FF0H040,L080FB80L01BA0H020,K0180FFC0L01FF0H020,L080EE80M0HEI020,K0100FFC010H0101FF01030,N0HB80L01FE0H020,K03007FC0L01FF0H030,K0200EAERE0,K0300FTF010,K02002BFBFBFBFBFBFBE010,K03007FSF010,K02002AAEQEH08,K07003FSF018,K02001ABBFOFB018,K060H0TF018,K020H0HAREH08,K06013FSF018,K02003FFBFBFBFBFBFBB018,K07007FSF018,K0200ESEA0,J01300FTF010,K0200FFBFIFBFKFA010,K0301FFC7FPF010,K0200EE80M0EA0H020,K0101FF80L01FF0H020,K0100FB80L01FA0H020,K0181FF0M01FE0H020,K0181AE80M0HEI020,K0181FF80L01FF0H040,L080BA80L01FE0,L0C0V040,L0E0V080,K0160U0180,L020U01,L070U01,L010U02,L0180T04,M08,M040T0C,M060T08,M070S010,M0380R020,M01C0R040,N0C0R080,N070Q01,N0380P02,N01C0P04,O060P08,O030O030,O0180N0E0,P070M03,P0180L0E," _
               & "P01D0K05810,Q03A0J0A0,R0H7H017,S0BFEA0,U"
      Print #1, "01,,::::::::::::::::^XA^LL0743"
      Print #1, "^PW815"
      Print #1, "^FT288,128^XG000.GRF,1,1^FS"
      Print #1, "^FT704,128^XG001.GRF,1,1^FS"
      Print #1, "^BY2,2.5,47^FT45,220^B3R,N,,Y,N"
      Print #1, "^FD" & sNumeroBarras & "^FS"
      Print #1, "^FT75,613^A0R,20,21^FH\^FD" & sDescuento & "^FS"
      Print #1, "^FT104,613^A0R,20,21^FH\^FD" & sValor & "^FS"
      Print #1, "^FT146,577^A0R,20,21^FH\^FD" & sturno & "^FS"
      Print #1, "^FT340,222^A0R,23,21^FH\^FD" & sNombreTerminal & "^FS"
      Print #1, "^FT356,526^A0R,14,12^FH\^FDNumero^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT181,29^A0R,20,19^FH\^FD" & snombrefechaingreso & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT181,29^A0R,20,19^FH\^FDFECHA LIMITE:^FS"
      'End If
      
      Print #1, "^FT287,272^A0R,20,16^FH\^FDHORA^FS"
      Print #1, "^FT214,232^A0R,20,19^FH\^FD" & sEmpresa & "^FS"
      Print #1, "^FT213,28^A0R,20,19^FH\^FDEMPRESA:^FS"
      Print #1, "^FT180,514^A0R,20,19^FH\^FDCAJA^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT118,28^A0R,20,19^FH\^FD" & sTiempoVigencia & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT118,28^A0R,20,19^FH\^FDVIGENCIA:^FS"
      'End If
      Print #1, "^FT150,29^A0R,20,19^FH\^FDFECHA SALIDA:^FS"
      Print #1, "^FT117,234^A0R,20,19^FH\^FD" & sTiempoPermanencia & "^FS"
      Print #1, "^FT290,538^A0R,20,16^FH\^FDPLACA^FS"
      Print #1, "^FT263,406^A0R,20,21^FH\^FD" & sNumeroInterno & "^FS"
      Print #1, "^FT287,406^A0R,20,16^FH\^FDNo INT.^FS"
      Print #1, "^FT263,271^A0R,20,21^FH\^FD" & sHora & "^FS"
      Print #1, "^FT150,234^A0R,20,21^FH\^FD" & sFechaSalida & "^FS"
      Print #1, "^FT40,615^A0R,25,24^FH\^FD" & sValorPagar & "^FS"
      Print #1, "^FT183,233^A0R,20,21^FH\^FD" & sFechaIngreso & "^FS"
      Print #1, "^FT77,493^A0R,20,19^FH\^FDDescuento^FS"
      Print #1, "^FT105,493^A0R,20,19^FH\^FDValor:^FS"
      Print #1, "^FT40,495^A0R,25,24^FH\^FDValor Pagar^FS"
      Print #1, "^FT147,514^A0R,20,19^FH\^FDTURNO^FS"
      Print #1, "^FT211,576^A0R,20,21^FH\^FD" & sCajero & "^FS"
      Print #1, "^FT179,577^A0R,20,21^FH\^FD" & sCaja & "^FS"
      Print #1, "^FT213,513^A0R,20,19^FH\^FDCAJERO^FS"
      Print #1, "^FT265,539^A0R,20,21^FH\^FD" & sPlaca & "^FS"
      Print #1, "^FT289,93^A0R,20,16^FH\^FDFECHA PAGO^FS"
      Print #1, "^FT262,92^A0R,20,21^FH\^FD" & sFecha & "^FS"
      Print #1, "^FT320,281^A0R,17,16^FH\^FD" & sNit & "^FS"
      Print #1, "^FT343,596^A0R,34,33^FH\^FD" & sNumero & "^FS"
      Print #1, "^FT367,247^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
      Print #1, "^BY2,2.5,47^FT453,220^B3R,N,,Y,N"
      Print #1, "^FD" & sNumeroBarras & "^FS"
      Print #1, "^FT483,613^A0R,20,21^FH\^FD" & sDescuento & "^FS"
      Print #1, "^FT512,613^A0R,20,21^FH\^FD" & sValor & "^FS"
      Print #1, "^FT554,577^A0R,20,21^FH\^FD" & sturno & "^FS"
      Print #1, "^FT748,222^A0R,23,21^FH\^FD" & sNombreTerminal & "^FS"
      Print #1, "^FT764,526^A0R,14,12^FH\^FDNumero^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT589,29^A0R,20,19^FH\^FD" & snombrefechaingreso & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT589,29^A0R,20,19^FH\^FDFECHA LIMITE:^FS"
      'End If
      
      Print #1, "^FT695,272^A0R,20,16^FH\^FDHORA^FS"
      Print #1, "^FT622,232^A0R,20,19^FH\^FD" & sEmpresa & "^FS"
      Print #1, "^FT621,28^A0R,20,19^FH\^FDEMPRESA:^FS"
      Print #1, "^FT588,514^A0R,20,19^FH\^FDCAJA^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT526,28^A0R,20,19^FH\^FD" & sTiempoVigencia & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT526,28^A0R,20,19^FH\^FDVIGENCIA:^FS"
      'End If
      
      Print #1, "^FT558,29^A0R,20,19^FH\^FDFECHA SALIDA:^FS"
      Print #1, "^FT525,234^A0R,20,19^FH\^FD" & sTiempoPermanencia & "^FS"
      Print #1, "^FT698,538^A0R,20,16^FH\^FDPLACA^FS"
      Print #1, "^FT671,406^A0R,20,21^FH\^FD" & sNumeroInterno & "^FS"
      Print #1, "^FT695,406^A0R,20,16^FH\^FDNo INT.^FS"
      Print #1, "^FT671,271^A0R,20,21^FH\^FD" & sHora & "^FS"
      Print #1, "^FT558,234^A0R,20,21^FH\^FD" & sFechaSalida & "^FS"
      Print #1, "^FT448,615^A0R,25,24^FH\^FD" & sValorPagar & "^FS"
      Print #1, "^FT591,233^A0R,20,21^FH\^FD" & sFechaIngreso & "^FS"
      Print #1, "^FT485,493^A0R,20,19^FH\^FDDescuento^FS"
      Print #1, "^FT513,493^A0R,20,19^FH\^FDValor:^FS"
      Print #1, "^FT448,495^A0R,25,24^FH\^FDValor Pagar^FS"
      Print #1, "^FT555,514^A0R,20,19^FH\^FDTURNO^FS"
      Print #1, "^FT619,576^A0R,20,21^FH\^FD" & sCajero & "^FS"
      Print #1, "^FT587,577^A0R,20,21^FH\^FD" & sCaja & "^FS"
      Print #1, "^FT621,513^A0R,20,19^FH\^FDCAJERO^FS"
      Print #1, "^FT673,539^A0R,20,21^FH\^FD" & sPlaca & "^FS"
      Print #1, "^FT697,93^A0R,20,16^FH\^FDFECHA PAGO^FS"
      Print #1, "^FT670,92^A0R,20,21^FH\^FD" & sFecha & "^FS"
      Print #1, "^FT728,281^A0R,17,16^FH\^FD" & sNit & "^FS"
      Print #1, "^FT751,596^A0R,34,33^FH\^FD" & sNumero & "^FS"
      Print #1, "^FT775,247^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
      Print #1, "^PQ1,0,1,Y^XZ"
      Print #1, "^XA^ID000.GRF^FS^XZ"
      Print #1, "^XA^ID001.GRF^FS^XZ"
      
      Close #1
      
   ElseIf G_ImpCon = "C" Or G_ImpCon = "N" Then

      If G_ImpEpsonZebra = "P" Then
      
         reporte = RpPath + "\pagparqueadero.rpt"
         
      ElseIf G_ImpEpsonZebra = "Z" Then
         
         reporte = RpPath + "\pagparqueaderozebra.rpt"
        
      End If
      rptPar.ReportFileName = reporte
      rptPar.Connect = G_Conexion_reportes
    
         rptPar.Formulas(0) = "nombreempresa = '" & nombrebasedatos & "'"                       'ok
         rptPar.Formulas(1) = "numrecibo = '" & sNumero & "'"                                   'ok
         rptPar.Formulas(2) = "fechapago = '" & sFecha & "'"                                    'ok
         rptPar.Formulas(3) = "hora = '" & sHora & "'"                                          'ok
         rptPar.Formulas(4) = "numint = '" & sNumeroInterno & "'"                               'ok
         rptPar.Formulas(5) = "numplaca = '" & sPlaca & "'"                                     'ok
         rptPar.Formulas(6) = "empveh = '" & sEmpresa & "'"                                     'ok
         rptPar.Formulas(7) = "fechaingreso = '" & sFechaIngreso & "'"                          'ok
         rptPar.Formulas(8) = "fechasalida = '" & sFechaSalida & "'"                            'ok
         rptPar.Formulas(9) = "tiempo = '" & sTiempoPermanencia & "'"                           'ok
         rptPar.Formulas(10) = "usuario = '" & sCajero & "'"                                    'ok
         rptPar.Formulas(11) = "pc = '" & sCaja & "'"                                           'ok
         rptPar.Formulas(12) = "turno = '" & sturno & "'"                                       'ok
         rptPar.Formulas(13) = "valor = '" & Format(sValor, "$##,##0") & "'"                    'ok
         rptPar.Formulas(14) = "descuento = '" & Format(sDescuento, "$##,##0") & "'"            'ok
         rptPar.Formulas(15) = "valortotal = '" & Format(sValorPagar, "$##,##0") & "'"          'ok
         rptPar.Formulas(16) = "codigobarras = '" & sNumeroBarras & "'"                         'ok
         rptPar.Formulas(17) = "snombrefechaingreso = '" & snombrefechaingreso & "'"            'ok
         rptPar.Formulas(18) = "stitulo = '" & sTitulo & "'"
         rptPar.Formulas(19) = "stiempovigencia = '" & sTiempoVigencia & "'"                    'ok
         
         If G_ImpEpsonZebra = "P" Then
            rptPar.PrinterName = G_PrinterName1
            rptPar.PrinterDriver = G_PrinterDriver1
            rptPar.PrinterPort = G_PrinterPort1
         End If
         
         rptPar.Destination = 1
         rptPar.Action = 1
         
   Else
        'si entra aqui significa que no tiene bien definido el tipo de impresora
        MsgBox "Verifique el tipo de impresión que tiene definida", 48, "Advertencia"
        ReImprimirPago = False
        Exit Function
   End If
   
   
    ReImprimirPago = True
    GoTo Finalizar  'MUY IMPORTANTE, de esta forma no se traba el ejecutable al salir del formulario de capconduces
   
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   ReImprimirPago = False
   Exit Function
   
Finalizar:
    On Error Resume Next
    Close #1
    Set rptPar = Nothing
    On Error GoTo 0
    Exit Function
End Function


