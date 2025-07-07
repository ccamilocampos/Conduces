VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Anucondu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Tasa de Uso"
   ClientHeight    =   6510
   ClientLeft      =   2355
   ClientTop       =   2070
   ClientWidth     =   7440
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcaja 
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtObserv 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Observaciones"
      Top             =   6120
      Width           =   6135
   End
   Begin VB.TextBox Txtestado 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Txtoperario 
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Operario que vendio el conduce"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Txtturno 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Turno en que se vendio el conduce"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txthora 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Hora de salida del vehículo"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txtfecha 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Fecha de salida del vehículo"
      Top             =   3000
      Width           =   1215
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
      Picture         =   "Anucondu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir de la forma actual"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Btnanular 
      Caption         =   "&Anular"
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
      Picture         =   "Anucondu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Anular tasa uso"
      Top             =   1680
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
      Picture         =   "Anucondu.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   4200
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
      Left            =   6480
      Picture         =   "Anucondu.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir recibo de anulación de Conduce"
      Top             =   3360
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
      Picture         =   "Anucondu.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar información de conduces ingresados"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Vehículo"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   6135
      Begin VB.TextBox Txtcodemp 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de transporte"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtnumpla 
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
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Placa del vehículo"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txtnumint 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Número interno del vehículo"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txttipveh 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Clase de servicio del vehículo"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Txtconact 
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
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Nombre del conductor actual"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Lblnomemp 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         Top             =   1080
         Width           =   3255
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   18
         Top             =   360
         Width           =   855
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
         Left            =   960
         TabIndex        =   17
         Top             =   360
         Width           =   735
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
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Conductor"
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
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lbltipveh 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.TextBox Txtnumcon 
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
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Digite el número del conduce a anular"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Txtcodrut 
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Código de la ruta"
      Top             =   2520
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1680
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin MSMask.MaskEdBox Mskvalor 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskalcoholemia 
      Height          =   375
      Left            =   5040
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Mskparqueadero 
      Height          =   375
      Left            =   5040
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Dinero de cambio"
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSMask.MaskEdBox MskValorTasaNorte 
      Height          =   375
      Left            =   1920
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   4560
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
   Begin MSMask.MaskEdBox MskValorMultas 
      Height          =   375
      Left            =   1920
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   5040
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
   Begin MSMask.MaskEdBox MskValorTag 
      Height          =   375
      Left            =   5040
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Valor conduce"
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label1 
      Caption         =   "TAG $"
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
      Left            =   3480
      TabIndex        =   47
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Multas $"
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
      TabIndex        =   46
      Top             =   5040
      Width           =   1455
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
      TabIndex        =   45
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "T. Permanencia$"
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
      Left            =   3480
      TabIndex        =   43
      Top             =   4560
      Width           =   1455
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
      Index           =   15
      Left            =   2880
      TabIndex        =   41
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Alcoholemia $"
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
      Left            =   3480
      TabIndex        =   39
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   37
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   6360
      Y1              =   3960
      Y2              =   3960
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
      Index           =   12
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Operario"
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
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Index           =   10
      Left            =   4440
      TabIndex        =   28
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   1335
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
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Tasa Uso$"
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
      TabIndex        =   23
      Top             =   4080
      Width           =   1455
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
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
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
      TabIndex        =   21
      Top             =   240
      Width           =   1335
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
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Lblnomrut 
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
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   6360
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   6480
   End
End
Attribute VB_Name = "Anucondu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LQryTransmitir As String
Dim L_Cupo As String
Dim L_Direccion As String
Dim L_Caja_Norte As String
Dim L_Numero_Terminal_Norte As Double
Dim L_Tipo_Cupo As String
Dim L_Cedula_Cupo As String
Dim L_Placa_Cupo As String
Dim L_Numero_Cupo As Double

Private Sub Traer_Tasas_Terminal_Norte()

   QRY = "SELECT tnnumero, tncaja, tnvalor " _
         & "FROM cotransitonorte " _
        & "WHERE tnnumerotasa = " & Txtnumcon.text & " " _
         & " AND tncajatasa = '" & txtcaja.text & "' "
          
   Set Tbl2 = Dbs.Execute(QRY)
   
   L_Caja_Norte = ""
   L_Numero_Terminal_Norte = 0
   MskValorTasaNorte.text = "0"
   If Not Tbl2.EOF Then
      L_Numero_Terminal_Norte = Tbl2!tnnumero
      L_Caja_Norte = Tbl2!tncaja
      MskValorTasaNorte.text = Tbl2!tnvalor
   End If
   Tbl2.Close
   
End Sub
   
Private Sub Traer_Datos_Cupo()
   QRY = "SELECT NVL(cptipo,'-') AS cptipo, NVL(cpcedula,'-') AS cpcedula, NVL(cpplaca  ,'-') AS cpplaca " _
         & "FROM cocupos " _
        & "WHERE cpnumcup = " & L_Numero_Cupo & " "
   
   Set Tbl2 = Dbs.Execute(QRY)
   
   
   If Not Tbl2.EOF Then
      L_Tipo_Cupo = Tbl2!cptipo
      L_Cedula_Cupo = Tbl2!cpcedula
      L_Placa_Cupo = Tbl2!cpplaca
   End If
   Tbl2.Close
End Sub


Private Sub Btnanular_Click()
   On Error GoTo Errores
   If Txtnumpla.text <> "" Then
      If Txtestado <> "Anulado" Then
      
         'Camilo Campos 07/08/2025 desde esta fecha solo se puede anular CON LA MISMA CAJA que se vendió
         If Not txtcaja.text = Pc Then
            MsgBox "La tasa de uso no pertenece a la caja actual", 48, "Advertencia"
            Exit Sub
         End If
         
         
         'qry = "SELECT MAX(lonumero) " _
               & "FROM gelogaud "
   
         'Set Ftb01 = Dbs.Execute(Qry)
   
         'If Not ftb01.EOF Then
         '   If IsNull(ftb01(0)) Then
         '      sNumero = 1
         '   Else
         '      sNumero = CDbl(ftb01(0)) + 1
         '   End If
         'Else
         '   sNumero = 1
         'End If
         'ftb01.Close
         
         respuesta = MsgBox("Desea ANULAR esta Tasa de Uso?", vbYesNo + 32, "ANULAR")
         If respuesta = vbYes Then
            estado = "Anulado"
            sTotal = CDbl(mskalcoholemia.text) + CDbl(Mskvalor.text)
            
            If L_Cupo = "P" Then
               QRY = "SELECT * " _
                     & "FROM cocupos " _
                    & "WHERE cpnumcup = " & L_Numero_Cupo & " " _
                      & "AND cpestado = 'A' "
      
               Set Tbl1 = Dbs.Execute(QRY)
      
               If Not Tbl1.EOF Then
                  sSaldoCupo = Tbl1!cpsaldo
               End If
            End If
            sFecha = Format(Fn_Parametros("HORASIS", 1), "MM/DD/YYYY HH:MM")
            Dbs.BeginTrans
            
            On Error GoTo Err_Transac
            Dbs.Execute "UPDATE coconduc " _
                      & "   SET coestado =  '" & estado & "'," _
                      & "       coobser =  '" & TxtObserv.text & "', " _
                      & "       copeaje = '0' " _
                      & " WHERE conumero = " & Txtnumcon.text & " " _
                              & "AND conomcaj = '" & Pc & "' "
            
            
            If L_Numero_Terminal_Norte <> 0 Then
               'se regresa al estado de P el cobro del transito
               QRY = "UPDATE cotransitonorte " _
                      & "SET TNNUMEROTASA = 0, " _
                         & " TNCAJATASA = NULL, " _
                         & " TNFECHATASA = NULL, " _
                         & " TNHORATASA = NULL, " _
                         & " TNTURNOTASA = 0, " _
                         & " TNCERRADOTASA = 'N', " _
                         & " TNFUNCIONARIOTASA = NULL, " _
                         & " TNESTADO = 'P' " _
                  & "WHERE tnnumero = " & L_Numero_Terminal_Norte & " " _
                   & " AND tncaja = '" & L_Caja_Norte & "' " _
                   & " AND tnnumerotasa = " & Txtnumcon.text & " " _
                   & " AND tncajatasa = '" & txtcaja.text & "'"
          
               Dbs.Execute QRY
               
            End If
            
            If L_Cupo = "P" Then
               sFecha1 = Format(sFecha, "MM/dd/yyyy")
               sHora = Format(sFecha, "HH:MM")
               
               QRY = "INSERT INTO opmovimientos (MOFECHA, MOHORA, MOTIPMOV, MOVALOR, MOSALDO, MOTIPO, MOTIPCUPO, MOTASAUSO, MOCAJA, MOCODEMP" _
            
               If L_Tipo_Cupo = "E" Then
                  QRY = QRY & " "
               ElseIf L_Tipo_Cupo = "P" Then
                  QRY = QRY & ", MOCEDULA, MOPLACA "
               End If
               
               QRY = QRY & ") VALUES(TO_DATE('" & sFecha1 & "','MM/DD/YYYY'),TO_DATE('" & sHora & "','hh24:mi'),'DB'," _
                        & sTotal & "," & sSaldoCupo & ",'ANTU','" & L_Tipo_Cupo & "', " & Txtnumcon.text & ", '" & txtcaja.text & "', '" & Txtcodemp.text & "' "
               
               If L_Tipo_Cupo = "E" Then
                  QRY = QRY & ""
               ElseIf L_Tipo_Cupo = "P" Then
                  QRY = QRY & ",'" & L_Cedula_Cupo & "','" & L_Placa_Cupo & "'"
               End If
               
               QRY = QRY & ")"
               
               Dbs.Execute QRY
               
               
               Dbs.Execute "UPDATE cocupos " _
                         & "   SET cpsaldo =  cpsaldo + " & sTotal & " " _
                         & " WHERE cpnumcup = " & L_Numero_Cupo & " " _
                         & "    AND cpestado = 'A'"
                      
            End If
            
            'HABILITAR EL VEHICULO DE LA TABLA DE CODESPACHOS
            QRY = "UPDATE codespachos " _
                   & "SET deestado = 'D', " _
                       & "denumtasa = 0,  " _
                       & "denomcaj = NULL " _
                & " WHERE denumtasa = " & Txtnumcon.text & " " _
                   & "AND denomcaj = '" & Pc & "' " _
                   & "AND deplaca = '" & Txtnumpla.text & "' " _
                   & "AND deestado = 'T' "
            
            Dbs.Execute QRY
            
            
            Dbs.Execute "DELETE FROM cosalveh " _
                            & "WHERE svnumero = " & Txtnumcon.text & " " _
                              & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "'"
            
            If G_Tipo_Terminal = "CENTRAL" Then
                 scero = 0
                 QRY = "UPDATE coparqueop " _
                        & "SET paaviso = paaviso - 1 " _
                      & "WHERE paplaca = '" & Txtnumpla.text & "' AND pavalor <> " & scero & " " _
                        & "AND paestado = 'A' "
                       
                 Dbs.Execute QRY
            End If
                     
            'Se arregla el pago de la multa
            'Buscamos el número de la multa primero
            QRY = "SELECT munumero " _
                    & "FROM opmultas " _
                    & "WHERE munumpla = '" & Txtnumpla.text & "' " _
                    & "AND mutasapag = " & Txtnumcon.text & " "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
            
                'Consultamos el próximo consecutivo de multas
                QRY = "SELECT psval FROM geparsis WHERE pscod = 'NUMMUL' "
                Set Ftb02 = Dbs.Execute(QRY)
                
                If Not Ftb02.EOF Then
                    sNumMulta = Ftb02!psval
                End If
            
                Do While Not Tbl1.EOF
                    'Teniendo el número de la multa, buscamos el total de abonos de la misma
                    QRY1 = "SELECT NVL(SUM(abvalor), 0) AS valmul FROM opabonos " _
                        & "WHERE abnumero = " & Tbl1!munumero & " " _
                        & "AND abestado = 'Activo' "
                        
                    Set Tbl2 = Dbs.Execute(QRY1)
                    
                    If Not Tbl2.EOF Then
                        'Si encontro abonos, pasamos estos a anulados, y la suma total
                        'la volvemos a dejar en el saldo de la multa
    '                    QRY = "UPDATE opmultas " _
    '                            & "SET muaviso = muaviso - 1, " _
    '                            & "musaldo = " & Tbl2!valmul & ", " _
    '                            & "mucerrado = 'N', " _
    '                            & "muciedia = 'N', " _
    '                            & "munomcaj = 'automatica', " _
    '                            & "mutasapag = 0, " _
    '                            & "muestado = 'Pendiente' " _
    '                            & "WHERE munumero = " & Tbl1!munumero & " " _
    '                            & "AND munumpla = '" & Txtnumpla.text & "' " _
    '                            & "AND mutasapag = " & Txtnumcon.text & " "
    '
    '                    Dbs.Execute (QRY)
                        
                        'Pasamos la multa como anulada
                        'Y creamos una nueva para que sea cobrada y guardar registro de la anulación
                        QRY = "UPDATE opmultas " _
                                & "SET muaviso = muaviso - 1, " _
                                & "musaldo = " & Tbl2!valmul & ", " _
                                & "munomcaj = 'automatica', " _
                                & "muestado = 'Anulado' " _
                                & "WHERE munumero = " & Tbl1!munumero & " " _
                                & "AND munumpla = '" & Txtnumpla.text & "' " _
                                & "AND mutasapag = " & Txtnumcon.text & " "
                                
                        Dbs.Execute (QRY)
                        
                        'Aumentamos en 1 el número de la multa
                        sNumMulta = sNumMulta + 1
                        
                        QRY = "INSERT INTO opmultas (munumero, mucodemp, mucodinf, munomnot, mufecha, munumpla, musitio, munomtes, muauxcon, " _
                                    & "mutipveh, muobserv, muestado, munetpag, muhora, mufecsis, musaldo, munomcaj, " _
                                    & "mudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete, mupasvia, " _
                                    & "mupaspla, mucedula, mutasapag) " _
                                    & "(SELECT " & sNumMulta & " AS munumero, mucodemp, mucodinf, munomnot, mufecha, munumpla, musitio, munomtes, muauxcon, " _
                                        & "mutipveh, muobserv, 'Pendiente' AS muestado, munetpag, muhora, mufecsis, musaldo, " _
                                        & "munomcaj, mudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete, mupasvia, " _
                                        & "mupaspla, mucedula, 0 AS mutasapag " _
                                        & "FROM opmultas " _
                                        & "WHERE munumero = " & Tbl1!munumero & " " _
                                        & "AND munumpla = '" & Txtnumpla.text & "' " _
                                        & "AND mutasapag = " & Txtnumcon.text & " " _
                                    & ") "
                        
                        Dbs.Execute (QRY)
                        
                        QRY = "UPDATE opabonos " _
                                & "SET abestado = 'Anulado' " _
                                & "WHERE abnumero = " & Tbl1!munumero & " "
                                
                        Dbs.Execute (QRY)
                        
                        'Sebastián Rondón - Mayo 08 de 2025
                        'Se elimina pago de la tabla temporal de pagos para que no se asigne a una tasa de uso corriente
                        QRY = "DELETE FROM opreccaja " _
                                    & "WHERE rcnumero = " & Tbl1!munumero & " "
                        
                        Dbs.Execute (QRY)
                    End If
                    Tbl1.MoveNext
                    Tbl2.Close
                Loop
                
                'Actualizamos el número de la multa luego de insertar las necesarias
                QRY = "UPDATE geparsis SET psval = " & sNumMulta & " WHERE pscod = 'NUMMUL' "
                
                Dbs.Execute (QRY)
                
            End If
            Tbl1.Close
            
            
            
'            QRY = "UPDATE opmultas " _
'                & "SET muaviso = muaviso - 1 " _
'              & "WHERE munumpla = '" & Txtnumpla.text & "' " _
'                & "AND muestado = 'Pendiente' " _
'
'            Dbs.Execute QRY
                     
                     
            'wecc 08/31/2010 se vuelve y se deja el registro del cobro de parqueo
            If MskParqueadero.text <> "" And MskParqueadero.text <> "0" Then
               QRY = "UPDATE coparqueop " _
                      & "SET paaviso = paaviso - 1 " _
                    & "WHERE paplaca = '" & Txtnumpla.text & "' " _
                      & "AND paestado = 'A' AND paaviso <> 0 " _
                      & "AND patasapag = " & Txtnumcon.text & " "
                  
               Dbs.Execute QRY
               
'               QRY = "UPDATE coparqueop " _
'                      & "SET pafecpag = NULL , " _
'                         & " patasa = 0, " _
'                         & " patasapag = 0, " _
'                         & " pacaja = NULL, " _
'                         & " paturno = 0, " _
'                         & " pacerrado= 'N', " _
'                         & " paestado = 'A' " _
'                    & "WHERE patasapag = " & Txtnumcon.text & " " _
'                      & "AND pacaja = '" & txtcaja.text & "' "
'
'               Dbs.Execute QRY

                'Sebastián Rondón - Julio 30 de 2024
                'Por orden del Ingeniero Álvaro, al anular una permanencia
                'Se debe guardar registro de la anulación, por eso se procede a crear una nueva permanencia en estado A
                
                'Buscamos todas las permanencias a reversar
                QRY = "SELECT panumero " _
                        & "FROM coparqueop " _
                        & "WHERE patasapag = " & Txtnumcon.text & " " _
                        & "AND pacaja = '" & txtcaja.text & "' "
                        
                Set Tbl1 = Dbs.Execute(QRY)
                
                If Not Tbl1.EOF Then
                    Do While Not Tbl1.EOF
                        'Actualizamos primero el pago a estado R (Reversado)
                        QRY = "UPDATE coparqueop " _
                                & "SET paestado = 'R' " _
                                & "WHERE panumero = " & Tbl1!panumero & " " _
                                & "AND patasapag = " & Txtnumcon.text & " " _
                                & "AND pacaja = '" & txtcaja.text & "' "
                                
                        Dbs.Execute (QRY)
                    
                        'Ahora creamos el registro nuevo, para que cobre al vender nuevamente
                        QRY = "INSERT INTO coparqueop (patasapag, paplaca, pacodemp, pafecini, pafecsal, patiempo, pavalor, patipo, paestado, " _
                                 & "patarifa, paaviso, paregpc, paobservacion, paterminal) " _
                                 & "(SELECT 0 AS patasapag, paplaca, pacodemp, pafecini, pafecsal, patiempo, pavalor, patipo, 'A' AS paestado, " _
                                    & "patarifa, paaviso, paregpc, paobservacion, paterminal " _
                                    & "FROM coparqueop " _
                                    & "WHERE panumero = " & Tbl1!panumero & " " _
                                    & "AND pacaja = '" & txtcaja.text & "' " _
                                 & ") "
                                 
                        'Sebastián Rondón - Mayo 08 de 2025
                        'Se elimina pago de la tabla temporal de pagos para que no se asigne a una tasa de uso corriente
                        QRY = "DELETE FROM opreccaja " _
                                    & "WHERE rcnumero = " & Tbl1!panumero & " "
                        
                        Dbs.Execute (QRY)
                                 
                        Dbs.Execute QRY
                        Tbl1.MoveNext
                    Loop
                End If
               
            End If
            
            'Se arregla el pago del TAG
'            QRY = "UPDATE optag " _
'                    & "SET tafecpag = NULL, " _
'                    & "tahorpag = NULL, " _
'                    & "taaviso = taaviso - 1, " _
'                    & "taturno = NULL, " _
'                    & "tanomcaj = NULL, " _
'                    & "tacerrado = NULL, " _
'                    & "taestado = 'A', " _
'                    & "tanumtasa = NULL " _
'                    & "WHERE tanumtasa = " & Txtnumcon.text & " " _
'                    & "AND taplaca = '" & Txtnumpla.text & "' "
'
'            Dbs.Execute QRY

            If MskValorTag.text <> "" And MskValorTag.text <> "0" Then
            
                QRY = "SELECT psval FROM geparsis WHERE pscod = 'NUMTAG' "
                Set Tbl1 = Dbs.Execute(QRY)
                
                If Not Tbl1.EOF Then
                    sNumTag = Tbl1!psval + 1
                End If
    
                QRY = "UPDATE optag " _
                        & "SET taestado = 'R' " _
                        & "WHERE tanumtasa = " & Txtnumcon.text & " " _
                        & "AND taplaca = '" & Txtnumpla.text & "' "
                        
                Dbs.Execute (QRY)
                
                QRY = "INSERT INTO optag (tanumero, taplaca, tacodemp, tafecreg, tahorreg, tacodtag, taconcep, taobserv, " _
                        & "tavalor, taestado, taaviso, taterminal, tanumtasa, tasubtot, taiva) " _
                        & "(SELECT " & sNumTag & " AS tanumero, taplaca, tacodemp, tafecreg, tahorreg, tacodtag, taconcep, taobserv, " _
                            & "tavalor, 'A' AS taestado, taaviso, taterminal, 0 AS tanumtasa, tasubtot, taiva " _
                            & "FROM optag " _
                            & "WHERE tanumtasa = " & Txtnumcon.text & " " _
                            & "AND taplaca = '" & Txtnumpla.text & "' " _
                        & ") "
                        
                Dbs.Execute (QRY)
                
                QRY = "UPDATE geparsis SET psval = '" & sNumTag & "' WHERE pscod = 'NUMTAG' "
                
                Dbs.Execute (QRY)
                
            End If
            
            
            'Si el código de la ruta es 000 se entiende que es Tasa Sin Viaje
            'Así que no es necesario envíar la anulación de la tasa a Integra
            If Txtcodrut.text <> "000" Then
                LQryTransmitir = "UPDATE coconduc " _
                           & "   SET coestado =  '" & estado & "' " _
                           & " WHERE conumero = " & Txtnumcon.text & " " _
                              & "AND conomcaj = '" & Pc & "' " _
                              & "AND coterminal = '" & G_Codigo_Terminal & "' "
                'SE ALMACENA EL REGSITRO EN UNA TABLA PARA LUEGO TRANSMITIR
                lQry = Replace(LQryTransmitir, "'", "<")
                QRY = "INSERT INTO optransmitir(trsql) VALUES('" & lQry & "')"
                Dbs.Execute QRY
            End If
            
            
            Txtestado.text = estado
            Btnanular.Enabled = False
            'sFecha = Format(Now, "mm/dd/yyyy hh:mm")
            Dbs.Execute " INSERT INTO gelogaud " _
               & "(lofecact, lofecnue, lofuncio, " _
               & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
               & "VALUES (to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),  to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),'" _
               & Login & "', to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi'),'SE ANULO UNA TASA DE USO', 'ANUTASA', '" _
               & Txtnumcon.text & "', '" & Mid(Trim(TxtObserv.text), 1, 50) & "')"
            Dbs.CommitTrans
            On Error GoTo Errores
            MsgBox "Tasa de Uso Anulada", 64, "Ok"
         Else
            Exit Sub
         End If
      Else
         MsgBox "La Tasa de Uso YA ESTA Anulada", 48, "Advertencia"
      End If
   End If
   
   Exit Sub
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
Err_Transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub


Private Sub BtnBuscar_Click()
  forma = "c1"
  forma1 = "c1"
  buscar.Show 1
End Sub

Private Sub Btnlimpiar_Click()
   Txtnumcon.text = ""
   Txtcodemp.text = ""
   Txtcodrut.text = ""
   Txtnumpla.text = ""
   Txtnumpla.Enabled = True
   TxtNumInt.text = ""
   Txttipveh.text = ""
   Txtconact.text = ""
   Lblnomemp.Caption = ""
   Lblnomrut.Caption = ""
   Lbltipveh.Caption = ""
   Txtfecha.text = ""
   Txthora.text = ""
   Txtturno.text = ""
   Txtoperario.text = ""
   Txtestado.text = ""
   Mskvalor.text = ""
   mskalcoholemia.text = ""
   MskParqueadero.text = ""
   MskValorTasaNorte.text = ""
   TxtObserv.text = ""
   Txtnumcon.SetFocus
   Btnanular.Enabled = True
   txtcaja.text = ""
   MskValorMultas.text = ""
   MskValorTag.text = ""
   L_Cupo = ""
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Anucondu")
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub Txtnumcon_KeyPress(KeyAscii As Integer)
If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumcon_LostFocus()
  If Txtnumcon.text <> "" Then
       fqry01 = "SELECT psval " _
                & "FROM geparsis " _
               & "WHERE pscod = 'HORASIS' "
          
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         Fecha = Mid(Ftb01!psval, 1, 10)
         Fecha1 = DateAdd("d", -1, Fecha)
      End If
      Ftb01.Close
     
     'Ahora se necesitan anular tasas de uso que vienen de rodamiento, estas manejan reglas diferentes así que entramos a validar eso
     
      L_caja = Fn_Parametros("NOMCAJTAQ", 1) 'En este parametro se define el nombre de la caja en rodamiento
      
      'verificamos que exista registro en dicha caja
      
      fqry01 = "SELECT * " _
              & "FROM coconduc " _
             & "WHERE conumero = " & Txtnumcon.text & " " _
              & " AND conomcaj = '" & L_caja & "' "
              
             
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
     'La tasa consultada es de rodamiento, ahora verificamos que no existe un cierre de turno
     
         QRY = "SELECT * FROM cocietur WHERE cifecha = TO_DATE('" & Ftb01!cofecsal & "','MM/DD/YYYY')  " _
                 & "  AND cinomcaj = '" & L_caja & "' "
   
   
            Set Ftb02 = Dbs.Execute(QRY)
            
            If Not Ftb02.EOF Then
                 MsgBox "La tasa ya tuvo cierre de turno, no se puede anular", 48, "Advertencia"
                Exit Sub
            End If
            
     Else
            'La tasa de uso es normal, por lo tanto tiene que cumplir las reglas que venian de siempre
           fqry01 = "SELECT * " _
                 & "FROM coconduc " _
                & "WHERE conumero = " & Txtnumcon.text & " " _
                  & "AND cofecsal BETWEEN to_date('" & Fecha1 & "','MM/DD/YYYY') AND to_date('" & Fecha & "','MM/DD/YYYY') " _
                 & " AND cocerrado = 'N' "
                 
                
        Set Ftb01 = Dbs.Execute(fqry01)
    
    
     End If
     
     
    
     If Not Ftb01.EOF Then
        Ftb01.MoveLast
        If Ftb01!conomcaj = Pc Then
            L_Cupo = Ftb01!comodo
            Txtestado.text = Ftb01!coestado
            Txtnumpla.text = Ftb01!coplaca
            TxtNumInt.text = Ftb01!conumint
            Txttipveh.text = Ftb01!coclase
            Txtconact.text = Ftb01!coconduc
            If L_Cupo = "P" Then
               L_Numero_Cupo = Ftb01!conumcupo
               Call Traer_Datos_Cupo
               
               If L_Tipo_Cupo = "E" Then
                  If Ftb01!cotipoconvenio <> "1" Then
                     Txtcodemp.text = Ftb01!cocodempcon
                  Else
                     Txtcodemp.text = Ftb01!cocodemp
                  End If
               Else
                  Txtcodemp.text = Ftb01!cocodemp
                  L_Cedula_Cupo = Ftb01!cocedula
                  L_Placa_Cupo = Ftb01!coplaca
               End If
            Else
               Txtcodemp.text = Ftb01!cocodemp
            End If
            Txtcodrut.text = Ftb01!cocodrut
            Txtfecha.text = Ftb01!cofecsal
            Txthora.text = Mid(Ftb01!cohorsal, 12, 13)
            Txtturno.text = Ftb01!coturno
            Mskvalor.text = Ftb01!covalor
            If IsNull(Ftb01!codirecc) Then
                L_Direccion = "N/A"
            Else
                L_Direccion = Ftb01!codirecc
            End If
            If IsNull(Ftb01!covalalc) Then
               mskalcoholemia.text = "0"
            Else
               mskalcoholemia.text = Ftb01!covalalc
            End If
            If IsNull(Ftb01!coparque) Or Ftb01!coparque = 0 Then
               QRY = "SELECT SUM(pavalor) AS vrPar " _
                        & "FROM coparqueop " _
                        & "WHERE patasapag = " & Txtnumcon.text & " "
               Set Tbl1 = Dbs.Execute(QRY)
               
               If Not Tbl1.EOF Then
                 If Not IsNull(Tbl1!vrPar) Then
                    MskParqueadero.text = Tbl1!vrPar
                 Else
                    MskParqueadero.text = "0"
                 End If
               Else
                 MskParqueadero.text = "0"
               End If
               
            Else
               MskParqueadero.text = Ftb01!coparque
            End If
            
            QRY = "SELECT NVL(SUM(munetpag), 0) AS valmul " _
                    & "FROM opmultas " _
                    & "WHERE mutasapag = " & Txtnumcon.text & " " _
                    & "AND munumpla = '" & Txtnumpla.text & "' " _
                    & "AND muestado = 'Pagada' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                MskValorMultas.text = Tbl1!valmul
            Else
                MskValorMultas.text = "0"
            End If
            
            Tbl1.Close
            
            QRY = "SELECT NVL(SUM(tavalor), 0) As valtag " _
                    & "FROM optag " _
                    & "WHERE tanumtasa = " & Txtnumcon.text & " " _
                    & "AND taplaca = '" & Txtnumpla.text & "' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                MskValorTag.text = Tbl1!valtag
            Else
                MskValorTag.text = "0"
            End If
            
            Tbl1.Close
            
            Txtoperario.text = Ftb01!cofuncio
            TxtObserv.text = Ftb01!coobser
            txtcaja.text = Ftb01!conomcaj
            Ftb01.Close
            
            fqry01 = "SELECT dsdes " _
                     & "FROM gesuptip, gedetsuptip " _
                    & "WHERE dscoddet = '" & Txttipveh.text & "' " _
                      & "AND dscodtip = stcodtip " _
                      & "AND stdes = 'TIPOS DE VEHICULO' "
                   
            Set Ftb01 = Dbs.Execute(fqry01)
            
            If Not Ftb01.EOF Then
               Lbltipveh.Caption = Ftb01!dsdes
            Else
               MsgBox "Tipo De Vehiculo No Existe", 48, "Advertencia"
            End If
            Ftb01.Close
            fqry01 = "SELECT * " _
                     & "FROM ophojemp " _
                    & "WHERE hecodigo = '" & Txtcodemp.text & "'"
                 
            Set Ftb01 = Dbs.Execute(fqry01)
            
            If Not Ftb01.EOF Then
               Lblnomemp.Caption = Ftb01!henombre
            Else
                fqry01 = "SELECT henombre " _
                         & "FROM ophojempcon " _
                        & "WHERE hecodigo = '" & Txtcodemp.text & "'"
             
                Set Ftb01 = Dbs.Execute(fqry01)
        
                If Not Ftb01.EOF Then
                   Lblnomemp.Caption = Ftb01!henombre
                  
                Else
                   MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
                End If
                Ftb01.Close
            End If
            fqry01 = "SELECT * " _
                     & "FROM opruta " _
                    & "WHERE rucodigo = '" & Txtcodrut.text & "'"
                 
            Set Ftb01 = Dbs.Execute(fqry01)
            
            If Not Ftb01.EOF Then
               Lblnomrut.Caption = Ftb01!rudestino
            Else
               MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
            End If
            Ftb01.Close
            
            If L_Direccion = "N" Then
               Call Traer_Tasas_Terminal_Norte
            End If
            
        Else
           MsgBox "El Tasa de Uso Pertenece a Otra Caja", 48, "Advertencia"
           Ftb01.Close
           Txtnumcon.SetFocus
           Exit Sub
        End If
     Else
        MsgBox "No Se Puede Anular La Tasa de Uso", 48, "Advertencia"
        Txtnumcon.Enabled = True
        Txtnumcon.SetFocus
        Ftb01.Close
     End If
  End If
End Sub

Private Sub TxtObserv_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtObserv_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))
   If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 38 Then
      KeyAscii% = 0
   End If
  
End Sub
