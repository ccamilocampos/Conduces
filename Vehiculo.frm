VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Vehiculo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de Vida Vehículo"
   ClientHeight    =   8640
   ClientLeft      =   2745
   ClientTop       =   2100
   ClientWidth     =   6525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnnovedades 
      Caption         =   "&Novedades"
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
      Left            =   720
      Picture         =   "Vehiculo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Novedades del vehículo"
      Top             =   7680
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
      Left            =   3600
      Picture         =   "Vehiculo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   7680
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
      Left            =   4560
      Picture         =   "Vehiculo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Limpia la forma actual"
      Top             =   7680
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
      Left            =   2640
      Picture         =   "Vehiculo.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   7680
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
      Left            =   1680
      Picture         =   "Vehiculo.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Ingresa la empresa a la base de datos"
      Top             =   7680
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
      Left            =   5520
      Picture         =   "Vehiculo.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7680
      Width           =   855
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   120
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame1 
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
      Height          =   7320
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtNumAvi 
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   51
         ToolTipText     =   "Capacidad máxima del vehículo"
         Top             =   4680
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Despacho"
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
         Height          =   720
         Left            =   120
         TabIndex        =   49
         Top             =   5160
         Width           =   6015
         Begin VB.OptionButton OptDespacho 
            Caption         =   "Intermunicipal"
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
            Left            =   2280
            TabIndex        =   15
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptDespacho 
            Caption         =   "Influencia"
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
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptDespacho 
            Caption         =   "Ambos"
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
            Left            =   4680
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox TxtEmpCon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   22
         ToolTipText     =   "Código de la empresa con la que trabaja el vehículo (doble Click para ayuda)"
         Top             =   6840
         Width           =   855
      End
      Begin VB.TextBox TxtClase 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         ToolTipText     =   "Tipo de clase del vehículo(doble Click para ayuda)"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtCedCon 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Frame Baño 
         Caption         =   "Servicio de Baño"
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
         Height          =   720
         Left            =   4440
         TabIndex        =   44
         Top             =   6000
         Width           =   1695
         Begin VB.OptionButton OpcBano 
            Caption         =   "Si"
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
            Left            =   840
            TabIndex        =   21
            Top             =   360
            Width           =   720
         End
         Begin VB.OptionButton OpcBano 
            Caption         =   "No"
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
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   720
         End
      End
      Begin VB.TextBox Txtcodcom 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   8
         ToolTipText     =   "Código de la empresa con la que trabaja el vehículo (doble Click para ayuda)"
         Top             =   3240
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo Vehículo"
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
         Height          =   720
         Left            =   120
         TabIndex        =   42
         Top             =   6000
         Width           =   4215
         Begin VB.OptionButton Opttipo 
            Caption         =   "Contrato"
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
            Left            =   3015
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Opttipo 
            Caption         =   "Normal"
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
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Opttipo 
            Caption         =   "Empresarial"
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
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtmodelo 
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   13
         ToolTipText     =   "Módelo del vehículo"
         Top             =   4200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   4200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "MMM/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txttipveh 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         ToolTipText     =   "Tipo de vehículo (doble Click para ayuda)"
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Cmbestado 
         Height          =   315
         ItemData        =   "Vehiculo.frx":198C
         Left            =   4680
         List            =   "Vehiculo.frx":1996
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Estado del vehículo"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Txtnumpas 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         ToolTipText     =   "Capacidad máxima del vehículo"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lblnomcon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "Nombre del conductor actual"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Txtpropietario 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         ToolTipText     =   "Nombre del propietario del vehículo"
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox Txtcodemp 
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   4
         ToolTipText     =   "Código de la empresa con la que trabaja el vehículo (doble Click para ayuda)"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Txtnumint 
         Height          =   285
         Left            =   4680
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Número interno del vehículo"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Txtnumpla 
         Height          =   285
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   0
         ToolTipText     =   "El número de la placa del vehículo (doble Click para ayuda)"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Num. Avi. Condic:"
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
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Emp. Contratante:"
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
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label LblNomEmpCon 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   6840
         Width           =   3375
      End
      Begin VB.Label LblClase 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label11 
         Caption         =   "Clase:"
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
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo Combustible:"
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
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Lblnomcom 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Modelo:"
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
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha :"
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
         Top             =   4200
         Width           =   1695
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
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Tip. Servicio:"
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
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Lblnomemp 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Estado del Veh:"
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
         TabIndex        =   36
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Núm. Pasajeros:"
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
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Conductor Activo:"
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
         TabIndex        =   34
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Propietario:"
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
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         TabIndex        =   32
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Número Interno:"
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
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Núm. Placa:"
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
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   6480
      X2              =   6480
      Y1              =   8640
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   6480
      Y1              =   7560
      Y2              =   7560
   End
End
Attribute VB_Name = "Vehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numeinterno As String
Dim tipovehiculo As String
Dim codempresa As String
Dim sEstado As String
Dim NumUni As Long
Dim dbsca() As Database
Dim Exi_Pla() As String
Dim Dir_Unix As String

Dim L_IntegracionProtech, L_UrlApiInsVehiculo, L_UrlApiUpdVehiculo, L_UrlServer As String

Dim L_Bandera_Venta_Condicionada As Boolean

Private Function Trae_Clase()
   On Error GoTo Errores
   ReDim sdato(10, 2)
                
   QRY = "SELECT dscoddet, dsdes " _
         & "FROM gesuptip, gedetsuptip " _
        & "WHERE  dscodtip = stcodtip " _
          & "AND stdes = 'RELACION CLASE-PASAJEROS' AND dsest = 'A' "
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   i = 0
   j = 0
   Do While Not Tbl1.EOF
      sdato(i, 0) = Tbl1(0)
      sdato(i, 1) = Tbl1(1)
      Tbl1.MoveNext
      i = i + 1
   Loop
   Tbl1.Close
   
   'se valida la cantidad de pasajeros para saber cual es el grupo
   If Val(Txtnumpas.text) >= 0 And Val(Txtnumpas.text) <= Val(sdato(0, 1)) Then
      TxtClase.text = sdato(0, 0)
   ElseIf Val(Txtnumpas.text) > Val(sdato(0, 1)) And Val(Txtnumpas.text) <= Val(sdato(1, 1)) Then
      TxtClase.text = sdato(1, 0)
   ElseIf Val(Txtnumpas.text) > Val(sdato(1, 1)) Then
      TxtClase.text = sdato(2, 0)
   End If
   TxtClase_LostFocus
   Exit Function
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function
End Function

Private Sub Trae_Estado(ByVal sPlaca As String)
   fqry01 = "SELECT hvestado  " _
            & "FROM ophojveh " _
           & "WHERE hvplaca = '" & sPlaca & "'"
     
   Set Ftb01 = Dbs.Execute(fqry01)
     
   If Not Ftb01.EOF Then
      sEstado = Ftb01!hvestado
      If Ftb01!hvestado = "Activo" Then
         Cmbestado.text = Cmbestado.List(0)
      Else
         Cmbestado.text = Cmbestado.List(1)
      End If
   End If
End Sub
Private Sub Btnbuscar_Click()
  forma = "2"
  forma1 = "1"
  buscar.Show 1
  Txtnumpla.Enabled = True
  Txtnumpla.SetFocus
  Txtnumpla_LostFocus
End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
   
   Dim lPlaca, lCodEmpresa, lNomEmpresa, lNivelServicio, lReserva, lEstado As String
   Dim lNumConvenio, lCodEmpConvenio, lNomEmpConvenio, lFecIniConvenio, lFecFinConvenio, lTag, lFecha As String
      
   If Txtnumpla.text = "" Then
      MsgBox "Falta Ingresar El Numero De la Placa", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
  
   If Txtnumint.text = "" Then
      MsgBox "Falta Ingresar El Numero Interno Del Vehiculo", 48, "Advertencia"
      Txtnumint.SetFocus
      Exit Sub
   End If
  
   If Txtcodemp.text = "" Then
      MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
      Txtcodemp.SetFocus
      Exit Sub
   End If
  
   If Txttipveh.text = "" Then
      MsgBox "Falta Ingresar El Tipo De Vehiculo", 48, "Advertencia"
      Txttipveh.SetFocus
      Exit Sub
   End If
  
  
   If TxtCedCon.text = "" Then
      MsgBox "Falta Ingresar la Cedula del conductor", 48, "Advertencia"
      TxtCedCon.SetFocus
      Exit Sub
   End If
  
   If Txtcodcom.text = "" Then
      MsgBox "Falta Ingresar El Tipo De Combustible", 48, "Advertencia"
      Txtcodcom.SetFocus
      Exit Sub
   End If
   If Txtpropietario.text = "" Then
      'MsgBox "Falta Ingresar El Propietario Del Vehiculo", 48, "Advertencia"
      'Txtpropietario.SetFocus
      'Exit Sub
      Txtpropietario.text = "PENDIENTE"
   End If
  
   If Lblnomcon.text = "" Then
      'MsgBox "Falta Ingresar El Campo Conductor Activo", 48, "Advertencia"
      'Lblnomcon.SetFocus
      'Exit Sub
      Lblnomcon.text = "PENDIENTE"
   End If
   
   If Txtnumpas.text = "" Then
      MsgBox "Falta Ingresar El Numero De Pasajeros", 48, "Advertencia"
      Txtnumpas.SetFocus
      Exit Sub
   End If
   
   If Cmbestado.text = "" Then
      MsgBox "Falta Ingresar El Estado", 48, "Advertencia"
      Cmbestado.SetFocus
      Exit Sub
   End If
   
   If Mskfecha.ClipText = "" Then
      MsgBox "Falta Ingresar La Fecha De La Ultima Prueba De Alcolemia", 48, "Advertencia"
      Mskfecha.SetFocus
      Exit Sub
   End If
   
   If Trim(txtmodelo.text) = "" Then
      txtmodelo.text = "0"
   End If
   sEmpresaConvenio = ""
   If OptTipo(0).Value = True Then
      sTipo = "N"
   ElseIf OptTipo(1).Value = True Or OptTipo(2).Value = True Then
      If OptTipo(1).Value = True Then
         sTipo = "E"
      ElseIf OptTipo(2).Value = True Then
         sTipo = "T"
      End If
        If Trim(TxtEmpCon.text) = "" Then
           MsgBox "Falta ingresar el código de la empresa del convenio", 48, "Advertencia"
           TxtEmpCon.SetFocus
           Exit Sub
        End If
        sEmpresaConvenio = Trim(TxtEmpCon.text)
      
   End If
   
   If OptDespacho(0).Value = True Then
       sTipoDespacho = "I"
   ElseIf OptDespacho(1).Value = True Then
       sTipoDespacho = "N"
   ElseIf OptDespacho(2).Value = True Then
       sTipoDespacho = "A"
   End If
   
   If OpcBano(0).Value = True Then
      sBano = "NO"
   Else
      sBano = "SI"
   End If
   
   If Trim(TxtClase.text) = "" Then
      Trae_Clase
   End If
   
   QRY = "SELECT * " _
         & "FROM ophojveh " _
        & "WHERE hvplaca <> '" & Txtnumpla.text & "' " _
          & "AND hvcodemp = '" & Txtcodemp.text & "' " _
          & "AND hvnumint = " & Txtnumint.text & " " _
          & "AND hvestado = 'Activo' " _
  
   Set Tbl1 = Dbs.Execute(QRY)
  
   If Not Tbl1.EOF Then
      MsgBox "El Vehiculo Con Placa " & Tbl1!hvplaca & Chr(13) & "Tiene El Mismo Número Interno", 48, "Advertencia"
      Txtnumint.Enabled = True
      Txtnumint.SetFocus
      Exit Sub
   End If
  
   If numeinterno <> "" And codempresa <> "" And tipovehiculo <> "" And sEstado <> "" Then
      If numeinterno <> Txtnumint.text Or codempresa <> Txtcodemp.text Or tipovehiculo <> Txttipveh.text Or sEstado <> Trim(Cmbestado.text) Then
         MsgBox "No Ha Ingresado Las Novedades", 48, "Advertencia"
      
         Exit Sub
      End If
   End If
   respuesta = MsgBox("Desea Actualizar El Vehiculo", 32 + vbYesNo, "GRABAR")
  
   If respuesta = vbNo Then
      Exit Sub
   End If
  
   If Txtnumpla.Enabled = False Then
      G_Clase_Autorizacion = "ACTVEH"
      Relogin.Show 1
      If G_Aut_Vender <> "SI" Then
         Exit Sub
      End If
      
   End If

   If Cmbestado.text = "Activo" Then
      sEstado = "1"
   Else
      sEstado = "0"
   End If
  
   Veh_Nuevo = "0"
   
   Dbs.BeginTrans
   
   lPlaca = Trim(Txtnumpla.text)
   lCodEmpresa = Trim(Txtcodemp.text)
   lNomEmpresa = Trim(Lblnomemp.Caption)
   lNivelServicio = Trim(Txttipveh.text)
   lReserva = "N"
   lEstado = sEstado

   Call CreandoJson("V", lPlaca, lCodEmpresa, lNomEmpresa, lNivelServicio, lReserva, lEstado, "", "", "")
   
   sBandera = CadenaJSON
   sParametro = ""
   On Error GoTo transac
   If Txtnumpla.Enabled = True Then
     
      QRY = " INSERT INTO ophojveh " _
                 & "(hvplaca, hvnumint, hvcodemp, hvpropie, " _
                 & " hvconact, hvnumpas, hvestado, hvtipveh, " _
                 & " hvfecha, hvprueba, hvmodelo, hvtipo," _
                 & " hvcombus, HVBANO, HVCONDUCT, hvclase, " _
                 & " hvempcon, hvtipodespacho, hvnumaviprot) VALUES ('" _
                 & Txtnumpla.text & "'," & Txtnumint.text & ",'" & Txtcodemp.text & "','" & Txtpropietario.text & "','" _
                 & Lblnomcon.text & "'," & Txtnumpas.text & ",'" & Cmbestado.text & "','" & Txttipveh.text & "',TO_DATE('" _
                 & Mskfecha.text & "','MM/DD/YYYY'),'No'," & Trim(txtmodelo.text) & ",'" & sTipo & "','" _
                 & Txtcodcom.text & "', '" & sBano & "', " & Trim(TxtCedCon.text) & ",'" & TxtClase.text & "', '" _
                 & sEmpresaConvenio & "','" & sTipoDespacho & "', " & txtNumAvi.text & ")"
     
     Dbs.Execute QRY
     Fecha1 = Format(Now, "MM/dd/yyyy hh:mm")
     'On Error GoTo transac
     
     QRY = " INSERT INTO gelogaud " _
               & "(lofecact, lofecnue, lofuncio, " _
               & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
               & "VALUES ( TO_DATE('" & Fecha1 & "', 'MM/dd/yyyy hh24:mi:ss'),  TO_DATE('" & Fecha1 & "', 'MM/dd/yyyy hh24:mi:ss'),'" _
               & Login & "', TO_DATE('" & Fecha1 & "', 'MM/dd/yyyy hh24:mi:ss'),'SE CREO UN VEHICULO', 'VEH', '" _
               & Txtnumpla.text & "', '" & Lblnomemp.Caption & "')"
     
     Dbs.Execute QRY
     
     Veh_Nuevo = "1"
     
      sUrl = L_UrlApiInsVehiculo
      sEvento = "INS"
  Else
     'bds del servidor
     
     QRY = "UPDATE ophojveh " _
                  & "SET hvnumint =  " & Txtnumint.text & " ," _
                  & "    hvcodemp = '" & Txtcodemp.text & "'," _
                  & "    hvpropie = '" & Txtpropietario.text & "'," _
                  & "    hvconact = '" & Lblnomcon.text & "'," _
                  & "    hvnumpas =  " & Txtnumpas.text & " ," _
                  & "    hvtipveh = '" & Txttipveh.text & "'," _
                  & "    hvestado = '" & Cmbestado.text & "'," _
                  & "    hvmodelo = " & Trim(txtmodelo.text) & ", " _
                  & "    hvtipo = '" & Trim(sTipo) & "', " _
                  & "    hvcombus = '" & Txtcodcom.text & "', " _
                  & "    hvfecha = to_date('" & Mskfecha.text & "', 'MM/dd/yyyy'), " _
                  & "    HVBANO =  '" & sBano & "', " _
                  & "    hvclase = '" & TxtClase.text & "', " _
                  & "    hvempcon = '" & sEmpresaConvenio & "', " _
                  & "    HVCONDUCT =  " & Trim(TxtCedCon.text) & ", " _
                  & "    hvtipodespacho = '" & sTipoDespacho & "', " _
                  & "    hvnumaviprot = " & txtNumAvi.text & " " _
                & "WHERE hvplaca  = '" & Txtnumpla.text & "' "
                
     Dbs.Execute QRY
     
     If txtNumAvi.text = "0" And L_Bandera_Venta_Condicionada = True Then
        QRY = "UPDATE ophojveh " _
                & "SET hvbandvencon = 1 " _
                & "WHERE hvplaca = '" & Txtnumpla.text & "' "
        Dbs.Execute QRY
     End If
      
      sUrl = L_UrlApiUpdVehiculo
      sEvento = "INS"
  End If
  
   If L_IntegracionProtech = "S" Then
      sRespuesta = Fn_Web_Sevices_Protech(L_UrlServer, sUrl, sEvento, sBandera, sParametro)
      
      If sRespuesta <> "0" Then
         sFechaSistemaProtech = Fn_Parametros("HORASIS", 1)
         
         QRY = "INSERT INTO logprotech(lotipo, lovalor, lodescripcion, lofecha)" _
         & "VALUES ('V','" & Txtnumpla.text & "','" & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta & "',TO_DATE('" & sFechaSistemaProtech & "','MM/dd/yyyy HH24:mi:ss'))"
         
         Dbs.Execute QRY
         
         MsgBox "Se presento error en el webservices de Protech por: " & Chr(13) & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta, 48, "Advertencia"
      End If
   End If
  Dbs.CommitTrans
  On Error GoTo Errores
  If Veh_Nuevo = "1" Then
      'Generación de plano
      NOMDIR = Dir_Unix & "bus.csv"
      
      NOMBRE = Dir(NOMDIR)
      
      If NOMBRE <> "" Then
         Kill NOMDIR
      End If
      
      
      On Error Resume Next
      Open NOMDIR For Append As #2
      If Err Then
         MsgBox Error$, 48
         
         Close #2
      Else
         Print #2, "KEY;NLG1;TYPE;LOGO"; Chr$(13)
         
         a = Trim(Txtnumpla.text) & ";" & Trim(Lblnomemp.Caption) & ";" & _
             Trim(Lbltipveh.Caption) & ";" & Trim(Txtcodemp.text) & Chr(13)
         Print #2, a
         Close #2
         
         
         NOMDIR = Dir_Unix & "Unix.BAT"
            
         NOMBRE = Dir(NOMDIR)
         
         
         If NOMBRE <> "" Then
           On Error Resume Next
            Close #2
            Kill NOMDIR
            On Error GoTo Errores
         End If
         
         On Error Resume Next
         Open NOMDIR For Append As #2
         If Err Then
            'MsgBox Error$, 48
            Close #2
            Exit Sub
         Else
            Dir_Unix1 = Mid(Dir_Unix, 1, Len(Dir_Unix) - 1)
            Print #2, "cd\"; Chr$(13)
            Print #2, "cd " & Mid(Dir_Unix1, 4); Chr$(13)
            Print #2, "ftp -s:unixcmd.txt -n " & G_IpUnix; Chr$(13)
                
         End If
         Close #2
         
         Shell Dir_Unix & "Unix.BAT", vbHide
         
      End If
  
  End If
  
  MsgBox "Vehículo Actualizado", 64, "O.K."
  
  LIMPIAR
  Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number

   Exit Sub
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo Errores
   
   If Txtnumpla.Enabled = False Then
     
      
      Dbs.BeginTrans
      On Error GoTo ErrTransac

      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      
      QRY = "INSERT INTO optempo1 " _
                & "(temtext1, temnumero1, temtext2, temtext3, temtext4, temnumero2, temtext5, tempc, temtext6, temtext7) " _
                & "SELECT hvplaca, hvnumint, hvcodemp, hvpropie, hvconact, hvnumpas, hvestado, '" & Pc & "', hvmodelo,hvcombus " _
                & "FROM ophojveh " _
                & "WHERE hvplaca = '" & Txtnumpla.text & "'"

      Dbs.Execute QRY
      
      Dbs.CommitTrans
      On Error GoTo Errores
      
      rptterminal.WindowTitle = "REPORTE DE VEHICULO"
      reporte = RpPath + "\vehiculo.rpt"
      'rptterminal.DataFiles(0) = dbpath
      'rptterminal.password = Passwd
      rptterminal.ReportFileName = reporte
      rptterminal.Connect = G_Conexion_reportes
      rptterminal.Formulas(0) = "nomemp = '" & Lblnomemp.Caption & "'"
      rptterminal.Formulas(1) = "tipveh = '" & Lbltipveh.Caption & "'"
      rptterminal.Formulas(2) = "empresa = '" & G_NomEmpresa & "'"
      Vehiculo.Enabled = False
      rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
      rptterminal.WindowState = crptMaximized
      rptterminal.Destination = 0
      rptterminal.Action = 1
     
     
      Vehiculo.Enabled = True
      
      Dbs.BeginTrans
      
         Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      
      Dbs.CommitTrans
   End If
   Exit Sub
     
Errores:
   Vehiculo.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
ErrTransac:
   Vehiculo.Enabled = True
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub

   
   
     
End Sub

Private Sub Btnlimpiar_Click()
  LIMPIAR
End Sub

Private Sub btnnovedades_Click()
   numeinterno = Txtnumint.text
   codempresa = Txtcodemp.text
   tipovehiculo = Txttipveh.text
   sEstado = Cmbestado.text
   G_Tipo_Novedad = "PLACA"
   detvehic.Show 1
   Trae_Estado (Txtnumpla.text)
End Sub

Private Sub Btnsalir_Click()
    On Error Resume Next
    'For j = 1 To NumUni
    '   dbsca(1, j).Close
    'Next j
   
    Unload Me
End Sub

Private Function LIMPIAR()
  Txtnumpla.Enabled = True
  Txtnumpla.text = ""
  Txtnumpla.SetFocus
  Txtnumint.text = ""
  Txtcodemp.text = ""
  Txtpropietario.text = ""
  Lblnomcon.text = ""
  Txtnumpas.text = ""
  Btngrabar.Caption = "&Grabar"
  btnnovedades.Enabled = False
  Lblnomemp.Caption = ""
  Txttipveh.text = ""
  Lbltipveh.Caption = ""
  Mskfecha.Mask = ""
  Mskfecha.text = ""
  Mskfecha.Mask = "##/##/####"
  Cmbestado.text = Cmbestado.List(0)
  txtmodelo.text = ""
  Txtcodcom.text = ""
  Lblnomcom.Caption = ""
  TxtClase.text = ""
  LblClase.Caption = ""
  OptDespacho(0).Value = True
  txtNumAvi.text = ""
  
  fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
 
  Set Ftb01 = Dbs.Execute(fqry01)
  
  If Not Ftb01.EOF Then
     Mskfecha.text = Mid(Ftb01!psval, 1, 10)
  End If
  Ftb01.Close
  numeinterno = ""
  codempresa = ""
  tipovehiculo = ""
  sEstado = ""
  OptTipo(0).Value = True
  OpcBano(0).Value = True
  TxtCedCon.text = ""
  TxtEmpCon.Enabled = False
  TxtEmpCon.text = ""
  LblNomEmpCon.Caption = ""
  L_Bandera_Venta_Condicionada = False
End Function

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Vehiculo")
End Sub

Private Sub Form_Load()
   On Error GoTo errorres
  
   Cmbestado.text = Cmbestado.List(0) ' Activo
   fqry01 = "SELECT psval " _
             & "FROM geparsis " _
            & "WHERE pscod = 'HORASIS' "
           
   Set Ftb01 = Dbs.Execute(fqry01)
   
    
   If Not Ftb01.EOF Then
      Mskfecha.text = Mid(Ftb01!psval, 1, 10)
   End If
   Ftb01.Close
   
   fqry01 = "SELECT psval " _
             & "FROM geparsis " _
            & "WHERE pscod = 'DIRUNIX' "
           
    
   Set Ftb01 = Dbs.Execute(fqry01)
    
   If Not Ftb01.EOF Then
      Dir_Unix = Trim(Ftb01!psval)
   End If
   Ftb01.Close
   
   
   L_IntegracionProtech = Fn_Parametros("INTPROTECH", 1)
  
   If L_IntegracionProtech = "" Then
      L_IntegracionProtech = "N"
   End If
   If L_IntegracionProtech = "S" Then
      L_UrlServer = Fn_Parametros("URLSERVER", 1)
      L_UrlApiInsVehiculo = Fn_Parametros("URLAPIINSV", 1)
      L_UrlApiUpdVehiculo = Fn_Parametros("URLAPIUPDV", 1)
   End If
   
   txtNumAvi.text = "0"
   
   Exit Sub
errorres:
  
  MsgBox Err.Description, 16, "Error " & Err.Number
  Exit Sub
End Sub



Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Mskfecha_GotFocus()
   Senalar_Texto Me
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
  End If
End Sub


Private Sub OptTipo_Click(index As Integer)
    TxtEmpCon.Enabled = False
    If OptTipo(2).Value = True Or OptTipo(1).Value = True Then
        TxtEmpCon.Enabled = True
    End If
End Sub

Private Sub TxtCedCon_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtCedCon_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtCedCon_LostFocus()
On Error GoTo Errores
   If TxtCedCon.text <> "" Then
      QRY = "SELECT conombres, coapellidos " & _
              "FROM coconductor " & _
             "WHERE cocedula = '" & TxtCedCon.text & "'"
      Set Tbl1 = Dbs.Execute(QRY)
      If Not Tbl1.EOF Then
         Lblnomcon.text = Tbl1!conombres & " " & Tbl1!coapellidos
      Else
         MsgBox "La cédula del conductor no existe", 48, "Advertencia"
         TxtCedCon.text = ""
         On Error Resume Next
         TxtCedCon.SetFocus
         Exit Sub
      End If
      Tbl1.Close
   Else
      Lblnomcon.text = ""
   End If
Exit Sub
Errores:
MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub TxtClase_DblClick()
  forma = "16"
  forma1 = "1"
  TxtClase.text = ""
  G_Nom_Catalogo = "TIPOS DE CLASES"
  buscar.Show 1
  TxtClase_LostFocus

End Sub

Private Sub TxtClase_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtClase_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub TxtClase_LostFocus()
   If TxtClase.text <> "" Then
     fqry01 = "SELECT dsdes, dsest " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & TxtClase.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'TIPOS DE CLASES' "
               
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        If Ftb01!dsest = "I" Then
           Ftb01.Close
           TxtClase.SetFocus
           MsgBox "La clase esta inactiva", 48, "Advertencia"
           Exit Sub
        End If
        LblClase.Caption = Ftb01!dsdes
     Else
        MsgBox "Tipo de clase no existe", 48, "Advertencia"
'        forma = "16"
'        forma1 = "1"
'        Buscar.Show 1
        
     End If
     Ftb01.Close
  End If

End Sub


Private Sub Txtcodcom_DblClick()
  forma = "15"
  forma1 = "30"
  Tabla_Catalogo = "TIPO COMBUSTIBLE"
  Txtcodcom.text = ""
  buscar.Show 1
  Txtcodcom.SetFocus
  Txtcodcom_LostFocus
End Sub

Private Sub Txtcodcom_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtcodcom_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub
Private Sub Txtcodcom_LostFocus()
If Txtcodcom.text <> "" Then
     fqry01 = "SELECT dsdes " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & Txtcodcom.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'TIPO COMBUSTIBLE' "
               
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lblnomcom.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        Ftb01.Close
        MsgBox "Tipo De Combustible  No Existe", 48, "Advertencia"
        forma = "15"
        forma1 = "30"
        Tabla_Catalogo = "TIPO COMBUSTIBLE"
        buscar.Show 1
        Txtcodcom.SetFocus
     End If
  End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "2"
  Txtcodemp.text = ""
  buscar.Show 1
  Txtcodemp.SetFocus
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
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
      fqry01 = "SELECT henombre " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & Txtcodemp.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         Lblnomemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
        
         fqry01 = "SELECT henombre " _
                  & "FROM ophojempcon " _
                 & "WHERE hecodigo = '" & Txtcodemp.text & "'"
             
         Set Ftb01 = Dbs.Execute(fqry01)
     
         If Not Ftb01.EOF Then
            Lblnomemp.Caption = Ftb01!henombre
            Ftb01.Close
         Else
        
            MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
            forma = "16"
            forma1 = "2"
            buscar.Show 1
            Txtcodemp.SetFocus
         End If
      End If
   End If
End Sub

Private Sub Lblnomcon_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


Private Sub txtempcon_DblClick()
    forma = "16"
    forma1 = "47"
    TxtEmpCon.text = ""
    buscar.Show 1
    TxtEmpCon.SetFocus
    txtempcon_LostFocus
End Sub

Private Sub txtempcon_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtempcon_KeyPress(KeyAscii As Integer)
If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtempcon_LostFocus()
    If TxtEmpCon.text <> "" Then
      fqry01 = "SELECT henombre " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & TxtEmpCon.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         LblNomEmpCon.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
         TxtEmpCon.SetFocus
         Exit Sub
        
      End If
   End If
End Sub

Private Sub txtmodelo_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtmodelo_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub txtNumAvi_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub txtNumAvi_KeyPress(KeyAscii As Integer)
    If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumint_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnumint_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumpas_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnumpas_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumpas_LostFocus()
   If Trim(Txtnumpas.text) <> "" Then

'      If Val(Txtnumpas.Text) > 55 Then
'         MsgBox "Sobrepasa la capacidad maxima de cupo de sillas establecido", 48, "Advertencia"
'         Txtnumpas.SetFocus
'         Exit Sub
'      End If
      
      Trae_Clase
      
   End If
End Sub

Private Sub Txtnumpla_DblClick()
  forma = "2"
  forma1 = "1"
  Txtnumpla.text = ""
  buscar.Show 1
  Txtnumpla.Enabled = True
  Txtnumpla.SetFocus
  Txtnumpla_LostFocus
End Sub

Private Sub Txtnumpla_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnumpla_KeyPress(KeyAscii As Integer)
'  Char = Chr(KeyAscii%)
'  KeyAscii = Asc(UCase(Char))  ' mayuscula
'  longitud = Len(Txtnumpla.Text)
'  longitud = longitud + 1
'  If longitud <= 2 Then
'     If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'         Exit Sub
'     Else
'     '    MsgBox "En El Numero De Placa las Tres Primeras Deben Ser Letras", 48, "Advertencia"
'         KeyAscii% = 0
'         Exit Sub
'     End If
'  ElseIf longitud > 3 Then
'     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
'        Exit Sub
'     Else
'     '   MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
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

Private Sub Txtnumpla_LostFocus()
  On Error GoTo Errores
  
  If Txtnumpla.text <> "" Then
     
     longitud = Len(Txtnumpla.text)
  
     If longitud < 6 Then
        MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
        Txtnumpla.SetFocus
        Exit Sub
     End If
     
     fqry01 = "SELECT * " _
              & "FROM ophojveh " _
             & "WHERE hvplaca = '" & Txtnumpla.text & "'"
     
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Txtnumpla.Enabled = False
        If IsNull(Ftb01!hvnumint) Then
           Txtnumint.text = 0
        Else
        Txtnumint.text = Ftb01!hvnumint
        End If
        numeinterno = Ftb01!hvnumint
        Txtcodemp.text = Ftb01!hvcodemp
        codempresa = Ftb01!hvcodemp
        Txtpropietario.text = Ftb01!hvpropie
        Lblnomcon.text = Ftb01!hvconact
        Txtnumpas.text = Ftb01!hvnumpas
        Txttipveh.text = Ftb01!hvtipveh
        tipovehiculo = Ftb01!hvtipveh
        txtNumAvi.text = Ftb01!hvnumaviprot
        If txtNumAvi.text = "1" Then
            L_Bandera_Venta_Condicionada = True
        End If
        If Not IsNull(Ftb01!hvfecha) Then
           Mskfecha.text = Format(Ftb01!hvfecha, "MM/DD/YYYY")
        Else
           Mskfecha.text = Date
        End If
        sEstado = Ftb01!hvestado
        If Ftb01!hvestado = "Activo" Then
           Cmbestado.text = Cmbestado.List(0)
        Else
           Cmbestado.text = Cmbestado.List(1)
        End If
        
        If IsNull(Ftb01!hvmodelo) Then
           txtmodelo.text = "0"
        Else
           txtmodelo.text = Ftb01!hvmodelo
        End If
        
        If Not IsNull(Ftb01!hvtipo) Then
           If Ftb01!hvtipo = "E" Then
               OptTipo(1).Value = True
           ElseIf Ftb01!hvtipo = "C" Or Ftb01!hvtipo = "N" Then
               OptTipo(0).Value = True
           ElseIf Ftb01!hvtipo = "T" Then
               OptTipo(2).Value = True
           End If
        Else
           OptTipo(0).Value = True
        End If
        
        If Not IsNull(Ftb01!hvtipodespacho) Then
           If Ftb01!hvtipodespacho = "I" Then
               OptDespacho(0).Value = True
           ElseIf Ftb01!hvtipodespacho = "N" Then
               OptDespacho(1).Value = True
           ElseIf Ftb01!hvtipodespacho = "A" Then
               OptDespacho(2).Value = True
           End If
        Else
           OptDespacho(1).Value = True
        End If
        
        If Not IsNull(Ftb01!hvcombus) Then
          Txtcodcom.text = Ftb01!hvcombus
        End If
        '--------
        
        If Not IsNull(Ftb01!hvbano) Then
           If Ftb01!hvbano = "NO" Then
               OpcBano(0).Value = True
           ElseIf Ftb01!hvbano = "SI" Then
               OpcBano(1).Value = True
           End If
        Else
           OpcBano(0).Value = True
        End If
        
        If Not IsNull(Ftb01!HVCONDUCT) Then
           TxtCedCon.text = Ftb01!HVCONDUCT
        Else
           TxtCedCon.text = ""
        End If
        If Not IsNull(Ftb01!hvclase) Then
           TxtClase.text = Ftb01!hvclase
        End If
        If Not IsNull(Ftb01!hvempcon) Then
           TxtEmpCon.text = Ftb01!hvempcon
        End If
                 
        '-----
        Ftb01.Close
        Txtcodcom_LostFocus
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
                 & "FROM ophojemp " _
                & "WHERE hecodigo = '" & Txtcodemp.text & "'"
          
         Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
           Lblnomemp.Caption = Ftb01!henombre
        End If
        Ftb01.Close
        TxtClase_LostFocus
        
        If Trim(Lblnomemp.Caption) = "" Then
            txtcodemp_LostFocus
        End If
        If Trim(TxtEmpCon.text) <> "" Then
            txtempcon_LostFocus
        End If
        
        Btngrabar.Caption = "&Actualizar"
        btnnovedades.Enabled = True
     Else
        numeinterno = ""
        codempresa = ""
        tipovehiculo = ""
        sEstado = ""
     End If
  
  End If
  Exit Sub
Errores:
  MsgBox Err.Description, 16, "Error " & Err.Number
  Exit Sub
  Resume
End Sub

Private Sub Txtpropietario_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtpropietario_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txttipveh_DblClick()
  forma = "15"
  forma1 = "2"
  Tabla_Catalogo = "TIPOS DE VEHICULO"
  Txttipveh.text = ""
  buscar.Show 1
  Txttipveh.SetFocus
  Txttipveh_LostFocus
End Sub

Private Sub Txttipveh_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txttipveh_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
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
        forma1 = "2"
        Tabla_Catalogo = "TIPOS DE VEHICULO"
        buscar.Show 1
        Txttipveh.SetFocus
     End If
  End If
End Sub
