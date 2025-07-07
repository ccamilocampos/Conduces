VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Conconcriterio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Tasas de Uso por Criterio"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   1965
   ClientWidth     =   11730
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Caption         =   "Tipo Ruta"
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
      Height          =   975
      Left            =   4200
      TabIndex        =   76
      Top             =   3840
      Width           =   1695
      Begin VB.ComboBox CmbTipoRuta 
         Height          =   315
         ItemData        =   "Conconcr.frx":0000
         Left            =   120
         List            =   "Conconcr.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Código Super"
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
      Left            =   4200
      TabIndex        =   75
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton OptCodSuper 
         Caption         =   "Sin Código"
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
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton OptCodSuper 
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
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   810
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptCodSuper 
         Caption         =   "Con Código"
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
         Left            =   120
         TabIndex        =   20
         Top             =   555
         Width           =   1335
      End
   End
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
      Height          =   1215
      Left            =   2640
      TabIndex        =   74
      Top             =   2520
      Width           =   1455
      Begin VB.OptionButton OptTipoTasa 
         Caption         =   "Sin Viaje"
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
         Index           =   3
         Left            =   120
         TabIndex        =   77
         Top             =   960
         Width           =   1095
      End
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptTipoTasa 
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
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Alistamiento"
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
      TabIndex        =   73
      Top             =   2520
      Width           =   2415
      Begin VB.OptionButton OptAlistamiento 
         Caption         =   "Con Alistamiento"
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
         Left            =   285
         TabIndex        =   12
         Top             =   300
         Width           =   1815
      End
      Begin VB.OptionButton OptAlistamiento 
         Caption         =   "Sin Alistamiento"
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
         Left            =   285
         TabIndex        =   13
         Top             =   555
         Width           =   1815
      End
      Begin VB.OptionButton OptAlistamiento 
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
         Height          =   255
         Index           =   2
         Left            =   285
         TabIndex        =   14
         Top             =   810
         Value           =   -1  'True
         Width           =   1215
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
      Height          =   615
      Left            =   4200
      TabIndex        =   72
      Top             =   1080
      Width           =   1695
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "Conconcr.frx":0032
         Left            =   120
         List            =   "Conconcr.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Tipos de terminales"
         Top             =   240
         Width           =   1455
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   6120
      TabIndex        =   71
      Top             =   7200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
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
      Height          =   615
      Left            =   4200
      TabIndex        =   70
      Top             =   1800
      Width           =   1695
      Begin VB.ComboBox Cmbcaja 
         Height          =   315
         ItemData        =   "Conconcr.frx":0036
         Left            =   120
         List            =   "Conconcr.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Estado de la empresa"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   67
      Top             =   1080
      Width           =   3975
      Begin VB.OptionButton Optgenerar 
         Caption         =   "G/rar Plano"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Optgenerar 
         Caption         =   "Ver X pantalla"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   66
      Top             =   1800
      Width           =   3975
      Begin VB.OptionButton OptTipo 
         Caption         =   "Totalizada"
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
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detallada"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
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
      Height          =   1695
      Left            =   6120
      TabIndex        =   53
      Top             =   5400
      Width           =   5535
      Begin MSMask.MaskEdBox Msktotcon 
         Height          =   300
         Left            =   3720
         TabIndex        =   54
         Top             =   960
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
         Left            =   840
         TabIndex        =   55
         Top             =   240
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
         Left            =   3720
         TabIndex        =   56
         Top             =   240
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
         Left            =   840
         TabIndex        =   57
         Top             =   600
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
      Begin MSMask.MaskEdBox mskcredito 
         Height          =   300
         Left            =   840
         TabIndex        =   64
         Top             =   960
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
      Begin VB.Label Label11 
         Caption         =   "Pasaj.:"
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
         TabIndex        =   69
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LblPasajeros 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         Left            =   840
         TabIndex        =   68
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Créd:"
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
         TabIndex        =   65
         Top             =   960
         Width           =   735
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
         Left            =   3720
         TabIndex        =   63
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label15 
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
         Left            =   2520
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Alcoh. :"
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
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Tasas. Vend.:"
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
         TabIndex        =   60
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Tot Gral :"
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
         TabIndex        =   59
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Perm.:"
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
         TabIndex        =   58
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton Btntotalizar 
      Caption         =   "&Totalizar"
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
      Left            =   1335
      Picture         =   "Conconcr.frx":003A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Totaliza todos los netos a pagar que se encuentran en la grilla"
      Top             =   6660
      Width           =   855
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   6240
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
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
      Left            =   2535
      Picture         =   "Conconcr.frx":0344
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   6660
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
      Left            =   3735
      Picture         =   "Conconcr.frx":0786
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Limpia la forma actual"
      Top             =   6660
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
      Left            =   4935
      Picture         =   "Conconcr.frx":0BC8
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Salir de la forma actual"
      Top             =   6660
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
      Left            =   120
      Picture         =   "Conconcr.frx":100A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Búscar informacion en la base de datos"
      Top             =   6660
      Width           =   855
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
      Height          =   1695
      Left            =   120
      TabIndex        =   43
      Top             =   4920
      Width           =   5775
      Begin VB.TextBox Txtoritra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   29
         ToolTipText     =   "Código del tipo de tasa de uso(Precione doble click para ayuda)"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Txttipser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Txtcodrut 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   27
         ToolTipText     =   "Código de la ruta (Precione doble click para ayuda)"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   26
         ToolTipText     =   "Código de la empresa (Precione doble click para ayuda)"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbloritra 
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
         TabIndex        =   51
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Lbltipser 
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
         TabIndex        =   50
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Lblcodrut 
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
         TabIndex        =   49
         Top             =   600
         Width           =   3255
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
         TabIndex        =   48
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Ori/Tra:"
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
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cód. Ruta:"
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
         Top             =   600
         Width           =   1095
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
         TabIndex        =   44
         Top             =   240
         Width           =   1335
      End
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
      Height          =   975
      Left            =   120
      TabIndex        =   42
      Top             =   3840
      Width           =   3975
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Tip. Servicio"
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
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Ori/Tra"
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
         Left            =   2280
         TabIndex        =   24
         Top             =   600
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Cód. Ruta"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Por:"
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
      TabIndex        =   35
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Optbuscar 
         Caption         =   "Números"
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
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Optbuscar 
         Caption         =   "Fechas"
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grdconduces 
      Height          =   5175
      Left            =   6120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frmnumeros 
      Caption         =   "Numeros de Tasas de Uso: "
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
      Left            =   1800
      TabIndex        =   36
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Txtnumfin 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Número final hasta donde quiere que llegue"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtnumini 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Número inicial por el que empiece a buscar"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta La Número:"
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
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "De la Número:"
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
         TabIndex        =   37
         Top             =   240
         Width           =   1695
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
      Height          =   855
      Left            =   1800
      TabIndex        =   39
      Top             =   120
      Width           =   4095
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Fecha incial donde quiere que empieze a buscar"
         Top             =   480
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
         Left            =   2520
         TabIndex        =   5
         ToolTipText     =   "Fecha final hasta donde quiere que busque"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
         Left            =   2520
         TabIndex        =   41
         Top             =   240
         Width           =   1215
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
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   135
      X2              =   11895
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6000
      X2              =   6000
      Y1              =   7440
      Y2              =   0
   End
End
Attribute VB_Name = "Conconcriterio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dat
Private Sub BtnBuscar_Click()
   Dim L_DirCristal As String

   On Error GoTo terminar
   
   bandera = 0
   Fecha = Mskfecini.text
   Fecha1 = Mskfecfin.text
   Msktotcon.text = ""
   Lblconven.Caption = ""
   Mskalcoholemia.text = ""
   msktasas.text = ""
   Mskparqueadero.text = ""
   mskcredito.text = ""
   LblPasajeros.Caption = ""
   ProBar.Value = ProBar.Min
   L_DirCristal = "N"
   dato1 = 0
   dato2 = 0
   dato3 = 0
   dato4 = 0
   dato5 = 0
   Dato6 = 0
   Dato7 = 0
   Dato8 = 0
   dato9 = 0
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
   If Optbuscar(0).Value = True Then
      If Txtnumini.text = "" Then
         MsgBox "Falta Ingresar El Numero Inicial", 48, "Advertencia"
         Txtnumini.SetFocus
         Exit Sub
      End If
      
      If Txtnumfin.text = "" Then
         MsgBox "Falta Ingresar El Numero Final", 48, "Advertencia"
         Txtnumfin.SetFocus
         Exit Sub
      End If
      
      If Chkcriterio(0).Value = 1 Then
         If Txtcodemp.text = "" Then
            MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(1).Value = 1 Then
         If Txtcodrut.text = "" Then
            MsgBox "Falta Ingresar El Codigo De La Ruta", 48, "Advertencia"
            Txtcodrut.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(2).Value = 1 Then
         If Txttipser.text = "" Then
            MsgBox "Falta Ingresar El Tipo De Servicio", 48, "Advertencia"
            Txttipser.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(3).Value = 1 Then
         If Txtoritra.text = "" Then
            MsgBox "Falta Ingresar El Codigo De Origen - Transito", 48, "Advertencia"
            Txtoritra.SetFocus
            Exit Sub
         End If
      End If
      
      DIB_GRILLA
     
      
      If OptTipo(1).Value = True Then
         'detallada
         QRY = "SELECT conumero, cocodrut, rudestino, cocodemp, henombre, to_char(cofecsal,'MON/DD/YYYY'), to_char(cohorsal,'MON/DD/YYYY HH24:MI'),  " _
                    & "coplaca, conumint, coclase, coconduc, cooritra, cofuncio, coturno, comodo, copasaje, covalor, coestado, covalalc, " _
                    & "coparque, conomcaj, cocredit, codirecc, copeaje, cocedula, NVL(cocedula1,' ') AS cocedula1, NVL(coconduc1,' ') AS coconduc1  " _
            & "FROM coconduc, opruta, ophojemp " _
           & "WHERE conumero BETWEEN " & Txtnumini.text & " " _
             & "AND " & Txtnumfin.text & " " _
             & "AND cocodrut = rucodigo " _
             & "AND cocodemp = hecodigo "
      ElseIf OptTipo(0).Value = True Then
         'Totalizada
         QRY = "SELECT COUNT(conumero), SUM(covalor), SUM(covalalc), SUM(coparque), SUM(copasaje) " _
               & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
              & "WHERE conumero BETWEEN " & Txtnumini.text & " " _
                & "AND " & Txtnumfin.text & " "
      End If
      
      If Cmbcaja.text <> "" Then
         QRY = QRY & " AND conomcaj = '" & Cmbcaja.text & "'"
      End If
      
      If CmbTerminal.ListIndex <> 0 Then
         QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
      End If
      
      '************************ CUANDO SE BUSCAR POR CODIGO DE EMPRESA ************************************
      If Chkcriterio(0).Value = 1 Then
         QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
      End If
     
      '************************ CUANDO SE BUSCA POR CODIGO DE RUTA ****************************************
      If Chkcriterio(1).Value = 1 Then
         QRY = QRY & "AND cocodrut = '" & Txtcodrut.text & "' "
      End If
      
      '************************ CUANDO SE BUSCA POR TIPO DE SERVICIO ***************************************
      If Chkcriterio(2).Value = 1 Then
         QRY = QRY & "AND coclase = '" & Txttipser.text & "' "
      End If
      
      '*********************** CUANDO SE BUSCA POR ORIGEN TRANSITO *****************************************
      If Chkcriterio(3).Value = 1 Then
         QRY = QRY & "AND cooritra = '" & Txtoritra.text & "' "
      End If
      
      If OptAlistamiento(0).Value = True Then
         QRY = QRY & " AND conumalistamiento <> '0' "
      ElseIf OptAlistamiento(1).Value = True Then
         QRY = QRY & " AND conumalistamiento = '0' "
      End If
      
      If OptTipoTasa(0).Value = True Then
         'cupo
         QRY = QRY & "AND comodo = 'P' "
      ElseIf OptTipoTasa(1).Value = True Then
         'vendido
         QRY = QRY & "AND comodo = 'D' "
      End If
      
      If OptTipoTasa(3).Value = True Then
         QRY = QRY & "AND cooritra = 'SV' "
      Else
         QRY = QRY & "AND cooritra <> 'SV' "
      End If
      
      'WECC 07/15/2019
      'CAU2 1258 se pueda consultar solo los que se le entregaron los códigos de la super
      '----------------------------------------------------------------
      If OptCodSuper(0).Value = True Then
         'cupo
         QRY = QRY & "AND copeaje = '0' "
      ElseIf OptCodSuper(1).Value = True Then
         'vendido
         QRY = QRY & "AND copeaje <> '0' "
      End If
      '----------------------------------------------------------------
      
      'WECC 09/23/2020
      'CAU2 3056 por tipo de ruta Intermunicipal (I) o Sabanera (S)
      '----------------------------------------------------------------
      If CmbTipoRuta.ListIndex = 1 Then
         'Intermunicipal
         QRY = QRY & "AND rutipo = 'I' "
      ElseIf CmbTipoRuta.ListIndex = 2 Then
         'Sabaneras
         QRY = QRY & "AND rutipo  = 'S' "
      End If
      '----------------------------------------------------------------
      
      If OptTipo(1).Value = True Then
         QRY = QRY & "AND coestado = 'Activo' " _
              & "ORDER BY conumero ASC "
      Else
         QRY = QRY & "AND coestado = 'Activo'"
         
         QRY1 = QRY
      End If
      Grdconduces.Visible = False
      If OptTipo(1).Value = True Then
      'detallada
     
         
         Set Tbl1 = Dbs.Execute(QRY)
         
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
         
         dato1 = 0
         
         
         If Optgenerar(1).Value = True Then
            Conconcriterio.Enabled = False
            Do While Not Tbl1.EOF
               DoEvents
               If IsNull(Tbl1(18)) Then
                  valalc = "0"
               Else
                  valalc = Trim(Tbl1(18))
               End If
               
               If IsNull(Tbl1(19)) Then
                  Parquea = "0"
               Else
                  Parquea = Trim(Tbl1(19))
               End If
               
               If IsNull(Tbl1(15)) Then
                  dt15 = "0"
               Else
                  dt15 = Trim(Tbl1(15))
               End If
               
               If Tbl1(22) = "O" Then
                    sdireccion = "Oriente"
               ElseIf Tbl1(22) = "C" Then
                    sdireccion = "Occidente"
               ElseIf Tbl1(22) = "S" Then
                    sdireccion = "Sur"
               ElseIf Tbl1(22) = "N" Then
                    sdireccion = "Norte"
               End If
               
               LPeaje = "0"
               If Not IsNull(Tbl1!copeaje) Then
                  LPeaje = Tbl1!copeaje
               End If
               
               If L_DirCristal = "N" Then
                  entry = str(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(9) + _
                          Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + Tbl1(5) + Chr(9) + _
                          Tbl1(6) + Chr(9) + Tbl1(7) + Chr(9) + str(Tbl1(8)) + Chr(9) + _
                          Tbl1(9) + Chr(9) + Tbl1!cocedula + Chr(9) + Tbl1!coconduc + Chr(9) + _
                          Tbl1!cocedula1 + Chr(9) + Tbl1!coconduc1 + Chr(9) + Tbl1(11) + Chr(9) + _
                          Tbl1(12) + Chr(9) + str(Tbl1(13)) + Chr(9) + Tbl1(14) + Chr(9) + _
                          str(dt15) + Chr(9) + str(Tbl1(16)) + Chr(9) + _
                          str(valalc) + Chr(9) + Tbl1(17) + Chr(9) + str(Parquea) + Chr(9) + _
                          Tbl1(20) + Chr(9) + Tbl1(21) + Chr(9) + sdireccion + Chr(9) + CStr(LPeaje)
                  
                  Grdconduces.AddItem entry
               End If
               'ElseIf L_DirCristal = "S" Then
                  QRY = "INSERT INTO optempo1 " _
                     & "(temnumero1, temtext9, temtext2, temdate1,  " _
                     & "temtext8, temtext4, temnumero3, temtext5,  " _
                     & "temtext6, temnumero2, temnumdob3, temnumdob2,  " _
                     & "temtext7, temtext1, temnumdob4, temnumero6,  " _
                     & "temtext11, tempc, temtext13, temtext14, temtext16,temtext17, temtext18, temtext19) VALUES (" _
                     & Tbl1(0) & ",'" & Tbl1(2) & "','" & Tbl1(4) & "',to_date('" & Tbl1(5) & "','MM/DD/YYYY'),'" _
                     & Format(Tbl1(6), "MMM/dd/yyyy HH:mm") & "','" & Tbl1(7) & "'," & Tbl1(8) & ",'" & Tbl1(11) & "','" _
                     & sdireccion & "'," & Tbl1(13) & "," & dt15 & "," & Tbl1(16) & ",'" _
                     & Tbl1(17) & "','" & Tbl1(9) & "'," & valalc & "," & Parquea & ",'" _
                     & Tbl1(20) & "','" & Pc & "','" & Tbl1(21) & "', '" & LPeaje & "','" & Tbl1!cocedula & "','" & Tbl1!coconduc & "','" & Tbl1!cocedula1 & "','" & Tbl1!coconduc1 & "')"
                  
                  Dbs.Execute QRY
                  
               'End If
               ProBar.Value = ProBar.Value + 1
               
               dato1 = dato1 + CDbl(valalc)
               dato4 = dato4 + CDbl(Tbl1(16))
               dato5 = dato5 + CDbl(Parquea)
               Dato6 = Dato6 + CDbl(valalc) + CDbl(Tbl1(16)) + CDbl(Parquea)
               If Tbl1(21) = "S" Then
                  Dato7 = Dato7 + Tbl1(16)
               End If
               dato9 = dato9 + Tbl1(15)
               
               Tbl1.MoveNext
               bandera = 1
            Loop
            Tbl1.Close
            Msktotcon.text = Dato6
            Mskalcoholemia.text = dato1
            msktasas.text = dato4
            Mskparqueadero.text = dato5
            Msktotcon.text = Dato6
            If IsNull(dato9) Then
                dato9 = 0
            End If
            LblPasajeros.Caption = dato9
            Lblconven.Caption = Grdconduces.rows - 1
            If Dato7 <> 0 Then
              mskcredito.text = Dato7
            Else
               mskcredito.text = "0"
            End If
         ElseIf Optgenerar(0).Value = True Then
            respuesta = InputBox("Digite un nombre para el plano", "GENERAR", "TASAS DE USO")
            If respuesta = "" Then
               MsgBox "Debe Escoger Un Nombre", 48, "Advertencia"
               Exit Sub
            End If
            NOMDIR = App.Path & "\" & respuesta & ".TXT"
             
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
               
               Print #2, "01|"; Now; "|REPORTE TASAS DE USO POR CRITERIO|"; Chr$(13)
               Print #2, "01|DEL NUMERO :|"; Trim(Txtnumini.text); "|HASTA EL NUMERO: |"; Trim(Txtnumfin.text); "|"; Chr$(13)
               Print #2, "HIDER|Numero|Cod.Ruta|Destino|Cod.Emp|Nombre|Fec.Sal|Hor.Sal|Num.Plac|Num.Interno|Tip. Servicio|Cedula1|Conductor1|Cedula2|Conductor2|Ori/Tra|Funcionario|Turno|Modo|Pasajeros|Valor|Alcoholemia|Parqueadero|Estado|Caja|Credito|Peaje"
            
               If Not Tbl1.EOF Then
                  Tbl1.MoveLast
                  ProBar.Max = Tbl1.RecordCount + 2
                  Tbl1.MoveFirst
               End If
               Do While Tbl1.EOF = False
                  DoEvents
                  QRY1 = "SELECT hvnumint " _
                         & "FROM ophojveh " _
                        & "WHERE hvplaca  = '" & Trim(Tbl1!coplaca) & "'"
                        
                  'Set tbl2 = dbs.OpenRecordset(qry1, dbOpenSnapshot)
                  
                  Set Tbl2 = Dbs.Execute(QRY1)
                  
                  LPeaje = "0"
                  If Not IsNull(Tbl1!copeaje) Then
                     LPeaje = Tbl1!copeaje
                  End If
                  
                  a = "02|" & Trim(Tbl1!conumero) & "|" & Trim(Tbl1!cocodrut) & "|" & _
                      Trim(Tbl1!rudestino) & "|" & Trim(Tbl1!cocodemp) & "|" & Trim(Tbl1!henombre) & "|" & _
                      Trim(Tbl1(5)) & "|" & Trim(Tbl1(6)) & "|" & Trim(Tbl1!coplaca) & "|" & _
                      Trim(Tbl2!hvnumint) & "|" & Trim(Tbl1!coclase) & "|" & Trim(Tbl1!cocedula) & "|" & Trim(Tbl1!coconduc) & "|" & _
                      Trim(Tbl1!cocedula1) & "|" & Trim(Tbl1!coconduc1) & "|" & Trim(Tbl1!Cooritra) & "|" & _
                      Trim(Tbl1!cofuncio) & "|" & Trim(Tbl1!coturno) & "|" & Trim(Tbl1!comodo) & "|" & _
                      Trim(Tbl1!copasaje) & "|" & Trim(Tbl1!covalor) & "|" & Trim(Tbl1!covalalc) & "|" & _
                      Trim(Tbl1!coparque) & "|" & Trim(Tbl1!coestado) & "|" & Trim(Tbl1!conomcaj) & "|" & _
                      Trim(Tbl1!cocredit) & "|" & Trim(LPeaje) & "|" & Chr(13)
                  
                  Print #2, a
                  con = con + 1
                  total = total + CDbl(Tbl1!covalor)
                  ProBar.Value = ProBar.Value + 1
                  bandera = 1
                  Tbl1.MoveNext
               Loop
               Print #2, "03|NUMERO DE REGISTROS|"; Trim(con); "|VALOR TOTAL TASAS DE USO|"; Trim(total); "|"
               Close #2
               Ftb01.Close
               MsgBox "Plano Generado con el nombre de : " & NOMDIR, 64, "O.K."
            End If
         End If
      ElseIf OptTipo(0).Value = True Then
         'totalizada
         
         Set Tbl1 = Dbs.Execute(QRY)
         
      
         If Not Tbl1.EOF Then
            If Not IsNull(Tbl1(1)) Then
               Mskalcoholemia.text = Tbl1(2)
               msktasas.text = Tbl1(1)
               Mskparqueadero.text = Tbl1(3)
               Lblconven.Caption = Tbl1(0)
               LblPasajeros.Caption = Tbl1(4)
               bandera = 1
               
            Else
               Mskalcoholemia.text = 0
               msktasas.text = 0
               Mskparqueadero.text = 0
               Lblconven.Caption = 0
               LblPasajeros.Caption = 0
            End If
         End If
            
         Tbl1.Close
        
         
         QRY = QRY1 & "AND cocredit = 'S'"
                       
         
         Set Tbl1 = Dbs.Execute(QRY)
         
         If Not Tbl1.EOF Then
            If Not IsNull(Tbl1(1)) Then
               mskcredito.text = Tbl1(1)
               bandera = 1
            End If
            
         End If
         Tbl1.Close

         If bandera = 1 Then
            Msktotcon.text = CDbl(Mskalcoholemia.text) + CDbl(Mskparqueadero.text) + CDbl(msktasas.text)
         End If
      
      
      End If
   End If
  
   If Optbuscar(1).Value = True Then
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
      
      If Chkcriterio(0).Value = 1 Then
         If Txtcodemp.text = "" Then
            MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(1).Value = 1 Then
         If Txtcodrut.text = "" Then
            MsgBox "Falta Ingresar El Codigo De La Ruta", 48, "Advertencia"
            Txtcodrut.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(2).Value = 1 Then
         If Txttipser.text = "" Then
            MsgBox "Falta Ingresar El Tipo De Servicio", 48, "Advertencia"
            Txttipser.SetFocus
            Exit Sub
         End If
      End If
      
      If Chkcriterio(3).Value = 1 Then
         If Txtoritra.text = "" Then
            MsgBox "Falta Ingresar El Codigo De Origen - Transito", 48, "Advertencia"
            Txtoritra.SetFocus
            Exit Sub
         End If
      End If
           
      DIB_GRILLA
      
      If OptTipo(1).Value = True Then
         'detallada
         QRY = "SELECT conumero, cocodrut, rudestino, cocodemp, henombre, to_char(cofecsal,'MON/DD/YYYY'), to_char(cohorsal,'MON/DD/YYYY HH24:MI'), " _
                    & "coplaca, conumint, coclase, coconduc, cooritra, cofuncio, coturno, comodo, copasaje, covalor, coestado, covalalc, " _
                    & "coparque, conomcaj, cocredit, codirecc, copeaje, cocedula, NVL(cocedula1,' ') AS cocedula1, NVL(coconduc1,' ') AS coconduc1 " _
               & "FROM coconduc, opruta, ophojemp " _
              & "WHERE cofecsal BETWEEN  to_date('" & Fecha & "','MM/DD/YYYY') " _
                & "AND  to_date('" & Fecha1 & "','MM/DD/YYYY') " _
                & "AND cocodrut = rucodigo " _
                & "AND cocodemp = hecodigo "
      ElseIf OptTipo(0).Value = True Then
         'Totalizada
         QRY = "SELECT COUNT(conumero), SUM(covalor), SUM(covalalc), SUM(coparque), SUM(copasaje) " _
               & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
              & "WHERE cofecsal BETWEEN  to_date('" & Fecha & "','MM/DD/YYYY') " _
                & "AND  to_date('" & Fecha1 & "','MM/DD/YYYY') "
      End If
      
      If Cmbcaja.text <> "" Then
         QRY = QRY & " AND conomcaj = '" & Cmbcaja.text & "'"
      End If
      
      If CmbTerminal.ListIndex <> 0 Then
         QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
      End If
      
      
      '************************ CUANDO SE BUSCAR POR CODIGO DE EMPRESA ************************************
      If Chkcriterio(0).Value = 1 Then
         QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
      End If
     
      '************************ CUANDO SE BUSCA POR CODIGO DE RUTA ****************************************
      If Chkcriterio(1).Value = 1 Then
         QRY = QRY & "AND cocodrut = '" & Txtcodrut.text & "' "
      End If
      
      '************************ CUANDO SE BUSCA POR TIPO DE SERVICIO ***************************************
      If Chkcriterio(2).Value = 1 Then
         QRY = QRY & "AND coclase = '" & Txttipser.text & "' "
      End If
            
      '*********************** CUANDO SE BUSCA POR ORIGEN TRANSITO *****************************************
      If Chkcriterio(3).Value = 1 Then
         QRY = QRY & "AND cooritra = '" & Txtoritra.text & "' "
      End If
            
      If OptAlistamiento(0).Value = True Then
         QRY = QRY & " AND conumalistamiento <> '0' "
      ElseIf OptAlistamiento(1).Value = True Then
         QRY = QRY & " AND conumalistamiento = '0' "
      End If
            
      If OptTipoTasa(0).Value = True Then
         'cupo
         QRY = QRY & "AND comodo = 'P' "
      ElseIf OptTipoTasa(1).Value = True Then
         'vendido
         QRY = QRY & "AND comodo = 'D' "
      End If
      
      If OptTipoTasa(3).Value = True Then
         QRY = QRY & "AND cooritra = 'SV' "
      Else
         QRY = QRY & "AND cooritra <> 'SV' "
      End If
      
      'WECC 07/15/2019
      'CAU2 1258 se pueda consultar solo los que se le entregaron los códigos de la super
      '----------------------------------------------------------------
      If OptCodSuper(0).Value = True Then
         'cupo
         QRY = QRY & "AND copeaje = '0' "
      ElseIf OptCodSuper(1).Value = True Then
         'vendido
         QRY = QRY & "AND copeaje <> '0' "
      End If
      '----------------------------------------------------------------
      
      'WECC 09/23/2020
      'CAU2 3056 por tipo de ruta Intermunicipal (I) o Sabanera (S)
      '----------------------------------------------------------------
      If CmbTipoRuta.ListIndex = 1 Then
         'Intermunicipal
         QRY = QRY & "AND rutipo = 'I' "
      ElseIf CmbTipoRuta.ListIndex = 2 Then
         'Sabaneras
         QRY = QRY & "AND rutipo  = 'S' "
      End If
      '----------------------------------------------------------------
      
      If OptTipo(1).Value = True Then
         QRY = QRY & "AND coestado = 'Activo' " _
              & "ORDER BY conumero ASC "
      Else
         QRY = QRY & "AND coestado = 'Activo'"
         
         QRY1 = QRY
      End If
      
      If OptTipo(1).Value = True Then
      'detallada
      
         Set Tbl1 = Dbs.Execute(QRY)
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
         dato1 = 0
         Grdconduces.Visible = False
         If Optgenerar(1).Value = True Then
            Conconcriterio.Enabled = False
            
            Do While Tbl1.EOF = False
               DoEvents
               If IsNull(Tbl1(18)) Then
                  valalc = "0"
               Else
                  valalc = Trim(Tbl1(18))
               End If
               
               If IsNull(Tbl1(19)) Then
                  Parquea = "0"
               Else
                  Parquea = Trim(Tbl1(19))
               End If
               
               If IsNull(Tbl1(15)) Then
                  dt15 = "0"
               Else
                  dt15 = Trim(Tbl1(15))
               End If
               
               If Tbl1(22) = "O" Then
                    sdireccion = "Oriente"
               ElseIf Tbl1(22) = "C" Then
                    sdireccion = "Occidente"
               ElseIf Tbl1(22) = "S" Then
                    sdireccion = "Sur"
               ElseIf Tbl1(22) = "N" Then
                    sdireccion = "Norte"
               End If
               
               LPeaje = "0"
               If Not IsNull(Tbl1!copeaje) Then
                  LPeaje = Tbl1!copeaje
               End If
               
               If L_DirCristal = "N" Then
                  entry = str(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(9) + _
                          Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + Tbl1(5) + Chr(9) + _
                          Tbl1(6) + Chr(9) + Tbl1(7) + Chr(9) + str(Tbl1(8)) + Chr(9) + _
                          Tbl1(9) + Chr(9) + Tbl1!cocedula + Chr(9) + Tbl1!coconduc + Chr(9) + _
                          Tbl1!cocedula1 + Chr(9) + Tbl1!coconduc1 + Chr(9) + Tbl1(11) + Chr(9) + _
                          Tbl1(12) + Chr(9) + str(Tbl1(13)) + Chr(9) + Tbl1(14) + Chr(9) + _
                          str(dt15) + Chr(9) + str(Tbl1(16)) + Chr(9) + _
                          str(valalc) + Chr(9) + Tbl1(17) + Chr(9) + str(Parquea) + Chr(9) + _
                          Tbl1(20) + Chr(9) + Tbl1(21) + Chr(9) + sdireccion + Chr(9) + CStr(LPeaje)
                          
                  Grdconduces.AddItem entry
               End If
               'ElseIf L_DirCristal = "S" Then
                  QRY = "INSERT INTO optempo1 " _
                     & "(temnumero1, temtext9, temtext2, temdate1,  " _
                     & "temtext8, temtext4, temnumero3, temtext5,  " _
                     & "temtext6, temnumero2, temnumdob3, temnumdob2,  " _
                     & "temtext7, temtext1, temnumdob4, temnumero6,  " _
                     & "temtext11, tempc, temtext13, temtext14, temtext16,temtext17, temtext18, temtext19) VALUES (" _
                     & Tbl1(0) & ",'" & Tbl1(2) & "','" & Tbl1(4) & "',to_date('" & Tbl1(5) & "','MM/DD/YYYY'),'" _
                     & Format(Tbl1(6), "MMM/dd/yyyy HH:mm") & "','" & Tbl1(7) & "'," & Tbl1(8) & ",'" & Tbl1(11) & "','" _
                     & sdireccion & "'," & Tbl1(13) & "," & dt15 & "," & Tbl1(16) & ",'" _
                     & Tbl1(17) & "','" & Tbl1(9) & "'," & valalc & "," & Parquea & ",'" _
                     & Tbl1(20) & "','" & Pc & "','" & Tbl1(21) & "', '" & LPeaje & "','" & Tbl1!cocedula & "','" & Tbl1!coconduc & "','" & Tbl1!cocedula1 & "','" & Tbl1!coconduc1 & "')"
                  Dbs.Execute QRY
               'End If
               
               ProBar.Value = ProBar.Value + 1
               dato1 = dato1 + CDbl(valalc)
               
               dato4 = dato4 + CDbl(Tbl1(16))
               dato5 = dato5 + CDbl(Parquea)
               Dato6 = Dato6 + CDbl(valalc) + CDbl(Tbl1(16)) + CDbl(Parquea)
               dato9 = dato9 + Tbl1(15)
               If Tbl1(21) = "S" Then
                  Dato7 = Dato7 + Tbl1(16)
               End If
               Tbl1.MoveNext
               bandera = 1
            Loop
            Tbl1.Close
            Msktotcon.text = Dato6
            Mskalcoholemia.text = dato1
            msktasas.text = dato4
            If IsNull(dato5) Then
                dato5 = 0
            End If
            Mskparqueadero.text = dato5
            Lblconven.Caption = Grdconduces.rows - 1
            If IsNull(dato9) Then
                dato9 = 0
            End If
            LblPasajeros.Caption = dato9
            If Dato7 <> 0 Then
              mskcredito.text = Dato7
            Else
               mskcredito.text = "0"
            End If
         ElseIf Optgenerar(0).Value = True Then
            respuesta = InputBox("Digite un nombre para el plano", "GENERAR", "TASAS DE USO")
            NOMDIR = App.Path & "\" & respuesta & ".TXT"
             
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
            
               Print #2, "01|"; Now; "|REPORTE TASAS DE USO POR CRITERIO|"; Chr$(13)
               Print #2, "01|FECHA INICIAL :|"; Trim(Mskfecini.FormattedText); "|FECHA FINAL: |"; Trim(Mskfecfin.FormattedText); "|"; Chr$(13)
               Print #2, "HIDER|Numero|Cod.Ruta|Destino|Cod.Emp|Nombre|Fec.Sal|Hor.Sal|Num.Plac|Num.Interno|Tip. Servicio|Cedula1|Conductor1|Cedula2|Conductor2|Ori/Tra|Funcionario|Turno|Modo|Pasajeros|Valor|Alcoholemia|Parqueadero|Estado|Caja|Peaje|"
         
               If Not Tbl1.EOF Then
                  Tbl1.MoveLast
                  ProBar.Max = Tbl1.RecordCount + 2
                  Tbl1.MoveFirst
               End If
               Do While Tbl1.EOF = False
                  DoEvents
                  QRY1 = "SELECT hvnumint " _
                         & "FROM ophojveh " _
                        & "WHERE hvplaca  = '" & Trim(Tbl1!coplaca) & "'"
                        
                  'Set tbl2 = dbs.OpenRecordset(qry1, dbOpenSnapshot)
                        
                  Set Tbl2 = Dbs.Execute(QRY1)
                  
                  LPeaje = "0"
                  If Not IsNull(Tbl1!copeaje) Then
                     LPeaje = Tbl1!copeaje
                  End If
                  
                  a = "02|" & Trim(Tbl1!conumero) & "|" & Trim(Tbl1!cocodrut) & "|" & _
                      Trim(Tbl1!rudestino) & "|" & Trim(Tbl1!cocodemp) & "|" & Trim(Tbl1!henombre) & "|" & _
                      Trim(Tbl1(5)) & "|" & Trim(Tbl1(6)) & "|" & Trim(Tbl1!coplaca) & "|" & _
                      Trim(Tbl2!hvnumint) & "|" & Trim(Tbl1!coclase) & "|" & Trim(Tbl1!cocedula) & "|" & _
                      Trim(Tbl1!coconduc) & "|" & Trim(Tbl1!cocedula1) & "|" & Trim(Tbl1!coconduc1) & "|" & Trim(Tbl1!Cooritra) & "|" & _
                      Trim(Tbl1!cofuncio) & "|" & Trim(Tbl1!coturno) & "|" & Trim(Tbl1!comodo) & "|" & _
                      Trim(Tbl1!copasaje) & "|" & Trim(Tbl1!covalor) & "|" & Trim(Tbl1!covalalc) & "|" & _
                      Trim(Tbl1!coparque) & "|" & Trim(Tbl1!coestado) & "|" & Trim(Tbl1!conomcaj) & "|" & Trim(LPeaje) & "|" & Chr(13)
                  Print #2, a
                  con = con + 1
                  total = total + CDbl(Tbl1!covalor)
                  ProBar.Value = ProBar.Value + 1
                  bandera = 1
                  Tbl1.MoveNext
               Loop
               Print #2, "03|NUMERO DE REGISTROS|"; Trim(con); "|VALOR TOTAL TASAS DE USO|"; Trim(total); "|"
               Close #2
               Ftb01.Close
               MsgBox "Plano Generado con el nombre de : " & NOMDIR, 64, "O.K."
            End If
         End If
      ElseIf OptTipo(0).Value = True Then
         'totalizada
         
         Set Tbl1 = Dbs.Execute(QRY)
            
         If Not Tbl1.EOF Then
            If Not IsNull(Tbl1(1)) Then
               Mskalcoholemia.text = Tbl1(2)
               msktasas.text = Tbl1(1)
               Mskparqueadero.text = Tbl1(3)
               Lblconven.Caption = Tbl1(0)
               LblPasajeros.Caption = Tbl1(4)
               bandera = 1
               
            Else
               Mskalcoholemia.text = 0
               msktasas.text = 0
               Mskparqueadero.text = 0
               Lblconven.Caption = 0
               LblPasajeros.Caption = 0
            End If
         End If
            
         Tbl1.Close
        
         
         QRY = QRY1 & "AND cocredit = 'S'"
         
         Set Tbl1 = Dbs.Execute(QRY)
         
         
         If Not Tbl1.EOF Then
            If Not IsNull(Tbl1(1)) Then
               mskcredito.text = Tbl1(1)
               bandera = 1
            End If
            
         End If
         Tbl1.Close

         If bandera = 1 Then
            Msktotcon.text = CDbl(Mskalcoholemia.text) + CDbl(Mskparqueadero.text) + CDbl(msktasas.text)
         End If
      
      End If
         
   End If
   Grdconduces.Visible = True
   Conconcriterio.Enabled = True
   ProBar.Value = ProBar.Max
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 48, "Advertencia"
      Exit Sub
   End If
    
   Call Aplicar_Perfil(Me, "Conconcriterio")
   If Btnimprimir.Enabled = True Then
      If L_DirCristal = "S" Then
        ImpReporte
      End If
   End If
   Exit Sub
  
terminar:
   Grdconduces.Visible = True
   Conconcriterio.Enabled = True
   MsgBox "Error al tratar de mostrar los registros o error de base de datos " & Chr(13) + Chr(13) & "ERROR >> " & Err.Description, 48, "Advertencia"
   Exit Sub
   Resume
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo err_mensaje
   
  If Grdconduces.rows > 1 Then
'     lineas = Grdconduces.Rows - 1
'     Grdconduces.Row = 1
'      ProBar.Value = ProBar.Min
'      ProBar.Max = Grdconduces.Rows + 2
'      Grdconduces.Visible = False
'      Conconcriterio.Enabled = False
'     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
'
'
'     Dbs.BeginTrans
'     On Error GoTo Err_Trans
'     Do
'         DoEvents
'        ReDim dato(1, 24)
'
'        For I = 0 To 23
'            Grdconduces.Col = I
'            dato(1, I) = Trim(Grdconduces.Text)
'        Next I
'
'        ' insertamos en la tabla
'
'        Qry = "INSERT INTO optempo1 " _
'        & "(temnumero1, temtext9, temtext2, temdate1,  " _
'        & "temtext8, temtext4, temnumero3, temtext5,  " _
'        & "temtext6, temnumero2, temnumdob3, temnumdob2,  " _
'        & "temtext7, temtext1, temnumdob4, temnumero6,  " _
'        & "temtext11, tempc, temtext13, temtext14) VALUES ('" _
'        & dato(1, 0) & "','" & dato(1, 2) & "','" & dato(1, 4) & "',to_date('" & dato(1, 5) & "','MM/DD/YYYY'),'" _
'        & Format(dato(1, 6), "MM/dd/yyyy HH:mm") & "','" & dato(1, 7) & "','" & dato(1, 8) & "','" & dato(1, 11) & "','" _
'        & dato(1, 22) & "','" & dato(1, 13) & "','" & dato(1, 15) & "','" & dato(1, 16) & "','" _
'        & dato(1, 18) & "','" & dato(1, 9) & "','" & dato(1, 17) & "','" & dato(1, 19) & "','" _
'        & dato(1, 20) & "','" & Pc & "','" & dato(1, 21) & "','" & dato(1, 23) & "')"
'
'        Dbs.Execute Qry
'
'        ' fin insertar
'
'        lineas = lineas - 1
'        ProBar.Value = ProBar.Value + 1
'        If lineas <> 0 Then
'           Grdconduces.Row = Grdconduces.Row + 1
'        End If
'     Loop Until lineas = 0  'do_while que se mueve x la grilla
'
'     Dbs.CommitTrans
'     On Error GoTo err_mensaje
'     Grdconduces.Visible = True
'     ProBar.Value = ProBar.Max
'
     ImpReporte
     
     
  End If
  
  Exit Sub
'Err_Trans:
'   Grdconduces.Visible = True
'   Conconcriterio.Enabled = True
'   Dbs.RollbackTrans
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Sub
'
err_mensaje:
   Grdconduces.Visible = True
   Conconcriterio.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   

End Sub

Private Sub Btnlimpiar_Click()
  Txtnumini.text = ""
  Txtnumfin.text = ""
  Mskfecini.Mask = ""
  Mskfecini.text = ""
  Mskfecini.Mask = "##/##/####"
  Mskfecfin.Mask = ""
  Mskfecfin.text = ""
  Mskfecfin.Mask = "##/##/####"
  For i = 0 To 3
     Chkcriterio(i).Value = 0
  Next i
  Grdconduces.rows = 1
  Grdconduces.Cols = 2
  Grdconduces.Clear
  Optbuscar(0).SetFocus
  Msktotcon.text = ""
  Lblconven.Caption = ""
  Mskalcoholemia.text = ""
  msktasas.text = ""
  Mskparqueadero.text = ""
  Optgenerar(1).Value = True
  mskcredito.text = ""
  LblPasajeros.Caption = ""
  Cmbcaja.ListIndex = -1
  CmbTerminal.ListIndex = 0
  CmbTipoRuta.ListIndex = 0
  CmbTerminal_LostFocus
  ProBar.Value = ProBar.Min
  OptAlistamiento(2).Value = True
  OptTipoTasa(2).Value = True
  OptCodSuper(2).Value = True
  DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub


Private Sub Btntotalizar_Click()
'  If Grdconduces.Text <> "" And Grdconduces.Rows > 1 Then
'     dato1 = 0
'     Dato6 = 0
'     dato2 = 0
'     dato5 = 0
'
'     Dato7 = 0
'     For I = 1 To Grdconduces.Rows - 1
'         Grdconduces.Row = I
'         Grdconduces.Col = 11
'         Grdconduces.Col = 16
'         dato1 = dato1 + CDbl(Grdconduces.Text)
'         Grdconduces.Col = 17
'         dato2 = dato2 + CDbl(Grdconduces.Text)
'         Grdconduces.Col = 19
'         dato5 = dato5 + CDbl(Grdconduces.Text)
'         Grdconduces.Col = 15
'         Dato7 = Dato7 + CDbl(Grdconduces.Text)
'         Grdconduces.Col = 21
'         If Grdconduces.Text = "S" Then
'            Grdconduces.Col = 16
'            Dato6 = Dato6 + CDbl(Grdconduces.Text)
'         End If
'
'     Next I
'     MskTasas.Text = dato1
'     mskalcoholemia.Text = dato2
'     Mskparqueadero.Text = dato5
'     LblPasajeros.Caption = Dato7
'     If Dato6 <> 0 Then
'        mskcredito.Text = Dato6
'     Else
'        mskcredito.Text = "0"
'     End If
'     Msktotcon.Text = dato1 + dato2 + dato3 + dato4 + dato5 + Dato7
'     Lblconven.Caption = Grdconduces.Rows - 1
'  End If
End Sub


Private Sub Chkcriterio_Click(index As Integer)
  If Chkcriterio(0).Value = 1 Then
     Txtcodemp.Enabled = True
  ElseIf Chkcriterio(0).Value = 0 Then
     Txtcodemp.Enabled = False
     Txtcodemp.text = ""
     Lblcodemp.Caption = ""
  End If
  
  If Chkcriterio(1).Value = 1 Then
     Txtcodrut.Enabled = True
  ElseIf Chkcriterio(1).Value = 0 Then
     Txtcodrut.Enabled = False
     Txtcodrut.text = ""
     Lblcodrut.Caption = ""
  End If
  
  If Chkcriterio(2).Value = 1 Then
     Txttipser.Enabled = True
  ElseIf Chkcriterio(2).Value = 0 Then
     Txttipser.Enabled = False
     Txttipser.text = ""
     Lbltipser.Caption = ""
  End If
  
  If Chkcriterio(3).Value = 1 Then
     Txtoritra.Enabled = True
  ElseIf Chkcriterio(3).Value = 0 Then
     Txtoritra.Enabled = False
     Txtoritra.text = ""
     Lbloritra.Caption = ""
  End If
  
End Sub

Private Sub CmbTerminal_LostFocus()
   On Error GoTo Errores
   
   Cmbcaja.Clear
   QRY = "SELECT canomcaj " _
         & "FROM cocaja "
   
   If CmbTerminal.ListIndex <> 0 Then
      QRY = QRY & "WHERE catipo = '" & Mid(CmbTerminal.text, 1, 1) & "' "
   End If
   
   QRY = QRY & "ORDER BY canomcaj "
            
   Set Ftb01 = Dbs.Execute(QRY)
      
   Cmbcaja.AddItem ""
   Do While Not Ftb01.EOF
      Cmbcaja.AddItem Ftb01(0)
      Ftb01.MoveNext
   Loop
   Ftb01.Close
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume '
End Sub

Private Sub Form_Activate()
Call Aplicar_Perfil(Me, "Conconcriterio")
End Sub

Private Sub Form_Load()
   DIB_GRILLA
   Optbuscar_Click (index)
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
   CmbTipoRuta.ListIndex = 0
   fqry01 = "SELECT canomcaj " _
            & "FROM cocaja " _
        & "ORDER BY canomcaj "
            
   Set Ftb01 = Dbs.Execute(fqry01)
      
   Cmbcaja.AddItem ""
   Do While Not Ftb01.EOF
      Cmbcaja.AddItem Ftb01(0)
      Ftb01.MoveNext
   Loop
   Ftb01.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Optbuscar_Click(index As Integer)
  If Optbuscar(0).Value = True Then
     Frmnumeros.Visible = True
     Frmfechas.Visible = False
     Txtnumini.TabIndex = 2
     Txtnumfin.TabIndex = 3
  ElseIf Optbuscar(1).Value = True Then
     Frmnumeros.Visible = False
     Frmfechas.Visible = True
     Mskfecini.TabIndex = 2
     Mskfecfin.TabIndex = 3
  End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "15"
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
              
      'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
      
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
         forma = "16"
         forma1 = "15"
         buscar.Show 1
         Txtcodemp.SetFocus
      End If
   End If
End Sub

Private Sub txtcodrut_DblClick()
  forma = "3"
  forma1 = "9"
  Txtcodrut.text = ""
  Lblcodrut.Caption = ""
  buscar.Show 1
  Txtcodrut_LostFocus
End Sub

Private Sub txtcodrut_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtcodrut_LostFocus()
  If Txtcodrut.text <> "" Then
     fqry01 = "SELECT * " _
              & "FROM opruta " _
             & "WHERE rucodigo = '" & Txtcodrut.text & "' "
             
     'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
     
     Set Ftb01 = Dbs.Execute(fqry01)
     
     
     If Not Ftb01.EOF Then
        Lblcodrut.Caption = Ftb01!rudestino
        Ftb01.Close
     Else
        MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
        forma = "3"
        forma1 = "9"
        buscar.Show 1
        Txtcodrut.SetFocus
     End If
  End If
End Sub

Private Sub Txtnumfin_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumini_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
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
         Mskfecini.SelStart = 0
         Mskfecini.SelLength = Len(Mskfecini.text)
         Exit Sub
      End If
  
      If Mskfecini.ClipText <> "" Then
         If Not IsDate(Mskfecini.text) Then
            MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
            Mskfecini.SetFocus
            Mskfecini.SelStart = 0
            Mskfecini.SelLength = Len(Mskfecini.text)
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub Mskfecfin_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecfin_LostFocus()
   If Mskfecfin.ClipText <> "" Then
      If Val(Left(Mskfecfin.text, 2)) < 1 Or Val(Left(Mskfecfin.text, 2)) > 12 Then
         MsgBox "No es un MES VALIDO", 48, "Advertencia"
         Mskfecfin.SetFocus
         Mskfecfin.SelStart = 0
         Mskfecfin.SelLength = Len(Mskfecfin.text)
         Exit Sub
      End If
  
      If Mskfecfin.ClipText <> "" Then
         If Not IsDate(Mskfecfin.text) Then
            MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
            Mskfecfin.SetFocus
            Mskfecfin.SelStart = 0
            Mskfecfin.SelLength = Len(Mskfecfin.text)
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub Txtoritra_DblClick()
  forma = "15"
  forma1 = "17"
  Tabla_Catalogo = "TIPOS DE CONDUCE"
  Txtoritra.text = ""
  Lbloritra.Caption = ""
  buscar.Show 1
  Txtoritra_LostFocus
End Sub

Private Sub Txtoritra_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 79 Or KeyAscii% = 84 Or KeyAscii% = 8 Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtoritra_LostFocus()
  If Txtoritra.text <> "" Then
     fqry01 = "SELECT dscoddet,dsdes,dsest " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & Txtoritra.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'TIPOS DE CONDUCE' "
  
     'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
     
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lbloritra.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        MsgBox "Codigo De Origen - Transito No Existe", 48, "Advertencia"
        forma = "15"
        forma1 = "17"
        Tabla_Catalogo = "TIPOS DE CONDUCE"
        buscar.Show 1
        Txtoritra.SetFocus
     End If
  End If
End Sub

Private Sub Txttipser_DblClick()
  forma = "15"
  forma1 = "4"
  Tabla_Catalogo = "TIPOS DE VEHICULO"
  Txttipser.text = ""
  Lbltipser.Caption = ""
  buscar.Show 1
  Txttipser_LostFocus
End Sub

Private Sub Txttipser_KeyPress(KeyAscii As Integer)
'  Char = Chr(KeyAscii%)
'  KeyAscii = Asc(UCase(Char))  ' mayuscula
'  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'      Exit Sub
'  Else
'     KeyAscii% = 0
'  End If
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub Txttipser_LostFocus()
  If Txttipser.text <> "" Then
     fqry01 = "SELECT dscoddet,dsdes,dsest " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & Txttipser.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'TIPOS DE VEHICULO' "
  
     'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
     
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lbltipser.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        MsgBox "Tipo De Servicio No Existe", 48, "Advertencia"
        forma = "15"
        forma1 = "4"
        Tabla_Catalogo = "TIPOS DE VEHICULO"
        buscar.Show 1
        Txttipser.SetFocus
     End If
  End If
End Sub

Private Function DIB_GRILLA()
     Grdconduces.rows = 1
     Grdconduces.Cols = 27
     Grdconduces.Clear
     
     s$ = "Número  |Cód. Ruta |Destino |Cód. Empresa |Nombre   |Fec. Salida |Hor. Salida  |Núm. Placa  |Núm. Int. |Tip. Servicio |" _
        & "Cedula1|Conductor1 |Cedula2|Conductor2|Origen/Transito |Funcionario |Turno |Modo |Pasajeros |Valor |Alcoholemia |Estado|Parqueadero|Caja |Crédito|Corredor|Peaje"
     Grdconduces.FormatString = s$
     Grdconduces.Row = 0
     Grdconduces.Col = 0
     Grdconduces.ColWidth(0) = 1000
     Grdconduces.ColWidth(1) = 1000
     Grdconduces.ColWidth(2) = 1800
     Grdconduces.ColWidth(3) = 1500
     Grdconduces.ColWidth(4) = 1800
     Grdconduces.ColWidth(5) = 1500
     Grdconduces.ColWidth(6) = 1500
     Grdconduces.ColWidth(7) = 1300
     Grdconduces.ColWidth(8) = 1300
     Grdconduces.ColWidth(9) = 1300
     Grdconduces.ColWidth(10) = 1100
     Grdconduces.ColWidth(11) = 2500
     Grdconduces.ColWidth(12) = 1100
     Grdconduces.ColWidth(13) = 2500
     
     Grdconduces.ColWidth(14) = 1300
     Grdconduces.ColWidth(15) = 1300
     Grdconduces.ColWidth(16) = 1000
     Grdconduces.ColWidth(17) = 1000
     Grdconduces.ColWidth(18) = 1000
     Grdconduces.ColWidth(19) = 1000
     Grdconduces.ColWidth(20) = 1000
     Grdconduces.ColWidth(21) = 1000
     Grdconduces.ColWidth(22) = 1000
     Grdconduces.ColWidth(23) = 1000
     Grdconduces.ColWidth(24) = 1000
     Grdconduces.ColWidth(25) = 1500
     Grdconduces.ColWidth(26) = 1200
     
     For i = 0 To Grdconduces.Cols - 1
         Grdconduces.Col = i
         Grdconduces.CellFontBold = 1
     Next i
End Function
Private Function ImpReporte()
   If Optbuscar(0).Value = True Then
      inicio1 = Txtnumini.text
      final = Txtnumfin.text
   ElseIf Optbuscar(1).Value = True Then
      inicio1 = Mskfecini.FormattedText
      final = Mskfecfin.FormattedText
   End If
   
   imprimir.Show 1
   If G_Imprimir = "0" Then
      Exit Function
   End If
   
   reporte = RpPath + "\conconcr.rpt"
   
   rptterminal.ReportFileName = reporte
   rptterminal.Connect = G_Conexion_reportes
   rptterminal.Formulas(0) = "totgen = '" & Msktotcon.FormattedText & "'"
   rptterminal.Formulas(1) = "conven = '" & Lblconven.Caption & "'"
   rptterminal.Formulas(2) = "inicio1 =  '" & inicio1 & "'"
   rptterminal.Formulas(3) = "final = '" & final & "'"
   rptterminal.Formulas(4) = "empresa = '" & G_NomEmpresa & "'"
   rptterminal.Formulas(6) = "tottasa = '" & msktasas.FormattedText & "'"
   rptterminal.Formulas(7) = "totalc = '" & Mskalcoholemia.FormattedText & "'"
   rptterminal.Formulas(9) = "parque = '" & Mskparqueadero.FormattedText & "'"
   rptterminal.Formulas(10) = "credito = '" & mskcredito.FormattedText & "'"
   rptterminal.Formulas(11) = "pasajeros = '" & LblPasajeros.Caption & "'"
   Conconcriterio.Enabled = False
   rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
   rptterminal.WindowState = crptMaximized
   rptterminal.PrinterName = G_PrinterName
   rptterminal.PrinterDriver = G_PrinterDriver
   rptterminal.PrinterPort = G_PrinterPort
   rptterminal.Destination = 0
   rptterminal.Action = 1
   Conconcriterio.Enabled = True
   
   


End Function


