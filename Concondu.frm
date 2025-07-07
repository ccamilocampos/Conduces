VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Conconduces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Tasas de Uso"
   ClientHeight    =   9420
   ClientLeft      =   2325
   ClientTop       =   1335
   ClientWidth     =   8160
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   8160
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
      Left            =   2520
      TabIndex        =   76
      Top             =   4080
      Width           =   1455
      Begin VB.ComboBox CmbTipoRuta 
         Height          =   315
         ItemData        =   "Concondu.frx":0000
         Left            =   120
         List            =   "Concondu.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   480
         Width           =   1215
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
      Left            =   5280
      TabIndex        =   75
      Top             =   2760
      Width           =   1695
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
         TabIndex        =   31
         Top             =   480
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
         TabIndex        =   29
         Top             =   750
         Value           =   -1  'True
         Width           =   1095
      End
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
         TabIndex        =   30
         Top             =   240
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
      Left            =   3600
      TabIndex        =   74
      Top             =   2760
      Width           =   1575
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
         TabIndex        =   26
         Top             =   240
         Width           =   1095
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
         TabIndex        =   28
         Top             =   720
         Value           =   -1  'True
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
         TabIndex        =   27
         Top             =   480
         Width           =   1215
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
      Height          =   1095
      Left            =   120
      TabIndex        =   73
      Top             =   2880
      Width           =   3375
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
         Left            =   840
         TabIndex        =   25
         Top             =   750
         Value           =   -1  'True
         Width           =   1695
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
         Left            =   840
         TabIndex        =   24
         Top             =   495
         Width           =   1815
      End
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
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Cajas"
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
      Left            =   5640
      TabIndex        =   72
      Top             =   4080
      Width           =   1335
      Begin VB.ComboBox cmbcaja 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Concondu.frx":0032
         Left            =   120
         List            =   "Concondu.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
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
      Height          =   975
      Left            =   4080
      TabIndex        =   71
      Top             =   4080
      Width           =   1455
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "Concondu.frx":0036
         Left            =   120
         List            =   "Concondu.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   34
         ToolTipText     =   "Tipos de terminales"
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   120
      TabIndex        =   67
      Top             =   1680
      Width           =   2295
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
         Left            =   240
         TabIndex        =   11
         Top             =   430
         Width           =   1215
      End
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
         TabIndex        =   10
         Top             =   200
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Height          =   375
      Left            =   120
      TabIndex        =   66
      Top             =   2400
      Width           =   2295
      Begin VB.CheckBox chkcredito 
         Caption         =   "Créditos"
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
         Left            =   480
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
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
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   120
      TabIndex        =   63
      Top             =   960
      Width           =   2295
      Begin VB.OptionButton Optgenerar 
         Caption         =   "Generar Plano"
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
         Width           =   1695
      End
      Begin VB.OptionButton Optgenerar 
         Caption         =   "Ver por pantalla"
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
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
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
      Height          =   1455
      Left            =   120
      TabIndex        =   52
      Top             =   7680
      Width           =   6855
      Begin MSMask.MaskEdBox Msktotcon 
         Height          =   300
         Left            =   3480
         TabIndex        =   53
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskalcoholemia 
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msktasas 
         Height          =   300
         Left            =   1680
         TabIndex        =   55
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskparqueadero 
         Height          =   300
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskcredito 
         Height          =   300
         Left            =   1800
         TabIndex        =   65
         ToolTipText     =   "Tasas de crédito "
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblPasajeros 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Left            =   5160
         TabIndex        =   69
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label13 
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
         Left            =   5280
         TabIndex        =   68
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Tasas Créd:"
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
         TabIndex        =   64
         Top             =   240
         Width           =   1095
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
         TabIndex        =   62
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Lblconven 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Left            =   3480
         TabIndex        =   60
         Top             =   480
         Width           =   1575
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
         Left            =   3480
         TabIndex        =   59
         Top             =   840
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
         Left            =   3480
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Alcoholimetría:"
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
         TabIndex        =   57
         Top             =   240
         Width           =   1335
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
         Left            =   1680
         TabIndex        =   56
         Top             =   840
         Width           =   975
      End
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
      Left            =   7200
      Picture         =   "Concondu.frx":003A
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   6840
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
      Left            =   7200
      Picture         =   "Concondu.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Limpia la forma actual"
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
      Left            =   7200
      Picture         =   "Concondu.frx":08BE
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Salir de la forma actual"
      Top             =   8520
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
      Left            =   7200
      Picture         =   "Concondu.frx":0D00
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Buscar informacion en la base de datos"
      Top             =   6000
      Width           =   855
   End
   Begin VB.Frame Frmcribus 
      Caption         =   "Rango de Búsqueda:"
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
      TabIndex        =   47
      Top             =   80
      Width           =   2415
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
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
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criterio de Busqueda:"
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
      Height          =   1815
      Left            =   2520
      TabIndex        =   42
      Top             =   960
      Width           =   4455
      Begin VB.OptionButton optcriterio 
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
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   21
         Top             =   1125
         Width           =   735
      End
      Begin VB.OptionButton optcriterio 
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
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   20
         Top             =   885
         Width           =   1815
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Operarios"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Consulta General"
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
         Index           =   7
         Left            =   2280
         TabIndex        =   22
         Top             =   1365
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Código Empresa"
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
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Código de Ruta"
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
         TabIndex        =   14
         Top             =   615
         Width           =   2175
      End
      Begin VB.OptionButton optcriterio 
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   885
         Width           =   1575
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Número Interno"
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
         TabIndex        =   16
         Top             =   1125
         Width           =   1695
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Origen/Transito"
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
         TabIndex        =   17
         Top             =   1365
         Width           =   1935
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Placa del Vehículo"
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
         Index           =   5
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valor de búsqueda:"
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
      TabIndex        =   40
      Top             =   4080
      Width           =   2295
      Begin VB.TextBox Txtvalbus 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Valor de búsqueda con respecto a lo que selecciono en criterio de búsqueda(doble Click)"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Lblvalbus 
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
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   43
      Top             =   5160
      Width           =   6855
      Begin Crystal.CrystalReport rptterminal 
         Left            =   5040
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconcon 
         Height          =   2055
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   70
      Top             =   9240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frmnummul 
      Caption         =   "Números de Tasas de Uso "
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
      Left            =   2640
      TabIndex        =   44
      Top             =   80
      Width           =   4335
      Begin VB.TextBox Txtnumini 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Número inicial de multa por la que quiere buscar"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Txtnumfin 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Número final de multa por la que quiere buscar"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Del Número:"
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
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Al Número:"
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
         Left            =   2760
         TabIndex        =   45
         Top             =   240
         Width           =   1215
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
      Left            =   2640
      TabIndex        =   48
      Top             =   80
      Width           =   4335
      Begin VB.OptionButton Opttipfec 
         Caption         =   "Día"
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
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Opttipfec 
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
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Fecha en la que empieze a buscar"
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
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Fecha hasta donde quiere que busque"
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
         Left            =   240
         TabIndex        =   50
         Top             =   240
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
         Left            =   1680
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   7080
      X2              =   7080
      Y1              =   9360
      Y2              =   0
   End
End
Attribute VB_Name = "Conconduces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btnbuscar_Click()
   On Error GoTo terminar
   Dim L_DirCristal As String
   
   bandera = 0
   Msktotcon.text = ""
   Lblconven.Caption = ""
   Mskalcoholemia.text = ""
   msktasas.text = ""
   Mskparqueadero.text = ""
   mskcredito.text = ""
   LblPasajeros.Caption = ""
   ProBar.Value = ProBar.Min
   Grdconcon.Visible = False
   L_DirCristal = "N"
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   If Optbuscar(0).Value = True Then
      If Txtnumini.text = "" Or Txtnumini.text = "0" Then
         MsgBox "Falta Ingresar El Numero Inicial", 48, "Advertencia"
         Txtnumini.SetFocus
         Exit Sub
      End If
     
      If Txtnumfin.text = "" Or Txtnumfin.text = "0" Then
         MsgBox "Falta Ingresar El Numero Final", 48, "Advertencia"
         Txtnumfin.SetFocus
         Exit Sub
      End If
     
      For i = 0 To 6
         If optcriterio(i).Value = True And Txtvalbus.text = "" Then
            MsgBox "Falta Ingresar El Valor de Busqueda", 48, "Advertencia"
            Txtvalbus.SetFocus
            Exit Sub
         End If
      Next i
          
      If optcriterio(8).Value = True And Txtvalbus.text = "" Then
         MsgBox "Falta Ingresar El Valor de Busqueda", 48, "Advertencia"
         Txtvalbus.SetFocus
         Exit Sub
      End If
     
      If CDbl(Txtnumini.text) > CDbl(Txtnumfin.text) Then
         MsgBox "El Numero Inicial No Puede Ser Mayor Al Numero Final", 48, "Advertencia"
         Txtnumini.SetFocus
         Exit Sub
      End If
     
      DIB_GRILLA
         
     
      If OptTipo(1).Value = True Then
         'detallada
         QRY = "SELECT conumero, cocodrut, rudestino, cocodemp, henombre, " _
                    & "to_char(cofecsal,'MON/DD/YYYY'), to_char(cohorsal,'MON/DD/YYYY HH24:MI'), coplaca, " _
                    & "conumint, coconduc, cooritra, cofuncio, coturno, comodo, copasaje, " _
                    & "covalor, coestado, covalalc, coparque, conomcaj, cocredit, codirecc, " _
                    & "coterminal, copeaje, cocedula, NVL(cocedula1,' ') AS cocedula1, NVL(coconduc1,' ') AS coconduc1 " _
               & "FROM coconduc, opruta, ophojemp " _
              & "WHERE conumero BETWEEN " & Txtnumini.text & " " _
                & "AND " & Txtnumfin.text & " " _
                & "AND cocodrut = rucodigo " _
                & "AND cocodemp = hecodigo "
      ElseIf OptTipo(0).Value = True Then
         'Totalizada
         QRY = "SELECT COUNT(conumero), SUM(covalor), SUM(covalalc), SUM(coparque), SUM(copasaje) " _
               & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo  " _
              & "WHERE conumero BETWEEN " & Txtnumini.text & " " _
                & "AND " & Txtnumfin.text & " "
      End If
      
      If CmbTerminal.ListIndex <> 0 Then
         'ha escogido un terminal
         QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
      End If
      
      
      '******************** CUANDO SOLO BUSCA POR NUMEROS **************************
      If optcriterio(7).Value = True Then
         
      End If
     
      '******************** CUANDO BUSCA CODIGO DE EMPRESA ********************************
      If optcriterio(0).Value = True Then
         QRY = QRY & "AND cocodemp = '" & Txtvalbus.text & "' "
      End If
     
      '******************** CUANDO BUSCA CODIGO DE RUTA ********************************
      If optcriterio(1).Value = True Then
         QRY = QRY & "AND cocodrut = '" & Txtvalbus.text & "' "
      End If
     
      '******************** CUANDO BUSCA POR TURNO *******************************
      If optcriterio(2).Value = True Then
         QRY = QRY & "AND coturno = " & Txtvalbus.text & " "
      End If
     
      '******************** CUANDO BUSCA POR NUMERO INTERNO ***********************
      If optcriterio(3).Value = True Then
         QRY = QRY & "AND conumint = " & Txtvalbus.text & " "
      End If
     
      '******************** CUANDO BUSCA POR ORIGEN/TRANSITO ********************
      If optcriterio(4).Value = True Then
         QRY = QRY & "AND cooritra = '" & Txtvalbus.text & "' "
      End If
     
      '******************** CUANDO BUSCA POR PLACA DEL VEHICULO ******************
      If optcriterio(5).Value = True Then
         QRY = QRY & "AND coplaca = '" & Txtvalbus.text & "' "
      End If
     
      '******************** CUANDO BUSCA CODIGO DE OPERARIO ********************************
      If optcriterio(6).Value = True Then
         QRY = QRY & "AND cofuncio = '" & Txtvalbus.text & "' "
      End If
                    
      '******************** CUANDO BUSCA POR ESTADO DE CONDUCES ********************************
      If optcriterio(8).Value = True Then
         QRY = QRY & "AND coestado LIKE '" & Mid(Lblvalbus.Caption, 1, 1) & LCase(Mid(Lblvalbus.Caption, 2)) & "' "
      End If
                     
      '******************** CUANDO BUSCA POR CAJAS ********************************
      If optcriterio(9).Value = True Then
         QRY = QRY & "AND conomcaj = '" & Txtvalbus.text & "' "
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
         If chkcredito.Value = 1 Then
            QRY = QRY & " AND cocredit = 'S' "
         End If
         
         QRY = QRY & "ORDER BY conumero ASC "
      Else
         
         If chkcredito.Value = 1 Then
            QRY = QRY & " AND cocredit = 'S' "
         End If
         QRY1 = QRY
         
         QRY = QRY & "AND coestado = 'Activo'"
      End If
     
      If OptTipo(1).Value = True Then
            
         
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
               'Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
            End If
         End If
         
         Conconduces.Enabled = False
         
         If Optgenerar(1).Value = True Then
            dato1 = 0
            dato2 = 0
            dato3 = 0
            dato4 = 0
            dato5 = 0
            Dato6 = 0
            Dato7 = 0
            Dato8 = 0
            dato9 = 0
            dato10 = 0
            dato11 = 0
            dato12 = 0
            dato13 = 0
            dato14 = 0
            dato15 = 0
            dato16 = 0
            
            i = 1
            j = 1
            Do While Not Tbl1.EOF
               DoEvents
               If IsNull(Tbl1(17)) Then
                   valalc = "0"
               Else
                  valalc = Trim(str(Tbl1(17)))
               End If
               
               If IsNull(Tbl1(18)) Then
                  Parquea = "0"
               Else
                  Parquea = Trim(str(Tbl1(18)))
               End If
               
               If IsNull(Tbl1(14)) Then
                  dt14 = "0"
               Else
                  dt14 = Trim(str(Tbl1(14)))
               End If
               
               If Tbl1(21) = "O" Then
                    sdireccion = "Oriente"
               ElseIf Tbl1(21) = "C" Then
                    sdireccion = "Occidente"
               ElseIf Tbl1(21) = "S" Then
                    sdireccion = "Sur"
               ElseIf Tbl1(21) = "N" Then
                    sdireccion = "Norte"
               End If
               LPeaje = "0"
               If Not IsNull(Tbl1!copeaje) Then
                  LPeaje = Tbl1!copeaje
               End If
               
               Call Traer_Clase_Puestos(Tbl1!coplaca)
               
               'Sebastián Rondón - Agosto 08 de 2023
                  'Consultamos el número de la factura antes de ingresarla en la grilla
                  fqry01 = "SELECT tfnumfac, tfprefijo FROM cotufe WHERE tfnumtu = " & Tbl1(0) & ""
                  Set Tbl2 = Dbs.Execute(fqry01)
                  
                sPref = ""
                sNumF = ""
                
                If Not Tbl2.EOF Then
                    sPref = Tbl2!tfprefijo
                    sNumF = str(Tbl2!tfnumfac)
                Else
                    sPref = "N/A"
                    sNumF = "N/A"
                End If
               
               If L_DirCristal = "N" Then
                  
                  entry = str(Tbl1(0)) + Chr(9) + sPref + Chr(9) + sNumF + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(9) + _
                          Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + Tbl1(5) + Chr(9) + _
                          Tbl1(6) + Chr(9) + Tbl1(7) + Chr(9) + str(Tbl1(8)) + Chr(9) + _
                          Tbl1!cocedula + Chr(9) + Tbl1!coconduc + Chr(9) + Tbl1!cocedula1 + Chr(9) + Tbl1!coconduc1 + Chr(9) + _
                          Tbl1(10) + Chr(9) + Tbl1(11) + Chr(9) + str(Tbl1(12)) + Chr(9) + Tbl1(13) + Chr(9) + str(dt14) + Chr(9) + _
                          str(Tbl1(15)) + Chr(9) + valalc + Chr(9) + Tbl1(16) + Chr(9) + _
                          Parquea + Chr(9) + Tbl1(19) + Chr(9) + Tbl1(20) + Chr(9) + _
                          sdireccion + Chr(9) + Tbl1!coterminal + Chr(9) + _
                          G_Clase_Vehiculo + Chr(9) + CStr(G_Cantidad_Puestos_Vehiculo) + Chr(9) + CStr(LPeaje)
                  
                  Grdconcon.AddItem entry
               End If
               'ElseIf L_DirCristal = "S" Then
                  QRY = "INSERT INTO optempo1 " _
                  & "(temnumero1, temtext9, temtext2, temdate1, temtext8, " _
                  & " temtext4, temnumero3, temtext5, temtext6, temnumero2, " _
                  & " temtext11, temnumdob3, temnumdob2, temnumdob4, temtext7, " _
                  & " temnumero6, temtext12, tempc, temtext13, temtext1, " _
                  & " temtext3, temtext14, temnumero4, temtext15, temtext16, " _
                  & " temtext17, temtext18, temtext19, temtext10, temtext20) VALUES (" _
                  & Tbl1(0) & ",'" & Tbl1(2) & "','" & Tbl1(4) & "',to_date('" & Tbl1(5) & "','MM/DD/YYYY'),'" & Format(Tbl1(6), "MMM/dd/yyyy HH:mm") & "','" _
                  & Tbl1(7) & "'," & Tbl1(8) & ", '" & Tbl1(10) & "','" & sdireccion & "'," & Tbl1(12) & ",'" _
                  & Tbl1(13) & "'," & dt14 & "," & Tbl1(15) & "," & valalc & ",'" & Tbl1(16) & "'," _
                  & Parquea & ",'" & Tbl1(19) & "','" & Pc & "','" & Tbl1(20) & "','" & Tbl1!coterminal & "','" _
                  & Tbl1!cofuncio & "','" & G_Clase_Vehiculo & "'," & G_Cantidad_Puestos_Vehiculo & ", '" & LPeaje & "','" & Tbl1!cocedula & "', '" _
                  & Tbl1!coconduc & "','" & Tbl1!cocedula1 & "','" & Tbl1!coconduc1 & "', '" & sPref & "', '" & sNumF & "')"
                  
                  Dbs.Execute QRY
               'End If
               ProBar.Value = ProBar.Value + 1
               If Tbl1(16) = "Activo" Then
                  dato1 = dato1 + CDbl(valalc)
                  dato4 = dato4 + CDbl(Tbl1(15))
                  
                  dato5 = dato5 + CDbl(Parquea)
                  Dato6 = Dato6 + CDbl(valalc) + CDbl(Tbl1(15)) + CDbl(Parquea)
                  dato15 = dato15 + Tbl1(14)
                  If Tbl1(20) = "S" Then
                     dato13 = dato13 + CDbl(Tbl1(15))
                  End If
                  i = 1 + i
               ElseIf Tbl1(16) = "Anulado" Then
                  Dato7 = Dato7 + CDbl(valalc)
                  dato10 = dato10 + CDbl(Tbl1(15))
                                
                  dato11 = dato11 + CDbl(Parquea)
                  dato12 = dato12 + CDbl(valalc) + CDbl(Tbl1(15)) + CDbl(Parquea)
                  dato16 = dato16 + Tbl1(14)
                  If Tbl1(20) = "S" Then
                     dato14 = dato14 + CDbl(Tbl1(15))
                  End If
                  j = 1 + j
               ElseIf Tbl1(16) = "Cambio" Then
                  dato1 = dato1 + CDbl(valalc)
                  Dato6 = Dato6 + CDbl(valalc)
               End If
               
               Tbl1.MoveNext
               bandera = 1
            Loop
            Tbl1.Close
            If optcriterio(8).Value = True Then   'cuando se lecciono la opcion estado
               If Lblvalbus.Caption = "ACTIVO" Then
                   Msktotcon.text = Dato6
                   Mskalcoholemia.text = dato1
                   msktasas.text = dato4
                   Mskparqueadero.text = dato5
                   LblPasajeros.Caption = dato15
                   Lblconven.Caption = i - 1
                   If dato13 <> 0 Then
                      mskcredito.text = dato13
                   Else
                      mskcredito.text = "0"
                   End If
                ElseIf Lblvalbus.Caption = "ANULADO" Then
                   Msktotcon.text = dato12
                   Mskalcoholemia.text = Dato7
                   msktasas.text = dato10
                   Mskparqueadero.text = dato11
                   Lblconven.Caption = j - 1
                   If IsNull(dato16) Then
                    dato16 = 0
                   End If
                   LblPasajeros.Caption = dato16
                   If dato14 <> 0 Then
                      mskcredito.text = dato14
                   Else
                      mskcredito.text = "0"
                   End If
                ElseIf Lblvalbus.Caption = "CAMBIO" Then
                   Msktotcon.text = Dato6
                   Mskalcoholemia.text = dato1
                   msktasas.text = dato4
                   Mskparqueadero.text = dato5
                   LblPasajeros.Caption = dato15
                   Lblconven.Caption = i - 1
                   If dato13 <> 0 Then
                      mskcredito.text = dato13
                   Else
                      mskcredito.text = "0"
                   End If
                End If
            
             Else
                Msktotcon.text = Dato6
                Mskalcoholemia.text = dato1
                msktasas.text = dato4
                Mskparqueadero.text = dato5
                Lblconven.Caption = i - 1
                If IsNull(dato15) Then
                    dato15 = 0
                End If
                LblPasajeros.Caption = dato15
                If dato13 <> 0 Then
                   mskcredito.text = dato13
                Else
                   mskcredito.text = "0"
                End If
             End If
              
         ElseIf Optgenerar(0).Value = True Then
            respuesta = InputBox("Digite un nombre para el plano", "GENERAR", "TASAS DE USO")
            If respuesta = "" Then
               Conconduces.Enabled = True
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
               Conconduces.Enabled = True
               Close #2
            Else
            
               Print #2, "01|"; Now; "|REPORTE TASAS DE USO POR CRITERIO|"; Chr$(13)
               Print #2, "01|DEL NUMERO :|"; Trim(Txtnumini.text); "|HASTA EL NUMERO: |"; Trim(Txtnumfin.text); "|"; Chr$(13)
               Print #2, "HIDER|Numero|Cod.Ruta|Destino|Cod.Emp|Nombre|Fec.Sal|Hor.Sal|Num.Plac|Num.Interno|Cedula1|Conductor1|Cedula2|Conductor2|Ori/Tra|Funcionario|Turno|Modo|Pasajeros|Valor|Alcoholemia|Parqueadero|Estado|Caja|Credito|Dirección|Terminal|Clase|Cant. Puestos|Peaje"
               If Not Tbl1.EOF Then
                  Tbl1.MoveLast
                  ProBar.Max = Tbl1.RecordCount + 2
                  Tbl1.MoveFirst
               End If
           
               Do While Not Tbl1.EOF
                  DoEvents
                  QRY1 = "SELECT hvnumint, hvclase, hvnumpas " _
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
                      Trim(Tbl2!hvnumint) & "|" & Trim(Tbl1!cocedula) & "|" & Trim(Tbl1!coconduc) & "|" & _
                      Trim(Tbl1!cocedula1) & "|" & Trim(Tbl1!coconduc1) & "|" & Trim(Tbl1!Cooritra) & "|" & _
                      Trim(Tbl1!cofuncio) & "|" & Trim(Tbl1!coturno) & "|" & Trim(Tbl1!comodo) & "|" & _
                      Trim(Tbl1!copasaje) & "|" & Trim(Tbl1!covalor) & "|" & Trim(Tbl1!covalalc) & "|" & _
                      Trim(Tbl1!coparque) & "|" & Trim(Tbl1!coestado) & "|" & Trim(Tbl1!conomcaj) & "|" & _
                      Trim(Tbl1!cocredit) & "|" & Trim(Tbl1!codirecc) & "|" & Trim(Tbl1!coterminal) & "|" & _
                      Trim(Tbl2!hvclase) & "|" & Trim(Tbl2!hvnumpas) & "|" & Trim(LPeaje) & "|" & Chr(13)
                  Print #2, a
                  con = con + 1
                  total = total + CDbl(Tbl1!covalor)
                  ProBar.Value = ProBar.Value + 1
                  bandera = 1
                  Tbl1.MoveNext
               Loop
               Print #2, "03|NUMERO DE REGISTROS|"; Trim(con); "|VALOR TOTAL TASAS DE USO|"; Trim(total); "|"
               Close #2
               Tbl1.Close
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
               If chkcredito.Value = 1 Then
                  mskcredito.text = Tbl1(1)
               End If
            Else
               Mskalcoholemia.text = 0
               msktasas.text = 0
               Mskparqueadero.text = 0
               Lblconven.Caption = 0
               LblPasajeros.Caption = 0
            End If
         End If
            
         Tbl1.Close
         If chkcredito.Value = 0 Then
         
            QRY = QRY1 & " AND cocredit = 'S' " _
                        & "AND coestado = 'Activo'"
                          
            
            Set Tbl1 = Dbs.Execute(QRY)

            If Not Tbl1.EOF Then
               If Not IsNull(Tbl1(1)) Then
                  mskcredito.text = Tbl1(1)
                  bandera = 1
               End If
               
            End If
            Tbl1.Close
         End If
         If bandera = 1 Then
            Msktotcon.text = CDbl(Mskalcoholemia.text) + CDbl(Mskparqueadero.text) + CDbl(msktasas.text)
         End If
      End If
   End If
  
   If Optbuscar(1).Value = True Then
      Fecha1 = Mskfecini.text
      Fecha2 = Mskfecfin.text
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
      
      For i = 0 To 6
          If optcriterio(i).Value = True And Txtvalbus.text = "" Then
             MsgBox "Falta Ingresar El Valor de Busqueda", 48, "Advertencia"
             Txtvalbus.SetFocus
             Exit Sub
          End If
      Next i
      
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
      
      DIB_GRILLA
     
      If OptTipo(1).Value = True Then
         'detallada
         If Opttipfec(0).Value = True Then
            sCantidadCiclo = 1
            sInicioCiclo = 1
         Else
            If Trim(Cmbcaja.text) = "" Then
               sCantidadCiclo = Cmbcaja.ListCount - 1
               sInicioCiclo = 1
            Else
               sCantidadCiclo = Cmbcaja.ListIndex
               sInicioCiclo = sCantidadCiclo
            End If
         End If
         
         dato1 = 0
         dato2 = 0
         dato3 = 0
         dato4 = 0
         dato5 = 0
         Dato6 = 0
         Dato7 = 0
         Dato8 = 0
         dato9 = 0
         dato10 = 0
         dato11 = 0
         dato12 = 0
         dato13 = 0
         dato14 = 0
         dato15 = 0
         dato16 = 0
         i = 1
         j = 1
         'se hace un ciclo para que recorra todas las cajas si la deja en blando si escoge una solo hace un ciclo
         'y cuando escoge por dia tambien hace un dia
         For k = sInicioCiclo To sCantidadCiclo
            QRY = "SELECT conumero, cocodrut, rudestino, cocodemp, henombre, " _
                       & "to_char(cofecsal,'MON/DD/YYYY'), to_char(cohorsal,'MON/DD/YYYY HH24:MI'), coplaca, " _
                           & "conumint, coconduc, cooritra, cofuncio, coturno, comodo, copasaje, " _
                           & "covalor, coestado, covalalc, coparque, conomcaj, cocredit, codirecc, " _
                           & "coterminal, copeaje, cocedula, NVL(cocedula1,' ') AS cocedula1, NVL(coconduc1,' ') AS coconduc1  " _
              & "FROM coconduc, opruta, ophojemp "
          
            If Opttipfec(0).Value = True Then
            'dia
               QRY = QRY & "WHERE cofecsal BETWEEN to_date('" & Fecha1 & "','MM/DD/YYYY') " _
                           & "AND  to_date('" & Fecha2 & "','MM/DD/YYYY') "
            Else
            'turno
               
               QRY1 = "SELECT MIN(TO_NUMBER(cinumini)), MAX(TO_NUMBER (cinumfin)) " _
                      & "FROM cocietur " _
                     & "WHERE cifecha BETWEEN TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                       & "AND TO_DATE('" & Fecha2 & "','MM/dd/yyyy') " _
                       & "AND cinomcaj = '" & Trim(Cmbcaja.List(k)) & "' AND TO_NUMBER(cinumini) <> 0 "
               
               Set Tbl2 = Dbs.Execute(QRY1)
         
               If Not Tbl2.EOF Then
                  If Not IsNull(Tbl2(0)) Then
                     vmin = Tbl2(0)
                     vmax = Tbl2(1)
                  Else
                     vmin = 0
                     vmax = 0
                  End If
               Else
                  vmin = 0
                  vmax = 0
               End If
               Tbl2.Close
         
               QRY = QRY & "WHERE conumero BETWEEN " & vmin & " " _
                           & "AND " & vmax & " " _
                           & "AND conomcaj = '" & Trim(Cmbcaja.List(k)) & "'"
                     
            End If
               
            QRY = QRY & "AND cocodrut = rucodigo " _
                      & "AND cocodemp = hecodigo "
      
      
            If CmbTerminal.ListIndex <> 0 Then
               'ha escogido un terminal
               QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
            End If
            
            '******************** CUANDO SOLO BUSCA POR NUMEROS **************************
            If optcriterio(7).Value = True Then
               
            End If
            
            '******************** CUANDO BUSCA CODIGO DE EMPRESA ********************************
            If optcriterio(0).Value = True Then
               QRY = QRY & "AND cocodemp = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA CODIGO DE RUTA ********************************
            If optcriterio(1).Value = True Then
               QRY = QRY & "AND cocodrut = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA POR TURNO *******************************
            If optcriterio(2).Value = True Then
               QRY = QRY & "AND coturno = " & Txtvalbus.text & " "
            End If
            
            '******************** CUANDO BUSCA POR NUMERO INTERNO ***********************
            If optcriterio(3).Value = True Then
               QRY = QRY & "AND conumint = " & Txtvalbus.text & " "
            End If
            
            '******************** CUANDO BUSCA POR ORIGEN/TRANSITO ********************
            If optcriterio(4).Value = True Then
               QRY = QRY & "AND cooritra = '" & Txtvalbus.text & "' "
            End If
      
            '******************** CUANDO BUSCA POR PLACA DEL VEHICULO ******************
            If optcriterio(5).Value = True Then
               QRY = QRY & "AND coplaca = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA CODIGO DE OPERARIO ********************************
            If optcriterio(6).Value = True Then
               QRY = QRY & "AND cofuncio = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA POR ESTADO DE CONDUCES ********************************
            If optcriterio(8).Value = True Then
               QRY = QRY & "AND coestado LIKE '" & Mid(Lblvalbus.Caption, 1, 1) & LCase(Mid(Lblvalbus.Caption, 2)) & "' "
            End If
            
            '******************** CUANDO BUSCA POR CAJA ********************************
            If optcriterio(9).Value = True Then
               QRY = QRY & "AND conomcaj = '" & Txtvalbus.text & "' "
            End If
      
            If chkcredito.Value = 1 Then
               QRY = QRY & " AND cocredit = 'S' "
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
            
            QRY = QRY & "ORDER BY conumero ASC "
      
            Set Tbl1 = Dbs.Execute(QRY)
            ProBar.Value = ProBar.Min
            
            If Not Tbl1.EOF Then
               Tbl1.MoveLast
               NumRegistros = Tbl1.RecordCount
               ProBar.Max = Tbl1.RecordCount + 2
               Tbl1.MoveFirst
               
               If Opttipfec(0).Value = True Then
                  If CDbl(G_CantRegist) < CDbl(NumRegistros) Then
                     MsgBox "El número de registros encontrados excede el limite para mostrar " & _
                     Chr(13) & "             en la grilla, se generará directo a pantalla", 48, "Advertencia"
                     L_DirCristal = "S"
                     
                  End If
               ElseIf Opttipfec(1).Value = True Then
                  L_DirCristal = "S"
               End If
            End If
            
            If Optgenerar(1).Value = True Then
               Conconduces.Enabled = False
               Do While Not Tbl1.EOF
                  DoEvents
                  If IsNull(Tbl1(17)) Then
                      valalc = "0"
                  Else
                     valalc = Trim(str(Tbl1(17)))
                  End If
                  
                  If IsNull(Tbl1(18)) Then
                     Parquea = "0"
                  Else
                     Parquea = Trim(str(Tbl1(18)))
                  End If
                  
                  If IsNull(Tbl1(14)) Then
                     dt14 = "0"
                  Else
                     dt14 = Trim(str(Tbl1(14)))
                  End If
                  
                  If Tbl1(21) = "O" Then
                       sdireccion = "Oriente"
                  ElseIf Tbl1(21) = "C" Then
                       sdireccion = "Occidente"
                  ElseIf Tbl1(21) = "S" Then
                       sdireccion = "Sur"
                  ElseIf Tbl1(21) = "N" Then
                       sdireccion = "Norte"
                  End If
                  
                  LPeaje = "0"
                  If Not IsNull(Tbl1!copeaje) Then
                     LPeaje = Tbl1!copeaje
                  End If
                  
                  Call Traer_Clase_Puestos(Tbl1!coplaca)
                  
                  'Sebastián Rondón - Agosto 08 de 2023
                    'Consultamos el número de la factura antes de ingresarla en la grilla
                    fqry01 = "SELECT tfnumfac, tfprefijo FROM cotufe WHERE tfnumtu = " & Tbl1(0) & ""
                    Set Tbl2 = Dbs.Execute(fqry01)
                    
                    sPref = ""
                    sNumF = ""
                    
                    If Not Tbl2.EOF Then
                        sPref = Tbl2!tfprefijo
                        sNumF = str(Tbl2!tfnumfac)
                    Else
                        sPref = "N/A"
                        sNumF = "N/A"
                    End If
                  
                  If L_DirCristal = "N" Then
                  
                     entry = str(Tbl1(0)) + Chr(9) + sPref + Chr(9) + sNumF + Chr(9) + Tbl1(1) + Chr(9) + _
                             Tbl1(2) + Chr(9) + Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + _
                             Tbl1(5) + Chr(9) + Tbl1(6) + Chr(9) + _
                             Tbl1(7) + Chr(9) + str(Tbl1(8)) + Chr(9) + _
                             Tbl1!cocedula + Chr(9) + Tbl1!coconduc + Chr(9) + Tbl1!cocedula1 + Chr(9) + Tbl1!coconduc1 + Chr(9) + _
                             Tbl1(10) + Chr(9) + Tbl1(11) + Chr(9) + str(Tbl1(12)) + Chr(9) + Tbl1(13) + Chr(9) + _
                             str(dt14) + Chr(9) + str(Tbl1(15)) + Chr(9) + valalc + Chr(9) + _
                             Tbl1(16) + Chr(9) + Parquea + Chr(9) + Tbl1(19) + Chr(9) + _
                             Tbl1(20) + Chr(9) + sdireccion + Chr(9) + Tbl1!coterminal + Chr(9) + _
                             G_Clase_Vehiculo + Chr(9) + CStr(G_Cantidad_Puestos_Vehiculo) + Chr(9) + CStr(LPeaje)
                  
                     Grdconcon.AddItem entry
                  End If
                  'ElseIf L_DirCristal = "S" Then
                     QRY = "INSERT INTO optempo1 " _
                     & "(temnumero1, temtext9, temtext2, temdate1, temtext8, " _
                     & " temtext4, temnumero3, temtext5, temtext6, temnumero2, " _
                     & " temtext11, temnumdob3, temnumdob2, temnumdob4, temtext7, " _
                     & " temnumero6, temtext12, tempc, temtext13, temtext1, " _
                     & " temtext3, temtext14, temnumero4, temtext15, temtext16, " _
                     & " temtext17, temtext18, temtext19, temtext10, temtext20) VALUES (" _
                     & Tbl1(0) & ",'" & Tbl1(2) & "','" & Tbl1(4) & "',to_date('" & Tbl1(5) & "','MM/DD/YYYY'),'" & Format(Tbl1(6), "MMM/dd/yyyy HH:mm") & "','" _
                     & Tbl1(7) & "'," & Tbl1(8) & ", '" & Tbl1(10) & "','" & sdireccion & "'," & Tbl1(12) & ",'" _
                     & Tbl1(13) & "'," & dt14 & "," & Tbl1(15) & "," & valalc & ",'" & Tbl1(16) & "'," _
                     & Parquea & ",'" & Tbl1(19) & "','" & Pc & "','" & Tbl1(20) & "','" & Tbl1!coterminal & "','" _
                     & Tbl1!cofuncio & "','" & G_Clase_Vehiculo & "'," & G_Cantidad_Puestos_Vehiculo & ", '" & LPeaje & "','" & Tbl1!cocedula & "', '" _
                     & Tbl1!coconduc & "','" & Tbl1!cocedula1 & "','" & Tbl1!coconduc1 & "', '" & sPref & "', '" & sNumF & "')"
                     
                     Dbs.Execute QRY
                  'End If
                  
                  ProBar.Value = ProBar.Value + 1
                  If Tbl1(16) = "Activo" Then
                     dato1 = dato1 + CDbl(valalc)
                     dato4 = dato4 + CDbl(Tbl1(15))
                     dato5 = dato5 + CDbl(Parquea)
                     Dato6 = Dato6 + CDbl(valalc) + CDbl(Tbl1(15)) + CDbl(Parquea)
                     dato15 = dato15 + Tbl1(14)
                     If Tbl1(20) = "S" Then
                        dato13 = dato13 + CDbl(Tbl1(15))
                     End If
                     i = 1 + i
                  Else
                     Dato7 = Dato7 + CDbl(valalc)
                     dato10 = dato10 + CDbl(Tbl1(15))
                     dato11 = dato11 + CDbl(Parquea)
                     dato16 = dato16 + Tbl1(14)
                     If Tbl1(20) = "S" Then
                        dato14 = dato14 + CDbl(Tbl1(15))
                     End If
                     dato12 = dato12 + CDbl(valalc) + CDbl(Tbl1(15)) + CDbl(Parquea)
                     j = 1 + j
                  End If
                  
                  Tbl1.MoveNext
                  bandera = 1
               Loop
               Tbl1.Close
               If optcriterio(8).Value = True Then   'cuando se lecciono la opcion estado
                  If Lblvalbus.Caption = "ACTIVO" Then
                     Msktotcon.text = Dato6
                      Mskalcoholemia.text = dato1
                      msktasas.text = dato4
                      Mskparqueadero.text = dato5
                      Lblconven.Caption = i - 1
                      If IsNull(dato15) Then
                       dato15 = 0
                      End If
                      LblPasajeros.Caption = dato15
                      If dato13 <> 0 Then
                         mskcredito.text = dato13
                      Else
                         mskcredito.text = "0"
                      End If
                  ElseIf Lblvalbus.Caption = "ANULADO" Then
                   
                     Msktotcon.text = dato12
                     Mskalcoholemia.text = Dato7
                     msktasas.text = dato10
                     Mskparqueadero.text = dato11
                   
                     If IsNull(dato16) Then
                      dato16 = 0
                     End If
                     LblPasajeros.Caption = dato16
                     Lblconven.Caption = j - 1
                     If dato14 <> 0 Then
                        mskcredito.text = dato14
                     Else
                        mskcredito.text = "0"
                     End If
                  End If
               Else
                  Msktotcon.text = Dato6
                  Mskalcoholemia.text = dato1
                  msktasas.text = dato4
                  Mskparqueadero.text = dato5
                  If IsNull(dato15) Then
                      dato15 = 0
                  End If
                  LblPasajeros.Caption = dato15
                  Lblconven.Caption = i - 1
                  If dato13 <> 0 Then
                     mskcredito.text = dato13
                  Else
                     mskcredito.text = "0"
                  End If
               End If
               LblPasajeros.Caption = dato15
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
                  Print #2, "01|FECHA INICIAL :|"; Trim(Mskfecini.FormattedText); "|FECHA FINAL : |"; Trim(Mskfecfin.FormattedText); "|"; Chr$(13)
                  Print #2, "HIDER|Numero|Cod.Ruta|Destino|Cod.Emp|Nombre|Fec.Sal|Hor.Sal|Num.Plac|Num.Interno|Cedula|Conductor|Cedula2|Conductor2|Ori/Tra|Funcionario|Turno|Modo|Pasajeros|Valor|Alcoholemia|Parqueadero|Estado|Caja|Credito|Dirección|Terminal|Clase|Cant. Puestos|Peaje"
                  If Not Tbl1.EOF Then
                     Tbl1.MoveLast
                     ProBar.Max = Tbl1.RecordCount + 2
                     Tbl1.MoveFirst
                  End If
           
                  Do While Not Tbl1.EOF
                     DoEvents
                     
                     QRY1 = "SELECT hvnumint, hvclase, hvnumpas " _
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
                         Trim(Tbl2!hvnumint) & "|" & Trim(Tbl1!cocedula) & "|" & Trim(Tbl1!coconduc) & "|" & _
                         Trim(Tbl1!cocedula1) & "|" & Trim(Tbl1!coconduc1) & "|" & Trim(Tbl1!Cooritra) & "|" & _
                         Trim(Tbl1!cofuncio) & "|" & Trim(Tbl1!coturno) & "|" & Trim(Tbl1!comodo) & "|" & _
                         Trim(Tbl1!copasaje) & "|" & Trim(Tbl1!covalor) & "|" & Trim(Tbl1!covalalc) & "|" & _
                         Trim(Tbl1!coparque) & "|" & Trim(Tbl1!coestado) & "|" & Trim(Tbl1!conomcaj) & "|" & _
                         Trim(Tbl1!cocredit) & "|" & Trim(Tbl1!codirecc) & "|" & Trim(Tbl1!coterminal) & "|" & _
                         Trim(Tbl2!hvclase) & "|" & Trim(Tbl2!hvnumpas) & "|" & Trim(LPeaje) & "|" & Chr(13)
                     Print #2, a
                     con = con + 1
                     total = total + CDbl(Tbl1!covalor)
                     ProBar.Value = ProBar.Value + 1
                     bandera = 1
                     Tbl1.MoveNext
                  Loop
                  Print #2, "03|NUMERO DE REGISTROS|"; Trim(con); "|VALOR TOTAL TASAS DE USO|"; Trim(total); "|"
                  Close #2
                  Tbl1.Close
                  MsgBox "Plano Generado con el nombre de : " & NOMDIR, 64, "O.K."
               End If
            End If
         Next k
      ElseIf OptTipo(0).Value = True Then
         'totalizada
         
         If Opttipfec(0).Value = True Then
            sCantidadCiclo = 1
            sInicioCiclo = 1
         Else
            If Trim(Cmbcaja.text) = "" Then
               sCantidadCiclo = Cmbcaja.ListCount - 1
               sInicioCiclo = 1
            Else
               sCantidadCiclo = Cmbcaja.ListIndex
               sInicioCiclo = sCantidadCiclo
            End If
         End If
         Mskalcoholemia.text = 0
         msktasas.text = 0
         Mskparqueadero.text = 0
         Lblconven.Caption = 0
         LblPasajeros.Caption = 0
         mskcredito.text = 0
         For k = sInicioCiclo To sCantidadCiclo
            QRY = "SELECT COUNT(conumero), SUM(covalor), SUM(covalalc), SUM(coparque), SUM(copasaje) " _
                  & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo "
             
             
            If Opttipfec(0).Value = True Then
            'dia
               QRY = QRY & "WHERE cofecsal BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
                           & "AND  TO_DATE('" & Fecha2 & "','MM/DD/YYYY') "
            Else
            'turno
               
               QRY1 = "SELECT MIN(TO_NUMBER(cinumini)), MAX(TO_NUMBER (cinumfin)) " _
                      & "FROM cocietur " _
                     & "WHERE cifecha BETWEEN TO_DATE('" & Fecha1 & "','MM/dd/yyyy') " _
                       & "AND TO_DATE('" & Fecha2 & "','MM/dd/yyyy') " _
                       & "AND cinomcaj = '" & Trim(Cmbcaja.List(k)) & "' AND TO_NUMBER(cinumini) <> 0 "
               
               Set Tbl2 = Dbs.Execute(QRY1)
         
               If Not Tbl2.EOF Then
                  If Not IsNull(Tbl2(0)) Then
                     vmin = Tbl2(0)
                     vmax = Tbl2(1)
                  Else
                     vmin = 0
                     vmax = 0
                  End If
               Else
                  vmin = 0
                  vmax = 0
               End If
               Tbl2.Close
         
               QRY = QRY & "WHERE conumero BETWEEN " & vmin & " " _
                           & "AND " & vmax & " " _
                           & "AND conomcaj = '" & Trim(Cmbcaja.List(k)) & "'"
                     
            End If
                   
            If CmbTerminal.ListIndex <> 0 Then
               'ha escogido un terminal
               QRY = QRY & " AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "'"
            End If
            
            
            '******************** CUANDO SOLO BUSCA POR NUMEROS **************************
            If optcriterio(7).Value = True Then
               
            End If
            
            '******************** CUANDO BUSCA CODIGO DE EMPRESA ********************************
            If optcriterio(0).Value = True Then
               QRY = QRY & "AND cocodemp = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA CODIGO DE RUTA ********************************
            If optcriterio(1).Value = True Then
               QRY = QRY & "AND cocodrut = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA POR TURNO *******************************
            If optcriterio(2).Value = True Then
               QRY = QRY & "AND coturno = " & Txtvalbus.text & " "
            End If
            
            '******************** CUANDO BUSCA POR NUMERO INTERNO ***********************
            If optcriterio(3).Value = True Then
               QRY = QRY & "AND conumint = " & Txtvalbus.text & " "
            End If
            
            '******************** CUANDO BUSCA POR ORIGEN/TRANSITO ********************
            If optcriterio(4).Value = True Then
               QRY = QRY & "AND cooritra = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA POR PLACA DEL VEHICULO ******************
            If optcriterio(5).Value = True Then
               QRY = QRY & "AND coplaca = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA CODIGO DE OPERARIO ********************************
            If optcriterio(6).Value = True Then
               QRY = QRY & "AND cofuncio = '" & Txtvalbus.text & "' "
            End If
            
            '******************** CUANDO BUSCA POR ESTADO DE CONDUCES ********************************
            If optcriterio(8).Value = True Then
               QRY = QRY & "AND coestado LIKE '" & Mid(Lblvalbus.Caption, 1, 1) & LCase(Mid(Lblvalbus.Caption, 2)) & "' "
            End If
            
            '******************** CUANDO BUSCA POR CAJA ********************************
            If optcriterio(9).Value = True Then
               QRY = QRY & "AND conomcaj = '" & Txtvalbus.text & "' "
            End If
            
            If chkcredito.Value = 1 Then
               QRY = QRY & " AND cocredit = 'S' "
            End If
            QRY1 = QRY
            
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
            
            QRY = QRY & "AND coestado = 'Activo'"
            
            Set Tbl1 = Dbs.Execute(QRY)
                        
            If Not Tbl1.EOF Then
               If Not IsNull(Tbl1(1)) Then
                  Mskalcoholemia.text = CDbl(Mskalcoholemia.text) + Tbl1(2)
                  msktasas.text = CDbl(msktasas.text) + Tbl1(1)
                  Mskparqueadero.text = CDbl(Mskparqueadero.text) + Tbl1(3)
                  Lblconven.Caption = CDbl(Lblconven.Caption) + Tbl1(0)
                  LblPasajeros.Caption = CDbl(LblPasajeros.Caption) + Tbl1(4)
                  bandera = 1
                  If chkcredito.Value = 1 Then
                     mskcredito.text = CDbl(mskcredito.text) + Tbl1(1)
                  End If
               Else
'                  Mskalcoholemia.Text = 0
'                  msktasas.Text = 0
'                  Mskparqueadero.Text = 0
'                  Lblconven.Caption = 0
'                  LblPasajeros.Caption = 0
'                  mskcredito.Text = 0
               End If
            End If
               
            Tbl1.Close
            If chkcredito.Value = 0 Then
            
               QRY = QRY1 & " AND cocredit = 'S' " _
                           & "AND coestado = 'Activo'"
                             
               
               Set Tbl1 = Dbs.Execute(QRY)
               
               
               If Not Tbl1.EOF Then
                  If Not IsNull(Tbl1(1)) Then
                     mskcredito.text = CDbl(mskcredito.text) + Tbl1(1)
                     bandera = 1
                  End If
                  
               End If
               Tbl1.Close
            End If
         Next k
         If bandera = 1 Then
            Msktotcon.text = CDbl(Mskalcoholemia.text) + CDbl(Mskparqueadero.text) + CDbl(msktasas.text)
         End If
      End If
  End If
  Conconduces.Enabled = True
  Grdconcon.Visible = True
  ProBar.Value = ProBar.Max
  If bandera = 0 Then
     MsgBox "No Se Encontro Ningun Registro Activo", 48, "Advertencia"
     Exit Sub
  End If
  Call Aplicar_Perfil(Me, "Conconduces")
  If Btnimprimir.Enabled = True Then
      If L_DirCristal = "S" Then
         ImpReporte
      End If
  End If
  Exit Sub
  
terminar:
   Grdconcon.Visible = True
   Conconduces.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
Resume
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo err_mensaje
   If Grdconcon.rows > 1 Then
'      lineas = Grdconcon.Rows - 1
'      Grdconcon.Row = 1
'      ProBar.Value = ProBar.Min
'      ProBar.Max = Grdconcon.Rows + 2
'      Grdconcon.Visible = False
'      Conconduces.Enabled = False
'      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
'
'      Dbs.BeginTrans
'      On Error GoTo Err_Trans
'      Do
'         DoEvents
'        ReDim dato(1, 26)
'
'        For I = 0 To 25
'            Grdconcon.Col = I
'            dato(1, I) = Trim(Grdconcon.Text)
'        Next I
'
'        ' insertamos en la tabla
'
'        Qry = "INSERT INTO optempo1 " _
'        & "(temnumero1, temtext9, temtext2, temdate1, " _
'        & " temtext8, temtext4, temnumero3, temtext5, " _
'        & " temtext6, temnumero2, temtext11, temnumdob3, " _
'        & " temnumdob2, temnumdob4, temtext7, temnumero6, " _
'        & " temtext12, tempc, temtext13, temtext1, temtext3, temtext14, temnumero4, temtext15) VALUES ('" _
'        & dato(1, 0) & "','" & dato(1, 2) & "','" & dato(1, 4) & "',to_date('" & dato(1, 5) & "','MM/DD/YYYY'),'" _
'        & Format(dato(1, 6), "MM/dd/yyyy HH:mm ") & "','" & dato(1, 7) & "','" & dato(1, 8) & "', '" & dato(1, 10) & "','" _
'        & dato(1, 21) & "','" & dato(1, 12) & "','" & dato(1, 13) & "','" & dato(1, 14) & "','" _
'        & dato(1, 15) & "','" & dato(1, 16) & "','" & dato(1, 17) & "','" & dato(1, 18) & "','" _
'        & dato(1, 19) & "','" & Pc & "','" & dato(1, 20) & "','" & dato(1, 22) & "','" & dato(1, 11) & "','" & dato(1, 23) & "', " & dato(1, 24) & ",'" & dato(1, 25) & "')"
'
'
'        Dbs.Execute Qry
'        ' fin insertar
'
'
'        lineas = lineas - 1
'        ProBar.Value = ProBar.Value + 1
'        If lineas <> 0 Then
'           Grdconcon.Row = Grdconcon.Row + 1
'        End If
'     Loop Until lineas = 0  'do_while que se mueve x la grilla
'
'
'
'     Dbs.CommitTrans
'     On Error GoTo err_mensaje
'     Grdconcon.Visible = True
     ImpReporte
     'ProBar.Value = ProBar.Max
     
     
  End If
  
  Exit Sub
'Err_Trans:
'   Grdconcon.Visible = True
'   Dbs.RollbackTrans
'   Conconduces.Enabled = True
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Sub
  
err_mensaje:
   Grdconcon.Visible = True
   Conconduces.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
Resume
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
  optcriterio(7).Value = True
  Txtvalbus.text = ""
  Lblvalbus.Caption = ""
  Grdconcon.rows = 1
  Grdconcon.Cols = 2
  Grdconcon.Clear
  Optbuscar(0).Value = True
  Optgenerar(1).Value = True
  Msktotcon.text = ""
  Lblconven.Caption = ""
  Mskalcoholemia.text = ""
  msktasas.text = ""
  Mskparqueadero.text = ""
  mskcredito.text = ""
  chkcredito.Value = 0
  LblPasajeros.Caption = ""
  DIB_GRILLA
  ProBar.Value = ProBar.Min
  CmbTerminal.ListIndex = 0
  Cmbcaja.ListIndex = -1
  Cmbcaja.Enabled = False
  CmbTipoRuta.ListIndex = 0
  OptAlistamiento(2).Value = True
  OptTipoTasa(2).Value = True
  OptCodSuper(2).Value = True
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Conconduces")
End Sub

Private Sub Form_Load()
  Optbuscar_Click (Index)
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
   QRY = " SELECT canomcaj FROM cocaja"
   Set Tbl2 = Dbs.Execute(QRY)
   If Not Tbl2.EOF Then
    Cmbcaja.Clear
    Cmbcaja.AddItem ""
    
    While Not Tbl2.EOF
     Cmbcaja.AddItem Tbl2!canomcaj
     
     Tbl2.MoveNext
     
    Wend
   End If
   CmbTerminal.ListIndex = 0
   CmbTipoRuta.ListIndex = 0
   If G_rolusu = "REPFIN" Then
      Opttipfec(1).Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Optbuscar_Click(Index As Integer)
  If Optbuscar(0).Value = True Then
     Frmnummul.Visible = True
     Frmfechas.Visible = False
     Txtnumini.TabIndex = 2
     Txtnumfin.TabIndex = 3
  ElseIf Optbuscar(1).Value = True Then
     Frmnummul.Visible = False
     Frmfechas.Visible = True
     Mskfecini.TabIndex = 2
     Mskfecfin.TabIndex = 3
  End If
End Sub

Private Sub optcriterio_Click(Index As Integer)
  Txtvalbus.text = ""
  Lblvalbus.Caption = ""
End Sub

Private Sub Opttipfec_Click(Index As Integer)
   Cmbcaja.Enabled = False
   
   If Index = 1 Then
      'turno
      Cmbcaja.Enabled = True
   End If
End Sub

Private Sub Txtnumfin_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Then
      KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumini_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Then
      KeyAscii% = 0
  End If
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

Private Sub Txtvalbus_DblClick()
  If optcriterio(7).Value = False Then
     If optcriterio(0).Value = True Then
        forma = "16"
        forma1 = "14"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
     
     If optcriterio(1).Value = True Then
        forma = "3"
        forma1 = "8"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
     
     If optcriterio(2).Value = True Then
     
     End If
     
     If optcriterio(3).Value = True Then
        forma = "2"
        forma1 = "5"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
     
     If optcriterio(4).Value = True Then
     
     End If
     
     If optcriterio(5).Value = True Then
        forma = "2"
        forma1 = "6"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
     
     If optcriterio(6).Value = True Then
     
     End If
     
     If optcriterio(8).Value = True Then
        forma = "15"
        forma1 = "20"
        Tabla_Catalogo = "ESTADOS"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
     
     If optcriterio(9).Value = True Then
        forma = "12"
        forma1 = "1"
        Txtvalbus.text = ""
        Lblvalbus.Caption = ""
        buscar.Show 1
        Txtvalbus_LostFocus
     End If
  End If
End Sub

Private Sub Txtvalbus_KeyPress(KeyAscii As Integer)
  If optcriterio(7).Value = True Then
     KeyAscii% = 0
     Exit Sub
  End If
  
  If optcriterio(0).Value = True Or optcriterio(1).Value = True Or optcriterio(2).Value = True Or optcriterio(3).Value = True Or optcriterio(8).Value = True Then
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
  
  If optcriterio(4).Value = True Then
     Char = Chr(KeyAscii%)
     KeyAscii = Asc(UCase(Char))  ' mayuscula
     If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
  
  If optcriterio(6).Value = True Then
     Char = Chr(KeyAscii%)
     KeyAscii% = Asc(LCase(Char))
     If (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% = 8) Then
         Exit Sub
     Else
         KeyAscii% = 0
     End If
  End If
  
  If optcriterio(5).Value = True Then
     Char = Chr(KeyAscii%)
     KeyAscii = Asc(UCase(Char))  ' mayuscula
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
  
  
  If optcriterio(9).Value = True Then
     Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))  ' mayuscula
      
      If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
         KeyAscii% = 0
      End If
  End If
End Sub

Private Sub Txtvalbus_LostFocus()
  If Txtvalbus.text <> "" And optcriterio(7).Value = False Then
     'Data1.RecordsetType = 1     'Recordset de tipo SnapShot
     If optcriterio(0).Value = True Then
        fqry01 = "SELECT * " _
                 & "FROM ophojemp " _
                & "WHERE hecodigo = '" & Txtvalbus.text & "'"
                
        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
        Set Ftb01 = Dbs.Execute(fqry01)
        
        
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Ftb01!henombre
           Ftb01.Close
        Else
           MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
           forma = "16"
           forma1 = "14"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
     
     If optcriterio(1).Value = True Then
        fqry01 = "SELECT * " _
                 & "FROM opruta " _
                & "WHERE rucodigo = '" & Txtvalbus.text & "'"
                
        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Ftb01!rudestino
           Ftb01.Close
        Else
           MsgBox "Codigo De Ruta No Existe", 48, "Advertencia"
           forma = "3"
           forma1 = "8"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
     
     If optcriterio(2).Value = True Then
        
     End If
     
     If optcriterio(3).Value = True Then
        fqry01 = "SELECT * " _
                 & "FROM ophojveh " _
                & "WHERE hvnumint = " & Txtvalbus.text & " "
         
        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Ftb01!hvconact
           Ftb01.Close
        Else
           MsgBox "Numero Interno De Vehiculo No Existe", 48, "Advertencia"
           forma = "2"
           forma1 = "5"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
     
     If optcriterio(4).Value = True Then
     
     End If
     
     If optcriterio(5).Value = True Then
        fqry01 = "SELECT * " _
                 & "FROM ophojveh " _
                & "WHERE hvplaca = '" & Txtvalbus.text & "' "
         
        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Ftb01!hvconact
           Ftb01.Close
        Else
           MsgBox "Placa De Vehiculo No Existe", 48, "Advertencia"
           forma = "2"
           forma1 = "6"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
     
     If optcriterio(6).Value = True Then
     
     End If
     
     If optcriterio(8).Value = True Then
        fqry01 = "SELECT dsdes " _
                 & "FROM gesuptip, gedetsuptip " _
                & "WHERE dscoddet = '" & Txtvalbus.text & "' " _
                  & "AND dscodtip = stcodtip " _
                  & "AND stdes = 'ESTADOS' "

        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
       
        Set Ftb01 = Dbs.Execute(fqry01)
       
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Ftb01!dsdes
           Ftb01.Close
        Else
           MsgBox "Codigo De Estado No Existe", 48, "Advertencia"
           forma = "15"
           forma1 = "20"
           Tabla_Catalogo = "ESTADOS"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
     
     If optcriterio(9).Value = True Then
        fqry01 = "SELECT * " _
                 & "FROM cocaja " _
                & "WHERE canomcaj = '" & Txtvalbus.text & "' "
         
        'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Ftb01.EOF Then
           Ftb01.Close
           MsgBox "Caja No Existe", 48, "Advertencia"
           Txtvalbus.SetFocus
           Exit Sub
        End If
        Ftb01.Close
     End If
  End If
End Sub

Private Function DIB_GRILLA()
     Grdconcon.rows = 1
     Grdconcon.Cols = 31
     Grdconcon.Clear
     
     s$ = "Número  |Pref. Fact. |Num. Fact. |Cód. Ruta |Destino |Cód. Empresa |Nombre   |Fec. Salida |Hor. Salida  |Núm. Placa  |Núm. Int |Cedula1|Conductor1 |Cedula2|Conductor2 |" _
         & "Origen/Transito |Funcionario |Turno |Modo |Pasajeros |Valor |Alcoholemia |Estado|Parqueadero |Caja |Crédito|Corredor|Terminal|Clase|Cant. Puestos|Peaje"
     Grdconcon.FormatString = s$
     Grdconcon.Row = 0
     Grdconcon.Col = 0
     Grdconcon.ColWidth(0) = 1000
     Grdconcon.ColWidth(1) = 1200
     Grdconcon.ColWidth(2) = 1200
     Grdconcon.ColWidth(3) = 1000
     Grdconcon.ColWidth(4) = 1800
     Grdconcon.ColWidth(5) = 1500
     Grdconcon.ColWidth(6) = 1800
     Grdconcon.ColWidth(7) = 1200
     Grdconcon.ColWidth(8) = 1700
     Grdconcon.ColWidth(9) = 1300
     Grdconcon.ColWidth(10) = 1300
     Grdconcon.ColWidth(11) = 1100
     Grdconcon.ColWidth(12) = 1700
     Grdconcon.ColWidth(13) = 1100
     Grdconcon.ColWidth(14) = 1700
     
     Grdconcon.ColWidth(15) = 1300
     Grdconcon.ColWidth(16) = 1300
     Grdconcon.ColWidth(17) = 1300
     Grdconcon.ColWidth(18) = 1000
     Grdconcon.ColWidth(19) = 1000
     Grdconcon.ColWidth(20) = 1000
     Grdconcon.ColWidth(21) = 1500
     Grdconcon.ColWidth(22) = 1000
     Grdconcon.ColWidth(23) = 1500
     Grdconcon.ColWidth(24) = 1500
     Grdconcon.ColWidth(25) = 1500
     Grdconcon.ColWidth(26) = 1500
     Grdconcon.ColWidth(27) = 1000
     Grdconcon.ColWidth(28) = 1000
     Grdconcon.ColWidth(29) = 1500
     Grdconcon.ColWidth(30) = 1200
     For i = 0 To Grdconcon.Cols - 1
        Grdconcon.Col = i
        Grdconcon.CellFontBold = 1
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
     
   respuesta = MsgBox("Desea Imprimir reporte general (Si)  o Evasión (No)", 32 + vbYesNo, "Imprimir")
  
   If respuesta = vbYes Then
      reporte = RpPath + "\concondu.rpt"
   Else
      reporte = RpPath + "\conconev.rpt"
   End If
   rptterminal.ReportFileName = reporte
   rptterminal.Connect = G_Conexion_reportes
   rptterminal.Formulas(0) = "inicio1 = '" & inicio1 & "'"
   rptterminal.Formulas(1) = "final = '" & final & "'"
   rptterminal.Formulas(2) = "empresa = '" & G_NomEmpresa & "'"
   rptterminal.Formulas(3) = "totgen = '" & Msktotcon.FormattedText & "'"
   rptterminal.Formulas(4) = "conven = '" & Lblconven.Caption & "'"
   rptterminal.Formulas(6) = "tottasa = '" & msktasas.FormattedText & "'"
   rptterminal.Formulas(7) = "totalc = '" & Mskalcoholemia.FormattedText & "'"
   rptterminal.Formulas(9) = "parque = '" & Mskparqueadero.FormattedText & "'"
   rptterminal.Formulas(10) = "credito = '" & mskcredito.FormattedText & "'"
   rptterminal.Formulas(11) = "pasajeros = '" & LblPasajeros.Caption & "'"
   rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
   Conconduces.Enabled = False
   rptterminal.PrinterName = G_PrinterName
   rptterminal.PrinterDriver = G_PrinterDriver
   rptterminal.PrinterPort = G_PrinterPort
   rptterminal.Destination = 0
   rptterminal.WindowState = crptMaximized
   rptterminal.Action = 1
   Conconduces.Enabled = True
   
End Function

