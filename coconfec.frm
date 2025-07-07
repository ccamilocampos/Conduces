VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form coconfec 
   Caption         =   "Consulta Facturación Electrónica Tasa de Uso"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Agrupar por"
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
      TabIndex        =   33
      Top             =   4320
      Width           =   2895
      Begin VB.OptionButton OptGrupo 
         Caption         =   "Empresas"
         BeginProperty Font 
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
         Left            =   1320
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptGrupo 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1095
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
      Left            =   7320
      Picture         =   "coconfec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7440
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
      Left            =   4920
      Picture         =   "coconfec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Limpia la forma actual"
      Top             =   7440
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
      Left            =   2640
      Picture         =   "coconfec.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   7440
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
      Left            =   720
      Picture         =   "coconfec.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   7440
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Totales"
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
      TabIndex        =   20
      Top             =   5040
      Width           =   8535
      Begin MSMask.MaskEdBox MskCanTasas 
         Height          =   300
         Left            =   480
         TabIndex        =   37
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
      Begin MSMask.MaskEdBox MskTotTasas 
         Height          =   300
         Left            =   480
         TabIndex        =   38
         Top             =   1440
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
      Begin MSMask.MaskEdBox MskCanFacGen 
         Height          =   300
         Left            =   2400
         TabIndex        =   39
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
      Begin MSMask.MaskEdBox MskTotFacGen 
         Height          =   300
         Left            =   2400
         TabIndex        =   40
         Top             =   1440
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
      Begin MSMask.MaskEdBox MskCanFacPen 
         Height          =   300
         Left            =   4440
         TabIndex        =   41
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
      Begin MSMask.MaskEdBox MskTotFacPen 
         Height          =   300
         Left            =   4440
         TabIndex        =   42
         Top             =   1440
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
      Begin MSMask.MaskEdBox MskCanFacRec 
         Height          =   300
         Left            =   6600
         TabIndex        =   43
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
      Begin MSMask.MaskEdBox MskTotFacRec 
         Height          =   300
         Left            =   6600
         TabIndex        =   44
         Top             =   1440
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
         Alignment       =   2  'Center
         Caption         =   "$"
         BeginProperty Font 
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
         Left            =   6600
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "$"
         BeginProperty Font 
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
         Left            =   4440
         TabIndex        =   27
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "$"
         BeginProperty Font 
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
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "$"
         BeginProperty Font 
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
         TabIndex        =   25
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Facturas Rechazadas"
         BeginProperty Font 
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
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Facturas Pendientes"
         BeginProperty Font 
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
         Left            =   4320
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Facturas Generadas"
         BeginProperty Font 
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
         Left            =   2280
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total Tasas"
         BeginProperty Font 
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
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   2895
      Begin VB.TextBox txtcodEmp 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cod. Empresa:"
         BeginProperty Font 
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
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNomEmpresa 
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
         TabIndex        =   17
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterios"
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
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
      Begin VB.ComboBox cmbCaja 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbTurno 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Txtnumfin 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txtnumini 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consultar por"
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
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton OptRango 
         Caption         =   "Conduces"
         BeginProperty Font 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptRango 
         Caption         =   "Fecha"
         BeginProperty Font 
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
         Width           =   1215
      End
   End
   Begin VB.Frame coconfec 
      Caption         =   "Rangos"
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
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
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
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
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
      Begin VB.Label lblFinal 
         Caption         =   "lblFinal"
         BeginProperty Font 
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
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblInicio 
         Caption         =   "lblInicio"
         BeginProperty Font 
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
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdFETU 
      Height          =   4815
      Left            =   3120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8493
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   36
      Top             =   7080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   1920
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   8880
      Y1              =   7320
      Y2              =   7320
   End
End
Attribute VB_Name = "coconfec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btnbuscar_Click()
    On Error GoTo Errores
    
    If OptRango(0).Value = True Then
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
         
    ElseIf OptRango(1).Value = True Then
        If Txtnumini.text = "" Then
            MsgBox "Debe digitar un número de inicio valido", 48, "Advertencia"
            Txtnumini.SetFocus
            Exit Sub
        End If
        
        If Txtnumfin.text = "" Then
            MsgBox "Debe digitar un numero de fin valido", 48, "Advertencia"
            Txtnumfin.SetFocus
            Exit Sub
        End If
        
        If CDbl(Txtnumini.text) > CDbl(Txtnumfin.text) Or CDbl(Txtnumfin.text) < CDbl(Txtnumini.text) Then
            MsgBox "Digite un rango de tasas de uso valido", 48, "Advertencia"
            Exit Sub
        End If
    End If
    
    CantidadTotalTasas = 0
    TotalTasas = 0
    CantidadTotalTasasGeneradas = 0
    TotalTasasGeneradas = 0
    CantidadTotalTasasPendientes = 0
    TotalTasasPendientes = 0
    CantidadTotalTasasRechazadas = 0
    TotalTasasRechazadas = 0
    
    DIB_GRILLA
    
    'Cabecera consulta
    If OptGrupo(0).Value = True Then
    
        QRY = "SELECT DISTINCT(conomcaj) AS conomcaj, SUM(CantTasasGen) AS CantTasasGen, SUM(TotTasasGen) AS TotTasasGen, SUM(CantTasasPen) AS CantTasasPen, SUM(TotTasasPen) AS TotTasasPen, SUM(CantTasasRec) AS CantTasasRec, SUM(TotTasasRec) AS TotTasasRec "
                
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = "SELECT DISTINCT(cocodemp) AS cocodemp, henombre, SUM(CantTasasGen) AS CantTasasGen, SUM(TotTasasGen) AS TotTasasGen, SUM(CantTasasPen) AS CantTasasPen, SUM(TotTasasPen) AS TotTasasPen, SUM(CantTasasRec) AS CantTasasRec, SUM(TotTasasRec) AS TotTasasRec "
    
    End If
    
    QRY = QRY & "FROM ("
    
    'Busqueda Facturas en estado OK
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(conomcaj) AS conomcaj, COUNT(conumero) AS CantTasasGen, SUM(covalor) AS TotTasasGen, 0 AS CantTasasPen, 0 AS TotTasasPen, 0 AS CantTasasRec, 0 AS TotTasasRec  " _
                        & "FROM coconduc " _
                        & "WHERE coestfe = 'OK' " _
                        & "AND coestado = 'Activo' "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(cocodemp) AS cocodemp, henombre, COUNT(conumero) AS CantTasasGen, SUM(covalor) AS TotTasasGen, 0 AS CantTasasPen, 0 AS TotTasasPen, 0 AS CantTasasRec, 0 AS TotTasasRec " _
                        & "FROM coconduc, ophojemp " _
                        & "WHERE coconduc.cocodemp = ophojemp.hecodigo " _
                        & "AND coestfe = 'OK' " _
                        & "AND coestado = 'Activo' "
    
    End If
    
    If OptRango(0).Value = True Then
    
        QRY = QRY & "AND cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                        & "AND TO_DATE('" & Mskfecfin.text & "', 'MM/DD/YYYY') "
                        
    ElseIf OptRango(1).Value = True Then
    
        QRY = QRY & "AND conumero BETWEEN " & Txtnumini.text & " AND " & Txtnumfin.text & " "
    
    End If
    
    If cmbCaja.ListIndex <> 0 Then
        QRY = QRY & "AND conomcaj = '" & cmbCaja.text & "' "
    End If
    
    If cmbTurno.ListIndex <> 0 Then
        QRY = QRY & "AND coturno = " & cmbTurno.text & " "
    End If
    
    If txtcodEmp.text <> "" Then
        QRY = QRY & "AND cocodemp = '" & txtcodEmp.text & "' "
    End If
    
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "GROUP BY conomcaj, coturno "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "GROUP BY cocodemp, henombre "
        
    End If
    
    QRY = QRY & " UNION "
    
    'Busqueda Facturas en estado PEN
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(conomcaj) AS conomcaj, 0 AS CantTasasGen, 0 AS TotTasasGen, COUNT(conumero) AS CantTasasPen, SUM(covalor) AS TotTasasPen, 0 AS CantTasasRec, 0 AS TotTasasRec " _
                        & "FROM coconduc " _
                        & "WHERE coestfe = 'PEN' " _
                        & "AND coestado = 'Activo' "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(cocodemp) AS cocodemp, henombre, 0 AS CantTasasGen, 0 AS TotTasasGen, COUNT(conumero) AS CantTasasPen, SUM(covalor) AS TotTasasPen, 0 AS CantTasasRec, 0 AS TotTasasRec " _
                        & "FROM coconduc, ophojemp " _
                        & "WHERE coconduc.cocodemp = ophojemp.hecodigo " _
                        & "AND coestfe = 'PEN' " _
                        & "AND coestado = 'Activo' "
    
    End If
    
    If OptRango(0).Value = True Then
    
        QRY = QRY & "AND cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                        & "AND TO_DATE('" & Mskfecfin.text & "', 'MM/DD/YYYY') "
                        
    ElseIf OptRango(1).Value = True Then
    
        QRY = QRY & "AND conumero BETWEEN " & Txtnumini.text & " AND " & Txtnumfin.text & " "
    
    End If
    
    If cmbCaja.ListIndex <> 0 Then
        QRY = QRY & "AND conomcaj = '" & cmbCaja.text & "' "
    End If
    
    If cmbTurno.ListIndex <> 0 Then
        QRY = QRY & "AND coturno = " & cmbTurno.text & " "
    End If
    
    If txtcodEmp.text <> "" Then
        QRY = QRY & "AND cocodemp = '" & txtcodEmp.text & "' "
    End If
    
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "GROUP BY conomcaj "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "GROUP BY cocodemp, henombre "
        
    End If
    
    QRY = QRY & " UNION "
    
    'Busqueda Facturas Rechazadas
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(conomcaj), 0 AS CantTasasGen, 0 AS TotTasasGen, 0 AS CantTasasPen, 0 AS TotTasasPen, COUNT(conumero) AS CantTasasRec, SUM(covalor) AS TotTasasRec " _
                        & "FROM coconduc, logfe " _
                        & "WHERE coconduc.conumero = logfe.loconduce " _
                        & "AND coestfe = 'PEN' " _
                        & "AND coestado = 'Activo' "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "SELECT DISTINCT(cocodemp), henombre, 0 AS CantTasasGen, 0 AS TotTasasGen, 0 AS CantTasasPen, 0 AS TotTasasPen, COUNT(conumero) AS CantTasasRec, SUM(covalor) AS TotTasasRec " _
                        & "FROM coconduc, logfe, ophojemp " _
                        & "WHERE coconduc.conumero = logfe.loconduce " _
                        & "AND coestfe = 'PEN' " _
                        & "AND coestado = 'Activo' " _
                        & "AND coconduc.cocodemp = ophojemp.hecodigo "
    
    End If
    
    If OptRango(0).Value = True Then
    
        QRY = QRY & "AND cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                        & "AND TO_DATE('" & Mskfecfin.text & "', 'MM/DD/YYYY') "
                        
    ElseIf OptRango(1).Value = True Then
    
        QRY = QRY & "AND conumero BETWEEN " & Txtnumini.text & " AND " & Txtnumfin.text & " "
    
    End If
    
    If cmbCaja.ListIndex <> 0 Then
        QRY = QRY & "AND conomcaj = '" & cmbCaja.text & "' "
    End If
    
    If cmbTurno.ListIndex <> 0 Then
        QRY = QRY & "AND coturno = " & cmbTurno.text & " "
    End If
    
    If txtcodEmp.text <> "" Then
        QRY = QRY & "AND cocodemp = '" & txtcodEmp.text & "' "
    End If
    
    If OptGrupo(0).Value = True Then
    
        QRY = QRY & "GROUP BY conomcaj " _
                        & ") GROUP BY conomcaj "
    
    ElseIf OptGrupo(1).Value = True Then
    
        QRY = QRY & "GROUP BY cocodemp, henombre " _
                        & ") GROUP BY cocodemp, henombre "
        
    End If
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el Progress Bar
    If Tbl1.EOF = False Then
        Tbl1.MoveLast
        ProBar.Value = ProBar.Min
        ProBar.Max = Tbl1.RecordCount + 2
        Tbl1.MoveFirst
    Else
        MsgBox "No se encontraron registros correspondientes", 48, "Advertencia"
        Exit Sub
    End If
    
    'Llenamos la grilla
    If OptGrupo(0).Value = True Then
    
        Do While Not Tbl1.EOF
            CantTasasCaja = 0
            TotTasasCaja = 0
            CantTasasCaja = CDbl(Tbl1!CantTasasGen) + CDbl(Tbl1!CantTasasPen) + CDbl(Tbl1!CantTasasRec)
            TotTasasCaja = CDbl(Tbl1!TotTasasGen) + CDbl(Tbl1!TotTasasPen) + CDbl(Tbl1!TotTasasRec)
            
            'Sumamos los valores totales
            entry = Tbl1!conomcaj + Chr(9) + CStr(CantTasasCaja) + Chr(9) + CStr(TotTasasCaja) + Chr(9) + _
                    CStr(Tbl1!CantTasasGen) + Chr(9) + CStr(Tbl1!TotTasasGen) + Chr(9) + _
                    CStr(Tbl1!CantTasasPen) + Chr(9) + CStr(Tbl1!TotTasasPen) + Chr(9) + _
                    CStr(Tbl1!CantTasasRec) + Chr(9) + CStr(Tbl1!TotTasasRec)
            
            'Sumamos los valores totales
            CantidadTotalTasas = CantidadTotalTasas + CantTasasCaja
            TotalTasas = TotalTasas + TotTasasCaja
            CantidadTotalTasasGeneradas = CantidadTotalTasasGeneradas + CDbl(Tbl1!CantTasasGen)
            TotalTasasGeneradas = TotalTasasGeneradas + CDbl(Tbl1!TotTasasGen)
            CantidadTotalTasasPendientes = CantidadTotalTasasPendientes + CDbl(Tbl1!CantTasasPen)
            TotalTasasPendientes = TotalTasasPendientes + CDbl(Tbl1!TotTasasPen)
            CantidadTotalTasasRechazadas = CantidadTotalTasasRechazadas + CDbl(Tbl1!CantTasasRec)
            TotalTasasRechazadas = TotalTasasRechazadas + CDbl(Tbl1!TotTasasRec)
            
            grdFETU.AddItem (entry)
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Loop
        
    ElseIf OptGrupo(1).Value = True Then
    
        Do While Not Tbl1.EOF
            CantTasasEmp = 0
            TotTasasEmp = 0
            CantTasasEmp = CDbl(Tbl1!CantTasasGen) + CDbl(Tbl1!CantTasasPen) + CDbl(Tbl1!CantTasasRec)
            TotTasasEmp = CDbl(Tbl1!TotTasasGen) + CDbl(Tbl1!TotTasasPen) + CDbl(Tbl1!TotTasasRec)
            
            entry = Tbl1!cocodemp + Chr(9) + Tbl1!henombre + Chr(9) + CStr(CantTasasEmp) + Chr(9) + CStr(TotTasasEmp) + Chr(9) + _
                        CStr(Tbl1!CantTasasGen) + Chr(9) + CStr(Tbl1!TotTasasGen) + Chr(9) + _
                        CStr(Tbl1!CantTasasPen) + Chr(9) + CStr(Tbl1!TotTasasPen) + Chr(9) + _
                        CStr(Tbl1!CantTasasRec) + Chr(9) + CStr(Tbl1!TotTasasRec)
            
            'Sumamos los valores totales
            CantidadTotalTasas = CantidadTotalTasas + CantTasasEmp
            TotalTasas = TotalTasas + TotTasasEmp
            CantidadTotalTasasGeneradas = CantidadTotalTasasGeneradas + CDbl(Tbl1!CantTasasGen)
            TotalTasasGeneradas = TotalTasasGeneradas + CDbl(Tbl1!TotTasasGen)
            CantidadTotalTasasPendientes = CantidadTotalTasasPendientes + CDbl(Tbl1!CantTasasPen)
            TotalTasasPendientes = TotalTasasPendientes + CDbl(Tbl1!TotTasasPen)
            CantidadTotalTasasRechazadas = CantidadTotalTasasRechazadas + CDbl(Tbl1!CantTasasRec)
            TotalTasasRechazadas = TotalTasasRechazadas + CDbl(Tbl1!TotTasasRec)
            
            grdFETU.AddItem (entry)
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Loop
    End If
    
    ProBar.Value = ProBar.Max
    
    MskCanTasas.text = CantidadTotalTasas
    MskTotTasas.text = "$ " & TotalTasas
    MskCanFacGen.text = CantidadTotalTasasGeneradas
    MskTotFacGen.text = "$ " & TotalTasasGeneradas
    MskCanFacPen.text = CantidadTotalTasasPendientes
    MskTotFacPen.text = "$ " & TotalTasasPendientes
    MskCanFacRec.text = CantidadTotalTasasRechazadas
    MskTotFacRec.text = "$ " & TotalTasasRechazadas
    
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub


Private Sub Btnimprimir_Click()
    On Error GoTo err_mensaje
    
    If grdFETU.rows > 1 Then
    
        lineas = grdFETU.rows - 1
        grdFETU.Row = 1
        grdFETU.Visible = False
        ProBar.Value = ProBar.Min
        ProBar.Max = grdFETU.rows + 2
        Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
        
        Dbs.BeginTrans
        On Error GoTo Err_Trans
        Do
        
            If OptGrupo(0).Value = True Then
                ReDim dato(1, 9)
                
                For i = 0 To 8
                    grdFETU.Col = i
                    dato(1, i) = Trim(grdFETU.text)
                Next i
            
            ElseIf OptGrupo(1).Value = True Then
                ReDim dato(1, 10)
                
                For i = 0 To 9
                    grdFETU.Col = i
                    dato(1, i) = Trim(grdFETU.text)
                Next i
            End If
            
            'Inserta en la tabla temporal
            If OptGrupo(0).Value = True Then
                QRY = "INSERT INTO optempo1 " _
                        & "(temtext1, temnumero1, temnumdob1, temnumero2, temnumdob2, temnumero3, temnumdob3, temnumero4, temnumdob4, tempc) " _
                        & "VALUES ('" & dato(1, 0) & "', " & dato(1, 1) & ", " & dato(1, 2) & ", " & dato(1, 3) & ", " _
                        & dato(1, 4) & ", " & dato(1, 5) & ", " & dato(1, 6) & ", " & dato(1, 7) & ", " & dato(1, 8) & ", '" & Pc & "') "
                Dbs.Execute (QRY)
                
            ElseIf OptGrupo(1).Value = True Then
                QRY = "INSERT INTO optempo1 " _
                        & "(temtext1, temtext2, temnumero1, temnumdob1, temnumero2, temnumdob2, temnumero3, temnumdob3, temnumero4, temnumdob4, tempc) " _
                        & "VALUES ('" & dato(1, 0) & "', '" & dato(1, 1) & "', " & dato(1, 2) & ", " & dato(1, 3) & ", " _
                        & dato(1, 4) & ", " & dato(1, 5) & ", " & dato(1, 6) & ", " & dato(1, 7) & ", " & dato(1, 8) & ", " _
                        & dato(1, 9) & ", '" & Pc & "') "
                Dbs.Execute (QRY)
            End If
            
            lineas = lineas - 1
            ProBar.Value = ProBar.Value + 1
            If lineas <> 0 Then
                grdFETU.Row = grdFETU.Row + 1
            End If
        Loop Until lineas = 0
        
        Dbs.CommitTrans
        grdFETU.Visible = True
        On Error GoTo err_mensaje
            respuesta = MsgBox("Desea escoger la impresora (SI) o impresora predeterminada (NO)", vbYesNo + 32, "Impresoras")
        
            If respuesta = 6 Then
               imprimir.Show 1
               If G_Imprimir = "0" Then
                  Exit Sub
               End If
            End If
            
            reporte = RpPath + "\coconfec.rpt"
            
            rptterminal.ReportFileName = reporte
            rptterminal.Connect = G_Conexion_reportes
            rptterminal.Formulas(1) = "fecini = '" & Mskfecini.FormattedText & "'"
            rptterminal.Formulas(2) = "fecfin = '" & Mskfecfin.FormattedText & "'"
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
            coconfec.Enabled = True
            Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
            ProBar.Value = ProBar.Max
    End If
    Exit Sub
Err_Trans:
    Dbs.RollbackTrans
err_mensaje:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Btnlimpiar_Click()
    Mskfecini.Mask = ""
    Mskfecini.text = ""
    Mskfecini.Mask = "##/##/####"
    Mskfecfin.Mask = ""
    Mskfecfin.text = ""
    Mskfecfin.Mask = "##/##/####"
    Txtnumini.text = ""
    Txtnumfin.text = ""
    txtcodEmp.text = ""
    lblNomEmpresa.Caption = ""
    cmbCaja.ListIndex = 0
    cmbTurno.ListIndex = 0
    OptRango(0).Value = True
    OptGrupo(0).Value = True
    DIB_GRILLA
    
    MskCanTasas.text = ""
    MskTotTasas.text = ""
    MskCanFacGen.text = ""
    MskTotFacGen.text = ""
    MskCanFacPen.text = ""
    MskTotFacPen.text = ""
    MskCanFacRec.text = ""
    MskTotFacRec.text = ""
    
    ProBar.Value = ProBar.Min
End Sub


Private Sub Btnsalir_Click()
    Unload Me
End Sub


Private Sub txtcodemp_DblClick()
    forma = "16"
    forma1 = "71"
    lblNomEmpresa.Caption = ""
    buscar.Show 1
    txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
    If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Or (KeyAscii% = 47) Then
      KeyAscii% = 0
  End If
  
End Sub

Private Sub txtcodemp_LostFocus()
    
    If txtcodEmp.text <> "" Then
        
        QRY = "SELECT henombre FROM ophojemp " _
            & "WHERE hecodigo = '" & txtcodEmp.text & "' "
            
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
            lblNomEmpresa.Caption = Tbl1!henombre
            Tbl1.Close
        Else
            
            QRY = "SELECT henombre FROM ophojempcon " _
                & "WHERE hecodigo = '" & txtcodEmp.text & "' "
                
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Tbl2.EOF = False Then
                lblNomEmpresa.Caption = Tbl2!henombre
            Else
                Tbl1.Close
                MsgBox "Empresa no existente", 48, "Advertencia"
                txtcodEmp.text = ""
                txtcodemp_DblClick
            End If
            Tbl2.Close
        End If
    End If
    
End Sub

Private Sub Form_Load()
    OptRango(0).Value = True
    OptGrupo(0).Value = True
    
    'Llenar combo cajas
    fqry01 = "SELECT canomcaj " _
             & "FROM cocaja " _
         & "ORDER BY canomcaj "
             
    Set Ftb01 = Dbs.Execute(fqry01)
    
    cmbCaja.AddItem "GENERAL"
    Do While Not Ftb01.EOF
       cmbCaja.AddItem Ftb01(0)
       Ftb01.MoveNext
    Loop
    Ftb01.Close
    
    cmbCaja.ListIndex = 0
    
    'Llenamos el combo de turnos
    cmbTurno.AddItem "GENERAL"
    cmbTurno.AddItem "1"
    cmbTurno.AddItem "2"
    cmbTurno.AddItem "3"
    cmbTurno.AddItem "4"
    
    cmbTurno.ListIndex = 0
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    
    
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

Public Function DIB_GRILLA()
    
    If OptGrupo(0).Value = True Then
        grdFETU.rows = 1
        grdFETU.Cols = 9
        grdFETU.Clear
        s$ = "Caja |"
    Else
        grdFETU.rows = 1
        grdFETU.Cols = 10
        grdFETU.Clear
        s$ = "Nit |Nom. empresa |"
    End If
    
    s$ = s$ & "Cant. Tasas |$ Tasas |Cant. Fac. Generadas |$ Fac. Generadas |Cant. Fac. Pendientes |$ Fac. Pendientes |Cant. Fac. Rechazadas |$ Fac. Rechazadas "
    
    grdFETU.FormatString = s$
    grdFETU.Row = 0
    grdFETU.Col = 0
    grdFETU.ColWidth(0) = 1100
    grdFETU.ColWidth(1) = 1200
    grdFETU.ColWidth(2) = 1800
    grdFETU.ColWidth(3) = 1800
    grdFETU.ColWidth(4) = 1800
    grdFETU.ColWidth(5) = 1800
    grdFETU.ColWidth(6) = 1800
    grdFETU.ColWidth(7) = 1800
    grdFETU.ColWidth(8) = 1800
    If OptGrupo(1).Value = True Then
        grdFETU.ColWidth(9) = 1800
    End If
    
    For i = 0 To grdFETU.Cols - 1
        grdFETU.Col = i
        grdFETU.CellFontBold = 1
    Next i
    
End Function

Private Sub OptGrupo_Click(index As Integer)
    DIB_GRILLA
End Sub

Private Sub OptRango_Click(index As Integer)
    If OptRango(0).Value = True Then
        Txtnumini.Enabled = False
        Txtnumini.Visible = False
        Txtnumini.text = ""
        Txtnumfin.Enabled = False
        Txtnumfin.Visible = False
        Txtnumfin.text = ""
        
        Mskfecini.Enabled = True
        Mskfecfin.Enabled = True
        Mskfecini.Visible = True
        Mskfecfin.Visible = True
        
        lblInicio.Caption = "Fecha Inicial"
        lblFinal.Caption = "Fecha Final"
        
    ElseIf OptRango(1).Value = True Then
        Txtnumini.Enabled = True
        Txtnumini.Visible = True
        Txtnumfin.Enabled = True
        Txtnumfin.Visible = True
        
        
        Mskfecini.Mask = ""
        Mskfecini.text = ""
        Mskfecini.Mask = "##/##/####"
        Mskfecfin.Mask = ""
        Mskfecfin.text = ""
        Mskfecfin.Mask = "##/##/####"
        Mskfecini.Enabled = False
        Mskfecfin.Enabled = False
        Mskfecini.Visible = False
        Mskfecfin.Visible = False
        
        lblInicio.Caption = "Num. Inicial"
        lblFinal.Caption = "Num. Final"
    End If
End Sub

Private Sub Txtnumfin_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub Txtnumfin_KeyPress(KeyAscii As Integer)
    If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Or (KeyAscii% = 47) Then
      KeyAscii% = 0
    End If
End Sub

Private Sub Txtnumfin_LostFocus()
    If CDbl(Txtnumfin.text) < 0 Then
        MsgBox "Debe ingresar un número de tasa de uso correcto", 48, "Advertencia"
        Txtnumfin.text = ""
    End If
End Sub

Private Sub Txtnumini_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub Txtnumini_KeyPress(KeyAscii As Integer)
    If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Or (KeyAscii% = 47) Then
      KeyAscii% = 0
    End If
End Sub

Private Sub Txtnumini_LostFocus()
    If CDbl(Txtnumini.text) < 0 Then
        MsgBox "Debe ingresar un número de tasa de uso correcto", 48, "Advertencia"
        Txtnumini.text = ""
    End If
End Sub
