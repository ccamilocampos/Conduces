VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Cuaturno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadre de Caja Por Turno"
   ClientHeight    =   8265
   ClientLeft      =   2325
   ClientTop       =   525
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txthora 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   1335
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Txtnumpla 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   480
      Width           =   1335
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton btnconsig 
      Caption         =   "&Consignar"
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Ingresa la consignación"
      Top             =   7320
      Width           =   855
   End
   Begin VB.Frame frmconsig 
      Caption         =   "Consignaciones"
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
      TabIndex        =   67
      Top             =   6720
      Width           =   11775
      Begin VB.CommandButton btnañadir 
         Height          =   495
         Left            =   4080
         Picture         =   "Cuaturno.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtbancos 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   68
         ToolTipText     =   "Código del banco (doble click  para ayuda)"
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Mskvalor 
         Height          =   285
         Left            =   840
         TabIndex        =   71
         ToolTipText     =   "Valor a consignar"
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconsigna 
         Height          =   975
         Left            =   4920
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox mskconsi 
         Height          =   285
         Left            =   840
         TabIndex        =   76
         ToolTipText     =   "Valor a consignar"
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Total:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   72
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Banco:"
         BeginProperty Font 
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
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblbancos 
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
         Left            =   2160
         TabIndex        =   69
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.ComboBox Cmbcaja 
      Height          =   315
      ItemData        =   "Cuaturno.frx":0884
      Left            =   10200
      List            =   "Cuaturno.frx":0886
      Style           =   2  'Dropdown List
      TabIndex        =   63
      ToolTipText     =   "Estado de la empresa"
      Top             =   480
      Width           =   1335
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ingresa el registro a la Base de Datos"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Txtobservacion 
      Height          =   285
      Left            =   120
      MaxLength       =   254
      TabIndex        =   1
      Top             =   6240
      Width           =   11775
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   6000
      TabIndex        =   23
      Top             =   960
      Width           =   5895
      Begin VB.TextBox Txtalcinf 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Txtalcori 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Txtalctra 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Txtnumneg 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txtnumanu 
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
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Txtnummul 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Msktotmul 
         Height          =   285
         Left            =   4440
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   600
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
      Begin MSMask.MaskEdBox Mskalcori 
         Height          =   285
         Left            =   4440
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1680
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
      Begin MSMask.MaskEdBox Mskalctra 
         Height          =   285
         Left            =   4440
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2040
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
      Begin MSMask.MaskEdBox Mskparqueadero 
         Height          =   285
         Left            =   4440
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   3840
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
      Begin MSMask.MaskEdBox MskMulElu 
         Height          =   285
         Left            =   1560
         TabIndex        =   80
         TabStop         =   0   'False
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
      Begin MSMask.MaskEdBox MskMulEva 
         Height          =   285
         Left            =   4440
         TabIndex        =   82
         TabStop         =   0   'False
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
      Begin MSMask.MaskEdBox MskMulPer 
         Height          =   285
         Left            =   1560
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1320
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
      Begin MSMask.MaskEdBox MskMulEvaTra 
         Height          =   285
         Left            =   4440
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
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
      Begin MSMask.MaskEdBox MskTag 
         Height          =   285
         Left            =   1560
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2880
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
      Begin MSMask.MaskEdBox MskVrCupAlc 
         Height          =   285
         Left            =   1560
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   3360
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
      Begin MSMask.MaskEdBox MskVigencia 
         Height          =   285
         Left            =   4440
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   2880
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
      Begin MSMask.MaskEdBox mskcredito 
         Height          =   285
         Left            =   4440
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   3360
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
      Begin MSMask.MaskEdBox Mskalcinf 
         Height          =   285
         Left            =   4440
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   2400
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
      Begin MSMask.MaskEdBox MskVrCupAlcInfluencia 
         Height          =   285
         Left            =   1560
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   3840
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
         Caption         =   "VrCupoAlc Inf.$"
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
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   123
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alcoh.Influencia"
         BeginProperty Font 
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
         Index           =   45
         Left            =   120
         TabIndex        =   117
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.Alc.Influen$"
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
         Height          =   255
         Index           =   46
         Left            =   3000
         TabIndex        =   116
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Total Credito  $"
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
         Height          =   255
         Index           =   23
         Left            =   3000
         TabIndex        =   105
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Total Vigencia $"
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
         Height          =   255
         Index           =   35
         Left            =   3000
         TabIndex        =   97
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "VrCupoAlc Ori$"
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
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   95
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Total TAG $"
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
         Height          =   255
         Index           =   36
         Left            =   120
         TabIndex        =   88
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "EvasTra.Multas$"
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
         Height          =   255
         Index           =   34
         Left            =   3000
         TabIndex        =   87
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Perdida Multas$"
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
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   85
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Evasión Multas$"
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
         Height          =   255
         Index           =   32
         Left            =   3000
         TabIndex        =   83
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Elusión Multas$"
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
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   81
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.Permanencia$"
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
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   62
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Alcohol.Origen "
         BeginProperty Font 
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
         TabIndex        =   60
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.AlcOrigen $"
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
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   59
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alcohol.Transito"
         BeginProperty Font 
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
         TabIndex        =   58
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.Alc.Transito$"
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
         Height          =   255
         Index           =   28
         Left            =   3000
         TabIndex        =   57
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "No. Anuladas"
         BeginProperty Font 
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
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "No. Multas"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Total Multas$"
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
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No. Negociadas"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   5775
      Begin VB.TextBox TxtCuposInfluencia 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Txtnuminf 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox TxtCupTraAut 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtCuposTransito 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtNumTransitoNorte 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TxtCuposOrigen 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txttiqcre 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Msktotori 
         Height          =   285
         Left            =   4320
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2040
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
      Begin VB.TextBox Txtnumtra 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Txtnumori 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Txttotal 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Txttiqanu 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txttottiq 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txttiqfin 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Txttiqini 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Msktottra 
         Height          =   285
         Left            =   4320
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2760
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
      Begin MSMask.MaskEdBox Msktotcon 
         Height          =   285
         Left            =   1560
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3480
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
      Begin MSMask.MaskEdBox MskVrCupTasaOrigen 
         Height          =   285
         Left            =   1560
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3840
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
      Begin MSMask.MaskEdBox MskTotTranNorte 
         Height          =   285
         Left            =   4320
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2400
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
      Begin MSMask.MaskEdBox MskVrCupTasaTransito 
         Height          =   285
         Left            =   4320
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3840
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
      Begin MSMask.MaskEdBox Msktotinf 
         Height          =   285
         Left            =   4320
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3120
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
      Begin MSMask.MaskEdBox MskVrCupTasaInfluencia 
         Height          =   285
         Left            =   4320
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   3480
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
         Caption         =   "VrCupTasInf$:"
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
         Height          =   255
         Index           =   50
         Left            =   3000
         TabIndex        =   120
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cupos Influen:"
         BeginProperty Font 
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
         Index           =   49
         Left            =   3000
         TabIndex        =   119
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tot. Influen.$:"
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
         Height          =   255
         Index           =   47
         Left            =   3000
         TabIndex        =   113
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No. Influencia:"
         BeginProperty Font 
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
         Index           =   48
         Left            =   120
         TabIndex        =   111
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "VrCupoTasaTra$:"
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
         Height          =   255
         Index           =   44
         Left            =   3000
         TabIndex        =   109
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "CuposTran.Aut:"
         BeginProperty Font 
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
         Index           =   43
         Left            =   3000
         TabIndex        =   107
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Cupos Transito:"
         BeginProperty Font 
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
         Index           =   42
         Left            =   120
         TabIndex        =   103
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.Tran.Norte$:"
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
         Height          =   255
         Index           =   41
         Left            =   3000
         TabIndex        =   101
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No Tran.Norte:"
         BeginProperty Font 
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
         Index           =   40
         Left            =   120
         TabIndex        =   99
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cupos Origen:"
         BeginProperty Font 
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
         Index           =   38
         Left            =   3000
         TabIndex        =   93
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Vr CupoTasaOri$:"
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
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   91
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tiq. Créditos:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tot.Transito$:"
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
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   38
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Total Tasas  $:"
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
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tot. Origen$:"
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
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No Transito:"
         BeginProperty Font 
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
         TabIndex        =   34
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "No. Origen:"
         BeginProperty Font 
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
         TabIndex        =   33
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total Tiquetes:"
         BeginProperty Font 
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
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tiq. Anulados:"
         BeginProperty Font 
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
         Left            =   3000
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
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
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tique. Final:"
         BeginProperty Font 
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
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tique. Inicial:"
         BeginProperty Font 
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
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   1215
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Cierre de Turno"
      Top             =   3960
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Borra datos del Formulario"
      Top             =   4800
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar Informacion en la Base de Datos"
      Top             =   3120
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
      Left            =   12120
      Picture         =   "Cuaturno.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir de la Forma Actual"
      Top             =   5640
      Width           =   855
   End
   Begin MSMask.MaskEdBox Mskbasmon 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
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
   Begin MSMask.MaskEdBox Msktotgen 
      Height          =   285
      Left            =   5640
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   13080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin MSMask.MaskEdBox Mskcaja 
      Height          =   285
      Left            =   9840
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msksalban 
      Height          =   285
      Left            =   9840
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   -120
      X2              =   13080
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   -120
      X2              =   12000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo x Consignar $"
      BeginProperty Font 
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
      Index           =   24
      Left            =   7920
      TabIndex        =   79
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   10200
      TabIndex        =   64
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Neto Caja Fuerte $"
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
      Height          =   255
      Index           =   22
      Left            =   7920
      TabIndex        =   51
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones:"
      BeginProperty Font 
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
      TabIndex        =   45
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Base Moneda $"
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
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   44
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Total General $"
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
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   43
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "No. Planilla:"
      BeginProperty Font 
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
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   1335
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
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Turno Activo:"
      BeginProperty Font 
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
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   -120
      X2              =   12960
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   12000
      X2              =   12000
      Y1              =   0
      Y2              =   8160
   End
End
Attribute VB_Name = "Cuaturno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hora
Dim horapet1
Dim horciet3
Dim no
Dim totanucon
Dim totanumul
Dim L_Buscar As String
Dim L_CiTElusion As String
Dim L_CiTEvasion As String
Dim L_CiTPerdida As String
Dim L_CiTEvasTra As String
Dim L_CiTExcreta As String
Dim L_Todas_Multas As String
  

Private Sub btnañadir_Click()
    On Error GoTo err_btnañadir_Click
    
    If Trim(txtbancos.text) = "" Then
        MsgBox "Falta Ingresar El Banco", 48, "Advertencia"
        txtbancos.SetFocus
        Exit Sub
    End If
    If Trim(Mskvalor.ClipText) = "" Then
        MsgBox "Falta Ingresar El Valor A Consignar", 48, "Advertencia"
        Mskvalor.SetFocus
        Exit Sub
    End If
    
    'jk 04/17/2013 - Validar Saldo en Bancos
    If Trim(msksalban.ClipText) = "" Then
        MsgBox "Falta el Valor del Saldo en Bancos", 48, "Advertencia"
        Exit Sub
    End If
    'fin jk
    
    If CDbl(Mskvalor.text) > CDbl(msksalban.text) Then
       MsgBox "El Valor Ha Consignar Es Mayor" & Chr(13) & "        Del Saldo En Bancos", 48, "Advertencia"
       Mskvalor.SetFocus
       Exit Sub
    End If
    
    no = "N"
    
    entry = Txtfecha.text + Chr(9) + Txthora.text + Chr(9) + _
            Cmbcaja.text + Chr(9) + Txtturno.text + Chr(9) + _
            Txtoperario.text + Chr(9) + txtbancos.text + Chr(9) + _
            Mskvalor.text + Chr(9) + no
            
    Grdconsigna.AddItem entry
    If mskconsi.text = "" Then
       mskconsi.text = CDbl(Mskvalor.text)
    Else
       mskconsi.text = CDbl(mskconsi.text) + CDbl(Mskvalor.text)
    End If
    
    If msksalban.text = "" Then
       msksalban.text = CDbl(Mskvalor.text)
    Else
       msksalban.text = msksalban.text - Mskvalor.text
    End If
    
    txtbancos.text = ""
    lblbancos.Caption = ""
    
    Mskvalor.text = ""
    On Error Resume Next
    If msksalban.text = "0" Then
       btnconsig.SetFocus
    Else
       txtbancos.SetFocus
    End If
    
    
    Exit Sub
err_btnañadir_Click:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnbuscar_Click()
   no = "N"
   On Error GoTo Errores
   
   MskMulElu.text = "0"
   MskMulEva.text = "0"
   MskMulPer.text = "0"
   MskMulEvaTra.text = "0"
   MskTag.text = "0"
   Msktotmul.text = "0"
   
   If Cmbcaja.Enabled = True Then
      If Cmbcaja.text = "" Then
         MsgBox "Falta Seleccionar La Caja", 48, "Advertencia"
         Cmbcaja.Enabled = True
         Cmbcaja.SetFocus
         Exit Sub
      End If
      fqry01 = "SELECT * " _
               & "FROM cocaja " _
              & "WHERE canomcaj = '" & Cmbcaja.text & "' "
             
      Set Ftb02 = Dbs.Execute(fqry01)
      If Ftb02.EOF = False Then
         G_maxtur = Ftb02!canumtur
      End If
      Ftb02.Close
   End If
   'wecc nuevo para el manejo de bancos y efectivo 10/02/2002
   
   Traer_Consig
   'wecc 11/14/2002
   
   'If Val(G_turno) < Val(G_maxtur) Then
   If Val(G_Turno) < 4 Then     'Cuando la el turno es menos al parametro de maximo de turnos
      fec = Txtfecha.text
      'Contar tiquetes
      fqry01 = "SELECT min(conumero) as minn, max(conumero) as maxn " _
               & "FROM coconduc " _
              & "WHERE cofecsal = TO_DATE('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND coturno = " & G_Turno & " "
             
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Not Ftb01!minn And Not Ftb01!maxn Then
            Txttiqini.text = Ftb01!minn
            Txttiqfin.text = Ftb01!maxn
            Txttottiq.text = Txttiqfin.text - Txttiqini.text + 1
         Else
            Txttiqini.text = "0"
            Txttiqfin.text = "0"
            Txttottiq.text = "0"
         End If
      End If
      Ftb01.Close
      'Tiquetes Anulados
      fqry01 = "SELECT count(conumero) as anu, sum(covalor) as valc " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY') " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND coestado = 'Anulado' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "'  "
              
      
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!anu <> "" And Ftb01!anu <> "0" Then
            Txttiqanu.text = Ftb01!anu
            totanucon = Ftb01!valc
         Else
            Txttiqanu.text = 0
            totanucon = "0"
         End If
      End If
      Txttotal.text = Txttottiq.text - CDbl(Ftb01!anu)
      Ftb01.Close
      'Contar Origen
      fqry01 = "SELECT count(conumero) as ori,sum(covalor) as valc " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY') " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'O' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "'  " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue <> 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!ori <> "" And Ftb01!ori <> "0" Then
            Txtnumori.text = Ftb01!ori
            Msktotori.text = Ftb01!valc
         Else
            Txtnumori.text = 0
            Msktotori.text = 0
         End If
      End If
      Ftb01.Close
      'Transito
      fqry01 = "SELECT count(conumero) as tran,sum(covalor) as valt " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'T' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "'  " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
       If Ftb01.EOF = False Then
         If Ftb01!tran <> "" And Ftb01!tran <> "0" Then
            Txtnumtra.text = Ftb01!tran
            Msktottra.text = Ftb01!valt
            Msktotcon.text = Val(Msktotori.text) + Val(Msktottra.text)
         Else
            Txtnumtra.text = 0
            Msktottra.text = 0
         End If
      End If
      Ftb01.Close
      
      'influencia
      'Contar Origen
      fqry01 = "SELECT count(conumero) as ori,sum(covalor) as valc " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY') " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'O' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "'  " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue = 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!ori <> "" And Ftb01!ori <> "0" Then
            Txtnuminf.text = Ftb01!ori
            Msktotinf.text = Ftb01!valc
         Else
            Txtnuminf.text = 0
            Msktotinf.text = 0
         End If
      End If
      Ftb01.Close
      
      'tiquetes en credito wecc 10/01/2002
      fqry01 = "SELECT count(conumero) as cre,sum(covalor) as valcre " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "'  " _
              & "  AND cocredit = 'S' " _
              & "  AND comodo = 'D' "
              
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!cre <> "" And Ftb01!cre <> "0" Then
            txttiqcre.text = Ftb01!cre
            mskcredito.text = Ftb01!valcre
         Else
            txttiqcre.text = 0
            mskcredito.text = 0
         End If
      End If
      Ftb01.Close
      '***************
      'wecc 04/25/2002
      valor = G_ValalcOri
      valinf = G_ValalcTra
      cero = 0
      'valor alcoholemia
      'fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc " _
              & "WHERE cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND cooritra = 'O' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND (covalalc = " & valor & " " _
              & "   OR covalalc = " & valinf & ") " _
              & "  AND conomcaj = '" & Cmbcaja.Text & "' "
              
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'O' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & " " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue <> 'Si' "

      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalcori.text = Ftb01!alco
            Mskalcori.text = Ftb01!vala
            
         Else
            Txtalcori.text = 0
            Mskalcori.text = 0
         End If
      End If
      Ftb01.Close
      
      'wecc 04/25/2002
      valor = G_ValalcTra
      'valor alcoholemia TRANSITO
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'T' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & "  " _
              & "  AND comodo = 'D' "

      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalctra.text = Ftb01!alco
            Mskalctra.text = Ftb01!vala
         Else
            Txtalctra.text = 0
            Mskalctra.text = 0
         End If
      End If
      Ftb01.Close
      
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' " _
              & "  AND cooritra = 'O' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & " " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue = 'Si' "

      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalcinf.text = Ftb01!alco
            Mskalcinf.text = Ftb01!vala
            
         Else
            Txtalcinf.text = 0
            Mskalcinf.text = 0
         End If
      End If
      Ftb01.Close
      '**********************
      
      'WECC MAY 06/2016
      'tasas de cupos origen
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' " _
            & " AND ruinflue <> 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposOrigen.text = Ftb01!sCantidadCupo
              MskVrCupTasaOrigen.text = Ftb01!sVrCupoTasa
              MskVrCupAlc.text = Ftb01!sVrCupoAlco
          Else
              TxtCuposOrigen.text = 0
              MskVrCupTasaOrigen.text = 0
              MskVrCupAlc.text = 0
          End If
      End If
      Ftb01.Close
      
      'WECC mAR/18/2021 CAU 3901
      'tasas de cupos influencia
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' " _
            & " AND ruinflue = 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposInfluencia.text = Ftb01!sCantidadCupo
              MskVrCupTasaInfluencia.text = Ftb01!sVrCupoTasa
              MskVrCupAlcInfluencia.text = Ftb01!sVrCupoAlco
          Else
              TxtCuposInfluencia.text = 0
              MskVrCupTasaInfluencia.text = 0
              MskVrCupAlcInfluencia.text = 0
          End If
      End If
      Ftb01.Close
      
      'tasas de cupos transito
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'T' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposTransito.text = Ftb01!sCantidadCupo
              MskVrCupTasaTransito.text = Ftb01!sVrCupoTasa
              MskVrCupAlc.text = CDbl(MskVrCupAlc.text) + Ftb01!sVrCupoAlco
          Else
              TxtCuposTransito.text = 0
              MskVrCupTasaTransito.text = 0
          End If
      End If
      Ftb01.Close
      '******************************************************
      '  wecc feb 03/2017 CAU No 14098
      fqry01 = "SELECT count(tnnumero) as qCantTransito,sum(tnvalor) as qValorTransito " _
               & "FROM cotransitonorte " _
              & "WHERE tnfechatasa = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND tncajatasa = '" & Cmbcaja.text & "' " _
              & "  AND tnturnotasa = " & G_Turno & " " _
              & "  AND tncerradotasa = '" & no & "' " _
              & "  AND tnestado = 'C' " _
              & "  AND tnmodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
       If Ftb01.EOF = False Then
         If Ftb01!qCantTransito <> "" And Ftb01!qCantTransito <> "0" Then
            TxtNumTransitoNorte.text = Ftb01!qCantTransito
            MskTotTranNorte.text = Ftb01!qValorTransito
            Msktotcon.text = CDbl(Msktotori.text) + CDbl(Msktottra.text) + CDbl(MskTotTranNorte.text)
         Else
            TxtNumTransitoNorte.text = 0
            MskTotTranNorte.text = 0
         End If
      End If
      Ftb01.Close
      
      'cupos de caja automatica
      fqry01 = "SELECT count(tnnumero) as qCantTransito,sum(tnvalor) as qValorTransito " _
               & "FROM cotransitonorte " _
              & "WHERE tnfechatasa = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND tncajatasa = '" & Cmbcaja.text & "' " _
              & "  AND tnturnotasa = " & G_Turno & " " _
              & "  AND tncerradotasa = '" & no & "' " _
              & "  AND tnestado = 'C' " _
              & "  AND tnmodo = 'P' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
       If Ftb01.EOF = False Then
         If Ftb01!qCantTransito <> "" And Ftb01!qCantTransito <> "0" Then
            TxtCupTraAut.text = Ftb01!qCantTransito
            MskVrCupTasaTransito.text = CDbl(MskVrCupTasaTransito.text) + Ftb01!qValorTransito
         Else
            TxtCupTraAut.text = "0"
         End If
      End If
      Ftb01.Close
      
      
      
      '************************************
      'wecc 08/26/2010
      'se parcha este qry porque el valor del vparqueo se va a generar de la tabla coparqueo
      
      'fqry01 = "SELECT sum(coparque) as valpar " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.Text & "' " _
              & "  AND coestado = 'Activo' " _
              & "  AND coturno = " & G_turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = '" & no & "'  "
      
      fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueop " _
              & "WHERE pafecpag = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND pacaja= '" & Cmbcaja.text & "' " _
              & "  AND paestado = 'C' " _
              & "  AND paturno = " & G_Turno & " " _
              & "  AND pacerrado = '" & no & "' " _
              & "  AND patipo = 'P' "
              
      '************************************
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
            Mskparqueadero.text = Ftb01!valpar
         Else
            Mskparqueadero.text = 0
         End If
      End If
      Ftb01.Close
      
      fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueop " _
              & "WHERE pafecpag = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND pacaja= '" & Cmbcaja.text & "' " _
              & "  AND paestado = 'C' " _
              & "  AND paturno = " & G_Turno & " " _
              & "  AND pacerrado = '" & no & "' " _
              & "  AND patipo = 'V' "
              
      '************************************
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
            MskVigencia.text = Ftb01!valpar
         Else
            MskVigencia.text = 0
         End If
      End If
      Ftb01.Close
      'abono de multas
      'wecc 08/05/2002
      
      
      'tabla de abonos todas las multas
      'se van a insertar en la tabla de abonos
      'wecc esto con el fin de hacer mas facil la busqueda
      
      'If G_BaseServidor = True Then
         fqry01 = "SELECT DISTINCT(abnumero) as nummul " _
                  & "FROM opabonos " _
                 & "WHERE abfecha = TO_DATE('" & fec & "','MM/DD/YYYY') " _
                   & "AND abnomcaj = '" & Cmbcaja.text & "' " _
                   & "AND abturno = " & G_Turno & " " _
                   & "AND abcerrado = '" & no & "' " _
                   & "AND abestado = 'Activo' "
        
         Set Ftb02 = Dbs.Execute(fqry01)
         If Ftb02.EOF = False Then
           'ftb02.MoveLast
           'ftb02.MoveFirst
           'Txtnummul.Text = Abs(ftb02.RecordCount)
            i = 0
            Ftb02.MoveLast
            Ftb02.MoveFirst
            Txtnummul.text = Ftb02.RecordCount
            
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
                  'Case 5
                     'sCodigo = L_CiTExcreta
               End Select
            
               fqry01 = "SELECT SUM(abvalor) as valt " _
                        & "FROM opabonos, opmultas " _
                       & "WHERE abfecha = TO_DATE('" & fec & "','MM/DD/YYYY') " _
                         & "AND abnomcaj = '" & Cmbcaja.text & "' " _
                         & "AND munumero = abnumero  " _
                         & "AND abturno = " & G_Turno & " " _
                         & "AND abcerrado = '" & no & "'  " _
                         & "AND mucodinf IN (" & sCodigo & ") " _
                         & "AND abestado = 'Activo' "
               Set Ftb02 = Dbs.Execute(fqry01)
            
               If Not Ftb02.EOF Then
                  If Not IsNull(Ftb02!valt) Then
                     Select Case k
                        Case 1
                           MskMulElu.text = Ftb02!valt
                        Case 2
                           MskMulEva.text = Ftb02!valt
                        Case 3
                           MskMulPer.text = Ftb02!valt
                        Case 4
                           MskMulEvaTra.text = Ftb02!valt
                        'Case 5
                           'MskTag.text = Ftb02!valt
                     End Select
                  End If
               End If
               Ftb02.Close
            
            Next k
            
            fqry01 = "SELECT SUM(abvalor) as valt " _
                     & "FROM opabonos, opmultas " _
                    & "WHERE abfecha = TO_DATE('" & fec & "','MM/DD/YYYY') " _
                      & "AND abnomcaj = '" & Cmbcaja.text & "' " _
                      & "AND munumero = abnumero  " _
                      & "AND abturno = " & G_Turno & " " _
                      & "AND abcerrado = '" & no & "'  " _
                      & "AND mucodinf NOT IN (" & L_Todas_Multas & ") " _
                      & "AND abestado = 'Activo' "
            
            Set Ftb02 = Dbs.Execute(fqry01)
         
            If Not Ftb02.EOF Then
               If Not IsNull(Ftb02!valt) Then
                  Msktotmul.text = Ftb02!valt
               End If
            End If
'            Do While Not Ftb02.EOF
'               i = i + 1
'               Ftb02.MoveNext
'            Loop
'            Ftb02.Close
'            Txtnummul.Text = i
'           fqry01 = "SELECT SUM(abvalor) as valt, mucodinf " _
'                    & "FROM opabonos, opmultas " _
'                   & "WHERE abfecha = TO_DATE('" & fec & "','MM/DD/YYYY') " _
'                     & "AND abnomcaj = '" & Cmbcaja.Text & "' " _
'                     & "AND munumero = abnumero  " _
'                     & "AND abturno = " & G_Turno & " " _
'                     & "AND abcerrado = '" & no & "'  " _
'                & "GROUP BY mucodinf"
'
'            Set Ftb02 = Dbs.Execute(fqry01)
'            'If Not ftb02.EOF Then
'            Do While Not Ftb02.EOF
'               Qry = " SELECT dsdes " _
'                      & "FROM gesuptip, gedetsuptip " _
'                     & "WHERE stdes = 'INFRACCIONES CIERRE TURNO' " _
'                       & "AND stcodtip = dscodtip " _
'                       & "AND dscoddet = '" & Ftb02(1) & "'"
'
'               Set Ftb01 = Dbs.Execute(Qry)
'
'               If Not Ftb01.EOF Then
'                  Mskmulotr.Text = CDbl(Mskmulotr.Text) + CDbl(Ftb02!valt)
'               Else
'                  Msktotmul.Text = CDbl(Msktotmul.Text) + CDbl(Ftb02!valt)
'               End If
'               Ftb01.Close
'               Ftb02.MoveNext
'            Loop
'            Ftb02.Close
            'End If
        Else
           Ftb02.Close
           Txtnummul.text = "0"
           Msktotmul.text = "0"
           
        End If
        
        '************************************************************
        'Sebastián Rondon - 21 de Octubre de 2022
        'Lo que antes se conocia como el concepto EXCRETA pasa ahora a ser TAG
        'Además este se empieza a cobrar directamente en la Tasa de Uso
        
        fqry01 = "SELECT sum(tavalor) as valtag " _
                    & "FROM optag " _
                    & "WHERE tafecpag = to_date('" & fec & "', 'MM/DD/YYYY') " _
                    & "AND tanomcaj = '" & Cmbcaja.text & "' " _
                    & "AND taestado = 'A' " _
                    & "AND tacerrado = 'N' " _
                    & "AND taturno = " & G_Turno & " "
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
            If Ftb01!valtag <> "" And Ftb01!valtag <> "0" Then
                MskTag.text = Ftb01!valtag
            Else
                MskTag.text = "0"
            End If
        End If
        
        '*****************************************************************
        
      
        fqry01 = "SELECT count(munumero) as anul, sum(munetpag) as valt " _
                 & "FROM opmultas " _
                & "WHERE mufecsis = to_date('" & fec & "','MM/DD/YYYY') " _
                & "  AND munomcaj = '" & Cmbcaja.text & "' " _
                  & "AND muturno = " & G_Turno & " " _
                  & "AND muestado = 'Anulado' " _
                  & "AND mucerrado = '" & no & "'  "

              
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
           'If Not IsNull(Ftb01!valt) Then
               If Ftb01!anul <> "" And Ftb01!anul <> "0" Then
                  Txtnumanu.text = Ftb01!anul
                  totanumul = Ftb01!valt
               Else
                  Txtnumanu.text = "0"
                  totanumul = "0"
               End If
           'End If
         End If
     
        fqry01 = "SELECT count(munumero) as neg " _
                 & "FROM opmultas " _
                & "WHERE mufecsis = to_date('" & fec & "','MM/DD/YYYY') " _
                  & "AND munomcaj = '" & Cmbcaja.text & "' " _
                  & "AND muturno = " & G_Turno & " " _
                  & "AND muestado = 'Pagada' " _
                  & "AND mudescue <> 0 " _
                  & "AND mucerrado = '" & no & "'  " _

              
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
           If Ftb01!neg <> "" And Ftb01!neg <> "0" Then
              Txtnumneg.text = Ftb01!neg
           Else
              Txtnumneg.text = "0"
           End If
        End If
      'Else
      '   Txtnummul.Text = "0"
      '   Msktotmul.Text = "0"
      '   Txtnumanu.Text = "0"
      '   totanumul = "0"
      '   Txtnumneg.Text = "0"
      'End If
      Msktotcon.text = CDbl(Msktotori.text) + CDbl(Msktottra.text) + CDbl(Msktotinf.text)
      'Mskbasmon.SetFocus
   
   'ElseIf G_turno = G_maxtur Then   'Cuando la global de turno es igual al parametro de turnos
   ElseIf G_Turno = 4 Then   'Cuando la global de turno es igual al parametro de turnos
      
      fqry01 = "SELECT to_char(cahoraap,'HH:MI AM') " _
               & "FROM cocaja " _
              & "WHERE canomcaj  = '" & Cmbcaja.text & "' "
            
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         horapet1 = Ftb01(0)
      End If
      
      fec = Txtfecha.text
      'wecc para manejar cuan las cajas tienen menos de 3 turnos
      'con el fin de no restarle un dia a la fecha
      
      'wecc 11/14/2002 se parcha xque cambiaron las horas de los turno
      'turno 1 10 pm a 06 am
      'turno 2 06 am a 02 pm
      'turno 3 02 pm a 10 pm
      'If G_maxtur >= 3 Then
         newfec = DateAdd("d", -1, Txtfecha.text)
      'ElseIf G_maxtur < 3 Then
      '   newfec = Txtfecha.Text
      'End If
      '************************
      fec1 = newfec
      horapet1 = Format(horapet1, "hh:mm")
      horciet3 = Format(Txthora.text, "hh:mm")
      horapet1 = horapet1
      horciet3 = horciet3
           
      'Contar tiquetes
      fqry01 = "SELECT min(conumero) as minn, max(conumero) as maxn " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Not Ftb01!minn And Not Ftb01!maxn Then
            Txttiqini.text = Ftb01!minn
            Txttiqfin.text = Ftb01!maxn
            Txttottiq.text = Txttiqfin.text - Txttiqini.text + 1
         Else
            Txttiqini.text = "0"
            Txttiqfin.text = "0"
            Txttottiq.text = "0"
         End If
      End If
      Ftb01.Close
      'Tiquetes Anulados
      fqry01 = "SELECT count(conumero) as anu, sum(covalor) as valt " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'   AND coestado = 'Anulado' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!anu <> "" And Ftb01!anu <> "0" Then
            Txttiqanu.text = Ftb01!anu
            totanucon = Ftb01!valt
         Else
            Txttiqanu.text = 0
            totanucon = "0"
         End If
      End If
      Txttotal.text = Txttottiq.text - CDbl(Ftb01!anu)
      Ftb01.Close
      'Contar Origen
      fqry01 = "SELECT count(conumero) as ori,sum(covalor) as valc " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'  AND cooritra = 'O' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND coestado = 'Activo' " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue <> 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!ori <> "" And Ftb01!ori <> "0" Then
            Txtnumori.text = Ftb01!ori
            Msktotori.text = Ftb01!valc
         Else
            Txtnumori.text = 0
            Msktotori.text = 0
         End If
      End If
      'Transito
      fqry01 = "SELECT count(conumero) as tran,sum(covalor) as valt " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'  AND cooritra = 'T' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND coestado = 'Activo' " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!tran <> "" And Ftb01!tran <> "0" Then
            Txtnumtra.text = Ftb01!tran
            Msktottra.text = Ftb01!valt
            Msktotcon.text = Val(Msktotori.text) + Val(Msktottra.text)
         Else
            Txtnumtra.text = 0
            Msktottra.text = 0
         End If
      End If
      
      'Contar influencia
      fqry01 = "SELECT count(conumero) as ori,sum(covalor) as valc " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'  AND cooritra = 'O' " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND coestado = 'Activo' " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = '" & no & "'  " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue = 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!ori <> "" And Ftb01!ori <> "0" Then
            Txtnuminf.text = Ftb01!ori
            Msktotinf.text = Ftb01!valc
         Else
            Txtnuminf.text = 0
            Msktotinf.text = 0
         End If
      End If
      
      
      'tiquetes en credito wecc 10/01/2002
      fqry01 = "SELECT count(conumero) as cred,sum(covalor) as valcred " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'  AND coturno = " & G_Turno & " " _
              & "  AND coestado = 'Activo' " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = 'S' " _
              & "  AND comodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!cred <> "" And Ftb01!cred <> "0" Then
            txttiqcre.text = Ftb01!cred
            mskcredito.text = Ftb01!valcred
         Else
            txttiqcre.text = 0
            mskcredito.text = 0
         End If
      End If
      Ftb01.Close
      
      '***************
      'wecc 04/25/2002
      valor = G_ValalcOri
      valinf = G_ValalcTra
      cero = 0
      'valor alcoholemia origen
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'   AND cooritra = 'O' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & " " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue <> 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalcori.text = Ftb01!alco
            Mskalcori.text = Ftb01!vala
            
         Else
            Txtalcori.text = 0
            Mskalcori.text = 0
         End If
      End If
      Ftb01.Close
      'wecc 04/25/2002
      valor = G_ValalcTra
      cero = 0
      'valor alcoholemia TRANSITO
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "' AND cooritra = 'T' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & " " _
              & "  AND comodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalctra.text = Ftb01!alco
            Mskalctra.text = Ftb01!vala
         Else
            Txtalctra.text = 0
            Mskalctra.text = 0
         End If
      End If
      Ftb01.Close
      
      'valor alcoholemia influencia
      fqry01 = "SELECT count(conumero) as alco,sum(covalalc) as vala " _
               & "FROM coconduc, opruta " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.text & "'   AND cooritra = 'O' " _
              & "  AND (coestado = 'Activo' OR coestado = 'Cambio') " _
              & "  AND coturno = " & G_Turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND covalalc <> " & cero & " " _
              & "  AND comodo = 'D' " _
              & "  AND cocodrut = rucodigo " _
              & "  AND ruinflue = 'Si' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!alco <> "" And Ftb01!alco <> "0" Then
            Txtalcinf.text = Ftb01!alco
            Mskalcinf.text = Ftb01!vala
            
         Else
            Txtalcinf.text = 0
            Mskalcinf.text = 0
         End If
      End If
      Ftb01.Close
      
      '*********************
      
      
      'WECC MAY 06/2016
      'tasas de cupos origen
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
            & " AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' " _
            & " AND ruinflue <> 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposOrigen.text = Ftb01!sCantidadCupo
              MskVrCupTasaOrigen.text = Ftb01!sVrCupoTasa
              MskVrCupAlc.text = Ftb01!sVrCupoAlco
          Else
              TxtCuposOrigen.text = 0
              MskVrCupTasaOrigen.text = 0
              MskVrCupAlc.text = 0
          End If
      End If
      Ftb01.Close
      
      'WECC mAR/18/2021 CAU 3901
      'tasas de cupos influencia
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
            & " AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' " _
            & " AND ruinflue = 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposInfluencia.text = Ftb01!sCantidadCupo
              MskVrCupTasaInfluencia.text = Ftb01!sVrCupoTasa
              MskVrCupAlcInfluencia.text = Ftb01!sVrCupoAlco
          Else
              TxtCuposInfluencia.text = 0
              MskVrCupTasaInfluencia.text = 0
              MskVrCupAlcInfluencia.text = 0
          End If
      End If
      Ftb01.Close
      
      'tasas de cupos transito
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc " _
           & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
            & " AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & Cmbcaja.text & "' " _
            & " AND cooritra = 'T' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND coturno = " & G_Turno & " " _
            & " AND cocerrado = '" & no & "' " _
            & " AND comodo = 'P' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              TxtCuposTransito.text = Ftb01!sCantidadCupo
              MskVrCupTasaTransito.text = Ftb01!sVrCupoTasa
              MskVrCupAlc.text = CDbl(MskVrCupAlc.text) + Ftb01!sVrCupoAlco
          Else
              TxtCuposTransito.text = 0
              MskVrCupTasaTransito.text = 0
          End If
      End If
      Ftb01.Close
      
      '******************************************************
      '  wecc feb 03/2017 CAU No 14098
      fqry01 = "SELECT count(tnnumero) as qCantTransito,sum(tnvalor) as qValorTransito " _
               & "FROM cotransitonorte " _
              & "WHERE tnfechatasa >= to_date('" & fec1 & "','MM/DD/YYYY') " _
                & "AND tnfechatasa <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND tncajatasa = '" & Cmbcaja.text & "' " _
              & "  AND tnturnotasa = " & G_Turno & " " _
              & "  AND tncerradotasa = '" & no & "' " _
              & "  AND tnestado = 'C' " _
              & "  AND tnmodo = 'D' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
       If Ftb01.EOF = False Then
         If Ftb01!qCantTransito <> "" And Ftb01!qCantTransito <> "0" Then
            TxtNumTransitoNorte.text = Ftb01!qCantTransito
            MskTotTranNorte.text = Ftb01!qValorTransito
            Msktotcon.text = CDbl(Msktotori.text) + CDbl(Msktottra.text) + CDbl(MskTotTranNorte.text)
         Else
            TxtNumTransitoNorte.text = 0
            MskTotTranNorte.text = 0
         End If
      End If
      Ftb01.Close
      
      fqry01 = "SELECT count(tnnumero) as qCantTransito,sum(tnvalor) as qValorTransito " _
               & "FROM cotransitonorte " _
              & "WHERE tnfechatasa >= to_date('" & fec1 & "','MM/DD/YYYY') " _
                & "AND tnfechatasa <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND tncajatasa = '" & Cmbcaja.text & "' " _
              & "  AND tnturnotasa = " & G_Turno & " " _
              & "  AND tncerradotasa = '" & no & "' " _
              & "  AND tnestado = 'C' " _
              & "  AND tnmodo = 'P' "
      
      
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!qCantTransito <> "" And Ftb01!qCantTransito <> "0" Then
            TxtCupTraAut.text = Ftb01!qCantTransito
            MskVrCupTasaTransito.text = CDbl(MskVrCupTasaTransito.text) + Ftb01!qValorTransito
         Else
            TxtCupTraAut.text = "0"
         End If
      End If
      Ftb01.Close
      '************************************
      'wecc 08/26/2010
      'se parcha este qry porque el valor del vparqueo se va a generar de la tabla coparqueo
      
      'Parqueadero
      ' fqry01 = "SELECT sum(coparque) as valpar " _
               & "FROM coconduc " _
              & "WHERE cofecsal >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND cofecsal <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & Cmbcaja.Text & "'   AND coestado = 'Activo' " _
              & "  AND coturno = " & G_turno & " " _
              & "  AND cocerrado = '" & no & "' " _
              & "  AND cocredit = '" & no & "'  "
     
     'fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueo " _
              & "WHERE pafecpag BETWEEN to_date('" & fec1 & "','MM/DD/YYYY')  " _
              & "  AND to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND pacaja= '" & Cmbcaja.Text & "' " _
              & "  AND paestado = 'C' " _
              & "  AND paturno = " & G_Turno & " " _
              & "  AND pacerrado = '" & no & "' "
     
     fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueop " _
              & "WHERE pafecpag >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND pafecpag <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND pacaja= '" & Cmbcaja.text & "' " _
              & "  AND paestado = 'C' " _
              & "  AND paturno = " & G_Turno & " " _
              & "  AND pacerrado = '" & no & "' " _
              & "  AND patipo = 'P' "
     
     
     Set Ftb01 = Dbs.Execute(fqry01)
     If Not Ftb01.EOF Then
         If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
            Mskparqueadero.text = Ftb01!valpar
         Else

            Mskparqueadero.text = 0
         End If
      End If
      Ftb01.Close
      
      fqry01 = "SELECT sum(pavalor) as valpar " _
               & "FROM coparqueop " _
              & "WHERE pafecpag >= to_date('" & fec1 & "','MM/DD/YYYY') " _
              & "  AND pafecpag <= to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND pacaja= '" & Cmbcaja.text & "' " _
              & "  AND paestado = 'C' " _
              & "  AND paturno = " & G_Turno & " " _
              & "  AND pacerrado = '" & no & "' " _
              & "  AND patipo = 'V' "
     
     
     Set Ftb01 = Dbs.Execute(fqry01)
     If Not Ftb01.EOF Then
         If Ftb01!valpar <> "" And Ftb01!valpar <> "0" Then
            MskVigencia.text = Ftb01!valpar
         Else
            MskVigencia.text = 0
         End If
      End If
      Ftb01.Close
      
      'abono de multas
      'If G_BaseServidor = True Then
         fqry01 = "SELECT DISTINCT(abnumero) as nummul " _
                  & "FROM opabonos " _
                 & "WHERE abfecha >= TO_DATE('" & fec1 & "','MM/DD/YYYY') " _
                   & "AND abfecha <= TO_DATE('" & fec & "','MM/DD/YYYY') " _
                   & "  AND abnomcaj = '" & Cmbcaja.text & "' AND abturno = " & G_Turno & " " _
                   & "AND abcerrado = '" & no & "' " _
                   & "AND abestado = 'Activo' "
           
         Set Ftb02 = Dbs.Execute(fqry01)
         If Not Ftb02.EOF Then
            i = 0
'            Do While Not Ftb02.EOF
'               i = i + 1
'               Ftb02.MoveNext
'            Loop
'            Txtnummul.Text = i
            Ftb02.MoveLast
            Ftb02.MoveFirst
            Txtnummul.text = Ftb02.RecordCount
            '**********************************
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
                  'Case 5
                     'sCodigo = L_CiTExcreta
               End Select
            
               fqry01 = "SELECT SUM(abvalor) as valt " _
                        & "FROM opabonos, opmultas " _
                       & "WHERE abfecha BETWEEN TO_DATE('" & fec1 & "','MM/DD/YYYY') " _
                         & "AND TO_DATE('" & fec & "','MM/DD/YYYY') " _
                         & "AND abnomcaj = '" & Cmbcaja.text & "' " _
                         & "AND munumero = abnumero  " _
                         & "AND abturno = " & G_Turno & " " _
                         & "AND abcerrado = '" & no & "'  " _
                         & "AND mucodinf IN ('" & sCodigo & "') " _
                         & "AND abestado = 'Activo' "
               Set Ftb02 = Dbs.Execute(fqry01)
            
               If Not Ftb02.EOF Then
                  If Not IsNull(Ftb02!valt) Then
                     Select Case k
                        Case 1
                           MskMulElu.text = Ftb02!valt
                        Case 2
                           MskMulEva.text = Ftb02!valt
                        Case 3
                           MskMulPer.text = Ftb02!valt
                        Case 4
                           MskMulEvaTra.text = Ftb02!valt
                        'Case 5
                           'MskTag.text = Ftb02!valt
                     End Select
                  End If
               End If
               Ftb02.Close
            
            Next k
            
            fqry01 = "SELECT SUM(abvalor) as valt " _
                     & "FROM opabonos, opmultas " _
                    & "WHERE abfecha BETWEEN TO_DATE('" & fec1 & "','MM/DD/YYYY') " _
                      & "AND TO_DATE('" & fec & "','MM/DD/YYYY') " _
                      & "AND abnomcaj = '" & Cmbcaja.text & "' " _
                      & "AND munumero = abnumero  " _
                      & "AND abturno = " & G_Turno & " " _
                      & "AND abcerrado = '" & no & "'  " _
                      & "AND mucodinf NOT IN (" & L_Todas_Multas & ") " _
                      & "AND abestado = 'Activo' "
            Set Ftb02 = Dbs.Execute(fqry01)
         
            If Not Ftb02.EOF Then
               If Not IsNull(Ftb02!valt) Then
                  Msktotmul.text = Ftb02!valt
               End If
            End If
            '*********************************
            
            
'            fqry01 = "SELECT SUM(abvalor) as valt, mucodinf " _
'                     & "FROM opabonos, opmultas  " _
'                    & "WHERE abfecha >= TO_DATE('" & fec1 & "','MM/DD/YYYY') " _
'                      & "AND abfecha <= TO_DATE('" & fec & "','MM/DD/YYYY') " _
'                      & "AND abnomcaj = '" & Cmbcaja.Text & "' AND munumero = abnumero   " _
'                      & "AND abcerrado = '" & no & "' " _
'                      & "AND abturno = " & G_Turno & " GROUP BY mucodinf  "
'
'            Set Ftb02 = Dbs.Execute(fqry01)
'            'If Not ftb02.EOF Then
'            Do While Not Ftb02.EOF
'               Qry = " SELECT dsdes " _
'                      & "FROM gesuptip, gedetsuptip " _
'                     & "WHERE stdes = 'INFRACCIONES CIERRE TURNO' " _
'                       & "AND stcodtip = dscodtip " _
'                       & "AND dscoddet = '" & Ftb02(1) & "'"
'
'               Set Ftb01 = Dbs.Execute(Qry)
'
'               If Not Ftb01.EOF Then
'                  Mskmulotr.Text = CDbl(Mskmulotr.Text) + CDbl(Ftb02!valt)
'               Else
'                  Msktotmul.Text = CDbl(Msktotmul.Text) + CDbl(Ftb02!valt)
'               End If
'               Ftb01.Close
'               Ftb02.MoveNext
'            Loop
'            Ftb02.Close
               
               
            'End If
   
         Else
            Txtnummul.text = "0"
            Msktotmul.text = "0"
            Ftb02.Close
         End If
         
        '************************************************************
        'Sebastián Rondon - 21 de Octubre de 2022
        'Lo que antes se conocia como el concepto EXCRETA pasa ahora a ser TAG
        'Además este se empieza a cobrar directamente en la venta de Tasas de Uso
        
        fqry01 = "SELECT sum(tavalor) as valtag " _
                    & "FROM optag " _
                    & "WHERE tafecpag >= to_date('" & fec1 & "', 'MM/DD/YYYY') " _
                    & "AND tafecpag <= to_date('" & fec & "', 'MM/DD/YYYY') " _
                    & "AND tanomcaj = '" & Cmbcaja.text & "' " _
                    & "AND taestado = 'A' " _
                    & "AND tacerrado = 'N' " _
                    & "AND taturno = " & G_Turno & " "
        
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
            If Ftb01!valtag <> "" And Ftb01!valtag <> "0" Then
                MskTag.text = Ftb01!valtag
            Else
                MskTag.text = "0"
            End If
        End If
        
        '*****************************************************************
               
         
         fqry01 = "SELECT count(munumero) as anul, sum(munetpag) as valt " _
                  & "FROM opmultas " _
                 & "WHERE mufecsis >= to_date('" & fec1 & "','MM/DD/YYYY') " _
                   & "AND mufecsis <= to_date('" & fec & "','MM/DD/YYYY') " _
                   & "AND munomcaj = '" & Cmbcaja.text & "' AND muturno = " & G_Turno & " " _
                   & "AND muestado = 'Anulado' " _
                   & "AND mucerrado = '" & no & "' "
                 
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
            'If Not IsNull(Ftb01!valt) Then
               If Ftb01!anul <> "" And Ftb01!anul <> "0" Then
                  Txtnumanu.text = Ftb01!anul
                  totanumul = Ftb01!valt
               Else
                  Txtnumanu.text = "0"
                  totanumul = "0"
               End If
            'End If
         End If
         
         fqry01 = "SELECT count(munumero) as neg " _
                  & "FROM opmultas " _
                 & "WHERE mufecsis >= to_date('" & fec1 & "','MM/DD/YYYY') " _
                   & "AND mufecsis <= to_date('" & fec & "','MM/DD/YYYY') " _
                   & "AND munomcaj = '" & Cmbcaja.text & "' AND muturno = " & G_Turno & " " _
                   & "AND muestado = 'Pagada' " _
                   & "AND mudescue <> 0 " _
                   & "AND mucerrado = '" & no & "' "
                 
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
            If Ftb01!neg <> "" And Ftb01!neg <> "0" Then
               Txtnumneg.text = Ftb01!neg
            Else
               Txtnumneg.text = "0"
            End If
         End If
      
      'Else
      '   Txtnummul.Text = "0"
      '   Msktotmul.Text = "0"
      '   Txtnumanu.Text = "0"
      '   totanumul = "0"
      '   Txtnumneg.Text = "0"
      'End If
      
      Msktotcon.text = CDbl(Msktotori.text) + CDbl(Msktottra.text) + CDbl(Msktotinf.text)     '+ CDbl(Msktotalc.Text)
   End If
   
   
   If Cmbcaja.Enabled = False Then
      
      QRY = "SELECT caconsig " _
            & "FROM cocaja " _
           & "WHERE canomcaj = '" & Cmbcaja.text & "'"
           
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         efectivo = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(MskVigencia.text) + CDbl(Mskalcinf.text) + CDbl(MskTag.text)
         If CDbl(Tbl1(0)) <> CDbl(efectivo) Then
            MsgBox "El Valor Efectivo Es Diferente Al Valor Consignado", 48, "Advertencia"
            Btngrabar.Enabled = False
            Mskbasmon.text = "0"
            Msktotgen.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(mskcredito.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text) + CDbl(MskTag.text)
            Mskcaja.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) - CDbl(Mskbasmon.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text) + CDbl(MskTag.text)
            msksalban.text = CDbl(Mskcaja.text) - CDbl(mskconsi.text)
            txtbancos.SetFocus
            Exit Sub
         Else
            btnconsig.Enabled = False
            btnañadir.Enabled = False
            Btngrabar.Enabled = True
            msksalban.text = "0"
         End If
      
      End If
      
    Else
      btnconsig.Enabled = False
      btnañadir.Enabled = False
      Btngrabar.Enabled = False
      msksalban.text = "0"
    End If
    Mskbasmon.text = "0"
    Msktotgen.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(mskcredito.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTag.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text)
    Mskcaja.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) - CDbl(Mskbasmon.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTag.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text)
    On Error Resume Next
    Txtobservacion.SetFocus
    
    Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    
    Exit Sub
Resume
End Sub

Private Sub btnconsig_Click()
   On Error GoTo err_btnconsig_Click
   
   If Grdconsigna.rows > 1 Then
   
    If G_CierreAutomatico = 0 Then
        respuesta = MsgBox("Desea Consignar El Dinero", 32 + vbYesNo, "CONSIGNAR")
    End If
     
  
     If respuesta = vbNo Then
        Exit Sub
     End If
      
      On Error GoTo err_btnconsig_Click
      
     
     Dbs.BeginTrans
     On Error GoTo err_transaccion
      
      'se parcha por el manejo de los turnos
      'If G_turno < 3 Then
      
      
      If G_Turno < 4 Then
      
         QRY = "DELETE FROM codetban " _
              & "WHERE dbfecha = TO_DATE('" & Txtfecha.text & "','MM/DD/YYYY') " _
                & "AND dbturno = " & Txtturno.text & " " _
                & "AND dbcaja = '" & Cmbcaja.text & "'"
      
      ElseIf G_Turno = 4 Then
         fec3 = Txtfecha.text
         newfec1 = DateAdd("d", -1, Txtfecha.text)
         fec2 = newfec1
      
         QRY = "DELETE FROM codetban " _
              & "WHERE dbfecha >= to_date('" & fec2 & "','MM/DD/YYYY') " _
                & "AND dbfecha <= to_date('" & fec3 & "','MM/DD/YYYY')  " _
                & "AND dbturno = " & Txtturno.text & " " _
                & "AND dbcaja = '" & Cmbcaja.text & "'"
      End If
      
      
      Dbs.Execute QRY
      
      'graba a la bs del servidor
      
      '***************
      
      valtotal = 0
      
      For i = 1 To Grdconsigna.rows - 1
          Grdconsigna.Row = i
          Grdconsigna.Col = 0
          feccon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 1
          Horcon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 2
          cajcon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 3
          turcon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 4
          funccon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 5
          bancon = Trim(Grdconsigna.text)
          Grdconsigna.Col = 6
          VALCON = Trim(Grdconsigna.text)
          Grdconsigna.Col = 7
          CAJACERCONS = Trim(Grdconsigna.text)
          valtotal = CDbl(valtotal) + CDbl(VALCON)
              
              
          'graba en bd del servidor
          'Guardo_Servidor = "N"
          'If G_BaseServidor = True Then
          '  dbsX.Execute " INSERT INTO codetban " _
          '                       & "(dbfecha, dbhora,dbcaja, dbturno, dboperar, dbcodban, dbvalor, dbcacerr) " _
          '                 & "VALUES (#" & feccon & "#,#" & Format(horcon, "hh:mm") & "#,'" & cajcon & "', " & turcon & ",'" & funccon & "','" & bancon & "'," & VALCON & ", '" & CAJACERCONS & "')"
          '
          '  Guardo_Servidor = "S"
          'End If
          '***************************
              
          Dbs.Execute " INSERT INTO codetban " _
                               & "(dbfecha, dbhora,dbcaja, dbturno, dboperar, dbcodban, dbvalor, dbcacerr) " _
                         & "VALUES (TO_DATE('" & feccon & "','MM/DD/YYYY'),TO_DATE('" & Horcon & "','HH24:MI:SS'),'" & cajcon & "', " & turcon & ",'" & funccon & "','" & bancon & "'," & VALCON & ", '" & CAJACERCONS & "')"
          
          
          
      Next i
      
      Dbs.Execute "UPDATE cocaja " _
                  & "SET  caconsig =  " & valtotal & " " _
                & "WHERE canomcaj = '" & Pc & "' "
      
      
      Dbs.CommitTrans
      MsgBox "Consignacion Grabada Satisfactoriamente", 64, "O.K"
      On Error Resume Next
      If msksalban.text = "0" Then
         Btngrabar.Enabled = True
         Txtobservacion.SetFocus
         btnconsig.Enabled = False
      End If
      
   Else
      MsgBox "No Hay Registros Para Consignar", 48, "Advertencia"
   End If
   
   Exit Sub
err_btnconsig_Click:
   If Err.Number = 52 Then
      MsgBox "Se Ha Desconectado del Servidor", 48, "Desconectado"
      G_BaseServidor = False
      Resume Next
    End If
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
err_transaccion:

   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
err_btnconsig_Click_Dbx:
   G_BaseServidor = False
   Resume Next
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo err_Btngrabar
  
  QRY = "SELECT * " _
        & "FROM cocaja " _
       & "WHERE canomcaj = '" & Pc & "'"
  
   Set Tbl1 = Dbs.Execute(QRY)
  
  If Not Tbl1.EOF Then
     If Val(G_Turno) = 1 Then
        horcierre1 = "13:01"
        horcierre2 = "23:59"
        
        If CDate(Format(Txthora.text, "hh:mm")) > CDate(Format(horcierre1, "hh:mm")) And CDate(Format(Txthora.text, "hh:mm")) < CDate(Format(horcierre2, "hh:mm")) Then
           MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
           Exit Sub
        End If
        
        If CDate(Format(Txthora.text, "hh:mm")) < CDate(Format(Tbl1!caturno1, "hh:mm")) Then
           MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
           Exit Sub
        End If
     ElseIf Val(G_Turno) = 2 Then
        If CDate(Format(Txthora.text, "hh:mm")) < CDate(Format(Tbl1!caturno2, "hh:mm")) Then
           MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
           Exit Sub
        End If
     ElseIf Val(G_Turno) = 3 Then
        'horcierre1 = "22:01"
        'horcierre2 = "23:59"
        
        'If CDate(Format(Txthora.Text, "hh:mm")) > CDate(Format(horcierre1, "hh:mm")) And CDate(Format(Txthora.Text, "hh:mm")) < CDate(Format(horcierre2, "hh:mm")) Then
        '   MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
        '   Exit Sub
        'End If
                
        If CDate(Format(Txthora.text, "hh:mm")) < CDate(Format(Tbl1!caturno3, "hh:mm")) Then
           MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
           Exit Sub
        End If
      ElseIf Val(G_Turno) = 4 Then
         horcierre1 = "21:01"
         horcierre2 = "23:59"
        
         If CDate(Format(Txthora.text, "hh:mm")) > CDate(Format(horcierre1, "hh:mm")) And CDate(Format(Txthora.text, "hh:mm")) <= CDate(Format(horcierre2, "hh:mm")) Then
            MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
            Exit Sub
         End If
      
         If CDate(Format(Txthora.text, "hh:mm")) < CDate(Format(Tbl1!caturno4, "hh:mm")) Then
            MsgBox "No Es Hora de Cerrar Turno ", 48, "Advertencia"
            Exit Sub
         End If
      End If
  
  End If
  Tbl1.Close
  
  If Mskbasmon.text = "" Then
     MsgBox "Falta Ingresar La Base De Moneda", 48, "Advertencia"
     Mskbasmon.SetFocus
     Exit Sub
  End If
  
  If Trim(Txtobservacion.text) = "" Then
     MsgBox "Falta Ingresar El Campo Observaciones", 48, "Advertencia"
     Txtobservacion.SetFocus
     Exit Sub
  End If
   
   On Error GoTo err_Btngrabar
    
  If G_CierreAutomatico = 0 Then
    respuesta = MsgBox("Desea Cerrar El Turno", 32 + vbYesNo, "CERRAR")
  End If
  
  If respuesta = vbNo Then
     Exit Sub
  End If
    
  If Msktotgen.text <> "" Then
     
     Dbs.BeginTrans
     On Error GoTo Err_Transac
     hora = Format(Txthora.text, "hh:mm")
          
     If G_Turno = G_maxtur Then
        
        'se parcha porque todos los turnos quedan con con la fecha del cierre
        If G_maxtur = 4 Then
        
           sFechaTurno = DateAdd("d", -1, Txtfecha.text)
        Else
           sFechaTurno = Txtfecha.text
        End If
        
        Dbs.Execute "UPDATE cocaja " _
                    & " SET caciedia = 'N' " _
                   & "WHERE canomcaj = '" & Pc & "' "
     Else
      sFechaTurno = Txtfecha.text
     End If
     
     If Txtoperario.text = "" Then
        Txtoperario.text = "automatico"
     End If
     
     Dbs.Execute "INSERT INTO cocietur " _
     & "(cifecha, citurno, cinumini, cinumfin, " _
     & "cicodope, cihoraap, cianula, ciorigen, " _
     & "citransi, citorige, cittrans, ciinfrac, " _
     & "ciobserv, citotgen, cibasmon, cinumpla, " _
     & "citinfra, citanula, civalalcori, citvalalcori, " _
     & "civalalctra, citvalalctra, citparqu, cinomcaj, " _
     & "cicredit, citcredi, citotrinf, citelusion, " _
     & "citevasion, citevastra, citvigencia, cittag, " _
     & "cicupos, cicupvaltasa, cicupvalalco, citransitonorte, " _
     & "cittransitonorte, cicupostransito, cicupostransitoaut, cicupvaltasatransito, " _
     & "civalalcinf, citvalalcinf, ciinfluencia, citinfluencia, " _
     & "cicuposinfluencia, cicupvaltasainfluencia, cicupalcinfluencia) VALUES (to_date('" _
     & sFechaTurno & "','MM/DD/YYYY')," & Txtturno.text & ",'" & Txttiqini.text & "','" & Txttiqfin.text & "','" _
     & Txtoperario.text & "',to_date('" & hora & "','HH24:MI:SS')," & Txttiqanu.text & "," & Txtnumori.text & "," _
     & Txtnumtra.text & "," & Msktotori.text & "," & Msktottra.text & "," & Txtnummul.text & ",'" _
     & Txtobservacion.text & "'," & Msktotgen.text & "," & Mskbasmon.text & "," & Txtnumpla.text & "," _
     & Msktotmul.text & "," & Txtnumanu.text & ", " & Txtalcori.text & "," & Mskalcori.text & ", " _
     & Txtalctra.text & "," & Mskalctra.text & "," & Mskparqueadero.text & ",'" & Pc & "'," _
     & txttiqcre.text & "," & mskcredito.text & ", " & MskMulPer.text & ", " & MskMulElu.text & ", " _
     & MskMulEva.text & ", " & MskMulEvaTra.text & "," & MskVigencia.text & "," & MskTag.text & ", " _
     & TxtCuposOrigen.text & ",  " & CDbl(MskVrCupTasaOrigen.text) & "," & CDbl(MskVrCupAlc.text) & "," & TxtNumTransitoNorte.text & "," _
     & MskTotTranNorte.text & "," & TxtCuposTransito.text & "," & TxtCupTraAut.text & "," & MskVrCupTasaTransito.text & ", " _
     & Txtalcinf.text & "," & Mskalcinf.text & "," & Txtnuminf.text & "," & Msktotinf.text & "," & TxtCuposInfluencia.text & "," & MskVrCupTasaInfluencia.text & "," & MskVrCupAlcInfluencia.text & ")"
     
     cero = "0"
     Dbs.Execute "UPDATE cocaja " _
                  & "SET cacercaj = 'S', " _
                  & "    caconsig = " & cero & ", " _
                  & "    canumpla = " & Txtnumpla.text & " " _
                & "WHERE canomcaj = '" & Pc & "' "
     
     'wecc 10/02/2002 actualiza el detalle de banco a la caje cerrada
     'en la consignaciones efectuadas por la caja
     
     Dbs.Execute "UPDATE codetban " _
                  & "SET dbcacerr = 'S' " _
                & "WHERE dbcaja = '" & Pc & "' " _
                  & "AND dbturno = " & Txtturno.text & " "
     
     If Val(G_Turno) < 4 Then
        sFecha1 = Txtfecha.text
        newfec = Txtfecha.text
     Else
        sFecha1 = Txtfecha.text
        newfec = DateAdd("d", -1, Txtfecha.text)
     End If
     
     If G_Turno = G_maxtur Then
        Dbs.Execute " UPDATE cocaja " _
                   & "SET caciedia = 'S' "
     
     'Btnimprimir_Click
     
         Dbs.Execute "UPDATE coconduc " _
                      & "SET cociedia = 'S' " _
                    & "WHERE conumero BETWEEN " & Txttiqini.text & " " _
                    & " AND " & Txttiqfin.text & " " _
                    & " AND conomcaj = '" & Cmbcaja.text & "' " _
                    & " AND cociedia = 'N' "
         
         Dbs.Execute "UPDATE opmultas " _
                      & "SET muciedia = 'S' " _
                    & "WHERE munomcaj = '" & Cmbcaja.text & "' " _
                      & "AND muciedia = 'N' "
                    
         Dbs.Execute "UPDATE opabonos " _
                      & "SET abciedia = 'S' " _
                    & "WHERE abnomcaj = '" & Cmbcaja.text & "' " _
                      & "AND abciedia = 'N' "
         
         Dbs.Execute "UPDATE cocietur " _
                      & "SET ciciecaj = 'S' " _
                    & "WHERE cinomcaj = '" & Cmbcaja.text & "' " _
                     & " AND ciciecaj = 'N' "
                     
         Dbs.Execute "UPDATE ophojveh " _
                  & "SET  hvnumsal = " & Val(cero) & ", " _
                  & "     hvestsel = 'No' "
     
     End If
     
     Dbs.Execute "UPDATE coconduc " _
                  & "SET cocerrado = 'S' " _
                & "WHERE conumero BETWEEN " & Txttiqini.text & " AND " & Txttiqfin.text & " " _
                & "  AND conomcaj = '" & Pc & "' " _
                 & " AND cocerrado = 'N' " _

     Dbs.Execute "UPDATE coparqueop " _
                  & "SET pacerrado = 'S', " _
                  & "    pafechadia = TO_DATE('" & sFecha1 & "','MM/DD/YYYY') " _
                & "WHERE pafecpag BETWEEN TO_DATE('" & newfec & "','MM/DD/YYYY') " _
                 & " AND TO_DATE('" & sFecha1 & "','MM/DD/YYYY') " _
                  & "AND pacaja = '" & Cmbcaja.text & "' " _
                  & "AND pacerrado = 'N' "
     
     QRY = "UPDATE cotransitonorte " _
            & "SET tncerradotasa = 'S' " _
          & "WHERE tnfechatasa BETWEEN TO_DATE('" & newfec & "','MM/DD/YYYY') " _
           & " AND TO_DATE('" & sFecha1 & "','MM/DD/YYYY') " _
           & " AND tncajatasa = '" & Cmbcaja.text & "' " _
           & " AND tnturnotasa = " & G_Turno & " " _
           & " AND tncerradotasa = '" & no & "' " _
           & " AND tnestado = 'C' "
           
     Dbs.Execute QRY
     
     Dbs.Execute "UPDATE opmultas " _
                  & "SET mucerrado = 'S' " _
                & "WHERE mucerrado = 'N' " _
                & "  AND munomcaj = '" & Pc & "' "
     
     Dbs.Execute "UPDATE opabonos " _
                  & "SET abcerrado = 'S' " _
                & "WHERE abcerrado = 'N' " _
                  & "AND abnomcaj = '" & Pc & "' "
     
     Dbs.Execute "UPDATE ophojveh " _
                  & "SET hvestsel = 'No' " _
                & "WHERE hvestsel = 'Si'"
                
    Dbs.Execute "UPDATE optag " _
                    & "SET tacerrado = 'S' " _
                    & "WHERE tacerrado = 'N' " _
                    & "AND tanomcaj = '" & Pc & "' "
                  
     Dbs.CommitTrans
     On Error GoTo err_Btngrabar
     main.Mnuapetur.Enabled = False
     main.Mnuvencon.Enabled = False
     main.Mnucancon.Enabled = False
     main.Mnuregpas.Enabled = False
     main.Mnucanmul.Enabled = False
     
     With main
        
        If .Mnuapetur.Enabled = False Then
           'cierre turno
           .Toolbar1.Buttons(6).Enabled = False
        End If
        
        If .Mnuvencon.Enabled = False Then
           'venta conduce
           .Toolbar1.Buttons(8).Enabled = False
        End If
        If .Mnucancon.Enabled = False Then
           'anulacion copnduce
           .Toolbar1.Buttons(9).Enabled = False
        End If
        
        If .Mnucanmul.Enabled = False Then
           'anulacion copnduce
           .Toolbar1.Buttons(12).Enabled = False
        End If
        
      End With
     MsgBox "Turno Cerrado", 64, "O.K."
     
     Btngrabar.Enabled = False
     Btnbuscar.Enabled = False
     Btnimprimir_Click
     'Btnimprimir.Enabled = False
  Else
     MsgBox "No Hay Nada Que Grabar", 48, "Advertencia"
     Exit Sub
  End If
  
  Exit Sub
err_Btngrabar:
   If Err.Number = 52 Then
      MsgBox "Se Ha Desconectado del Servidor", 48, "Desconectado"
      G_BaseServidor = False
      Resume Next
   End If
   MsgBox Err.Description, 16, "Error " & Err.Number
   
   Exit Sub
err_Btngrabar_Dbx:
   G_BaseServidor = False
   Resume Next
   
Err_Transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnimprimir_Click()
    On Error GoTo Errores
    'If Msktotgen.Text <> "" And Btngrabar.Enabled = False Then
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
     
      fec = Txtfecha.text
      'dbs.Execute "INSERT INTO optempo1 " _
                            & "(temdate1, temnumero1, temtext1, temtext2, temtext3, temtext4,  temnumero2, " _
                            "temnumero3, temnumero4, temnumdob1, temnumdob2, temnumero5, temnumdob3, temnumdob4 , " _
                         & "temnummon1, temnumdob5, temnummon2, temnummon3, temnumero6, temtext5, " _
                         temtext6 ) " _


      QRY = "SELECT cifecha, citurno, cinumini, cinumfin, cicodope, cihoraap, cianula, " _
                 & "ciorigen, citransi, citorige, cittrans, ciinfrac, citinfra, citotgen, " _
                 & "cibasmon, citanula, citvalalcori, citvalalctra, citparqu, cinomcaj, cinumpla, " _
                 & "cicredit, citcredi, citotrinf, citelusion, citevasion, citevastra, citvigencia, " _
                 & "cittag, cicupos, cicupvaltasa, cicupvalalco, citransitonorte, cittransitonorte, " _
                 & "cicupostransito, cicupostransitoaut, cicupvaltasatransito, ciinfluencia, citinfluencia, citvalalcinf, " _
                 & "cicuposinfluencia, cicupvaltasainfluencia, cicupalcinfluencia " _
            & "FROM cocietur " _
           & "WHERE cinumpla = " & Txtnumpla.text & " " _
             & "AND cinomcaj = '" & Pc & "' "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
                           
      If Not Tbl1.EOF Then
         netpago = Tbl1!citorige + Tbl1!cittrans + Tbl1!citinfra + Tbl1!citvalalcori + Tbl1!citvalalctra + _
                   Tbl1!citparqu - Tbl1!cibasmon + Tbl1!citotrinf + Tbl1!citelusion + Tbl1!citevasion + _
                   Tbl1!citevastra + Tbl1!citvigencia + Tbl1!cittag + Tbl1!cittransitonorte + Tbl1!citinfluencia + Tbl1!citvalalcinf
         netpago = Format(netpago, "##,##")
         reporte = RpPath + "\cuaturno.rpt"
         
         rptterminal.Connect = G_Conexion_reportes
         rptterminal.ReportFileName = reporte
         rptterminal.Formulas(0) = "observaciones = '" & Txtobservacion.text & "'"
         rptterminal.Formulas(1) = "total = '" & Msktotgen.FormattedText & "'"
         rptterminal.Formulas(2) = "totanucon = '" & totanucon & "'"
         rptterminal.Formulas(3) = "totanumul = '" & totanumul & "'"
         rptterminal.Formulas(4) = "empresa = '" & G_NomEmpresa & "'"
         rptterminal.Formulas(5) = "fecha = '" & str(Tbl1!cifecha) & "'"
         rptterminal.Formulas(6) = "turno = '" & str(Tbl1!citurno) & "'"
         rptterminal.Formulas(7) = "numini = '" & Tbl1!cinumini & "'"
         rptterminal.Formulas(8) = "numfin = '" & Tbl1!cinumfin & "'"
         rptterminal.Formulas(9) = "operador = '" & Tbl1!cicodope & "'"
         rptterminal.Formulas(10) = "horaap = '" & Format(Tbl1!cihoraap, "HH:mm:ss") & "'"
         rptterminal.Formulas(11) = "anula = '" & str(Tbl1!cianula) & "'"
         rptterminal.Formulas(12) = "origen = '" & str(Tbl1!ciorigen) & "'"
         rptterminal.Formulas(13) = "transi = '" & str(Tbl1!citransi) & "'"
         rptterminal.Formulas(14) = "torige = '" & Format(Tbl1!citorige, "##,##0") & "'"
         rptterminal.Formulas(15) = "ttrans = '" & Format(Tbl1!cittrans, "##,##0") & "'"
         rptterminal.Formulas(16) = "infra = '" & str(Tbl1!ciinfrac) & "'"
         rptterminal.Formulas(17) = "tinfra = '" & Format(Tbl1!citinfra, "##,##0") & "'"
         rptterminal.Formulas(18) = "basmon = '" & Format(Tbl1!cibasmon, "##,##0") & "'"
         rptterminal.Formulas(19) = "tanula = '" & str(Tbl1!citanula) & "'"
         rptterminal.Formulas(20) = "tvalalcori = '" & Format(Tbl1!citvalalcori, "##,##0") & "'"
         rptterminal.Formulas(21) = "tvalalctra = '" & Format(Tbl1!citvalalctra, "##,##0") & "'"
         rptterminal.Formulas(22) = "tparqu = '" & Format(Tbl1!citparqu, "##,##0") & "'"
         rptterminal.Formulas(23) = "nomcaj = '" & Tbl1!cinomcaj & "'"
         rptterminal.Formulas(24) = "numpla = '" & str(Tbl1!cinumpla) & "'"
         rptterminal.Formulas(25) = "netpag = '" & Format(netpago, "##,##0") & "'"
         rptterminal.Formulas(26) = "credito = '" & str(Tbl1!cicredit) & "'"
         rptterminal.Formulas(27) = "tcredito = '" & Format(Tbl1!citcredi, "##,##0") & "'"
         rptterminal.Formulas(28) = "totrmul = '" & Format(Tbl1!citotrinf, "##,##0") & "'"
         
         rptterminal.Formulas(29) = "totmulelu = '" & Format(Tbl1!citelusion, "##,##0") & "'"
         rptterminal.Formulas(30) = "totmuleva = '" & Format(Tbl1!citevasion, "##,##0") & "'"
         rptterminal.Formulas(31) = "totmuletr = '" & Format(Tbl1!citevastra, "##,##0") & "'"
         rptterminal.Formulas(32) = "totvigencia = '" & Format(Tbl1!citvigencia, "##,##0") & "'"
         rptterminal.Formulas(33) = "totexcreta = '" & Format(Tbl1!cittag, "##,##0") & "'"
         
         rptterminal.Formulas(34) = "totcupos = '" & Format(Tbl1!cicupos, "##,##0") & "'"
         rptterminal.Formulas(35) = "tottasacupo = '" & Format(Tbl1!cicupvaltasa, "##,##0") & "'"
         rptterminal.Formulas(36) = "totalccupo = '" & Format(Tbl1!cicupvalalco, "##,##0") & "'"
         
         rptterminal.Formulas(37) = "transitonorte = '" & Format(Tbl1!citransitonorte, "##,##0") & "'"
         rptterminal.Formulas(38) = "totaltransitonorte = '" & Format(Tbl1!cittransitonorte, "##,##0") & "'"
         rptterminal.Formulas(39) = "totcupostransito = '" & Format(Tbl1!cicupostransito, "##,##0") & "'"
         rptterminal.Formulas(40) = "totcupostraaut = '" & Format(Tbl1!cicupostransitoaut, "##,##0") & "'"
         
         rptterminal.Formulas(41) = "tottasacupotransito = '" & Format(Tbl1!cicupvaltasatransito, "##,##0") & "'"
         
         rptterminal.Formulas(42) = "influencia = '" & str(Tbl1!ciinfluencia) & "'"
         rptterminal.Formulas(43) = "tinfluencia = '" & Format(Tbl1!citinfluencia, "##,##0") & "'"
         rptterminal.Formulas(44) = "tvalalcinf = '" & Format(Tbl1!citvalalcinf, "##,##0") & "'"
               
         rptterminal.Formulas(45) = "totcuposinfluencia = '" & Format(Tbl1!cicuposinfluencia, "##,##0") & "'"
         rptterminal.Formulas(46) = "tottasacupoinfluencia = '" & Format(Tbl1!cicupvaltasainfluencia, "##,##0") & "'"
         rptterminal.Formulas(47) = "totalccupoinfluencia = '" & Format(Tbl1!cicupalcinfluencia, "##,##0") & "'"
         
         'insertamos en la tabla temporal para luego
         'hacer el select de la caja
         For i = 1 To Grdconsigna.rows - 1
             Grdconsigna.Row = i
                         
             Grdconsigna.Col = 1
             hora = Trim(Grdconsigna.text)
             Grdconsigna.Col = 5
             banco = Trim(Grdconsigna.text)
             Grdconsigna.Col = 6
             valor = Trim(Grdconsigna.text)
             Dbs.Execute "INSERT INTO optempo1 " _
                            & "(temtext1, temtext2, temtext3, temnumdob1, tempc) " _
                            & "SELECT '" & Format(hora, "HH:mm:ss") & "','" & banco & "',banombre, " & valor & ",'" & Pc & "' " _
                              & "FROM cobancos " _
                            & "WHERE bacodigo = '" & banco & "' "
                            
         Next i
            
         QRY = "SELECT temtext1, temtext2,  temtext3, temnumdob1, tempc " _
               & "FROM optempo1  " _
              & "WHERE tempc = '" & Pc & "' "
        
         Cuaturno.Enabled = False
         
         If G_ImpCon = "Z" Or G_ImpCon = "C" Or G_ImpCon = "N" Then
            L_Buscar = "1"
            imprimir.Show 1
            If G_Imprimir = "1" Then
               rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
               rptterminal.PrinterName = G_PrinterName
               rptterminal.PrinterPort = G_PrinterPort
               rptterminal.PrinterDriver = G_PrinterDriver
               
               rptterminal.Destination = 1
               'rptterminal.CopiesToPrinter = 2
               rptterminal.WindowState = crptMaximized
               rptterminal.Action = 1
               Cuaturno.Enabled = True
               Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
            End If
         ElseIf G_ImpCon = "E" Then
            rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
            rptterminal.Destination = 1
            'rptterminal.CopiesToPrinter = 2
            rptterminal.WindowState = crptMaximized
            rptterminal.Action = 1
            Cuaturno.Enabled = True
            Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
         End If
         Cuaturno.Enabled = True
         Tbl1.Close
      Else
         Tbl1.Close
         '*************WECC 02/12/2005*************
         'Se imprime un reporte sin grabar para auditoria
      
         netpago = CDbl(Msktotori.text) + CDbl(Msktottra.text) + CDbl(Msktotinf.text) + CDbl(Msktotmul.text) + _
                   CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskalcinf.text) + CDbl(Mskparqueadero.text) - _
                   CDbl(Mskbasmon.text) + CDbl(MskMulPer.text) + CDbl(MskMulElu.text) + _
                   CDbl(MskMulEva.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTotTranNorte.text)   '+ CDbl(MskExcreta.Text)
                   
         netpago = Format(netpago, "##,##")
         reporte = RpPath + "\cuaturno.rpt"
         
         
         rptterminal.Connect = G_Conexion_reportes
         rptterminal.ReportFileName = reporte
         rptterminal.Formulas(0) = "observaciones = '" & Txtobservacion.text & "'"
         rptterminal.Formulas(1) = "total = '" & Msktotgen.FormattedText & "'"
         rptterminal.Formulas(2) = "totanucon = '" & totanucon & "'"
         rptterminal.Formulas(3) = "totanumul = '" & totanumul & "'"
         rptterminal.Formulas(4) = "empresa = '" & G_NomEmpresa & "'"
         rptterminal.Formulas(5) = "fecha = '" & Txtfecha.text & "'"
         rptterminal.Formulas(6) = "turno = '" & Txtturno.text & "'"
         rptterminal.Formulas(7) = "numini = '" & Txttiqini.text & "'"
         rptterminal.Formulas(8) = "numfin = '" & Txttiqfin.text & "'"
         rptterminal.Formulas(9) = "operador = '" & Login & "'"
         rptterminal.Formulas(10) = "horaap = '" & Format(Txthora.text, "HH:mm:ss") & "'"
         rptterminal.Formulas(11) = "anula = '" & Txttiqanu.text & "'"
         rptterminal.Formulas(12) = "origen = '" & Txtnumori.text & "'"
         rptterminal.Formulas(13) = "transi = '" & Txtnumtra.text & "'"
         rptterminal.Formulas(14) = "torige = '" & Format(Msktotori.text, "##,##0") & "'"
         rptterminal.Formulas(15) = "ttrans = '" & Format(Msktottra.text, "##,##0") & "'"
         rptterminal.Formulas(16) = "infra = '" & str(Txtnummul.text) & "'"
         rptterminal.Formulas(17) = "tinfra = '" & Format(Msktotmul.text, "##,##0") & "'"
         rptterminal.Formulas(18) = "basmon = '" & Format(Mskbasmon.text, "##,##0") & "'"
         rptterminal.Formulas(19) = "tanula = '" & str(Txtnumanu.text) & "'"
         rptterminal.Formulas(20) = "tvalalcori = '" & Format(Mskalcori.text, "##,##0") & "'"
         rptterminal.Formulas(21) = "tvalalctra = '" & Format(Mskalctra.text, "##,##0") & "'"
         rptterminal.Formulas(22) = "tparqu = '" & Format(Mskparqueadero.text, "##,##0") & "'"
         rptterminal.Formulas(23) = "nomcaj = '" & Cmbcaja.text & "'"
         rptterminal.Formulas(24) = "numpla = '" & str(Txtnumpla.text) & "'"
         rptterminal.Formulas(25) = "netpag = '" & Format(netpago, "##,##0") & "'"
         rptterminal.Formulas(26) = "credito = '" & str(txttiqcre.text) & "'"
         rptterminal.Formulas(27) = "tcredito = '" & Format(mskcredito.text, "##,##0") & "'"
         rptterminal.Formulas(28) = "totrmul = '" & Format(MskMulPer.text, "##,##0") & "'"
         
         rptterminal.Formulas(29) = "totmulelu = '" & Format(MskMulElu.text, "##,##0") & "'"
         rptterminal.Formulas(30) = "totmuleva = '" & Format(MskMulEva.text, "##,##0") & "'"
         rptterminal.Formulas(31) = "totmuletr = '" & Format(MskMulEvaTra.text, "##,##0") & "'"
         rptterminal.Formulas(32) = "totvigencia = '" & Format(MskVigencia.text, "##,##0") & "'"
         rptterminal.Formulas(33) = "totexcreta = '" & Format(MskTag.text, "##,##0") & "'"
         
         rptterminal.Formulas(34) = "totcupos = '" & Format(TxtCuposOrigen.text, "##,##0") & "'"
         rptterminal.Formulas(35) = "tottasacupo = '" & Format(MskVrCupTasaOrigen.text, "##,##0") & "'"
         rptterminal.Formulas(36) = "totalccupo = '" & Format(MskVrCupAlc.text, "##,##0") & "'"
         
         rptterminal.Formulas(37) = "transitonorte = '" & Format(TxtNumTransitoNorte.text, "##,##0") & "'"
         rptterminal.Formulas(38) = "totaltransitonorte = '" & Format(MskTotTranNorte.text, "##,##0") & "'"
                  
         rptterminal.Formulas(39) = "totcupostransito = '" & Format(TxtCuposTransito.text, "##,##0") & "'"
         rptterminal.Formulas(40) = "totcupostraaut = '" & Format(TxtCupTraAut.text, "##,##0") & "'"
                  
         rptterminal.Formulas(41) = "tottasacupotransito = '" & Format(MskVrCupTasaTransito.text, "##,##0") & "'"
         
         rptterminal.Formulas(42) = "influencia = '" & Txtnuminf.text & "'"
         rptterminal.Formulas(43) = "tinfluencia = '" & Format(Msktotinf.text, "##,##0") & "'"
         rptterminal.Formulas(44) = "tvalalcinf = '" & Format(Mskalcinf.text, "##,##0") & "'"
      
         rptterminal.Formulas(45) = "totcuposinfluencia = '" & Format(TxtCuposInfluencia.text, "##,##0") & "'"
         rptterminal.Formulas(46) = "tottasacupoinfluencia = '" & Format(MskVrCupTasaInfluencia.text, "##,##0") & "'"
         rptterminal.Formulas(47) = "totalccupoinfluencia = '" & Format(MskVrCupAlcInfluencia.text, "##,##0") & "'"
      
         'insertamos en la tabla temporal para luego
         'hacer el select de la caja
         For i = 1 To Grdconsigna.rows - 1
             Grdconsigna.Row = i
                         
             Grdconsigna.Col = 1
             hora = Trim(Grdconsigna.text)
             Grdconsigna.Col = 5
             banco = Trim(Grdconsigna.text)
             Grdconsigna.Col = 6
             valor = Trim(Grdconsigna.text)
             Dbs.Execute "INSERT INTO optempo1 " _
                            & "(temtext1, temtext2, temtext3, temnumdob1, tempc) " _
                            & "SELECT '" & Format(hora, "hh:mm:ss") & "','" & banco & "',banombre, " & valor & ",'" & Pc & "' " _
                              & "FROM cobancos " _
                            & "WHERE bacodigo = '" & banco & "' "
                            
         Next i
            
         'qry = "SELECT temtext1, temtext2,  temtext3, temnumdob1, tempc " _
               & "FROM optempo1  " _
              & "WHERE tempc = '" & pc & "' "
        
         Cuaturno.Enabled = True
         
         If G_ImpCon = "Z" Or G_ImpCon = "C" Or G_ImpCon = "N" Then
            L_Buscar = "1"
            imprimir.Show 1
            If G_Imprimir = "1" Then
               rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
               rptterminal.PrinterName = G_PrinterName
               rptterminal.PrinterPort = G_PrinterPort
               rptterminal.PrinterDriver = G_PrinterDriver
               If Cmbcaja.Enabled = False Then
                  rptterminal.Destination = 1
               Else
                  rptterminal.Destination = 0
               End If
               
               'rptterminal.CopiesToPrinter = 2
               rptterminal.WindowState = crptMaximized
               rptterminal.Action = 1
               Cuaturno.Enabled = True
               Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
            End If
         ElseIf G_ImpCon = "E" Then
            rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
            If Cmbcaja.Enabled = False Then
               rptterminal.Destination = 1
            Else
               rptterminal.Destination = 0
            End If
            'rptterminal.CopiesToPrinter = 2
            rptterminal.WindowState = crptMaximized
            rptterminal.Action = 1
            Cuaturno.Enabled = True
            Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
         End If
         Cuaturno.Enabled = True
      
      End If
      
  'End If
  
  Exit Sub
Errores:
  Cuaturno.Enabled = True
  MsgBox Err.Description, 16, "Error " & Err.Number
  Exit Sub
  Resume
End Sub

Private Sub Btnlimpiar_Click()
   
   fqry01 = "SELECT * " _
            & "FROM cocaja " _
           & "WHERE canomcaj = '" & Pc & "' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Not Ftb01.EOF Then
      Cmbcaja.text = Pc
      Cmbcaja.Enabled = False
   Else
      Cmbcaja.ListIndex = -1
      Cmbcaja.Enabled = True
   End If
    
   Txttiqini.text = ""
   Txttiqfin.text = ""
   Txttottiq.text = ""
   Txttiqanu.text = ""
   Txttotal.text = ""
   Txtnumori.text = ""
   Txtnumtra.text = ""
   TxtNumTransitoNorte.text = ""
   Msktotori.text = ""
   Msktottra.text = ""
   MskTotTranNorte.text = ""
   Msktotcon.text = ""
   Txtnummul.text = ""
   Txtnumanu.text = ""
   Txtnumneg.text = ""
   Msktotmul.text = ""
   Mskbasmon.text = ""
   Msktotgen.text = ""
   Txtobservacion.text = ""
   Mskcaja.text = ""
   Txtalcori.text = ""
   Mskalcori.text = ""
   Txtalctra.text = ""
   Mskalctra.text = ""
   Mskparqueadero.text = ""
   MskVigencia.text = ""
   mskcredito.text = ""
   txttiqcre.text = ""
   txtbancos.text = ""
   lblbancos.Caption = ""
   Mskvalor.text = ""
   MskMulElu.text = ""
   MskMulEva.text = ""
   MskMulPer.text = ""
   MskMulEvaTra.text = ""
   MskTag.text = ""
   TxtCuposOrigen.text = ""
   TxtCuposTransito.text = ""
   TxtCupTraAut.text = ""
   MskVrCupAlc.text = ""
   MskVrCupTasaOrigen.text = ""
   MskVrCupTasaTransito.text = ""
   Btngrabar.Enabled = True
   btnconsig.Enabled = True
   btnañadir.Enabled = True
   mskconsi.text = ""
   msksalban.text = ""
   Txtalcinf.text = ""
   Mskalcinf.text = ""
   
   TxtCuposInfluencia.text = ""
   MskVrCupTasaInfluencia.text = ""
   MskVrCupAlcInfluencia.text = ""
   
   Txtnuminf.text = ""
   Msktotinf.text = ""
   DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Cmbcaja_LostFocus()
   fqry01 = "SELECT * " _
            & "FROM cocaja " _
           & "WHERE canomcaj = '" & Cmbcaja.text & "' "
             
   Set Ftb01 = Dbs.Execute(fqry01)
   If Not Ftb01.EOF Then
      Txtnumpla.text = Ftb01!canumpla + 1
      Txtturno.text = Ftb01!caturcaj
   End If
   Ftb01.Close
End Sub

Private Sub Form_Activate()
   Call Aplicar_Perfil(Me, "Cuaturno")
  
   If L_Buscar = "0" Then
      Btnbuscar_Click
      
      If G_CierreAutomatico = 1 Then
        Mskbasmon.text = "0"
        Txtobservacion.text = "Cierre automatico. No se realizo cierre por parte del operario."
        
        'Realizamos la consignación del banco
        'Comenta Nelson que siempre se realiza sobre el banco 006
        txtbancos.text = "006"
        'El valor siempre es el total generado
        Mskvalor.text = Msktotgen.text
        'Añadimos sobre la grilla
        btnañadir_Click
        'Guardamos los registros
        btnconsig_Click
        
        'Ahora guardamos el cierre
        BtnGrabar_Click
        'Una vez realizamos el cierre, cerramos la pestaña
        Btnsalir_Click
      End If
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo Errores
   'Set dbs = OpenDatabase(dbpath)
   L_Buscar = "0"
   Txtturno.text = G_Turno
   If Login = "" Then
    Login = "automatico"
   End If
   Txtoperario.text = Login
   
   fqry01 = "SELECT canomcaj " _
            & "FROM cocaja "
            
   Set Ftb01 = Dbs.Execute(fqry01)
   If Not Ftb01.EOF Then
      i = 0
      Do While Not Ftb01.EOF
         Cmbcaja.List(i) = Ftb01(0)
         i = i + 1
         Ftb01.MoveNext
      Loop
      Ftb01.Close
      
      fqry01 = "SELECT * " _
               & "FROM cocaja " _
              & "WHERE canomcaj = '" & Pc & "' "
             
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         Txtnumpla.text = Ftb01!canumpla + 1
         G_maxtur = Ftb01!canumtur
         Cmbcaja.text = Pc
         Cmbcaja.Enabled = False
         
         QRY = "SELECT * " _
               & "FROM cocietur " _
              & "WHERE cinomcaj = '" & Pc & "' " _
                & "AND cinumpla = " & Txtnumpla.text & ""
         
         Set Ftb02 = Dbs.Execute(QRY)
         
         If Not Ftb02.EOF Then
            Ftb02.Close
            QRY = "SELECT MAX(cinumpla ) as numplanilla " _
                  & "FROM cocietur " _
                 & "WHERE cinomcaj = '" & Pc & "'"
                 
            Set Ftb02 = Dbs.Execute(QRY)
            If Not Ftb02.EOF Then
               Txtnumpla.text = Ftb02(0) + 1
            End If
            Ftb02.Close
         Else
            Ftb02.Close
         End If
         
      Else
         
         Btngrabar.Enabled = False
         Cmbcaja.Enabled = True
         btnconsig.Enabled = False
      End If
      Ftb01.Close
   End If
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Not Ftb01.EOF Then
      hora = Mid(Ftb01!psval, 12, 8)
      Fecha = Mid(Ftb01!psval, 1, 10)
      Txthora.text = hora
      Txtfecha.text = Fecha
   End If
   Ftb01.Close
   
   'wecc 04/03/2013 caso cau 6849 _
   conceptos de multas para cierre de turno
   'COBRO POR ELUSION DE TUT: 51
   'COBRO POR EVACION DE TUT: 52
   'COBRO POR PERDIDA DE DOCUMENTOS: 53
   'COBRO POR EVACION DE TUT EN TRANSITO: 54
   
   
   
   L_CiTElusion = Fn_Parametros("CITELUSION", 1)
   L_CiTEvasion = Fn_Parametros("CITEVASION", 1)
   L_CiTPerdida = Fn_Parametros("CITPERDIDA", 1)
   L_CiTEvasTra = Fn_Parametros("CITEVASTRA", 1)
   'L_CiTExcreta = Fn_Parametros("CITEXCRETA", 1)
   
   
   L_Todas_Multas = Trim(L_CiTElusion & ", " & L_CiTEvasion & ", " & L_CiTPerdida & ", " & L_CiTEvasTra)
   'L_Todas_Multas = Trim(L_CiTElusion & ", " & L_CiTEvasion & ", " & L_CiTPerdida)
   
   If Mid(L_Todas_Multas, Len(L_Todas_Multas), 1) = "," Then
      L_Todas_Multas = Mid(Trim(L_Todas_Multas), 1, Len(L_Todas_Multas) - 1)
   End If
   
   Traer_Consig
   
   '*******
   'wecc 03/25/2003
   'fqry01 = "SELECT * " _
            & "FROM coconpen " _
          & "WHERE pecaja = '" & pc & "' "
          
   'Set ftb01 = dbs.OpenRecordset(fqry01, dbOpenSnapshot)
  
   'If Not ftb01.EOF Then
   '   On Error Resume Next
   '   MsgBox "Existen Conduces Pendientes", 48, "Advertencia"
   '   Btngrabar.Enabled = False
   '   Btnbuscar.Enabled = False
   '   Btnlimpiar.Enabled = False
   '   Btnimprimir.Enabled = False
   '   btnconsig.Enabled = False
   '   ftb01.Close
   '   Exit Sub
   'End If
   
   '*******
   
   Exit Sub
   
Errores:
   
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Grdconsigna_DblClick()
If Grdconsigna.rows > 1 Then
     y = Grdconsigna.rows - 1
     If y = 1 And Grdconsigna.Row = 1 Then
        Grdconsigna.Col = 5
        txtbancos.text = Trim(Grdconsigna.text)
        txtbancos_LostFocus
        Grdconsigna.Col = 6
        Mskvalor.text = Trim(Grdconsigna.text)
        mskconsi.text = CDbl(mskconsi.text) - CDbl(Grdconsigna.text)
        msksalban.text = CDbl(msksalban.text) + CDbl(Grdconsigna.text)
        DIB_GRILLA
        
     ElseIf y > 1 Then
        Z = Grdconsigna.Row      ' COLOCA EL NUMERO DE "FILA" EN Z PARA PODER BORRARLO
        
        Grdconsigna.Col = 5
        txtbancos.text = Trim(Grdconsigna.text)
        txtbancos_LostFocus
        Grdconsigna.Col = 6
        Mskvalor.text = Trim(Grdconsigna.text)
        mskconsi.text = CDbl(mskconsi.text) - CDbl(Grdconsigna.text)
        msksalban.text = CDbl(msksalban.text) + CDbl(Grdconsigna.text)
        Grdconsigna.RemoveItem Z ' POR QUE Y TIENE EL N° DE "FILAS"
                
     End If
  End If
End Sub


Private Sub Mskbasmon_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48) Or (KeyAscii% > 57) And (KeyAscii% <> 8) Then
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskbasmon_LostFocus()
  If Mskbasmon.text <> "" And Msktotcon.text <> "" And Msktotmul.text <> "" Then
     Msktotgen.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(mskcredito.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTag.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text)
     Mskcaja.text = CDbl(Msktotcon.text) + CDbl(Msktotmul.text) - CDbl(Mskbasmon.text) + CDbl(Mskalcori.text) + CDbl(Mskalctra.text) + CDbl(Mskparqueadero.text) + CDbl(MskMulElu.text) + CDbl(MskMulEva.text) + CDbl(MskMulPer.text) + CDbl(MskMulEvaTra.text) + CDbl(MskVigencia.text) + CDbl(MskTag.text) + CDbl(MskTotTranNorte.text) + CDbl(Mskalcinf.text)
  End If
End Sub


Private Sub MskValor_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Then
      KeyAscii% = 0
  End If

End Sub

Private Sub txtbancos_DblClick()
    forma = "13"
    forma1 = "1"
    txtbancos.text = ""
    buscar.Show 1
    txtbancos_LostFocus
End Sub

Private Sub txtbancos_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
   
  If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
     KeyAscii% = 0
  End If
End Sub

Private Sub txtbancos_LostFocus()
  On Error GoTo err_txtbancos_LostFocus
  
    If Trim(txtbancos.text) <> "" Then
        QRY = "SELECT * " _
                & "FROM cobancos " _
               & "WHERE bacodigo ='" & Trim(txtbancos.text) & "' "
               
          Set Tbl1 = Dbs.Execute(QRY)
         
           If Not Tbl1.EOF Then
               If Tbl1!baestado = "I" Then
                  MsgBox "El Banco Esta Inactivo", 48, "Advertencia"
                  txtbancos.SetFocus
                  Tbl1.Close
                  Exit Sub
               End If
               
               lblbancos.Caption = Tbl1!banombre
           Else
               MsgBox "El Banco No Existe", 48, "Advertencia"
               txtbancos.SetFocus
           End If
           Tbl1.Close
    End If
   Exit Sub
err_txtbancos_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

Private Sub Txtobservacion_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
  End If
End Sub

Private Function DIB_GRILLA()
     Grdconsigna.rows = 1
     Grdconsigna.Cols = 8
     Grdconsigna.Clear
     
     Grdconsigna.RowHeight(0) = 200
     Grdconsigna.Font = 4
     
     
     s$ = "Fecha  |Hora  |Caja  |Turno  |Funcionario| Banco  |Valor |Caja Cerrada "
     Grdconsigna.FormatString = s$
     Grdconsigna.Row = 0
     Grdconsigna.Col = 0
     Grdconsigna.ColWidth(0) = 1000
     Grdconsigna.ColWidth(1) = 1000
     Grdconsigna.ColWidth(2) = 1000
     Grdconsigna.ColWidth(3) = 1000
     Grdconsigna.ColWidth(4) = 1200
     Grdconsigna.ColWidth(5) = 1000
     Grdconsigna.ColWidth(6) = 1500
     Grdconsigna.ColWidth(7) = 1500
         
     For i = 0 To Grdconsigna.Cols - 1
     
        Grdconsigna.Col = i
        Grdconsigna.CellFontBold = 1
     Next i
End Function

Private Function Traer_Consig()
    On Error GoTo err_Traer_Consig
    DIB_GRILLA
    
    'wecc 11/14/2002 se parcha porque se cambio las horas de los turnos
    'If G_maxtur < 3 Then
    If G_Turno < 4 Then
    QRY = "SELECT dbfecha, dbhora,dbcaja, dbturno, dboperar, dbcodban, dbvalor, dbcacerr " _
          & "FROM codetban " _
         & "WHERE dbfecha = TO_DATE('" & Txtfecha.text & "','MM/DD/YYYY HH24:MI:SS') " _
           & "AND dbturno = " & Txtturno.text & " " _
           & "AND dbcaja = '" & Cmbcaja.text & "' " _
           & "AND dbcacerr = 'N' "
    
    
    ElseIf G_Turno = 4 Then
       fec3 = Txtfecha.text
       newfec1 = DateAdd("d", -1, Txtfecha.text)
       fec2 = newfec1
       
       QRY = "SELECT dbfecha, dbhora,dbcaja, dbturno, dboperar, dbcodban, dbvalor, dbcacerr " _
          & "FROM codetban " _
         & "WHERE dbfecha >= to_date('" & fec2 & "','MM/DD/YYYY HH24:MI:SS') " _
           & "AND dbfecha <= to_date('" & fec3 & "','MM/DD/YYYY HH24:MI:SS') " _
           & "AND dbturno = " & Txtturno.text & " " _
           & "AND dbcaja = '" & Cmbcaja.text & "' " _
           & "AND dbcacerr = 'N' "
    End If
    Set Tbl1 = Dbs.Execute(QRY)
         
    valcons = 0   'valor consignado
    Do While Not Tbl1.EOF
       
        entry = str(Tbl1(0)) + Chr(9) + Mid(Tbl1(1), 12) + Chr(9) + _
            Tbl1(2) + Chr(9) + str(Tbl1(3)) + Chr(9) + _
            Tbl1(4) + Chr(9) + Tbl1(5) + Chr(9) + _
            str(Tbl1(6)) + Chr(9) + Tbl1(7)
            
        Grdconsigna.AddItem entry
        valcons = CDbl(valcons) + CDbl(Tbl1(6))
       Tbl1.MoveNext
    Loop
    mskconsi.text = valcons
    Tbl1.Close
    Exit Function
err_Traer_Consig:

     MsgBox Err.Description, 16, "Error " & Err.Number

     Exit Function
End Function
