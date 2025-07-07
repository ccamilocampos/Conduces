VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Anumultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación /   Educativo de Evasiones y Elusiones"
   ClientHeight    =   8445
   ClientLeft      =   1455
   ClientTop       =   1470
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3615
      Left            =   6240
      TabIndex        =   42
      Top             =   4680
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
         TabIndex        =   44
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Txtrecpag 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
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
         TabIndex        =   43
         ToolTipText     =   "Recibo de pago"
         Top             =   720
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Mskvalmul 
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
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
         TabIndex        =   46
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         TabIndex        =   50
         Top             =   2760
         Width           =   1575
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
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
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
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
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
         TabIndex        =   47
         Top             =   1080
         Width           =   1575
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
      Height          =   615
      Left            =   120
      TabIndex        =   39
      Top             =   5760
      Width           =   6015
      Begin VB.OptionButton Optestado 
         Caption         =   "Educativo"
         BeginProperty Font 
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
         Left            =   4440
         TabIndex        =   78
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Optestado 
         Caption         =   "Anulado"
         BeginProperty Font 
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
         Left            =   3120
         TabIndex        =   41
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   1680
         TabIndex        =   40
         Top             =   240
         Width           =   1335
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
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Adicionales"
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
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   6015
      Begin VB.TextBox Txtobsanu 
         Height          =   1215
         Left            =   3000
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Txtobservaciones 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   2775
      End
      Begin Crystal.CrystalReport rptterminal 
         Left            =   5520
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin VB.Label Label19 
         Caption         =   "Razón de Anulación:"
         BeginProperty Font 
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
         TabIndex        =   53
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label12 
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
         Left            =   120
         TabIndex        =   38
         Top             =   240
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
      Height          =   4455
      Left            =   6240
      TabIndex        =   33
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
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Número de la multa(doble Click)"
         Top             =   720
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   2400
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
         Left            =   240
         TabIndex        =   14
         Top             =   3240
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
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         Format          =   "mmm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCaja 
         Height          =   285
         Left            =   240
         TabIndex        =   82
         Top             =   3960
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
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   81
         Top             =   3720
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
         Left            =   240
         TabIndex        =   52
         Top             =   1200
         Width           =   1455
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
         Left            =   240
         TabIndex        =   36
         Top             =   3000
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
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   975
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
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
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
      Height          =   5535
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6015
      Begin VB.TextBox TxtCedula 
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
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   79
         ToolTipText     =   "Número cédula del conductor"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Txtauxalt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   76
         ToolTipText     =   "Código del auxiliar de control que hizo el parte(doble Click)"
         Top             =   5040
         Width           =   855
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
         TabIndex        =   65
         Top             =   1800
         Width           =   2775
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
            TabIndex        =   67
            ToolTipText     =   "Código de la ruta (presione DblClick para ayuda)"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txtnumcon 
            Height          =   285
            Left            =   840
            TabIndex        =   66
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1320
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Mskvalor 
            Height          =   285
            Left            =   840
            TabIndex        =   68
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
            TabIndex        =   69
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
            TabIndex        =   74
            Top             =   600
            Width           =   2535
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
            TabIndex        =   73
            Top             =   240
            Width           =   735
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
            TabIndex        =   72
            Top             =   960
            Width           =   1455
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
            TabIndex        =   70
            Top             =   1680
            Width           =   735
         End
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
         Height          =   2055
         Left            =   3120
         TabIndex        =   55
         Top             =   1800
         Width           =   2775
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
            TabIndex        =   58
            ToolTipText     =   "Código de la ruta (presione DblClick para Ayuda)"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Txtplanilla 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   57
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Txttiquete 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   56
            ToolTipText     =   "Número inicial por el que empiece a buscar"
            Top             =   1680
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Mskvalpla 
            Height          =   285
            Left            =   1320
            TabIndex        =   59
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
            TabIndex        =   64
            Top             =   600
            Width           =   2535
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
            TabIndex        =   63
            Top             =   240
            Width           =   735
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
            TabIndex        =   62
            Top             =   960
            Width           =   1455
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
            TabIndex        =   61
            Top             =   1320
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
            TabIndex        =   60
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.TextBox Txtnumint 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtnomnot 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   4080
         Width           =   2715
      End
      Begin VB.TextBox Txtnomtes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   4440
         Width           =   4395
      End
      Begin VB.TextBox Txtnomaux 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Txtsitio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txttipveh 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txtcodinf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
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
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   2
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
         Top             =   4080
         Width           =   1335
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
         TabIndex        =   77
         Top             =   5040
         Width           =   1815
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
         TabIndex        =   75
         Top             =   4800
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
         TabIndex        =   54
         Top             =   1080
         Width           =   1935
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
         TabIndex        =   32
         Top             =   5040
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
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   30
         Top             =   4440
         Width           =   1455
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
         TabIndex        =   29
         Top             =   4800
         Width           =   975
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
         TabIndex        =   28
         Top             =   840
         Width           =   2055
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
         TabIndex        =   27
         Top             =   840
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
         Left            =   2880
         TabIndex        =   26
         Top             =   1440
         Width           =   3015
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
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
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
         TabIndex        =   24
         Top             =   240
         Width           =   1215
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   480
         Width           =   2175
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
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
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
      Picture         =   "Anumulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ingresa la empresa a la base de datos"
      Top             =   4200
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
      Picture         =   "Anumulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7560
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
      Picture         =   "Anumulta.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Limpia la forma actual"
      Top             =   6720
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
      Picture         =   "Anumulta.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   5880
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
      Picture         =   "Anumulta.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   9480
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   8280
      X2              =   8280
      Y1              =   0
      Y2              =   8400
   End
End
Attribute VB_Name = "Anumultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim estado As String
Dim tasapago As Integer

Private Sub BtnBuscar_Click()
  forma = "7"
  forma1 = "3"
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
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
   
   If Trim(Txtobsanu.text) = "" Then
      MsgBox "Falta Ingresar las Observaciones de la Anulación", 48, "Advertencia"
      Txtobsanu.SetFocus
      Exit Sub
   End If
   
   If OptEstado(2).Value = True Then
      estado = "Anulado"
   ElseIf OptEstado(3).Value = True Then
      estado = "Educativo"
   End If
     
     
   'Camilo Campos 07/02/2025 desde esta fecha solo puede anularse una multa CON LA MISMA CAJA que se anulo
   If Not Trim(txtcaja.text) = Pc Then
     MsgBox "Para anular una multa debe ser con la misma caja que la generó", 48, "Advertencia"
     Exit Sub
   End If
   

   respuesta = MsgBox("Desea Grabar La Multa", 32 + vbYesNo, "GRABAR")
  
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
  
   Dbs.BeginTrans
   On Error GoTo transac
   QRY = "UPDATE opmultas " _
          & "SET muestado = '" & estado & "'," _
          & "    muturno  =  " & G_Turno & "," _
          & "    muoperar = '" & Login & "'," _
          & "    muhora   = to_date('" & hora & "','hh24:mi:ss')," _
          & "    mucerrado = '" & no & "'," _
          & "    muciedia  = '" & no & "'," _
          & "    mufecsis = to_date('" & Mskfecsis.text & "','MM/DD/YYYY'), " _
          & "    muobsanu = '" & Trim(Txtobsanu.text) & "' " _
        & "WHERE munumero = " & Txtnumero.text & ""
       
   
   Dbs.Execute QRY

    'Se arregla el pago de la multa
    'Buscamos el número de la multa primero
'    QRY = "SELECT munumero " _
'            & "FROM opmultas " _
'            & "WHERE munumpla = '" & Txtnumero.text & "' "
'
'    Set Tbl1 = Dbs.Execute(QRY)
'
'    If Not Tbl1.EOF Then
'        'Consultamos el próximo consecutivo de multas
'        QRY = "SELECT psval FROM geparsis WHERE pscod = 'NUMMUL' "
'        Set Ftb02 = Dbs.Execute(QRY)
'
'        If Not Ftb02.EOF Then
'            sNumMulta = Ftb02!psval
'        End If
'
'        'Teniendo el número de la multa, buscamos el total de abonos de la misma
'        QRY1 = "SELECT NVL(SUM(abvalor), 0) AS valmul FROM opabonos " _
'            & "WHERE abnumero = " & Tbl1!munumero & " " _
'            & "AND abestado = 'Activo' "
'
'        Set Tbl2 = Dbs.Execute(QRY1)
'
'        If Not Tbl2.EOF Then
'            'Si encontro abonos, pasamos estos a anulados, y la suma total
'            'la volvemos a dejar en el saldo de la multa
'
'            'Pasamos la multa como anulada
'            'Y creamos una nueva para que sea cobrada y guardar registro de la anulación
'            QRY = "UPDATE opmultas " _
'                    & "SET muaviso = muaviso - 1, " _
'                    & "musaldo = " & Tbl2!valmul & ", " _
'                    & "munomcaj = 'automatica', " _
'                    & "muestado = 'Anulado' " _
'                    & "WHERE munumero = " & Tbl1!munumero & " " _
'                    & "AND munumpla = '" & Txtplaca.text & "' "
'
'            Dbs.Execute (QRY)
'
'            'Aumentamos en 1 el número de la multa
'            sNumMulta = sNumMulta + 1
'
'            QRY = "INSERT INTO opmultas (munumero, mucodemp, mucodinf, munomnot, mufecha, munumpla, musitio, munomtes, muauxcon, " _
'                        & "mutipveh, muobserv, muestado, munetpag, muhora, mufecsis, musaldo, munomcaj, " _
'                        & "mudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete, mupasvia, " _
'                        & "mupaspla, mucedula, mutasapag) " _
'                        & "(SELECT " & sNumMulta & " AS munumero, mucodemp, mucodinf, munomnot, mufecha, munumpla, musitio, munomtes, muauxcon, " _
'                            & "mutipveh, muobserv, 'Pendiente' AS muestado, munetpag, muhora, mufecsis, musaldo, " _
'                            & "munomcaj, mudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete, mupasvia, " _
'                            & "mupaspla, mucedula, 0 AS mutasapag " _
'                            & "FROM opmultas " _
'                            & "WHERE munumero = " & Tbl1!munumero & " " _
'                            & "AND munumpla = '" & Txtplaca.text & "' " _
'                        & ") "
'
'            Dbs.Execute (QRY)
'
'            QRY = "UPDATE opabonos " _
'                        & "SET abestado = 'Anulado' " _
'                        & "WHERE abnumero = " & Tbl1!munumero & " "
'
'            Dbs.Execute (QRY)
'        End If
'        Tbl2.Close
'
'        'Actualizamos el número de la multa luego de insertar las necesarias
'        QRY = "UPDATE geparsis SET psval = " & sNumMulta & " WHERE pscod = 'NUMMUL' "
'
'        Dbs.Execute (QRY)
'
'    End If
'    Tbl1.Close
       
   QRY = " INSERT INTO gelogaud " _
               & "(lofecact, lofecnue, lofuncio, " _
               & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
               & "VALUES ( to_date('" & Now & "', 'MM/dd/yyyy hh24:mi:ss'),  to_date('" & Now & "', 'MM/dd/yyyy hh24:mi:ss'),'" _
               & Login & "', to_date('" & Now & "', 'MM/dd/yyyy hh24:mi:ss'),'SE ANULO UNA EVASION', 'ANUMUL', '" _
               & Txtnumero.text & "', '" & Mid(Trim(Txtobsanu.text), 1, 50) & "')"
        
   Dbs.CommitTrans
   On Error GoTo Errores
   MsgBox "Evasion " & estado, 64, "O.K."
   Btngrabar.Enabled = False
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
'  If Btngrabar.Enabled = False And Txtnumero.Enabled = False Then
'     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
'
'     Dbs.Execute "INSERT INTO optempo1 " _
'                            & "(temnumero1, temtext1, temnumero2, temtext2, temdate1, temtext3, temtext4, temtext5, temtext6, temtext7, temtext8, temnumdob1,  temtext9, temtext10, temnumdob2, temnumero3, temnumdob3, temdate2, tempc ) " _
'                    & "SELECT munumero,   mucodemp, mucodinf,   munomnot, mufecha, muhora,   munumpla,  musitio,  munomtes, muauxcon, mutipveh, mudescue,  muobserv, muestado,  munetpag,  murecpag, muvaldes, mufecsis, '" & Pc & "' " _
'                      & "FROM opmultas " _
'                     & "WHERE munumero = " & Txtnumero.Text & ""
'
'     reporte = RpPath + "\regmulta.rpt"
'     rptterminal.DataFiles(0) = DbPath
'     rptterminal.password = Passwd
'     rptterminal.ReportFileName = reporte
'     rptterminal.Formulas(0) = "empresa = '" & G_NomEmpresa & "'"
'     Canmultas.Enabled = False
'     rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
'     rptterminal.Destination = 0
'     rptterminal.Action = 1
'     Canmultas.Enabled = True
'     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
'  End If
End Sub

Private Sub Btnlimpiar_Click()
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
  txtCedula.text = ""
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
  OptEstado(2).Enabled = True
  tasapago = 0
  
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
  
  txtPlaca.Enabled = True
  txtPlaca.SetFocus
  Mskfecsis.text = ""
  fqry01 = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'HORASIS' "
 
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
   End If
   Ftb01.Close
     Txtobsanu.text = ""

End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Anumultas")
End Sub

Private Sub Form_Load()
  
   On Error GoTo errorres
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
 
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecsis.text = Fecha
   End If
   Ftb01.Close
   
   tasapago = 0
   
   Exit Sub
errorres:
  MsgBox Err.Description, 16, "Error " & Err.Number
  Exit Sub
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
              & "WHERE rucodigo = '" & Txtdespla.text & "' "
             
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         'wecc 02/27/2002 cuando la ruta esta inactiva
         If Ftb01!ruestado = "Inactiva" Then
            Ftb01.Close
            Txtdespla.SetFocus
            MsgBox "Ruta Inactiva o Inexistente", 48, "Advertencia"
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
              & "WHERE rucodigo = '" & Txtdestino.text & "' "
             
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
End Sub


Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Then
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
        txtPlaca.text = Ftb01!munumpla
        Txtnumero.text = Ftb01!munumero
        Txtcodemp.text = Ftb01!mucodemp
        Txtcodinf.text = Ftb01!mucodinf
        Txtnomnot.text = Ftb01!munomnot
        Mskfecha.text = Format(Ftb01!mufecha, "MM/DD/YYYY")
        Mskhora.text = Ftb01!muhora
        Txtsitio.text = Ftb01!musitio
        Txtnomtes.text = Ftb01!munomtes
        Txtnomaux.text = Ftb01!muauxcon
        Txttipveh.text = Ftb01!mutipveh
        Txtdescuento.text = Ftb01!mudescue
        TxtObservaciones.text = Ftb01!muobserv
        Msknetpag.text = Ftb01!munetpag
        Mskvalmul.text = Ftb01!munetpag
        If Not IsNull(Ftb01!mutasapag) Then
            tasapago = Ftb01!mutasapag
        End If
         txtCedula.text = ""
         If Not IsNull(Ftb01!mucedula) Then
            txtCedula.text = Ftb01!mucedula
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
         End If
        
        
        fqry01 = "SELECT * " _
                 & "FROM ophojveh " _
                & "WHERE hvplaca = '" & txtPlaca.text & "' "
             
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
                & "WHERE hecodigo = '" & Txtcodemp.text & "' "
        
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
        Txtauxalt_LostFocus
        If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
           Mskvalmul.text = (Msknetpag.text * 100) / (100 - Txtdescuento.text)
        End If
        
        
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
      
   

        Btngrabar.Enabled = True
        
     Else
       MsgBox "Numero De Evasion No Existe", 48, "Advertencia"
       Ftb01.Close
       Txtnumero.SetFocus
       Exit Sub
     End If
   End If
End Sub

Private Sub Txtobsanu_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Then
     KeyAscii% = 0
  End If

End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
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
  
      If longitud < 6 Then
         MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
         txtPlaca.SetFocus
         Exit Sub
      End If
     
      fqry01 = "SELECT * " _
               & "FROM opmultas " _
              & "WHERE munumpla = '" & txtPlaca.text & "' " _
                & "AND muestado = 'Pendiente' "
     
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         Txtnumero.Enabled = False
         txtPlaca.Enabled = False
         Txtnumero.text = Ftb01!munumero
         Txtcodemp.text = Ftb01!mucodemp
         Txtcodinf.text = Ftb01!mucodinf
         Txtnomnot.text = Ftb01!munomnot
         Mskfecha.text = Format(Ftb01!mufecha, "MM/DD/YYYY")
         Mskhora.text = Ftb01!muhora
         Txtsitio.text = Ftb01!musitio
         Txtnomtes.text = Ftb01!munomtes
         Txtnomaux.text = Ftb01!muauxcon
         Txttipveh.text = Ftb01!mutipveh
         Txtdescuento.text = Ftb01!mudescue
         txtcaja.text = Ftb01!munomcaj
         TxtObservaciones.text = Ftb01!muobserv
         Msknetpag.text = Ftb01!munetpag
         txtCedula.text = ""
         If Not IsNull(Ftb01!mucedula) Then
            txtCedula.text = Ftb01!mucedula
         End If
        Ftb01.Close
        fqry01 = "SELECT * " _
                 & "FROM ophojveh " _
                & "WHERE hvplaca = '" & txtPlaca.text & "' "
             
  
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
        
           'qry01 = "SELECT psval " _
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
        If Txtcodinf.text = "51" Or Txtcodinf.text = "52" Then
           Mskvalmul.text = (Msknetpag.text * 100) / (100 - Txtdescuento.text)
        End If
        
        Btngrabar.Enabled = True
     Else
        MsgBox "Codigo De Placa No Tiene Ningun Registro De Evasiones o Elusiones", 48, "Advertencia"
        Ftb01.Close
        txtPlaca.SetFocus
        Exit Sub
     End If
  End If
End Sub

Private Sub Txtrecpag_KeyPress(KeyAscii As Integer)
  If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Then
      KeyAscii% = 0
  End If
End Sub

