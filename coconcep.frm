VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Coconconcep 
   Caption         =   "Consulta Tasas de Uso y sus Conceptos"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnExportar 
      Caption         =   "&Exportar"
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
      Left            =   11160
      Picture         =   "coconcep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   2880
      Width           =   855
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
      TabIndex        =   66
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   480
         Width           =   1095
      End
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
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Txtnumini 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Número inicial por el que empiece a buscar"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtnumfin 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Número final hasta donde quiere que llegue"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "De Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1455
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
      TabIndex        =   50
      Top             =   120
      Width           =   3375
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   240
         TabIndex        =   3
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
         Left            =   1800
         TabIndex        =   4
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
         TabIndex        =   52
         Top             =   240
         Width           =   1215
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
         Left            =   1800
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
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
      Left            =   11160
      Picture         =   "coconcep.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Búscar informacion en la base de datos"
      Top             =   1200
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
      Left            =   11160
      Picture         =   "coconcep.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4560
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
      Left            =   11160
      Picture         =   "coconcep.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3720
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
      Left            =   11160
      Picture         =   "coconcep.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Totales:"
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
      Left            =   5400
      TabIndex        =   42
      Top             =   3240
      Width           =   5535
      Begin MSMask.MaskEdBox MskTotal 
         Height          =   300
         Left            =   3840
         TabIndex        =   43
         Top             =   1680
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
      Begin MSMask.MaskEdBox MskAlcohol 
         Height          =   300
         Left            =   1080
         TabIndex        =   44
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
      Begin MSMask.MaskEdBox MskTasas 
         Height          =   300
         Left            =   1080
         TabIndex        =   45
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
      Begin MSMask.MaskEdBox MskParqueadero 
         Height          =   300
         Left            =   1080
         TabIndex        =   46
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
      Begin MSMask.MaskEdBox MskTag 
         Height          =   300
         Left            =   1080
         TabIndex        =   55
         Top             =   1680
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
      Begin MSMask.MaskEdBox MskElusion 
         Height          =   300
         Left            =   3840
         TabIndex        =   56
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
      Begin MSMask.MaskEdBox MskEvasion 
         Height          =   300
         Left            =   3840
         TabIndex        =   58
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
      Begin MSMask.MaskEdBox MskEvasionTransito 
         Height          =   300
         Left            =   3840
         TabIndex        =   60
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
      Begin MSMask.MaskEdBox MskPDocumento 
         Height          =   300
         Left            =   3840
         TabIndex        =   62
         Top             =   1320
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
      Begin MSMask.MaskEdBox MskVigencias 
         Height          =   300
         Left            =   1080
         TabIndex        =   64
         Top             =   1320
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
      Begin VB.Label Label18 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2760
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Tag:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   63
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Perd. Doc:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   61
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "E. Transito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   59
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Evasion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2760
         TabIndex        =   57
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Vigencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Permanen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Alcohol:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Tasas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Elusion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   47
         Top             =   240
         Width           =   975
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
      TabIndex        =   41
      Top             =   2520
      Width           =   5055
      Begin VB.CheckBox chkOriTra 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkTipoServicio 
         Caption         =   "Tipo Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkRuta 
         Caption         =   "Cod. Ruta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkEmpresa 
         Caption         =   "Cod. Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   240
         Width           =   1575
      End
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
      TabIndex        =   32
      Top             =   3600
      Width           =   5055
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   15
         ToolTipText     =   "Código de la empresa (Precione doble click para ayuda)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Txtcodrut 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   16
         ToolTipText     =   "Código de la ruta (Precione doble click para ayuda)"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Txttipser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Txtoritra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   18
         ToolTipText     =   "Código del tipo de tasa de uso(Precione doble click para ayuda)"
         Top             =   1320
         Width           =   735
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
         TabIndex        =   40
         Top             =   240
         Width           =   1335
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
         TabIndex        =   39
         Top             =   600
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
         TabIndex        =   38
         Top             =   960
         Width           =   1215
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
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LblNomEmpresa 
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
         TabIndex        =   36
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label LblNomRuta 
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
         TabIndex        =   35
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label LblNomServicio 
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
         TabIndex        =   34
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label LblNomOriTra 
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
         TabIndex        =   33
         Top             =   1320
         Width           =   2535
      End
   End
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
      Height          =   1335
      Left            =   3480
      TabIndex        =   31
      Top             =   1080
      Width           =   1695
      Begin VB.ComboBox CmbTipoRuta 
         Height          =   315
         ItemData        =   "coconcep.frx":1412
         Left            =   120
         List            =   "coconcep.frx":141F
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   1455
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
      Height          =   1335
      Left            =   1920
      TabIndex        =   30
      Top             =   1080
      Width           =   1455
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
         TabIndex        =   7
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
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptTipoTasa 
         Caption         =   "Credito"
         BeginProperty Font 
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
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
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
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
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
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      Width           =   1695
      Begin VB.ComboBox Cmbcaja 
         Height          =   315
         ItemData        =   "coconcep.frx":1444
         Left            =   120
         List            =   "coconcep.frx":1446
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Estado de la empresa"
         Top             =   240
         Width           =   1455
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
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   1695
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "coconcep.frx":1448
         Left            =   120
         List            =   "coconcep.frx":144A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Tipos de terminales"
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grdconduces 
      Height          =   3015
      Left            =   5400
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   69
      Top             =   5520
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   11400
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   11040
      X2              =   11040
      Y1              =   5400
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   12600
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   5280
      X2              =   5280
      Y1              =   5400
      Y2              =   -120
   End
End
Attribute VB_Name = "Coconconcep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sebastián Rondón - 04 Marzo de 2025
'Conceptos multas
Dim L_CiTElusion As String
Dim L_CiTEvasion As String
Dim L_CiTPerdida As String
Dim L_CiTExcreta As String
Dim L_CiTEvasTra As String
Dim L_Todas_Multas As String

'Variables para los totales
Dim L_Total_Tasas As Double
Dim L_Total_Alcohol As Double
Dim L_Total_Permanencia As Double
Dim L_Total_Vigencia As Double
Dim L_Total_Elusion As Double
Dim L_Total_Evasion As Double
Dim L_Total_Evasion_Transito As Double
Dim L_Total_P_Documento As Double
Dim L_Total_Tag As Double
Dim L_Total As Double


Private Sub DIB_GRILLA()
    Grdconduces.rows = 1
    Grdconduces.Cols = 17
    Grdconduces.Clear
    
    s$ = "Tasa de Uso |Pref. Fact. |Num. Fact. |Cod. Empresa |Nom. Empresa |NIT |Vr. TU Origen |Vr. TU Transito |Vr. Alcohol |Vr. Permanencia " _
            & "|Vr. Vigencias |Vr. Elusion |Vr. Evasion |Vr. Evasion Transito |Vr. P. Documento |Vr Tag |Total"
            
    Grdconduces.FormatString = s$
    Grdconduces.Row = 0
    Grdconduces.Col = 0
    
    Grdconduces.ColWidth(0) = 1200
    Grdconduces.ColWidth(1) = 1200
    Grdconduces.ColWidth(2) = 1200
    Grdconduces.ColWidth(3) = 1200
    Grdconduces.ColWidth(4) = 1200
    Grdconduces.ColWidth(5) = 1200
    Grdconduces.ColWidth(6) = 1200
    Grdconduces.ColWidth(7) = 1200
    Grdconduces.ColWidth(8) = 1200
    Grdconduces.ColWidth(9) = 1200
    Grdconduces.ColWidth(10) = 1200
    Grdconduces.ColWidth(11) = 1200
    Grdconduces.ColWidth(12) = 1200
    Grdconduces.ColWidth(13) = 1200
    Grdconduces.ColWidth(14) = 1200
    Grdconduces.ColWidth(15) = 1200
    Grdconduces.ColWidth(16) = 1200
    
    For i = 0 To Grdconduces.Cols - 1
        Grdconduces.Col = i
        Grdconduces.CellFontBold = True
    Next i
End Sub

Private Sub BtnBuscar_Click()
    'Limpiamos los valores generales
    L_Total_Tasas = 0
    L_Total_Alcohol = 0
    L_Total_Permanencia = 0
    L_Total_Vigencia = 0
    L_Total_Elusion = 0
    L_Total_Evasion = 0
    L_Total_Evasion_Transito = 0
    L_Total_P_Documento = 0
    L_Total_Tag = 0
    L_Total = 0

    'Declaramos las variables a utilizar
    Dim sPref, sNumFac As String
    Dim sValorOrigen, sValorTransito, sValorAlcohol As Double
    Dim sValorPermanencia, sValorVigencia As Double
    Dim sValorElusion, sValorEvasion, sValorEvasionTransito, sValorPDocumento As Double
    Dim sValorTag As Double
    Dim sTotal As Double
    
    'Inicializamos el ProgressBar
    ProBar.Value = ProBar.Min
    
    On Error GoTo Errores
    
    'Borramos los datos de la tabla temporal
    Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"

    'Validamos el tipo de busqueda
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
    ElseIf Optbuscar(1).Value = True Then
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
    End If
    
    'Validamos el resto de los campos
    If chkEmpresa.Value = 1 Then
        If Txtcodemp.text = "" Then
            MsgBox "Falta digita el código de la empresa", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
        End If
    End If
    
    If chkRuta.Value = 1 Then
        If Txtcodrut.text = "" Then
            MsgBox "Falta digitar el código de ruta", 48, "Advertencia"
            Txtcodrut.SetFocus
            Exit Sub
        End If
    End If
    
    If chkTipoServicio.Value = 1 Then
        If Txttipser.text = "" Then
            MsgBox "Falta digitar el código del tipo de servicio", 48, "Advertencia"
            Txttipser.SetFocus
            Exit Sub
        End If
    End If
    
    If chkOriTra.Value = 1 Then
        If Txtoritra.text = "" Then
            MsgBox "Falta digita el código de Tipo de Despacho", 48, "Advertencia"
            Txtoritra.SetFocus
            Exit Sub
        End If
    End If
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    'Realizamos la consulta
    QRY = "SELECT conumero, cocodemp, cocodrut, henombre, henit, cooritra, covalor, covalalc, ruinflue " _
                & "FROM coconduc " _
                & "INNER JOIN ophojemp ON ophojemp.hecodigo = coconduc.cocodemp " _
                & "INNER JOIN opruta ON opruta.rucodigo = coconduc.cocodrut "
                
    If Optbuscar(0).Value = True Then
        QRY = QRY & "WHERE conumero BETWEEN " & Txtnumini.text & " AND " & Txtnumfin.text & " "
    ElseIf Optbuscar(1).Value = True Then
        QRY = QRY & "WHERE cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                    & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY') "
    End If
    
    'Si seleccionó terminal
    If CmbTerminal.text <> "GENERAL" Then
        QRY = QRY & "AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "' "
    End If
    
    'Si seleccionó caja
    If Cmbcaja.text <> "" Then
        QRY = QRY & "AND conomcaj = '" & Cmbcaja.text & "' "
    End If
    
    'Si seleccionó un tipo de ruta
    If CmbTipoRuta.text = "SABANERA" Then
        QRY = QRY & "AND ruinflue = 'Si' "
    ElseIf CmbTipoRuta.text = "INTERMUN." Then
        QRY = QRY & "AND ruinflue = 'No' "
    End If
    
    'Si seleccionó tipo de tasa
    If OptTipoTasa(0).Value = True Then
        QRY = QRY & "AND comodo = 'P' "
    ElseIf OptTipoTasa(1).Value = True Then
        QRY = QRY & "AND cocredito = 'S' "
    ElseIf OptTipoTasa(3).Value = True Then
        QRY = QRY & "AND cooritra = 'SV' "
    End If
    
    'Si busca por código de empresa
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
    End If
    
    'Si busca por código de ruta
    If chkRuta.Value = 1 Then
        QRY = QRY & "AND cocodrut = '" & Txtcodrut.text & "' "
    End If
    
    'Si busca por tipo de servicio
    If chkTipoServicio.Value = 1 Then
        QRY = QRY & "AND coclase = '" & Txttipser.text & "' "
    End If
    
    'Si busca por tipo de despacho
    If chkOriTra.Value = 1 Then
        QRY = QRY & "AND cooritra = '" & Txtoritra.text & "' "
    End If
    
    QRY = QRY & "AND coestado = 'Activo' " _
                & "ORDER BY conumero, cofecsal, cohorven "
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        
        ProBar.Max = Tbl1.RecordCount + 2
        
        'Inhabilitamos la pantalla
        Coconconcep.Enabled = False
        
        Do While Not Tbl1.EOF
        
            'Inicializamos las variables
            sPref = "N/A"
            sNumFac = "N/A"
            
            sValorOrigen = 0
            sValorTransito = 0
            sValorAlcohol = 0
            sValorPermanencia = 0
            sValorVigencia = 0
            sValorElusion = 0
            sValorEvasion = 0
            sValorEvasionTransito = 0
            sValorPDocumento = 0
            sValorTag = 0
            sTotal = 0
            
            'Consultamos el número de la factura con que quedo la Tasa de Uso en la DIAN
            QRY = "SELECT tfnumfac, tfprefijo " _
                        & "FROM cotufe " _
                        & "WHERE tfnumtu = " & Tbl1!conumero & " "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sPref = Tbl2!tfprefijo
                sNumFac = Trim(str(Tbl2!tfnumfac))
            End If
            Tbl2.Close
            
            'Asignamos valor por Origen o Transito
            If Tbl1!Cooritra = "O" Then
                sValorOrigen = CDbl(Tbl1!covalor)
            ElseIf Tbl1!Cooritra = "T" Then
                sValorTransito = CDbl(Tbl1!covalor)
            End If
            
            'Asignamos valor por Alcoholimetria
            sValorAlcohol = Tbl1!covalalc
            
            'Buscamos lo pagado por Permanencias
            QRY = "SELECT NVL(SUM(pavalor), 0) As TotalPermanencia " _
                        & "FROM coparqueop " _
                        & "WHERE pafecpag IS NOT NULL " _
                        & "AND patipo = 'P' " _
                        & "AND paestado = 'C' " _
                        & "AND patasapag = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorPermanencia = CDbl(Tbl2!TotalPermanencia)
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Vigencias
            QRY = "SELECT NVL(SUM(pavalor), 0) As TotalVigencia " _
                        & "FROM coparqueop " _
                        & "WHERE pafecpag IS NOT NULL " _
                        & "AND patipo = 'V' " _
                        & "AND paestado = 'C' " _
                        & "AND patasapag = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorVigencia = CDbl(Tbl2!TotalVigencia)
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Multas
            QRY = "SELECT NVL(SUM(abvalor), 0) AS TotalMulta, mucodinf " _
                        & "FROM opabonos " _
                        & "INNER JOIN opmultas ON opmultas.munumero = opabonos.abnumero " _
                        & "WHERE abfecha IS NOT NULL " _
                        & "AND munumero = abnumero " _
                        & "AND abestado = 'Activo' " _
                        & "AND mutasapag = " & Tbl1!conumero & " " _
                        & "GROUP BY mucodinf "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                Do While Not Tbl2.EOF
                    If Tbl2!mucodinf = L_CiTElusion Then
                        sValorElusion = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTEvasion Then
                        sValorEvasion = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTEvasTra Then
                        sValorEvasionTransito = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTPerdida Then
                        sValorPDocumento = CDbl(Tbl2!TotalMulta)
                    End If
                    Tbl2.MoveNext
                Loop
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Tag
            QRY = "SELECT NVL(SUM(tavalor), 0) As TotalTag " _
                        & "FROM optag " _
                        & "WHERE tafecpag IS NOT NULL " _
                        & "AND taestado = 'A' " _
                        & "AND tanumtasa = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorTag = CDbl(Tbl2!TotalTag)
            End If
            Tbl2.Close
            
            'Calculamos el Total
            sTotal = sValorOrigen + sValorTransito + sValorAlcohol + sValorPermanencia + sValorVigencia + _
                        sValorElusion + sValorEvasion + sValorEvasionTransito + sValorPDocumento + sValorTag
            
            'Llenamos los datos en la grilla
            entry = str(Tbl1!conumero) + Chr(9) + sPref + Chr(9) + sNumFac + Chr(9) + Tbl1!cocodemp + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + _
                        str(sValorOrigen) + Chr(9) + str(sValorTransito) + Chr(9) + str(sValorAlcohol) + Chr(9) + _
                        str(sValorPermanencia) + Chr(9) + str(sValorVigencia) + Chr(9) + str(sValorElusion) + Chr(9) + _
                        str(sValorEvasion) + Chr(9) + str(sValorEvasionTransito) + Chr(9) + str(sValorPDocumento) + Chr(9) + _
                        str(sValorTag) + Chr(9) + str(sTotal)
                        
            'Inserto los datos en una tabla temporal para la exportación del reporte
            QRY = "INSERT INTO optempo1 " _
                        & "(temnumero1, temtext1, temtext2, temtext3, temnumdob1, temnumdob2, temnumdob3, " _
                        & "temnumdob4, temnumdob5, temnumdob6, temnumdob7, temnumdob8, temnumdob9, temnumdob10, " _
                        & "temnummon1, tempc, temtext4, temtext5) VALUES(" _
                        & Tbl1!conumero & ", '" & Tbl1!cocodemp & "', '" & Tbl1!henombre & "', '" & Tbl1!henit & "', " _
                        & sValorOrigen & ", " & sValorTransito & ", " & sValorAlcohol & ", " & sValorPermanencia & ", " _
                        & sValorVigencia & ", " & sValorElusion & ", " & sValorEvasion & ", " & sValorEvasionTransito & ", " _
                        & sValorPDocumento & ", " & sValorTag & ", " & sTotal & ", '" & Pc & "', '" _
                        & sPref & "', '" & sNumFac & "') "
                        
            Dbs.Execute (QRY)
                        
            Grdconduces.AddItem entry
            
            'Sumamos los valores para los totales generales
            L_Total_Tasas = L_Total_Tasas + sValorOrigen + sValorTransito
            L_Total_Alcohol = L_Total_Alcohol + sValorAlcohol
            L_Total_Permanencia = L_Total_Permanencia + sValorPermanencia
            L_Total_Vigencia = L_Total_Vigencia + sValorVigencia
            L_Total_Elusion = L_Total_Elusion + sValorElusion
            L_Total_Evasion = L_Total_Evasion + sValorEvasion
            L_Total_Evasion_Transito = L_Total_Evasion_Transito + sValorEvasionTransito
            L_Total_P_Documento = L_Total_P_Documento + sValorPDocumento
            L_Total_Tag = L_Total_Tag + sValorTag
            L_Total = L_Total + sTotal
            
            ProBar.Value = ProBar.Value + 1
            
            Tbl1.MoveNext
        Loop
        
        'Mostramos los totales en pantalla
        MskTasas.text = L_Total_Tasas
        MskAlcohol.text = L_Total_Alcohol
        MskParqueadero.text = L_Total_Permanencia
        MskVigencias.text = L_Total_Vigencia
        MskElusion.text = L_Total_Elusion
        MskEvasion.text = L_Total_Evasion
        MskEvasionTransito.text = L_Total_Evasion_Transito
        MskPDocumento.text = L_Total_P_Documento
        MskTag.text = L_Total_Tag
        MskTotal.text = L_Total
        
        ProBar.Value = ProBar.Max
        Coconconcep.Enabled = True
    Else
        MsgBox "No se encontró información con los datos digitados", 48, "Advertencia"
        Exit Sub
    End If
    
    Exit Sub
Errores:
    MsgBox Err.Description, 48, "Advertencia"
    Coconconcep.Enabled = True
    Exit Sub
    Resume
End Sub

Private Sub BtnExportar_Click()
    On Error GoTo Errores
    
    'Limpiamos los valores generales
    L_Total_Tasas = 0
    L_Total_Alcohol = 0
    L_Total_Permanencia = 0
    L_Total_Vigencia = 0
    L_Total_Elusion = 0
    L_Total_Evasion = 0
    L_Total_Evasion_Transito = 0
    L_Total_P_Documento = 0
    L_Total_Tag = 0
    L_Total = 0
    
    'Inicializamos el ProgressBar
    ProBar.Value = ProBar.Min
    
    On Error GoTo Errores
    
    'Borramos los datos de la tabla temporal
    Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"

    'Validamos el tipo de busqueda
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
    ElseIf Optbuscar(1).Value = True Then
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
    End If
    
    'Validamos el resto de los campos
    If chkEmpresa.Value = 1 Then
        If Txtcodemp.text = "" Then
            MsgBox "Falta digita el código de la empresa", 48, "Advertencia"
            Txtcodemp.SetFocus
            Exit Sub
        End If
    End If
    
    If chkRuta.Value = 1 Then
        If Txtcodrut.text = "" Then
            MsgBox "Falta digitar el código de ruta", 48, "Advertencia"
            Txtcodrut.SetFocus
            Exit Sub
        End If
    End If
    
    If chkTipoServicio.Value = 1 Then
        If Txttipser.text = "" Then
            MsgBox "Falta digitar el código del tipo de servicio", 48, "Advertencia"
            Txttipser.SetFocus
            Exit Sub
        End If
    End If
    
    If chkOriTra.Value = 1 Then
        If Txtoritra.text = "" Then
            MsgBox "Falta digita el código de Tipo de Despacho", 48, "Advertencia"
            Txtoritra.SetFocus
            Exit Sub
        End If
    End If
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    'Realizamos la consulta
    QRY = "SELECT conumero, cocodemp, cocodrut, henombre, henit, cooritra, covalor, covalalc, ruinflue " _
                & "FROM coconduc " _
                & "INNER JOIN ophojemp ON ophojemp.hecodigo = coconduc.cocodemp " _
                & "INNER JOIN opruta ON opruta.rucodigo = coconduc.cocodrut "
                
    If Optbuscar(0).Value = True Then
        QRY = QRY & "WHERE conumero BETWEEN " & Txtnumini.text & " AND " & Txtnumfin.text & " "
    ElseIf Optbuscar(1).Value = True Then
        QRY = QRY & "WHERE cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                    & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY') "
    End If
    
    'Si seleccionó terminal
    If CmbTerminal.text <> "GENERAL" Then
        QRY = QRY & "AND coterminal = '" & Mid(CmbTerminal.text, 1, 1) & "' "
    End If
    
    'Si seleccionó caja
    If Cmbcaja.text <> "" Then
        QRY = QRY & "AND conomcaj = '" & Cmbcaja.text & "' "
    End If
    
    'Si seleccionó un tipo de ruta
    If CmbTipoRuta.text = "SABANERA" Then
        QRY = QRY & "AND ruinflue = 'Si' "
    ElseIf CmbTipoRuta.text = "INTERMUN." Then
        QRY = QRY & "AND ruinflue = 'No' "
    End If
    
    'Si seleccionó tipo de tasa
    If OptTipoTasa(0).Value = True Then
        QRY = QRY & "AND comodo = 'P' "
    ElseIf OptTipoTasa(1).Value = True Then
        QRY = QRY & "AND cocredito = 'S' "
    ElseIf OptTipoTasa(3).Value = True Then
        QRY = QRY & "AND cooritra = 'SV' "
    End If
    
    'Si busca por código de empresa
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND cocodemp = '" & Txtcodemp.text & "' "
    End If
    
    'Si busca por código de ruta
    If chkRuta.Value = 1 Then
        QRY = QRY & "AND cocodrut = '" & Txtcodrut.text & "' "
    End If
    
    'Si busca por tipo de servicio
    If chkTipoServicio.Value = 1 Then
        QRY = QRY & "AND coclase = '" & Txttipser.text & "' "
    End If
    
    'Si busca por tipo de despacho
    If chkOriTra.Value = 1 Then
        QRY = QRY & "AND cooritra = '" & Txtoritra.text & "' "
    End If
    
    QRY = QRY & "AND coestado = 'Activo' " _
                & "ORDER BY conumero, cofecsal, cohorven "
    
    Coconconcep.Enabled = False
    ExportarExcel (QRY)
    Coconconcep.Enabled = True
    
    Exit Sub
    
Errores:
    Coconconcep.Enabled = True
    MsgBox Err.Description, 48, "Advertencia"
    Exit Sub
    Resume
End Sub

Private Sub ExportarExcel(ByVal QRY As String)

    'Declaramos las variables a utilizar
    Dim sPref, sNumFac As String
    Dim sValorOrigen, sValorTransito, sValorAlcohol As Double
    Dim sValorPermanencia, sValorVigencia As Double
    Dim sValorElusion, sValorEvasion, sValorEvasionTransito, sValorPDocumento As Double
    Dim sValorTag As Double
    Dim sTotal As Double
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "TASA USO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "PREF. FACTURA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "NUM. FACTURA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "COD. EMPRESA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "EMPRESA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "NIT"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "VR. TU ORIGEN"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "VR. TU TRANSITO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "VR. ALCOHOL"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "VR. PERMANENCIA"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "VR. VIGENCIA"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "VR. ELUSION"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "VR. EVASION"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "VR. EVASION T."
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "VR. P. DOCUMENTO"
    ApExcel.Cells(fila, 15).Font.Size = 10
    ApExcel.Cells(fila, 16).formula = "VR. TAG"
    ApExcel.Cells(fila, 16).Font.Size = 10
    ApExcel.Cells(fila, 17).formula = "TOTAL"
    ApExcel.Cells(fila, 17).Font.Size = 10
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
    
        ProBar.Max = Tbl1.RecordCount + 2
        
        'Inhabilitamos la pantalla
        Coconconcep.Enabled = False
        Do While Not Tbl1.EOF
            'Inicializamos las variables
            sPref = "N/A"
            sNumFac = "N/A"
            
            sValorOrigen = 0
            sValorTransito = 0
            sValorAlcohol = 0
            sValorPermanencia = 0
            sValorVigencia = 0
            sValorElusion = 0
            sValorEvasion = 0
            sValorEvasionTransito = 0
            sValorPDocumento = 0
            sValorTag = 0
            sTotal = 0
            
            'Consultamos el número de la factura con que quedo la Tasa de Uso en la DIAN
            QRY = "SELECT tfnumfac, tfprefijo " _
                        & "FROM cotufe " _
                        & "WHERE tfnumtu = " & Tbl1!conumero & " "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sPref = Tbl2!tfprefijo
                sNumFac = Trim(str(Tbl2!tfnumfac))
            End If
            Tbl2.Close
            
            'Asignamos valor por Origen o Transito
            If Tbl1!Cooritra = "O" Then
                sValorOrigen = CDbl(Tbl1!covalor)
            ElseIf Tbl1!Cooritra = "T" Then
                sValorTransito = CDbl(Tbl1!covalor)
            End If
            
            'Asignamos valor por Alcoholimetria
            sValorAlcohol = Tbl1!covalalc
            
            'Buscamos lo pagado por Permanencias
            QRY = "SELECT NVL(SUM(pavalor), 0) As TotalPermanencia " _
                        & "FROM coparqueop " _
                        & "WHERE pafecpag IS NOT NULL " _
                        & "AND patipo = 'P' " _
                        & "AND paestado = 'C' " _
                        & "AND patasapag = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorPermanencia = CDbl(Tbl2!TotalPermanencia)
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Vigencias
            QRY = "SELECT NVL(SUM(pavalor), 0) As TotalVigencia " _
                        & "FROM coparqueop " _
                        & "WHERE pafecpag IS NOT NULL " _
                        & "AND patipo = 'V' " _
                        & "AND paestado = 'C' " _
                        & "AND patasapag = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorVigencia = CDbl(Tbl2!TotalVigencia)
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Multas
            QRY = "SELECT NVL(SUM(abvalor), 0) AS TotalMulta, mucodinf " _
                        & "FROM opabonos " _
                        & "INNER JOIN opmultas ON opmultas.munumero = opabonos.abnumero " _
                        & "WHERE abfecha IS NOT NULL " _
                        & "AND munumero = abnumero " _
                        & "AND abestado = 'Activo' " _
                        & "AND mutasapag = " & Tbl1!conumero & " " _
                        & "GROUP BY mucodinf "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                Do While Not Tbl2.EOF
                    If Tbl2!mucodinf = L_CiTElusion Then
                        sValorElusion = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTEvasion Then
                        sValorEvasion = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTEvasTra Then
                        sValorEvasionTransito = CDbl(Tbl2!TotalMulta)
                    End If
                    If Tbl2!mucodinf = L_CiTPerdida Then
                        sValorPDocumento = CDbl(Tbl2!TotalMulta)
                    End If
                    Tbl2.MoveNext
                Loop
            End If
            Tbl2.Close
            
            'Buscamos lo pagado por Tag
            QRY = "SELECT NVL(SUM(tavalor), 0) As TotalTag " _
                        & "FROM optag " _
                        & "WHERE tafecpag IS NOT NULL " _
                        & "AND taestado = 'A' " _
                        & "AND tanumtasa = " & Tbl1!conumero & " "
                        
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sValorTag = CDbl(Tbl2!TotalTag)
            End If
            Tbl2.Close
            
            'Calculamos el Total
            sTotal = sValorOrigen + sValorTransito + sValorAlcohol + sValorPermanencia + sValorVigencia + _
                        sValorElusion + sValorEvasion + sValorEvasionTransito + sValorPDocumento + sValorTag
                        
            'Llenamos los datos en Excel
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!conumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & sPref & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & sNumFac & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!cocodemp & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!henit & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & sValorOrigen & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & sValorTransito & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & sValorAlcohol & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & sValorPermanencia & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & sValorVigencia & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & sValorElusion & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & sValorEvasion & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & sValorEvasionTransito & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & sValorPDocumento & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            ApExcel.Cells(fila, 16).formula = "" & sValorTag & ""
            ApExcel.Cells(fila, 16).Font.Size = 10
            ApExcel.Cells(fila, 17).formula = "" & sTotal & ""
            ApExcel.Cells(fila, 17).Font.Size = 10
            
            'Sumamos los valores para los totales generales
            L_Total_Tasas = L_Total_Tasas + sValorOrigen + sValorTransito
            L_Total_Alcohol = L_Total_Alcohol + sValorAlcohol
            L_Total_Permanencia = L_Total_Permanencia + sValorPermanencia
            L_Total_Vigencia = L_Total_Vigencia + sValorVigencia
            L_Total_Elusion = L_Total_Elusion + sValorElusion
            L_Total_Evasion = L_Total_Evasion + sValorEvasion
            L_Total_Evasion_Transito = L_Total_Evasion_Transito + sValorEvasionTransito
            L_Total_P_Documento = L_Total_P_Documento + sValorPDocumento
            L_Total_Tag = L_Total_Tag + sValorTag
            L_Total = L_Total + sTotal
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            
            Tbl1.MoveNext
        Loop
        Tbl1.Close
        
        'Mostramos los totales en pantalla
        MskTasas.text = L_Total_Tasas
        MskAlcohol.text = L_Total_Alcohol
        MskParqueadero.text = L_Total_Permanencia
        MskVigencias.text = L_Total_Vigencia
        MskElusion.text = L_Total_Elusion
        MskEvasion.text = L_Total_Evasion
        MskEvasionTransito.text = L_Total_Evasion_Transito
        MskPDocumento.text = L_Total_P_Documento
        MskTag.text = L_Total_Tag
        MskTotal.text = L_Total
        
        ApExcel.Worksheets(1).Name = "CONSULTA TASA POR TASA"
        ApExcel.Visible = True
        
        ProBar.Value = ProBar.Max
        Coconconcep.Enabled = True
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub
Errores:
    Coconconcep.Enabled = True
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
    On Error GoTo Errores
    
    If Grdconduces.rows > 1 Then
        'Consultamos la fecha actual
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'HORASIS' "
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        Fecha = ""
        If Not Tbl1.EOF Then
            Fecha = Mid(Tbl1!psval, 1, 10)
        End If
        
        If Optbuscar(0).Value = True Then
            inic = Txtnumini.text
            final = Txtnumfin.text
        ElseIf Optbuscar(1).Value = True Then
            inic = Mskfecini.FormattedText
            final = Mskfecfin.FormattedText
        End If
        
        imprimir.Show 1
        If G_Imprimir = "0" Then
          Exit Sub
        End If
        
        reporte = RpPath + "\coconconcep.rpt"
        rptterminal.ReportFileName = reporte
        rptterminal.Connect = G_Conexion_reportes
        rptterminal.Formulas(0) = "inicio = '" & inic & "' "
        rptterminal.Formulas(1) = "final = '" & final & "' "
        rptterminal.Formulas(2) = "fecha = '" & Fecha & "' "
        Coconconcep.Enabled = False
        rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
        rptterminal.WindowState = crptMaximized
        rptterminal.PrinterName = G_PrinterName
        rptterminal.PrinterDriver = G_PrinterDriver
        rptterminal.PrinterPort = G_PrinterPort
        rptterminal.Destination = 0
        rptterminal.Action = 1
        Coconconcep.Enabled = True
    End If
    
    Exit Sub
Errores:
    Coconconcep.Enabled = True
    MsgBox Err.Description, 48, "Advertencia"
    Exit Sub
    Resume
End Sub

Private Sub Btnlimpiar_Click()

    'Limpiamos el ProgressBar
    ProBar.Value = ProBar.Min
    
    'Limpiamos la grilla
    DIB_GRILLA
    
    'Limpiamos los campos de búsqueda
    Optbuscar(1).Value = True
    Txtnumini.text = ""
    Txtnumfin.text = ""
    Mskfecini.text = "__/__/____"
    Mskfecfin.text = "__/__/____"
    CmbTerminal.ListIndex = 0
    Cmbcaja.ListIndex = 0
    CmbTipoRuta.ListIndex = 0
    OptTipoTasa(2).Value = True
    chkEmpresa.Value = 0
    chkRuta.Value = 0
    chkTipoServicio.Value = 0
    chkOriTra.Value = 0
    
    'Limpiamos los campos de los totales
    MskTasas.text = ""
    MskAlcohol.text = ""
    MskParqueadero.text = ""
    MskVigencias.text = ""
    MskElusion.text = ""
    MskEvasion.text = ""
    MskEvasionTransito.text = ""
    MskPDocumento.text = ""
    MskTag.text = ""
    MskTotal.text = ""
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub chkEmpresa_Click()
    If chkEmpresa.Value = 1 Then
        Txtcodemp.Enabled = True
    Else
        Txtcodemp.text = ""
        LblNomEmpresa.Caption = ""
        Txtcodemp.Enabled = False
    End If
End Sub

Private Sub chkOriTra_Click()
    If chkOriTra.Value = 1 Then
        Txtoritra.Enabled = True
    Else
        Txtoritra.text = ""
        LblNomOriTra.Caption = ""
        Txtoritra.Enabled = False
    End If
End Sub

Private Sub chkRuta_Click()
    If chkRuta.Value = 1 Then
        Txtcodrut.Enabled = True
    Else
        Txtcodrut.text = ""
        LblNomRuta.Caption = ""
        Txtcodrut.Enabled = False
    End If
End Sub

Private Sub chkTipoServicio_Click()
    If chkTipoServicio.Value = 1 Then
        Txttipser.Enabled = True
    Else
        Txttipser.text = ""
        LblNomServicio.Caption = ""
        Txttipser.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    'Inicializamos el ProgressBar
    ProBar.Value = ProBar.Min

    'Inicializamos la Grilla
    DIB_GRILLA
    
    'Llenamos la lista de terminales
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
    
    'Llenamos la lista de tipos de rutas
    CmbTipoRuta.Clear
    CmbTipoRuta.AddItem "GENERAL"
    CmbTipoRuta.AddItem "INTERMUN."
    CmbTipoRuta.AddItem "SABANERA"
    CmbTipoRuta.ListIndex = 0
    
    'Llenamos la lista de cajas
    QRY = " SELECT canomcaj FROM cocaja"
    Set Tbl1 = Dbs.Execute(QRY)
    
    Cmbcaja.Clear
    Cmbcaja.AddItem ""
    
    Do While Not Tbl1.EOF
        Cmbcaja.AddItem Tbl1(0)
        Tbl1.MoveNext
    Loop
    Tbl1.Close
    Cmbcaja.ListIndex = 0
    
    L_CiTElusion = Fn_Parametros("CITELUSION", 1)
    L_CiTEvasion = Fn_Parametros("CITEVASION", 1)
    L_CiTPerdida = Fn_Parametros("CITPERDIDA", 1)
    L_CiTEvasTra = Fn_Parametros("CITEVASTRA", 1)
    L_CiTExcreta = Fn_Parametros("CITEXCRETA", 1)
    
    Optbuscar(1).Value = True
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
    forma1 = "73"
    LblNomEmpresa.Caption = ""
    buscar.Show 1
    txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub txtcodemp_LostFocus()
    If Txtcodemp.text <> "" Then
        
        'Si ya fue digitada la ruta, verificamos que pertenezca a la empresa
        If Txtcodrut.text <> "" Then
            QRY = "SELECT * " _
                    & "FROM opruthor " _
                    & "WHERE rhcodemp = '" & Txtcodemp.text & "' " _
                    & "AND rhcodrut = '" & txtRuta.text & "' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                Tbl1.Close
                MsgBox "La empresa no tiene autorizada esta ruta", 48, "Advertencia"
                txtRuta.SetFocus
                txtRuta.text = ""
                Exit Sub
            Else
                If Tbl1!rhestado = "Inactiva" Then
                    Tbl1.Close
                    txtRuta.text = ""
                    txtRuta.SetFocus
                    MsgBox "La ruta esta inactiva para esta empresa"
                End If
            End If
            Tbl1.Close
        End If
        
        'Verificamos que la empresa exista
        QRY = "SELECT henombre FROM ophojemp " _
            & "WHERE hecodigo = '" & Txtcodemp.text & "' "
            
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
            LblNomEmpresa.Caption = Tbl1!henombre
            Tbl1.Close
        Else
            
            QRY = "SELECT henombre FROM ophojempcon " _
                & "WHERE hecodigo = '" & Txtcodemp.text & "' "
                
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Tbl2.EOF = False Then
                LblNomEmpresa.Caption = Tbl2!henombre
            Else
                Tbl1.Close
                MsgBox "Empresa no existente", 48, "Advertencia"
                Txtcodemp.text = ""
                txtCodEmp_DblClick
            End If
            Tbl2.Close
        End If
    End If
End Sub

Private Sub txtcodrut_DblClick()
    forma = "3"
    If Txtcodemp.text <> "" Then
        forma1 = "36"
        G_Codigo_Empresa = Txtcodemp.text
    Else
        forma1 = "35"
    End If
    LblNomRuta.Caption = ""
    buscar.Show 1
    Txtcodrut_LostFocus
End Sub

Private Sub Txtcodrut_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub txtcodrut_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub Txtcodrut_LostFocus()
    If Txtcodrut.text <> "" Then
        If Txtcodemp.text <> "" Then
            QRY = "SELECT * " _
                    & "FROM opruthor " _
                    & "WHERE rhcodemp = '" & Txtcodemp.text & "' " _
                    & "AND rhcodrut = '" & Txtcodrut.text & "' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                Tbl1.Close
                Txtcodrut.SetFocus
                MsgBox "La empresa no tiene autorizada esta ruta", 48, "Advertencia"
                Txtcodrut.text = ""
                Exit Sub
            Else
                If Tbl1!rhestado = "Inactiva" Then
                    Tbl1.Close
                    Txtcodrut.text = ""
                    MsgBox "La ruta esta inactiva para esta empresa", 48, "Advertencia"
                    Exit Sub
                End If
            End If
            Tbl1.Close
        End If
        
        QRY = "SELECT rucodigo, rudestino " _
                & "FROM opruta " _
                & "WHERE rucodigo = '" & Txtcodrut.text & "' "
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF Then
            MsgBox "La ruta ingresada no existe", 48, "Advertencia"
            Txtcodrut.text = ""
            Txtcodrut.SetFocus
            Tbl1.Close
        Else
            LblNomRuta.Caption = Tbl1!rudestino
        End If
        Tbl1.Close
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

Private Sub Txtoritra_DblClick()
    forma = "15"
    forma1 = "67"
    Tabla_Catalogo = "TIPOS DE CONDUCE"
    LblNomOriTra.Caption = ""
    buscar.Show 1
    Txtoritra_LostFocus
End Sub

Private Sub Txtoritra_GotFocus()
    Senalar_Texto Me
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
           LblNomOriTra.Caption = Ftb01!dsdes
           Ftb01.Close
       Else
           MsgBox "Codigo De Origen - Transito No Existe", 48, "Advertencia"
           forma = "15"
           forma1 = "67"
           Tabla_Catalogo = "TIPOS DE CONDUCE"
           buscar.Show 1
           Txtoritra.SetFocus
       End If
  End If
End Sub

Private Sub Txttipser_DblClick()
    forma = "15"
    forma1 = "66"
    Tabla_Catalogo = "TIPOS DE VEHICULO"
    LblNomServicio.Caption = ""
    buscar.Show 1
    Txttipser_LostFocus
End Sub

Private Sub Txttipser_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub Txttipser_KeyPress(KeyAscii As Integer)
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
     
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        LblNomServicio.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        MsgBox "Tipo De Servicio No Existe", 48, "Advertencia"
        forma = "15"
        forma1 = "66"
        Tabla_Catalogo = "TIPOS DE VEHICULO"
        buscar.Show 1
        Txttipser.SetFocus
     End If
  End If
End Sub
