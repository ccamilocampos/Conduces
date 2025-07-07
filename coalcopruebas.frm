VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form coalcopruebas 
   Caption         =   "Reporte AlcoPruebas"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
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
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   2895
      Begin VB.OptionButton OptConduc 
         Caption         =   "Todos"
         BeginProperty Font 
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
         TabIndex        =   44
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton OptConduc 
         Caption         =   "2 Conductores"
         BeginProperty Font 
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
         TabIndex        =   43
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OptConduc 
         Caption         =   "1 Conductor"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   240
         Width           =   2175
      End
   End
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
      Left            =   7800
      Picture         =   "coalcopruebas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   6000
      Width           =   855
   End
   Begin VB.Frame Frame4 
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
      Height          =   855
      Left            =   1560
      TabIndex        =   33
      Top             =   2400
      Width           =   1455
      Begin VB.ComboBox cmbTipRut 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   1335
      Begin VB.ComboBox cmbTerminal 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   7680
      Width           =   7455
      Begin VB.Label lblTotReg 
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
         Left            =   3120
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Total Registros:"
         BeginProperty Font 
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
         TabIndex        =   28
         Top             =   240
         Width           =   1455
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
      Left            =   7800
      Picture         =   "coalcopruebas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   4440
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
      Left            =   7800
      Picture         =   "coalcopruebas.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   5160
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
      Left            =   7800
      Picture         =   "coalcopruebas.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Limpia la forma actual"
      Top             =   6840
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
      Left            =   7800
      Picture         =   "coalcopruebas.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7680
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Caption         =   "Valores de Busqueda"
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
      TabIndex        =   19
      Top             =   5760
      Width           =   7455
      Begin VB.TextBox txtRuta 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCedula 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtCodEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblRuta 
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
         Left            =   2520
         TabIndex        =   39
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Ruta:"
         BeginProperty Font 
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
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblPlaca 
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
         Left            =   2520
         TabIndex        =   25
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "Placa:"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblNomConductor 
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
         Left            =   2520
         TabIndex        =   23
         Top             =   720
         Width           =   4335
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
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Cedula:"
         BeginProperty Font 
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
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Crit de Busqueda"
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
      TabIndex        =   18
      Top             =   3360
      Width           =   2895
      Begin VB.CheckBox chkRuta 
         Caption         =   "Ruta"
         BeginProperty Font 
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
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkPlaca 
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
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkCedula 
         Caption         =   "Cedula"
         BeginProperty Font 
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
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkEmpresa 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   2895
      Begin VB.OptionButton OptEstado 
         Caption         =   "Todos"
         BeginProperty Font 
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
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Negativo"
         BeginProperty Font 
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
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Positivo"
         BeginProperty Font 
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
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
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
      Top             =   120
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
         TabIndex        =   3
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
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         BeginProperty Font 
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
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
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
   Begin Crystal.CrystalReport rptterminal 
      Left            =   8040
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid grdAlcoPruebas 
      Height          =   5535
      Left            =   3120
      TabIndex        =   26
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   30
      Top             =   8280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   8880
      X2              =   0
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   8520
   End
End
Attribute VB_Name = "coalcopruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalRegistros As Integer
Private Sub Btnbuscar_Click()
    On Error GoTo Errores
    
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
    
    If chkEmpresa.Value = 1 Then
        If txtCodEmp.text = "" Then
            MsgBox "El codigo de la empresa no puede estar vacio", 48, "Advertencia"
            txtCodEmp.SetFocus
            Exit Sub
        End If
    End If
    If chkCedula.Value = 1 Then
        If txtCedula.text = "" Then
            MsgBox "El campo de cedula no puede estar vacio", 48, "Advertencia"
            txtCedula.SetFocus
            Exit Sub
        End If
    End If
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de la placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    If chkRuta.Value = 1 Then
        If txtRuta.text = "" Then
            MsgBox "El campo de ruta no puede estar vacio", 48, "Advertencia"
            txtRuta.SetFocus
            Exit Sub
        End If
    End If
    
    
    'Limpiamos la Grilla
    DIB_GRILLA
    
    'Declaramos variables de consulta
    Fecha = Mskfecini.text
    Fecha1 = Mskfecfin.text
    
    If OptEstado(0).Value = True Then
        opt = "P"
    ElseIf OptEstado(1).Value = True Then
        opt = "N"
    End If
    
    'Hacemos la consulta
    QRY = "SELECT apidprueba, aptasauso, apfechor, applaca, apcedcon, hecodigo, henombre, apresult, apenumer, apestado, " _
            & "apfecreg, apipreg, apnumcaj, apuser, coterminal, cocodrut, rudestino, rutipo " _
            & "FROM alcopruebas, ophojveh, ophojemp, coconduc, opruta " _
            & "WHERE apfecreg BETWEEN TO_DATE('" & Fecha & " 00:00', 'MM/DD/YYYY HH24:mi') " _
            & "AND TO_DATE('" & Fecha1 & " 23:59','MM/DD/YYYY HH24:mi') " _
            & "AND applaca = hvplaca " _
            & "AND hvcodemp = hecodigo " _
            & "AND aptasauso = conumero " _
            & "AND cocodrut = rucodigo "
            
    If chkPlaca.Value = 1 Then
        QRY = QRY & "AND applaca = '" & txtPlaca.text & "' "
    End If
    
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND hecodigo = '" & txtCodEmp.text & "' "
    End If
    
    If chkCedula.Value = 1 Then
        QRY = QRY & "AND apcedcon = '" & txtCedula.text & "' "
    End If
    
    If chkRuta.Value = 1 Then
        QRY = QRY & "AND cocodrut = '" & txtRuta.text & "' "
    End If
    
    If cmbTerminal.ListIndex <> 0 Then
        QRY = QRY & "AND coterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' "
    End If
    
    If cmbTipRut.ListIndex <> 0 Then
        QRY = QRY & "AND rutipo = '" & Mid(cmbTipRut.text, 1, 1) & "' "
    End If
    
    If opt = "P" Then
        QRY = QRY & "AND apresult > 0 "
    ElseIf opt = "N" Then
        QRY = QRY & "AND apresult < 0 "
    End If
    
    If OptConduc(0).Value = True Then
        QRY = QRY & "AND cocedula1 IS NULL "
    ElseIf OptConduc(1).Value = True Then
        QRY = QRY & "AND cocedula1 IS NOT NULL"
    End If
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el Progress Bar
    If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
    End If
            
    TotalRegistros = 0
    
    'Si no es vacio empezamos a agregar los registros a la grilla
    If Not Tbl1.EOF Then
        
        If Tbl1!coterminal = "N" Then
            Terminal = "NORTE"
        ElseIf Tbl1!coterminal = "C" Then
            Terminal = "CENTRAL"
        ElseIf Tbl1!coterminal = "S" Then
            Terminal = "SUR"
        End If
        
        If Tbl1!rutipo = "I" Then
            TRuta = "INTERMUNICIPAL"
        ElseIf Tbl1!rutipo = "S" Then
            TRuta = "SABANERA"
        End If
        
        Do While Not Tbl1.EOF
            entry = CStr(Tbl1!apidprueba) + Chr(9) + CStr(Tbl1!aptasauso) + Chr(9) + Tbl1!applaca + Chr(9) + _
                    Tbl1!apcedcon + Chr(9) + Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!apresult + Chr(9) + _
                    CStr(Tbl1!apenumer) + Chr(9) + Tbl1!apestado + Chr(9) + CStr(Tbl1!apfecreg) + Chr(9) + _
                    Tbl1!apnumcaj + Chr(9) + Terminal + Chr(9) + Tbl1!cocodrut + Chr(9) + Tbl1!rudestino + Chr(9) + TRuta
                    
            grdAlcoPruebas.AddItem (entry)
            Tbl1.MoveNext
            
            TotalRegistros = TotalRegistros + 1
            ProBar.Value = ProBar.Value + 1
        Loop
        Tbl1.Close
        
        lblTotReg.Caption = TotalRegistros
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros ", 48, "Advertencia"
    End If
    
    
Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
    
End Sub

Private Sub BtnExportar_Click()

    On Error GoTo err_mensaje
    
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
    
    If chkEmpresa.Value = 1 Then
        If txtCodEmp.text = "" Then
            MsgBox "El codigo de la empresa no puede estar vacio", 48, "Advertencia"
            txtCodEmp.SetFocus
            Exit Sub
        End If
    End If
    If chkCedula.Value = 1 Then
        If txtCedula.text = "" Then
            MsgBox "El campo de cedula no puede estar vacio", 48, "Advertencia"
            txtCedula.SetFocus
            Exit Sub
        End If
    End If
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de la placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    If chkRuta.Value = 1 Then
        If txtRuta.text = "" Then
            MsgBox "El campo de ruta no puede estar vacio", 48, "Advertencia"
            txtRuta.SetFocus
            Exit Sub
        End If
    End If
    
    
    'Limpiamos la Grilla
    DIB_GRILLA
    
    'Declaramos variables de consulta
    Fecha = Mskfecini.text
    Fecha1 = Mskfecfin.text
    
    If OptEstado(0).Value = True Then
        opt = "P"
    ElseIf OptEstado(1).Value = True Then
        opt = "N"
    End If
    
    'Hacemos la consulta
    QRY = "SELECT apidprueba, aptasauso, apfechor, applaca, apcedcon, hecodigo, henombre, apresult, apenumer, apestado, " _
            & "apfecreg, apipreg, apnumcaj, apuser, coterminal, cocodrut, rudestino, rutipo " _
            & "FROM alcopruebas, ophojveh, ophojemp, coconduc, opruta " _
            & "WHERE apfecreg BETWEEN TO_DATE('" & Fecha & " 00:00', 'MM/DD/YYYY HH24:mi') " _
            & "AND TO_DATE('" & Fecha1 & " 23:59','MM/DD/YYYY HH24:mi') " _
            & "AND applaca = hvplaca " _
            & "AND hvcodemp = hecodigo " _
            & "AND aptasauso = conumero " _
            & "AND cocodrut = rucodigo "
            
    If chkPlaca.Value = 1 Then
        QRY = QRY & "AND applaca = '" & txtPlaca.text & "' "
    End If
    
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND hecodigo = '" & txtCodEmp.text & "' "
    End If
    
    If chkCedula.Value = 1 Then
        QRY = QRY & "AND apcedcon = '" & txtCedula.text & "' "
    End If
    
    If chkRuta.Value = 1 Then
        QRY = QRY & "AND cocodrut = '" & txtRuta.text & "' "
    End If
    
    If cmbTerminal.ListIndex <> 0 Then
        QRY = QRY & "AND coterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' "
    End If
    
    If cmbTipRut.ListIndex <> 0 Then
        QRY = QRY & "AND rutipo = '" & Mid(cmbTipRut.text, 1, 1) & "' "
    End If
    
    If opt = "P" Then
        QRY = QRY & "AND apresult > 0 "
    ElseIf opt = "N" Then
        QRY = QRY & "AND apresult < 0 "
    End If
    
    If OptConduc(0).Value = True Then
        QRY = QRY & "AND cocedula1 IS NULL "
    ElseIf OptConduc(1).Value = True Then
        QRY = QRY & "AND cocedula1 IS NOT NULL "
    End If
    
    QRY = QRY & "ORDER BY apidprueba"
    
    ExportarExcel (QRY)
    
'    If grdAlcoPruebas.rows > 1 Then
'        ExportarExcel
'    Else
'        MsgBox "Debe consultar datos primero", 48, "Advertencia"
'        Exit Sub
'    End If
'
    Exit Sub
err_mensaje:
    MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub ExportarExcel(QRY As String)
    On Error GoTo Errores
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "CONSECUTIVO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "TASA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "PLACA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "CEDULA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "COD. EMPRESA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "NOM. EMPRESA"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "RESULTADO"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "NUMERO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "ESTADO"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "FEC. REGISTRO"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "CAJA"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "TERMINAL"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "COD. RUTA"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "DESTINO"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "TIPO RUTA"
    ApExcel.Cells(fila, 15).Font.Size = 10
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el Progress Bar
    If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
    End If
    
    If Not Tbl1.EOF Then
        
        While Not Tbl1.EOF
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!apidprueba & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!aptasauso & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!applaca & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!apcedcon & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!hecodigo & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!apresult & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!apenumer & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!apestado & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!apfecreg & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & Tbl1!apnumcaj & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!coterminal & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & Tbl1!cocodrut & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & Tbl1!rudestino & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & Tbl1!rutipo & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA" & Month(Mskfecini.text) & Day(Mskfecini.text) & Year(Mskfecini.text)
        ApExcel.Visible = True
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub
    
    'Sebastián Rondón - Enero 02 - 2022
    'Se cambia de donde se saca la información, debido por la magnitud de los datos
    'Aveces no deja consultar y traer datos a la grilla
    '************************************************************
    'Llenamos según la grilla
    'lineas = grdAlcoPruebas.rows - 1
    'grdAlcoPruebas.Row = 1
    'grdAlcoPruebas.Visible = False
    'ProBar.Value = ProBar.Min
    'ProBar.Max = grdAlcoPruebas.rows + 2
'
'    Do
'
'        ReDim dato(1, 15)
'
'        For i = 0 To 14
'            grdAlcoPruebas.Col = i
'            dato(1, i) = Trim(grdAlcoPruebas.Text)
'        Next i
'
'        ApExcel.cells(fila, 1).formula = "" & dato(1, 0) & ""
'        ApExcel.cells(fila, 1).Font.Size = 10
'        ApExcel.cells(fila, 2).formula = "" & dato(1, 1) & ""
'        ApExcel.cells(fila, 2).Font.Size = 10
'        ApExcel.cells(fila, 3).formula = "" & dato(1, 2) & ""
'        ApExcel.cells(fila, 3).Font.Size = 10
'        ApExcel.cells(fila, 4).formula = "" & dato(1, 3) & ""
'        ApExcel.cells(fila, 4).Font.Size = 10
'        ApExcel.cells(fila, 5).formula = "" & dato(1, 4) & ""
'        ApExcel.cells(fila, 5).Font.Size = 10
'        ApExcel.cells(fila, 6).formula = "" & dato(1, 5) & ""
'        ApExcel.cells(fila, 6).Font.Size = 10
'        ApExcel.cells(fila, 7).formula = "" & dato(1, 6) & ""
'        ApExcel.cells(fila, 7).Font.Size = 10
'        ApExcel.cells(fila, 8).formula = "" & dato(1, 7) & ""
'        ApExcel.cells(fila, 8).Font.Size = 10
'        ApExcel.cells(fila, 9).formula = "" & dato(1, 8) & ""
'        ApExcel.cells(fila, 9).Font.Size = 10
'        ApExcel.cells(fila, 10).formula = "" & dato(1, 9) & ""
'        ApExcel.cells(fila, 10).Font.Size = 10
'        ApExcel.cells(fila, 11).formula = "" & dato(1, 10) & ""
'        ApExcel.cells(fila, 11).Font.Size = 10
'        ApExcel.cells(fila, 12).formula = "" & dato(1, 11) & ""
'        ApExcel.cells(fila, 12).Font.Size = 10
'        ApExcel.cells(fila, 13).formula = "" & dato(1, 12) & ""
'        ApExcel.cells(fila, 13).Font.Size = 10
'        ApExcel.cells(fila, 14).formula = "" & dato(1, 13) & ""
'        ApExcel.cells(fila, 14).Font.Size = 10
'        ApExcel.cells(fila, 15).formula = "" & dato(1, 14) & ""
'        ApExcel.cells(fila, 15).Font.Size = 10
'
'        fila = fila + 1
'        lineas = lineas - 1
'        ProBar.Value = ProBar.Value + 1
'        If lineas <> 0 Then
'            grdAlcoPruebas.Row = grdAlcoPruebas.Row + 1
'        End If
'    Loop Until lineas = 0
'***********************************************************************************************
    
Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
    
     On Error GoTo err_mensaje
     
     If grdAlcoPruebas.rows > 1 Then
        
        lineas = grdAlcoPruebas.rows - 1
        grdAlcoPruebas.Row = 1
        grdAlcoPruebas.Visible = False
        ProBar.Value = ProBar.Min
        ProBar.Max = grdAlcoPruebas.rows + 2
        Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
        
        Dbs.BeginTrans
        On Error GoTo Err_Trans
        Do
            
            ReDim dato(1, 15)
                
            For i = 0 To 14
                grdAlcoPruebas.Col = i
                dato(1, i) = Trim(grdAlcoPruebas.text)
            Next i
            
            'Inserta en la tabla temporal
            QRY = "INSERT INTO optempo1 " _
                            & "(temnumdob2, temtext2, temtext3, temtext4, " _
                            & " temtext5, temtext6, temtext7, " _
                            & " temtext8, temtext9, temtext10, temtext11, temtext12, temtext13, tempc ) VALUES (" _
                            & dato(1, 0) & ",'" & dato(1, 1) & "','" & dato(1, 2) & "','" & dato(1, 3) & "','" _
                            & dato(1, 5) & "','" & dato(1, 6) & "', '" _
                            & dato(1, 8) & "','" & dato(1, 9) & "','" & dato(1, 10) & "', '" & dato(1, 11) & "','" _
                            & dato(1, 12) & "','" & dato(1, 13) & "','" & dato(1, 14) & "','" & Pc & "') "
                            
            Dbs.Execute QRY
            
            lineas = lineas - 1
            ProBar.Value = ProBar.Value + 1
            If lineas <> 0 Then
                grdAlcoPruebas.Row = grdAlcoPruebas.Row + 1
            End If
            
        Loop Until lineas = 0
        
        Dbs.CommitTrans
         grdAlcoPruebas.Visible = True
         On Error GoTo err_mensaje
          respuesta = MsgBox("Desea escoger la impresora (SI) o impresora predeterminada (NO)", vbYesNo + 32, "Impresoras")
        
         If respuesta = 6 Then
            imprimir.Show 1
            If G_Imprimir = "0" Then
               Exit Sub
            End If
         End If
         
         reporte = RpPath + "\coalcopruebas.rpt"
             
             rptterminal.ReportFileName = reporte
             rptterminal.Connect = G_Conexion_reportes
             rptterminal.Formulas(1) = "fecini = '" & Mskfecini.FormattedText & "'"
             rptterminal.Formulas(2) = "fecfin = '" & Mskfecfin.FormattedText & "'"
             rptterminal.Formulas(3) = "empresa = '" & G_NomEmpresa & "'"
             rptterminal.Formulas(4) = "titulo = '" & TITULO & "'"
             rptterminal.Formulas(5) = "totalregistros = '" & TotalRegistros & "'"
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
             coalcopruebas.Enabled = True
            
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
    chkEmpresa.Value = False
    chkCedula.Value = False
    chkPlaca.Value = False
    chkRuta.Value = False
    txtCodEmp.text = ""
    lblNomEmpresa.Caption = ""
    txtCedula.text = ""
    lblNomConductor.Caption = ""
    txtPlaca.text = ""
    lblPlaca.Caption = ""
    txtRuta.text = ""
    lblRuta.Caption = ""
    OptEstado(2).Value = True
    OptConduc(0).Value = False
    OptConduc(1).Value = False
    OptConduc(2).Value = True
    cmbTerminal.ListIndex = 0
    cmbTipRut.ListIndex = 0
    lblTotReg.Caption = ""
    DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    'Creamos la grilla
    DIB_GRILLA
    
    'Ponemos el foco en la fecha
    Mskfecini_GotFocus
    
    'Llenamos el combo de Terminal
    QRY = "SELECT dsdes " _
            & "FROM gesuptip, gedetsuptip " _
            & "WHERE dscodtip = stcodtip " _
            & "AND stdes = 'TIPOS DE TERMINAL' " _
            & "and dsest = 'A' "
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    cmbTerminal.Clear
    cmbTerminal.AddItem "GENERAL"
    Do While Not Tbl1.EOF
        cmbTerminal.AddItem Tbl1(0)
        Tbl1.MoveNext
    Loop
    Tbl1.Close
    cmbTerminal.ListIndex = 0
    
    'Llenamos el combo de Tipo de Ruta
    cmbTipRut.AddItem "GENERAL"
    cmbTipRut.AddItem "INTERMUN."
    cmbTipRut.AddItem "SABANERA"
    cmbTipRut.ListIndex = 0
    
    
    OptEstado(2).Value = True
    
    OptConduc(0).Value = False
    OptConduc(1).Value = False
    OptConduc(2).Value = True
    

End Sub

Private Function DIB_GRILLA()

   grdAlcoPruebas.rows = 1
   grdAlcoPruebas.Cols = 15
   grdAlcoPruebas.Clear
   
   s$ = "Consecutivo |Num. Tasa   |Placa |Cedula |Cod. Empresa |Empresa |Resultado |Numero |Estado |Fec. Registro |Caja |Terminal |Cod. Ruta |Destino |Tipo Ruta  "
   grdAlcoPruebas.FormatString = s$
   grdAlcoPruebas.Row = 0
   grdAlcoPruebas.Col = 0
   grdAlcoPruebas.ColWidth(0) = 1100
   grdAlcoPruebas.ColWidth(1) = 1600
   grdAlcoPruebas.ColWidth(2) = 1200
   grdAlcoPruebas.ColWidth(3) = 1400
   grdAlcoPruebas.ColWidth(4) = 1200
   grdAlcoPruebas.ColWidth(5) = 1600
   grdAlcoPruebas.ColWidth(6) = 1200
   grdAlcoPruebas.ColWidth(7) = 1600
   grdAlcoPruebas.ColWidth(8) = 1200
   grdAlcoPruebas.ColWidth(9) = 2100
   grdAlcoPruebas.ColWidth(10) = 1200
   grdAlcoPruebas.ColWidth(11) = 1200
   grdAlcoPruebas.ColWidth(12) = 1200
   grdAlcoPruebas.ColWidth(13) = 1800
   grdAlcoPruebas.ColWidth(14) = 1500
   
   
   For i = 0 To grdAlcoPruebas.Cols - 1
       grdAlcoPruebas.Col = i
       grdAlcoPruebas.CellFontBold = 1
   Next i
End Function


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
Private Sub chkCedula_Click()
    
    If chkCedula.Value = 1 Then
        txtCedula.Enabled = True
    Else
        txtCedula.text = ""
        lblNomConductor.Caption = ""
        txtCedula.Enabled = False
    End If
    
End Sub

Private Sub chkEmpresa_Click()
    
    If chkEmpresa.Value = 1 Then
        txtCodEmp.Enabled = True
    Else
        txtCodEmp.text = ""
        lblNomEmpresa.Caption = ""
        txtCodEmp.Enabled = False
    End If
    
End Sub
Private Sub chkPlaca_Click()

    If chkPlaca.Value = 1 Then
        txtPlaca.Enabled = True
    Else
        txtPlaca.text = ""
        lblPlaca.Caption = ""
        txtPlaca.Enabled = False
    End If

End Sub

Private Sub chkRuta_Click()
    If chkRuta.Value = 1 Then
        txtRuta.Enabled = True
    Else
       txtRuta.text = ""
       lblRuta.Caption = ""
       txtPlaca.Enabled = False
    End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "70"
  lblNomEmpresa.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub
Private Sub txtcedula_DblClick()
    forma = "20"
    forma1 = "9"
    lblNomConductor.Caption = ""
    buscar.Show 1
    txtcedula_LostFocus
End Sub

Private Sub TxtPlaca_DblClick()
    forma = "2"
    forma1 = "25"
    lblPlaca.Caption = ""
    buscar.Show 1
    TxtPlaca_LostFocus
End Sub

Private Sub Txtruta_DblClick()
    forma = "3"
    If txtCodEmp.text <> "" Then
        forma1 = "34"
        G_Codigo_Empresa = txtCodEmp.text
    Else
        forma1 = "33"
    End If
    lblRuta.Caption = ""
    buscar.Show 1
    Txtruta_LostFocus
End Sub


Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_LostFocus()
    
    If txtCodEmp.text <> "" Then
        
        'Si ya digitaron la cedula validamos que el conductor pertenezca a la empresa
        If txtCedula.text <> "" Then
            QRY = "SELECT cocedula " _
                    & "FROM coconductor " _
                    & "WHERE cocedula = '" & txtCedula.text & "' " _
                    & "AND cocodemp = '" & txtCodEmp.text & "' "
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                MsgBox "El conductor no pertenece a la empresa", "48", "Error"
                lblNomEmpresa.Caption = ""
                txtCodEmp.text = ""
                Txtcodemp_GotFocus
                Exit Sub
            End If
            Tbl1.Close
        End If
        
        'Si ya fue digitada la ruta, verificamos que pertenezca a la empresa
        If txtRuta.text <> "" Then
            QRY = "SELECT * " _
                    & "FROM opruthor " _
                    & "WHERE rhcodemp = '" & txtCodEmp.text & "' " _
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
            & "WHERE hecodigo = '" & txtCodEmp.text & "' "
            
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
            lblNomEmpresa.Caption = Tbl1!henombre
            Tbl1.Close
        Else
            
            QRY = "SELECT henombre FROM ophojempcon " _
                & "WHERE hecodigo = '" & txtCodEmp.text & "' "
                
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Tbl2.EOF = False Then
                lblNomEmpresa.Caption = Tbl2!henombre
            Else
                Tbl1.Close
                MsgBox "Empresa no existente", 48, "Advertencia"
                txtCodEmp.text = ""
                txtCodEmp_DblClick
            End If
            Tbl2.Close
        End If
    End If
End Sub

Private Sub TxtPlaca_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   longitud = Len(txtPlaca.text)
   longitud = longitud + 1
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
          If longitud = 7 Then
             If KeyAscii% <> 8 Then
                TxtPlaca_LostFocus
             End If
          End If
          Exit Sub
       Else
          If longitud >= 7 Then
             txtPlaca.text = ""
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
End Sub

Private Sub TxtPlaca_LostFocus()
    If txtPlaca.text <> "" Then
        
        If txtCodEmp.text <> "" Then
            QRY = "SELECT hvplaca " _
                    & "FROM ophojveh " _
                    & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                    & "AND hvcodemp = '" & txtCodEmp.text & "' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                MsgBox "Vehículo no pertenece a la empresa ", 48, "Advertencia"
                txtPlaca.text = ""
                lblPlaca.Caption = ""
                Exit Sub
            End If
            Tbl1.Close
        End If
            
        QRY = "SELECT hvplaca, hvpropie " _
                & "FROM ophojveh " _
                & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                & "AND hvestado = 'Activo' "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            lblPlaca.Caption = "Propietario: " & Tbl1!hvpropie
        Else
            txtPlaca.text = ""
            lblPlaca.Caption = ""
            MsgBox "El vehículo no existe o se encuentra inactivo ", 48, "Advertencia"
        End If
    End If
End Sub
Private Sub Txtruta_GotFocus()
    Senalar_Texto Me
End Sub
Private Sub Txtruta_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub Txtruta_LostFocus()
    If txtRuta.text <> "" Then
        If txtCodEmp.text <> "" Then
            QRY = "SELECT * " _
                    & "FROM opruthor " _
                    & "WHERE rhcodemp = '" & txtCodEmp.text & "' " _
                    & "AND rhcodrut = '" & txtRuta.text & "' "
                    
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                Tbl1.Close
                txtRuta.SetFocus
                MsgBox "La empresa no tiene autorizada esta ruta", 48, "Advertencia"
                txtRuta.text = ""
                Exit Sub
            Else
                If Tbl1!rhestado = "Inactiva" Then
                    Tbl1.Close
                    txtRuta.text = ""
                    MsgBox "La ruta esta inactiva para esta empresa", 48, "Advertencia"
                    Exit Sub
                End If
            End If
            Tbl1.Close
        End If
        
        QRY = "SELECT rucodigo, rudestino " _
                & "FROM opruta " _
                & "WHERE rucodigo = '" & txtRuta.text & "' "
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF Then
            MsgBox "La ruta ingresada no existe", 48, "Advertencia"
            txtRuta.text = ""
            txtRuta.SetFocus
            Tbl1.Close
        Else
            lblRuta.Caption = Tbl1!rudestino
        End If
        Tbl1.Close
    End If
End Sub

Private Sub txtcedula_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcedula_KeyPress(KeyAscii As Integer)
    If (KeyAscii% < 48 Or KeyAscii% > 57) And (KeyAscii% <> 8) Or (KeyAscii% = 47) Then
      KeyAscii% = 0
  End If
End Sub


Private Sub txtcedula_LostFocus()

    If txtCedula.text <> "" Then
        
        If txtCodEmp.text <> "" Then
            QRY = "SELECT cocedula " _
                    & "FROM coconductor " _
                    & "WHERE cocedula = '" & txtCedula.text & "' " _
                    & "AND cocodemp = '" & txtCodEmp.text & "'"
            Set Tbl1 = Dbs.Execute(QRY)
            
            If Tbl1.EOF Then
                MsgBox "El conductor no pertenece a la empresa", "48", "Advertencia"
                lblNomConductor.Caption = ""
                txtCedula.text = ""
                txtcedula_GotFocus
                Exit Sub
            End If
            Tbl1.Close
        End If
        
        QRY = "SELECT cocedula, conombres, coapellidos " _
                & "FROM coconductor " _
                & "WHERE cocedula = '" & txtCedula.text & "' "
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            lblNomConductor.Caption = Tbl1!conombres & " " & Tbl1!coapellidos
        Else
            lblNomConductor.Caption = ""
            txtCedula.text = ""
            MsgBox "El conductor no existe", 48, "Advertencia"
            txtcedula_GotFocus
        End If
        Tbl1.Close
        
    End If
End Sub
