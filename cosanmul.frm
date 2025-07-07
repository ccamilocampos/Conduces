VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form cosanmul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Sanciones"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7440
      Picture         =   "cosanmul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   4080
      Width           =   855
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
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   7095
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         ToolTipText     =   "Fecha incial donde quiere que empieze a buscar"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   285
         Left            =   5280
         TabIndex        =   1
         ToolTipText     =   "Fecha final hasta donde quiere que busque"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         Left            =   3720
         TabIndex        =   34
         Top             =   360
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
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   7095
      Begin Crystal.CrystalReport rptterminal 
         Left            =   600
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconsalmul 
         Height          =   2535
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores De Búsqueda:"
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
      Left            =   2400
      TabIndex        =   24
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox Txtcodinf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   9
         ToolTipText     =   "Código de infracción (doble Click para ayuda)"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Txtnumpla 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   8
         ToolTipText     =   "Número de placa del vehículo(doble Click para ayuda)"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Cmbestado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "cosanmul.frx":030A
         Left            =   1440
         List            =   "cosanmul.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Estado de la multa por la que quiere que busque"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Código de la empresa(doble Click para ayuda) "
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblCodInf 
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
         Left            =   2280
         TabIndex        =   36
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Infrac.:"
         BeginProperty Font 
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
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Lblnumpla 
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
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label8 
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
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
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
         Left            =   2280
         TabIndex        =   27
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Estado: "
         BeginProperty Font 
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
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frmbuscar 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   2175
      Begin VB.CheckBox Chkbuscar 
         Caption         =   "Cód. Infracción"
         BeginProperty Font 
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Chkbuscar 
         Caption         =   "Núm. Placa"
         BeginProperty Font 
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Chkbuscar 
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Chkbuscar 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1695
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
      Left            =   7440
      Picture         =   "cosanmul.frx":0351
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   3240
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
      Left            =   7440
      Picture         =   "cosanmul.frx":0793
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir de la forma actual"
      Top             =   6600
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
      Left            =   7440
      Picture         =   "cosanmul.frx":0BD5
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpia la forma actual"
      Top             =   5760
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
      Left            =   7440
      Picture         =   "cosanmul.frx":1017
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   7095
      Begin VB.Label Label12 
         Caption         =   "Anuladas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lblanuladas 
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
         Height          =   300
         Left            =   5160
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Lbleducativas 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Educativas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lblpendientes 
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
         Height          =   300
         Left            =   5160
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Pendientes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lblpagadas 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Pagadas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   8400
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   7320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   7320
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   7320
      X2              =   7320
      Y1              =   7440
      Y2              =   0
   End
End
Attribute VB_Name = "cosanmul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Integer
Dim dato1 As Double
Dim longitud
Private Sub BtnBuscar_Click()
  
   On Error GoTo terminar
  
   bandera = 0
   Lblpagadas.Caption = ""
   Lblpendientes.Caption = ""
   Lbleducativas.Caption = ""
   Lblanuladas.Caption = ""
   dato1 = 0
   dato2 = 0
   dato3 = 0
   dato4 = 0
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
   
   If Chkbuscar(0).Value = 1 Then
      If Cmbestado.text = "" Then
         MsgBox "Falta Ingresar El Estado", 48, "Advertencia"
         Cmbestado.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkbuscar(1).Value = 1 Then
      If Txtcodemp.text = "" Then
         MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkbuscar(2).Value = 1 Then
      If Txtnumpla.text = "" Then
         MsgBox "Falta Ingresar El Numero De La Placa", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
   If Chkbuscar(3).Value = 1 Then
      If Txtcodinf.text = "" Then
         MsgBox "Falta ingresar el código de infracción", 48, "Advertencia"
         Txtcodinf.SetFocus
         Exit Sub
      End If
   End If
   DIB_GRILLA
   QRY = "SELECT munumero, mufecha, muhora, musitio, mucodemp, munumpla, hvnumint,  " _
               & "mudestino, rudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete,  " _
               & "mucodinf, muauxcon, muauxalt,  muestado, mupasvia, mupaspla  " _
         & "FROM opmultas, ophojveh, opruta " _
        & "WHERE mufecha BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
          & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
          & "AND munumpla = hvplaca " _
          & "AND mudestino = rucodigo "
   
   
   '******************** CUANDO BUSCA POR ESTADO ********************************
   If Chkbuscar(0).Value = 1 Then
      QRY = QRY & "AND muestado LIKE '" & Cmbestado.text & "%' "
   End If
   
   '******************** CUANDO BUSCA POR EMPRESA *******************************
   If Chkbuscar(1).Value = 1 Then
      QRY = QRY & "AND mucodemp = '" & Txtcodemp.text & "' "
   End If
   
   '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
   If Chkbuscar(2).Value = 1 Then
      QRY = QRY & "AND munumpla = '" & Txtnumpla.text & "' "
   End If
   '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
   If Chkbuscar(2).Value = 1 Then
      QRY = QRY & "AND mucodinf = '" & Txtcodinf.text & "' "
   End If
   
   
   QRY = QRY & "ORDER BY munumero ASC "
   
   
   Set Tbl1 = Dbs.Execute(QRY)
   cosanmul.Enabled = False
   Do While Tbl1.EOF = False
      DoEvents
      
      QRY = "SELECT henombre "
      If Len(Tbl1!mucodemp) = 4 Then
         QRY = QRY & "FROM ophojempcon  "
      Else
         QRY = QRY & "FROM ophojemp  "
      End If
      QRY = QRY & "WHERE hecodigo = '" & Tbl1!mucodemp & "'"
          
      Set Tbl2 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         sEmpresa = Tbl2!henombre
      End If
      Tbl2.Close
      
      sdiferencia = Tbl1(10) - Tbl1(13)
      
      
      QRY = "SELECT rudestino FROM opruta WHERE rucodigo = '" & Tbl1(12) & "'"
      
      Set Ftb01 = Dbs.Execute(QRY)
      
      If Not Ftb01.EOF Then
         sDesPla = Ftb01(0)
      Else
         sDesPla = ""
      End If
      Ftb01.Close
      
      If IsNull(Tbl1(18)) Then
         sAuxAlt = ""
      Else
         sAuxAlt = Tbl1(18)
      End If
      
      sDifRec = Tbl1(10) - Tbl1(13)
      entry = CStr(Tbl1(0)) + Chr(9) + CStr(Tbl1(1)) + Chr(9) + CStr(Format(Tbl1(2), "HH:mm")) + Chr(9) + Tbl1(3) + Chr(9) + _
              Tbl1(4) & "-" & sEmpresa + Chr(9) + Tbl1(5) + Chr(9) + CStr(Tbl1(6)) + Chr(9) + _
              Tbl1(7) & "-" & Tbl1(8) + Chr(9) + CStr(Tbl1(9)) + Chr(9) + CStr(Tbl1(10)) + Chr(9) + _
              Tbl1(11) & "-" & sDesPla + Chr(9) + CStr(Tbl1(12)) + Chr(9) + Tbl1(13) + Chr(9) + Tbl1(14) + Chr(9) + CStr(sDifRec) + Chr(9) + _
              CStr(Tbl1(15)) + Chr(9) + Tbl1(16) + Chr(9) + sAuxAlt + Chr(9) + Tbl1(18)
              
      If Tbl1(19) = "Pagada" Then
         dato1 = dato1 + 1
      ElseIf Tbl1(19) = "Educativo" Then
         dato2 = dato2 + 1
      ElseIf Tbl1(19) = "Anulado" Then
         dato3 = dato3 + 1
      ElseIf Tbl1(19) = "Pendiente" Then
         dato4 = dato4 + 1
      End If
              
      Grdconsalmul.AddItem entry
      Tbl1.MoveNext
      bandera = 1
   Loop
   Tbl1.Close
   cosanmul.Enabled = True
   Lblpagadas.Caption = dato1
   Lbleducativas.Caption = dato2
   Lblanuladas.Caption = dato3
   Lblpendientes.Caption = dato4
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 16, "Busqueda"
      Exit Sub
   End If
   Call Aplicar_Perfil(Me, "cosanmul")
   Exit Sub
  
  
terminar:
  cosanmul.Enabled = True
  MsgBox "Error al tratar de mostrar los registros o error de la base de datos " & Chr(13) + Chr(13) & "ERROR >> " & Err.Description, 48, "Advertencia"
 
  Exit Sub
Resume
End Sub

Private Sub BtnExportar_Click()
    On Error GoTo terminar
    
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
    
    If Chkbuscar(0).Value = 1 Then
       If Cmbestado.text = "" Then
          MsgBox "Falta Ingresar El Estado", 48, "Advertencia"
          Cmbestado.SetFocus
          Exit Sub
       End If
    End If
    
    If Chkbuscar(1).Value = 1 Then
       If Txtcodemp.text = "" Then
          MsgBox "Falta Ingresar El Codigo De La Empresa", 48, "Advertencia"
          Txtcodemp.SetFocus
          Exit Sub
       End If
    End If
    
    If Chkbuscar(2).Value = 1 Then
       If Txtnumpla.text = "" Then
          MsgBox "Falta Ingresar El Numero De La Placa", 48, "Advertencia"
          Txtnumpla.SetFocus
          Exit Sub
       End If
    End If
    If Chkbuscar(3).Value = 1 Then
       If Txtcodinf.text = "" Then
          MsgBox "Falta ingresar el código de infracción", 48, "Advertencia"
          Txtcodinf.SetFocus
          Exit Sub
       End If
    End If
    DIB_GRILLA
    QRY = "SELECT munumero, mufecha, muhora, musitio, mucodemp, munumpla, hvnumint,  " _
                & "mudestino, rudestino, muvalcon, muconduce, mudespla, muvalpla, muplanilla, mutiquete,  " _
                & "mucodinf, muauxcon, muauxalt,  muestado, mupasvia, mupaspla  " _
          & "FROM opmultas, ophojveh, opruta " _
         & "WHERE mufecha BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
           & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
           & "AND munumpla = hvplaca " _
           & "AND mudestino = rucodigo "
    
    
    '******************** CUANDO BUSCA POR ESTADO ********************************
    If Chkbuscar(0).Value = 1 Then
       QRY = QRY & "AND muestado LIKE '" & Cmbestado.text & "%' "
    End If
    
    '******************** CUANDO BUSCA POR EMPRESA *******************************
    If Chkbuscar(1).Value = 1 Then
       QRY = QRY & "AND mucodemp = '" & Txtcodemp.text & "' "
    End If
    
    '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
    If Chkbuscar(2).Value = 1 Then
       QRY = QRY & "AND munumpla = '" & Txtnumpla.text & "' "
    End If
    '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
    If Chkbuscar(2).Value = 1 Then
       QRY = QRY & "AND mucodinf = '" & Txtcodinf.text & "' "
    End If
    
    QRY = QRY & "ORDER BY munumero ASC "
    
    cosanmul.Enabled = False
    ExportarExcel (QRY)
    cosanmul.Enabled = True
    
    Call Aplicar_Perfil(Me, "cosanmul")
    Exit Sub
terminar:
  cosanmul.Enabled = True
  MsgBox "Error al tratar de mostrar los registros o error de la base de datos " & Chr(13) + Chr(13) & "ERROR >> " & Err.Description, 48, "Advertencia"
 
  Exit Sub
Resume
End Sub

Private Sub ExportarExcel(ByVal QRY As String)
    On Error GoTo Errores
    
    Lblpagadas.Caption = ""
    Lblpendientes.Caption = ""
    Lbleducativas.Caption = ""
    Lblanuladas.Caption = ""
    dato1 = 0
    dato2 = 0
    dato3 = 0
    dato4 = 0
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "CONSECUTIVO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "FECHA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "HORA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "LUGAR"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "EMPRESA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "PLACA"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "NUM. INTERNO"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "DESTINO TUT"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "VR. DESTINO"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "NUM. TUT"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "DES PLANILLA"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "VR. DESTINO PLANILLA"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "NUM. PLANILLA"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "NUM. TIQUETE"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "DIFERENCIA REC."
    ApExcel.Cells(fila, 15).Font.Size = 10
    ApExcel.Cells(fila, 16).formula = "COD. INFRACCION"
    ApExcel.Cells(fila, 16).Font.Size = 10
    ApExcel.Cells(fila, 17).formula = "AUXILIAR 1"
    ApExcel.Cells(fila, 17).Font.Size = 10
    ApExcel.Cells(fila, 18).formula = "AUXILIAR 2"
    ApExcel.Cells(fila, 18).Font.Size = 10
    ApExcel.Cells(fila, 19).formula = "ESTADO"
    ApExcel.Cells(fila, 19).Font.Size = 10
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        While Not Tbl1.EOF
            'Consultamos los datos de la empresa
            '*********************************************
            QRY = "SELECT henombre "
            If Len(Tbl1!mucodemp) = 4 Then
                QRY = QRY & "FROM ophojempcon "
            Else
                QRY = QRY & "FROM ophojemp "
            End If
            QRY = QRY & "WHERE hecodigo = '" & Tbl1!mucodemp & "' "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                sEmpresa = Tbl2!henombre
            End If
            '*********************************************
            
            'Destino planilla
            '*********************************************
            If IsNull(Tbl1!mudespla) Then
                sDesPla = ""
            Else
                QRY = "SELECT rudestino " _
                        & "FROM opruta " _
                        & "WHERE rucodigo = '" & Tbl1!mudespla & "' "
                        
                Set Tbl2 = Dbs.Execute(QRY)
                
                If Not Tbl2.EOF Then
                    sDesPla = Tbl1!mudespla & "-" & Tbl2!rudestino
                Else
                    sDesPla = Tbl1!mudespla
                End If
            End If
            '*********************************************
            
            'Diferencia
            '*********************************************
            sdiferencia = Val(Tbl1!muvalcon) - Val(Tbl1!muvalpla)
            '*********************************************
            
            'Auxiliar Control
            '*********************************************
            If IsNull(Tbl1!muauxalt) Then
                sAuxAlt = ""
            Else
                sAuxAlt = Tbl1!muauxalt
            End If
            '*********************************************
            
            'Nombre infracción
            '*********************************************
            QRY = "SELECT psdes FROM geparsis WHERE pscod = '" & Tbl1!mucodinf & "' "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sDesInfr = Tbl2!psdes
            Else
                sDesInfr = ""
            End If
            '*********************************************
            
            'Llenamos los datos en Excel
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!munumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!mufecha & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!muhora & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!musitio & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & sEmpresa & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!munumpla & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!hvnumint & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!mudestino & " - " & Tbl1!rudestino & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!muvalcon & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!muconduce & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & Tbl1!mudespla & " - " & sDesPla & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!muvalpla & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & Tbl1!muplanilla & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & Tbl1!mutiquete & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & sdiferencia & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            ApExcel.Cells(fila, 16).formula = "" & sDesInfr & ""
            ApExcel.Cells(fila, 16).Font.Size = 10
            ApExcel.Cells(fila, 17).formula = "" & Tbl1!muauxcon & ""
            ApExcel.Cells(fila, 17).Font.Size = 10
            ApExcel.Cells(fila, 18).formula = "" & sAuxAlt & ""
            ApExcel.Cells(fila, 18).Font.Size = 10
            ApExcel.Cells(fila, 19).formula = "" & Tbl1!muestado & ""
            ApExcel.Cells(fila, 19).Font.Size = 10
            
            If Tbl1!muestado = "Pagada" Then
                dato1 = dato1 + 1
            ElseIf Tbl1!muestado = "Educativo" Then
                dato2 = dato2 + 1
            ElseIf Tbl1!muestado = "Anulado" Then
                dato3 = dato3 + 1
            ElseIf Tbl1!muestado = "Pendiente" Then
                dato4 = dato4 + 1
            End If
            
            fila = fila + 1
            Tbl1.MoveNext
        Wend
        
        Tbl2.Close
        Tbl1.Close
        
        Lblpagadas.Caption = dato1
        Lbleducativas.Caption = dato2
        Lblanuladas.Caption = dato3
        Lblpendientes.Caption = dato4
        
        ApExcel.Worksheets(1).Name = "CONSULTA MULTAS"
        ApExcel.Visible = True
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub
    
Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
On Error GoTo err_mensaje
  
   If Grdconsalmul.rows > 1 Then
      lineas = Grdconsalmul.rows - 1
      Grdconsalmul.Row = 1
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      cosanmul.Enabled = False
      Dbs.BeginTrans
         
      On Error GoTo Err_Trans
      Do
      
         ReDim dato(1, 18)
                
         For i = 0 To 17
            Grdconsalmul.Col = i
            dato(1, i) = Trim(Grdconsalmul.text)
         Next i
            
         ' insertamos en la tabla
             
         Dbs.Execute "INSERT INTO optempo1 " _
         & "(temnumdob1, temtext1, temtext2, temtext3, " _
         & " temtext4, temtext5, temnumdob2, temtext6, " _
         & " temnumdob3, temnumdob4, temtext7, temnumdob5, " _
         & " temtext8, temtext9, temnummon1, temtext10, " _
         & " temtext11,  temtext12, tempc ) VALUES (" _
         & dato(1, 0) & " ,'" & dato(1, 1) & "','" & dato(1, 2) & "','" & dato(1, 3) & "','" _
         & dato(1, 4) & "','" & dato(1, 5) & "' , " & dato(1, 6) & ",'" & dato(1, 7) & "'," _
         & dato(1, 8) & "," & dato(1, 9) & ",'" & dato(1, 10) & "'," & dato(1, 11) & ",'" _
         & dato(1, 12) & "','" & dato(1, 13) & "'," & dato(1, 14) & ",'" & dato(1, 15) & "','" _
         & dato(1, 16) & "','" & dato(1, 17) & "','" & Pc & "')"
             
         ' fin insertar
              
             
         lineas = lineas - 1
             
         If lineas <> 0 Then
            Grdconsalmul.Row = Grdconsalmul.Row + 1
         End If
      Loop Until lineas = 0  'do_while que se mueve x la grilla
    
     
      inicio1 = Mskfecini.FormattedText
      final = Mskfecfin.FormattedText
     
     
      Dbs.CommitTrans
      On Error GoTo err_mensaje

      reporte = RpPath + "\cosalmul.rpt"
     
      rptterminal.ReportFileName = reporte
      rptterminal.Connect = G_Conexion_reportes
      rptterminal.Formulas(0) = "totpag = '" & Lblpagadas.Caption & "'"
      rptterminal.Formulas(1) = "inicio1 = '" & inicio1 & "'"
      rptterminal.Formulas(2) = "final = '" & final & "'"
      rptterminal.Formulas(3) = "empresa = '" & G_NomEmpresa & "'"
      rptterminal.Formulas(4) = "totedu = '" & Lbleducativas.Caption & "'"
      rptterminal.Formulas(5) = "totpen = '" & Lblpendientes.Caption & "'"
      rptterminal.Formulas(6) = "totanu = '" & Lblanuladas.Caption & "'"
      
      rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
      rptterminal.WindowState = crptMaximized
      rptterminal.Destination = 0
      rptterminal.Action = 1
      cosanmul.Enabled = True
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   End If
  
   Exit Sub
Err_Trans:
   Dbs.RollbackTrans
   
err_mensaje:
   cosanmul.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnlimpiar_Click()
  Mskfecini.Mask = ""
  Mskfecini.text = ""
  Mskfecini.Mask = "##/##/####"
  Mskfecfin.Mask = ""
  Mskfecfin.text = ""
  Mskfecfin.Mask = "##/##/####"
  
  For i = 0 To 3
   Chkbuscar(i).Value = 0
  Next i
  Txtcodemp.text = ""
  Lblcodemp.Caption = ""
  Txtnumpla.text = ""
  Lblnumpla.Caption = ""
  Txtcodinf.text = ""
  LblCodInf.Caption = ""
  Grdconsalmul.rows = 1
  Grdconsalmul.Cols = 2
  Grdconsalmul.Clear
  Lblpagadas.Caption = ""
  Lblpendientes.Caption = ""
  Lbleducativas.Caption = ""
  Lblanuladas.Caption = ""
  
  dato1 = 0
  DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Btntotalizar_Click()
  If Grdconsalmul.rows > 1 Then
     dato1 = 0
     Grdconsalmul.Col = 12
     For i = 1 To Grdconsalmul.rows - 1
         Grdconsalmul.Row = i
         dato1 = dato1 + CDbl(Grdconsalmul.text)
     Next i
     MskTotal.text = dato1
  End If
End Sub

Private Sub Chkbuscar_Click(index As Integer)
  If Chkbuscar(0).Value = 1 Then
     Cmbestado.Enabled = True
  Else
     Cmbestado.Enabled = False
  End If
  
  If Chkbuscar(1).Value = 1 Then
     Txtcodemp.Enabled = True
     Chkbuscar(2).Enabled = True
  Else
     Txtcodemp.Enabled = False
     Txtcodemp.text = ""
     Lblcodemp.Caption = ""
  End If
  
  If Chkbuscar(2).Value = 1 Then
     Txtnumpla.Enabled = True
  Else
     Txtnumpla.Enabled = False
     Txtnumpla.text = ""
     Lblnumpla.Caption = ""
  End If
  If Chkbuscar(3).Value = 1 Then
     Txtcodinf.Enabled = True
  Else
     Txtcodinf.Enabled = False
     Txtcodinf.text = ""
     LblCodInf.Caption = ""
  End If
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "cosanmul")
End Sub

Private Sub Form_Load()
   Optbuscar_Click (index)
   DIB_GRILLA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
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
        Exit Sub
     End If
  
     If Mskfecfin.ClipText <> "" Then
        If Not IsDate(Mskfecfin.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
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
        Exit Sub
     End If
  
     If Mskfecini.ClipText <> "" Then
        If Not IsDate(Mskfecini.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Exit Sub
        End If
     End If
  End If
End Sub

Private Sub Optbuscar_Click(index As Integer)
  
End Sub

Private Sub txtCodEmp_DblClick()
  If Chkbuscar(1).Value = 1 Then
     forma = "16"
     forma1 = "41"
     Txtcodemp.text = ""
     buscar.Show 1
     txtcodemp_LostFocus
  End If
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   If Chkbuscar(1).Value = 1 And Txtcodemp.text <> "" Then
      QRY = "SELECT henombre "
      If Len(Txtcodemp.text) = 4 Then
         QRY = QRY & "FROM ophojempcon  "
      Else
         QRY = QRY & "FROM ophojemp  "
      End If
      QRY = QRY & "WHERE hecodigo = '" & Txtcodemp.text & "'"
          
      Set Ftb01 = Dbs.Execute(QRY)
      
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         MsgBox "El Codigo De Empresa Que Digito No Existe", 48, "Advertencia"
         Ftb01.Close
         forma = "16"
         forma1 = "41"
         buscar.Show 1
         Txtcodemp.SetFocus
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

Private Sub Txtcodinf_DblClick()
   forma = "4"
   forma1 = "3"
   buscar.Show 1
   Txtcodinf_LostFocus
End Sub

Private Sub Txtcodinf_KeyPress(KeyAscii As Integer)
   If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub Txtcodinf_LostFocus()
      
   If Trim(Txtcodinf.text) <> "" Then
      fqry01 = "SELECT * " _
               & "FROM geparsis " _
              & "WHERE pscod = '" & Txtcodinf.text & "' "
      
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         LblCodInf.Caption = Ftb01!psdes
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Codigo De Infraccion No Existe", 48, "Advertencia"
         Txtcodinf.SetFocus
         Exit Sub
         
      End If
   End If
End Sub

Private Sub Txtnumpla_DblClick()
  If Chkbuscar(2).Value = 1 Then
     forma = "2"
     forma1 = "14"
     Txtnumpla.text = ""
     buscar.Show 1
     Txtnumpla_LostFocus
  End If
End Sub

Private Sub Txtnumpla_KeyPress(KeyAscii As Integer)
  'Char = Chr(KeyAscii%)
  'KeyAscii = Asc(UCase(Char))  ' mayuscula
  'longitud = Len(Txtnumpla.Text)
  'longitud = longitud + 1
  'If longitud <= 3 Then
  '   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
  '       Exit Sub
  '   Else
  '       MsgBox "En El Numero De Placa las Tres Primeras Deben Ser Letras", 48, "Advertencia"
  '       KeyAscii% = 0
  '       Exit Sub
  '   End If
  'ElseIf longitud > 3 Then
  '   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
  '      Exit Sub
  '   Else
  '      MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
  '      KeyAscii% = 0
  '      Exit Sub
  '   End If
  'End If
  Char = Chr(KeyAscii%)
     KeyAscii = Asc(UCase(Char))  ' mayuscula
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
End Sub

Private Sub Txtnumpla_LostFocus()
  If Chkbuscar(2).Value = 1 And Txtnumpla.text <> "" Then
     fqry01 = "SELECT * " _
              & "FROM ophojveh " _
             & "WHERE hvplaca = '" & Txtnumpla.text & "'"
             
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        Lblnumpla.Caption = Ftb01!hvpropie
        Ftb01.Close
     Else
        MsgBox " Numero De Placa No Existe", 48, "Advertencia"
        Ftb01.Close
        forma = "2"
        forma1 = "14"
        buscar.Show 1
        Txtnumpla.SetFocus
     End If
  End If
End Sub

Private Function DIB_GRILLA()
     Grdconsalmul.rows = 1
     Grdconsalmul.Cols = 18
     Grdconsalmul.Clear
     
     s$ = "Número  |Fecha  |Hora|Lugar|Código y Empresa|Placa|Núm. Interno|Destino TUT|" _
        & "Vr Destino TUT|Núm. TUT|Dest. Planilla/Tiquete|Vr Destino Planilla/Tiquete|" _
        & "Planilla No.|Tiquete No.|Diferencia Recaudo|Cód. Infracción|Axuliar1|Axuliar2|Estado"
     Grdconsalmul.FormatString = s$
     Grdconsalmul.Row = 0
     Grdconsalmul.Col = 0
     Grdconsalmul.ColWidth(0) = 1000
     Grdconsalmul.ColWidth(1) = 1200
     Grdconsalmul.ColWidth(2) = 1200
     Grdconsalmul.ColWidth(3) = 1200
     Grdconsalmul.ColWidth(4) = 1900
     Grdconsalmul.ColWidth(5) = 1000
     Grdconsalmul.ColWidth(6) = 1300
     Grdconsalmul.ColWidth(7) = 1800
     Grdconsalmul.ColWidth(8) = 1300
     Grdconsalmul.ColWidth(9) = 1300
     Grdconsalmul.ColWidth(10) = 2000
     Grdconsalmul.ColWidth(11) = 2400
     Grdconsalmul.ColWidth(12) = 1300
     Grdconsalmul.ColWidth(13) = 1500
     Grdconsalmul.ColWidth(14) = 1800
     Grdconsalmul.ColWidth(15) = 1800
     Grdconsalmul.ColWidth(16) = 1500
     Grdconsalmul.ColWidth(17) = 1500
     Grdconsalmul.ColWidth(18) = 1500
     
     For i = 0 To Grdconsalmul.Cols - 1
        Grdconsalmul.Col = i
        Grdconsalmul.CellFontBold = 1
     Next i
End Function


