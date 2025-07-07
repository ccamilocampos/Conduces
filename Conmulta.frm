VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Conmultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Evasiones y Elusiones"
   ClientHeight    =   6900
   ClientLeft      =   2325
   ClientTop       =   1335
   ClientWidth     =   7800
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7800
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
      Left            =   6840
      Picture         =   "Conmulta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   2880
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
      Left            =   6840
      Picture         =   "Conmulta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   3840
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
      Left            =   6840
      Picture         =   "Conmulta.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Limpia la forma actual"
      Top             =   4920
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
      Left            =   6840
      Picture         =   "Conmulta.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir de la forma actual"
      Top             =   6000
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
      Left            =   6840
      Picture         =   "Conmulta.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   6495
      Begin VB.TextBox Txtvalbus 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   15
         ToolTipText     =   "Valor de búsqueda con respecto a lo que selecciono en criterio de búsqueda(doble Click)"
         Top             =   360
         Width           =   1335
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
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   27
      Top             =   1200
      Width           =   6495
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
         Index           =   5
         Left            =   4320
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Código Infracción"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Código Guarda"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Número Placa"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optcriterio 
         Caption         =   "Código de la Empresa"
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
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optcriterio 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   6495
      Begin Crystal.CrystalReport rptterminal 
         Left            =   600
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconmul 
         Height          =   2175
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
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
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2055
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
         Top             =   600
         Width           =   1695
      End
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
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   6495
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
         Left            =   3240
         TabIndex        =   40
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
         Left            =   4440
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
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
         Left            =   3240
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frmnummul 
      Caption         =   "Numeros de Evasiones: "
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
      Left            =   2280
      TabIndex        =   20
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Txtnumfin 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "Número final de multa por la que quiere buscar"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txtnumini 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Número de inicial de multa por la que quiere buscar"
         Top             =   600
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   1215
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
         TabIndex        =   21
         Top             =   360
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
      Height          =   975
      Left            =   2280
      TabIndex        =   23
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Optfecha 
         Caption         =   "Registro"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Optfecha 
         Caption         =   "Comparendo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Fecha en la que empieze a buscar"
         Top             =   600
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
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Fecha hasta donde quiere que busque"
         Top             =   600
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
         Left            =   1560
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   7560
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   6720
      X2              =   6720
      Y1              =   6840
      Y2              =   0
   End
End
Attribute VB_Name = "Conmultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Integer
Dim lineas As Integer
Dim estado As String
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
      
     
      
     For i = 0 To 4
         If optcriterio(i).Value = True And Txtvalbus.text = "" Then
            MsgBox "Falta Ingresar El Valor de Busqueda", 48, "Advertencia"
            Txtvalbus.SetFocus
            Exit Sub
         End If
     Next i
          
     If CDbl(Txtnumini.text) > CDbl(Txtnumfin.text) Then
        MsgBox "El Numero Inicial No Puede Ser Mayor Al Numero Final", 48, "Advertencia"
        Txtnumini.SetFocus
        Exit Sub
     End If
     
     DIB_GRILLA
     
     QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, " _
                & " muestado, hvnumint, musaldo, munomcaj, muobsanu, muauxalt, " _
                & " mudestino, mudespla, muplanilla, mutiquete " _
          & "FROM opmultas, ophojveh " _
         & "WHERE munumero >= " & Txtnumini.text & " " _
           & "AND munumero <= " & Txtnumfin.text & " " _
     
     '******************** CUANDO SOLO BUSCA POR NUMEROS **************************
     If optcriterio(5).Value = True Then
        
     End If
     
     '******************** CUANDO BUSCA POR ESTADO ********************************
     If optcriterio(0).Value = True Then
        QRY = QRY & "AND muestado LIKE '" & Lblvalbus.Caption & "%' "
     End If
     '******************** CUANDO BUSCA POR EMPRESA *******************************
     If optcriterio(1).Value = True Then
        QRY = QRY & "AND mucodemp = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
     If optcriterio(2).Value = True Then
        QRY = QRY & "AND munumpla = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR CODIGO DE PLACA DE GUARDA ********************
     If optcriterio(3).Value = True Then
        QRY = QRY & "AND muauxcon = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR CODIGO DE INFRACCION ******************
     If optcriterio(4).Value = True Then
        QRY = QRY & "AND mucodinf = " & Txtvalbus.text & " "
     End If
     
     QRY = QRY & "AND munumpla = hvplaca " _
          & "ORDER BY munumero ASC "
     
     Set Tbl1 = Dbs.Execute(QRY)
     Conmultas.Enabled = False
     Do While Not Tbl1.EOF
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
        
        If IsNull(Tbl1(16)) Then
           obser = "NO"
        Else
           obser = Tbl1(16)
        End If
        
        If IsNull(Tbl1(14)) Then
           saldo = "0"
        Else
           saldo = str(Tbl1(14))
        End If
        
        QRY = "SELECT psdes FROM geparsis WHERE pscod = '" & Tbl1(2) & "'"
        
        Set Ftb01 = Dbs.Execute(QRY)
        
        If Not Ftb01.EOF Then
           sDesInfr = Ftb01(0)
        Else
           sDesInfr = ""
        End If
        Ftb01.Close
        
        If IsNull(Tbl1(17)) Then
           scodaux = Tbl1(7)
        Else
           scodaux = Tbl1(7) & "-" & Tbl1(17)
        End If
        
        If IsNull(Tbl1(18)) Then
           sdestut = ""
        Else
           fqry01 = "SELECT * " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Tbl1(18) & "'"
             
           Set Ftb01 = Dbs.Execute(fqry01)
        
           If Not Ftb01.EOF Then
              sdestut = Tbl1(18) & "-" & Ftb01!rudestino
           Else
              sdestut = Tbl1(18)
           End If
           Ftb01.Close
        End If
        If IsNull(Tbl1(19)) Then
           sDesPla = ""
        Else
            fqry01 = "SELECT * " _
                     & "FROM opruta " _
                    & "WHERE rucodigo = '" & Tbl1(19) & "'"
             
           Set Ftb01 = Dbs.Execute(fqry01)
        
           If Not Ftb01.EOF Then
              sDesPla = Tbl1(19) & "-" & Ftb01!rudestino
           Else
              sDesPla = Tbl1(19)
           End If
           Ftb01.Close
           
        End If
        If IsNull(Tbl1(20)) Then
           snumplani = ""
        Else
           snumplani = Tbl1(20)
        End If
        If IsNull(Tbl1(21)) Then
           snumtiq = ""
        Else
           snumtiq = Tbl1(21)
        End If
        
        If IsNull(Tbl1(15)) Then
           sNombreCaja = ""
        Else
           sNombreCaja = Tbl1(15)
        End If
        
        
        
        entry = str(Tbl1(0)) + Chr(9) + sEmpresa + Chr(9) + Trim(sDesInfr) + Chr(9) + _
                Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + str(Tbl1(13)) + Chr(9) + _
                str(Tbl1(5)) + Chr(9) + str(Tbl1(6)) + Chr(9) + scodaux + Chr(9) + _
                str(Tbl1(8)) + Chr(9) + str(Tbl1(9)) + Chr(9) + str(Tbl1(10)) + Chr(9) + _
                str(Tbl1(11)) + Chr(9) + Tbl1(12) + Chr(9) + saldo + Chr(9) + sNombreCaja + Chr(9) + obser + Chr(9) + _
                sdestut + Chr(9) + sDesPla + Chr(9) + snumplani + Chr(9) + snumtiq
        
        If Tbl1(12) = "Pagada" Then
           dato1 = dato1 + 1
        ElseIf Tbl1(12) = "Educativo" Then
           dato2 = dato2 + 1
        ElseIf Tbl1(12) = "Anulado" Then
           dato3 = dato3 + 1
        ElseIf Tbl1(12) = "Pendiente" Then
           dato4 = dato4 + 1
        End If
        
        
        Grdconmul.AddItem entry
        Tbl1.MoveNext
        bandera = 1
     Loop
     Tbl1.Close
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
     
     For i = 0 To 4
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
     
     QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                    & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, " _
                    & "muestado, hvnumint, musaldo, munomcaj, muobsanu, muauxalt,  " _
                    & " mudestino, mudespla, muplanilla, mutiquete " _
              & "FROM opmultas, ophojveh "
     
     If Optfecha(0).Value = True Then
        QRY = QRY & "WHERE mufecha BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
                    & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') "
     Else
        QRY = QRY & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
                    & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') "
     End If
     
     '******************** SOLO BUSCA POR FECHA ***********************************
     If optcriterio(5).Value = True Then
        
     End If
     
     '******************** CUANDO BUSCA POR ESTADO ********************************
     If optcriterio(0).Value = True Then
        QRY = QRY & "AND muestado LIKE '" & Lblvalbus.Caption & "%' "
     End If
     
     '******************** CUANDO BUSCA POR EMPRESA *******************************
     If optcriterio(1).Value = True Then
        QRY = QRY & "AND mucodemp = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
     If optcriterio(2).Value = True Then
        QRY = QRY & "AND munumpla = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR CODIGO DE PLACA DE GUARDA **********************
     If optcriterio(3).Value = True Then
        QRY = QRY & "AND muauxcon = '" & Txtvalbus.text & "' "
     End If
     
     '******************** CUANDO BUSCA POR CODIGO DE INFRACCION *******************
     If optcriterio(4).Value = True Then
        QRY = QRY & "AND mucodinf = " & Txtvalbus.text & " "
     End If
     
     
     QRY = QRY & "AND munumpla = hvplaca " _
          & "ORDER BY munumero ASC "
          
     Set Tbl1 = Dbs.Execute(QRY)
     Conmultas.Enabled = False
     Do While Not Tbl1.EOF
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
        
        If IsNull(Tbl1(14)) Then
           saldo = "NO0"
        Else
           saldo = str(Tbl1(14))
        End If
        
        If IsNull(Tbl1(16)) Then
           obser = "0"
        Else
           obser = Tbl1(16)
        End If
        
        QRY = "SELECT psdes FROM geparsis WHERE pscod = '" & Tbl1(2) & "'"
        
        Set Ftb01 = Dbs.Execute(QRY)
        
        If Not Ftb01.EOF Then
           sDesInfr = Ftb01(0)
        Else
           sDesInfr = ""
        End If
        Ftb01.Close
        
        If IsNull(Tbl1(17)) Then
           scodaux = Tbl1(7)
        Else
           scodaux = Tbl1(7) & "-" & Tbl1(17)
        End If
        If IsNull(Tbl1(18)) Then
           sdestut = ""
        Else
           fqry01 = "SELECT * " _
               & "FROM opruta " _
              & "WHERE rucodigo = '" & Tbl1(18) & "'"
             
           Set Ftb01 = Dbs.Execute(fqry01)
        
           If Not Ftb01.EOF Then
              sdestut = Tbl1(18) & "-" & Ftb01!rudestino
           Else
              sdestut = Tbl1(18)
           End If
           Ftb01.Close
        End If
        If IsNull(Tbl1(19)) Then
           sDesPla = ""
        Else
            fqry01 = "SELECT * " _
                     & "FROM opruta " _
                    & "WHERE rucodigo = '" & Tbl1(19) & "'"
             
           Set Ftb01 = Dbs.Execute(fqry01)
        
           If Not Ftb01.EOF Then
              sDesPla = Tbl1(19) & "-" & Ftb01!rudestino
           Else
              sDesPla = Tbl1(19)
           End If
           Ftb01.Close
           
        End If
        If IsNull(Tbl1(20)) Then
           snumplani = ""
        Else
           snumplani = Tbl1(20)
        End If
        If IsNull(Tbl1(21)) Then
           snumtiq = ""
        Else
           snumtiq = Tbl1(21)
        End If
        
        If IsNull(Tbl1(15)) Then
           sNombreCaja = ""
        Else
           sNombreCaja = Tbl1(15)
        End If
        
        entry = str(Tbl1(0)) + Chr(9) + sEmpresa + Chr(9) + sDesInfr + Chr(9) + _
                Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + str(Tbl1(13)) + Chr(9) + _
                str(Tbl1(5)) + Chr(9) + str(Tbl1(6)) + Chr(9) + scodaux + Chr(9) + _
                str(Tbl1(8)) + Chr(9) + str(Tbl1(9)) + Chr(9) + str(Tbl1(10)) + Chr(9) + _
                str(Tbl1(11)) + Chr(9) + Tbl1(12) + Chr(9) + saldo + Chr(9) + sNombreCaja + Chr(9) + obser + Chr(9) + _
                sdestut + Chr(9) + sDesPla + Chr(9) + snumplani + Chr(9) + snumtiq
                
        If Tbl1(12) = "Pagada" Then
           dato1 = dato1 + 1
        ElseIf Tbl1(12) = "Educativo" Then
           dato2 = dato2 + 1
        ElseIf Tbl1(12) = "Anulado" Then
           dato3 = dato3 + 1
        ElseIf Tbl1(12) = "Pendiente" Then
           dato4 = dato4 + 1
        End If
                
        Grdconmul.AddItem entry
        Tbl1.MoveNext
        bandera = 1
     Loop
     Tbl1.Close
  End If
  Conmultas.Enabled = True
  Lblpagadas.Caption = dato1
  Lbleducativas.Caption = dato2
  Lblanuladas.Caption = dato3
  Lblpendientes.Caption = dato4
  If bandera = 0 Then
     MsgBox "No Se Encontro Ningun Registro Activo", 16, "Busqueda"
     Exit Sub
  End If
  Call Aplicar_Perfil(Me, "Conmultas")
  Exit Sub
  
terminar:
  Conmultas.Enabled = True
  MsgBox "Error al tratar de mostrar los registros o error en la base de datos " & Chr(13) + Chr(13) & " ERROR >> " & Err.Description, 48, "Advertencia"
  
  Exit Sub
Resume
End Sub

Private Sub BtnExportar_Click()
    On Error GoTo terminar
    
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
         
        
         
        For i = 0 To 4
            If optcriterio(i).Value = True And Txtvalbus.text = "" Then
               MsgBox "Falta Ingresar El Valor de Busqueda", 48, "Advertencia"
               Txtvalbus.SetFocus
               Exit Sub
            End If
        Next i
             
        If CDbl(Txtnumini.text) > CDbl(Txtnumfin.text) Then
           MsgBox "El Numero Inicial No Puede Ser Mayor Al Numero Final", 48, "Advertencia"
           Txtnumini.SetFocus
           Exit Sub
        End If
        
        DIB_GRILLA
        
        QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, " _
                & " muestado, hvnumint, musaldo, munomcaj, muobsanu, muauxalt, " _
                & " mudestino, mudespla, muplanilla, mutiquete " _
          & "FROM opmultas, ophojveh " _
         & "WHERE munumero >= " & Txtnumini.text & " " _
           & "AND munumero <= " & Txtnumfin.text & " " _
     
        '******************** CUANDO SOLO BUSCA POR NUMEROS **************************
        If optcriterio(5).Value = True Then
           
        End If
        
        '******************** CUANDO BUSCA POR ESTADO ********************************
        If optcriterio(0).Value = True Then
           QRY = QRY & "AND muestado LIKE '" & Lblvalbus.Caption & "%' "
        End If
        '******************** CUANDO BUSCA POR EMPRESA *******************************
        If optcriterio(1).Value = True Then
           QRY = QRY & "AND mucodemp = '" & Txtvalbus.text & "' "
        End If
        
        '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
        If optcriterio(2).Value = True Then
           QRY = QRY & "AND munumpla = '" & Txtvalbus.text & "' "
        End If
        
        '******************** CUANDO BUSCA POR CODIGO DE PLACA DE GUARDA ********************
        If optcriterio(3).Value = True Then
           QRY = QRY & "AND muauxcon = '" & Txtvalbus.text & "' "
        End If
        
        '******************** CUANDO BUSCA POR CODIGO DE INFRACCION ******************
        If optcriterio(4).Value = True Then
           QRY = QRY & "AND mucodinf = " & Txtvalbus.text & " "
        End If
        
        QRY = QRY & "AND munumpla = hvplaca " _
             & "ORDER BY munumero ASC "
    
    ElseIf Optbuscar(1).Value = True Then
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
           
           For i = 0 To 4
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
           
           QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                          & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, " _
                          & "muestado, hvnumint, musaldo, munomcaj, muobsanu, muauxalt,  " _
                          & " mudestino, mudespla, muplanilla, mutiquete " _
                    & "FROM opmultas, ophojveh "
           
           If Optfecha(0).Value = True Then
              QRY = QRY & "WHERE mufecha BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
                          & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') "
           Else
              QRY = QRY & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
                          & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') "
           End If
           
           '******************** SOLO BUSCA POR FECHA ***********************************
           If optcriterio(5).Value = True Then
              
           End If
           
           '******************** CUANDO BUSCA POR ESTADO ********************************
           If optcriterio(0).Value = True Then
              QRY = QRY & "AND muestado LIKE '" & Lblvalbus.Caption & "%' "
           End If
           
           '******************** CUANDO BUSCA POR EMPRESA *******************************
           If optcriterio(1).Value = True Then
              QRY = QRY & "AND mucodemp = '" & Txtvalbus.text & "' "
           End If
           
           '******************** CUANDO BUSCA POR NUMERO DE PLACA ***********************
           If optcriterio(2).Value = True Then
              QRY = QRY & "AND munumpla = '" & Txtvalbus.text & "' "
           End If
           
           '******************** CUANDO BUSCA POR CODIGO DE PLACA DE GUARDA **********************
           If optcriterio(3).Value = True Then
              QRY = QRY & "AND muauxcon = '" & Txtvalbus.text & "' "
           End If
           
           '******************** CUANDO BUSCA POR CODIGO DE INFRACCION *******************
           If optcriterio(4).Value = True Then
              QRY = QRY & "AND mucodinf = " & Txtvalbus.text & " "
           End If
           
           QRY = QRY & "AND munumpla = hvplaca " _
                & "ORDER BY munumero ASC "
    End If
    
    Conmultas.Enabled = False
    ExportarExcel (QRY)
    Conmultas.Enabled = True
    
    Call Aplicar_Perfil(Me, "Conmultas")
    Exit Sub
terminar:
  Conmultas.Enabled = True
  MsgBox "Error al tratar de mostrar los registros o error en la base de datos " & Chr(13) + Chr(13) & " ERROR >> " & Err.Description, 48, "Advertencia"
  Exit Sub
End Sub

Private Sub ExportarExcel(QRY As String)
    
    Lblpagadas.Caption = ""
    Lblpendientes.Caption = ""
    Lbleducativas.Caption = ""
    Lblanuladas.Caption = ""
    dato1 = 0
    dato2 = 0
    dato3 = 0
    dato4 = 0
    
    On Error GoTo Errores
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "CONSECUTIVO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "EMPRESA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "COD. INFRAC."
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "NOMBRE NOT."
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "PLACA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "NUMERO INT."
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "FECHA"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "HORA"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "AUX. CONTROL"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "REC. PAGO"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "DESCUENTO"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "VAL. DESCUENTO"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "NETO PAGAR"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "ESTADO"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "SALDO"
    ApExcel.Cells(fila, 15).Font.Size = 10
    ApExcel.Cells(fila, 16).formula = "CAJA"
    ApExcel.Cells(fila, 16).Font.Size = 10
    ApExcel.Cells(fila, 17).formula = "OBS. ANULACION"
    ApExcel.Cells(fila, 17).Font.Size = 10
    ApExcel.Cells(fila, 18).formula = "DES. TUT"
    ApExcel.Cells(fila, 18).Font.Size = 10
    ApExcel.Cells(fila, 19).formula = "DES. PLANILLA"
    ApExcel.Cells(fila, 19).Font.Size = 10
    ApExcel.Cells(fila, 20).formula = "NUM. PLANILLA"
    ApExcel.Cells(fila, 20).Font.Size = 10
    ApExcel.Cells(fila, 21).formula = "NUM. TIQUETE"
    ApExcel.Cells(fila, 21).Font.Size = 10
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
            
            'Saldos
            '*********************************************
            If IsNull(Tbl1!musaldo) Then
                saldo = "NO"
            Else
                saldo = Tbl1!musaldo
            End If
            '*********************************************
            
            'Observacion
            '*********************************************
            If IsNull(Tbl1!muobsanu) Then
                obser = "0"
            Else
                obser = Tbl1!muobsanu
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
            
            'Auxiliar control
            '*********************************************
            If IsNull(Tbl1!muauxalt) Then
                scodaux = Tbl1!muauxcon
            Else
                scodaux = Tbl1!muauxcon & "-" & Tbl1!muauxalt
            End If
            '*********************************************
            
            'Destino Tasa de Uso
            '*********************************************
            If IsNull(Tbl1!mudestino) Then
                sdestut = ""
            Else
                QRY = "SELECT rudestino " _
                        & "FROM opruta " _
                        & "WHERE rucodigo = '" & Tbl1!mudestino & "' "
                
                Set Tbl2 = Dbs.Execute(QRY)
                
                If Not Tbl2.EOF Then
                    sdestut = Tbl1!mudestino & "-" & Tbl2!rudestino
                Else
                    sdestut = Tbl1!mudestino
                End If
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
            
            'Número Planilla
            '*********************************************
            If IsNull(Tbl1!muplanilla) Then
                snumplani = ""
            Else
                snumplani = Tbl1!muplanilla
            End If
            '*********************************************
            
            'Número de Tiquete
            '*********************************************
            If IsNull(Tbl1!mutiquete) Then
                snumtiq = ""
            Else
                snumtiq = Tbl1!mutiquete
            End If
            '*********************************************
            
            'Nombre Caja
            '*********************************************
            If IsNull(Tbl1!munomcaj) Then
                sNombreCaja = ""
            Else
                sNombreCaja = Tbl1!munomcaj
            End If
            '*********************************************
            
            'Llenamos los datos en Excel
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!munumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & sEmpresa & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & sDesInfr & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!munomnot & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!munumpla & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!hvnumint & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!mufecha & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!muhora & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & scodaux & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!murecpag & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & Tbl1!mudescue & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!muvaldes & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & Tbl1!munetpag & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & Tbl1!muestado & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & saldo & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            ApExcel.Cells(fila, 16).formula = "" & sNombreCaja & ""
            ApExcel.Cells(fila, 16).Font.Size = 10
            ApExcel.Cells(fila, 17).formula = "" & obser & ""
            ApExcel.Cells(fila, 17).Font.Size = 10
            ApExcel.Cells(fila, 18).formula = "" & sdestut & ""
            ApExcel.Cells(fila, 18).Font.Size = 10
            ApExcel.Cells(fila, 19).formula = "" & sDesPla & ""
            ApExcel.Cells(fila, 19).Font.Size = 10
            ApExcel.Cells(fila, 20).formula = "" & snumplani & ""
            ApExcel.Cells(fila, 20).Font.Size = 10
            ApExcel.Cells(fila, 21).formula = "" & snumtiq & ""
            ApExcel.Cells(fila, 21).Font.Size = 10
            
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
  'Grdconmul.Text <> "" And
  If Grdconmul.rows > 1 Then
     lineas = Grdconmul.rows - 1
     Grdconmul.Row = 1
     Dbs.Execute "DELETE FROM optempo1 WHERE  tempc = '" & Pc & "' "
            
     Conmultas.Enabled = False
     Dbs.BeginTrans
     On Error GoTo Err_Trans
     Do
      
        ReDim dato(1, 21)
                
        For i = 0 To 20
            Grdconmul.Col = i
            dato(1, i) = Trim(Grdconmul.text)
        Next i
            
        ' insertamos en la tabla
             
        Dbs.Execute "INSERT INTO optempo1 " _
        & "(temnumero1,      temtext8,           temtext2,           temtext3,            temtext4,            temnumero3,          temdate1,           temtext5,             temtext6,            temnumero2,          temnumdob1,          temnumdob3,           temnumdob2,           temtext7,   temnumdob4, temtext10, tempc, temmemo1, temtext11, temtext12, temtext13, temtext14) VALUES ('" _
        & dato(1, 0) & "' ,'" & dato(1, 1) & "','" & dato(1, 2) & "','" & dato(1, 3) & "','" & dato(1, 4) & "','" & dato(1, 5) & "' ,to_date('" & dato(1, 6) & "','MM/DD/YYYY'),'" & Format(dato(1, 7), "HH:mm") & "','" & dato(1, 8) & "','" & dato(1, 9) & "','" & dato(1, 10) & "','" & dato(1, 11) & "','" & dato(1, 12) & "','" & dato(1, 13) & "'," & dato(1, 14) & ",'" & dato(1, 15) & "','" & Pc & "','" & dato(1, 16) & "','" & dato(1, 17) & "','" & dato(1, 18) & "','" & dato(1, 19) & "','" & dato(1, 20) & "')"
             
        ' fin insertar
              
             
        lineas = lineas - 1
             
        If lineas <> 0 Then
           Grdconmul.Row = Grdconmul.Row + 1
        End If
     Loop Until lineas = 0  'do_while que se mueve x la grilla
     
     Dbs.CommitTrans
     On Error GoTo err_mensaje
       
     If Optbuscar(0).Value = True Then
        inicio1 = Txtnumini.text
        final = Txtnumfin.text
     ElseIf Optbuscar(1).Value = True Then
        inicio1 = Mskfecini.FormattedText
        final = Mskfecfin.FormattedText
     End If
       
     reporte = RpPath + "\conmulta.rpt"
     rptterminal.ReportFileName = reporte
     rptterminal.Connect = G_Conexion_reportes
     rptterminal.Formulas(0) = "totpag = '" & Lblpagadas.Caption & "'"
     rptterminal.Formulas(1) = "inicio1 = '" & inicio1 & "'"
     rptterminal.Formulas(2) = "final = '" & final & "'"
     rptterminal.Formulas(3) = "empresa = '" & G_NomEmpresa & "'"
     rptterminal.Formulas(4) = "totedu = '" & Lbleducativas.Caption & "'"
     rptterminal.Formulas(5) = "totpen = '" & Lblpendientes.Caption & "'"
     rptterminal.Formulas(6) = "totanu = '" & Lblanuladas.Caption & "'"
     
     rptterminal.WindowState = crptMaximized
     rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
     rptterminal.Destination = 0
     rptterminal.Action = 1
     Conmultas.Enabled = True
     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
  End If
  
  Exit Sub
Err_Trans:
   Dbs.RollbackTrans
   
err_mensaje:
   Conmultas.Enabled = True
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
  optcriterio(5).Value = True
  Txtvalbus.text = ""
  Lblvalbus.Caption = ""
  Grdconmul.rows = 1
  Grdconmul.Cols = 2
  Grdconmul.Clear
  Optbuscar(0).Value = True
  Lblpagadas.Caption = ""
  Lblpendientes.Caption = ""
  Lbleducativas.Caption = ""
  Lblanuladas.Caption = ""
  dato1 = 0
  Optfecha(0).Value = True
  DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Btntotalizar_Click()
  If Grdconmul.text <> "" And Grdconmul.rows > 1 Then
     dato1 = 0
     Grdconmul.Col = 12
     For i = 1 To Grdconmul.rows - 1
         Grdconmul.Row = i
         dato1 = dato1 + CDbl(Grdconmul.text)
     Next i
     MskTotal.text = dato1
  End If
End Sub

Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Conmultas")
End Sub

Private Sub Form_Load()
  Optbuscar_Click (index)
  DIB_GRILLA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Mskfecfin_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Or (KeyAscii% = 47) Then
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

Private Sub Msktotal_Change()

End Sub

Private Sub Optbuscar_Click(index As Integer)
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


Private Sub optcriterio_Click(index As Integer)
  Txtvalbus.text = ""
  Lblvalbus.Caption = ""
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


Private Sub Txtvalbus_DblClick()
  If optcriterio(5).Value = False Then
     Txtvalbus.text = ""
     If optcriterio(0).Value = True Then
        forma = "15"
        forma1 = "19"
        Tabla_Catalogo = "ESTADOS"
        buscar.Show 1
     End If
     
     If optcriterio(1).Value = True Then
        forma = "16"
        forma1 = "11"
        buscar.Show 1
     End If
     
     If optcriterio(2).Value = True Then
        forma = "2"
        forma1 = "3"
        buscar.Show 1
     End If
     
     If optcriterio(3).Value = True Then
        forma = "15"
        forma1 = "16"
        Tabla_Catalogo = "AUXILIARES DE CONTROL"
        buscar.Show 1
     End If
     
     If optcriterio(4).Value = True Then
        forma = "4"
        forma1 = "2"
        buscar.Show 1
     End If
     Txtvalbus_LostFocus
  End If
End Sub

Private Sub Txtvalbus_KeyPress(KeyAscii As Integer)
  If optcriterio(2).Value = True Then
     'Char = Chr(KeyAscii%)
     'KeyAscii = Asc(UCase(Char))  ' mayuscula
     'longitud = Len(Txtvalbus.Text)
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
     '       Exit Sub
     '   Else
     '       MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
     '       KeyAscii% = 0
     '       Exit Sub
     '   End If
     'End If
     Char = Chr(KeyAscii%)
     KeyAscii = Asc(UCase(Char))  ' mayuscula
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  Else
     Char = Chr(KeyAscii%)
     KeyAscii = Asc(UCase(Char))  ' mayuscula
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
End Sub

Private Sub Txtvalbus_LostFocus()
  If Txtvalbus.text <> "" And optcriterio(5).Value = False Then
     If optcriterio(0).Value = True Then
        fqry01 = "SELECT dsdes " _
                 & "FROM gesuptip, gedetsuptip " _
                & "WHERE dscoddet = '" & Txtvalbus.text & "' " _
                  & "AND dscodtip = stcodtip " _
                  & "AND stdes = 'ESTADOS' "

        Set Ftb01 = Dbs.Execute(fqry01)
       
        If Not Ftb01.EOF Then
           Lblvalbus.Caption = Mid(Ftb01!dsdes, 1, 1) & LCase(Mid(Ftb01!dsdes, 2))
           Ftb01.Close
        Else
           MsgBox "Codigo De Estado No Existe", 48, "Advertencia"
           Ftb01.Close
           forma = "15"
           forma1 = "19"
           Tabla_Catalogo = "ESTADOS"
           buscar.Show 1
           Txtvalbus.SetFocus
        End If
     End If
  
     If optcriterio(1).Value = True Then
         QRY = "SELECT henombre "
         If Len(Txtvalbus.text) = 4 Then
            QRY = QRY & "FROM ophojempcon  "
         Else
            QRY = QRY & "FROM ophojemp  "
         End If
         QRY = QRY & "WHERE hecodigo = '" & Txtvalbus.text & "'"
             
         Set Ftb01 = Dbs.Execute(QRY)
         
         If Not Ftb01.EOF Then
            Lblvalbus.Caption = Ftb01!henombre
            Ftb01.Close
         Else
            MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
            Ftb01.Close
            forma = "16"
            forma1 = "11"
            buscar.Show 1
            Txtvalbus.SetFocus
         End If
        
      End If
  
      If optcriterio(2).Value = True Then
         fqry01 = "SELECT * " _
                  & "FROM ophojveh " _
                 & "WHERE hvplaca = '" & Txtvalbus.text & "'"
                 
         Set Ftb01 = Dbs.Execute(fqry01)
         
         If Not Ftb01.EOF Then
            Lblvalbus.Caption = Ftb01!hvpropie
            Ftb01.Close
         Else
            MsgBox "Numero De Placa De Vehiculo No Existe", 48, "Advertencia"
            Ftb01.Close
            forma = "2"
            forma1 = "3"
            buscar.Show 1
            Txtvalbus.SetFocus
         End If
      End If
      
      If optcriterio(3).Value = True Then
         fqry01 = "SELECT dsdes " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txtvalbus.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'AUXILIARES DE CONTROL' "
                
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lblvalbus.Caption = Ftb01!dsdes
            Ftb01.Close
         Else
            MsgBox "Auxiliar De Control No Existe", 48, "Advertencia"
            Ftb01.Close
            forma = "15"
            forma1 = "16"
            Tabla_Catalogo = "AUXILIARES DE CONTROL"
            buscar.Show 1
            Txtvalbus.SetFocus
         End If
      End If
      
      If optcriterio(4).Value = True Then
         fqry01 = "SELECT * " _
                  & "FROM geparsis " _
                 & "WHERE pscod = '" & Txtvalbus.text & "' "
         
         Set Ftb01 = Dbs.Execute(fqry01)
      
         If Not Ftb01.EOF Then
            Lblvalbus.Caption = Ftb01!psdes
            Ftb01.Close
         Else
            MsgBox "Codigo De Infraccion No Existe", 48, "Advertencia"
            Ftb01.Close
            forma = "4"
            forma1 = "2"
            buscar.Show 1
            Txtvalbus.SetFocus
         End If
      End If
   End If
End Sub

Private Function DIB_GRILLA()
     Grdconmul.rows = 1
     Grdconmul.Cols = 21
     Grdconmul.Clear
     
     
     s$ = "Número  |Empresa  |Cód. Infracción   |Nom. Notificado  |Núm. Placa  |Núm. Int| Fec. Evasion  |Hora |Aux. Control |Rec. Pago |Descuento |Valor Descuento |Neto Pagar |Estado|Saldo |Caja|Obs. Anulación|Dest. TUT|Dest. PLanilla|Núm. Planilla|Núm. Tiquete"
     Grdconmul.FormatString = s$
     Grdconmul.Row = 0
     Grdconmul.Col = 0
     Grdconmul.ColWidth(0) = 1000
     Grdconmul.ColWidth(1) = 1400
     Grdconmul.ColWidth(2) = 1400
     Grdconmul.ColWidth(3) = 2500
     Grdconmul.ColWidth(4) = 1200
     Grdconmul.ColWidth(5) = 1200
     Grdconmul.ColWidth(6) = 1200
     Grdconmul.ColWidth(7) = 1200
     Grdconmul.ColWidth(8) = 1200
     Grdconmul.ColWidth(9) = 1200
     Grdconmul.ColWidth(10) = 1200
     Grdconmul.ColWidth(11) = 1200
     Grdconmul.ColWidth(12) = 1200
     Grdconmul.ColWidth(13) = 1200
     Grdconmul.ColWidth(14) = 1200
     Grdconmul.ColWidth(15) = 1200
     Grdconmul.ColWidth(16) = 1800
     Grdconmul.ColWidth(17) = 1800
     Grdconmul.ColWidth(18) = 1800
     Grdconmul.ColWidth(19) = 1800
     Grdconmul.ColWidth(20) = 1800
         
     For i = 0 To Grdconmul.Cols - 1
     
        Grdconmul.Col = i
        Grdconmul.CellFontBold = 1
     Next i
End Function
