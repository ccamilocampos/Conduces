VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Concietur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cierre de Turno"
   ClientHeight    =   4650
   ClientLeft      =   1485
   ClientTop       =   2265
   ClientWidth     =   10515
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      Picture         =   "Concietu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton BtnRodamiento 
      Caption         =   "&Rodamiento"
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
      Picture         =   "Concietu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Totaliza todos los netos a pagar que se encuentran en la grilla"
      Top             =   3840
      Width           =   855
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
      Left            =   1920
      TabIndex        =   25
      Top             =   3000
      Width           =   1695
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "Concietu.frx":0598
         Left            =   120
         List            =   "Concietu.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Tipos de terminales"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
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
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
      Begin VB.TextBox txtcaja 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Caja "
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
      Left            =   120
      Picture         =   "Concietu.frx":059C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   3840
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
      Left            =   3960
      Picture         =   "Concietu.frx":09DE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir de la forma actual"
      Top             =   3840
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
      Left            =   2040
      Picture         =   "Concietu.frx":0E20
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   3840
      Width           =   855
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
      Left            =   1080
      Picture         =   "Concietu.frx":1262
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Totaliza todos los netos a pagar que se encuentran en la grilla"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   3840
      TabIndex        =   19
      Top             =   0
      Width           =   6495
      Begin Crystal.CrystalReport rptterminal 
         Left            =   600
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconcie 
         Height          =   3255
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox Txtvalbus 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Valor de búsqueda "
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Valor de Búsqueda:"
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
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterio"
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
      TabIndex        =   16
      Top             =   1200
      Width           =   3495
      Begin VB.OptionButton Optcriterio 
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
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Optcriterio 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Optcriterio 
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
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Optcriterio 
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
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   855
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
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   3495
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Fecha en la que empiece a buscar"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "Mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Fecha hasta donde quiere que busque"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "Mmm/dd/yyyy"
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
         Left            =   2040
         TabIndex        =   15
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Msktotal 
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      ToolTipText     =   "Valor total de las multas que se encuentran en la grilla"
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
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
      Format          =   "$#,##0;($#,##0)"
      PromptChar      =   "_"
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   5880
      TabIndex        =   24
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   3720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   3720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   10440
      X2              =   10440
      Y1              =   4920
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   10440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   3720
      X2              =   3720
      Y1              =   3720
      Y2              =   0
   End
End
Attribute VB_Name = "Concietur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecini
Dim fecfin
Dim VehCentralOrigen As Double
Dim VehSurOrigen As Double
Dim VehNorteOrigen As Double
Dim TotalTUCentralOrigen As Double
Dim TotalTUSurOrigen As Double
Dim TotalTUNorteOrigen As Double
Dim TotalAlcCentralOrigen As Double
Dim TotalAlcSurOrigen As Double
Dim TotalAlcNorteOrigen As Double
Dim TotalAlcCentralInfluencia As Double
Dim TotalAlcSurInfluencia As Double
Dim TotalAlcNorteInfluencia As Double
Dim TotalTagNorte As Double
Dim TotalTagCentral As Double
Dim TotalTagSur As Double

Private Function Traer_Cant_Alc_influencia(ByVal sNumInicial As Double, ByVal sNumFinal As Double, ByVal sCaja As String) As Double
'Trae la cantidad de pruebas que se cobraron de influencia

   QRY = "SELECT COUNT(*) AS sCantidad " _
         & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
        & "WHERE conumero between " & sNumInicial & " AND " & sNumFinal & " " _
          & "AND conomcaj = '" & sCaja & "' " _
          & "AND coestado = 'Activo' " _
          & "AND ruinflue = 'Si' " _
          & "AND covalalc > 0 " _
          & "AND comodo = 'P' " _

   Set Tbl10 = Dbs.Execute(QRY)
   
   
   Traer_Cant_Alc_influencia = 0
   
   If Not Tbl10.EOF Then
      Traer_Cant_Alc_influencia = Tbl10!sCantidad
   End If
   Tbl10.Close


End Function



Private Sub Imprimir_Cierre(ByVal fPLanilla As Double, ByVal fCaja As String)
   QRY = "SELECT cifecha, citurno, cinumini, cinumfin, cicodope, cihoraap, cianula, " _
            & "ciorigen, citransi, citorige, cittrans, ciinfrac, citinfra, citotgen, " _
            & "cibasmon, citanula, citvalalcori, citvalalctra, citparqu, cinomcaj, " _
            & "cinumpla, cicredit, citcredi, citotrinf, citelusion, citevasion, " _
            & "citevastra, ciobserv, citvigencia, cittag, cicupos, cicupvaltasa, " _
            & "cicupvalalco, citransitonorte, cittransitonorte, cicupostransito, " _
            & "cicupostransitoaut, cicupvaltasatransito, ciinfluencia, citinfluencia, " _
            & "citvalalcinf, cicuposinfluencia, cicupvaltasainfluencia, cicupalcinfluencia  " _
       & "FROM cocietur " _
      & "WHERE cinumpla = " & fPLanilla & " " _
        & "AND cinomcaj = '" & fCaja & "' "
                            
   Set Tbl1 = Dbs.Execute(QRY)
                           
   If Not Tbl1.EOF Then
      netpago = Tbl1!citorige + Tbl1!cittrans + Tbl1!citinfra + Tbl1!citvalalcori + Tbl1!citvalalctra + _
                Tbl1!citparqu - Tbl1!cibasmon + Tbl1!citotrinf + Tbl1!citelusion + Tbl1!citevasion + _
                Tbl1!citevastra + Tbl1!citvigencia + Tbl1!cittag + Tbl1!cittransitonorte + Tbl1!citinfluencia + Tbl1!citvalalcinf
      netpago = Format(netpago, "##,##")
      
      imprimir.Show 1
      If G_Imprimir = "0" Then
         Exit Sub
      End If
      reporte = RpPath + "\cuaturno.rpt"
      
      totanumul = "0"
      rptterminal.Connect = G_Conexion_reportes
      rptterminal.ReportFileName = reporte
      rptterminal.Formulas(0) = "observaciones = '" & Tbl1!ciobserv & "'"
      rptterminal.Formulas(1) = "total = '" & Format(Tbl1!citotgen, "##,##0") & "'"
      rptterminal.Formulas(2) = "totanucon = '" & Tbl1!citanula & "'"
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
      rptterminal.Formulas(46) = "totalccupoinfluencia = '" & Format(Tbl1!cicupvaltasainfluencia, "##,##0") & "'"
      rptterminal.Formulas(47) = "tottasacupoinfluencia = '" & Format(Tbl1!cicupalcinfluencia, "##,##0") & "'"
      
      
      QRY = "INSERT INTO optempo1 " _
               & "(temtext1, temtext2, temtext3, temnumdob1, tempc) " _
               & "SELECT TO_CHAR(dbhora, 'HH24:MI'),dbcodban, banombre, dbvalor,'" & Pc & "' " _
                 & "FROM codetban, cobancos " _
                & "WHERE dbfecha = TO_DATE('" & Tbl1!cifecha & "','MM/DD/YYYY') " _
                  & "AND dbcaja = '" & fCaja & "' " _
                  & "AND dbturno = " & Tbl1!citurno & " " _
                  & "AND dbcodban = bacodigo "
      
      Dbs.Execute QRY
      rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
      rptterminal.PrinterName = G_PrinterName
      rptterminal.PrinterPort = G_PrinterPort
      rptterminal.PrinterDriver = G_PrinterDriver
      rptterminal.Destination = 0
      rptterminal.WindowState = crptMaximized
      rptterminal.Action = 1
      
   End If
End Sub


Private Sub BtnBuscar_Click()
  
   On Error GoTo terminar
    
   fecini = Mskfecini.text
   fecfin = Mskfecfin.text
     
   VehCentralOrigen = 0
   VehSurOrigen = 0
   VehNorteOrigen = 0
   TotalTUCentralOrigen = 0
   TotalTUSurOrigen = 0
   TotalTUNorteOrigen = 0
   TotalAlcCentralOrigen = 0
   TotalAlcSurOrigen = 0
   TotalAlcNorteOrigen = 0
   TotalAlcCentralInfluencia = 0
   TotalAlcSurInfluencia = 0
   TotalAlcNorteInfluencia = 0
   TotalTagCentral = 0
   TotalTagSur = 0
   TotalTagNorte = 0
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
  
  If (Optcriterio(0).Value = True Or Optcriterio(1).Value = True Or Optcriterio(3).Value = True) And Txtvalbus.text = "" Then
     MsgBox "Falta Ingresar El Valor De Busqueda", 48, "Advertencia"
     Txtvalbus.SetFocus
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
  
  bandera = 0
  Msktotal.text = ""
  
  DIB_GRILLA
   QRY = "SELECT cifecha, citurno, cinumini, cinumfin, cicodope, cihoraap, cimanual, cianula, " _
              & "ciorigen, citransi, citorige, cittrans, citanula, ciinfrac, citinfra, citotgen, " _
              & "civalalcori, citvalalcori, civalalctra, citvalalctra, civalalcinf, citvalalcinf, citparqu, cinomcaj,  " _
              & "cicredit, citcredi, catipo, citelusion, citevasion, citotrinf, citevastra, cinumpla, " _
              & "citvigencia, cittag, cicupos, cicupvaltasa, cicupvalalco, citransitonorte, " _
              & "cittransitonorte, cicupostransito, cicupostransitoaut, cicupvaltasatransito, " _
              & "ciinfluencia, citinfluencia, cicuposinfluencia, cicupvaltasainfluencia, cicupalcinfluencia " _
         & "FROM cocietur, cocaja " _
        & "WHERE cifecha BETWEEN TO_DATE(' " & fecini & "','MM/DD/YYYY') " _
          & "AND TO_DATE('" & fecfin & "','MM/DD/YYYY') " _
          & "AND cinomcaj = canomcaj "
          
  If Optcriterio(0).Value = True Then
     QRY = QRY & "AND citurno = " & Txtvalbus.text & " " _
               & "AND cinomcaj LIKE '" & Trim(txtcaja.text) & "%' "
  ElseIf Optcriterio(1).Value = True Then
     QRY = QRY & "AND cicodope = '" & Txtvalbus.text & "' "
  ElseIf Optcriterio(2).Value = True Then
     
   ElseIf Optcriterio(3).Value = True Then
     QRY = QRY & "AND cinomcaj = '" & Txtvalbus.text & "' "
  End If
  
  If CmbTerminal.ListIndex > 0 Then
      'ha escogido un terminal
      QRY = QRY & " AND catipo = '" & Mid(CmbTerminal.text, 1, 1) & "'"
   End If
      
  
  QRY = QRY & "ORDER BY cifecha, citurno ASC "
  
  Set Tbl1 = Dbs.Execute(QRY)
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
   End If
     
  total = 0
  Concietur.Enabled = False
   Do While Tbl1.EOF = False
      DoEvents
      alcori = 0
      talcori = 0
      alctra = 0
      talctra = 0
      parque = 0
      scitelusion = 0
      scitevasion = 0
      scitotrinf = 0
      scitevastra = 0
      scittag = 0
      alcinf = 0
      talcinf = 0
      
      If Not IsNull(Tbl1!civalalcori) Then
         alcori = str(Tbl1!civalalcori)
      End If
      If Not IsNull(Tbl1!citvalalcori) Then
         talcori = str(Tbl1!citvalalcori)
      End If
      
      If Not IsNull(Tbl1!civalalctra) Then
         alctra = str(Tbl1!civalalctra)
      End If
      If Not IsNull(Tbl1!citvalalctra) Then
         talctra = str(Tbl1!citvalalctra)
      End If
      
      If Not IsNull(Tbl1!civalalcinf) Then
         alcinf = CDbl(Tbl1!civalalcinf)
      End If
      
      If Not IsNull(Tbl1!citvalalcinf) Then
         talcinf = CDbl(Tbl1!citvalalcinf)
      End If
      
      If Not IsNull(Tbl1!citparqu) Then
         parque = Tbl1!citparqu
      End If
      
      If Not IsNull(Tbl1!citelusion) Then
         scitelusion = Tbl1!citelusion
      End If
      
      If Not IsNull(Tbl1!citevasion) Then
         scitevasion = Tbl1!citevasion
      End If
      
      If Not IsNull(Tbl1!citotrinf) Then
         scitotrinf = Tbl1!citotrinf
      End If
      
      If Not IsNull(Tbl1!citevastra) Then
         scitevastra = Tbl1!citevastra
      End If
     
      If Not IsNull(Tbl1!cittag) Then
         'scitexcreta = Tbl1!cittag
         scittag = Tbl1!cittag
      End If
      
      sCantCuposOrigen = 0
      If Not IsNull(Tbl1!cicupos) Then
          sCantCuposOrigen = Tbl1!cicupos
      End If
      sCantCuposTransito = 0
      If Not IsNull(Tbl1!cicupostransito) Then
          sCantCuposTransito = Tbl1!cicupostransito
      End If
      sCantCuposInfluencia = 0
      If Not IsNull(Tbl1!cicuposinfluencia) Then
          sCantCuposInfluencia = Tbl1!cicuposinfluencia
      End If
      
      
      sCantCuposTranAut = 0
      If Not IsNull(Tbl1!cicupostransitoaut) Then
          sCantCuposTranAut = Tbl1!cicupostransitoaut
      End If
      sCupoTasaOrigen = 0
      If Not IsNull(Tbl1!cicupvaltasa) Then
          sCupoTasaOrigen = Tbl1!cicupvaltasa
      End If
      
      sCupoTasaTransito = 0
      If Not IsNull(Tbl1!cicupvaltasatransito) Then
          sCupoTasaTransito = Tbl1!cicupvaltasatransito
      End If
      sCupoTasaInfluencia = 0
      If Not IsNull(Tbl1!cicupvaltasainfluencia) Then
          sCupoTasaInfluencia = Tbl1!cicupvaltasainfluencia
      End If


      sCupoAlcohol = 0
      If Not IsNull(Tbl1!cicupvalalco) Then
          sCupoAlcohol = Tbl1!cicupvalalco
      End If
      sCupoAlcoholInfluencia = 0
      If Not IsNull(Tbl1!cicupalcinfluencia) Then
          sCupoAlcoholInfluencia = Tbl1!cicupalcinfluencia
      End If
       
      sCanAlcInfluencia = Traer_Cant_Alc_influencia(Tbl1!cinumini, Tbl1!cinumfin, Tbl1!cinomcaj)
      
      
      entry = str(Tbl1!cifecha) + Chr(9) + str(Tbl1!citurno) + Chr(9) + Tbl1!cinumini + Chr(9) + Tbl1!cinumfin + Chr(9) + _
              Tbl1!cicodope + Chr(9) + Format(Tbl1!cihoraap, "hh:mm") + Chr(9) + str(Tbl1!cimanual) + Chr(9) + str(Tbl1!cianula) + Chr(9) + _
              str(Tbl1!ciorigen) + Chr(9) + str(Tbl1!citransi) + Chr(9) + str(Tbl1!citorige) + Chr(9) + str(Tbl1!cittrans) + Chr(9) + _
              str(Tbl1!citanula) + Chr(9) + str(Tbl1!ciinfrac) + Chr(9) + str(Tbl1!citinfra) + Chr(9) + str(scitelusion) + Chr(9) + _
              str(scitevasion) + Chr(9) + str(scitotrinf) + Chr(9) + str(scitevastra) + Chr(9) + str(alcori) + Chr(9) + _
              str(talcori) + Chr(9) + str(alctra) + Chr(9) + str(talctra) + Chr(9) + str(parque) + Chr(9) + _
              str(Tbl1!citotgen) + Chr(9) + Tbl1!cinomcaj + Chr(9) + str(Tbl1!cicredit) + Chr(9) + str(Tbl1!citcredi) + Chr(9) + _
              CStr(Tbl1!cinumpla) + Chr(9) + CStr(Tbl1!citvigencia) + Chr(9) + CStr(scittag) + Chr(9) + CStr(sCantCuposOrigen) + Chr(9) + _
              CStr(sCantCuposTransito) + Chr(9) + CStr(sCantCuposInfluencia) + Chr(9) + CStr(sCantCuposTranAut) + Chr(9) + CStr(sCupoTasaOrigen) + Chr(9) + _
              CStr(sCupoTasaTransito) + Chr(9) + CStr(sCupoTasaInfluencia) + Chr(9) + CStr(sCupoAlcohol) + Chr(9) + CStr(sCupoAlcoholInfluencia) + Chr(9) + _
              CStr(Tbl1!citransitonorte) + Chr(9) + CStr(Tbl1!cittransitonorte) + Chr(9) + Tbl1!catipo + Chr(9) + CStr(Tbl1!ciinfluencia) + Chr(9) + _
              CStr(Tbl1!citinfluencia) + Chr(9) + str(alcinf) + Chr(9) + str(talcinf) + Chr(9) + CStr(sCanAlcInfluencia)

       If Tbl1!catipo = "C" Then
          VehCentralOrigen = VehCentralOrigen + Tbl1!ciorigen + Tbl1!cicredit + Tbl1!ciinfluencia
          TotalTUCentralOrigen = TotalTUCentralOrigen + Tbl1!citorige + Tbl1!citcredi + Tbl1!citinfluencia
          TotalAlcCentralOrigen = TotalAlcCentralOrigen + talcori
          TotalAlcCentralInfluencia = TotalAlcCentralInfluencia + talcinf
          TotalTagCentral = TotalTagCentral + scittag
       ElseIf Tbl1!catipo = "S" Then
          VehSurOrigen = VehSurOrigen + Tbl1!ciorigen + Tbl1!cicredit + Tbl1!ciinfluencia
          TotalTUSurOrigen = TotalTUSurOrigen + Tbl1!citorige + Tbl1!citcredi + Tbl1!citinfluencia
          TotalAlcSurOrigen = TotalAlcSurOrigen + talcori
          TotalAlcSurInfluencia = TotalAlcSurInfluencia + talcinf
          TotalTagSur = TotalTagSur + scittag
       ElseIf Tbl1!catipo = "N" Then
          VehNorteOrigen = VehNorteOrigen + Tbl1!ciorigen + Tbl1!cicredit + Tbl1!ciinfluencia
          TotalTUNorteOrigen = TotalTUNorteOrigen + Tbl1!citorige + Tbl1!citcredi + Tbl1!citinfluencia
          TotalAlcNorteOrigen = TotalAlcNorteOrigen + talcori
          TotalAlcNorteInfluencia = TotalAlcNorteInfluencia + talcinf
          TotalTagNorte = TotalTagNorte + scittag
       End If
       
       
    
      Grdconcie.AddItem entry
      ProBar.Value = ProBar.Value + 1
      total = total + CDbl(Tbl1!citorige) + CDbl(Tbl1!cittrans) + CDbl(Tbl1!citinfluencia) + CDbl(talcori) + CDbl(talcinf) + CDbl(talctra) + _
              CDbl(parque) + CDbl(Tbl1!citcredi) + CDbl(Tbl1!citinfra) + CDbl(scitelusion) + _
              CDbl(scitevasion) + CDbl(scitotrinf) + CDbl(scitevastra) + CDbl(Tbl1!citvigencia) + CDbl(scittag) + CDbl(Tbl1!cittransitonorte)
      
      Tbl1.MoveNext
      bandera = 1
   Loop
   Tbl1.Close
   Concietur.Enabled = True
   ProBar.Value = ProBar.Max
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 16, "Busqueda"
      Exit Sub
   End If
   Msktotal.text = total
   
   Call Aplicar_Perfil(Me, "Concietur")
   Exit Sub
  
terminar:
  Concietur.Enabled = True
  MsgBox "Error al tratar de mostrar los registros o error de base de datos " & Chr(13) + Chr(13) & "ERROR >> " & Err.Description, 48, "Advertencia"

  Exit Sub
  
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo err_mensaje

   If Grdconcie.rows > 1 Then
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      
      If Grdconcie.rows > 2 Then
         lineas = Grdconcie.rows - 1
         Grdconcie.Row = 1
         Grdconcie.Visible = False
         ProBar.Value = ProBar.Min
         ProBar.Max = Grdconcie.rows + 2
         Concietur.Enabled = False
         Dbs.BeginTrans
        
         On Error GoTo Err_Trans
         Do
         
            ReDim dato(1, 48)
                   
            For i = 0 To 47
               Grdconcie.Col = i
               dato(1, i) = Trim(Grdconcie.text)
            Next i
               
            ' insertamos en la tabla
                
            QRY = "INSERT INTO optempo1 " _
                         & "(temdate1, temnumero1, temtext1, temtext2, " _
                         & "temtext3, temtext4, temnumero2, temnumero3, " _
                         & "temnumero4, temnumero5, temnumdob1, temnumdob2, " _
                         & "temnumdob3, temnummon1, temnumdob4, temnummon2, " _
                         & "temnummon3, temnumdob5, temnumero7, temnummon4, " _
                         & "temnummon5, temtext5, tempc, temnumero8, " _
                         & "temnumero6, temnumero9, temnumero10, temnumdob7, " _
                         & "temnumdob6, temnumdob8, temnumero11, temnumero12, " _
                         & "temnumero16, temnumero23, temnumero17, temnumero13, " _
                         & "temnumero18, temnumero24, temnumero14, temnumero25, " _
                         & "temnumero15, temnumdob9, temtext6, temnumero19, " _
                         & "temnumero20, temnumero21, temnumero22, temnumero26, temnumero27) VALUES (TO_DATE('" _
                         & dato(1, 0) & "','MM/dd/yyyy')," & dato(1, 1) & ",'" & dato(1, 2) & "','" & dato(1, 3) & "','" _
                         & dato(1, 4) & "','" & dato(1, 5) & "'," & dato(1, 6) & "," & dato(1, 7) & "," _
                         & dato(1, 8) & "," & dato(1, 9) & "," & dato(1, 10) & "," & dato(1, 11) & "," _
                         & dato(1, 12) & "," & dato(1, 13) & "," & dato(1, 14) & "," & dato(1, 19) & "," _
                         & dato(1, 20) & "," & dato(1, 24) & "," & dato(1, 21) & "," & dato(1, 22) & ", " _
                         & dato(1, 23) & ",  '" & dato(1, 25) & "','" & Pc & "'," & dato(1, 26) & "," _
                         & dato(1, 15) & ",  '" & dato(1, 16) & "','" & dato(1, 17) & "'," & dato(1, 18) & "," _
                         & dato(1, 27) & "," & dato(1, 29) & "," & dato(1, 30) & "," & dato(1, 31) & "," _
                         & dato(1, 32) & "," & dato(1, 33) & "," & dato(1, 34) & "," & dato(1, 35) & "," _
                         & dato(1, 36) & "," & dato(1, 37) & "," & dato(1, 38) & "," & dato(1, 39) & "," _
                         & dato(1, 40) & "," & dato(1, 41) & ",'" & dato(1, 42) & "'," & dato(1, 43) & "," _
                         & dato(1, 44) & "," & dato(1, 45) & "," & dato(1, 46) & "," & dato(1, 47) & "," & dato(1, 30) & ")"
         
            Dbs.Execute QRY
            'fin insertar
                
            lineas = lineas - 1
            ProBar.Value = ProBar.Value + 1
            If lineas <> 0 Then
               Grdconcie.Row = Grdconcie.Row + 1
            End If
         Loop Until lineas = 0  'do_while que se mueve x la grilla
   
         Dbs.CommitTrans
         Grdconcie.Visible = True
         On Error GoTo err_mensaje
         
         imprimir.Show 1
         If G_Imprimir = "0" Then
            Exit Sub
         End If
         
         reporte = RpPath + "\concietu.rpt"
         sFecha = "FECHA INICIAL:  " & Mskfecini.FormattedText & "      FECHA FINAL:  " & Mskfecfin.FormattedText
         rptterminal.ReportFileName = reporte
         sTitulo = "REPORTE DE CIERRES DE TURNO"
         sTerminal = ""
         If CmbTerminal.ListIndex > 0 Then
            'ha escogido un terminal
            If Mid(CmbTerminal.text, 1, 1) = "C" Then
               sTerminal = " TERMINAL CENTRAL"
            ElseIf Mid(CmbTerminal.text, 1, 1) = "S" Then
               sTerminal = " TERMINAL SUR"
            ElseIf Mid(CmbTerminal.text, 1, 1) = "N" Then
               sTerminal = " TERMINAL NORTE"
            End If
          End If
         sTitulo = sTitulo & sTerminal
         rptterminal.WindowTitle = sTitulo
         rptterminal.Connect = G_Conexion_reportes
         rptterminal.Formulas(0) = "total = '" & Msktotal.FormattedText & "'"
         rptterminal.Formulas(1) = "fecini = '" & sFecha & "'"
         rptterminal.Formulas(2) = "empresa = '" & G_NomEmpresa & "'"
         rptterminal.Formulas(3) = "VehCentralOrigen = '" & Format(VehCentralOrigen, "#,##0;(#,##0)") & "'"
         rptterminal.Formulas(4) = "TotalTUCentralOrigen = '" & Format(TotalTUCentralOrigen, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(5) = "TotalAlcCentralOrigen = '" & Format(TotalAlcCentralOrigen, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(6) = "VehSurOrigen = '" & Format(VehSurOrigen, "#,##0;(#,##0)") & "'"
         rptterminal.Formulas(7) = "TotalTUSurOrigen = '" & Format(TotalTUSurOrigen, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(8) = "TotalAlcSurOrigen = '" & Format(TotalAlcSurOrigen, "$#,##0;($#,##0)") & "'"
         
         rptterminal.Formulas(9) = "VehNorteOrigen = '" & Format(VehNorteOrigen, "#,##0;(#,##0)") & "'"
         rptterminal.Formulas(10) = "TotalTUNorteOrigen = '" & Format(TotalTUNorteOrigen, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(11) = "TotalAlcNorteOrigen = '" & Format(TotalAlcNorteOrigen, "$#,##0;($#,##0)") & "'"
         
         rptterminal.Formulas(12) = "TotalAlcCentralInfluencia = '" & Format(TotalAlcCentralInfluencia, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(13) = "TotalAlcSurInfluencia = '" & Format(TotalAlcSurInfluencia, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(14) = "TotalAlcNorteInfluencia = '" & Format(TotalAlcNorteInfluencia, "$#,##0;($#,##0)") & "'"
         
         rptterminal.Formulas(15) = "TotalTagNorte = '" & Format(TotalTagNorte, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(16) = "TotalTagCentral = '" & Format(TotalTagCentral, "$#,##0;($#,##0)") & "'"
         rptterminal.Formulas(17) = "TotalTagSur = '" & Format(TotalTagSur, "$#,##0;($#,##0)") & "'"
         
         'For j = 15 To 47
         For j = 18 To 47
            rptterminal.Formulas(j) = "TotalAlcSurOrigen = '" & Format(TotalAlcSurOrigen, "$#,##0;($#,##0)") & "'"
         Next j
   '     rptterminal.Formulas(1) = "fecini = '" & Mskfecini.FormattedText & "'"
   '     rptterminal.Formulas(2) = "fecfin = '" & Mskfecfin.FormattedText & "'"
        
         Concietur.Enabled = False
         
         rptterminal.PrinterName = G_PrinterName
         rptterminal.PrinterDriver = G_PrinterDriver
         rptterminal.PrinterPort = G_PrinterPort
         
         rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
         rptterminal.WindowState = crptMaximized
         rptterminal.Destination = 0
         rptterminal.Action = 1
         Concietur.Enabled = True
     
      Else
         Call Imprimir_Cierre(Grdconcie.TextMatrix(1, 28), Grdconcie.TextMatrix(1, 25))
      End If
     
      Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      ProBar.Value = ProBar.Max
   End If
  
   Exit Sub
  
Err_Trans:
   Grdconcie.Visible = True
   Concietur.Enabled = True
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
  
err_mensaje:
   Grdconcie.Visible = True
   Concietur.Enabled = True
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
  Optcriterio(0).Value = False
  Optcriterio(1).Value = False
  Txtvalbus.text = ""
  Grdconcie.Cols = 2
  Grdconcie.rows = 1
  Grdconcie.Clear
  Mskfecini.SetFocus
  Optcriterio(2).Value = True
  Msktotal.text = ""
  txtcaja.text = ""
  txtcaja.Enabled = False
  CmbTerminal.ListIndex = 0
  DIB_GRILLA
  ProBar.Value = ProBar.Min
End Sub

Private Sub BtnRodamiento_Click()



    fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      hora = Mid(Ftb01!psval, 12, 8)
   End If
   Ftb01.Close


   L_caja = Fn_Parametros("NOMCAJTAQ", 1) 'En este parametro se define el nombre de la caja en rodamiento
   
   
   
   'Verificamos que la caja actual sea la caja de rodamiento
   
   If L_caja <> Pc Then
        MsgBox "Actualmente no se encuentra en la caja de rodamiento, no es posible realizar el cierre", 48, "Advertencia"
        Exit Sub
   End If
   
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
   
   fecini = Mskfecini.text
   fecfin = Mskfecfin.text
   
   If Not fecini = fecfin Then
        MsgBox "Solo puede contabilizar un día a la vez", 48, "Advertencia"
        Exit Sub
   End If
   
   
   fec = fecini
   
   If fec >= Fecha Then
        MsgBox "Solo se pueden cerrar dias anteriores a hoy", 48, "Advertencia"
        Exit Sub
   End If
   
   'Verificamos si existe un cierre de turno para ese dia
   
   QRY = "SELECT * FROM cocietur WHERE cifecha = TO_DATE('" & fec & "','MM/DD/YYYY')  " _
                 & "  AND cinomcaj = '" & L_caja & "' "
   
   
    Set Ftb01 = Dbs.Execute(QRY)
    
    If Not Ftb01.EOF Then
         MsgBox "Ya existe un cierre para este día", 48, "Advertencia"
        Exit Sub
    End If

    
    fqry01 = "SELECT min(conumero) as minn, max(conumero) as maxn " _
               & "FROM coconduc " _
              & "WHERE cofecsal = TO_DATE('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & L_caja & "' " _
   
     Set Ftb01 = Dbs.Execute(fqry01)
     
    If Ftb01.EOF = False Then
    
        If Not IsNull(Ftb01!minn) And Not IsNull(Ftb01!maxn) Then
           sTiqIni = Ftb01!minn
           sTiqFin = Ftb01!maxn
           sTotalTiq = sTiqFin - sTiqIni + 1
        Else
            MsgBox "No hay tasas para contabilizar", 48, "Advertencia"
            Exit Sub
        End If
    Else
            MsgBox "No hay tasas para contabilizar", 48, "Advertencia"
        Exit Sub
    End If
    
   Ftb01.Close
    
    
    'Consulta para sacar las tasas CUPOS en origen
    
     QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & L_caja & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND cocredit = 'N' " _
            & " AND comodo = 'P' " _
            & " AND ruinflue <> 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              sCantCupo = Ftb01!sCantidadCupo
              sValorCupo = Ftb01!sVrCupoTasa
              sValorCupoAlco = Ftb01!sVrCupoAlco
          Else
              sCantCupo = 0
              sValorCupo = 0
              sValorCupoAlco = 0
          End If
      End If
      Ftb01.Close
      
      'Consulta para tasas CUPOS INFLUENCIA
      
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc INNER JOIN opruta ON cocodrut = rucodigo " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & L_caja & "' " _
            & " AND cooritra = 'O' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND comodo = 'P' " _
            & " AND cocredit = 'N' " _
            & " AND ruinflue = 'Si' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              sCantCupoInflu = Ftb01!sCantidadCupo
              sValorCupoInflu = Ftb01!sVrCupoTasa
              sValorCupoAlcoInflu = Ftb01!sVrCupoAlco
          Else
              sCantCupoInflu = 0
              sValorCupoInflu = 0
              sValorCupoAlcoInflu = 0
          End If
      End If
      Ftb01.Close
      
      'Consulta para tasas CUPOS TRANSITO
      
      QRY = "SELECT count(conumero) as sCantidadCupo, SUM(covalor) AS sVrCupoTasa, SUM(covalalc) AS sVrCupoAlco " _
            & "FROM coconduc " _
           & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
            & " AND conomcaj = '" & L_caja & "' " _
            & " AND cooritra = 'T' " _
            & " AND (coestado = 'Activo' OR coestado = 'Cambio') " _
            & " AND cocredit = 'N' " _
            & " AND comodo = 'P' "

      Set Ftb01 = Dbs.Execute(QRY)
      If Not Ftb01.EOF Then
          If Ftb01!sCantidadCupo <> "" And Ftb01!sCantidadCupo <> "0" Then
              sCantCuposTra = Ftb01!sCantidadCupo
              sValCuposTra = Ftb01!sVrCupoTasa
              sValCuposAlcoTra = CDbl(sValCuposAlcoTra) + Ftb01!sVrCupoAlco
          Else
              sCantCuposTra = 0
              sValCuposTra = 0
          End If
      End If
      Ftb01.Close
    
    
    'Consulta para tasas CUPO - CREDITO
    
    'tiquetes en credito wecc 10/01/2002
      fqry01 = "SELECT count(conumero) as cre,sum(covalor) as valcre " _
               & "FROM coconduc " _
              & "WHERE cofecsal = to_date('" & fec & "','MM/DD/YYYY')  " _
              & "  AND conomcaj = '" & L_caja & "' " _
              & "  AND coestado = 'Activo' " _
              & "  AND cocredit = 'S' " _
              & "  AND comodo = 'P' "
              
              
      Set Ftb01 = Dbs.Execute(fqry01)
      If Ftb01.EOF = False Then
         If Ftb01!cre <> "" And Ftb01!cre <> "0" Then
            sCantCupoCre = Ftb01!cre
            sValCupoCre = Ftb01!valcre
         Else
            sCantCupoCre = 0
            sValCupoCre = 0
         End If
      End If
      Ftb01.Close
    
    
    
    sTotalGeneral = CDbl(sValCuposTra) + CDbl(sValCuposAlcoTra) + CDbl(sValorCupoInflu) + CDbl(sValorCupoAlcoInflu) + CDbl(sValorCupo) + CDbl(sValorCupoAlco) + CDbl(sValCupoCre)
    
    'Cabe recalcar que este es un proceso de cierre de turno exclusivo de RODAMIENTOS.
    'La razón por la que se queman los otros datos es para que la tabla COCIETUR permita la inserción
    On Error GoTo terminar
    
    
    QRY = "INSERT INTO cocietur " _
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
     & fec & "','MM/DD/YYYY'), '1' ,'" & sTiqIni & "','" & sTiqFin & "','" _
     & Login & "',to_date('" & hora & "','HH24:MI:SS'), '0' , 0 ," _
     & " 0 , 0, 0 , 0, 'Cierre de turno rodamiento' ," _
     & sTotalGeneral & ", 0 ," & sTiqIni & "," _
     & " 0, 0, 0, 0, " _
     & " 0, 0, 0,'" & L_caja & "'," _
     & sCantCupoCre & "," & sValCupoCre & ", 0, 0, " _
     & " 0, 0 , 0, 0, " _
     & sCantCupo & ",  " & CDbl(sValorCupo) & "," & CDbl(sValorCupoAlco) & ", 0," _
     & " 0," & sCantCuposTra & ", 0," & sValCuposTra & ", " _
     & " 0, 0, 0, 0," & sCantCupoInflu & "," & sValorCupoInflu & "," & sValorCupoAlcoInflu & ")"
     
      Dbs.Execute (QRY)
    
        MsgBox "El cierre de turno se realizo con éxito", 64, "O.K."
    
    Exit Sub
    
terminar:

    MsgBox "Solo puede contabilizar un día a la vez", 48, "Advertencia"
    Exit Sub
    
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub Btntotalizar_Click()
'  If Grdconcie.Rows > 1 Then
'     total = 0
'     For i = 1 To Grdconcie.Rows - 1
'         Grdconcie.Row = i
'         Grdconcie.Col = 10
'         dato1 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 11
'         dato2 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 14
'         dato3 = CDbl(Grdconcie.Text)
'
'         Grdconcie.Col = 15
'         dato4 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 16
'         dato5 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 17
'         Dato6 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 18
'         Dato7 = CDbl(Grdconcie.Text)
'
'         Grdconcie.Col = 20
'         Dato8 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 22
'         dato9 = CDbl(Grdconcie.Text)
'         Grdconcie.Col = 23
'         dato10 = CDbl(Grdconcie.Text)
'
'         Grdconcie.Col = 27
'         dato11 = CDbl(Grdconcie.Text)
'         total = total + dato1 + dato2 + dato3 + dato4 + dato5 + Dato6 + Dato7 + Dato8 + dato9 + dato10 + dato11
'     Next i
'     Msktotal.Text = total
'  End If
End Sub



Private Sub Form_Activate()
  Call Aplicar_Perfil(Me, "Concietur")
End Sub

Private Sub Form_Load()
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
   
   
   L_usuario = Fn_Parametros("USURODA", 1) 'Aqui tomamos el usuario que tiene permitido el cierre de dia en rodamiento
   
   If L_usuario = Login Then
   
      BtnRodamiento.Enabled = True
   
   Else
   
      BtnRodamiento.Enabled = False
   
   End If
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "'"
End Sub

Private Sub Mskfecfin_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskfecini_GotFocus()
   Senalar_Texto Me
End Sub



Private Sub optcriterio_Click(index As Integer)
  Txtvalbus.text = ""
  txtcaja.text = ""
  If Optcriterio(0).Value = True Then
     txtcaja.Enabled = True
  Else
     txtcaja.Enabled = False
  End If
End Sub

Private Sub Txtcaja_DblClick()
    forma = "12"
    forma1 = "2"
    txtcaja.text = ""
    buscar.Show 1
    
End Sub

Private Sub txtcaja_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
    
   If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
      KeyAscii% = 0
   End If
End Sub

Private Sub Txtvalbus_KeyPress(KeyAscii As Integer)
  If Optcriterio(0).Value = True Then
     If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
  
  If Optcriterio(1).Value = True Then
     Char = Chr(KeyAscii%)
     KeyAscii% = Asc(LCase(Char))
     If (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% = 8) Then
        Exit Sub
     Else
        KeyAscii% = 0
     End If
  End If
  If Optcriterio(3).Value = True Then
      Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))  ' mayuscula
      
      If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
         KeyAscii% = 0
      End If

  End If
  If Optcriterio(0).Value = False And Optcriterio(1).Value = False And Optcriterio(3).Enabled = False Then
     KeyAscii% = 0
  End If
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

Private Function DIB_GRILLA()
   Grdconcie.rows = 1
   Grdconcie.Cols = 48
   Grdconcie.Clear

   s$ = "Fecha|Turno|Núm. Inicial de Tiquete|Núm. Final de Tiquete|" _
      & "Cód. Operario|Hora Cierre|Tiquetes Manuales|Tiquetes Anulados |" _
      & "Por Origen|Por Transito|Total Origen|Total Transito|" _
      & "Total Anulados|Cant. Infracciones|Total Otras Inf.|Total Elusión TUT|" _
      & "Total Evasion TUT|Total Perdida|Total Eva. Trans.|Cant. Alcoh Origen|" _
      & "Total Alc. Origen|Cant. Alcoh Transito|Total Alc. Transito|Permanencia| " _
      & "Total General|Caja|Crédito|Total Crédito|" _
      & "Planilla|Vigencia|TAG|Cant. Cupos Origen|" _
      & "Cant. Cupos Transito|Cant. Cupos Influencia|Cant. Cupos Tran Aut|Cupo Tasa Origen|" _
      & "Cupo Tasa Transito|Cupo Tasa Influencia|Cupo Alcohol Origen|Cupo Alcohol Influencia|" _
      & "Por Trans.Norte|Total Trans.Norte|Tipo|Por Influencia|" _
      & "Total Influencia|Cant. Alcoh Influencia|Total Alc. Influencia|Cant. Alc. Cupos Influencia"
   Grdconcie.FormatString = s$
   Grdconcie.Row = 0
   Grdconcie.Col = 0
   Grdconcie.ColWidth(0) = 1000
   Grdconcie.ColWidth(1) = 1000
   Grdconcie.ColWidth(2) = 2100
   Grdconcie.ColWidth(3) = 2100
   Grdconcie.ColWidth(4) = 1500
   Grdconcie.ColWidth(5) = 1200
   Grdconcie.ColWidth(6) = 1700
   Grdconcie.ColWidth(7) = 1700
   Grdconcie.ColWidth(8) = 1500
   Grdconcie.ColWidth(9) = 1500
   Grdconcie.ColWidth(10) = 1500
   Grdconcie.ColWidth(11) = 1500
   Grdconcie.ColWidth(12) = 1500
   Grdconcie.ColWidth(13) = 1700
   Grdconcie.ColWidth(14) = 1700
   Grdconcie.ColWidth(15) = 1700
   Grdconcie.ColWidth(16) = 1700
   Grdconcie.ColWidth(17) = 1700
   Grdconcie.ColWidth(18) = 1700
   Grdconcie.ColWidth(19) = 1700
   Grdconcie.ColWidth(20) = 1700
   Grdconcie.ColWidth(21) = 1900
   Grdconcie.ColWidth(22) = 1700
   Grdconcie.ColWidth(23) = 1700
   Grdconcie.ColWidth(24) = 1700
   Grdconcie.ColWidth(25) = 1700
   Grdconcie.ColWidth(26) = 1700
   Grdconcie.ColWidth(27) = 1700
   Grdconcie.ColWidth(28) = 0
   Grdconcie.ColWidth(29) = 1700
   Grdconcie.ColWidth(30) = 1500
   Grdconcie.ColWidth(31) = 1900
   Grdconcie.ColWidth(32) = 1900
   Grdconcie.ColWidth(33) = 2200
   Grdconcie.ColWidth(34) = 2000
   Grdconcie.ColWidth(35) = 1700
   Grdconcie.ColWidth(36) = 1900
   Grdconcie.ColWidth(37) = 2000
   Grdconcie.ColWidth(38) = 1900
   Grdconcie.ColWidth(39) = 2200
   Grdconcie.ColWidth(40) = 1700
   Grdconcie.ColWidth(41) = 1700
   Grdconcie.ColWidth(42) = 0
   Grdconcie.ColWidth(43) = 1700
   Grdconcie.ColWidth(44) = 1700
   Grdconcie.ColWidth(45) = 2100
   Grdconcie.ColWidth(46) = 1900
   Grdconcie.ColWidth(47) = 2300
      
   For i = o To Grdconcie.Cols - 1
      Grdconcie.Col = i
      Grdconcie.CellFontBold = 1
   Next i
  
End Function
