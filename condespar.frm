VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form condespar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Descuento de Parqueaderos"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   29
      Top             =   120
      Width           =   7695
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Fecha en la que empiece a buscar"
         Top             =   360
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
         Left            =   5880
         TabIndex        =   1
         ToolTipText     =   "Fecha hasta donde quiere que busque"
         Top             =   360
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
         Left            =   4560
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   360
         Width           =   1335
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
      Left            =   2520
      TabIndex        =   22
      Top             =   1080
      Width           =   5295
      Begin VB.ComboBox Cmbestado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "condespar.frx":0000
         Left            =   1680
         List            =   "condespar.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Estado del registro"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox TxtTipo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   7
         ToolTipText     =   "Código del tipo de descuento (Precione doble click para ayuda)"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   3
         ToolTipText     =   "Número de la placa del vehículo (Precione doble click para ayuda)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Txtcodemp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   5
         ToolTipText     =   "Código de la empresa (Precione doble click para ayuda)"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Estado:"
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
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LblTipo 
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
         Left            =   2520
         TabIndex        =   28
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label LblPlaca 
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
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   2655
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
         Left            =   2520
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1335
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
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Descuento:"
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
         Top             =   960
         Width           =   1455
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
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   2295
      Begin VB.CheckBox Chkcriterio 
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Chkcriterio 
         Caption         =   "Tipo Descuento"
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
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Chkcriterio 
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
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
      Left            =   8040
      Picture         =   "condespar.frx":002F
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   4080
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
      Left            =   8040
      Picture         =   "condespar.frx":0471
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Left            =   8040
      Picture         =   "condespar.frx":08B3
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir de la forma actual"
      Top             =   5760
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
      Left            =   8040
      Picture         =   "condespar.frx":0CF5
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar informacion en la base de datos"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   7695
      Begin Crystal.CrystalReport rptterminal 
         Left            =   1920
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDescuento 
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   7695
      Begin MSMask.MaskEdBox Msktotal 
         Height          =   300
         Left            =   5400
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDescuento 
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Valor Cancelado:"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Descuentos:"
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
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   32
      Top             =   6360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   7920
      X2              =   7920
      Y1              =   6480
      Y2              =   0
   End
End
Attribute VB_Name = "condespar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DIB_GRILLA()
     GrdDescuento.rows = 1
     GrdDescuento.Cols = 11
     GrdDescuento.Clear
     
     s$ = "Número|Placa  |Empresa|Fecha Ingreso |Fecha Salida |Fecha Pago |Valor a Pagar|" _
        & "Vr Total Descuento| Tipo Descuento|Vr Descuento Tipo|Estado|Usuario"
        
     GrdDescuento.FormatString = s$
     GrdDescuento.Row = 0
     GrdDescuento.Col = 0
     GrdDescuento.ColWidth(0) = 1000
     GrdDescuento.ColWidth(1) = 1000
     GrdDescuento.ColWidth(2) = 1800
     GrdDescuento.ColWidth(3) = 1600
     GrdDescuento.ColWidth(4) = 1600
     GrdDescuento.ColWidth(5) = 1600
     GrdDescuento.ColWidth(6) = 1300
     GrdDescuento.ColWidth(7) = 1800
     GrdDescuento.ColWidth(8) = 2000
     GrdDescuento.ColWidth(9) = 1800
     GrdDescuento.ColWidth(10) = 1300
     GrdDescuento.ColWidth(11) = 1300
     
     For i = 0 To GrdDescuento.Cols - 1
         GrdDescuento.Col = i
         GrdDescuento.CellFontBold = 1
         GrdDescuento.MergeCol(i) = True
     Next i
End Sub

Private Function ImpReporte()
     
   reporte = RpPath + "\condespa.rpt"
   rptterminal.ReportFileName = reporte
   rptterminal.Connect = G_Conexion_reportes
   rptterminal.Formulas(0) = "empresa = '" & G_NomEmpresa & "'"
   rptterminal.Formulas(1) = "inicio = '" & Mskfecini.FormattedText & "'"
   rptterminal.Formulas(2) = "final = '" & Mskfecfin.FormattedText & "'"
   rptterminal.Formulas(3) = "total = '" & Msktotal.FormattedText & "'"
   rptterminal.Formulas(4) = "totdescuento = '" & MskDescuento.FormattedText & "'"
   condespar.Enabled = False
   rptterminal.WindowState = crptMaximized
   rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
   rptterminal.Destination = 0
   rptterminal.Action = 1
   condespar.Enabled = True
   
End Function


Private Sub Btnbuscar_Click()
   On Error GoTo Errores
   Dim L_DirCristal As String
   
   Msktotal.text = "0"
   MskDescuento.text = "0"
   L_DirCristal = "N"
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
        
   If Chkcriterio(0).Value = 1 Then
      If TxtPlaca.text = "" Then
         MsgBox "Falta ingresar el código de la placa", 48, "Advertencia"
         TxtPlaca.SetFocus
         Exit Sub
      End If
   End If
        
   If Chkcriterio(1).Value = 1 Then
      If Txtcodemp.text = "" Then
         MsgBox "Falta ingresar el código de la empresa", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(2).Value = 1 Then
      If TxtTipo.text = "" Then
         MsgBox "Falta ingresar el tipo de descuento", 48, "Advertencia"
         TxtTipo.SetFocus
         Exit Sub
      End If
   End If
   
   If Chkcriterio(3).Value = 1 Then
      If Cmbestado.ListIndex = -1 Then
         MsgBox "Falta seleccionar el estado", 48, "Advertencia"
         Cmbestado.SetFocus
         Exit Sub
      End If
   End If
   
   DIB_GRILLA
   
   QRY = "SELECT paplaca, henombre, pafecini, pafecsal, pafecpag, " _
              & "pavalor, padescuento, paestado, dpcoddes, dpvalor, panumero, PAFUNDES " _
         & "FROM ophojemp inner join " _
         & "(coparqueop inner join opdesparque2 on coparqueop.panumero = opdesparque2.dpconpar) " _
         & "on ophojemp.hecodigo = coparqueop.pacodemp " _
        & "WHERE pafecdes BETWEEN TO_DATE('" & Mskfecini.text & " 00:00', 'MM/dd/yyyy HH24:MI') " _
         & "AND TO_DATE('" & Mskfecfin.text & " 23:59', 'MM/dd/yyyy HH24:MI') "
   
   If Trim(TxtPlaca.text) <> "" Then
      QRY = QRY & " AND paplaca = '" & TxtPlaca.text & "' "
   End If
   
   If Trim(Txtcodemp.text) <> "" Then
      QRY = QRY & " AND pacodemp = '" & Txtcodemp.text & "' "
   End If
   If Trim(TxtTipo.text) <> "" Then
      QRY = QRY & " AND dpcoddes = '" & TxtTipo.text & "' "
   End If
   
   If Trim(Cmbestado.text) <> "" Then
      QRY = QRY & " AND paestado = '" & Mid(Cmbestado.text, 1, 1) & "' "
   End If
   
   QRY = QRY & " ORDER BY henombre "
   
   Set Tbl1 = Dbs.Execute(QRY)
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
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
   
         
   bandera = 0
   sValDescuento = 0
   ProBar.Value = ProBar.Min
   condespar.Enabled = False
   Do While Not Tbl1.EOF
      DoEvents
      sDescripcion = Traer_Descripcion(Tbl1!dpcoddes, "TIPOS DESCUENTOS PARQUEADEROS")
      sFecPag = ""
      If Not IsNull(Tbl1!pafecpag) Then
         sFecPag = Format(Tbl1!pafecpag, "MMM/dd/yyyy")
      End If
      
      If Tbl1!paestado = "A" Then
         sEstado = "Activa"
      ElseIf Tbl1!paestado = "C" Then
         sEstado = "Cancelada"
      End If
      
      sDescuento = 0
      If Not IsNull(Tbl1!padescuento) Then
         sDescuento = Tbl1!padescuento
      End If
      
      Dim funcionario
      If Not IsNull(Tbl1!PAFUNDES) Then
        funcionario = Tbl1!PAFUNDES
      Else
        funcionario = ""
      End If
      
      If L_DirCristal = "N" Then
         entry = CStr(Tbl1!panumero) + Chr(9) + Tbl1!paplaca + Chr(9) + Tbl1!henombre + Chr(9) + Format(Tbl1!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
                 Format(Tbl1!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + sFecPag + Chr(9) + CStr(Tbl1!pavalor) + Chr(9) + _
                 CStr(sDescuento) + Chr(9) + sDescripcion + Chr(9) + CStr(Tbl1!dpvalor) + Chr(9) + sEstado + Chr(9) + funcionario
      
         GrdDescuento.AddItem entry
      End If
      sPlaca = Tbl1!paplaca
      sEmpresa = Tbl1!henombre
      sFecIni = Format(Tbl1!pafecini, "MMM/dd/yyyy hh:mm")
      sFecSal = Format(Tbl1!pafecsal, "MMM/dd/yyyy hh:mm")
      sValor = Tbl1!pavalor
      sTipoDescuento = sDescripcion
      sVrDescuento = Tbl1!dpvalor
      
      ' insertamos en la tabla
           
      QRY = "INSERT INTO optempo1 " _
      & "(temtext1,temtext2, temtext3, temtext4, " _
      & "temtext5, temnumdob1, temnumdob3, temtext7, " _
      & "temnumdob2, temtext8, tempc, temnumero2, TEMTEXT10 ) VALUES ('" _
      & sPlaca & "' ,'" & sEmpresa & "','" & sFecIni & "','" & sFecSal & "','" _
      & sFecPag & "'," & sValor & " ," & sDescuento & ",'" & sTipoDescuento & "'," _
      & sVrDescuento & ",'" & sEstado & "','" & Pc & "'," & Tbl1!panumero & ", '" & funcionario & "')"
      
      Dbs.Execute QRY
      
      sValDescuento = sValDescuento + sDescuento
      sValActiva = sValActiva + Tbl1!pavalor
      
      bandera = 1
      ProBar.Value = ProBar.Value + 1
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   MskDescuento.text = sValDescuento
   Msktotal.text = sValActiva
   condespar.Enabled = True
   If bandera = 0 Then
      MsgBox "No se encontro ningún registro activo", 64, "Busqueda"
      Exit Sub
   End If
   If L_DirCristal = "S" Then
      ImpReporte
   End If
   ProBar.Value = ProBar.Max
   Call Aplicar_Perfil(Me, "condespar")
   Exit Sub
Errores:
   condespar.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo Errores
   
'   ProBar.Value = ProBar.Min
'   ProBar.Max = GrdDescuento.Rows + 2
'
'   If GrdDescuento.Rows > 1 Then
'      Dbs.Execute "DELETE FROM optempo1 WHERE  tempc = '" & Pc & "' "
'
'
'      Dbs.BeginTrans
'      On Error GoTo Err_Trans
'      For i = 1 To GrdDescuento.Rows - 1
'
'         sPlaca = GrdDescuento.TextMatrix(i, 0)
'         sEmpresa = GrdDescuento.TextMatrix(i, 1)
'         sFecIni = GrdDescuento.TextMatrix(i, 2)
'         sFecSal = GrdDescuento.TextMatrix(i, 3)
'         sFecPag = GrdDescuento.TextMatrix(i, 4)
'         sValor = GrdDescuento.TextMatrix(i, 5)
'         sDescuento = GrdDescuento.TextMatrix(i, 6)
'         sTipoDescuento = GrdDescuento.TextMatrix(i, 7)
'         sVrDescuento = GrdDescuento.TextMatrix(i, 8)
'         sEstado = GrdDescuento.TextMatrix(i, 9)
'
'        ' insertamos en la tabla
'
'         qry = "INSERT INTO optempo1 " _
'         & "(temtext1,temtext2, temtext3, temtext4, " _
'         & "temtext5, temnumdob1, temnumdob3, temtext7, " _
'         & "temnumdob2, temtext8, tempc ) VALUES ('" _
'         & sPlaca & "' ,'" & sEmpresa & "','" & sFecIni & "','" & sFecSal & "','" _
'         & sFecPag & "'," & sValor & " ," & sDescuento & ",'" & sTipoDescuento & "'," _
'         & sVrDescuento & ",'" & sEstado & "','" & Pc & "')"
'
'         Dbs.Execute qry
'         ' fin insertar
'
'         ProBar.Value = ProBar.Value + 1
'
'      Next i
'
'      Dbs.CommitTrans
'      On Error GoTo Errores
'      ProBar.Value = ProBar.Max
       
      Call ImpReporte
      
'   End If
   
   Exit Sub
Err_Trans:
   Dbs.RollbackTrans
Errores:
   condespar.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   
   Mskfecini.Mask = ""
   Mskfecini.text = ""
   Mskfecini.Mask = "##/##/####"
   Mskfecfin.Mask = ""
   Mskfecfin.text = ""
   Mskfecfin.Mask = "##/##/####"
   
   For i = 0 To 2
      Chkcriterio(i).Value = 0
   Next i
   MskDescuento.text = ""
   Msktotal.text = ""
   ProBar.Value = ProBar.Min
   Mskfecini.SetFocus
   DIB_GRILLA

   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Chkcriterio_Click(index As Integer)
   On Error GoTo Errores
   
   If Chkcriterio(index).Value = 1 Then
      If index = 0 Then
         TxtPlaca.Enabled = True
      ElseIf index = 1 Then
         Txtcodemp.Enabled = True
      ElseIf index = 2 Then
         TxtTipo.Enabled = True
      ElseIf index = 3 Then
         Cmbestado.Enabled = True
      End If
      
   ElseIf Chkcriterio(index).Value = 0 Then
      
      If index = 0 Then
         TxtPlaca.Enabled = False
         TxtPlaca.text = ""
         LblPlaca.Caption = ""
      ElseIf index = 1 Then
         Txtcodemp.Enabled = False
         Txtcodemp.text = ""
         Lblcodemp.Caption = ""
      ElseIf index = 2 Then
         TxtTipo.Enabled = True
         TxtTipo.text = ""
         LblTipo.Caption = ""
      ElseIf index = 3 Then
         Cmbestado.Enabled = True
         Cmbestado.ListIndex = -1
      End If
      
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub

End Sub

Private Sub Form_Activate()
   Call Aplicar_Perfil(Me, "condespar")
End Sub

Private Sub Form_Load()
   DIB_GRILLA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
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
      
      If Mskfecini.ClipText <> "" Then
         If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
            MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
            Mskfecini.SetFocus
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

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "44"
  Txtcodemp.text = ""
  Lblcodemp.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   On Error GoTo Errores
   If Txtcodemp.text <> "" Then
      fqry01 = "SELECT henombre " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & Txtcodemp.text & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
         Txtcodemp.SetFocus
         Exit Sub
      End If
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub TxtPlaca_DblClick()
   forma = "2"
   forma1 = "16"
   TxtPlaca.text = ""
   LblPlaca.Caption = ""
   buscar.Show 1
   TxtPlaca_LostFocus
End Sub

Private Sub TxtPlaca_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub TxtPlaca_LostFocus()
   On Error GoTo Errores
   
   If Trim(TxtPlaca.text) <> "" Then
      fqry01 = "SELECT hvconact " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & TxtPlaca.text & "' "
       
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         LblPlaca.Caption = Ftb01!hvconact
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Placa De Vehiculo No Existe", 48, "Advertencia"
         TxtPlaca.SetFocus
         Exit Sub
      End If
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub TxtTipo_DblClick()
   forma = "15"
   forma1 = "38"
   Tabla_Catalogo = "TIPOS DESCUENTOS PARQUEADEROS"
   TxtTipo.text = ""
   buscar.Show 1
   TxtTipo.SetFocus
   TxtTipo_LostFocus
End Sub

Private Sub TxtTipo_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtTipo_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtTipo_LostFocus()
   On Error GoTo Errores
   If TxtTipo.text <> "" Then
      fqry01 = "SELECT dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscoddet = '" & TxtTipo.text & "' " _
                & "AND dscodtip = stcodtip " _
                & "AND stdes = 'TIPOS DESCUENTOS PARQUEADEROS' "
                
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         LblTipo.Caption = Ftb01!dsdes
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Tipo de descuento de parqueadero no existe", 48, "Advertencia"
         TxtTipo.SetFocus
      End If
   End If

   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub
