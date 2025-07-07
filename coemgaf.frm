VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form coemgaf 
   Caption         =   "Consulta Registros EMGAF"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
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
      Picture         =   "coemgaf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4320
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
      Picture         =   "coemgaf.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3480
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
      Picture         =   "coemgaf.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   2640
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
      Picture         =   "coemgaf.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   1800
      Width           =   855
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
      TabIndex        =   16
      Top             =   4200
      Width           =   6495
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
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
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
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
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
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   6495
      Begin VB.TextBox txtCodEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtCedula 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   1335
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
         TabIndex        =   14
         Top             =   360
         Width           =   855
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
         TabIndex        =   13
         Top             =   720
         Width           =   1215
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
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   4575
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
         TabIndex        =   11
         Top             =   720
         Width           =   3855
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
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
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
         TabIndex        =   7
         Top             =   360
         Width           =   1575
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
         TabIndex        =   6
         Top             =   720
         Width           =   1215
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   240
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
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
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
         Left            =   240
         TabIndex        =   2
         Top             =   960
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdEmgaf 
      Height          =   2775
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport rptterminal 
      Left            =   7080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      DiscardSavedData=   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   7800
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6720
      X2              =   6720
      Y1              =   0
      Y2              =   5160
   End
End
Attribute VB_Name = "coemgaf"
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
    
    'Limpiamos la Grilla
    DIB_GRILLA
    
    'Declaramos variables de consulta
    Fecha = Mskfecini.text
    Fecha1 = Mskfecfin.text
    
    'Consulta
    QRY = "SELECT emnumero, emnumcon, emcedula, hecodigo, henombre, emfecreg, emconmed, emtension, emfrecar, emimc, emfatiga, emgluco, " _
            & "emaudio, emvisual, NVL(TO_CHAR(emfecinicer, 'MM/DD/YYYY'), 'N/A') AS emfecinicer, NVL(TO_CHAR(emfecfincer, 'MM/DD/YYYY'), 'N/A') AS emfecfincer " _
            & "FROM opemgaf, coconductor, ophojemp " _
            & "WHERE emfecreg BETWEEN TO_DATE('" & Fecha & "','MM/DD/YYYY') " _
            & "AND TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
            & "AND emcedula = cocedula " _
            & "AND cocodemp = hecodigo "
            
    'Si buscan empresa
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND cocodemp = '" & txtCodEmp.text & "' "
    End If
    
    'Si buscan conductor especifico
    If chkCedula.Value = 1 Then
        QRY = QRY & "AND emcedula = '" & txtCedula.text & "' "
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
        
        Do While Not Tbl1.EOF
            entry = CStr(Tbl1!emnumero) + Chr(9) + Tbl1!emnumcon + Chr(9) + Tbl1!emcedula + Chr(9) + Tbl1!hecodigo + Chr(9) + _
                Tbl1!henombre + Chr(9) + CStr(Tbl1!emfecreg) + Chr(9) + Tbl1!emconmed + Chr(9) + Tbl1!emtension + Chr(9) + _
                Tbl1!emfrecar + Chr(9) + Tbl1!emimc + Chr(9) + Tbl1!emfatiga + Chr(9) + Tbl1!emgluco + Chr(9) + _
                Tbl1!emaudio + Chr(9) + Tbl1!emvisual + Chr(9) + CStr(Tbl1!emfecinicer) + Chr(9) + CStr(Tbl1!emfecfincer)
                
            grdEmgaf.AddItem (entry)
            Tbl1.MoveNext
            
            TotalRegistros = TotalRegistros + 1
            ProBar.Value = ProBar.Value + 1
        Loop
        Tbl1.Close
        
        lblTotReg.Caption = TotalRegistros
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros", "48", "Advertencia"
    End If
    
    Exit Sub
    
Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
    
    
End Sub

Private Sub Btnimprimir_Click()
        
    On Error GoTo err_mensaje
        
        If grdEmgaf.rows > 1 Then
            
            lineas = grdEmgaf.rows - 1
            grdEmgaf.Row = 1
            grdEmgaf.Visible = False
            ProBar.Value = ProBar.Min
            ProBar.Max = grdEmgaf.rows + 2
            Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
      
            Dbs.BeginTrans
            On Error GoTo Err_Trans
            Do
                ReDim dato(1, 16)
                
                For i = 0 To 13
                    grdEmgaf.Col = i
                    dato(1, i) = Trim(grdEmgaf.text)
                Next i
                
                'Inserta en la tabla temporal
                QRY = "INSERT INTO optempo1 " _
                                & "(temnumdob2, temtext2, temtext3, temtext4, " _
                                & " temtext5, temtext6, temtext7, " _
                                & " temtext8, temtext9, temtext10, temtext11, temtext12, temtext13, temtext14, temtext15, temtext16, tempc ) VALUES (" _
                                & dato(1, 0) & ",'" & dato(1, 1) & "','" & dato(1, 2) & "','" & dato(1, 3) & "','" _
                                & dato(1, 4) & "','" & dato(1, 5) & "','" & dato(1, 6) & "','" & dato(1, 7) & "','" _
                                & dato(1, 8) & "','" & dato(1, 9) & "','" & dato(1, 10) & "','" & dato(1, 11) & "', '" _
                                & dato(1, 12) & "', '" & dato(1, 13) & "', '" & dato(1, 14) & "', '" & dato(1, 15) & "', '" & Pc & "') "
                                
                Dbs.Execute QRY
                lineas = lineas - 1
                ProBar.Value = ProBar.Value + 1
                If lineas <> 0 Then
                    grdEmgaf.Row = grdEmgaf.Row + 1
                End If
                
            Loop Until lineas = 0  'do_while que se mueve x la grilla
    
             Dbs.CommitTrans
             grdEmgaf.Visible = True
             On Error GoTo err_mensaje
              respuesta = MsgBox("Desea escoger la impresora (SI) o impresora predeterminada (NO)", vbYesNo + 32, "Impresoras")
            
             If respuesta = 6 Then
                imprimir.Show 1
                If G_Imprimir = "0" Then
                   Exit Sub
                End If
             End If
             
             reporte = RpPath + "\coemgaf.rpt"
             
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
             coemgaf.Enabled = True
            
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
txtCodEmp.text = ""
lblNomEmpresa.Caption = ""
txtCedula.text = ""
lblNomConductor.Caption = ""
lblTotReg.Caption = ""
DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    'Creamos la grilla
    DIB_GRILLA
    
    Mskfecini_GotFocus

End Sub

Private Function DIB_GRILLA()

   grdEmgaf.rows = 1
   grdEmgaf.Cols = 16
   grdEmgaf.Clear
   
   's$ = "Cód. Empresa  |Nombre   |Cant.Origen |Tot. Origen |Cant.Transito |Tot. Transito |Vr.Alc. Origen |Vr.Alc. Transito |Parqueadero |Elusion |Evasion |Evasion Transito |Per. Documento |Total Multas |Vigencias |Tasas de Uso Vendidos  |Total General |Saldo Inicial | Saldo Final| Consignacion|"
   s$ = "Consecutivo |Num. Consulta   |Cedula |Cod. Empresa |Nom. Empresa |Fec. Registro |Cons. Medica |Tension |Frec. Cardiaca |IMC |Fatiga |Glucometria |Audio |Visual  |Fec. Inicio Certificado |Fec. Fin Certificado"
   grdEmgaf.FormatString = s$
   grdEmgaf.Row = 0
   grdEmgaf.Col = 0
   grdEmgaf.ColWidth(0) = 1400
   grdEmgaf.ColWidth(1) = 1400
   grdEmgaf.ColWidth(2) = 1400
   grdEmgaf.ColWidth(3) = 2200
   grdEmgaf.ColWidth(4) = 1600
   grdEmgaf.ColWidth(5) = 1600
   grdEmgaf.ColWidth(6) = 1600
   grdEmgaf.ColWidth(7) = 1800
   grdEmgaf.ColWidth(8) = 1400
   grdEmgaf.ColWidth(9) = 2200
   grdEmgaf.ColWidth(10) = 1400
   grdEmgaf.ColWidth(11) = 1400
   grdEmgaf.ColWidth(12) = 1400
   grdEmgaf.ColWidth(13) = 1400
   grdEmgaf.ColWidth(14) = 2000
   grdEmgaf.ColWidth(15) = 2000
   
   For i = 0 To grdEmgaf.Cols - 1
       grdEmgaf.Col = i
       grdEmgaf.CellFontBold = 1
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

Private Sub txtcodemp_DblClick()
  forma = "16"
  forma1 = "69"
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
    
    If txtCodEmp.text <> "" Then
        
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
                txtcodemp_DblClick
            End If
            Tbl2.Close
        End If
    End If
    
End Sub
Private Sub txtcedula_DblClick()
    forma = "20"
    forma1 = "8"
    lblNomConductor.Caption = ""
    buscar.Show 1
    txtcedula_LostFocus
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
