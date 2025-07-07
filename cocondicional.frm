VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form cocondicional 
   Caption         =   "Reporte Tasas de Uso Condicional"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8940
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
      Left            =   7920
      Picture         =   "cocondicional.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salir de la forma actual"
      Top             =   3480
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
      Left            =   7920
      Picture         =   "cocondicional.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Limpia la forma actual"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Btnimprimir 
      Caption         =   "&Imprimir"
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
      Left            =   7920
      Picture         =   "cocondicional.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   960
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
      Left            =   7920
      Picture         =   "cocondicional.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   120
      Width           =   855
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
      Left            =   7920
      Picture         =   "cocondicional.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprimir Documento Soporte"
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
      TabIndex        =   18
      Top             =   3360
      Width           =   7575
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   7575
      Begin VB.TextBox txtCodEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   12
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
         TabIndex        =   17
         Top             =   360
         Width           =   855
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
         TabIndex        =   16
         Top             =   360
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
         TabIndex        =   15
         Top             =   720
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
         TabIndex        =   14
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Crit Busqueda"
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
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
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
         TabIndex        =   10
         Top             =   240
         Width           =   1095
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
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
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
      Begin VB.ComboBox cmbTerminal 
         Height          =   315
         Left            =   120
         TabIndex        =   7
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
      Top             =   0
      Width           =   2895
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   120
         TabIndex        =   1
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
         TabIndex        =   4
         Top             =   360
         Width           =   1215
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
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConcond 
      Height          =   1935
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   8880
      X2              =   0
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7800
      X2              =   7800
      Y1              =   0
      Y2              =   8520
   End
End
Attribute VB_Name = "cocondicional"
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
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de la placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    
    'Limpiamos la Grilla
    DIB_GRILLA
    
    'Declaramos variables de consulta
    Fecha = Mskfecini.text
    Fecha1 = Mskfecfin.text
    
    'Hacemos la consulta
    QRY = "SELECT cdnumero, cdplaca, cdcodemp, NVL(henombre, 'N/A') AS henombre, cdfecha, " _
            & "TO_CHAR(cdhora, 'HH24:MI:SS') AS cdhora, " _
            & "CASE cdterminal WHEN 'C' THEN 'CENTRAL' " _
            & "WHEN 'N' THEN 'NORTE' " _
            & "ELSE 'SUR' END AS cdterminal, " _
            & "NVL(cdtasauso, 0) AS cdtasauso, NVL(cdcoddes, 'N/A') AS cdcoddes, " _
            & "NVL(rudestino, 'N/A') AS rudestino " _
            & "FROM vehconcond " _
            & "LEFT JOIN opruta " _
            & "ON vehconcond.cdcoddes = opruta.rucodigo " _
            & "LEFT JOIN ophojemp " _
            & "ON vehconcond.cdcodemp = ophojemp.hecodigo " _
            & "WHERE cdfecha BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
            & "AND TO_DATE('" & Fecha1 & "','MM/DD/YYYY' ) "

    If chkPlaca.Value = 1 Then
        QRY = QRY & "AND cdplaca = '" & txtPlaca.text & "' "
    End If
    
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND cdcodemp = '" & txtCodEmp.text & "' "
    End If
    
    If cmbTerminal.ListIndex <> 0 Then
        QRY = QRY & "AND cdterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' "
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
    
    'Si no es vacio agregamos los registros a la grilla
    If Not Tbl1.EOF Then
        
        Do While Not Tbl1.EOF
        
            entry = CStr(Tbl1!cdnumero) + Chr(9) + Tbl1!cdplaca + Chr(9) + Tbl1!cdcodemp + Chr(9) + Tbl1!henombre + Chr(9) + _
                    CStr(Tbl1!cdfecha) + Chr(9) + CStr(Tbl1!cdhora) + Chr(9) + Tbl1!cdterminal + Chr(9) + _
                    CStr(Tbl1!cdtasauso) + Chr(9) + Tbl1!cdcoddes + Chr(9) + Tbl1!rudestino
                    
        
            grdConcond.AddItem (entry)
            Tbl1.MoveNext
            
            TotalRegistros = TotalRegistros + 1
            ProBar.Value = ProBar.Value + 1
        Loop
        
        Tbl1.Close
        
        lblTotReg.Caption = TotalRegistros
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros", 48, "Advertencia"
    End If
Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error: " & Err.Number
    
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
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de la placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    
    'Limpiamos la Grilla
    DIB_GRILLA
    
    'Declaramos variables de consulta
    Fecha = Mskfecini.text
    Fecha1 = Mskfecfin.text
    
    'Hacemos la consulta
    QRY = "SELECT cdnumero, cdplaca, cdcodemp, NVL(henombre, 'N/A') AS henombre, cdfecha, " _
            & "TO_CHAR(cdhora, 'HH24:MI:SS') AS cdhora, " _
            & "CASE cdterminal WHEN 'C' THEN 'CENTRAL' " _
            & "WHEN 'N' THEN 'NORTE' " _
            & "ELSE 'SUR' END AS cdterminal, " _
            & "NVL(cdtasauso, 0) AS cdtasauso, NVL(cdcoddes, 'N/A') AS cdcoddes, " _
            & "NVL(rudestino, 'N/A') AS rudestino " _
            & "FROM vehconcond " _
            & "LEFT JOIN opruta " _
            & "ON vehconcond.cdcoddes = opruta.rucodigo " _
            & "LEFT JOIN ophojemp " _
            & "ON vehconcond.cdcodemp = ophojemp.hecodigo " _
            & "WHERE cdfecha BETWEEN TO_DATE('" & Fecha & "', 'MM/DD/YYYY') " _
            & "AND TO_DATE('" & Fecha1 & "','MM/DD/YYYY' ) "

            
    If chkPlaca.Value = 1 Then
        QRY = QRY & "AND cdplaca = '" & txtPlaca.text & "' "
    End If
    
    If chkEmpresa.Value = 1 Then
        QRY = QRY & "AND cdcodemp = '" & txtCodEmp.text & "' "
    End If
    
    If cmbTerminal.ListIndex <> 0 Then
        QRY = QRY & "AND cdterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' "
    End If
    
    QRY = QRY & "ORDER BY cdnumero "
    
    ExportarExcel (QRY)
    
    Exit Sub
err_mensaje:
    MsgBox Err.Description, 16, "Error: " & Err.Number
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
    ApExcel.Cells(fila, 2).formula = "PLACA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "CODIGO EMPRESA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "NOMBRE EMPRESA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "FECHA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "HORA"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "TERMINAL"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "TASA USO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "COD. DESTINO"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "NOM. DESTINO"
    ApExcel.Cells(fila, 10).Font.Size = 10
    
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el ProgressBar
    If Not Tbl1.EOF Then
        Tbl1.MoveLast
        ProBar.Value = ProBar.Min
        ProBar.Max = Tbl1.RecordCount + 2
        Tbl1.MoveFirst
    End If
    
    If Not Tbl1.EOF Then
        
        While Not Tbl1.EOF
        
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!cdnumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!cdplaca & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!cdcodemp & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!cdfecha & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!cdhora & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!cdterminal & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!cdtasauso & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!cdcoddes & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!rudestino & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA}" & Month(Mskfecini.text) & Day(Mskfecini.text) & Year(Mskfecini.text)
        ApExcel.Visible = True
        ProBar.Value = ProBar.Max
        
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub

Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia: " & Err.Number
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
    chkEmpresa.Value = False
    chkPlaca.Value = False
    txtCodEmp.text = ""
    lblNomEmpresa.Caption = ""
    txtPlaca.text = ""
    lblPlaca.Caption = ""
    cmbTerminal.ListIndex = 0
    lblTotReg.Caption = ""
    DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Inicializamos la Grilla
    DIB_GRILLA

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
End Sub

Private Function DIB_GRILLA()
    grdConcond.rows = 1
    grdConcond.Cols = 10
    grdConcond.Clear
    
    s$ = "Consecutivo |Placa |Cod. Empresa |Nom. Empresa |Fecha |Hora |Terminal |Numero Tasa |Cod. Destino |Nom. Destino "
    grdConcond.FormatString = s$
    grdConcond.ColWidth(0) = 1100
    grdConcond.ColWidth(1) = 1600
    grdConcond.ColWidth(2) = 1200
    grdConcond.ColWidth(3) = 1600
    grdConcond.ColWidth(4) = 1800
    grdConcond.ColWidth(5) = 1400
    grdConcond.ColWidth(6) = 1400
    grdConcond.ColWidth(7) = 1600
    grdConcond.ColWidth(8) = 1200
    grdConcond.ColWidth(9) = 1800
    
    
    For i = 0 To grdConcond.Cols - 1
        grdConcond.Col = i
        grdConcond.CellFontBold = 1
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

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "73"
  lblNomEmpresa.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub TxtPlaca_DblClick()
    forma = "2"
    forma1 = "27"
    lblPlaca.Caption = ""
    buscar.Show 1
    TxtPlaca_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_LostFocus()
    
    If txtCodEmp.text <> "" Then
        
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
