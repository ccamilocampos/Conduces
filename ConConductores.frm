VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ConConductores 
   Caption         =   "Consulta de Conductores"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8880
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
      Picture         =   "ConConductores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Salir de la forma actual"
      Top             =   3720
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
      Picture         =   "ConConductores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Limpia la forma actual"
      Top             =   2880
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
      Picture         =   "ConConductores.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   1200
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
      Picture         =   "ConConductores.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   480
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
      Picture         =   "ConConductores.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   2040
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
      TabIndex        =   15
      Top             =   3720
      Width           =   7455
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      TabIndex        =   8
      Top             =   2280
      Width           =   7455
      Begin VB.TextBox txtCodEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   10
         Top             =   360
         Width           =   1335
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
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   4335
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
         Width           =   4335
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
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
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
         Width           =   1335
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
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1215
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
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.OptionButton OptEstado 
         Caption         =   "Activo"
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
      Begin VB.OptionButton OptEstado 
         Caption         =   "Inactivo"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
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
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConductores 
      Height          =   2055
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7800
      X2              =   7800
      Y1              =   0
      Y2              =   8520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   8880
      X2              =   0
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "ConConductores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnExportar_Click()
    On Error GoTo Errores
    
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
    
    If OptEstado(0).Value = True Then
        estado = "Activo"
    ElseIf OptEstado(1).Value = True Then
        estado = "Inactivo"
    Else
        estado = ""
    End If
    
    'Iniciamos la consulta
    QRY = "SELECT cocedula, CONCAT(conombres, CONCAT('', coapellidos)) AS conombres, henombre, coestado " _
            & "FROM coconductor c " _
            & "INNER JOIN ophojemp e " _
            & "ON c.cocodemp = e.hecodigo " _
            & "WHERE "
            
    If OptEstado(2).Value = True Then
        QRY = QRY & "coestado IN('Activo', 'Inactivo') "
    Else
        QRY = QRY & "coestado = '" & estado & "' "
    End If
            
    If txtCodEmp.text <> "" Then
        QRY = QRY & "AND cocodemp = '" & txtCodEmp.text & "' "
    End If
    
    If txtCedula.text <> "" Then
        QRY = QRY & "AND cocedula = '" & txtCedula.text & "' "
    End If
    
    QRY = QRY & "ORDER BY henombre ASC "
    
    ExportarExcel (QRY)
Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub ExportarExcel(QRY As String)
    On Error GoTo Errores
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "CEDULA"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "NOMBRES Y APELLIDOS"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "EMPRESA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "ESTADO"
    ApExcel.Cells(fila, 4).Font.Size = 10
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
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!cocedula & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!conombres & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!coestado & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA CONDUCTORES"
        ApExcel.Visible = True
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub

Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub BtnBuscar_Click()
    On Error GoTo Errores
    
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
    
    If OptEstado(0).Value = True Then
        estado = "Activo"
    ElseIf OptEstado(1).Value = True Then
        estado = "Inactivo"
    Else
        estado = ""
    End If
    
    'Iniciamos la consulta
    QRY = "SELECT cocedula, CONCAT(conombres, CONCAT('', coapellidos)) AS conombres, henombre, coestado " _
            & "FROM coconductor c " _
            & "INNER JOIN ophojemp e " _
            & "ON c.cocodemp = e.hecodigo " _
            & "WHERE "
            
    If OptEstado(2).Value = True Then
        QRY = QRY & "coestado IN('Activo', 'Inactivo') "
    Else
        QRY = QRY & "coestado = '" & estado & "' "
    End If
            
    If txtCodEmp.text <> "" Then
        QRY = QRY & "AND cocodemp = '" & txtCodEmp.text & "' "
    End If
    
    If txtCedula.text <> "" Then
        QRY = QRY & "AND cocedula = '" & txtCedula.text & "' "
    End If
    
    QRY = QRY & "ORDER BY henombre ASC "
            
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
            
            entry = CStr(Tbl1!cocedula) + Chr(9) + Tbl1!conombres + Chr(9) _
                        & Tbl1!henombre + Chr(9) + Tbl1!coestado
                        
            grdConductores.AddItem (entry)
            Tbl1.MoveNext
            
            TotalRegistros = TotalRegistros + 1
            ProBar.Value = ProBar.Value + 1
        
        Loop
        Tbl1.Close
        
        lblTotReg.Caption = TotalRegistros
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros", 48, "Advertencia"
        Exit Sub
    End If

Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error" & Err.Number
End Sub

Private Sub Btnlimpiar_Click()
    chkEmpresa.Value = False
    chkCedula.Value = False
    txtCodEmp.text = ""
    lblNomEmpresa.Caption = ""
    txtCedula.text = ""
    lblNomConductor.Caption = ""
    OptEstado(2).Value = True
    lblTotReg.Caption = ""
    DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    'Inicializamos las opciones
    OptEstado(2).Value = True
    
    
End Sub

Private Function DIB_GRILLA()

   grdConductores.rows = 1
   grdConductores.Cols = 4
   grdConductores.Clear
   
   s$ = "Cedula |Nombres y Apellidos |Empresa |Estado "
   grdConductores.FormatString = s$
   grdConductores.Row = 0
   grdConductores.Col = 0
   grdConductores.ColWidth(0) = 1400
   grdConductores.ColWidth(1) = 3000
   grdConductores.ColWidth(2) = 2000
   grdConductores.ColWidth(3) = 1000
   
   
   For i = 0 To grdConductores.Cols - 1
       grdConductores.Col = i
       grdConductores.CellFontBold = 1
   Next i
End Function

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

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "74"
  lblNomEmpresa.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub
Private Sub txtcedula_DblClick()
    forma = "20"
    forma1 = "10"
    lblNomConductor.Caption = ""
    buscar.Show 1
    txtcedula_LostFocus
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
        
        'Verificamos si ya digitaron la empresa
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
