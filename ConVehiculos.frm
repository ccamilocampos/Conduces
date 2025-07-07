VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ConVehiculos 
   Caption         =   "Consulta de Vehículos"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8775
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
      Left            =   7800
      Picture         =   "ConVehiculos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   4200
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
      Left            =   7800
      Picture         =   "ConVehiculos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Buscar información en la base de datos"
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
      Left            =   7800
      Picture         =   "ConVehiculos.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   3360
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
      Picture         =   "ConVehiculos.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Limpia la forma actual"
      Top             =   5040
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
      Picture         =   "ConVehiculos.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Salir de la forma actual"
      Top             =   5880
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
      TabIndex        =   14
      Top             =   5760
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   240
         Width           =   1455
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
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   7455
      Begin VB.TextBox txtClase 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   38
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Txtmodelo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtCombustible 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtServicio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtCodEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblClase 
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
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Clase:"
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
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblModelo 
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
         TabIndex        =   30
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Modelo:"
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
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblCombustible 
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
         TabIndex        =   27
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Combus.:"
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
         Width           =   1215
      End
      Begin VB.Label lblServicio 
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
         TabIndex        =   24
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio:"
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
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
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
         TabIndex        =   10
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
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
      Begin VB.CheckBox chkClase 
         Caption         =   "Clase"
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
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkModelo 
         Caption         =   "Modelo"
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
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkCombustible 
         Caption         =   "Combusti."
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
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkServicio 
         Caption         =   "Nivel Ser."
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
         TabIndex        =   19
         Top             =   720
         Width           =   1335
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
         TabIndex        =   6
         Top             =   360
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
         TabIndex        =   5
         Top             =   360
         Width           =   1335
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
      TabIndex        =   0
      Top             =   120
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
         TabIndex        =   3
         Top             =   720
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
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
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
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdVehiculos 
      Height          =   2775
      Left            =   3120
      TabIndex        =   18
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7680
      X2              =   7680
      Y1              =   -120
      Y2              =   8400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   8880
      X2              =   0
      Y1              =   6720
      Y2              =   6720
   End
End
Attribute VB_Name = "ConVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBuscar_Click()
    On Error GoTo Errores
    
    If chkEmpresa.Value = 1 Then
        If txtCodEmp.text = "" Then
            MsgBox "El codigo de la empresa no puede estar vacio", 48, "Advertencia"
            txtCodEmp.SetFocus
            Exit Sub
        End If
    End If
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    If chkServicio.Value = 1 Then
        If txtServicio.text = "" Then
            MsgBox "El campo de tipo de servicio no puede estar vacio", 48, "Advertencia"
            txtServicio.SetFocus
            Exit Sub
        End If
    End If
    If chkCombustible.Value = 1 Then
        If txtCombustible.text = "" Then
            MsgBox "El campo de tipo de combustible no puede estar vacio", 48, "Advertencia"
            txtCombustible.text = ""
            Exit Sub
        End If
    End If
    If chkModelo.Value = 1 Then
        If Txtmodelo.text = "" Then
            MsgBox "El campo de modelo no puede estar vacio", 48, "Advertencia"
            Txtmodelo.SetFocus
            Exit Sub
        End If
    End If
    If chkClase.Value = 1 Then
        If txtClase.text = "" Then
            MsgBox "El campo de clase no puede estar vacio", 48, "Advertencia"
            txtClase.SetFocus
            Exit Sub
        End If
    End If
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    If OptEstado(0).Value = True Then
        estado = "Activo"
    ElseIf OptEstado(1).Value = True Then
        estado = "Inactivo"
    Else
        estado = ""
    End If
    
    'Hacemos la consulta
    QRY = "SELECT hvplaca, henombre, hvnumpas, hvtipveh, hvnumint, hvmodelo, hvclase, hvcombus, hvtipveh, hvestado " _
                & "FROM ophojveh v " _
                & "INNER JOIN ophojemp e " _
                & "ON v.hvcodemp = e.hecodigo " _
                & "WHERE "
                
    If OptEstado(2).Value = True Then
        QRY = QRY & "hvestado IN ('Activo', 'Inactivo') "
    Else
        QRY = QRY & "hvestado = '" & estado & "' "
    End If
    
    If txtCodEmp.text <> "" Then
        QRY = QRY & "AND hvcodemp = '" & txtCodEmp.text & "' "
    End If
    
    If txtPlaca.text <> "" Then
        QRY = QRY & "AND hvplaca = '" & txtPlaca.text & "' "
    End If
    
    If txtServicio.text <> "" Then
        QRY = QRY & "AND hvtipveh = '" & txtServicio.text & "' "
    End If
    
    If txtCombustible.text <> "" Then
        QRY = QRY & "AND hvcombus = '" & txtCombustible.text & "' "
    End If
    
    If Txtmodelo.text <> "" Then
        QRY = QRY & "AND hvmodelo = " & Txtmodelo.text & " "
    End If
    
    If txtClase.text <> "" Then
        QRY = QRY & "AND hvclase = '" & txtClase.text & "' "
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
        
            'Traemos la clase
            If Not IsNull(Tbl1!hvclase) Then
                clase = Consultar_Catalogo(Tbl1!hvclase, "TIPOS DE CLASES")
            Else
                clase = "N/A"
            End If
            'Traemos el tipo de combustible
            If Not IsNull(Tbl1!hvcombus) Then
                combustible = Consultar_Catalogo(Tbl1!hvcombus, "TIPO COMBUSTIBLE")
            Else
                combustible = "N/A"
            End If
            'Traemos el tipo de servicio
            If Not IsNull(Tbl1!hvtipveh) Then
                servicio = Consultar_Catalogo(Tbl1!hvtipveh, "TIPOS DE VEHICULO")
            Else
                servicio = "N/A"
            End If
            
            'Verificamos si el módelo es nulo
            If Not IsNull(Tbl1!hvmodelo) Then
                modelo = CStr(Tbl1!hvmodelo)
            Else
                modelo = "N/A"
            End If
            
            entry = Tbl1!hvplaca + Chr(9) + CStr(Tbl1!hvnumint) + Chr(9) + Tbl1!henombre + Chr(9) + clase + Chr(9) _
                        & CStr(Tbl1!hvnumpas) + Chr(9) + combustible + Chr(9) + servicio + Chr(9) _
                        & modelo + Chr(9) + Tbl1!hvestado
                        
            grdVehiculos.AddItem (entry)
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

Private Sub BtnExportar_Click()
    On Error GoTo Errores
    
    If chkEmpresa.Value = 1 Then
        If txtCodEmp.text = "" Then
            MsgBox "El codigo de la empresa no puede estar vacio", 48, "Advertencia"
            txtCodEmp.SetFocus
            Exit Sub
        End If
    End If
    If chkPlaca.Value = 1 Then
        If txtPlaca.text = "" Then
            MsgBox "El campo de placa no puede estar vacio", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
        End If
    End If
    If chkServicio.Value = 1 Then
        If txtServicio.text = "" Then
            MsgBox "El campo de tipo de servicio no puede estar vacio", 48, "Advertencia"
            txtServicio.SetFocus
            Exit Sub
        End If
    End If
    If chkCombustible.Value = 1 Then
        If txtCombustible.text = "" Then
            MsgBox "El campo de tipo de combustible no puede estar vacio", 48, "Advertencia"
            txtCombustible.text = ""
            Exit Sub
        End If
    End If
    If chkModelo.Value = 1 Then
        If Txtmodelo.text = "" Then
            MsgBox "El campo de modelo no puede estar vacio", 48, "Advertencia"
            Txtmodelo.SetFocus
            Exit Sub
        End If
    End If
    If chkClase.Value = 1 Then
        If txtClase.text = "" Then
            MsgBox "El campo de clase no puede estar vacio", 48, "Advertencia"
            txtClase.SetFocus
            Exit Sub
        End If
    End If
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    If OptEstado(0).Value = True Then
        estado = "Activo"
    ElseIf OptEstado(1).Value = True Then
        estado = "Inactivo"
    Else
        estado = ""
    End If
    
    'Hacemos la consulta
    QRY = "SELECT hvplaca, henombre, hvnumpas, hvtipveh, hvnumint, hvmodelo, hvclase, hvcombus, hvtipveh, hvestado " _
                & "FROM ophojveh v " _
                & "INNER JOIN ophojemp e " _
                & "ON v.hvcodemp = e.hecodigo " _
                & "WHERE "
                
    If OptEstado(2).Value = True Then
        QRY = QRY & "hvestado IN ('Activo', 'Inactivo') "
    Else
        QRY = QRY & "hvestado = '" & estado & "' "
    End If
    
    If txtCodEmp.text <> "" Then
        QRY = QRY & "AND hvcodemp = '" & txtCodEmp.text & "' "
    End If
    
    If txtPlaca.text <> "" Then
        QRY = QRY & "AND hvplaca = '" & txtPlaca.text & "' "
    End If
    
    If txtServicio.text <> "" Then
        QRY = QRY & "AND hvtipveh = '" & txtServicio.text & "' "
    End If
    
    If txtCombustible.text <> "" Then
        QRY = QRY & "AND hvcombus = '" & txtCombustible.text & "' "
    End If
    
    If Txtmodelo.text <> "" Then
        QRY = QRY & "AND hvmodelo = " & Txtmodelo.text & " "
    End If
    
    If txtClase.text <> "" Then
        QRY = QRY & "AND hvclase = '" & txtClase.text & "' "
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
    ApExcel.Cells(fila, 1).formula = "PLACA"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "NUM. INTERNO"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "EMPRESA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "CLASE"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "CANT. SILLAS"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "TIP. COMBUSTIBLE"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "NIVEL SERVICIO"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "MODELO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "ESTADO"
    ApExcel.Cells(fila, 9).Font.Size = 10
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
            
            'Traemos la clase
            If Not IsNull(Tbl1!hvclase) Then
                clase = Consultar_Catalogo(Tbl1!hvclase, "TIPOS DE CLASES")
            Else
                clase = "N/A"
            End If
            'Traemos el tipo de combustible
            If Not IsNull(Tbl1!hvcombus) Then
                combustible = Consultar_Catalogo(Tbl1!hvcombus, "TIPO COMBUSTIBLE")
            Else
                combustible = "N/A"
            End If
            'Traemos el tipo de servicio
            If Not IsNull(Tbl1!hvtipveh) Then
                servicio = Consultar_Catalogo(Tbl1!hvtipveh, "TIPOS DE VEHICULO")
            Else
                servicio = "N/A"
            End If
            
            'Verificamos si el módelo es nulo
            If Not IsNull(Tbl1!hvmodelo) Then
                modelo = CStr(Tbl1!hvmodelo)
            Else
                modelo = "N/A"
            End If
            
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!hvplaca & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!hvnumint & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & clase & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!hvnumpas & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & combustible & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & servicio & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & modelo & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!hvestado & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA VEHICULOS"
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

Private Sub Btnlimpiar_Click()
    chkEmpresa.Value = False
    chkPlaca.Value = False
    chkServicio.Value = False
    chkCombustible.Value = False
    chkModelo.Value = False
    chkClase.Value = False
    txtCodEmp.text = ""
    lblNomEmpresa.Caption = ""
    txtPlaca.text = ""
    lblPlaca.Caption = ""
    txtServicio.text = ""
    lblServicio.Caption = ""
    txtCombustible.text = ""
    lblCombustible.Caption = ""
    Txtmodelo.text = ""
    lblModelo.Caption = ""
    txtClase.text = ""
    lblClase.Caption = ""
    OptEstado(2).Value = True
    lblTotReg.Caption = ""
    DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub chkClase_Click()
    If chkClase.Value = 1 Then
        txtClase.Enabled = True
    Else
        txtClase.text = ""
        lblClase.Caption = ""
        txtClase.Enabled = False
    End If
End Sub

Private Sub chkCombustible_Click()
    If chkCombustible.Value = 1 Then
        txtCombustible.Enabled = True
    Else
        txtCombustible.text = ""
        lblCombustible.Caption = ""
        txtCombustible.Enabled = False
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

Private Sub chkModelo_Click()
    If chkModelo.Value = 1 Then
        Txtmodelo.Enabled = True
    Else
        Txtmodelo.text = ""
        lblModelo.Caption = ""
        Txtmodelo.Enabled = False
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

Private Sub chkServicio_Click()
    If chkServicio.Value = 1 Then
        txtServicio.Enabled = True
    Else
        txtServicio.text = ""
        lblServicio.Caption = ""
        txtServicio.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    
    'Inicializamos la grilla
    DIB_GRILLA
    
    'Inicializamos las opciones
    OptEstado(2).Value = True
    
End Sub

Private Function DIB_GRILLA()

   grdVehiculos.rows = 1
   grdVehiculos.Cols = 9
   grdVehiculos.Clear
   
   s$ = "Placa |Num. Interno |Empresa |Clase |Cant. Sillas |Tip. Combustible |Nivel Servicio |Modelo |Estado "
   grdVehiculos.FormatString = s$
   grdVehiculos.Row = 0
   grdVehiculos.Col = 0
   grdVehiculos.ColWidth(0) = 1500
   grdVehiculos.ColWidth(1) = 1200
   grdVehiculos.ColWidth(2) = 3000
   grdVehiculos.ColWidth(3) = 1000
   grdVehiculos.ColWidth(4) = 1100
   grdVehiculos.ColWidth(5) = 1300
   grdVehiculos.ColWidth(6) = 1300
   grdVehiculos.ColWidth(7) = 1200
   grdVehiculos.ColWidth(8) = 1300
   
   
   For i = 0 To grdVehiculos.Cols - 1
       grdVehiculos.Col = i
       grdVehiculos.CellFontBold = 1
   Next i
End Function

Private Sub TxtClase_DblClick()
    forma = "15"
    forma1 = "65"
    Tabla_Catalogo = "TIPOS DE CLASES"
    lblClase.Caption = ""
    buscar.Show 1
    TxtClase_LostFocus
End Sub

Private Sub TxtClase_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub TxtClase_LostFocus()
    If txtClase.text <> "" Then
        QRY = "SELECT dsdes " _
                & "FROM gesuptip " _
                & "INNER JOIN gedetsuptip " _
                & "ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                & "WHERE stdes = 'TIPOS DE CLASES' " _
                & "AND dscoddet = '" & txtClase.text & "' "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            lblClase.Caption = Tbl1!dsdes
        Else
            txtClase.text = ""
            lblClase.Caption = ""
            MsgBox "El tipo de clase no existe o se encuentra inactivo", 48, "Advertencia"
        End If
    End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "75"
  lblNomEmpresa.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
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

Private Sub txtCombustible_DblClick()
    forma = "15"
    forma1 = "64"
    Tabla_Catalogo = "TIPO COMBUSTIBLE"
    lblCombustible.Caption = ""
    buscar.Show 1
    txtCombustible_LostFocus
End Sub

Private Sub txtCombustible_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub txtCombustible_LostFocus()
    If txtCombustible.text <> "" Then
        QRY = "SELECT dsdes " _
                & "FROM gesuptip " _
                & "INNER JOIN gedetsuptip " _
                & "ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                & "WHERE stdes = 'TIPO COMBUSTIBLE' " _
                & "AND dscoddet = '" & txtCombustible.text & "' "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            lblCombustible.Caption = Tbl1!dsdes
        Else
            txtCombustible.text = ""
            lblCombustible.Caption = ""
            MsgBox "El tipo de combustible no existe o se encuentra inactivo", 48, "Advertencia"
        End If
    End If
End Sub

Private Sub txtmodelo_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub Txtmodelo_LostFocus()
    If Txtmodelo.text <> "" Then
        lblModelo.Caption = Txtmodelo.text
    End If
End Sub

Private Sub TxtPlaca_DblClick()
    forma = "2"
    forma1 = "28"
    lblPlaca.Caption = ""
    buscar.Show 1
    TxtPlaca_LostFocus
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

Private Sub txtServicio_DblClick()
    forma = "15"
    forma1 = "63"
    Tabla_Catalogo = "TIPOS DE VEHICULO"
    lblServicio.Caption = ""
    buscar.Show 1
    txtServicio_LostFocus
End Sub

Private Sub txtServicio_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub txtServicio_LostFocus()
    If txtServicio.text <> "" Then
        QRY = "SELECT dsdes " _
                & "FROM gesuptip " _
                & "INNER JOIN gedetsuptip " _
                & "ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                & "WHERE stdes = 'TIPOS DE VEHICULO' " _
                & "AND dscoddet = '" & txtServicio.text & "' "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            lblServicio.Caption = Tbl1!dsdes
        Else
            txtServicio.text = ""
            lblServicio.Caption = ""
            MsgBox "El tipo de servicio no existe o se encuentra inactivo", 48, "Advertencia"
        End If
    End If
End Sub

Private Function Consultar_Catalogo(fDato As String, fCatalogo As String) As String
    QRY = "SELECT dsdes " _
                & "FROM gesuptip " _
                & "INNER JOIN gedetsuptip " _
                & "ON gesuptip.stcodtip = gedetsuptip.dscodtip " _
                & "WHERE stdes = '" & fCatalogo & "' " _
                & "AND dscoddet = '" & fDato & "' "
                
    Set Tbl2 = Dbs.Execute(QRY)
    
    If Not Tbl2.EOF Then
        dato = Tbl2!dsdes
    Else
        dato = ""
    End If
    
    Consultar_Catalogo = dato
End Function
