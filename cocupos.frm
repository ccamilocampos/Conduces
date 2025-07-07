VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form cocupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cupos"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Información General"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   5895
      Begin VB.ComboBox CmbTerminal 
         Height          =   315
         ItemData        =   "cocupos.frx":0000
         Left            =   1680
         List            =   "cocupos.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Tipos de terminales"
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox Cmbestado 
         Height          =   315
         ItemData        =   "cocupos.frx":0004
         Left            =   3960
         List            =   "cocupos.frx":000E
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Estado del cupo"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txtobservaciones 
         Height          =   1335
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Observaciones del cupo"
         Top             =   1680
         Width           =   5175
      End
      Begin MSMask.MaskEdBox Mskminimo 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Valor mínimo del cupo"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskmaximo 
         Height          =   285
         Left            =   3960
         TabIndex        =   8
         ToolTipText     =   "Valor máximo del cupo"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Fecha del cupo"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         Caption         =   "Terminal:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Mínimo:"
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
         Index           =   16
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Máximo:"
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
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
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
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   3120
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Observaciones:"
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
         Top             =   1440
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Empresa"
      TabPicture(0)   =   "cocupos.frx":0024
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vehículo"
      TabPicture(1)   =   "cocupos.frx":0040
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Cupo"
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
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   5415
         Begin VB.CommandButton btnañadir 
            Caption         =   "A&ñadir"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            OLEDropMode     =   1  'Manual
            Picture         =   "cocupos.frx":005C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Añadir el registro"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox TxtCodEmpVeh 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   4
            ToolTipText     =   "Código de la Empresa de Transporte, Presione DblClick para Ayuda"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TxtPlaca 
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
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   5
            ToolTipText     =   "Placa del vehículo"
            Top             =   960
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid GrdCupo 
            Height          =   1695
            Left            =   2640
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   2990
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            AllowUserResizing=   1
         End
         Begin VB.Label Label1 
            Caption         =   "Código Empresa:"
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
            Index           =   6
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label LblEmpVehiculo 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label1 
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
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Cupo"
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
         Left            =   -74760
         TabIndex        =   21
         Top             =   600
         Width           =   5415
         Begin VB.TextBox Txtcodemp 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   3
            ToolTipText     =   "Código de la Empresa de Transporte, Presione DblClick para Ayuda"
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Lblnomemp 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   18
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Código Empresa:"
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
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox Txtnumcup 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Número consecutivo de cupos"
      Top             =   240
      Width           =   1215
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
      Left            =   6240
      Picture         =   "cocupos.frx":049E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salir de la forma actual"
      Top             =   5760
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
      Left            =   6240
      Picture         =   "cocupos.frx":08E0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Limpia la forma actual"
      Top             =   4920
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
      Left            =   6240
      Picture         =   "cocupos.frx":0D22
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   4080
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
      Left            =   6240
      Picture         =   "cocupos.frx":1164
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Btngrabar 
      Caption         =   "&Grabar"
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
      Left            =   6240
      Picture         =   "cocupos.frx":15A6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ingresa el registro en la base de datos"
      Top             =   2400
      Width           =   855
   End
   Begin MSMask.MaskEdBox Msksaldo 
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Saldo del cupo"
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0;(#,##0)"
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   6600
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo:"
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
      Index           =   11
      Left            =   3240
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "No. Cupo:"
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
      TabIndex        =   19
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "cocupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sBanderaEmpresa As Boolean


Private Function Valida_Cupo_Empresa(ByVal fCodigoEmpresa As String) As Boolean
   QRY = "SELECT * " _
         & "FROM cocupos " _
        & "WHERE cpempresa = '" & fCodigoEmpresa & "' " _
          & "AND cpestado = 'A'"
   
   Set Ftb04 = Dbs.Execute(QRY)
   Valida_Cupo_Empresa = False
   If Not Ftb04.EOF Then
      Valida_Cupo_Empresa = True
   End If
   Ftb04.Close
   

End Function

Private Function DIB_GRILLA()
   
   GrdCupo.rows = 1
   GrdCupo.Cols = 2
   GrdCupo.Clear
   
   s$ = "Placa  |Estado "
   GrdCupo.FormatString = s$
   GrdCupo.Row = 0
   GrdCupo.Col = 0
   GrdCupo.ColWidth(0) = 1000
   GrdCupo.ColWidth(1) = 1000
   
   For i = 0 To GrdCupo.Cols - 1
       GrdCupo.Col = i
       GrdCupo.CellFontBold = 1
   Next i
End Function

Private Function Cargar_Vehiculos(ByVal sNumero As Double)
   QRY = "SELECT cdplaca, " _
              & "CASE cdestado " _
              & "WHEN 'A' THEN 'Activo' " _
              & "WHEN 'I' THEN 'Inactivo' " _
              & "END AS cdestado " _
         & "FROM cocuposdet " _
        & "WHERE cdnumcupo = " & sNumero & " " _
     & "ORDER BY cdplaca"
   
   Set Tbl10 = Dbs.Execute(QRY)
   
   Do While Not Tbl10.EOF
      entry = Tbl10!cdplaca + Chr(9) + Tbl10!cdestado
      GrdCupo.AddItem entry
      Tbl10.MoveNext
   Loop
   Tbl10.Close
End Function

Private Sub btnañadir_Click()
   On Error GoTo Errores
   
   If txtPlaca.text = "" Then
      MsgBox "Falta ingresar el número de la placa", 48, "Advertencia"
      txtPlaca.SetFocus
      Exit Sub
   End If
   
   
   For i = 1 To GrdCupo.rows - 1
      sPlaca = GrdCupo.TextMatrix(i, 0)
      
      If sPlaca = txtPlaca.text Then
         MsgBox "La Placa ya esta asociada en el cupo", 48, "Advertencia"
         txtPlaca.text = ""
         txtPlaca.SetFocus
         Exit Sub
      End If
   Next i
   
   sEstado = "Activo"
   entry = txtPlaca.text + Chr(9) + sEstado

   GrdCupo.AddItem entry
   txtPlaca.text = ""
   txtPlaca.SetFocus
   Exit Sub
Errores:
   
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnbuscar_Click()
   On Error Resume Next
   forma = "22"
   forma1 = "1"
   Txtnumcup.text = ""
   buscar.Show 1
   Txtnumcup_LostFocus

End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
      
   If Trim(Txtnumcup.text) = "" Then
      MsgBox "Falta ingresar el número de cupo", 48, "Advertencia"
      Exit Sub
   End If
   
   
   If SSTab1.Tab = 0 Then
      If Trim(txtcodEmp.text) = "" Then
         MsgBox "Falta ingresar el código de empresa", 48, "Advertencia"
         txtcodEmp.SetFocus
         Exit Sub
      End If
      sTipo = "E"
   ElseIf SSTab1.Tab = 1 Then
      
'      If Trim(TxtPlaca.Text) = "" Then
'         MsgBox "Falta ingresar el número de la placa", 48, "Advertencia"
'         TxtPlaca.SetFocus
'         Exit Sub
'      End If
      If Trim(TxtCodEmpVeh.text) = "" Then
         MsgBox "Falta ingresar el código de empresa", 48, "Advertencia"
         TxtCodEmpVeh.SetFocus
         Exit Sub
      End If
      
      If GrdCupo.rows = 1 Then
         MsgBox "No se han relacionado las placas del cupo", 48, "Advertencia"
         txtPlaca.SetFocus
         Exit Sub
      End If
      sTipo = "P"
   End If
   If Trim(Mskminimo.ClipText) = "" Then
      MsgBox "Falta ingresar el valor mínimo", 48, "Advertencia"
      Mskminimo.SetFocus
      Exit Sub
   End If
   
   If Trim(Mskmaximo.ClipText) = "" Then
      MsgBox "Falta ingresar el valor máximo", 48, "Advertencia"
      Mskmaximo.SetFocus
      Exit Sub
   End If
   
   If Cmbestado.ListIndex = -1 Then
      MsgBox "Falta seleccionar el estado del cupo", 48, "Advertencia"
      Cmbestado.SetFocus
      Exit Sub
   End If
   
   If Trim(TxtObservaciones.text) = "" Then
      MsgBox "Falta ingresar las observaciones", 48, "Advertencia"
      TxtObservaciones.SetFocus
      Exit Sub
   End If
   
   respuesta = MsgBox("Desea ingresar el registro", 32 + vbYesNo, "Grabar")
  
   If respuesta = vbNo Then
      Exit Sub
   End If
   
   If Btngrabar.Caption = "&Actualizar" Then
      sFechaSistema = Mid(Fn_Parametros("HORASIS", 1), 1, 10)
   End If
   
   Dbs.BeginTrans
   On Error GoTo Err_Transac
   
   If SSTab1.Tab = 0 Then
      sCodigoEmpresa = Trim(txtcodEmp.text)
   Else
      sCodigoEmpresa = Trim(TxtCodEmpVeh.text)
   End If
   If Btngrabar.Caption = "&Grabar" Then
      QRY = "INSERT INTO cocupos (" _
            & "cpnumcup, cpfecha, cpminimo, " _
            & "cpmaximo, cpfuncio, cpestado, cpobserv, cpempresa, cptipo, cpterminal) VALUES (" _
            & Trim(Txtnumcup.text) & ",TO_DATE('" & Format(Mskfecha.text, "MM/dd/yyyy") & "','MM/DD/YYYY') , " _
            & Mskminimo.text & ", " & Mskmaximo.text & ",'" & Login & "','" & Mid(Cmbestado.text, 1, 1) & "','" _
            & Trim(TxtObservaciones.text) & "','" & sCodigoEmpresa & "','" & sTipo & "', '" & Mid(cmbTerminal.text, 1, 1) & "')"
         
      Dbs.Execute QRY
      
      Dbs.Execute "UPDATE geparsis " _
               & "SET psval = psval + 1 " _
             & "WHERE pscod = 'NUMCUP' "
             
      If SSTab1.Tab = 1 Then
         For i = 1 To GrdCupo.rows - 1
            sPlaca = GrdCupo.TextMatrix(i, 0)
            If GrdCupo.TextMatrix(i, 1) = "Activo" Then
               sEstado = "A"
            Else
               sEstado = "I"
            End If
            
            QRY = "INSERT INTO cocuposdet(cdnumcupo, cdplaca,cdestado, cdfecha, cdfuncionario) " _
            & "VALUES(" _
            & Txtnumcup.text & ",'" & sPlaca & "','" & sEstado & "',TO_DATE('" & Format(Mskfecha.text, "MM/dd/yyyy") & "','MM/DD/YYYY'),'" & Login & "')"
            
            Dbs.Execute QRY
            
            
         Next i
      End If
      sMensaje = "El registro se grabó exitosamente"
   ElseIf Btngrabar.Caption = "&Actualizar" Then
   
      QRY = "UPDATE cocupos " _
             & "SET cpminimo = " & Mskminimo.text & ", " _
                 & "cpmaximo = " & Mskmaximo.text & ", " _
                 & "cpfuncio = '" & Login & "' , " _
                 & "cpestado = '" & Mid(Cmbestado.text, 1, 1) & "' , " _
                 & "cpobserv = '" & Trim(TxtObservaciones.text) & "', " _
                 & "cpterminal = '" & Mid(cmbTerminal.text, 1, 1) & "' " _
            & "WHERE cpnumcup = " & Trim(Txtnumcup.text) & ""
      
      Dbs.Execute QRY
      
      If SSTab1.Tab = 1 Then
         For i = 1 To GrdCupo.rows - 1
            sPlaca = GrdCupo.TextMatrix(i, 0)
            
            If GrdCupo.TextMatrix(i, 1) = "Activo" Then
               sEstado = "A"
            Else
               sEstado = "I"
            End If
            
            'Se busca en la tabla para hacer cambio de estado del vehiculo
            QRY = "UPDATE cocuposdet " _
                   & "SET cdestado = '" & sEstado & "' " _
                 & "WHERE cdnumcupo = " & Txtnumcup.text & " " _
                   & "AND cdplaca ='" & sPlaca & "'"
                   
            Dbs.Execute QRY, sRecordAffected
            
            If sRecordAffected = 0 Then
               'si no se afecta se graba la placa con la fecha y login para auditoria
               QRY = "INSERT INTO cocuposdet(cdnumcupo, cdplaca,cdestado, cdfecha, cdfuncionario) " _
               & "VALUES(" _
               & Txtnumcup.text & ",'" & sPlaca & "','" & sEstado & "',TO_DATE('" & sFechaSistema & "','MM/DD/YYYY'),'" & Login & "')"
               
               Dbs.Execute QRY
            End If
         Next i
      End If
      
      sMensaje = "El registro se actualizó exitosamente"
   End If
   
   
   Dbs.CommitTrans
   On Error GoTo Errores
   
   MsgBox sMensaje, 64, "O.K."
   Btnlimpiar_Click
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
Err_Transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnlimpiar_Click()
   DIB_GRILLA
   sBanderaEmpresa = True
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMCUP' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
  
   If Not Ftb01.EOF Then
      Txtnumcup.text = Ftb01!psval + 1
   End If
   
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecha.text = Fecha
   End If
   Ftb01.Close
   SSTab1.Enabled = True
   SSTab1.Tab = 0
   cmbTerminal.ListIndex = 0
   txtcodEmp.Enabled = True
   txtPlaca.Enabled = True
   MskSaldo.text = ""
   txtcodEmp.text = ""
   Lblnomemp.Caption = ""
   Mskminimo.text = ""
   Mskmaximo.text = ""
   Cmbestado.ListIndex = 0
   TxtObservaciones.text = ""
   txtPlaca.text = ""
   TxtCodEmpVeh.text = ""
   LblEmpVehiculo.Caption = ""
   
   txtcodEmp.SetFocus
   Btngrabar.Caption = "&Grabar"
   Btngrabar.Enabled = True
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   sBanderaEmpresa = True
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMCUP' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
  
   If Not Ftb01.EOF Then
      Txtnumcup.text = Ftb01!psval + 1
   End If
  
   'hora actual del sistema
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      Mskfecha.text = Fecha
   End If
   Ftb01.Close
   Cmbestado.ListIndex = 0
   
   QRY = "SELECT dsdes " _
         & "FROM gesuptip, gedetsuptip " _
        & "WHERE dscodtip = stcodtip " _
          & "AND stdes = 'TIPOS DE TERMINAL' " _
          & "and dsest = 'A' " _
          & "ORDER BY dsdes "
    
   Set Tbl1 = Dbs.Execute(QRY)
   
   cmbTerminal.Clear
   cmbTerminal.AddItem "GENERAL"
   Do While Not Tbl1.EOF
      cmbTerminal.AddItem Tbl1(0)
      
      Tbl1.MoveNext
   Loop
   Tbl1.Close
   cmbTerminal.ListIndex = 0
   
   DIB_GRILLA
   SSTab1.Tab = 0
   
End Sub

Private Sub GrdCupo_DblClick()
   If GrdCupo.rows > 1 Then
      Y = GrdCupo.Row
      
      If GrdCupo.TextMatrix(Y, 1) = "Activo" Then
         GrdCupo.TextMatrix(Y, 1) = "Inactivo"
      Else
         GrdCupo.TextMatrix(Y, 1) = "Activo"
      End If
   End If

End Sub

Private Sub Mskmaximo_GotFocus()
   Mskmaximo.SelStart = 0
   Mskmaximo.SelLength = Len(Mskmaximo.text)
End Sub

Private Sub Mskmaximo_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub Mskmaximo_LostFocus()
   On Error GoTo Errores
   If Mskmaximo.ClipText <> "" Then
      If Mskminimo.ClipText <> "" Then
         If CDbl(Mskminimo.text) > CDbl(Mskmaximo.text) Then
            MsgBox "El Valor máximo no puede ser menor al valor mínimo", 48, "Advertencia"
            Mskmaximo.SetFocus
            Exit Sub
         End If
      End If
   
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Mskminimo_GotFocus()
   Mskminimo.SelStart = 0
   Mskminimo.SelLength = Len(Mskminimo.text)
End Sub

Private Sub Mskminimo_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   On Error Resume Next
   If PreviousTab = 1 Then
      If sBanderaEmpresa = True Then
         txtcodEmp.Enabled = True
         txtcodEmp.SetFocus
      End If
      
      TxtCodEmpVeh.Enabled = False
      txtPlaca.Enabled = False
      Btnañadir.Enabled = False
   Else
      txtPlaca.Enabled = True
      Btnañadir.Enabled = True
      txtcodEmp.Enabled = False
      If sBanderaEmpresa = True Then
         TxtCodEmpVeh.Enabled = True
         TxtCodEmpVeh.SetFocus
      Else
         txtPlaca.SetFocus
      End If
      
      
   End If
End Sub

Private Sub txtcodemp_DblClick()
  forma = "16"
  forma1 = "54"
  txtcodEmp.text = ""
  buscar.Show 1
  txtcodEmp.SetFocus
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   txtcodEmp.SelStart = 0
   txtcodEmp.SelLength = Len(txtcodEmp.text)
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
   If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub txtcodemp_LostFocus()
   On Error GoTo Errores
   If txtcodEmp.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & txtcodEmp.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         If sBanderaEmpresa = True Then
            If Valida_Cupo_Empresa(txtcodEmp.text) = True Then
               Ftb01.Close
               MsgBox "La empresa tiene un cupo asignado", 48, "Advertencia"
               Lblnomemp.Caption = ""
               txtcodEmp.SetFocus
               Exit Sub
            End If
         End If
         Lblnomemp.Caption = Ftb01!henombre
      Else
         MsgBox "Código de empresa no existe", 48, "Advertencia"
         txtcodEmp.text = ""
         txtcodEmp.SetFocus
      End If
      Ftb01.Close
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub



Private Sub TxtCodEmpVeh_DblClick()
   forma = "16"
   forma1 = "66"
   TxtCodEmpVeh.text = ""
   LblNomEmpVeh.Caption = ""
   buscar.Show 1
   TxtCodEmpVeh_LostFocus
End Sub

Private Sub TxtCodEmpVeh_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtCodEmpVeh_KeyPress(KeyAscii As Integer)
If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub TxtCodEmpVeh_LostFocus()
   On Error GoTo Errores
   If TxtCodEmpVeh.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & TxtCodEmpVeh.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         If sBanderaEmpresa = True Then
            If Valida_Cupo_Empresa(TxtCodEmpVeh.text) = True Then
               Ftb01.Close
               MsgBox "La empresa que pertenece el vehículo tiene un cupo asignado", 48, "Advertencia"
               LblEmpVehiculo.Caption = ""
               TxtCodEmpVeh.SetFocus
               Exit Sub
            End If
         End If
         LblEmpVehiculo.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
         MsgBox "Código de empresa no existe", 48, "Advertencia"
         TxtCodEmpVeh.text = ""
         TxtCodEmpVeh.SetFocus
         Exit Sub
      End If
      
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Txtnumcup_DblClick()
   Btnbuscar_Click
End Sub


Private Sub Txtnumcup_LostFocus()
   On Error GoTo Errores
   
   If Trim(Txtnumcup.text) <> "" Then
      QRY = "SELECT * " _
            & "FROM cocupos " _
           & "WHERE cpnumcup = " & Trim(Txtnumcup.text) & " "
           
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         sBanderaEmpresa = False
         If Tbl1!cpestado = "I" Then
            Btngrabar.Enabled = False
            Cmbestado.ListIndex = 1
         Else
            Cmbestado.ListIndex = 0
            
         End If
         
         MskSaldo.text = Tbl1!cpsaldo
         Mskminimo.text = Tbl1!cpminimo
         Mskmaximo.text = Tbl1!cpmaximo
         TxtObservaciones.text = Tbl1!cpobserv
         Mskfecha.text = Format(Tbl1!cpfecha, "MM/dd/yyyy")
         
         If Not IsNull(Tbl1!cpterminal) Then
            If Tbl1!cpterminal = "G" Then
               cmbTerminal.ListIndex = 0
            ElseIf Tbl1!cpterminal = "C" Then
               cmbTerminal.ListIndex = 1
            ElseIf Tbl1!cpterminal = "N" Then
               cmbTerminal.ListIndex = 2
            ElseIf Tbl1!cpterminal = "S" Then
               cmbTerminal.ListIndex = 3
            End If
         Else
            cmbTerminal.ListIndex = 0
         End If
         
         If Tbl1!cptipo = "E" Then
            SSTab1.Tab = 0
            txtcodEmp.text = Tbl1!cpempresa
            txtcodEmp.Enabled = False
            Tbl1.Close
            txtcodemp_LostFocus
         Else
            SSTab1.Tab = 1
            TxtCodEmpVeh.text = Tbl1!cpempresa
            TxtCodEmpVeh.Enabled = False
            Tbl1.Close
            TxtCodEmpVeh_LostFocus
            Cargar_Vehiculos (Txtnumcup.text)
         End If
         
         
         Btngrabar.Caption = "&Actualizar"
      Else
         Tbl1.Close
      End If
      
   End If
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Txtobservaciones_GotFocus()
   TxtObservaciones.SelStart = 0
   TxtObservaciones.SelLength = Len(TxtObservaciones.text)
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
  End If

End Sub

Private Sub TxtPlaca_GotFocus()
   txtPlaca.SelStart = 0
   txtPlaca.SelLength = Len(txtPlaca.text)
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
   If txtPlaca.text <> "" Then
          
      longitud = Len(txtPlaca.text)
  
      If longitud < 6 Then
         MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
         txtPlaca.SetFocus
         Exit Sub
      End If
     
      If Trim(TxtCodEmpVeh.text) = "" Then
         MsgBox "Falta ingresar el código de empresa", 48, "Advertencia"
         TxtCodEmpVeh.SetFocus
         Exit Sub
      End If
      
      fqry01 = "SELECT * " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & txtPlaca.text & "' "
              
      Set Ftb02 = Dbs.Execute(fqry01)
     
      If Not Ftb02.EOF Then
         
         If Trim(TxtCodEmpVeh.text) <> Ftb02!hvcodemp Then
            MsgBox "El vehículo pertenece a otra empresa", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
         End If
         
         If Ftb02!hvestado = "Inactivo" Then
            MsgBox "El vehículo esta inactivo", 48, "Advertencia"
            txtPlaca.SetFocus
            Exit Sub
         End If
         
         Ftb02.Close
         
      Else
         Ftb02.Close
         MsgBox "La placa del vehículo no existe", 48, "Advertencia"
         txtPlaca.SetFocus
         Exit Sub
      End If
   End If
End Sub
