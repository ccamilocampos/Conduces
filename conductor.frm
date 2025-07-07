VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C0A57E85-C72B-44C4-A7EF-EBE05047E8C6}#1.1#0"; "CheckPeople3.ocx"
Begin VB.Form conductor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conductores"
   ClientHeight    =   6090
   ClientLeft      =   1485
   ClientTop       =   1635
   ClientWidth     =   7890
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnnovedades 
      Caption         =   "&Novedades"
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
      Left            =   6960
      Picture         =   "conductor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Novedades del conductor"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton btnhuella 
      Caption         =   "&Huella"
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
      Left            =   6960
      Picture         =   "conductor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ingresar huella al conductor"
      Top             =   120
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
      Left            =   6960
      Picture         =   "conductor.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   3600
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
      Left            =   6960
      Picture         =   "conductor.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir de la Forma Actual"
      Top             =   5280
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
      Left            =   6960
      Picture         =   "conductor.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Limpia la forma actual"
      Top             =   4440
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
      Left            =   6960
      Picture         =   "conductor.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar Informacion en la Base de Datos"
      Top             =   2760
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
      Left            =   6960
      Picture         =   "conductor.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ingresa El Conductor  a la Base de Datos"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
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
      Height          =   5895
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox CmbTipoConductor 
         Height          =   315
         ItemData        =   "conductor.frx":1DCE
         Left            =   1680
         List            =   "conductor.frx":1DD8
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Tipo Conductor"
         Top             =   3240
         Width           =   4695
      End
      Begin Biometria2.CheckPeople CheckPeople1 
         Height          =   1155
         Left            =   480
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2037
         DefaultPicture  =   "conductor.frx":1DEF
         PermitirCancelar=   0   'False
         Sensibilidad    =   0
         Seguridad       =   0
         ImagenFondo     =   "conductor.frx":1E0B
      End
      Begin VB.ComboBox Cmbestado 
         Height          =   315
         ItemData        =   "conductor.frx":1E27
         Left            =   5280
         List            =   "conductor.frx":1E31
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Estado del conductor"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtpeso 
         Height          =   285
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   8
         ToolTipText     =   "Peso del conductor"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtlicencia 
         Height          =   285
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Número de licencia"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtestatura 
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Estatuta del conductor"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtedad 
         Height          =   285
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   6
         ToolTipText     =   "Edad del conductor"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Txtobservacion 
         Height          =   1965
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         ToolTipText     =   "Observaciones"
         Top             =   3720
         Width           =   4695
      End
      Begin VB.TextBox Txtcedula 
         Height          =   285
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Número de cédula del conductor"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtapellidos 
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Apellidos del conductor"
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox Txtnombres 
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   3
         ToolTipText     =   "Nombre del conductor"
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox Txtcodemp 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         ToolTipText     =   "Empresa en que trabaja el conductor"
         Top             =   2760
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Mskfecing 
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         ToolTipText     =   "Fecha de ingreso del conductor"
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "MMM/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecnac 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Fecha de nacimiento del conductor"
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "MMM/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport Rpt1 
         Left            =   5880
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Conductor:"
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
         TabIndex        =   36
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblCodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label6 
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
         Left            =   4560
         TabIndex        =   33
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Peso:"
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
         TabIndex        =   32
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Licencia:"
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
         Left            =   3600
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Estatura:"
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
         TabIndex        =   30
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Edad:"
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
         Left            =   5280
         TabIndex        =   29
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   28
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblempresa 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   3120
         TabIndex        =   20
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Cédula :"
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
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Apellidos:"
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
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nombres:"
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
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Fec. Nacimiento:"
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
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Ingreso:"
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
         TabIndex        =   26
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label16 
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
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      Visible         =   0   'False
      X1              =   6840
      X2              =   7920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6840
      X2              =   6840
      Y1              =   6120
      Y2              =   0
   End
End
Attribute VB_Name = "conductor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paso1 As Integer
'Dim ha As New HuellaDigital
Dim cEstado As String
Dim Respuesta_huella
Dim sEstado As String

Dim L_IntegracionProtech As String
Dim L_UrlApiInsConductor As String
Dim L_UrlApiUpdConductor As String
Dim L_UrlServer As String

Private Sub Trae_Estado(ByVal sCedula As String)
   fqry01 = "SELECT coestado  " _
            & "FROM coconductor " _
           & "WHERE cocedula = '" & sCedula & "'"
     
   Set Ftb01 = Dbs.Execute(fqry01)
     
   If Not Ftb01.EOF Then
      sEstado = Ftb01!coestado
      If Ftb01!coestado = "Activo" Then
         Cmbestado.text = Cmbestado.List(0)
      Else
         Cmbestado.text = Cmbestado.List(1)
      End If
   End If
End Sub


Private Sub Btnbuscar_Click()
   forma = "20"
   forma1 = "2"
   buscar.Show 1
   txtcedula_LostFocus
End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo err_Btngrabar_Click
   Dim lCedula, lNombre, lEstado As String
   
   If Trim(txtCedula.text) = "" Then
      MsgBox "Falta Ingresar el Número de la CEDULA", 48, "Advertencia"
      txtCedula.SetFocus
      Exit Sub
   End If
       
   If Trim(Txtapellidos.text) = "" Then
      MsgBox "Falta Ingresar los APELLIDOS", 48, "Advertencia"
      Txtapellidos.SetFocus
      Exit Sub
   End If
   If Trim(Txtnombres.text) = "" Then
      MsgBox "Falta Ingresar los NOMBRES", 48, "Advertencia"
      Txtnombres.SetFocus
      Exit Sub
   End If
   
   If Mskfecnac.ClipText = "" Then
      MsgBox "Falta Ingresar La FECHA DE NACIMIENTO", 48, "Advertencia"
      Mskfecnac.SetFocus
      Exit Sub
   End If
   
   If Mskfecing.ClipText = "" Then
      MsgBox "Falta Ingresar La FECHA DE INGRESO", 48, "Advertencia"
      Mskfecing.SetFocus
      Exit Sub
   End If
   
   If Trim(txtCodEmp.text) = "" Then
      MsgBox "Falta Ingresar la EMPRESA", 48, "Advertencia"
      txtCodEmp.SetFocus
      Exit Sub
   End If
   
   If Trim(Txtobservacion.text) = "" Then
      Txtobservacion.text = "NINGUNA"
   End If
       
   If Trim(txtedad.text) = "" Then
      MsgBox "Falta Ingresar la EDAD", 48, "Advertencia"
      txtedad.SetFocus
      Exit Sub
   End If
            
   If DateValue(Mskfecnac.text) > DateValue(Mskfecing.text) Then
      MsgBox "La Fecha de Nacimiento No Puede Ser Mayor A La Fecha de Ingreso", 48, "Advertencia"
      Mskfecnac.SetFocus
      Exit Sub
   End If
              
   If Trim(txtlicencia.text) = "" Then
      txtlicencia.text = 0
   End If
            
   If Trim(txtpeso.text) = "" Then
      txtpeso.text = 0
   End If
   
   If sEstado <> "" Then
     If sEstado <> Trim(Cmbestado.text) Then
        MsgBox "No Ha Ingresado Las Novedades", 48, "Advertencia"
        Exit Sub
     End If
  End If
   respuesta = MsgBox("Desea Grabar los Datos del Conductor?", 32 + vbYesNo, "Grabar")
      
   If respuesta = vbNo Then
      Exit Sub
   End If
       
   If Cmbestado.ListIndex = 0 Then
      cEstado = "Activo"
      sEstado = "1"
   Else
      cEstado = "Inactivo"
      sEstado = "0"
   End If
       
   
   Dbs.BeginTrans
   On Error GoTo err_transaccion
   
   
   x = InStr(1, Trim(Txtapellidos.text), " ", 1)        'Busca la posicion del espacio
   lCedula = Trim(txtCedula.text)
   lNombre = Trim(Mid(Txtnombres.text, 1, 25))
   
   If x <> 0 Then
      lApellido1 = Trim(Mid(Trim(Txtapellidos.text), 1, x - 1))
      lApellido2 = Trim(Mid(Trim(Txtapellidos.text), x + 1))
   Else
      lApellido1 = Trim(Txtapellidos.text)
      lApellido2 = ""
   End If
   lEstado = sEstado
   
   Call CreandoJson("C", lCedula, lNombre, lApellido1, lApellido2, lEstado, "", "", "", "")
   
   sBandera = CadenaJSON
   sParametro = ""
   
   ingreso = Format(Mskfecing.text, "MM/dd/yyyy")
   nacimiento = Format(Mskfecnac.text, "MM/dd/yyyy")
   If Btngrabar.Caption = "&Grabar" Then
      Dbs.Execute " INSERT INTO coconductor " _
                    & "(cocedula, coapellidos, conombres, " _
                    & "cofecing, cofecnac, cocodemp, coobserv, " _
                    & "coedad,  coestatura, colicencia, copeso, coestado, cotipoconductor) " _
                    & "VALUES ('" _
                    & txtCedula.text & "' ,'" & Txtapellidos.text & "','" & Txtnombres.text & "',TO_DATE('" _
                    & ingreso & "','MM/DD/YYYY'),TO_DATE('" & nacimiento & "','MM/DD/YYYY'),'" & txtCodEmp.text & "','" & Txtobservacion.text & "'," _
                    & txtedad.text & ", '" & txtestatura.text & "','" & txtlicencia.text & "','" & txtpeso.text & "','" & cEstado & "','" & Mid(CmbTipoConductor.text, 1, 1) & "')"
            
      sUrl = L_UrlApiInsConductor
      sEvento = "INS"
      
      
   Else
      Dbs.Execute " UPDATE coconductor " _
                 & "SET coapellidos = '" & Txtapellidos.text & "'," _
                 & "    conombres   = '" & Txtnombres.text & "', " _
                 & "    cocodemp    = '" & txtCodEmp.text & "', " _
                 & "    coobserv    = '" & Txtobservacion.text & "', " _
                 & "    cofecing    = TO_DATE('" & ingreso & "','MM/DD/YYYY'), " _
                 & "    cofecnac    = TO_DATE('" & nacimiento & "','MM/DD/YYYY'), " _
                 & "    coedad   =  " & txtedad.text & ", " _
                 & "    colicencia   =  '" & txtlicencia.text & "', " _
                 & "    copeso   =  '" & txtpeso.text & "', " _
                 & "    coestatura    = '" & txtestatura.text & "', " _
                 & "    coestado = '" & cEstado & "', " _
                 & "    cotipoconductor = '" & Mid(CmbTipoConductor.text, 1, 1) & "' " _
                 & "WHERE cocedula  = '" & txtCedula.text & "' "
         
      sUrl = L_UrlApiUpdConductor
      sEvento = "INS"
      'sParametro = "?sUsuario=sa1"
     
   End If
   
   If L_IntegracionProtech = "S" Then
      sRespuesta = Fn_Web_Sevices_Protech(L_UrlServer, sUrl, sEvento, sBandera, sParametro)
      
      If sRespuesta <> "0" Then
         sFechaSistemaProtech = Fn_Parametros("HORASIS", 1)
         
         QRY = "INSERT INTO logprotech(lotipo, lovalor, lodescripcion, lofecha)" _
         & "VALUES ('C','" & txtCedula.text & "','" & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta & "',TO_DATE('" & sFechaSistemaProtech & "','MM/dd/yyyy HH24:mi:ss'))"
         
         Dbs.Execute QRY
         MsgBox "Se presento error en el webservices de Protech por: " & Chr(13) & L_Mensaje_Fn_Web_Sevices_Protech & " - Error " & sRespuesta, 48, "Advertencia"
      End If
   End If
   
   Dbs.CommitTrans
   MsgBox "Conductor Ingresado", 64, "O.K."
   
   On Error GoTo err_Btngrabar_Click
   Btnlimpiar_Click
   'btnhuella.Enabled = True
  
   Exit Sub
err_Btngrabar_Click:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
err_transaccion:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

'Private Sub btnhuella_Click()
'On Error GoTo Errores
'
'    Dim l As Long
'
'    If Txtcedula.Text <> "" Then
'
'        Qry = " SELECT * " & _
'                " FROM coconductor " & _
'               " WHERE cocedula = '" & Txtcedula.Text & "' "
'
'        Set Tbl1 = Dbs.Execute(Qry)
'
'        If Not Tbl1.EOF Then
'            If Not IsNull(Tbl1!cohuella) Then
'
'              If G_rolusu = "ADMSIS" Then
'
'                If MsgBox("La huella ya fue ingresada, Desea ingresar de nuevo?", 64 + vbYesNo, "Error") = vbYes Then
'
'                    Qry = " UPDATE coconductor " & _
'                             " SET cohuella = '' " & _
'                           " WHERE cocedula = '" & Txtcedula.Text & "' "
'
'                    Dbs.Execute Qry
'                Else
'                    Tbl1.Close
'                    Exit Sub
'                End If
'
'              Else
'               Tbl1.Close
'               MsgBox "No tiene permiso para actualizar huellas", 48, "Advertencia"
'               Exit Sub
'              End If
'
'            End If
'
'
'        Else
'            Tbl1.Close
'            Exit Sub
'        End If
'        Tbl1.Close
'
'        If Not BIOSENSOR1.Inicializar(G_rutallave) Then
'            MsgBox "No se puede Activar el Lector.", 48, "Advertencia"
'            Exit Sub
'        End If
'
'        BIOSENSOR1.ActivarSensor Me.hWnd
'
'        If BIOSENSOR1.CapturarHuella(ha) Then
'
'            If Not IsEmpty(ha.Paquete) Then
'
'                If ha.Imagen <> 0 Then
'
'                    Qry = " SELECT cohuella " & _
'                            " FROM coconductor " & _
'                           " WHERE cocedula = '" & Txtcedula.Text & "' "
'
'                    Set Tbl1 = Dbs.Execute(Qry)
'
'                    'Tbl1.Edit
'                    'Tbl1(0).AppendChunk (ha.Paquete)
'                    'Tbl1.Update
'                    'Tbl1.Close
'                End If
'            End If
'
'            BIOSENSOR1.DesactivarSensor
'        End If
'    End If
'
'Exit Sub
'Errores:
'    MsgBox Err.Description, 16, "Error " & Err.Number
'End Sub


Private Sub btnhuella_Click()
   On Error GoTo Errores

    Dim l As Long
    
    If txtCedula.text <> "" Then
        Respuesta_huella = MsgBox("Desea Usar La Huella Principal(Si) o La Huella Alternativa(No)", 64 + vbYesNo, "Advertencia")
        
        If Respuesta_huella = vbYes Then
            QRY = "SELECT cohuella " & _
                  "FROM coconductor " & _
                  "WHERE cocedula = '" & txtCedula.text & "' "

        
            Set Tbl20 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
            If Not Tbl20.EOF Then
               If Not IsNull(Tbl20!cohuella) Then
                  If MsgBox("La Huella ya fue Ingresada, Desea Ingresar de Nuevo?", 64 + vbYesNo, "Cambio") = vbYes Then
                     QRY = " UPDATE coconductor SET cohuella = NULL " & _
                          " WHERE cocedula = '" & txtCedula.text & "' "

                     Dbs.Execute QRY

                     GoTo huella
                  Else
                     Exit Sub
                  End If
                  Tbl20.Close
               Else
                  GoTo huella
               End If
            Else
               Tbl20.Close
               Exit Sub
            End If
            Tbl20.Close
        
        ElseIf Respuesta_huella = vbNo Then
            QRY = "SELECT cohuealt " & _
                  "FROM coconductor " & _
                  "WHERE cocedula = '" & txtCedula.text & "' "

        
            Set Tbl20 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
            If Not Tbl20.EOF Then
               If Not IsNull(Tbl20!cohuealt) Then
                  If MsgBox("La Huella ya fue Ingresada, Desea Ingresar de Nuevo?", 64 + vbYesNo, "Cambio") = vbYes Then
                     QRY = " UPDATE coconductor SET cohuealt = NULL " & _
                          " WHERE cocedula = '" & txtCedula.text & "' "

                     Dbs.Execute QRY

                     GoTo huella
                  Else
                     Exit Sub
                  End If
                  Tbl20.Close
               Else
huella:
                  If G_ManLector = "S" Then
                     INICIALIZAR_LECTOR_HUELLA
                     CAPTURAR_HUELLA
                     FINALIZAR_LECTOR_HUELLA
                  End If
               End If
            Else
               Tbl20.Close
               Exit Sub
            End If
            Tbl20.Close
        End If
    End If
        
    Btnlimpiar_Click
    
    
Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
    On Error GoTo Btnimprimir_Click
    
    TITULO = "REPORTE DE CONDUCTORES"
    
'    Qry = "INSERT INTO optempo1 (t)"
'
'    "SELECT cocedula, conombres, coapellidos, colicencia, TO_CHAR(cofecnac,'MON/dd/yyyy') as cofecnac, TO_CHAR(cofecing,'MON/dd/yyyy') as cofecing, " _
'& "ROUND((SYSDATE - cofecnac)/365,0) as coedad , coestatura, coobserv, cocodemp, henombre, coestado " _
'& "FROM coconductor INNER JOIN ophojemp ON hecodigo = cocodemp"
    
    reporte = RpPath + "\conducto.rpt"
    Rpt1.Connect = G_Conexion_reportes
    Rpt1.ReportFileName = reporte
    
    Rpt1.Formulas(0) = "empresa = '" & G_NomEmpresa & "'"
    Rpt1.Formulas(1) = "titulo = '" & TITULO & "'"
    Rpt1.WindowState = crptMaximized
    Rpt1.Destination = 0
    Rpt1.Action = 1
    
    Exit Sub
Btnimprimir_Click:
  If Err.Number <> 20000 Then
  MsgBox Err.Description, 16, "Error " & Err.Number
  End If
  Exit Sub
End Sub

Private Sub Btnlimpiar_Click()
   Txtobservacion.text = ""
   txtCedula.text = ""
   Txtapellidos.text = ""
   Txtnombres.text = ""
   txtCedula.Enabled = True
   lblempresa.Caption = ""
   txtCodEmp.text = ""
   Mskfecing.Mask = ""
   Mskfecing.text = ""
   Mskfecing.Mask = "##/##/####"
   Mskfecnac.Mask = ""
   Mskfecnac.text = ""
   Mskfecnac.Mask = "##/##/####"
   Btngrabar.Caption = "&Grabar"
   txtedad.text = ""
   txtestatura.text = ""
   txtlicencia.text = ""
   txtpeso.text = ""
   txtlicencia.Enabled = True
   Mskfecing.Enabled = True
   Mskfecnac.Enabled = True
   txtCedula.SetFocus
   Mskfecing.text = Date
   Mskfecnac.text = Date
   txtedad.text = "0"
   txtestatura.text = "0"
   txtpeso.text = "0"
   Txtobservacion.text = "NINGUNA"
   If G_Sitio = "SERVIDOR" Then
      Btngrabar.Enabled = False
   End If
   Cmbestado.ListIndex = 0
   CmbTipoConductor.ListIndex = 0
   btnnovedades.Enabled = False
   sEstado = ""
End Sub

Private Sub btnnovedades_Click()
    
   sEstado = Cmbestado.text
   G_Tipo_Novedad = "CONDU"
   detvehic.Show 1
   Trae_Estado (txtCedula.text)
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Errores

   
   On Error GoTo Errores
   
   Mskfecing.text = Date
   Mskfecnac.text = Date
   Mskfecing.Enabled = False
   Mskfecnac.Enabled = False
   txtedad.text = "0"
   txtestatura.text = "0"
   txtpeso.text = "0"
   Txtobservacion.text = "NINGUNA"

   Cmbestado.ListIndex = 0
   Cmbestado.Enabled = True
   CmbTipoConductor.ListIndex = 0
   If G_ManLector = "S" Then
      connectdb = ";DATABASE=" & DbPathHuella & ";PWD=cta2000"
      
      Set dbs4 = OpenDatabase("", False, False, connectdb)
      
   End If
   
   L_IntegracionProtech = Fn_Parametros("INTPROTECH", 1)
  
   If L_IntegracionProtech = "" Then
      L_IntegracionProtech = "N"
   End If
   If L_IntegracionProtech = "S" Then
      L_UrlServer = Fn_Parametros("URLSERVER", 1)
      L_UrlApiInsConductor = Fn_Parametros("URLAPIINSC", 1)
      L_UrlApiUpdConductor = Fn_Parametros("URLAPIUPDC", 1)
   End If
   
Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   
End Sub

Private Sub Mskfecing_KeyPress(KeyAscii As Integer)
 If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecing_LostFocus()
  If Mskfecing.ClipText <> "" Then
     If Val(Left(Mskfecing.text, 2)) < 1 Or Val(Left(Mskfecing.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecing.SetFocus
        Exit Sub
     End If
  
     If Mskfecing.ClipText <> "" Then
        If Not IsDate(Mskfecing.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Mskfecing.SetFocus
           Exit Sub
        End If
     End If
     
  End If

End Sub

Private Sub Mskfecnac_KeyPress(KeyAscii As Integer)
 If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecnac_LostFocus()
  If Mskfecnac.ClipText <> "" Then
     If Val(Left(Mskfecnac.text, 2)) < 1 Or Val(Left(Mskfecnac.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecnac.SetFocus
        Exit Sub
     End If
  
     If Mskfecnac.ClipText <> "" Then
        If Not IsDate(Mskfecnac.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Mskfecnac.SetFocus
           Exit Sub
        End If
     End If
  End If

End Sub

Private Sub txtcedula_DblClick()
   Btnbuscar_Click
End Sub

Private Sub txtcedula_KeyPress(KeyAscii As Integer)
  If (KeyAscii% <> 8) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) Then
       KeyAscii% = 0
  End If
End Sub

Private Sub txtcedula_LostFocus()
   On Error GoTo Errores
   If txtCedula.text <> "" Then
      fqry02 = "SELECT coapellidos, conombres, cocodemp, coobserv, cofecing, cofecnac, coedad, colicencia, " _
                  & " coestatura, coestado, copeso, cotipoconductor, cofecinicer, cofecfincer " _
            & "FROM coconductor " _
            & "WHERE cocedula = '" & txtCedula.text & "' "
        
      Set Tbl1 = Dbs.Execute(fqry02)
      If Not Tbl1.EOF Then
       
         Txtapellidos.text = Tbl1!coapellidos
         Txtnombres.text = Tbl1!conombres
         txtCodEmp.text = Tbl1!cocodemp
         Txtobservacion.text = Tbl1!coobserv
         Mskfecing.text = Format(Tbl1!cofecing, "MM/DD/YYYY")
         Mskfecnac.text = Format(Tbl1!cofecnac, "MM/DD/YYYY")
         txtCedula.Enabled = False
         txtedad.text = Tbl1!coedad
         
         If IsNull(Tbl1!colicencia) Then
            txtlicencia.text = ""
         Else
            txtlicencia.text = Tbl1!colicencia
         End If
         
         If IsNull(Tbl1!coestatura) Then
            txtestatura.text = ""
         Else
            txtestatura.text = Tbl1!coestatura
         End If
         sEstado = Tbl1!coestado
         If IsNull(Tbl1!coestado) Then
           Cmbestado.ListIndex = 0
         Else
           If Tbl1!coestado = "Activo" Then
              Cmbestado.ListIndex = 0
           Else
              Cmbestado.ListIndex = 1
           End If
         End If
         
         If IsNull(Tbl1!cotipoconductor) Then
           CmbTipoConductor.ListIndex = 0
         Else
           If Tbl1!cotipoconductor = "P" Then
              CmbTipoConductor.ListIndex = 0
           Else
              CmbTipoConductor.ListIndex = 1
           End If
         End If
         
         
         If IsNull(Tbl1!copeso) Then
            txtpeso.text = ""
         Else
            txtpeso.text = Tbl1!copeso
         End If
         
         
         txtcodemp_LostFocus
         Btngrabar.Caption = "&Actualizar"
         btnnovedades.Enabled = True
         'btnhuella.Enabled = True
      Else
         sEstado = ""
         txtlicencia.Enabled = False
         txtlicencia.text = txtCedula.text
      End If
      Tbl1.Close
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Txtapellidos_KeyPress(KeyAscii As Integer)
 Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtedad_GotFocus()
   txtedad.SelStart = 0
   txtedad.SelLength = Len(txtedad.text)
End Sub

Private Sub txtedad_KeyPress(KeyAscii As Integer)
  If (KeyAscii% <> 8) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) Then
       KeyAscii% = 0
  End If
End Sub

Private Sub txtestatura_GotFocus()
   txtestatura.SelStart = 0
   txtestatura.SelLength = Len(txtestatura.text)
End Sub

Private Sub txtestatura_KeyPress(KeyAscii As Integer)
  If (KeyAscii% <> 8) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) And (KeyAscii% <> 46) Then
       KeyAscii% = 0
  End If

End Sub

Private Sub txtlicencia_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))
   If KeyAscii% = 124 Or KeyAscii% = 39 Then
      KeyAscii% = 0
   End If
End Sub

Private Sub Txtnombres_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
       Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub txtcodemp_DblClick()
   txtCodEmp.text = ""
   forma = "16"
   forma1 = "53"
   buscar.Show 1
   txtcodemp_LostFocus
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Then
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   On Error GoTo errTxtcodemp_LostFocus
   If txtCodEmp.text <> "" Then
     fqry01 = "SELECT henombre, heestado " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & txtCodEmp.text & "'"
             
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         If Ftb01!heestado = "Inactiva" Then
             MsgBox "Empresa Inactiva", 48, "Advertencia"
             txtCodEmp.SetFocus
             Ftb01.Close
             txtCodEmp.SelStart = 0
             txtCodEmp.SelLength = Len(txtCodEmp.text)
             Exit Sub
          End If
         lblempresa.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         Ftb01.Close
        
         fqry01 = "SELECT henombre, heestado " _
                  & "FROM ophojempcon " _
                 & "WHERE hecodigo = '" & txtCodEmp.text & "'"
             
         Set Ftb01 = Dbs.Execute(fqry01)
     
         If Not Ftb01.EOF Then
            If Ftb01!heestado = "Inactiva" Then
               MsgBox "Empresa Inactiva", 48, "Advertencia"
               txtCodEmp.SetFocus
               Ftb01.Close
               txtCodEmp.SelStart = 0
               txtCodEmp.SelLength = Len(txtCodEmp.text)
               Exit Sub
            End If
            lblempresa.Caption = Ftb01!henombre
            Ftb01.Close
         Else
            MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
            txtCodEmp.SetFocus
            Ftb01.Close
            txtCodEmp.SelStart = 0
            txtCodEmp.SelLength = Len(txtCodEmp.text)
            
            Exit Sub
         End If
      End If
   Else
      lblempresa.Caption = ""
   End If
   
   Exit Sub
   
errTxtcodemp_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Txtobservacion_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 38 Then
     KeyAscii% = 0
  End If

End Sub

Private Sub txtpeso_GotFocus()
   txtpeso.SelStart = 0
   txtpeso.SelLength = Len(txtpeso.text)
End Sub

Private Sub txtpeso_KeyPress(KeyAscii As Integer)
If (KeyAscii% <> 8) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) Then
       KeyAscii% = 0
  End If
End Sub
'funciones Huellero YR Dic/11/2015
Public Sub LIMPIAR_LECTOR_HUELLA()
    'Txtcedula = ""
    lblCodigo.Caption = ""
    lblNombreDedo.Caption = ""
    txtCedula.Enabled = True
    txtCedula.SetFocus
    CheckPeople1.Limpiar
End Sub

Public Sub INICIALIZAR_LECTOR_HUELLA()
    CheckPeople1.Inicializar
End Sub

Public Sub FINALIZAR_LECTOR_HUELLA()
    CheckPeople1.Finalizar
End Sub

Public Sub CAPTURAR_HUELLA()
   'Captura la huella
   CheckPeople1.Registrar
   lblCodigo.Caption = CheckPeople1.Dedo
   'lblNombreDedo.Caption = CheckPeople1.NombreDedo
      
   'Inserta el registro con la huella
   If Not IsEmpty(CheckPeople1.HuellaByte) Then
         
         QRY = " SELECT * FROM COCONDUCTOR " & _
               " WHERE COCEDULA = '" & txtCedula.text & "' "
         
         Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
         
         If Not Tbl21.EOF Then
            
            Tbl21.Close
            If Respuesta_huella = vbYes Then
        
               QRY = "UPDATE coconductor " _
                        & "SET CODEDO = " & lblCodigo.Caption & " " _
                        & "WHERE cocedula = '" & txtCedula.text & "' "
                        
            ElseIf Respuesta_huella = vbNo Then
               QRY = "UPDATE coconductor " _
                        & "SET CODEDALT = " & lblCodigo.Caption & " " _
                        & "WHERE cocedula = '" & txtCedula.text & "' "
            End If
            
            dbs4.Execute QRY
            
            If Respuesta_huella = vbYes Then
               QRY = " SELECT cohuella " & _
                     " FROM coconductor " & _
                     " WHERE cocedula = '" & txtCedula.text & "' "
            ElseIf Respuesta_huella = vbNo Then
               QRY = " SELECT cohuealt " & _
                     " FROM coconductor " & _
                     " WHERE cocedula = '" & txtCedula.text & "' "
            End If
            
            Set Tbl21 = dbs4.OpenRecordset(QRY, dbOpenDynaset)
            
            Tbl21.Edit
            Tbl21(0).AppendChunk (CheckPeople1.HuellaByte)
            Tbl21.Update
            'Tbl21.Close
'
'         Else
'            qry = " INSERT INTO coconductor (COCEDULA, CODEDO) " _
'                & " VALUES (" & Txtcedula & "," & CheckPeople1.Dedo & ")"
'
'            dbs4.Execute qry
'
'            If Not dbs4.RecordsAffected = 0 Then
'               qry = "SELECT * FROM coconductor WHERE cocedula = " & Txtcedula
'               Set RS = dbs4.OpenRecordset(qry, dbOpenDynaset)
'
'               If Not RS.EOF Then
'                  If Not IsEmpty(CheckPeople1.HuellaByte) Then
'                     RS.Edit
'                     RS("cohuella").AppendChunk (CheckPeople1.HuellaByte)
'                     RS.Update
'                     RS.Close
'                  End If
'               End If
'            End If
               
         End If
         Tbl21.Close
   End If
End Sub


