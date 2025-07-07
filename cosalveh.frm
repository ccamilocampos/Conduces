VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form cosalveh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desbloqueo de Venta de TUT"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "cosalveh.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   5280
      Picture         =   "cosalveh.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   3480
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
      Left            =   4320
      Picture         =   "cosalveh.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir recibo de conduce"
      Top             =   3480
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
      Left            =   3360
      Picture         =   "cosalveh.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar Informacion de Conduces Ingresados"
      Top             =   3480
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
      Left            =   2400
      Picture         =   "cosalveh.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Desbloquea la placa para la venta de TUT"
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "Times New Roman"
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
      TabIndex        =   14
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TxtCedula2 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Número de cedula segundo conductor"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtCedula1 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Número de cedula primer conductor"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Txtnumcon 
         Alignment       =   2  'Center
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Número de tasa de uso"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txtcodrut 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Código de la ruta "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Txtcodemp 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de Transporte "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Txtnumpla 
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
         Left            =   960
         MaxLength       =   7
         TabIndex        =   0
         ToolTipText     =   "Placa del vehículo"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txtnumint 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Número interno del vehículo"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Txttipveh 
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
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Clase de servicio del vehículo"
         Top             =   840
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Mskfecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Fecha y hora de salida"
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "MMM/dd/yyyy HH:mm"
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label LblNombre2 
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
         Left            =   2280
         TabIndex        =   28
         ToolTipText     =   "Nombre del segundo conductor"
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula 2:"
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
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label LblNombre1 
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
         Left            =   2280
         TabIndex        =   26
         ToolTipText     =   "Nombre del primer conductor"
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula 1:"
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
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   855
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
         Left            =   4200
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "TUT:"
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
         Left            =   2280
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ruta:"
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
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Lblnomrut 
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
         Left            =   2280
         TabIndex        =   21
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1320
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
         Left            =   2280
         TabIndex        =   19
         Top             =   1320
         Width           =   4695
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
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Núm.Int.:"
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
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Lbltipveh 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   7200
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "cosalveh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Cedula1 As String
Dim L_Cedula2 As String

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
   
   If Trim(Txtnumpla.text) = "" Then
      MsgBox "Falta ingresar el numero de placa del vehículo", 48, "Advertencia"
      Txtnumcon.Enabled = True
      Txtnumpla.SetFocus
      Exit Sub
   End If
   
   
   If Txtnumpla.Enabled = False Then
      respuesta = MsgBox("Desea desbloquear el vehículo para la venta de TUT", 32 + vbYesNo, "GRABAR")
  
      If respuesta = vbNo Then
         Exit Sub
      End If
      
      Dbs.BeginTrans
      On Error GoTo Err_Transac
      
      QRY = "DELETE FROM cosalveh " _
           & "WHERE svnumero = " & Trim(Txtnumcon.text) & " " _
             & "AND svplaca = '" & Trim(Txtnumpla.text) & "' " _
             & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "

      Dbs.Execute QRY
      
      QRY = "UPDATE ophojveh " _
             & "SET hvestviaje = 'N' " _
           & "WHERE hvplaca = '" & Trim(Txtnumpla.text) & "'"
         
      Dbs.Execute QRY
      If L_Cedula1 <> "-" Then
         QRY = "UPDATE coconductor " _
                & "SET coestviaje = 'N' " _
              & "WHERE cocedula = '" & L_Cedula1 & "'"
         
         Dbs.Execute QRY
      End If
      
      If L_Cedula2 <> "-" Then
         QRY = "UPDATE coconductor " _
                & "SET coestviaje = 'N' " _
              & "WHERE cocedula = '" & L_Cedula2 & "'"
         
         Dbs.Execute QRY
      
      End If
      
      'Camilo Campos 06/05/2025. Al dia de hoy esta funcionalidad no termina por desbloquear la TU.
      'para ello, en conjunto con el ingeniero de la terminal Nelson Ovalle nos comento que el proceso manual que realiza
      'se basa en poner estado "I" el último despacho del vehiculo. Según él es normal cuando un bus termina varado y nunca se cierra dicho despacho
      'Por ello buscamos el último despacho de la placa y procedemos a cambiar el estado
      
      QRY = "SELECT MAX(denumero) as despacho FROM codespachos WHERE deplaca = '" & Txtnumpla.text & "'"
      
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
      
        QRY = "UPDATE codespachos SET deestado = 'I' WHERE denumero = " & Tbl1!despacho
        Dbs.Execute (QRY)
        
      Else
      
        MsgBox "No existe un despacho para inhabiltar", 16, "Error " & Err.Number
        Dbs.RollbackTrans
        Exit Sub
      
      End If
      
      
      
      Dbs.CommitTrans
      MsgBox "El vehículo de desbloqueo exitósamente", 64, "O.K."
      
      On Error GoTo Errores
      Btnlimpiar_Click
   End If
   Exit Sub
   
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
Err_Transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
End Sub

Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   Txtnumpla.Enabled = True
   Txtnumpla.text = ""
   Txtnumcon.text = ""
   Mskfecha.Mask = ""
   Mskfecha.text = ""
   Mskfecha.Mask = "##/##/#### ##:##"
   TxtNumInt.text = ""
   Txttipveh.text = ""
   Lbltipveh.Caption = ""
   Txtcodemp.text = ""
   Lblnomemp.Caption = ""
   Txtcodrut.text = ""
   Lblnomrut.Caption = ""
   TxtCedula1.text = ""
   LblNombre1.Caption = ""
   TxtCedula2.text = ""
   LblNombre2.Caption = ""
   Txtnumpla.SetFocus
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error "
   Exit Sub
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Txtnumpla_GotFocus()
   Txtnumpla.SelStart = 0
   Txtnumpla.SelLength = Len(Txtnumpla.text)
End Sub

Private Sub Txtnumpla_KeyPress(KeyAscii As Integer)
'   Char = Chr(KeyAscii%)
'   KeyAscii = Asc(UCase(Char))  ' mayuscula
'   longitud = Len(Txtnumpla.Text)
'   longitud = longitud + 1
'   If longitud <= 2 Then
'     If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'         Exit Sub
'     Else
'         KeyAscii% = 0
'         Exit Sub
'      End If
'   ElseIf longitud > 3 Then
'      If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
'         Exit Sub
'      Else
'         KeyAscii% = 0
'         Exit Sub
'      End If
'   End If
'   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
'      Exit Sub
'   Else
'      'KeyAscii% = 0
'   End If
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
   
End Sub

Private Sub Txtnumpla_LostFocus()
 
   On Error GoTo Message_Next
 
   If Txtnumpla.text <> "" Then
      sFechaSistema = Format(Fn_Parametros("HORASIS", 1), "MM/DD/YYYY")
      sFecha = DateAdd("d", -60, sFechaSistema)
      TxtCedula1.text = ""
      TxtCedula2.text = ""
      fqry01 = "SELECT hvnumint, hvcodemp, hvtipveh, hvestado " _
               & "FROM ophojveh " _
              & "WHERE hvplaca = '" & Txtnumpla.text & "' "
  
      Set Ftb01 = Dbs.Execute(fqry01)
      If Not Ftb01.EOF Then
         If Ftb01!hvestado = "Inactivo" Then
            MsgBox "EL VEHICULO SE ENCUENTRA INACTIVO," & vbCrLf & "  NO SE LE PUEDE VENDER TASAS DE USO", 16, "AVISO"
            Txtnumpla.SetFocus
            Exit Sub
         End If
         TxtNumInt.text = Ftb01!hvnumint
         Txtcodemp.text = Ftb01!hvcodemp
         Txttipveh.text = Ftb01!hvtipveh
         Ftb01.Close
         'busca que el registro exista en la tabla cosalveh
         QRY = "SELECT svnumero, svfecha " _
               & "FROM cosalveh " _
              & "WHERE svplaca = '" & Trim(Txtnumpla.text) & "' " _
                & "AND svterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
               
         Set Tbl1 = Dbs.Execute(QRY)
          
         If Not Tbl1.EOF Then
            Txtnumcon.text = Tbl1!svnumero
            Mskfecha.text = Format(Tbl1!svfecha, "MM/dd/yyyy HH:mm")
            sFecha = Mid(Tbl1!svfecha, 1, 10)
            Tbl1.Close
            
            
            'busca el codigo de la ruta de la TUT
            QRY = "SELECT cocodrut, NVL(cocedula,'-') AS cocedula, NVL(cocedula1,'-') AS cocedula1  " _
                  & "FROM coconduc " _
                 & "WHERE conumero = " & Trim(Txtnumcon.text) & " " _
                   & "AND cofecsal = TO_DATE('" & sFecha & "', 'MM/dd/yyyy' )" _
                   & "AND coterminal = '" & Mid(G_Tipo_Terminal, 1, 1) & "' "
               
            Set Tbl1 = Dbs.Execute(QRY)
          
            If Not Tbl1.EOF Then
               Txtcodrut.text = Tbl1(0)
               L_Cedula1 = Trim(Tbl1!cocedula)
               L_Cedula2 = Trim(Tbl1!cocedula1)
               If Trim(Tbl1!cocedula) <> "-" Then
                  TxtCedula1.text = Tbl1!cocedula
               End If
               
               If Trim(Tbl1!cocedula1) <> "-" Then
                  TxtCedula2.text = Tbl1!cocedula1
               End If
            End If
            Tbl1.Close
         Else
            Tbl1.Close
            
            QRY = "SELECT conumero, cohorsal, cocodrut, NVL(cocedula,'-') AS cocedula, NVL(cocedula1,'-') AS cocedula1 " _
                  & "FROM coconduc " _
                 & "WHERE cohorsal = (SELECT MAX(cohorsal) " _
                                     & "FROM coconduc " _
                                    & "WHERE cofecsal BETWEEN TO_DATE('" & sFecha & "','MM/dd/yyyy')  " _
                                      & "AND TO_DATE('" & sFechaSistema & "','MM/dd/yyyy') " _
                                      & "AND coplaca = '" & Txtnumpla.text & "' " _
                                      & "AND coestado = 'Activo' " _
                                      & "AND cooritra ='O') " _
                  & "AND coplaca = '" & Txtnumpla.text & "' " _
                  & "AND coestado = 'Activo' " _
                  & "AND cooritra ='O'"
            
            Set Tbl1 = Dbs.Execute(QRY)
          
            If Not Tbl1.EOF Then
               Txtnumcon.text = Tbl1!conumero
               Mskfecha.text = Format(Tbl1!cohorsal, "MM/dd/yyyy HH:mm")
               Txtcodrut.text = Tbl1!cocodrut
               L_Cedula1 = Trim(Tbl1!cocedula)
               L_Cedula2 = Trim(Tbl1!cocedula1)
               If Trim(Tbl1!cocedula) <> "-" Then
                  TxtCedula1.text = Tbl1!cocedula
               End If
               
               If Trim(Tbl1!cocedula1) <> "-" Then
                  TxtCedula2.text = Tbl1!cocedula1
               End If
            Else
               Tbl1.Close
               MsgBox "El vehículo no ha comprado TUT en los ultimos 60 días", 48, "Advertencia"
               Txtnumpla.SetFocus
               Exit Sub
            End If
         End If
         
         fqry01 = "SELECT heestado, henombre " _
                 & "FROM ophojemp " _
                & "WHERE hecodigo = '" & Txtcodemp.text & "' "
        
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
            If Ftb01!heestado = "Inactiva" Then
               Ftb01.Close
               MsgBox "La empresa esta inactiva", 48, "Advertencia"
               Txtnumpla.SetFocus
               Exit Sub
            End If
              
            Lblnomemp.Caption = Ftb01!henombre
         End If
         Ftb01.Close
         
         fqry01 = "SELECT rudestino " _
                  & "FROM opruta " _
                 & "WHERE rucodigo = '" & Txtcodrut.text & "' "
             
         Set Ftb01 = Dbs.Execute(fqry01)
     
         If Not Ftb01.EOF Then
            Lblnomrut.Caption = Ftb01!rudestino
         End If
         Ftb01.Close
         
         fqry01 = "SELECT dsdes " _
                  & "FROM gesuptip, gedetsuptip " _
                 & "WHERE dscoddet = '" & Txttipveh.text & "' " _
                   & "AND dscodtip = stcodtip " _
                   & "AND stdes = 'TIPOS DE VEHICULO' "
               
         Set Ftb01 = Dbs.Execute(fqry01)
         If Not Ftb01.EOF Then
            Lbltipveh.Caption = Ftb01!dsdes
         End If
         Ftb01.Close
         
         If Trim(TxtCedula1.text) <> "" Then
            LblNombre1.Caption = Traer_Nombre_Conductor(TxtCedula1.text)
         End If
         If Trim(TxtCedula2.text) <> "" Then
            LblNombre2.Caption = Traer_Nombre_Conductor(TxtCedula2.text)
         End If
         
         Txtnumpla.Enabled = False
      Else
         MsgBox "El vehículo no existe", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
   End If
  
salir:
   Exit Sub
   
Message_Next:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub


