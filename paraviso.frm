VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form paraviso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descontar Avisos Parqueadero"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7410
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
      Left            =   6480
      Picture         =   "paraviso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir de la forma actual"
      Top             =   2880
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
      Left            =   6480
      Picture         =   "paraviso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   2040
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
      Left            =   6480
      Picture         =   "paraviso.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir recibo de conduce"
      Top             =   1200
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
      Left            =   6480
      Picture         =   "paraviso.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ingresa el registro a la base de datos"
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid GrdParque 
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Vehículo"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Txtcodemp 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de Transporte (presione DblClick para ayuda)"
         Top             =   720
         Width           =   855
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
         Left            =   120
         MaxLength       =   7
         TabIndex        =   2
         ToolTipText     =   "Placa del vehículo"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Txtnumint 
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Número interno del vehículo"
         Top             =   720
         Width           =   735
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
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1575
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
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Núm.Int."
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
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   3600
   End
End
Attribute VB_Name = "paraviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DIB_GRILLA()
   GrdParque.rows = 1
   GrdParque.Cols = 9
   GrdParque.Clear
   
   s$ = "|Consecutivo|Placa|Avisos|Fecha Ingreso|Fecha Salida|Valor|Desconto|Tipo"
   GrdParque.FormatString = s$
   GrdParque.Row = 0
   GrdParque.Col = 0
   GrdParque.ColWidth(0) = 500
   GrdParque.ColWidth(1) = 1200
   GrdParque.ColWidth(2) = 1000
   GrdParque.ColWidth(3) = 1000
   GrdParque.ColWidth(4) = 1500
   GrdParque.ColWidth(5) = 1500
   GrdParque.ColWidth(6) = 1000
   GrdParque.ColWidth(7) = 1100
    GrdParque.ColWidth(8) = 1500
                    
   For i = 0 To GrdParque.Cols - 1
       GrdParque.Col = i
       GrdParque.CellFontBold = 1
   Next i
End Sub

Private Sub Traer_Datos()
   Dim sDatos As String
   On Error GoTo Errores
   
   DIB_GRILLA
   scero = 0
   
   'Sebastian Rondon - 08/23/2023
   'Se busca parametro donde se establece la fecha a empezar a buscar permanencias
   QRY = "SELECT psval FROM geparsis WHERE pscod = 'FECPARPRO' "
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
    fecpar = Tbl1!psval
   End If
   
   fqry01 = "SELECT panumero, paplaca, pafecini, " _
                 & "pafecsal, pavalor, paaviso, " _
                 & "paestado, padesavi, patipo " _
            & "FROM coparqueop " _
           & "WHERE paplaca = '" & Txtnumpla.text & "' " _
             & "AND pavalor <> " & scero & " " _
             & "AND pafecini >= TO_DATE('" & fecpar & "','MM/DD/YYYY') " _
             & "AND paestado = 'A' ORDER BY  paaviso DESC, panumero "

   Set Ftb01 = Dbs.Execute(fqry01)
   sTotal = 0
   Do While Not Ftb01.EOF
      Txtnumpla.Enabled = False
      sValor = Ftb01!pavalor
      sAviso = "0"
      If Not IsNull(Ftb01!paaviso) Then
         sAviso = CStr(Ftb01!paaviso)
      End If
      sDesAvi = ""
      If Not IsNull(Ftb01!padesavi) Then
         sDesAvi = Ftb01!padesavi
      End If
      
      If Ftb01!patipo = "P" Then
         sTipo = "Permanencia"
      ElseIf Ftb01!patipo = "V" Then
         sTipo = "Vigencia"
      End If
      
      
      entry = "" + Chr(9) + CStr(Ftb01!panumero) + Chr(9) + Ftb01!paplaca + Chr(9) + _
               CStr(Ftb01!paaviso) + Chr(9) + Format(Ftb01!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
               Format(Ftb01!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + CStr(sValor) + Chr(9) + _
               sDesAvi + Chr(9) + sTipo
      GrdParque.AddItem entry            ' Agrega la multa.
      
      Ftb01.MoveNext
   Loop
   Ftb01.Close
   
   
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
End Sub

Private Sub BtnGrabar_Click()
   On Error GoTo Errores
   
   If Trim(Txtnumpla.text) = "" Then
      MsgBox "Falta ingresar el número de la placa", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
      
   End If
   
   If GrdParque.rows = 1 Then
      MsgBox "No hay registros para grabar", 48, "Advertencia"
      Exit Sub
   End If
   
   seleccion = "0"
   For i = 1 To GrdParque.rows - 1
      GrdParque.Row = i
      GrdParque.Col = 0
      
      If GrdParque.text = "->" Then
         seleccion = "1"
         Exit For
      End If
   Next i
       
   If seleccion = "0" Then
      MsgBox "No se ha seleccionado ningún registro", 48, "Advertencia"
      GrdParque.Row = 1
      GrdParque.Col = 0
      GrdParque.SetFocus
      Exit Sub
   End If
   
   rtamsg = MsgBox("Esta seguro de grabar la los registros?", 36, "Confirmación")
      
   If rtamsg = vbNo Then
      Exit Sub
   End If
   sFechaSistema = Fn_Parametros("HORASIS", 1)
   sFecha = Format(sFechaSistema, "MM/dd/yyyy")
   sHora = Format(sFechaSistema, "hh:mm")
   
   Dbs.BeginTrans
   paraviso.Enabled = False
   On Error GoTo Err_Transac
   
   For i = 1 To GrdParque.rows - 1
      GrdParque.Col = 0
      GrdParque.Row = i
      
      If GrdParque.text = "->" Then
         sNumero = GrdParque.TextMatrix(i, 1)     'numero tiquete
         suno = 1
         QRY = "UPDATE coparqueop " _
             & "SET paaviso =  paaviso - " & suno & ", " _
             & "    padesavi = 'S' " _
           & "WHERE panumero = " & sNumero & ""
      
         Dbs.Execute QRY
         
         QRY = "INSERT INTO coaudparqueo (" _
                           & "applaca, apfecha, aphora, " _
                           & "apnumparqueo, apfuncionario) " _
                  & "VALUES('" _
                  & Txtnumpla.text & "', TO_DATE('" & sFecha & "','mm/dd/yyyy'), TO_DATE('" & sHora & "','hh24:mi'), " _
                  & sNumero & ",'" & Login & "')"
         
         Dbs.Execute QRY
         
      End If
   Next i
   
   Dbs.CommitTrans
   On Error GoTo Errores
   paraviso.Enabled = True
   MsgBox "Los registros se grabaron exitosamente", 64, "O.K."
   Btnlimpiar_Click
   
   
   
   Exit Sub
   
Err_Transac:
   Dbs.RollbackTrans
Errores:
   paraviso.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   Txtnumpla.Enabled = True
   Txtnumpla.text = ""
   Txtnumint.text = ""
   Txtcodemp.text = ""
   Lblnomemp.Caption = ""
   
   DIB_GRILLA
   
   Txtnumpla.SetFocus
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   DIB_GRILLA
End Sub

Private Sub GrdParque_DblClick()
   
   
   If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
      
'      For i = 1 To GrdParque.Rows - 1
'         GrdParque.TextMatrix(i, 0) = ""
'      Next i
      
      
      
      If CDbl(GrdParque.TextMatrix(GrdParque.Row, 3)) = 0 Then
         MsgBox "El número de avisos del parqueadero es de cero," & Chr(13) & "               no se puede descontar un aviso", 48, "Advertencia"
         Exit Sub
      End If
      
      If GrdParque.TextMatrix(GrdParque.Row, 7) = "S" Then
         MsgBox "Al parqueadero ya se le desconto un aviso", 48, "Advertencia"
'         Exit Sub
      End If
      
      If GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = ""
      Else
         GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
      End If
      
      GrdParque.Col = 5
      
   ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
      GrdParque.TextMatrix(GrdParque.Row, 0) = ""
      GrdParque.Col = 5
      
   End If
   
   GrdParque.Col = 0
   If GrdParque.rows - 1 > GrdParque.Row Then
      GrdParque.Row = GrdParque.Row + 1
   Else
      
   End If
End Sub

Private Sub GrdParque_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If KeyAscii% = 32 Then
      If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
         GrdParque.Col = 5
         If mskafac.text = "" Then
            mskafac.text = GrdParque.text
         Else
            mskafac.text = CDbl(mskafac.text) + CDbl(GrdParque.text)
         End If
         
      ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = ""
         GrdParque.Col = 5
         mskafac.text = CDbl(mskafac.text) - CDbl(GrdParque.text)
      End If
      
      GrdParque.Col = 0
      If GrdParque.rows - 1 > GrdParque.Row Then
         GrdParque.Row = GrdParque.Row + 1
      Else
         'TxtObserv.SetFocus
      End If
   End If
   
End Sub


Private Sub Txtnumpla_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtnumpla_KeyPress(KeyAscii As Integer)
'   Char = Chr(KeyAscii%)
'   KeyAscii = Asc(UCase(Char))  ' mayuscula
'   longitud = Len(Txtnumpla.Text)
'   longitud = longitud + 1
'   If longitud <= 2 Then
'      If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'
'          Exit Sub
'      Else
'          'MsgBox "En El Numero De Placa las Tres Primeras Deben Ser Letras", 48, "Advertencia"
'          KeyAscii% = 0
'
'          Exit Sub
'      End If
'   ElseIf longitud > 3 Then
'      If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
'
'         If longitud = 6 Then
'            If KeyAscii% <> 8 Then
'            Txtnumpla_LostFocus
'            GrdParque.SetFocus
'            End If
'         End If
'         Exit Sub
'      Else
'         If longitud >= 7 Then
'            Txtnumpla.Text = ""
'            Exit Sub
'         End If
'         'MsgBox "En El Numero De Placa las Tres Ultimas Deben Ser Numeros", 48, "Advertencia"
'         KeyAscii% = 0
'         Exit Sub
'      End If
'   End If
'   If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
'       Exit Sub
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
      'If Len(Txtnumpla.Text) = 6 Then
         fqry01 = "SELECT * " _
                  & "FROM ophojveh " _
                 & "WHERE hvplaca = '" & Txtnumpla.text & "' "
                
         Set Tbl1 = Dbs.Execute(fqry01)
         If Not Tbl1.EOF Then
            Txtnumint.text = Tbl1!hvnumint
            Txtcodemp.text = Tbl1!hvcodemp
            Tbl1.Close
            fqry01 = "SELECT * " _
                    & "FROM ophojemp " _
                   & "WHERE hecodigo = '" & Txtcodemp.text & "' "
           
            Set Tbl1 = Dbs.Execute(fqry01)
            If Not Tbl1.EOF Then
               Lblnomemp.Caption = Tbl1!henombre
            End If
            Tbl1.Close
            Traer_Datos
            'Traer_Datos_Descuentos
         Else
            Tbl1.Close
            MsgBox "El vehículo no existe", 48, "Advertencia"
            Txtnumpla.SetFocus
            Exit Sub
         End If
      'End If
   End If
salir:
   Exit Sub
   
Message_Next:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub


