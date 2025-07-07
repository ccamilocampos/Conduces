VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form comulcan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Evasiones y Elusiones Canceladas"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6720
      Picture         =   "comulcan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterio de Búsqueda"
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
      TabIndex        =   19
      Top             =   960
      Width           =   2295
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
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
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
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Value           =   -1  'True
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Optcriterio 
         Caption         =   "Cajero"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valor de Búsqueda"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   960
      Width           =   3975
      Begin VB.TextBox Txtvalbus 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Lblvalbus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   2415
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
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6375
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Fecha en la que empiece a buscar"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mskfecfin 
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         ToolTipText     =   "Fecha hasta donde quiere que busque"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mmm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1335
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
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   6375
      Begin Crystal.CrystalReport rptterminal 
         Left            =   600
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         DiscardSavedData=   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Grdconmul 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
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
      Left            =   6720
      Picture         =   "comulcan.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Informacion en la Base de Datos"
      Top             =   480
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
      Left            =   6720
      Picture         =   "comulcan.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir de la Forma Actual"
      Top             =   4680
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
      Left            =   6720
      Picture         =   "comulcan.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Limpia la forma actual"
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
      Left            =   6720
      Picture         =   "comulcan.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   3000
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
      Left            =   6720
      Picture         =   "comulcan.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Totaliza todos los netos a pagar que se encuentran en la grilla"
      Top             =   2160
      Width           =   855
   End
   Begin MSMask.MaskEdBox Msktotal 
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "Valor total de las multas que se encuentran en la grilla"
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   -2147483633
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0;($#,##0)"
      PromptChar      =   "_"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   6600
      X2              =   6600
      Y1              =   5520
      Y2              =   0
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   7560
      Y1              =   5520
      Y2              =   5520
   End
End
Attribute VB_Name = "comulcan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function DIB_GRILLA()
     
     Grdconmul.rows = 1
     Grdconmul.Cols = 17
     Grdconmul.Clear
     
     
     s$ = "Número  |Empresa  |Cód. Infracción   |Nom. Notificado  |Núm. Placa  |Núm. Int| Fec. Evasion  |Hora |Aux. Control |Rec. Pago |Descuento |Valor Descuento |Neto Pagar |Turno |Cajero | Estado|Saldo|Caja "
     Grdconmul.FormatString = s$
     Grdconmul.Row = 0
     Grdconmul.Col = 0
     Grdconmul.ColWidth(0) = 1000
     Grdconmul.ColWidth(1) = 1400
     Grdconmul.ColWidth(2) = 1400
     Grdconmul.ColWidth(3) = 2500
     Grdconmul.ColWidth(4) = 1200
     Grdconmul.ColWidth(5) = 1200
     Grdconmul.ColWidth(6) = 1200
     Grdconmul.ColWidth(7) = 1200
     Grdconmul.ColWidth(8) = 1200
     Grdconmul.ColWidth(9) = 1200
     Grdconmul.ColWidth(10) = 1200
     Grdconmul.ColWidth(11) = 1500
     Grdconmul.ColWidth(12) = 1200
     Grdconmul.ColWidth(13) = 1200
     Grdconmul.ColWidth(14) = 1200
     Grdconmul.ColWidth(15) = 1200
     Grdconmul.ColWidth(16) = 1200
     Grdconmul.ColWidth(17) = 1200
     
     For i = 0 To Grdconmul.Cols - 1

        Grdconmul.Col = i
        Grdconmul.CellFontBold = 1
     Next i
End Function


Private Sub BtnBuscar_Click()
   On Error GoTo Err_Btnbuscar_Click
   
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
   
   If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
      MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
      Mskfecini.SetFocus
      Exit Sub
   End If
   
   If Optcriterio(2).Value = False Then
      If Trim(Txtvalbus.text) = "" Then
         MsgBox "Falta Ingresar el Valor de Busqueda", 48, "Advertencia"
         Txtvalbus.SetFocus
         Exit Sub
      End If
   End If
   
   
   DIB_GRILLA
   
   Fecha1 = Mskfecini.text
   Fecha2 = Mskfecfin.text
   Msktotal.text = ""
   
   estado = "Pagada"
   
   If Optcriterio(0).Value = True Then
      QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
               & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt " _
            & "FROM opmultas, ophojveh " _
           & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
             & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
             & "AND munumpla = hvplaca " _
             & "AND muestado = '" & estado & "' " _
             & "AND muturno = " & Txtvalbus.text & " " _
        & "ORDER BY munumero ASC "
   ElseIf Optcriterio(1).Value = True Then
      QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
            & "FROM opmultas, ophojveh " _
           & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
             & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
             & "AND munumpla = hvplaca " _
             & "AND muestado = '" & estado & "' " _
             & "AND muoperar = '" & Txtvalbus.text & "' " _
        & "ORDER BY munumero ASC "
  
   ElseIf Optcriterio(2).Value = True Then
      QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
            & "FROM opmultas, ophojveh " _
           & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
             & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
             & "AND munumpla = hvplaca " _
             & "AND muestado = '" & estado & "' " _
        & "ORDER BY munumero ASC "
   ElseIf Optcriterio(3).Value = True Then
      QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
            & "FROM opmultas, ophojveh " _
           & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
             & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
             & "AND munumpla = hvplaca " _
             & "AND muestado = '" & estado & "' " _
             & "AND munomcaj = '" & Txtvalbus.text & "' " _
        & "ORDER BY munumero ASC "
   
   End If
   
   Set Tbl1 = Dbs.Execute(QRY)
   total = 0
   comulcan.Enabled = False
   Do While Not Tbl1.EOF
   
      QRY = "SELECT henombre "
      If Len(Tbl1!mucodemp) = 4 Then
         QRY = QRY & "FROM ophojempcon  "
      Else
         QRY = QRY & "FROM ophojemp  "
      End If
      QRY = QRY & "WHERE hecodigo = '" & Tbl1!mucodemp & "'"
          
      Set Tbl2 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         sEmpresa = Tbl2!henombre
      End If
      Tbl2.Close
      
      If IsNull(Tbl1(16)) Then
         saldo = "0"
      Else
         saldo = str(Tbl1(16))
      End If
        
      QRY = "SELECT psdes FROM geparsis WHERE pscod = '" & Tbl1(2) & "'"
        
      Set Ftb01 = Dbs.Execute(QRY)
        
      If Not Ftb01.EOF Then
         sDesInfr = Ftb01(0)
      Else
         sDesInfr = ""
      End If
      Ftb01.Close
        
      entry = str(Tbl1(0)) + Chr(9) + sEmpresa + Chr(9) + sDesInfr + Chr(9) + _
              Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + str(Tbl1(13)) + Chr(9) + _
              str(Tbl1(5)) + Chr(9) + str(Tbl1(6)) + Chr(9) + Tbl1(7) + Chr(9) + _
              str(Tbl1(8)) + Chr(9) + str(Tbl1(9)) + Chr(9) + str(Tbl1(10)) + Chr(9) + _
              str(Tbl1(11)) + Chr(9) + str(Tbl1(14)) + Chr(9) + Tbl1(15) + Chr(9) + Tbl1(12) + Chr(9) + saldo + Chr(9) + Tbl1(17)
              
      Grdconmul.AddItem entry
      total = total + Val(Tbl1(11))
      Tbl1.MoveNext
      bandera = 1
   Loop
    Tbl1.Close
    comulcan.Enabled = True
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 16, "Busqueda"
      Exit Sub
   End If
   Call Aplicar_Perfil(Me, "comulcan")
   Msktotal.text = total
salir:
  Exit Sub
  
Err_Btnbuscar_Click:
  comulcan.Enabled = True
  MsgBox Err.Description, 16, "Error " & Err.Number
  Resume salir
Resume
End Sub

Private Sub BtnExportar_Click()
    On Error GoTo Err_Btnbuscar_Click
    
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
     
     If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
        MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
        Mskfecini.SetFocus
        Exit Sub
     End If
     
     If Optcriterio(2).Value = False Then
        If Trim(Txtvalbus.text) = "" Then
           MsgBox "Falta Ingresar el Valor de Busqueda", 48, "Advertencia"
           Txtvalbus.SetFocus
           Exit Sub
        End If
     End If
     
     
     DIB_GRILLA
     
     Fecha1 = Mskfecini.text
     Fecha2 = Mskfecfin.text
     Msktotal.text = ""
     
     estado = "Pagada"
     
     If Optcriterio(0).Value = True Then
        QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                 & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt " _
              & "FROM opmultas, ophojveh " _
             & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
               & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
               & "AND munumpla = hvplaca " _
               & "AND muestado = '" & estado & "' " _
               & "AND muturno = " & Txtvalbus.text & " " _
          & "ORDER BY munumero ASC "
     ElseIf Optcriterio(1).Value = True Then
        QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                  & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
              & "FROM opmultas, ophojveh " _
             & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
               & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
               & "AND munumpla = hvplaca " _
               & "AND muestado = '" & estado & "' " _
               & "AND muoperar = '" & Txtvalbus.text & "' " _
          & "ORDER BY munumero ASC "
    
     ElseIf Optcriterio(2).Value = True Then
        QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                  & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
              & "FROM opmultas, ophojveh " _
             & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
               & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
               & "AND munumpla = hvplaca " _
               & "AND muestado = '" & estado & "' " _
          & "ORDER BY munumero ASC "
     ElseIf Optcriterio(3).Value = True Then
        QRY = "SELECT munumero, mucodemp, mucodinf, munomnot, munumpla, mufecha, " _
                  & " muhora, muauxcon, murecpag, mudescue, muvaldes, munetpag, muestado, hvnumint, muturno, muoperar, musaldo,  munomcaj, muauxalt  " _
              & "FROM opmultas, ophojveh " _
             & "WHERE mufecsis BETWEEN TO_DATE('" & Fecha1 & "','MM/DD/YYYY') " _
               & "AND TO_DATE('" & Fecha2 & "','MM/DD/YYYY') " _
               & "AND munumpla = hvplaca " _
               & "AND muestado = '" & estado & "' " _
               & "AND munomcaj = '" & Txtvalbus.text & "' " _
          & "ORDER BY munumero ASC "
     End If
     
     comulcan.Enabled = False
     ExportarExcel (QRY)
     comulcan.Enabled = True
     
     Call Aplicar_Perfil(Me, "comulcan")
     
salir:
  Exit Sub
  
Err_Btnbuscar_Click:
  comulcan.Enabled = True
  MsgBox Err.Description, 16, "Error " & Err.Number
  Resume salir
  Resume
End Sub

Private Sub ExportarExcel(ByVal QRY As String)
    On Error GoTo Errores
    
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "CONSECUTIVO"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "EMPRESA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "COD. INFRAC."
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "NOMBRE NOT."
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "PLACA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "NUMERO INT."
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "FECHA"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "HORA"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "AUX. CONTROL"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "REC. PAGO"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "DESCUENTO"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "VAL. DESCUENTO"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "NETO PAGAR"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "TURNO"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "CAJERO"
    ApExcel.Cells(fila, 15).Font.Size = 10
    ApExcel.Cells(fila, 16).formula = "ESTADO"
    ApExcel.Cells(fila, 16).Font.Size = 10
    ApExcel.Cells(fila, 17).formula = "SALDO"
    ApExcel.Cells(fila, 17).Font.Size = 10
    ApExcel.Cells(fila, 18).formula = "CAJA"
    ApExcel.Cells(fila, 18).Font.Size = 10
    fila = fila + 1
    
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        While Not Tbl1.EOF
            
            'Consultamos los datos de la empresa
            '*********************************************
            QRY = "SELECT henombre "
            If Len(Tbl1!mucodemp) = 4 Then
                QRY = QRY & "FROM ophojempcon "
            Else
                QRY = QRY & "FROM ophojemp "
            End If
            QRY = QRY & "WHERE hecodigo = '" & Tbl1!mucodemp & "' "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl1.EOF Then
                sEmpresa = Tbl2!henombre
            End If
            '*********************************************
            
            'Saldos
            '*********************************************
            If IsNull(Tbl1!musaldo) Then
                saldo = "NO"
            Else
                saldo = Tbl1!musaldo
            End If
            '*********************************************
            
            'Nombre infracción
            '*********************************************
            QRY = "SELECT psdes FROM geparsis WHERE pscod = '" & Tbl1!mucodinf & "' "
            
            Set Tbl2 = Dbs.Execute(QRY)
            
            If Not Tbl2.EOF Then
                sDesInfr = Tbl2!psdes
            Else
                sDesInfr = ""
            End If
            '*********************************************
            
            'Llenamos los datos en Excel
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!munumero & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & sEmpresa & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & sDesInfr & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!munomnot & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!munumpla & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!hvnumint & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!mufecha & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!muhora & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!muauxcon & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!murecpag & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & Tbl1!mudescue & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!muvaldes & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & Tbl1!munetpag & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & Tbl1!muturno & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & Tbl1!muoperar & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            ApExcel.Cells(fila, 16).formula = "" & Tbl1!muestado & ""
            ApExcel.Cells(fila, 16).Font.Size = 10
            ApExcel.Cells(fila, 17).formula = "" & Tbl1!musaldo & ""
            ApExcel.Cells(fila, 17).Font.Size = 10
            ApExcel.Cells(fila, 18).formula = "" & Tbl1!munomcaj & ""
            ApExcel.Cells(fila, 18).Font.Size = 10
            
            total = total + Val(Tbl1!munetpag)
            fila = fila + 1
            Tbl1.MoveNext
        Wend
        
        Tbl2.Close
        Tbl1.Close
        
        Msktotal.text = total
        
        ApExcel.Worksheets(1).Name = "CONSULTA MULTAS"
        ApExcel.Visible = True
    Else
        MsgBox "No se encontraron registros para los datos digitados", 48, "Advertencia"
    End If
    
    Exit Sub
    
Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Btnimprimir_Click()
On Error GoTo err_mensaje
  
  If Grdconmul.text <> "" And Grdconmul.rows > 1 Then
     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
     lineas = Grdconmul.rows - 1
     Grdconmul.Row = 1
            
     
     Dbs.BeginTrans
     On Error GoTo Err_Trans
     Do
        ReDim dato(1, 18)
                
        For i = 0 To 17
            Grdconmul.Col = i
            dato(1, i) = Trim(Grdconmul.text)
        Next i
            
        ' insertamos en la tabla
             
        Dbs.Execute "INSERT INTO optempo1 " _
        & "(temnumero1,      temtext8,           temtext2,           temtext3,            temtext4,            temnumero3,          temdate1,           temtext5,             temtext6,            temnumero2,          temnumdob1,          temnumdob3,           temnumdob2,           temtext7,              temtext11,            temtext9, temnumdob4, temtext10,tempc ) VALUES ('" _
        & dato(1, 0) & "' ,'" & dato(1, 1) & "','" & dato(1, 2) & "','" & dato(1, 3) & "','" & dato(1, 4) & "','" & dato(1, 5) & "' ,TO_DATE('" & dato(1, 6) & "','MM/DD/YYYY'),'" & Format(dato(1, 7), "HH:mm") & "','" & dato(1, 8) & "','" & dato(1, 9) & "','" & dato(1, 10) & "','" & dato(1, 11) & "','" & dato(1, 12) & "','" & dato(1, 15) & "','" & dato(1, 13) & "','" & dato(1, 14) & "'," & dato(1, 16) & ",'" & dato(1, 17) & "',  '" & Pc & "' )"
             
        ' fin insertar
              
             
        lineas = lineas - 1
             
        If lineas <> 0 Then
           Grdconmul.Row = Grdconmul.Row + 1
        End If
     Loop Until lineas = 0  'do_while que se mueve x la grilla
     
     Dbs.CommitTrans
     On Error GoTo err_mensaje
     inicio1 = Mskfecini.FormattedText
     final = Mskfecfin.FormattedText
     If Optcriterio(0).Value = True Then
        turcaje = "Turno:"
     ElseIf Optcriterio(1).Value = True Then
        turcaje = "Cajero:"
     ElseIf Optcriterio(2).Value = True Then
        turcaje = ""
      ElseIf Optcriterio(3).Value = True Then
        turcaje = "Caja:"
     End If
     
     reporte = RpPath + "\conmulca.rpt"
     rptterminal.ReportFileName = reporte
     rptterminal.Connect = G_Conexion_reportes
     rptterminal.Formulas(0) = "total = '" & Msktotal.FormattedText & "'"
     rptterminal.Formulas(1) = "inicio1 = '" & inicio1 & "'"
     rptterminal.Formulas(2) = "final = '" & final & "'"
     rptterminal.Formulas(3) = "empresa = '" & G_NomEmpresa & "'"
     rptterminal.Formulas(4) = "turcaj = '" & turcaje & "'"
     rptterminal.Formulas(5) = "valor = '" & Txtvalbus.text & "'"
     rptterminal.WindowState = crptMaximized
     rptterminal.SelectionFormula = "{optempo1.tempc}= '" & Pc & "' "
     Conmultas.Enabled = False
     rptterminal.Destination = 0
     rptterminal.Action = 1
     Conmultas.Enabled = True
     Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
  End If
  
  
  Exit Sub
  
err_mensaje:
   Conmultas.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
Err_Trans:
   Dbs.RollbackTrans
   Conmultas.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnlimpiar_Click()
  
  Mskfecini.Mask = ""
  Mskfecini.text = ""
  Mskfecini.Mask = "##/##/####"
  Mskfecfin.Mask = ""
  Mskfecfin.text = ""
  Mskfecfin.Mask = "##/##/####"
  Grdconmul.rows = 1
  Grdconmul.Cols = 2
  Grdconmul.Clear
  Msktotal.text = ""
  Txtvalbus.text = ""
  Lblvalbus.Caption = ""
  Optcriterio(0).Value = True
  DIB_GRILLA

End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Btntotalizar_Click()
  If Grdconmul.text <> "" And Grdconmul.rows > 1 Then
     dato1 = 0
     Grdconmul.Col = 12
     For i = 1 To Grdconmul.rows - 1
         Grdconmul.Row = i
         dato1 = dato1 + CDbl(Grdconmul.text)
     Next i
     Msktotal.text = dato1
  End If

End Sub

Private Sub Form_Activate()
Call Aplicar_Perfil(Me, "comulcan")
End Sub

Private Sub Form_Load()
   DIB_GRILLA
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
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

Private Sub optcriterio_Click(index As Integer)
   Txtvalbus.text = ""
   Lblvalbus.Caption = ""
   Txtvalbus.Enabled = True
   If Optcriterio(0).Value = True Then
      Txtvalbus.MaxLength = 1
   ElseIf Optcriterio(1).Value = True Then
      Txtvalbus.MaxLength = 15
   ElseIf Optcriterio(2).Value = True Then
      Txtvalbus.Enabled = False
    ElseIf Optcriterio(3).Value = True Then
      Txtvalbus.MaxLength = 10
   End If
End Sub



Private Sub Txtvalbus_DblClick()
  If Optcriterio(3).Value = True Then
      forma = "12"
      forma1 = "5"
      Txtvalbus.text = ""
      buscar.Show 1
      Txtvalbus_LostFocus
   End If
  
End Sub

Private Sub Txtvalbus_KeyPress(KeyAscii As Integer)
   If Optcriterio(0).Value = True Then
      If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Or (KeyAscii% = 47) Then
         Exit Sub
      Else
         KeyAscii% = 0
      End If
   ElseIf Optcriterio(0).Value = True Then
      Char = Chr(KeyAscii%)
      KeyAscii% = Asc(LCase(Char))
      If (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% = 8) Then
         Exit Sub
      Else
         KeyAscii% = 0
      End If
   ElseIf Optcriterio(3).Value = True Then
       Char = Chr(KeyAscii%)
      KeyAscii = Asc(UCase(Char))  ' mayuscula
      
      If KeyAscii% = 39 Or KeyAscii% = 34 Or KeyAscii% = 124 Then
         KeyAscii% = 0
      End If

   End If
End Sub

Private Sub Txtvalbus_LostFocus()
On Error GoTo err_Txtvalbus_LostFocus
   
   If Trim(Txtvalbus.text) <> "" Then
   
      If Optcriterio(3).Value = True Then
         QRY = "SELECT * " _
                & "FROM cocaja " _
               & "WHERE canomcaj ='" & Trim(Txtvalbus.text) & "' "
               
         Set Tbl1 = Dbs.Execute(QRY)
         
         If Not Tbl1.EOF Then
            Lblvalbus.Caption = Tbl1!canomcaj
         End If
         Tbl1.Close
      End If
   End If
   Exit Sub
   
err_Txtvalbus_LostFocus:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub
