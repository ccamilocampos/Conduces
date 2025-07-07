VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form corecasaldo 
   Caption         =   "Recalcular Sados Cupos por Empresa"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
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
      Left            =   4800
      Picture         =   "corecalsaldo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCodEmp 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
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
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label3 
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
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
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
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   2520
   End
End
Attribute VB_Name = "corecasaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   
   If txtCodEmp.text = "" Then
    MsgBox "Debe seleccionar una empresa", 48, "Advertencia"
    txtCodEmp.SetFocus
    Exit Sub
   End If
   
   sFechaIni = Mskfecini.text
   sFechaFin = Mskfecfin.text
   
   'Debido a un error a la hora de la venta de tasas de uso por cupos
   'Nos aseguramos que las tasas vendidas correspondan a la empresa
   
   ProBar.Value = ProBar.Min
   
   DoEvents
   
   QRY = "UPDATE opmovimientos_prueba " _
            & "SET mocodemp = (SELECT coconduc.cocodemp " _
                                & "FROM coconduc " _
                                & "WHERE coconduc.conumero = opmovimientos_prueba.motasauso " _
                                & "AND cocodemp = '" & txtCodEmp.text & "' " _
                                & "AND cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                                & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY')) " _
            & "WHERE EXISTS (SELECT coconduc.cocodemp " _
                                & "FROM coconduc " _
                                & "WHERE coconduc.conumero = opmovimientos_prueba.motasauso " _
                                & "AND cocodemp = '" & txtCodEmp.text & "' " _
                                & "AND cofecsal BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
                                & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY')) " _
            & "AND mofecha BETWEEN TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') " _
            & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY') "
            
    Dbs.Execute (QRY)
    
    'Una vez arreglados los registros, disponemos de recalcular los saldos
    'Para ello, buscamos la última consignación dentro del rango de fechas digitado
    
    QRY = "SELECT * FROM (SELECT monumero, mocodemp, mofecha, motipmov, movalor + mosaldo AS saldo, motipo " _
                            & "FROM opmovimientos_prueba " _
                            & "WHERE mocodemp = '" & txtCodEmp.text & "' " _
                            & "AND mofecha BETWEEN TO_DATE((SELECT TO_CHAR((SELECT TO_DATE('" & Mskfecini.text & "','MM/DD/YYYY') - 30 FROM DUAL ), 'MM/DD/YYYY') FROM DUAL), 'MM/DD/YYYY') " _
                            & "AND TO_DATE('" & Mskfecfin.text & "','MM/DD/YYYY') " _
                            & "AND motipmov = 'DB' " _
                            & "AND motipo = 'FACT' " _
                            & "ORDER BY monumero ASC " _
            & ") WHERE rownum <= 1 "
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        sSaldo = CDbl(Tbl1!saldo)
        sFechaConsignacion = Tbl1!mofecha
        sConsecutivo = Tbl1!monumero
        
    End If
    
    'Teniendo la primera consignación, consultamos todos los movimientos que sean mayor a esta consignación
    QRY = "SELECT monumero, mocodemp, mofecha, motipmov, motipo, movalor " _
            & "FROM opmovimientos_prueba " _
            & "WHERE mocodemp = '" & txtCodEmp.text & "' " _
            & "AND mofecha >= TO_DATE('" & sFechaConsignacion & "', 'MM/DD/YYYY') " _
            & "AND monumero > " & sConsecutivo & " " _
            & "ORDER BY monumero ASC "
            
    Set Tbl2 = Dbs.Execute(QRY)
    
    'Recorremos todos los movimientos realizando los respectivos calculos, para así recalcular el saldo final
    If Not Tbl2.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl2.RecordCount + 2
      Tbl2.MoveFirst
    End If
    
    If Not Tbl2.EOF Then
        Do While Not Tbl2.EOF
            
            If Tbl2!motipmov = "CR" Then
                QRY = "UPDATE opmovimientos_prueba " _
                        & "SET mosaldo = " & sSaldo & " " _
                        & "WHERE monumero = " & Tbl2!monumero & " "
                        
                Dbs.Execute (QRY)
                
                sSaldo = sSaldo - CDbl(Tbl2!movalor)
                
            ElseIf Tbl2!motipmov = "DB" Then
                QRY = "UPDATE opmovimientos_prueba " _
                        & "SET mosaldo = " & sSaldo & " " _
                        & "WHERE monumero = " & Tbl2!monumero & " "
            
                sSaldo = sSaldo + CDbl(Tbl2!movalor)
            End If
            
            ProBar.Value = ProBar.Value + 1
            Tbl2.MoveNext
            
        Loop
    
    End If
    
    'Por último, actualizamos el saldo de la empresa en la tabla de cupos
    QRY = "UPDATE cocupos_prueba " _
            & "SET cpsaldo = " & sSaldo & " " _
            & "WHERE cpempresa = '" & txtCodEmp.text & "' " _
            & "AND cpestado = 'A' "
            
    Dbs.Execute (QRY)
    
    ProBar.Value = ProBar.Max
    MsgBox "Proceso terminado existosamente", 48, "Advertencia"
    
    Exit Sub

Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    
    
End Sub

Private Sub Form_Load()
    QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'HORASIS' "
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    If Not Tbl1.EOF Then
        Mskfecfin.text = Mid(Tbl1!psval, 1, 10)
    End If
    
    ProBar.Value = ProBar.Max
    
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

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub



Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "72"
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
