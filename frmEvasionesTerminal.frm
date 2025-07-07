VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEvasionesTerminal 
   Caption         =   "Consulta de Evasiones Terminales"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5040
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
      Left            =   4080
      Picture         =   "frmEvasionesTerminal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir de la Forma Actual"
      Top             =   960
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
      Left            =   4080
      Picture         =   "frmEvasionesTerminal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   120
      Width           =   855
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   840
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
         Left            =   2160
         TabIndex        =   2
         Top             =   840
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
         Left            =   2160
         TabIndex        =   4
         Top             =   480
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   5160
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmEvasionesTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnExportar_Click()
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
   
   ExportarExcel
   
End Sub
Private Sub ExportarExcel()
    On Error GoTo Errores
    'Exportar a Excel
    Set ApExcel = CreateObject("Excel.application")
    'Añadimos una hoja nueva
    ApExcel.Workbooks.Add
    
    Columna = 1
    fila = 1
    ApExcel.Cells(fila, 1).formula = "EMPRESA"
    ApExcel.Cells(fila, 1).Font.Size = 10
    ApExcel.Cells(fila, 2).formula = "TASA USO"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "FECHA SALIDA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "HORA SALIDA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "PLACA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "CLASE"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "ORIGEN"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "DESTINO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "TERMINAL NORTE"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "TERMINAL SUR"
    ApExcel.Cells(fila, 10).Font.Size = 10
    fila = fila + 1
    
    'Hacemos la consulta a la tabla
    QRY = "SELECT ophojemp.henombre, conumero, to_char(cofecsal,'MM/dd/yyyy') as cofecsal, " _
            & "to_char(cohorsal,'HH24:MI') as cohorsal, coplaca, " _
            & "CASE WHEN coterminal = 'C' THEN 'CENTRAL' END as coterminal, opruta.rudestino, coclase " _
            & "FROM coconduc, opruta, ophojemp " _
            & "WHERE cofecsal >= TO_DATE('" & Mskfecini.text & "','MM/dd/yyyy') " _
            & "AND cofecsal <= TO_DATE('" & Mskfecfin.text & "','MM/dd/yyyy') " _
            & "AND coterminal = 'C' " _
            & "AND cooritra = 'O' " _
            & "AND ophojemp.hecodigo = coconduc.cocodemp " _
            & "AND coconduc.cocodrut = opruta.rucodigo "
            
            
    Set Tbl1 = Dbs.Execute(QRY)
    
    'Si no es vacio inicializamos el Progress Bar
    If Not Tbl1.EOF Then
      Tbl1.MoveLast
      ProBar.Value = ProBar.Min
      ProBar.Max = Tbl1.RecordCount + 2
      Tbl1.MoveFirst
    End If
    
    'Si la consulta no es vacia, creamos el archivo
    If Not Tbl1.EOF Then
    
        While Not Tbl1.EOF
            
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!conumero & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!cofecsal & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!cohorsal & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!coplaca & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!coclase & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!coterminal & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!rudestino & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            
            'Variables para la consulta de la tasa de uso transito
            Fecha1 = Tbl1!cofecsal
            Fecha2 = DateAdd("d", 1, Tbl1!cofecsal)
            Norte = consultarTransito(CStr(Fecha1), CStr(Fecha2), "T", Tbl1!coplaca, "N", Tbl1!cohorsal)
            Sur = consultarTransito(CStr(Fecha1), CStr(Fecha2), "T", Tbl1!coplaca, "S", Tbl1!cohorsal)
            
            'Agregamos lo consultado
            ApExcel.Cells(fila, 9).formula = "" & Norte & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Sur & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Wend
        ApExcel.Worksheets(1).Name = "CONSULTA" & Month(Mskfecini.text) & Day(Mskfecini.text) & Year(Mskfecini.text)
        ApExcel.Visible = True
        ProBar.Value = ProBar.Max
    Else
        MsgBox "No se encontraron registros para las fechas digitadas", 48, "Advertencia"
    End If
    
    Exit Sub
Errores:
    MsgBox Err.Description, vbExclamation, "Advertencia " & Err.Number
    Exit Sub
    Resume

End Sub
Private Function consultarTransito(Fecha1 As String, Fecha2 As String, Cooritra As String, Placa As String, Terminal As String, hora As String) As String
    fqry01 = "SELECT to_char(cohorsal, 'MM/dd/yyyy HH24:MI') as hora FROM coconduc " _
                        & "WHERE cofecsal >= TO_DATE('" & Fecha1 & "', 'MM/DD/YYYY') " _
                        & "AND cofecsal <= TO_DATE('" & Fecha2 & "', 'MM/DD/YYYY') " _
                        & "AND cooritra = '" & Cooritra & "' " _
                        & "AND coplaca = '" & Placa & "' " _
                        & "AND coterminal = '" & Terminal & "' " _
                        & "AND cohorsal > TO_DATE('" & Fecha1 & " " & hora & "', 'MM/DD/YYYY HH24:MI') " _
                        & "ORDER BY cohorsal ASC"
    Set Tbl2 = Dbs.Execute(fqry01)
    
    If Not Tbl2.EOF Then
        transito = Mid(Tbl2!hora, 12)
    Else
        transito = "No"
    End If
    
    consultarTransito = transito
End Function
Private Sub Btnsalir_Click()
    Unload Me
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
