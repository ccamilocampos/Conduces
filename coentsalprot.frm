VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form coentsalprot 
   Caption         =   "Consulta Procesos Entradas y Salidas"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Criterios de Busqueda"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   2895
      Begin VB.TextBox txtPlaca 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
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
      TabIndex        =   17
      Top             =   4800
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
      Begin VB.OptionButton OptTipo 
         Caption         =   "C. Auto."
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
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptTipo 
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
         Index           =   5
         Left            =   1440
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Vigencias"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Entradas"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Salidas"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Perman."
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
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
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
      Picture         =   "coentsalprot.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4800
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
      Picture         =   "coentsalprot.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3840
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
      Picture         =   "coentsalprot.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   2040
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
      Picture         =   "coentsalprot.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton BtnExportar 
      Caption         =   "&Exportar"
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
      Picture         =   "coentsalprot.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Documento Soporte"
      Top             =   3000
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin MSFlexGridLib.MSFlexGrid grdDatos 
      Height          =   4575
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8070
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   20
      Top             =   5520
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   8760
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   6600
   End
End
Attribute VB_Name = "coentsalprot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function DIB_GRILLA()

   grdDatos.rows = 1
   grdDatos.Cols = 7
   grdDatos.Clear
   
   s$ = "Consecutivo |Tipo   |Placa |ID Ciclo |Fecha Ingreso |Fecha Salida |Fecha Registro "
   grdDatos.FormatString = s$
   grdDatos.Row = 0
   grdDatos.Col = 0
   grdDatos.ColWidth(0) = 1100
   grdDatos.ColWidth(1) = 1600
   grdDatos.ColWidth(2) = 1200
   grdDatos.ColWidth(3) = 1400
   grdDatos.ColWidth(4) = 1600
   grdDatos.ColWidth(5) = 1600
   grdDatos.ColWidth(6) = 2000
   
   
   For i = 0 To grdDatos.Cols - 1
       grdDatos.Col = i
       grdDatos.CellFontBold = 1
   Next i
End Function

Private Sub Btnbuscar_Click()
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
   
   'Limpiamos al Grilla
   DIB_GRILLA
   
   'Declaramos las variables
   Fecha = Mskfecini.text
   Fecha1 = Mskfecfin.text
   
   QRY = "SELECT lonumero, lotipo, loplaca, loidciclo, lofecing, " _
            & "lofecsal, lofecreg FROM logcierreciclos " _
            & "WHERE lofecreg BETWEEN TO_DATE('" & Fecha & " 00:00', 'MM/DD/YYYY HH24:mi') " _
            & "AND TO_DATE('" & Fecha1 & " 23:59', 'MM/DD/YYYY HH24:mi') " _
            
   If OptTipo(0).Value = True Then
        QRY = QRY & "AND lotipo = 'ENTRADA PROTECH' "
   ElseIf OptTipo(1).Value = True Then
        QRY = QRY & "AND lotipo = 'SALIDA PROTECH' "
   ElseIf OptTipo(2).Value = True Then
        QRY = QRY & "AND lotipo = 'PARQUEADERO' "
   ElseIf OptTipo(3).Value = True Then
        QRY = QRY & "AND lotipo = 'SALIDA CTR VIGENCIA' "
   ElseIf OptTipo(4).Value = True Then
        QRY = QRY & "AND lotipo = 'CIERRE AUTOMATICO' "
   ElseIf OptTipo(5).Value = True Then
        QRY = QRY & "AND lotipo IN ('ENTRADA PROTECH', 'SALIDA PROTECH', 'PARQUEADERO', 'SALIDA CTR VIGENCIA', 'CIERRE AUTOMATICO') "
   End If
   
   If txtPlaca.text <> "" Then
        QRY = QRY & "AND loplaca = '" & txtPlaca.text & "' "
   End If
   
   QRY = QRY & "ORDER BY lonumero "
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
        'Si no es vacio inicializamos el Progress Bar
        TotalRegistros = 0
        Tbl1.MoveLast
        ProBar.Value = ProBar.Min
        ProBar.Max = Tbl1.RecordCount + 2
        Tbl1.MoveFirst
        
        If IsNull(Tbl1!lofecing) Then
            FechaInicio = "NO APLICA"
        Else
            FechaInicio = Tbl1!lofecing
        End If
        If IsNull(Tbl1!lofecsal) Then
            FechaSalida = "NO APLICA"
        Else
            FechaSalida = Tbl1!lofecsal
        End If
        If IsNull(Tbl1!loidciclo) Then
            IdCiclo = "NO APLICA"
        Else
            IdCiclo = Tbl1!loidciclo
        End If
        'Llenamos la Grilla
        Do While Not Tbl1.EOF
            entry = CStr(Tbl1!lonumero) + Chr(9) + Tbl1!lotipo + Chr(9) + Tbl1!loplaca + Chr(9) + CStr(IdCiclo) + Chr(9) + _
                    CStr(FechaInicio) + Chr(9) + CStr(FechaSalida) + Chr(9) + CStr(Tbl1!lofecreg)
            
            grdDatos.AddItem (entry)
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
End Sub

Private Sub BtnExportar_Click()
    'On Error GoTo err_mensaje
    
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
   
   'Limpiamos al Grilla
   DIB_GRILLA
   
   'Declaramos las variables
   Fecha = Mskfecini.text
   Fecha1 = Mskfecfin.text
   
   QRY = "SELECT lonumero, lotipo, loplaca, loidciclo, lofecing, " _
            & "lofecsal, lofecreg FROM logcierreciclos " _
            & "WHERE lofecreg BETWEEN TO_DATE('" & Fecha & " 00:00', 'MM/DD/YYYY HH24:mi') " _
            & "AND TO_DATE('" & Fecha1 & " 23:59', 'MM/DD/YYYY HH24:mi') " _
            
   If OptTipo(0).Value = True Then
        QRY = QRY & "AND lotipo = 'ENTRADA PROTECH' "
   ElseIf OptTipo(1).Value = True Then
        QRY = QRY & "AND lotipo = 'SALIDA PROTECH' "
   ElseIf OptTipo(2).Value = True Then
        QRY = QRY & "AND lotipo = 'PARQUEADERO' "
   ElseIf OptTipo(3).Value = True Then
        QRY = QRY & "AND lotipo = 'SALIDA CTR VIGENCIA' "
   ElseIf OptTipo(4).Value = True Then
        QRY = QRY & "AND lotipo = 'CIERRE AUTOMATICO' "
   ElseIf OptTipo(5).Value = True Then
        QRY = QRY & "AND lotipo IN ('ENTRADA PROTECH', 'SALIDA PROTECH', 'PARQUEADERO', 'SALIDA CTR VIGENCIA', 'CIERRE AUTOMATICO') "
   End If
   
   If txtPlaca.text <> "" Then
        QRY = QRY & "AND loplaca = '" & txtPlaca.text & "' "
   End If
   
   QRY = QRY & "ORDER BY lonumero "
   
   
   
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
    ApExcel.Cells(fila, 2).formula = "TASA"
    ApExcel.Cells(fila, 2).Font.Size = 10
    ApExcel.Cells(fila, 3).formula = "PLACA"
    ApExcel.Cells(fila, 3).Font.Size = 10
    ApExcel.Cells(fila, 4).formula = "CEDULA"
    ApExcel.Cells(fila, 4).Font.Size = 10
    ApExcel.Cells(fila, 5).formula = "COD. EMPRESA"
    ApExcel.Cells(fila, 5).Font.Size = 10
    ApExcel.Cells(fila, 6).formula = "NOM. EMPRESA"
    ApExcel.Cells(fila, 6).Font.Size = 10
    ApExcel.Cells(fila, 7).formula = "RESULTADO"
    ApExcel.Cells(fila, 7).Font.Size = 10
    ApExcel.Cells(fila, 8).formula = "NUMERO"
    ApExcel.Cells(fila, 8).Font.Size = 10
    ApExcel.Cells(fila, 9).formula = "ESTADO"
    ApExcel.Cells(fila, 9).Font.Size = 10
    ApExcel.Cells(fila, 10).formula = "FEC. REGISTRO"
    ApExcel.Cells(fila, 10).Font.Size = 10
    ApExcel.Cells(fila, 11).formula = "CAJA"
    ApExcel.Cells(fila, 11).Font.Size = 10
    ApExcel.Cells(fila, 12).formula = "TERMINAL"
    ApExcel.Cells(fila, 12).Font.Size = 10
    ApExcel.Cells(fila, 13).formula = "COD. RUTA"
    ApExcel.Cells(fila, 13).Font.Size = 10
    ApExcel.Cells(fila, 14).formula = "DESTINO"
    ApExcel.Cells(fila, 14).Font.Size = 10
    ApExcel.Cells(fila, 15).formula = "TIPO RUTA"
    ApExcel.Cells(fila, 15).Font.Size = 10
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
            ApExcel.Cells(fila, 1).formula = "" & Tbl1!apidprueba & ""
            ApExcel.Cells(fila, 1).Font.Size = 10
            ApExcel.Cells(fila, 2).formula = "" & Tbl1!aptasauso & ""
            ApExcel.Cells(fila, 2).Font.Size = 10
            ApExcel.Cells(fila, 3).formula = "" & Tbl1!applaca & ""
            ApExcel.Cells(fila, 3).Font.Size = 10
            ApExcel.Cells(fila, 4).formula = "" & Tbl1!apcedcon & ""
            ApExcel.Cells(fila, 4).Font.Size = 10
            ApExcel.Cells(fila, 5).formula = "" & Tbl1!hecodigo & ""
            ApExcel.Cells(fila, 5).Font.Size = 10
            ApExcel.Cells(fila, 6).formula = "" & Tbl1!henombre & ""
            ApExcel.Cells(fila, 6).Font.Size = 10
            ApExcel.Cells(fila, 7).formula = "" & Tbl1!apresult & ""
            ApExcel.Cells(fila, 7).Font.Size = 10
            ApExcel.Cells(fila, 8).formula = "" & Tbl1!apenumer & ""
            ApExcel.Cells(fila, 8).Font.Size = 10
            ApExcel.Cells(fila, 9).formula = "" & Tbl1!apestado & ""
            ApExcel.Cells(fila, 9).Font.Size = 10
            ApExcel.Cells(fila, 10).formula = "" & Tbl1!apfecreg & ""
            ApExcel.Cells(fila, 10).Font.Size = 10
            ApExcel.Cells(fila, 11).formula = "" & Tbl1!apnumcaj & ""
            ApExcel.Cells(fila, 11).Font.Size = 10
            ApExcel.Cells(fila, 12).formula = "" & Tbl1!coterminal & ""
            ApExcel.Cells(fila, 12).Font.Size = 10
            ApExcel.Cells(fila, 13).formula = "" & Tbl1!cocodrut & ""
            ApExcel.Cells(fila, 13).Font.Size = 10
            ApExcel.Cells(fila, 14).formula = "" & Tbl1!rudestino & ""
            ApExcel.Cells(fila, 14).Font.Size = 10
            ApExcel.Cells(fila, 15).formula = "" & Tbl1!rutipo & ""
            ApExcel.Cells(fila, 15).Font.Size = 10
            
            fila = fila + 1
            ProBar.Value = ProBar.Value + 1
            Tbl1.MoveNext
        Wend
        
        ApExcel.Worksheets(1).Name = "CONSULTA" & Month(Mskfecini.text) & Day(Mskfecini.text) & Year(Mskfecini.text)
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
    Mskfecini.Mask = ""
    Mskfecini.text = ""
    Mskfecini.Mask = "##/##/####"
    Mskfecfin.Mask = ""
    Mskfecfin.text = ""
    Mskfecfin.Mask = "##/##/####"
    OptTipo(5).Value = True
    txtPlaca.text = ""
    DIB_GRILLA
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    OptTipo(5).Value = True
    DIB_GRILLA
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
        
'        If txtCodEmp.text <> "" Then
'            QRY = "SELECT hvplaca " _
'                    & "FROM ophojveh " _
'                    & "WHERE hvplaca = '" & txtPlaca.text & "' " _
'                    & "AND hvcodemp = '" & txtCodEmp.text & "' "
'
'            Set Tbl1 = Dbs.Execute(QRY)
'
'            If Tbl1.EOF Then
'                MsgBox "Vehículo no pertenece a la empresa ", 48, "Advertencia"
'                txtPlaca.text = ""
'                'lblPlaca.Caption = ""
'                Exit Sub
'            End If
'            Tbl1.Close
'        End If
            
        QRY = "SELECT hvplaca, hvpropie " _
                & "FROM ophojveh " _
                & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                & "AND hvestado = 'Activo' "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            'lblPlaca.Caption = "Propietario: " & Tbl1!hvpropie
        Else
            txtPlaca.text = ""
            'lblPlaca.Caption = ""
            MsgBox "El vehículo no existe o se encuentra inactivo ", 48, "Advertencia"
        End If
    End If
End Sub

Private Sub TxtPlaca_DblClick()
    forma = "2"
    forma1 = "26"
    'lblPlaca.Caption = ""
    buscar.Show 1
    TxtPlaca_LostFocus
End Sub

