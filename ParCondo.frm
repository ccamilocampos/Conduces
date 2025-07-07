VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ParCondo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Condonar Deuda Parqueadero"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   1815
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   6135
      Begin VB.CommandButton BtnAsignar 
         Caption         =   "&Asignar"
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
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ingresa el descuento en horas"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TxtValor 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDescuentos 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         Enabled         =   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Caption         =   "Valor:"
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
         Left            =   5040
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Observaciones"
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
      TabIndex        =   24
      Top             =   6960
      Width           =   6135
      Begin VB.TextBox TxtObserv 
         Height          =   1125
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Observaciones para condonar la deuda del parqueadero"
         Top             =   240
         Width           =   5895
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
      TabIndex        =   19
      Top             =   120
      Width           =   6135
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Número interno del vehículo"
         Top             =   720
         Width           =   735
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
         TabIndex        =   0
         ToolTipText     =   "Placa del vehículo"
         Top             =   720
         Width           =   975
      End
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código de la empresa de Transporte (presione DblClick para ayuda)"
         Top             =   720
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
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   735
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
         TabIndex        =   22
         Top             =   360
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
         Left            =   3000
         TabIndex        =   21
         Top             =   720
         Width           =   3015
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
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid GrdParque 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
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
      Left            =   6360
      Picture         =   "ParCondo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ingresa el registro a la base de datos"
      Top             =   4800
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
      Left            =   6360
      Picture         =   "ParCondo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir recibo de conduce"
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
      Left            =   6360
      Picture         =   "ParCondo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   6720
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
      Left            =   6360
      Picture         =   "ParCondo.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir de la forma actual"
      Top             =   7680
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   6135
      Begin MSMask.MaskEdBox MskValor 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDescuento 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskValPagar 
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.0;(#,##0.0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Valor a Pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Descuento Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Valor Total Deuda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ParCondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Fila_Descuentos As Double
Dim L_Valor_Parqueadero
Dim L_Valor_Tarifa As Double

Private Sub DIB_GRILLA()
   GrdParque.rows = 1
   GrdParque.Cols = 10
   GrdParque.Clear
   
   s$ = "|Consecutivo|Placa|Fecha Ingreso|Fecha Salida|Permanencia|Valor|Observaciones|vrhora|Tipo|Caja"
   GrdParque.FormatString = s$
   GrdParque.Row = 0
   GrdParque.col = 0
   GrdParque.ColWidth(0) = 500
   GrdParque.ColWidth(1) = 1200
   GrdParque.ColWidth(2) = 1000
   GrdParque.ColWidth(3) = 1500
   GrdParque.ColWidth(4) = 1500
   GrdParque.ColWidth(5) = 1500
   GrdParque.ColWidth(6) = 1000
   GrdParque.ColWidth(7) = 2000
   GrdParque.ColWidth(8) = 0
   GrdParque.ColWidth(9) = 1500
   GrdParque.ColWidth(10) = 0  'para evitar confusiones en la información, dejamos esta columna "invisible"
   For i = 0 To GrdParque.Cols - 1
       GrdParque.col = i
       GrdParque.CellFontBold = 1
   Next i

End Sub

Private Sub Dib_Grilla_Descuentos()
   GrdDescuentos.rows = 1
   GrdDescuentos.Cols = 4
   GrdDescuentos.Clear
   
   s$ = "Código|Descuento|Valor|Fecha"
   GrdDescuentos.FormatString = s$
   GrdDescuentos.Row = 0
   GrdDescuentos.col = 0
   GrdDescuentos.ColWidth(0) = 0
   GrdDescuentos.ColWidth(1) = 3200
   GrdDescuentos.ColWidth(2) = 1500
   GrdDescuentos.ColWidth(3) = 0
   For i = 0 To GrdDescuentos.Cols - 1
       GrdDescuentos.col = i
       GrdDescuentos.CellFontBold = 1
   Next i
   
End Sub

Private Sub Traer_Datos_Descuentos(ByVal fNumero As Double, ByVal fValor As Double)
   On Error GoTo Errores
   Dib_Grilla_Descuentos
   
'   QRY = "SELECT dpnumero, dpcoddes, dpvalor, dphora " _
'         & "FROM opdesparque " _
'        & "WHERE dpnumero = " & fNumero & ""

    QRY = "SELECT dpnumero, dpconpar, dpcoddes, dpvalor, dpfecmod " _
            & "FROM opdesparque2 " _
            & "WHERE dpconpar = " & fNumero & " "
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Tbl1.EOF Then
      Btngrabar.Enabled = True
      fqry01 = "SELECT dscoddet, dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscodtip = stcodtip " _
                & "AND stdes = 'TIPOS DESCUENTOS PARQUEADEROS' " _
                & "AND dsest = 'A'"
              
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         Do While Not Ftb01.EOF
            
            entry = Ftb01!dscoddet + Chr(9) + Ftb01!dsdes
            
            GrdDescuentos.AddItem entry
            Ftb01.MoveNext
         Loop
         Ftb01.Close
           
      Else
         Ftb01.Close
         MsgBox "Catalogo Tipos de Descuento de Parqueaderos no Existe", 48, "Advertencia"
      End If
      
      MskValor.text = fValor
   Else
      Btngrabar.Caption = "Actualizar"
      mskDescuento.text = 0
      fqry01 = "SELECT dscoddet, dsdes " _
               & "FROM gesuptip, gedetsuptip " _
              & "WHERE dscodtip = stcodtip " _
                & "AND stdes = 'TIPOS DESCUENTOS PARQUEADEROS' " _
                & "AND dsest = 'A'"
              
      Set Ftb01 = Dbs.Execute(fqry01)
        
      If Not Ftb01.EOF Then
         Do While Not Ftb01.EOF
            
            entry = Ftb01!dscoddet + Chr(9) + Ftb01!dsdes
            
            GrdDescuentos.AddItem entry
            Ftb01.MoveNext
         Loop
         Ftb01.Close
           
      Else
         Ftb01.Close
         MsgBox "Catalogo Tipos de Descuento de Parqueaderos no Existe", 48, "Advertencia"
      End If
      
      'Ahora asignamos a cada uno su valor si es que tiene
      For i = 1 To GrdDescuentos.rows - 1
        Do While Not Tbl1.EOF
            If Tbl1!dpcoddes = GrdDescuentos.TextMatrix(i, 0) Then
                GrdDescuentos.TextMatrix(i, 2) = CDbl(Tbl1!dpvalor)
                mskDescuento.text = CDbl(mskDescuento.text) + CDbl(Tbl1!dpvalor)
            End If
            Tbl1.MoveNext
        Loop
        Tbl1.MoveFirst
      Next
      
      'Llenamos la grilla con los conceptos
      
'      Do While Not Tbl1.EOF
'         sDescripcion = Traer_Descripcion(Tbl1!dpcoddes, "TIPOS DESCUENTOS PARQUEADEROS")
'         mskDescuento.text = CDbl(mskDescuento.text) + CDbl(Tbl1!dpvalor)
'         entry = Tbl1!dpcoddes + Chr(9) + sDescripcion + Chr(9) + CStr(Tbl1!dpvalor) + Chr(9) + CStr(Tbl1!dpfecmod)
'
'         GrdDescuentos.AddItem entry
'         Tbl1.MoveNext
'      Loop
      Tbl1.Close
      
      'Una vez traido los datos, sumamos para conocer el valor original de la deuda
      sValorOrg = 0
      For i = 1 To GrdDescuentos.rows - 1
        If GrdDescuentos.TextMatrix(i, 2) <> "" Then
            sValorOrg = sValorOrg + CDbl(GrdDescuentos.TextMatrix(i, 2))
        End If
      Next
   End If
   
   MskValor.text = sValorOrg + fValor
   
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
    Resume
End Sub

Private Sub Traer_Datos()
   Dim sDatos As String
   On Error GoTo Errores
   
   'Sebastián Rondón - 08/23/2023
   'Consultamos la fecha parametrizada desde donde se consultaran los cobros
   QRY = "SELECT psval FROM geparsis WHERE pscod = 'FECPARPRO' "
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
    fecpar = Tbl1!psval
   End If
   
   DIB_GRILLA
   scero = 0
   fqry01 = "SELECT panumero, paplaca, pacaja, pafecini, " _
                 & "pafecsal, pavalor, paaviso, " _
                 & "paestado, paobservacion, patiempo, patarifa, patipo " _
            & "FROM coparqueop " _
           & "WHERE paplaca = '" & Txtnumpla.text & "' AND pavalor <> " & scero & " " _
             & "AND pafecini >= TO_DATE('" & fecpar & "','MM/DD/YYYY') " _
             & "AND paestado = 'A' ORDER BY  paaviso DESC, panumero "

   Set Ftb01 = Dbs.Execute(fqry01)
   sTotal = 0
   Do While Not Ftb01.EOF
      Txtnumpla.Enabled = False
      sValor = Ftb01!pavalor
      sObservacion = ""
      If Not IsNull(Ftb01!paobservacion) Then
         sObservacion = Ftb01!paobservacion
      End If
      
      permanencia1 = CDate(Ftb01!pafecsal) - CDate(Ftb01!pafecini)
      
      If permanencia1 > 1 Then
        If permanencia1 >= 86400 Then
           horest = permanencia1 / 86400
    
           perma1 = Fix(horest)
           perma = horest - perma1
    
           permanencia = perma1 & " " & Format(perma, "hh:mm") 'para revisar
        Else
           perma1 = Fix(permanencia1)
           perma = permanencia1 - perma1
    
           permanencia = perma1 & " " & Format(perma, "hh:mm") 'para revisar
        End If
      Else
        permanencia = Format(CDate(Ftb01!pafecsal) - CDate(Ftb01!pafecini), "hh:mm")
      End If
      
      If Ftb01!patipo = "P" Then
         sTipo = "Permanencia"
      ElseIf Ftb01!patipo = "V" Then
         sTipo = "Vigencia"
      End If
      
      'Para evitar el posible error de que aparezca un parqueadero de otra caja (parqueadero ya pago) validamos que en caso de traer una caja registrada compare si existe, de no tener dejamos la actual
      
      If IsNull(Ftb01!pacaja) Then
        sCaja = Pc
      End If
      
      entry = "" + Chr(9) + CStr(Ftb01!panumero) + Chr(9) + Ftb01!paplaca + Chr(9) + Format(Ftb01!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
               Format(Ftb01!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + CStr(permanencia) + Chr(9) + CStr(sValor) + Chr(9) + sObservacion + Chr(9) + CStr(Ftb01!patarifa) + Chr(9) + sTipo + Chr(9) + sCaja
      GrdParque.AddItem entry            ' Agrega la multa.
      
      Ftb01.MoveNext
   Loop
   Ftb01.Close
   mskDescuento.text = "0"
   
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
End Sub



Private Sub BtnAsignar_Click()
   On Error GoTo Errores
   If TxtValor.text <> "" Or TxtValor.text <> "0" Then
            
            
      'sValor = CDbl(MskDescuento.text) + (CDbl(L_Valor_Tarifa) * CDbl(TxtValor.text))
      sValor = CDbl(TxtValor.text)
      'Asignamos el valor en la grilla
      GrdDescuentos.TextMatrix(L_Fila_Descuentos, 2) = CDbl(sValor)
      
      sTotalDes = 0
      'Hacemos la sumatoria total de descuentos
      For i = 1 To GrdDescuentos.rows - 1
        If GrdDescuentos.TextMatrix(i, 2) <> "" Then
            sTotalDes = sTotalDes + CDbl(GrdDescuentos.TextMatrix(i, 2))
        End If
        
      Next
      If sTotalDes <= CDbl(MskValor.text) Then
        mskDescuento.text = CDbl(sTotalDes)
        MskValPagar.text = CDbl(MskValor) - CDbl(sTotalDes)
        TxtValor.text = ""
        TxtValor.Enabled = False
      Else
        MsgBox "No se puede hacer el descuento porque el valor del parqueadero" & Chr(13) & "es menor al valor del descuento", 48, "Advertencia"
        GrdDescuentos.TextMatrix(L_Fila_Descuentos, 2) = 0
        TxtValor.SetFocus
      End If
      
'      sValorTotal = CDbl(MskDescuento.text) + CDbl(Mskvalor.text)
'
'      If CDbl(sValorTotal) >= sValor Then
'         BtnAsignar.Enabled = False
'         GrdDescuentos.TextMatrix(L_Fila_Descuentos, 2) = CDbl(TxtValor.text)
'         MskDescuento.text = CDbl(MskDescuento.text) + (CDbl(L_Valor_Tarifa) * CDbl(TxtValor.text))
'         MskValPagar.text = CDbl(Mskvalor.text) - CDbl(MskDescuento.text)
'         MskDescuento.text = CDbl(TxtValor.text)
'         MskValPagar.text = CDbl(sValorTotal) - CDbl(MskDescuento.text)
'
'         TxtValor.text = ""
'         TxtValor.Enabled = False
'      Else
'         MsgBox "No se puede hacer el descuento porque el valor del parqueadero" & Chr(13) & "es menor al valor del descuento", 48, "Advertencia"
'         TxtValor.SetFocus
'      End If
      
      
      
      
      
   End If
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
      GrdParque.col = 0
      
      If GrdParque.text = "->" Then
         seleccion = "1"
         Exit For
      End If
   Next i
       
   If seleccion = "0" Then
      MsgBox "No se ha seleccionado ningún registro", 48, "Advertencia"
      GrdParque.Row = 1
      GrdParque.col = 0
      GrdParque.SetFocus
      Exit Sub
   End If
   seleccion = "0"
   For i = 1 To GrdDescuentos.rows - 1
      If GrdDescuentos.TextMatrix(i, 2) <> "" And GrdDescuentos.TextMatrix(i, 2) <> "0" Then
         seleccion = "1"
         Exit For
      End If
   Next i
   
   If seleccion = "0" Then
      MsgBox "No se ha seleccionado ningún registro para descuento", 48, "Advertencia"
      GrdDescuentos.Row = 1
      GrdDescuentos.col = 0
      GrdDescuentos.SetFocus
      Exit Sub
   End If
   If Trim(TxtObserv.text) = "" Then
      MsgBox "Falta ingresar las observaciones para condonar la deuda de parqueaderos", 48, "Advertencia"
      TxtObserv.SetFocus
      Exit Sub
   End If
   
   If MskValPagar.text = "" Then
    MsgBox "Debe realizar un cambio para guardar", 48, "Advertencia"
    Exit Sub
   End If
   
   rtamsg = MsgBox("Esta seguro de grabar la los registros?", 36, "Confirmación")
      
   If rtamsg = vbNo Then
      Exit Sub
   End If
   
   sFecSistema = Fn_Parametros("HORASIS", 1)
   sFecSis = Format(sFecSistema, "MM/dd/yyyy")
   Dbs.BeginTrans
   ParCondo.Enabled = False
   On Error GoTo Err_Transac
   
   For i = 1 To GrdParque.rows - 1
      GrdParque.col = 0
      GrdParque.Row = i
      
      'Consultamos la fecha del sistema
      QRY = "SELECT psval FROM geparsis " _
                & "WHERE pscod = 'HORASIS' "
                
     Set Tbl1 = Dbs.Execute(QRY)
     
     If Not Tbl1.EOF Then
        fechaMod = Mid(Tbl1!psval, 1, 10)
        horaMod = Mid(Tbl1!psval, 12, 5)
    Else
        MsgBox "No se pudo consultar la fecha y hora del sistema", 48, "Advertencia"
        Exit Sub
     End If
      
      If GrdParque.text = "->" Then
         dato1 = GrdParque.TextMatrix(i, 1)     'numero tiquete
         
         QRY = "UPDATE coparqueop " _
             & "SET pafecdes = to_date('" & sFecSis & "','MM/dd/yyyy') , " _
                & " pafundes = '" & Login & "', " _
                & " pavalor = " & MskValPagar.text & ", " _
                & " padescuento = " & mskDescuento.text & ", " _
                & " paobservacion = '" & TxtObserv.text & "' " _
           & "WHERE panumero = " & dato1 & ""
      
         Dbs.Execute QRY
         '& " pavalor = pavalor - " & mskDescuento.text & ", " _

         'Eliminamos la información ya existente para registrar la nueva
         QRY = "DELETE FROM opdesparque2 WHERE dpconpar = " & dato1 & " "
         Dbs.Execute (QRY)
         
         For j = 1 To GrdDescuentos.rows - 1
            If GrdDescuentos.TextMatrix(j, 2) <> "" And GrdDescuentos.TextMatrix(j, 2) <> "0" Then
               QRY = "INSERT INTO OPDESPARQUE2(dpconpar, dpcoddes, dpvalor, dpfecmod, dphormod, dpobservacion) VALUES(" _
               & dato1 & ",'" & GrdDescuentos.TextMatrix(j, 0) & "'," & GrdDescuentos.TextMatrix(j, 2) & ", " _
               & "TO_DATE('" & fechaMod & "' ,'MM/DD/YYYY'), TO_DATE('" & horaMod & "','HH24:mi'), '" & TxtObserv.text & "')"
               Dbs.Execute QRY
               
            End If
         Next j
         
         Exit For
      End If
   Next i
   
   Dbs.CommitTrans
   ParCondo.Enabled = True
   MsgBox "Los registros se grabaron exitosamente", 64, "O.K."
   Btnlimpiar_Click
   
   On Error GoTo Errores
   
   Exit Sub
   
Err_Transac:
   Dbs.RollbackTrans
Errores:
   ParCondo.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub


Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   Btngrabar.Caption = "Grabar"
   Txtnumpla.Enabled = True
   Txtnumpla.text = ""
   Txtnumint.text = ""
   Txtcodemp.text = ""
   Lblnomemp.Caption = ""
   MskValor.text = ""
   mskDescuento.text = ""
   MskValPagar.text = ""
   TxtObserv.text = ""
   DIB_GRILLA
   Dib_Grilla_Descuentos
   TxtValor.text = ""
   GrdDescuentos.Enabled = False
   TxtValor.Enabled = False
   BtnAsignar.Enabled = False
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
   Dib_Grilla_Descuentos
   L_Valor_Parqueadero = Fn_Parametros("VALPARQ1", 1)
End Sub




Private Sub GrdDescuentos_Click()
   On Error GoTo Errores
   If TxtValor.Enabled = False Then
      TxtValor.Enabled = True
      BtnAsignar.Enabled = True
      If GrdDescuentos.TextMatrix(GrdDescuentos.Row, 2) <> "" Then
         TxtValor.text = GrdDescuentos.TextMatrix(GrdDescuentos.Row, 2)
      Else
         TxtValor.text = "0"
      End If
      If TxtValor.text = "0" Then
        mskDescuento.text = CDbl(mskDescuento.text) + CDbl(TxtValor.text)
        MskValPagar.text = CDbl(MskValor.text) - CDbl(mskDescuento.text)
      End If
      'MskDescuento.text = CDbl(MskDescuento.text) - (CDbl(L_Valor_Tarifa) * CDbl(TxtValor.text))
      'MskValPagar.text = CDbl(Mskvalor.text) - CDbl(MskDescuento.text)
      
      L_Fila_Descuentos = GrdDescuentos.Row
   Else
      MsgBox "No se le ha asiganado el valor al descuento anterior", 48, "Advertencia"
   End If
   TxtValor.SetFocus
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub GrdParque_DblClick()
   Dib_Grilla_Descuentos
   GrdDescuentos.Enabled = False
   mskDescuento.text = "0"
   L_Valor_Tarifa = 0
   
   'Camilo Campos 07/02/2025 a partir de ahora se valida que pertenezca a la misma caja
   
   If Not GrdParque.TextMatrix(GrdParque.Row, 10) = Pc Then
        MsgBox "El parqueadero no pertenece a la caja actual", 48, "Advertencia"
        Exit Sub
   End If
   
   
   If GrdParque.rows > 1 Then
      If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
        numero = CDbl(GrdParque.TextMatrix(GrdParque.Row, 1))
        valor = CDbl(GrdParque.TextMatrix(GrdParque.Row, 6))
         Traer_Datos_Descuentos numero, valor
         For i = 1 To GrdParque.rows - 1
            GrdParque.TextMatrix(i, 0) = ""
         Next i
         
         GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
         'MskValor.text = GrdParque.TextMatrix(GrdParque.Row, 6)
         TxtObserv.text = GrdParque.TextMatrix(GrdParque.Row, 7)
         L_Valor_Tarifa = GrdParque.TextMatrix(GrdParque.Row, 8)
         GrdDescuentos.Enabled = True
      ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = ""
         GrdParque.col = 6
         'MskValor.text = ""
      End If
      
      GrdParque.col = 0
      If GrdParque.rows - 1 > GrdParque.Row Then
         GrdParque.Row = GrdParque.Row + 1
      Else
         'TxtObserv.SetFocus
      End If
   End If
End Sub

Private Sub GrdParque_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If KeyAscii% = 32 Then
      If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
         GrdParque.col = 6
         If mskafac.text = "" Then
            mskafac.text = GrdParque.text
         Else
            mskafac.text = CDbl(mskafac.text) + CDbl(GrdParque.text)
         End If
         
      ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = ""
         GrdParque.col = 6
         mskafac.text = CDbl(mskafac.text) - CDbl(GrdParque.text)
      End If
      
      GrdParque.col = 0
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
      'If Len(Txtnumpla.Text) = 7 Then
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

Private Sub TxtObserv_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtObserv_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))
   If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
      KeyAscii% = 0
   End If
End Sub

Private Sub TxtValor_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
   If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub
