VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form capparqu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro Tiempo Permanencia"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkReImprimir 
      Caption         =   "Re Imprimir Ultimo Pago"
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
      Left            =   3720
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   6840
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   6135
      Begin MSMask.MaskEdBox Msktotliq 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   240
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
      Begin MSMask.MaskEdBox mskafac 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Top             =   240
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
      Begin VB.Label Label9 
         Caption         =   "Total :"
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
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Total A Facturar:"
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
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnconduces 
      Caption         =   "&Tasas Uso"
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
      Picture         =   "capparqu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Venta de tasas de uso"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Btnsalir 
      Caption         =   "&Salir"
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
      Picture         =   "capparqu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4920
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
      Picture         =   "capparqu.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpia campos de la forma actual"
      Top             =   4080
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
      Picture         =   "capparqu.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir recibo de conduce"
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
      Left            =   6360
      Picture         =   "capparqu.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ingresa el registro a la base de datos"
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid GrdParque 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4895
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
      Top             =   480
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "capparqu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Bandera_Mensaje As Boolean

Private Sub DIB_GRILLA()
   GrdParque.rows = 1
   GrdParque.Cols = 8
   GrdParque.Clear
   
   s$ = "|Consecutivo|Placa|Fecha Ingreso|Fecha Salida|Valor|Descuento|Tipo"
   GrdParque.FormatString = s$
   GrdParque.Row = 0
   GrdParque.col = 0
   GrdParque.ColWidth(0) = 500
   GrdParque.ColWidth(1) = 1200
   GrdParque.ColWidth(2) = 1000
   GrdParque.ColWidth(3) = 1500
   GrdParque.ColWidth(4) = 1500
   GrdParque.ColWidth(5) = 1000
   GrdParque.ColWidth(6) = 0
   GrdParque.ColWidth(7) = 1000
                    
   For i = 0 To GrdParque.Cols - 1
       GrdParque.col = i
       GrdParque.CellFontBold = 1
   Next i

End Sub

Private Sub Traer_Datos()
   Dim sDatos As String
   On Error GoTo Errores
   
   DIB_GRILLA
   
   'Sebastián Rondón 20 Enero - 2023
   'Debido a que Bogotá empezará a utilizar ahora la tabla de coparqueo de Protech
   'Se consulta primero el parámetro que definirá la fecha desde donde se van a empezar a tomar
   'Las vigencias y tiempos de permanencias
   QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'FECPARPRO' "
    Set Tbl1 = Dbs.Execute(QRY)

    If Not Tbl1.EOF Then
        FechaParametro = Tbl1!psval
    End If
   
   fqry01 = "SELECT panumero, paplaca, pafecini, " _
                 & "pafecsal, pavalor, paaviso, " _
                 & "paestado, NVL(padescuento, 0) AS padescuento, patipo " _
            & "FROM coparqueop " _
           & "WHERE paplaca = '" & Txtnumpla.text & "' " _
             & "AND pafecini >= TO_DATE('" & FechaParametro & "','MM/DD/YYYY') " _
             & "AND paestado = 'A' AND pavalor <> 0 ORDER BY  paaviso DESC, panumero "

   Set Ftb01 = Dbs.Execute(fqry01)
   sTotal = 0
   Do While Not Ftb01.EOF
      Txtnumpla.Enabled = False
      entry = "" + Chr(9) + CStr(Ftb01!panumero) + Chr(9) + Ftb01!paplaca + Chr(9) + Format(Ftb01!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
               Format(Ftb01!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + CStr(Ftb01!pavalor) + Chr(9) + CStr(Ftb01!padescuento) + Chr(9) + Ftb01!patipo
      GrdParque.AddItem entry            ' Agrega el registro.
      sTotal = sTotal + Ftb01!pavalor
      Ftb01.MoveNext
   Loop
   Ftb01.Close
   Msktotliq.text = sTotal
   
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
End Sub

Private Sub ReImprimirPago(ByVal sTiempoPermanencia As String, ByVal sNombreTerminal As String, ByVal sNit As String, ByVal sFecha As String, _
                           ByVal sHora As String, ByVal sNumeroInterno As String, ByVal sPlaca As String, ByVal sEmpresa As String, _
                           ByVal sFechaIngreso As String, ByVal sFechaSalida As String, ByVal sValorPagar As String, ByVal sDescuento As String, _
                           ByVal sValor As String, ByVal sCajero As String, ByVal sCaja As String, ByVal sturno As String, _
                           ByVal sNumero As String, ByVal sNumeroBarras As String, ByVal snombrefechaingreso As String, ByVal sTitulo As String, _
                           ByVal sTiempoVigencia As String)
   On Error GoTo Errores
   
   
   On Error Resume Next
   Close #1
   On Error GoTo Errores
   
   If G_ImpCon = "Z" Then
      
      Open "LPT1" For Output As #1

      Print #1, "^XA~TA000~JSN^LT0^MMT^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD0^JUS^LRN^CI0^XZ"
      Print #1, "~DG000.GRF,02048,016,"
      Print #1, ",:::::::::::::::::::::::::T010H010,,U07FHF0,T02F0H0FC0,S01F10H0171,S0280K08,R0140L03,R0680M080,Q01D0N070,Q0180N028,Q060O01C,Q0E0P0E,P0180P03,P020Q0180,P040R040,P0C0R0A0,N0H180R030,O020S018,O060T08,O040T0C,O0C0T04,O080T02,N0180T03,N030U01,N030U0180,N020V080,N060V040,N080EE80M0EA0H080,N0C0FFC0L01FF0H040,N080FB80L01BA0H020,M0180FFC0L01FF0H020,N080EE80M0HEI020,M0100FFC010H0101FF01030,P0HB80L01FE0H020,M03007FC0L01FF0H030,M0200EAERE0,M0300FTF010,M02002BFBFBFBFBFBFBE010,M03007FSF010,M02002AAEQEH08,M07003FSF018,M02001ABBFOFB018,M060H0TF018,M020H0HAREH08,M06013FSF018,M02003FFBFBFBFBFBFBB018,M07007FSF018,M0200ESEA0,L01300FTF010,M0200FFBFIFBFKFA010,M0301FFC7FPF010,M0200EE80M0EA0H020,M0101FF80L01FF0H020,M0100FB80L01FA0H020,M0181FF0M01FE0H020,M0181AE80M0HEI020,M0181FF80L01FF0H040,N080BA80L01FE0,N0C0V040,N0E0V080,M0160U0180,N020U01,N070U01,N010U02,N0180T04,O08,O040T0C,O060T08,O070S010,O0380R020,O01C0R040,P0C0R080,P070Q01,P0380P02,P01C0P04,Q060P08,Q030O030,Q0180N0E0,R070M03," _
               & "R0180L0E,R01D0K05810,S03A0J0A0,T0H7H017,U0BFEA0,W"
      Print #1, "01,,::::::::::::::::~DG001.GRF,02048,016,"
      Print #1, ",:::::::::::::::::::::::::R010H010,,S07FHF0,R02F0H0FC0,Q01F10H0171,Q0280K08,P0140L03,P0680M080,O01D0N070,O0180N028,O060O01C,O0E0P0E,N0180P03,N020Q0180,N040R040,N0C0R0A0,L0H180R030,M020S018,M060T08,M040T0C,M0C0T04,M080T02,L0180T03,L030U01,L030U0180,L020V080,L060V040,L080EE80M0EA0H080,L0C0FFC0L01FF0H040,L080FB80L01BA0H020,K0180FFC0L01FF0H020,L080EE80M0HEI020,K0100FFC010H0101FF01030,N0HB80L01FE0H020,K03007FC0L01FF0H030,K0200EAERE0,K0300FTF010,K02002BFBFBFBFBFBFBE010,K03007FSF010,K02002AAEQEH08,K07003FSF018,K02001ABBFOFB018,K060H0TF018,K020H0HAREH08,K06013FSF018,K02003FFBFBFBFBFBFBB018,K07007FSF018,K0200ESEA0,J01300FTF010,K0200FFBFIFBFKFA010,K0301FFC7FPF010,K0200EE80M0EA0H020,K0101FF80L01FF0H020,K0100FB80L01FA0H020,K0181FF0M01FE0H020,K0181AE80M0HEI020,K0181FF80L01FF0H040,L080BA80L01FE0,L0C0V040,L0E0V080,K0160U0180,L020U01,L070U01,L010U02,L0180T04,M08,M040T0C,M060T08,M070S010,M0380R020,M01C0R040,N0C0R080,N070Q01,N0380P02,N01C0P04,O060P08,O030O030,O0180N0E0,P070M03,P0180L0E," _
               & "P01D0K05810,Q03A0J0A0,R0H7H017,S0BFEA0,U"
      Print #1, "01,,::::::::::::::::^XA^LL0743"
      Print #1, "^PW815"
      Print #1, "^FT288,128^XG000.GRF,1,1^FS"
      Print #1, "^FT704,128^XG001.GRF,1,1^FS"
      Print #1, "^BY2,2.5,47^FT45,220^B3R,N,,Y,N"
      Print #1, "^FD" & sNumeroBarras & "^FS"
      Print #1, "^FT75,613^A0R,20,21^FH\^FD" & sDescuento & "^FS"
      Print #1, "^FT104,613^A0R,20,21^FH\^FD" & sValor & "^FS"
      Print #1, "^FT146,577^A0R,20,21^FH\^FD" & sturno & "^FS"
      Print #1, "^FT340,222^A0R,23,21^FH\^FD" & sNombreTerminal & "^FS"
      Print #1, "^FT356,526^A0R,14,12^FH\^FDNumero^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT181,29^A0R,20,19^FH\^FD" & snombrefechaingreso & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT181,29^A0R,20,19^FH\^FDFECHA LIMITE:^FS"
      'End If
      
      Print #1, "^FT287,272^A0R,20,16^FH\^FDHORA^FS"
      Print #1, "^FT214,232^A0R,20,19^FH\^FD" & sEmpresa & "^FS"
      Print #1, "^FT213,28^A0R,20,19^FH\^FDEMPRESA:^FS"
      Print #1, "^FT180,514^A0R,20,19^FH\^FDCAJA^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT118,28^A0R,20,19^FH\^FD" & sTiempoVigencia & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT118,28^A0R,20,19^FH\^FDVIGENCIA:^FS"
      'End If
      Print #1, "^FT150,29^A0R,20,19^FH\^FDFECHA SALIDA:^FS"
      Print #1, "^FT117,234^A0R,20,19^FH\^FD" & sTiempoPermanencia & "^FS"
      Print #1, "^FT290,538^A0R,20,16^FH\^FDPLACA^FS"
      Print #1, "^FT263,406^A0R,20,21^FH\^FD" & sNumeroInterno & "^FS"
      Print #1, "^FT287,406^A0R,20,16^FH\^FDNo INT.^FS"
      Print #1, "^FT263,271^A0R,20,21^FH\^FD" & sHora & "^FS"
      Print #1, "^FT150,234^A0R,20,21^FH\^FD" & sFechaSalida & "^FS"
      Print #1, "^FT40,615^A0R,25,24^FH\^FD" & sValorPagar & "^FS"
      Print #1, "^FT183,233^A0R,20,21^FH\^FD" & sFechaIngreso & "^FS"
      Print #1, "^FT77,493^A0R,20,19^FH\^FDDescuento^FS"
      Print #1, "^FT105,493^A0R,20,19^FH\^FDValor:^FS"
      Print #1, "^FT40,495^A0R,25,24^FH\^FDValor Pagar^FS"
      Print #1, "^FT147,514^A0R,20,19^FH\^FDTURNO^FS"
      Print #1, "^FT211,576^A0R,20,21^FH\^FD" & sCajero & "^FS"
      Print #1, "^FT179,577^A0R,20,21^FH\^FD" & sCaja & "^FS"
      Print #1, "^FT213,513^A0R,20,19^FH\^FDCAJERO^FS"
      Print #1, "^FT265,539^A0R,20,21^FH\^FD" & sPlaca & "^FS"
      Print #1, "^FT289,93^A0R,20,16^FH\^FDFECHA PAGO^FS"
      Print #1, "^FT262,92^A0R,20,21^FH\^FD" & sFecha & "^FS"
      Print #1, "^FT320,281^A0R,17,16^FH\^FD" & sNit & "^FS"
      Print #1, "^FT343,596^A0R,34,33^FH\^FD" & sNumero & "^FS"
      Print #1, "^FT367,247^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
      Print #1, "^BY2,2.5,47^FT453,220^B3R,N,,Y,N"
      Print #1, "^FD" & sNumeroBarras & "^FS"
      Print #1, "^FT483,613^A0R,20,21^FH\^FD" & sDescuento & "^FS"
      Print #1, "^FT512,613^A0R,20,21^FH\^FD" & sValor & "^FS"
      Print #1, "^FT554,577^A0R,20,21^FH\^FD" & sturno & "^FS"
      Print #1, "^FT748,222^A0R,23,21^FH\^FD" & sNombreTerminal & "^FS"
      Print #1, "^FT764,526^A0R,14,12^FH\^FDNumero^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT589,29^A0R,20,19^FH\^FD" & snombrefechaingreso & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT589,29^A0R,20,19^FH\^FDFECHA LIMITE:^FS"
      'End If
      
      Print #1, "^FT695,272^A0R,20,16^FH\^FDHORA^FS"
      Print #1, "^FT622,232^A0R,20,19^FH\^FD" & sEmpresa & "^FS"
      Print #1, "^FT621,28^A0R,20,19^FH\^FDEMPRESA:^FS"
      Print #1, "^FT588,514^A0R,20,19^FH\^FDCAJA^FS"
      'If GrdParque.TextMatrix(i, 7) = "P" Then
         Print #1, "^FT526,28^A0R,20,19^FH\^FD" & sTiempoVigencia & "^FS"
      'ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
      '   Print #1, "^FT526,28^A0R,20,19^FH\^FDVIGENCIA:^FS"
      'End If
      
      Print #1, "^FT558,29^A0R,20,19^FH\^FDFECHA SALIDA:^FS"
      Print #1, "^FT525,234^A0R,20,19^FH\^FD" & sTiempoPermanencia & "^FS"
      Print #1, "^FT698,538^A0R,20,16^FH\^FDPLACA^FS"
      Print #1, "^FT671,406^A0R,20,21^FH\^FD" & sNumeroInterno & "^FS"
      Print #1, "^FT695,406^A0R,20,16^FH\^FDNo INT.^FS"
      Print #1, "^FT671,271^A0R,20,21^FH\^FD" & sHora & "^FS"
      Print #1, "^FT558,234^A0R,20,21^FH\^FD" & sFechaSalida & "^FS"
      Print #1, "^FT448,615^A0R,25,24^FH\^FD" & sValorPagar & "^FS"
      Print #1, "^FT591,233^A0R,20,21^FH\^FD" & sFechaIngreso & "^FS"
      Print #1, "^FT485,493^A0R,20,19^FH\^FDDescuento^FS"
      Print #1, "^FT513,493^A0R,20,19^FH\^FDValor:^FS"
      Print #1, "^FT448,495^A0R,25,24^FH\^FDValor Pagar^FS"
      Print #1, "^FT555,514^A0R,20,19^FH\^FDTURNO^FS"
      Print #1, "^FT619,576^A0R,20,21^FH\^FD" & sCajero & "^FS"
      Print #1, "^FT587,577^A0R,20,21^FH\^FD" & sCaja & "^FS"
      Print #1, "^FT621,513^A0R,20,19^FH\^FDCAJERO^FS"
      Print #1, "^FT673,539^A0R,20,21^FH\^FD" & sPlaca & "^FS"
      Print #1, "^FT697,93^A0R,20,16^FH\^FDFECHA PAGO^FS"
      Print #1, "^FT670,92^A0R,20,21^FH\^FD" & sFecha & "^FS"
      Print #1, "^FT728,281^A0R,17,16^FH\^FD" & sNit & "^FS"
      Print #1, "^FT751,596^A0R,34,33^FH\^FD" & sNumero & "^FS"
      Print #1, "^FT775,247^A0R,17,16^FH\^FDAlcaldia Mayor de Bogota D.C.^FS"
      Print #1, "^PQ1,0,1,Y^XZ"
      Print #1, "^XA^ID000.GRF^FS^XZ"
      Print #1, "^XA^ID001.GRF^FS^XZ"
      
      Close #1
      
   ElseIf G_ImpCon = "C" Or G_ImpCon = "N" Then
   
'            imprimir.Show 1
'            If G_Imprimir = "0" Then
'               MsgBox "Se canceló la impresión no se grabó la información", 48, "Advertencia"
'               Dbs.RollbackTrans
'               Exit Sub
'            End If
      
      If G_ImpEpsonZebra = "P" Then
         
         reporte = RpPath + "\pagparqueadero.rpt"
         sCantidad = 2
      ElseIf G_ImpEpsonZebra = "Z" Then
         
         reporte = RpPath + "\pagparqueaderozebra.rpt"
         sCantidad = 1
      End If
      rpt.ReportFileName = reporte
      rpt.Connect = G_Conexion_reportes
      
      For j = 1 To sCantidad
   
      
         rpt.Formulas(0) = "nombreempresa = '" & nombrebasedatos & "'"                       'ok
         rpt.Formulas(1) = "numrecibo = '" & sNumero & "'"                                   'ok
         rpt.Formulas(2) = "fechapago = '" & sFecha & "'"                                    'ok
         rpt.Formulas(3) = "hora = '" & sHora & "'"                                          'ok
         rpt.Formulas(4) = "numint = '" & sNumeroInterno & "'"                               'ok
         rpt.Formulas(5) = "numplaca = '" & sPlaca & "'"                                     'ok
         rpt.Formulas(6) = "empveh = '" & sEmpresa & "'"                                     'ok
         rpt.Formulas(7) = "fechaingreso = '" & sFechaIngreso & "'"                          'ok
         rpt.Formulas(8) = "fechasalida = '" & sFechaSalida & "'"                            'ok
         rpt.Formulas(9) = "tiempo = '" & sTiempoPermanencia & "'"                           'ok
         rpt.Formulas(10) = "usuario = '" & sCajero & "'"                                    'ok
         rpt.Formulas(11) = "pc = '" & sCaja & "'"                                           'ok
         rpt.Formulas(12) = "turno = '" & sturno & "'"                                       'ok
         rpt.Formulas(13) = "valor = '" & Format(sValor, "$##,##0") & "'"                    'ok
         rpt.Formulas(14) = "descuento = '" & Format(sDescuento, "$##,##0") & "'"            'ok
         rpt.Formulas(15) = "valortotal = '" & Format(sValorPagar, "$##,##0") & "'"          'ok
         rpt.Formulas(16) = "codigobarras = '" & sNumeroBarras & "'"                         'ok
         rpt.Formulas(17) = "snombrefechaingreso = '" & snombrefechaingreso & "'"            'ok
         rpt.Formulas(18) = "stitulo = '" & sTitulo & "'"
         rpt.Formulas(19) = "stiempovigencia = '" & sTiempoVigencia & "'"                    'ok
         
         If G_ImpEpsonZebra = "P" Then
            rpt.PrinterName = G_PrinterName1
            rpt.PrinterDriver = G_PrinterDriver1
            rpt.PrinterPort = G_PrinterPort1
            'Rpt.ProgressDialog = False
         End If
         
         rpt.Destination = 1
         
         rpt.Action = 1
         
       Next j
   End If
   
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub btnconduces_Click()
   On Error Resume Next
   Unload Me
   Capconduce.Show 1
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
   
   'rtamsg = MsgBox("Esta seguro de grabar la los registros?", 36, "Confirmación")
    rtamsg = MsgBox("¿Está seguro de grabar los registros?", 32 + vbYesNo + vbQuestion + vbDefaultButton2, "Confirmación")
      
   If rtamsg = vbNo Then
      Exit Sub
   End If
   
   sFecSistema = Fn_Parametros("HORASIS", 1)
   sFecSis = Format(sFecSistema, "MM/dd/yyyy")
   sHorSis = Format(sFecSistema, "HH:mm")
   
   'Sebastián Rondón - Febrero 27 de 2025
   'Si el vehículo no cuenta con rodamiento, preguntamos al operario si desea continuar con la operación
'   If Not L_Bandera_Rodamiento Then
'        G_Mensaje_Pago = "VEHÍCULO SIN RODAMIENTO PROGRAMADO, DESEA CONTINUAR EL PAGO?"
'        G_Mensaje_Tipo = "E"
'        G_Mensaje_Opc = False
'        AviRecMult.Show 1
'
'        If Not G_Mensaje_Opc Then
'            Exit Sub
'        End If
'   End If

    'Sebastián Rondón - Abril 02 de 2025
    'Mensaje solamente informativo para confirmar si se desea realizar el proceso
'    G_Mensaje_Pago = "¿DESEA REALIZAR EL PROCESO PARA LA PLACA " & Txtnumpla.text & "? "
'    G_Mensaje_Tipo = "E"
'    G_Mensaje_Opc = False
    
'    AviRecMult.Show 1
    
'    If Not G_Mensaje_Opc Then
'        Exit Sub
'    End If
   
   'Dbs.BeginTrans
   capparqu.Enabled = False
   On Error GoTo Err_Transac
   
   For i = 1 To GrdParque.rows - 1
      GrdParque.col = 0
      GrdParque.Row = i
      
      If GrdParque.text = "->" Then
         
         Dbs.BeginTrans
         
         GrdParque.col = 1
         dato1 = GrdParque.text     'numero tiquete
         
         QRY = "UPDATE coparqueop " _
             & "SET pafecpag = to_date('" & sFecSis & "','MM/dd/yyyy') , " _
                & " pacaja = '" & Pc & "', " _
                & " paturno = " & G_Turno & ", " _
                & " pacerrado= 'N', " _
                & " paestado = 'C' " _
           & "WHERE panumero = " & dato1 & ""
      
         Dbs.Execute QRY
         
         'WECC 0813/2019 CAU No 1254
         '**********************************************************************
'         QRY = "UPDATE optempo1 " _
'                & "SET temnumero1 = temnumero1 + " & GrdParque.TextMatrix(i, 5) & " " _
'              & "WHERE temtext4 = '" & Trim(Txtnumpla.text) & "' " _
'                & "AND temnumero3 = " & dato1 & " " _
'                & "AND tempc = '" & Trim(Pc) & "'"
'
'         Dbs.Execute QRY, sRecordAffected
'
'         If sRecordAffected = 0 Then
'            QRY = "INSERT INTO optempo1(temtext4, temnumero1, temnumero3, tempc)" _
'                     & "VALUES('" & Trim(Txtnumpla.text) & "'," & GrdParque.TextMatrix(i, 5) & ", " & dato1 & ", '" & Trim(Pc) & "' )"
'
'            Dbs.Execute QRY
'         End If
         '**********************************************************************
         
        'Sebastián Rondón - Marzo 20 de 2024
        'Por orden del jefe Álvaro se cambia el pago de las permanencias y multas
        'Si una infracción es pagada en la terminal Norte y Sur, se crea automaticamente una tasa sin viaje
        'Una tasa sin viaje es una tasa con valor 0, la cual solo se crea, para asociar el pago de la multa, y esta se pueda enviar a la DIAN
        'Si es un pago en la terminal Central, valida si tiene rodamiento, en cuyo caso no, crea de igual manera la tasa sin viaje
        
        'Sebastián Rondón - Febrero 26 de 2025
        'Se empieza a validar rodamiento para todas las terminales
        If Mid(G_Tipo_Terminal, 1, 1) = "N" Or Mid(G_Tipo_Terminal, 1, 1) = "S" Or Mid(G_Tipo_Terminal, 1, 1) = "C" Then
            
            'Sebastián Rondón - Marzo 10 de 2025
            'Por orden del jefe Álvaro, el sistema ahora siempre en el pago de una permanencia, va generar TU Sin Viaje
            'Se hace esto ya que muchas veces los operarios de la caja, no realizan el procedimiento de pago correctamente
            'L_Bandera_Rodamiento = False
            
            'Sebastián Rondón - Abril 01 de 2025
            'Desde ahora, cada pago genera Tasa de Uso Sin Viaje siempre
            L_Bandera_Rodamiento = False
            If L_Bandera_Rodamiento Then
                 'Felipe Mendoza - Marzo 12/2024
                 'Se hace el insert de los pagos en una nueva tabla NO TEMPORAL
                 QRY = "UPDATE opreccaja " _
                        & "SET rcvalper = rcvalper + " & GrdParque.TextMatrix(i, 5) & ", " _
                        & "rcfecha = TO_DATE('" & sFecSis & "', 'MM/DD/YYYY'), " _
                        & "rchora = TO_DATE('" & sHorSis & "','HH24:MI') " _
                      & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
                        & "AND rcnumero = " & dato1 & " " _
                        & "AND rcnomcaj = '" & Trim(Pc) & "'"
        
                 Dbs.Execute QRY, sRecordAffected
        
                 If sRecordAffected = 0 Then
                    QRY = "INSERT INTO opreccaja(rcplaca, rcvalper, rcnumero, rcnomcaj, rcfecha, rchora)" _
                             & "VALUES('" & Trim(Txtnumpla.text) & "'," & GrdParque.TextMatrix(i, 5) & ", " _
                             & dato1 & ", '" & Trim(Pc) & "', TO_DATE('" & sFecSis & "','MM/DD/YYYY'), " _
                             & "TO_DATE('" & sHorSis & "','HH24:MI') )"
                    Dbs.Execute QRY
                 End If
            Else
                sFechaCompleta = sFecSis & " " & sHorSis
             
                'Consultamos el tipo de vehículo
                QRY = "SELECT hvtipveh FROM ophojveh WHERE hvplaca = '" & Trim(Txtnumpla.text) & "' "
                Set Tbl1 = Dbs.Execute(QRY)
                
                If Not Tbl1.EOF Then
                   TipVeh = Tbl1!hvtipveh
                End If
                 
                'Consultamos el último consecutivo de conduces para la caja
                QRY = "SELECT caconcaj FROM cocaja WHERE canomcaj = '" & Pc & "' "
                Set Tbl1 = Dbs.Execute(QRY)
                
                If Not Tbl1.EOF Then
                    numconduce = Tbl1!caconcaj + 1
                    
                    'Creamos la Tasa de uso sin viaje
                    QRY = "INSERT INTO coconduc (conumero, cocodrut, cocodemp, cofecsal, cohorven, cohorsal, coplaca, coclase, " _
                                & "coconduc, cooritra, cofuncio, coturno, coestado, comodo, conumint, coobser, conomcaj , " _
                                & "cocredit, coterminal, cohorasig) " _
                                & "VALUES(" & numconduce & ", '000', '" & Txtcodemp.text & "', TO_DATE('" & sFecSis & "','MM/DD/YYYY'), " _
                                & "TO_DATE('" & sHorSis & "','HH24:MI'), TO_DATE('" & sHorSis & "','HH24:MI'), '" & Txtnumpla.text & "', '" & TipVeh & "', " _
                                & "'0', 'SV', '" & Login & "', " & G_Turno & ", 'Activo', 'D', " & Txtnumint.text & ", " _
                                & "'TASA DE USO PARA PAGO DE PERMANENCIA', '" & Pc & "', 'N', '" & Mid(G_Tipo_Terminal, 1, 1) & "', " _
                                & "TO_DATE('" & sFechaCompleta & "','MM/DD/YYYY HH24:MI')) "
                    Dbs.Execute (QRY)
                    
                    'Asignamos la tasa de uso en la multa pagada
                    QRY = "UPDATE coparqueop " _
                            & "SET patasapag = " & numconduce & " " _
                            & "WHERE panumero = " & dato1 & " "
                    Dbs.Execute (QRY)
                    
                    'Aumentamos el número consecutivo del conduce
                    QRY = "UPDATE cocaja " _
                            & "SET caconcaj = " & numconduce & " " _
                            & "WHERE canomcaj = '" & Pc & "' "
                    Dbs.Execute (QRY)
                    
                    'Sebastián Rondón - Abril 02 de 2025
                    'Registro en OPRECCAJA por si van a pagar la tasa de uso de al instante y se vea la permanencia en la calculadora
            
                     '06/17/2025 Camilo Campos se comenta por razones operativas (estaba dejando multas/permanencias sin TU relacionada)

'                    QRY = "UPDATE opreccaja " _
'                                & "SET rcvalper = rcvalper + " & GrdParque.TextMatrix(i, 5) & ", " _
'                                & "rcfecha = TO_DATE('" & sFecSis & "', 'MM/DD/YYYY'), " _
'                                & "rchora = TO_DATE('" & sHorSis & "','HH24:MI') " _
'                                & "WHERE rcplaca = '" & Trim(Txtnumpla.text) & "' " _
'                                & "AND rcnumero = " & dato1 & " " _
'                                & "AND rcnomcaj = '" & Trim(Pc) & "'"
            
                     Dbs.Execute QRY, sRecordAffected
            
                     If sRecordAffected = 0 Then
                        QRY = "INSERT INTO opreccaja(rcplaca, rcvalper, rcnumero, rcnomcaj, rcfecha, rchora)" _
                                 & "VALUES('" & Trim(Txtnumpla.text) & "'," & GrdParque.TextMatrix(i, 5) & ", " _
                                 & dato1 & ", '" & Trim(Pc) & "', TO_DATE('" & sFecSis & "','MM/DD/YYYY'), " _
                                 & "TO_DATE('" & sHorSis & "','HH24:MI') )"
                        Dbs.Execute QRY
                     End If
                    
                    'Mostramos el mensaje de que el recaudo no se reflejará en la venta de la tasa
                    'L_Bandera_Mensaje = True
                End If
            End If
        End If
        
        
         sminutosDif = (DateDiff("n", GrdParque.TextMatrix(i, 3), GrdParque.TextMatrix(i, 4)))
         sdiferencia = (DateDiff("n", GrdParque.TextMatrix(i, 3), GrdParque.TextMatrix(i, 4)))
         
         horest = DateDiff("s", GrdParque.TextMatrix(i, 3), GrdParque.TextMatrix(i, 4))

         horest = horest / 86400
         horper = horest
               
               
         sMinutosIniciarCobro = G_Minutos_Minimos_Parqueadero        'minutos para iniciar cobro de parqueadero
         sMinutosMaximos = 60            'el cobro es por horas
               
        'Controlamos que solocobre parqueadero  contando 12 horas a partir de los minutos inial de cobro
         sdiferencia = sminutosDif
         If sdiferencia > sMinutosIniciarCobro Then
            sdiferencia = sdiferencia - sMinutosIniciarCobro
            cociente = sdiferencia \ sMinutosMaximos
            residuo = sdiferencia Mod sMinutosMaximos
            
            If residuo < sMinutosIniciarCobro Then
               If residuo <> 0 Then
                  cociente = cociente
               End If
            Else
               cociente = cociente + 1
            End If
            
            sTiempoPermanencia = cociente & " Horas " & residuo & " minutos"
         Else
            sTiempoPermanencia = "0 Horas 0 minutos"
         End If
         
         
         sNombreTerminal = G_NomEmpresa
         sNit = "NIT.860.052.155-6"
         sFecha = Format(sFecSistema, "MMM/dd/yyyy")
         sHora = Format(sFecSistema, "hh:mm")
         sNumeroInterno = Txtnumint.text
         sPlaca = Txtnumpla.text
         sEmpresa = Lblnomemp.Caption
         sFechaIngreso = GrdParque.TextMatrix(i, 3)
         sFechaSalida = GrdParque.TextMatrix(i, 4)
         sValorPagar = GrdParque.TextMatrix(i, 5)
         sDescuento = GrdParque.TextMatrix(i, 6)
         sValor = CDbl(sValorPagar) + CDbl(sDescuento)
         sCajero = Login
         sCaja = Pc
         sturno = G_Turno
         sNumero = GrdParque.TextMatrix(i, 1)
         sNumeroBarras = "*" & GrdParque.TextMatrix(i, 1) & "*"
         If GrdParque.TextMatrix(i, 7) = "P" Then
            snombrefechaingreso = "FECHA INGRESO"
            sTitulo = "PAGO POR TIEMPO DE PERMANENCIA"
            sTiempoVigencia = "TIEMPO PERMANENCIA:"
         ElseIf GrdParque.TextMatrix(i, 7) = "V" Then
            snombrefechaingreso = "FECHA LIMITE"
            sTitulo = "PAGO POR VIGENCIA"
            sTiempoVigencia = "VIGENCIA:"
         End If
         
'         If L_Bandera_Mensaje Then
'            G_Mensaje_Pago = "POR FAVOR RECAUDE EL VALOR DE LA MULTA/PERMANENCIA, YA QUE NO SE VERÁ REFLEJADO EN LA VENTA DE LA TASA DE USO"
'            G_Mensaje_Tipo = "M"
'            AviRecMult.Show 1
'         End If
         
         'Sebastián Rondón - Marzo 03  de 2025
         'Realizó el commit antes de realizar la impresión de los pagos
         Dbs.CommitTrans
         
         
         On Error GoTo Errores
         
         Call ReImprimirPago(sTiempoPermanencia, sNombreTerminal, sNit, sFecha, _
                           sHora, sNumeroInterno, sPlaca, sEmpresa, _
                           sFechaIngreso, sFechaSalida, sValorPagar, sDescuento, _
                           sValor, sCajero, sCaja, sturno, _
                           sNumero, sNumeroBarras, snombrefechaingreso, sTitulo, _
                           sTiempoVigencia)
      End If
   Next i
   
   'Dbs.CommitTrans
   capparqu.Enabled = True
   
   MsgBox "Los registros se grabaron exitosamente", 64, "O.K."
   Btnlimpiar_Click
   
   Exit Sub
   
Err_Transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
Errores:
   capparqu.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub



Private Sub Btnimprimir_Click()
   On Error GoTo Errores
   If Txtnumpla.text <> "" Then
         
         
      sFecSistema = Fn_Parametros("HORASIS", 1)
      sFecSis = Format(sFecSistema, "MM/dd/yyyy")
   
      QRY = "SELECT panumero, paplaca, pafecini, " _
                 & "pafecsal, pavalor, paaviso, " _
                 & "paestado, padescuento, patipo, pafecpag " _
            & "FROM coparqueop " _
          & " WHERE panumero =  (SELECT MAX(panumero) " _
                                & "FROM coparqueop " _
                               & "WHERE paplaca = '" & Txtnumpla.text & "' " _
                                & " AND pafecpag = TO_DATE('" & sFecSis & "','MM/dd/yyyy') " _
                                & " AND pacaja = '" & Pc & "' " _
                                & " AND pavalor <> 0 " _
                                & " AND paestado = 'C' )"
      
      Set Tbl1 = Dbs.Execute(QRY)
      
      If Not Tbl1.EOF Then
         sTiempoPermanencia = "0 Horas 0 minutos"
         
         sminutosDif = (DateDiff("n", Tbl1!pafecini, Tbl1!pafecsal))
         sdiferencia = (DateDiff("n", Tbl1!pafecini, Tbl1!pafecsal))
         
         horest = DateDiff("s", Tbl1!pafecini, Tbl1!pafecsal)

         horest = horest / 86400
         horper = horest
               
               
         sMinutosIniciarCobro = G_Minutos_Minimos_Parqueadero        'minutos para iniciar cobro de parqueadero
         sMinutosMaximos = 60            'el cobro es por horas
               
               'Controlamos que solocobre parqueadero  contando 12 horas a partir de los minutos inial de cobro
         sdiferencia = sminutosDif
         If sdiferencia > sMinutosIniciarCobro Then
            sdiferencia = sdiferencia - sMinutosIniciarCobro
            cociente = sdiferencia \ sMinutosMaximos
            residuo = sdiferencia Mod sMinutosMaximos
            
            If residuo < sMinutosIniciarCobro Then
               If residuo <> 0 Then
                  cociente = cociente
               End If
               
            Else
               cociente = cociente + 1
            End If
            
            sTiempoPermanencia = cociente & " Horas " & residuo & " minutos"
         Else
            sTiempoPermanencia = "0 Horas 0 minutos"
         End If
         
         
         sNombreTerminal = G_NomEmpresa
         sNit = "NIT.860.052.155-6"
         sFecha = Format(Tbl1!pafecpag, "MMM/dd/yyyy")
         sHora = Format(sFecSistema, "hh:mm")
         sNumeroInterno = Txtnumint.text
         sPlaca = Txtnumpla.text
         sEmpresa = Lblnomemp.Caption
         sFechaIngreso = Format(Tbl1!pafecini, "MMM/dd/yyyy hh:mm")
         sFechaSalida = Format(Tbl1!pafecsal, "MMM/dd/yyyy hh:mm")
         sValorPagar = Tbl1!pavalor
         sDescuento = Tbl1!padescuento
         sValor = CDbl(sValorPagar) + CDbl(sDescuento)
         sCajero = Login
         sCaja = Pc
         sturno = G_Turno
         sNumero = Tbl1!panumero
         sNumeroBarras = "*" & Tbl1!panumero & "*"
         If Tbl1!patipo = "P" Then
            snombrefechaingreso = "FECHA INGRESO"
            sTitulo = "PAGO POR TIEMPO DE PERMANENCIA"
            sTiempoVigencia = "TIEMPO PERMANENCIA:"
         ElseIf Tbl1!patipo = "V" Then
            snombrefechaingreso = "FECHA LIMITE"
            sTitulo = "PAGO POR VIGENCIA"
            sTiempoVigencia = "VIGENCIA:"
         End If
         
         Call ReImprimirPago(sTiempoPermanencia, sNombreTerminal, sNit, sFecha, _
                           sHora, sNumeroInterno, sPlaca, sEmpresa, _
                           sFechaIngreso, sFechaSalida, sValorPagar, sDescuento, _
                           sValor, sCajero, sCaja, sturno, _
                           sNumero, sNumeroBarras, snombrefechaingreso, sTitulo, _
                           sTiempoVigencia)
      Else
         MsgBox "El vehículo no reporta un pago en el día de hoy", 48, "Advertencia"
         Txtnumpla.SetFocus
         Exit Sub
      End If
      
   Else
      MsgBox "Falta digitar el número de placa", 48, "Advertencia"
      Txtnumpla.SetFocus
      Exit Sub
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnlimpiar_Click()
   On Error GoTo Errores
   Txtnumpla.Enabled = True
   ChkReImprimir.Value = 0
   Btnimprimir.Enabled = False
   Txtnumpla.text = ""
   Txtnumint.text = ""
   Txtcodemp.text = ""
   Lblnomemp.Caption = ""
   mskafac.text = ""
   Msktotliq.text = ""
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

Private Sub ChkReImprimir_Click()
   Btnimprimir.Enabled = False
   Btngrabar.Enabled = True
   If ChkReImprimir.Value = 1 Then
      Btnimprimir.Enabled = True
      Btngrabar.Enabled = False
   End If
End Sub

Private Sub Form_Activate()
   If Txtnumpla.text <> "" Then
      Txtnumpla_LostFocus
   End If
End Sub

Private Sub Form_Load()
   btnconduces.Enabled = False
   If main.Mnuconduces.Enabled = True Then
      btnconduces.Enabled = True
   End If
   DIB_GRILLA
   L_Bandera_Mensaje = False
End Sub

Private Sub GrdParque_DblClick()
   If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
      GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
      GrdParque.col = 5
      If mskafac.text = "" Then
         mskafac.text = GrdParque.text
      Else
         mskafac.text = CDbl(mskafac.text) + CDbl(GrdParque.text)
      End If
      
   ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
      GrdParque.TextMatrix(GrdParque.Row, 0) = ""
      GrdParque.col = 5
      mskafac.text = CDbl(mskafac.text) - CDbl(GrdParque.text)
   End If
   
   GrdParque.col = 0
   If GrdParque.rows - 1 > GrdParque.Row Then
      GrdParque.Row = GrdParque.Row + 1
   Else
      'Btngrabar.SetFocus
   End If


End Sub

Private Sub GrdParque_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If KeyAscii% = 32 Then
      If GrdParque.TextMatrix(GrdParque.Row, 0) = "" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = "->"
         GrdParque.col = 5
         If mskafac.text = "" Then
            mskafac.text = GrdParque.text
         Else
            mskafac.text = CDbl(mskafac.text) + CDbl(GrdParque.text)
         End If
         
      ElseIf GrdParque.TextMatrix(GrdParque.Row, 0) = "->" Then
         GrdParque.TextMatrix(GrdParque.Row, 0) = ""
         GrdParque.col = 5
         mskafac.text = CDbl(mskafac.text) - CDbl(GrdParque.text)
      End If
      
      GrdParque.col = 0
      If GrdParque.rows - 1 > GrdParque.Row Then
         GrdParque.Row = GrdParque.Row + 1
      Else
            'Btngrabar.SetFocus
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
'            Btngrabar.SetFocus
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
      If Len(Txtnumpla.text) = 6 Then
         fqry01 = "SELECT * " _
                  & "FROM ophojveh " _
                 & "WHERE hvplaca = '" & Txtnumpla.text & "' "
                
         Set Tbl1 = Dbs.Execute(fqry01)
         If Not Tbl1.EOF Then
            Txtnumint.text = Tbl1!hvnumint
            Txtcodemp.text = Tbl1!hvcodemp
            Tbl1.Close
            fqry01 = "SELECT henombre " _
                    
            If Len(Trim(Txtcodemp.text)) = 3 Then
               fqry01 = fqry01 & "FROM ophojemp "
            ElseIf Len(Trim(Txtcodemp.text)) = 4 Then
               fqry01 = fqry01 & "FROM ophojempcon "
            End If
                    
            fqry01 = fqry01 & "WHERE hecodigo = '" & Txtcodemp.text & "' "
           
            Set Tbl1 = Dbs.Execute(fqry01)
            If Not Tbl1.EOF Then
               Lblnomemp.Caption = Tbl1!henombre
            End If
            Tbl1.Close
            Traer_Datos
         Else
            Tbl1.Close
            MsgBox "El vehículo no existe", 48, "Advertencia"
            Txtnumpla.SetFocus
            Exit Sub
         End If
      End If
   End If
salir:
   Exit Sub
   
Message_Next:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
 
End Sub

