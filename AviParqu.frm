VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AviParqu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parqueaderos y Vigencias Pendientes de Pago"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7485
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
      Left            =   5760
      Picture         =   "AviParqu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir de la forma actual"
      Top             =   2520
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grdavimul 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSMask.MaskEdBox Msktotpar 
      Height          =   300
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.0;(#,##0.0)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskTotalPagar 
      Height          =   300
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.0;(#,##0.0)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "Total a pagar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Total Parqueaderos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Aviso No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblnumavi 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblnummul 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "No Parqueadero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "AviParqu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DIB_GRILLA()
   Grdavimul.rows = 1
   Grdavimul.Cols = 9
   Grdavimul.Clear
   
   s$ = "Consec.|Placa|Fecha Ingreso|Fecha Salida|Valor| Aviso No.|Tipo|Estado|Terminal"
   Grdavimul.FormatString = s$
   Grdavimul.Row = 0
   Grdavimul.col = 0
   Grdavimul.ColWidth(0) = 800
   Grdavimul.ColWidth(1) = 800
   Grdavimul.ColWidth(2) = 1500
   Grdavimul.ColWidth(3) = 1500
   Grdavimul.ColWidth(4) = 800
   Grdavimul.ColWidth(5) = 1000
   Grdavimul.ColWidth(6) = 1000
   Grdavimul.ColWidth(7) = 1000
   Grdavimul.ColWidth(8) = 1200
   
   For i = 0 To Grdavimul.Cols - 1
       Grdavimul.col = i
       Grdavimul.CellFontBold = 1
   Next i

End Sub
   Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        ' El usuario hizo clic en la "X"
        MsgBox "No se permite cerrar el formulario desde aquí.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub CARGAR()
   Dim sDatos As String
   On Error GoTo Errores
   
   'Sebastián Rondón - Enero 24 - 2023
   'Se consulta parámetro para la consulta a la tabla de corparqueop
   'Mientras se decide que se hace con los datos de ahí
   QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'FECPARPRO' "

    Set Tbl1 = Dbs.Execute(QRY)

    If Tbl1.EOF = False Then
        fecpar = Tbl1!psval
    End If
   
   
   scero = 0
   fqry01 = "SELECT panumero, paplaca, pafecini, " _
                 & "pafecsal, pavalor, NVL(paaviso, 0) AS paaviso, " _
                 & "paestado, patipo, " _
                 & "CASE paterminal " _
                 & "WHEN 'C' THEN 'CENTRAL'" _
                 & "WHEN 'N' THEN 'NORTE'" _
                 & "WHEN 'S' THEN 'SUR'" _
                 & "END paterminal " _
            & "FROM coparqueop " _
           & "WHERE paplaca = '" & Capconduce.Txtnumpla.text & "' AND pavalor <> " & scero & " " _
             & "AND pafecini >= TO_DATE('" & fecpar & "','MM/DD/YYYY') " _
             & "AND paestado = 'A' ORDER BY  paaviso DESC, panumero "

   Set Ftb01 = Dbs.Execute(fqry01)
   i = 0
   G_Numero_Parqueadero = 0
   G_Tipo_Parqueadero = ""
   Do While Not Ftb01.EOF
      avisos = Ftb01!paaviso '+ 1
       
      If i = 0 Then
         G_Numero_Parqueadero = Ftb01!panumero
         G_Tipo_Parqueadero = Ftb01!patipo
      End If
      i = i + 1
      entry = CStr(Ftb01!panumero) + Chr(9) + Ftb01!paplaca + Chr(9) + Format(Ftb01!pafecini, "MMM/dd/yyyy hh:mm") + Chr(9) + _
               Format(Ftb01!pafecsal, "MMM/dd/yyyy hh:mm") + Chr(9) + CStr(Ftb01!pavalor) + Chr(9) + CStr(avisos) + Chr(9) + Ftb01!patipo + Chr(9) + Ftb01!paestado + Chr(9) + Ftb01!paterminal
      
      Grdavimul.AddItem entry            ' Agrega la multa.

      Ftb01.MoveNext
   Loop
       
   'numero maximo de avisos
   fqry01 = "SELECT NVL(MAX(paaviso), 0), COUNT(panumero) " _
            & "FROM coparqueop " _
           & "WHERE paplaca = '" & Capconduce.Txtnumpla.text & "' " _
             & "AND pafecini >= TO_DATE('" & fecpar & "', 'MM/DD/YYYY') " _
             & "AND paestado = 'A' AND pavalor <> " & scero & "  "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   G_Cantidad_Parqueaderos = 0
   If Not Ftb01.EOF Then
      If IsNull(Ftb01(0)) Then
        lblnumavi.Caption = 1
        lblnummul.Caption = Ftb01(1)
        G_Avisos_Parqueadero = Ftb01(0) '+ 1
      Else
        lblnumavi.Caption = Ftb01(0) + 1
        lblnummul.Caption = Ftb01(1)
        G_Avisos_Parqueadero = Ftb01(0) '+ 1
     End If
     G_Cantidad_Parqueaderos = lblnummul.Caption
   End If
   
   'total de multas

   If Grdavimul.rows > 1 Then
      dato1 = 0
      Grdavimul.col = 4
      For i = 1 To Grdavimul.rows - 1
         Grdavimul.Row = i
         dato1 = dato1 + CDbl(Grdavimul.text)
      Next i
      Msktotpar.text = dato1
   End If
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
    Resume
End Sub


Private Sub Btnsalir_Click()
     Dim i As Integer
    Dim fila As Integer

    If filasSeleccionadasParqueadero.Count = 0 Then
        MsgBox "No se seleccionaron multas.", vbInformation
        G_parqueo_vacio = True
        Unload Me
        Exit Sub
    End If

    ReDim parqueGuardado(1 To filasSeleccionadasParqueadero.Count)

    For i = 1 To filasSeleccionadasParqueadero.Count
        fila = filasSeleccionadasParqueadero(i)

        Grdavimul.Row = fila

        Grdavimul.col = 0
        parqueGuardado(i).id = Trim(Grdavimul.text)

        Grdavimul.col = 2
         parqueGuardado(i).fecing = Trim(Grdavimul.text)
         
        Grdavimul.col = 3
        parqueGuardado(i).fecsal = Trim(Grdavimul.text)
        
        Grdavimul.col = 4
         parqueGuardado(i).valor = Trim(Grdavimul.text)
         
        Grdavimul.col = 6
        parqueGuardado(i).tipo = Trim(Grdavimul.text)
         
    Next i
   
    G_parqueo_vacio = False
   Unload Me
End Sub


Private Sub Form_Load()
   DIB_GRILLA
   CARGAR
   mskTotalPagar.text = "0"
    Set filasSeleccionadasParqueadero = New Collection 'para resetear los parqueos seleccionadas
 
End Sub



Private Sub Grdavimul_Click()
    Dim filaActual As Integer
    Dim i As Integer
    Dim yaSeleccionada As Boolean

    filaActual = Grdavimul.Row
    yaSeleccionada = False


    'IMPORTANTE El siguiente código tiene como funcionalidad UNICA Y EXCLUSIVAMENTE sumar los valores seleccionados o restarlos (netamente visual e informativo)
    For i = 1 To filasSeleccionadasParqueadero.Count
        If filasSeleccionadasParqueadero(i) = filaActual Then
            yaSeleccionada = True
            Exit For
        End If
    Next i

    If yaSeleccionada Then
        ' Deseleccionar
        Call RestaurarColorFila(filaActual, Grdavimul) 'Metodo publico en el cual agregamos la fila y el grid a pintar
        For i = 1 To filasSeleccionadasParqueadero.Count
            If filasSeleccionadasParqueadero(i) = filaActual Then
            
            valorCol3 = Grdavimul.TextMatrix(filaActual, 4) 'la columna 4 corresponde al valor de parqueo
            mskTotalPagar.text = CDbl(mskTotalPagar.text) - CDbl(valorCol3)
            
                filasSeleccionadasParqueadero.Remove i
                Exit For
            End If
        Next i
    Else
        valorCol3 = Grdavimul.TextMatrix(filaActual, 4) 'la columna 4 corresponde al valor de parqueo
            mskTotalPagar.text = CDbl(mskTotalPagar.text) + CDbl(valorCol3)
        ' Seleccionar
        Call PintarFilaVerde(filaActual, Grdavimul)
        filasSeleccionadasParqueadero.Add filaActual
    End If
End Sub
