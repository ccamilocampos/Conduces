VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Avimul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multas Pendientes"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Descuentos"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   6015
      Begin VB.TextBox Txtpor4 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtpor1 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtdia2 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtpor2 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtdia3 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtpor3 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtdia4 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Txtdia1 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "%"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5640
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "%"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta Dias"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Mas de:"
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
         Left            =   4800
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Dias"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Dias"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
      Left            =   5280
      Picture         =   "avimul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir de la forma actual"
      Top             =   3600
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grdavimul 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSMask.MaskEdBox Msktotmul 
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
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
   Begin MSMask.MaskEdBox mskValorPagar 
      Height          =   300
      Left            =   2400
      TabIndex        =   25
      Top             =   4320
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
   Begin VB.Label Label11 
      Caption         =   "Valor a pagar :"
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
      TabIndex        =   24
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "No Multas.:"
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
      Top             =   3240
      Width           =   1575
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
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
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
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
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
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Total Multas :"
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
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "Avimul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub DIB_GRILLA()
   Grdavimul.rows = 1
   Grdavimul.Cols = 8
   Grdavimul.Clear
   
   s$ = "Núm. Multa|Fec. Multa|Fec.Grab|Valor |Concepto |Aviso No.|Codigo|Saldo"
   Grdavimul.FormatString = s$
   Grdavimul.Row = 0
   Grdavimul.col = 0
   Grdavimul.ColWidth(0) = 1200
   Grdavimul.ColWidth(1) = 1300
   Grdavimul.ColWidth(2) = 1300
   Grdavimul.ColWidth(3) = 1000
   Grdavimul.ColWidth(4) = 3500
   Grdavimul.ColWidth(5) = 1000
   Grdavimul.ColWidth(6) = 800
   Grdavimul.ColWidth(7) = 1200
                    
   For i = 0 To 7
       Grdavimul.col = i
       Grdavimul.CellFontBold = 1
   Next i

End Sub
   

Private Sub CARGAR()
   Dim sDatos As String
   On Error GoTo Errores
   
   
   fqry01 = "SELECT * " _
            & "FROM opmultas " _
           & "WHERE munumpla = '" & Capconduce.Txtnumpla.text & "' " _
             & "AND (muestado = 'Pendiente' OR muestado = 'Abonada') ORDER BY mufecha "

   Set Ftb01 = Dbs.Execute(fqry01)
   
   G_BloqTasaUso = False
   G_BloqSin_Aviso = False
   sDatos = ""
   sMultasSinAvisos = ""
   Do While Not Ftb01.EOF
      avisos = Ftb01!muaviso + 1
       
      fqry02 = "SELECT * " _
               & "FROM geparsis " _
              & "WHERE pscod = '" & Ftb01!mucodinf & "' "
       
      Set Ftb02 = Dbs.Execute(fqry02)
      
      If Not Ftb02.EOF Then
         desmul = Ftb02!psdes
      End If
      Ftb02.Close
          
    '      If ftb02!mucodinf = 13 Then
    '        G_BloqTasaUso = True
    '      End If
          
      G_bloqTUxMulta = ""
      G_bloqTUxMulta = Traer_Descripcion(Ftb01!mucodinf, "MULTAS BLOQUEADAS EN TU")
      
      If G_bloqTUxMulta <> "" Then
        G_BloqTasaUso = True
        sDatos = G_bloqTUxMulta
      End If
      
      sBandera = Traer_Descripcion(Ftb01!mucodinf, "MULTAS A COBRAR SIN AVISOS")
      
      If sBandera <> "" Then
         If sMultasSinAvisos = "" Then
            sMultasSinAvisos = Ftb01!mucodinf
         Else
            sMultasSinAvisos = sMultasSinAvisos & "," & Ftb01!mucodinf
         End If
         G_BloqSin_Aviso = True
      End If
      
      entry = CStr(Ftb01!munumero) + Chr(9) + CStr(Ftb01!mufecha) + Chr(9) + CStr(Ftb01!mufecsis) + Chr(9) + _
                CStr(Ftb01!munetpag) + Chr(9) + desmul + Chr(9) + CStr(avisos) + Chr(9) + CStr(Ftb01!mucodinf) + Chr(9) + CStr(Ftb01!musaldo)
      Grdavimul.AddItem entry            ' Agrega la multa.

      Ftb01.MoveNext
   Loop
       
   If G_bloqTUxMulta = "" And sDatos <> "" Then
        G_bloqTUxMulta = sDatos
   End If
       
       
   'numero maximo de avisos
   fqry01 = "SELECT MAX(muaviso), COUNT(munumero) " _
            & "FROM opmultas " _
           & "WHERE munumpla = '" & Capconduce.Txtnumpla.text & "' " _
             & "AND (muestado = 'Pendiente' OR muestado = 'Abonada') "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      If IsNull(Ftb01(0)) Then
        lblnumavi.Caption = 1
       lblnummul.Caption = Ftb01(1)
       NUMAVISOS = Ftb01(0) + 1
      Else
       lblnumavi.Caption = Ftb01(0) + 1
       lblnummul.Caption = Ftb01(1)
       NUMAVISOS = Ftb01(0) + 1
     End If
     Ftb01.Close
     
      fqry01 = "SELECT munumero " _
              & "FROM opmultas " _
             & "WHERE munumpla = '" & Capconduce.Txtnumpla.text & "' "
             
      If G_BloqSin_Aviso = True Then
         fqry01 = fqry01 & "AND mucodinf IN (" & sMultasSinAvisos & ")"
      Else
         fqry01 = fqry01 & "AND muaviso = " & NUMAVISOS - 1 & " "
      End If
      
      fqry01 = fqry01 & " AND (muestado = 'Pendiente' OR muestado = 'Abonada') " _
          & "ORDER BY mufecha "
          
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         G_Numero_Comparendo_Antiguo = Ftb01!munumero
      End If
      Ftb01.Close
   End If
   
   
   
   
   'total de multas

   If Grdavimul.text <> "" And Grdavimul.rows > 1 Then
      dato1 = 0
      Grdavimul.col = 3
      For i = 1 To Grdavimul.rows - 1
         Grdavimul.Row = i
         dato1 = dato1 + CDbl(Grdavimul.text)
      Next i
      Msktotmul.text = dato1
   End If
   Exit Sub
Errores:
    MsgBox Err.Description, 16, "Error " & Err.Number
    Exit Sub
   
       
       'MsgBox "El Vehiculo Actualmente Tiene Una Multa PENDIENTE." & Chr(13) & Chr(13) & "MULTA No.  " & ftb01!munumero & " Por $ " & ftb01!munetpag & Chr(13) & Chr(13) & "AVISO No. " & avisos & Chr(13) & Chr(13) & "Fecha Multa: " & Chr(9) & Format(ftb01!mufecha, "mmm/dd/yyyy") & Chr(13) & Chr(13) & "Por Concepto de: " & Chr(9) & desmul, 64, "AVISO."
End Sub


Private Sub Btnsalir_Click()
    Dim i As Integer
    Dim fila As Integer

    If filasSeleccionadasMultas.Count = 0 Then
        MsgBox "No se seleccionaron multas.", vbInformation
        G_multas_vacias = True
        Unload Me
        Exit Sub
    End If

    ReDim multasGuardadas(1 To filasSeleccionadasMultas.Count)

    For i = 1 To filasSeleccionadasMultas.Count
        fila = filasSeleccionadasMultas(i)

        Grdavimul.Row = fila

        Grdavimul.col = 0
        multasGuardadas(i).id = Trim(Grdavimul.text)

        Grdavimul.col = 1
         multasGuardadas(i).fechaMulta = Trim(Grdavimul.text)
         
        Grdavimul.col = 3
        multasGuardadas(i).valor = Trim(Grdavimul.text)
        
        Grdavimul.col = 6
         multasGuardadas(i).codinf = Trim(Grdavimul.text)
         DoEvents
    Next i
    
    G_multas_vacias = False
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        ' El usuario hizo clic en la "X"
        MsgBox "No se permite cerrar el formulario desde aquí.", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
   DIB_GRILLA
   CARGAR
   mskValorPagar.text = "0"
   
   Set filasSeleccionadasMultas = New Collection 'para resetear las multas seleccionadas
   

   
   'carga paremetro de descuentos en multas
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'DIASMUL1' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Txtdia1 = Ftb01!psval
   End If
   Ftb01.Close
   
   fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'DIASMUL2' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Txtdia2 = Ftb01!psval
   End If
   Ftb01.Close
   fqry01 = "SELECT psval " _
         & "FROM geparsis " _
        & "WHERE pscod = 'DIASMUL3' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Txtdia3 = Ftb01!psval
   End If
   Ftb01.Close
   fqry01 = "SELECT psval " _
         & "FROM geparsis " _
        & "WHERE pscod = 'DIASMUL4' "
   
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Txtdia4 = Ftb01!psval
   End If
   Ftb01.Close
   'Busco % de acuerdo a los dias
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'PORMUL1' "
   
   Set Ftb02 = Dbs.Execute(fqry02)
   
   If Not Ftb02.EOF Then
      Txtpor1 = Ftb02!psval
   End If
   Ftb02.Close
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'PORMUL2' "
   
   Set Ftb02 = Dbs.Execute(fqry02)
   
   If Not Ftb02.EOF Then
      Txtpor2 = Ftb02!psval
   End If
   Ftb02.Close
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'PORMUL3' "
   
   Set Ftb02 = Dbs.Execute(fqry02)
   
   If Not Ftb02.EOF Then
      Txtpor3 = Ftb02!psval
   End If
   Ftb02.Close
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'PORMUL4' "
   
   Set Ftb02 = Dbs.Execute(fqry02)
   
   If Not Ftb02.EOF Then
      Txtpor4 = Ftb02!psval
   End If
   Ftb02.Close
End Sub

Private Sub Grdavimul_Click()
    Dim filaActual As Integer
    Dim i As Integer
    Dim yaSeleccionada As Boolean

    filaActual = Grdavimul.Row
    yaSeleccionada = False

    For i = 1 To filasSeleccionadasMultas.Count
        If filasSeleccionadasMultas(i) = filaActual Then
            yaSeleccionada = True
            Exit For
        End If
    Next i

    If yaSeleccionada Then
        ' Deseleccionar
        Call RestaurarColorFila(filaActual, Grdavimul)
        For i = 1 To filasSeleccionadasMultas.Count
            If filasSeleccionadasMultas(i) = filaActual Then
            
                valorCol3 = Grdavimul.TextMatrix(filaActual, 3) 'la columna 3 corresponde al valor de multas
                
                mskValorPagar.text = CDbl(mskValorPagar.text) - CDbl(valorCol3)
                filasSeleccionadasMultas.Remove i
                
                Exit For
            End If
        Next i
    Else
        ' Seleccionar
        Call PintarFilaVerde(filaActual, Grdavimul)
        
           valorCol3 = Grdavimul.TextMatrix(filaActual, 3) 'la columna 3 corresponde al valor de multas
            
                mskValorPagar.text = CDbl(mskValorPagar.text) + CDbl(valorCol3)
        
        filasSeleccionadasMultas.Add filaActual
    End If
End Sub

Private Sub MaskEdBox1_Change()

End Sub
