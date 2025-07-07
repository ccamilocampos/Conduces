VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Buscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   3555
   ClientLeft      =   2925
   ClientTop       =   2805
   ClientWidth     =   6105
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5160
      Picture         =   "Buscar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   120
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
      Left            =   5160
      Picture         =   "Buscar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir de la forma actual"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Btnescoger 
      Caption         =   "&Escoger"
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
      Left            =   5160
      MousePointer    =   4  'Icon
      Picture         =   "Buscar.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar los 20 siguientes registros"
      Top             =   960
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
      Left            =   5160
      Picture         =   "Buscar.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borra datos del formulario"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Txtcribus 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Criterio de búqueda por la cual quiere que busque"
      Top             =   120
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid Grdbuscar 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   "Criterio de Búsqueda:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   6120
      Y1              =   3480
      Y2              =   3480
   End
End
Attribute VB_Name = "Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y As Integer
Dim bandera As Integer

Private Sub Dib_Grilla_Factura()
   Grdbuscar.Cols = 3
   Grdbuscar.SetFocus
   s$ = "Número   |Empresa      |Estado "
   Grdbuscar.FormatString = s$
   Grdbuscar.Row = 0
   Grdbuscar.Col = 0
   Grdbuscar.ColWidth(0) = 1300
   Grdbuscar.ColWidth(1) = 2500
   Grdbuscar.ColWidth(2) = 1500
   
   For i = 0 To Grdbuscar.Cols - 1
      Grdbuscar.Col = i
      Grdbuscar.CellFontBold = 1
   Next i
      
End Sub

Private Sub Dib_Grilla_Cupos()
   Grdbuscar.Cols = 4
   Grdbuscar.SetFocus
   s$ = "Cupo   |Nombre |Tipo|Estado "
   Grdbuscar.FormatString = s$
   Grdbuscar.Row = 0
   Grdbuscar.Col = 0
   Grdbuscar.ColWidth(0) = 1300
   Grdbuscar.ColWidth(1) = 2500
   Grdbuscar.ColWidth(2) = 1500
   Grdbuscar.ColWidth(3) = 1500
   
   
   For i = 0 To Grdbuscar.Cols - 1
      Grdbuscar.Col = i
      Grdbuscar.CellFontBold = 1
   Next i
      
End Sub

Private Sub Dib_Grilla_Usuarios()
   Grdbuscar.Cols = 3
   s$ = "Usuario|Nombre  |Estado "
   Grdbuscar.FormatString = s$
   Grdbuscar.Row = 0
   Grdbuscar.Col = 0
   Grdbuscar.ColWidth(0) = 2000
   Grdbuscar.ColWidth(1) = 3500
   Grdbuscar.ColWidth(2) = 2000
   For i = 0 To Grdbuscar.Cols - 1
      Grdbuscar.Col = i
      Grdbuscar.CellFontBold = 1
   Next i
End Sub

Private Sub Btnbuscar_Click()
   On Error GoTo Err_Btnbuscar_Click
   bandera = 0
   Grdbuscar.Clear
   Grdbuscar.rows = 1
   If forma = "c1" Then    'Consulta Conduces
      Btnescoger.Enabled = False
      Grdbuscar.Cols = 5
      Grdbuscar.SetFocus
      s$ = "Numero   |Placa  |Num. Int.|Cla.Veh  |Cod. Emp.|Cod. Ruta |Valor |Fecha|Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 1000
      Grdbuscar.ColWidth(2) = 1000
      Grdbuscar.ColWidth(3) = 1000
      Grdbuscar.ColWidth(4) = 1000
      Grdbuscar.ColWidth(5) = 1000
      Grdbuscar.ColWidth(6) = 1000
      Grdbuscar.ColWidth(7) = 1000
      Grdbuscar.ColWidth(8) = 1000
      
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
       
      QRY = "SELECT conumero, coplaca, conumint, coclase, cocodemp, cocodrut,covalor, cofecsal,coestado " _
            & "FROM coconduc  " _
           & "WHERE conumero LIKE '" & Txtcribus.text & "%' "
         
      Set Tbl1 = Dbs.Execute(QRY)
      
      Do While Tbl1.EOF = False
         entry = str(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + str(Tbl1(2)) + Chr(9) + Tbl1(3) + Chr(9) + Tbl1(4) + Chr(9) + Tbl1(5) + Chr(9) + str(Tbl1(6)) + Chr(9) + str(Tbl1(7)) + Chr(9) + Tbl1(8)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
   If forma = "1" Then    'Detecta cual es la forma empresas
      Grdbuscar.Cols = 5
      Grdbuscar.SetFocus
      s$ = "Codigo   |Nombre  |NIT  |Gerente  |Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 2000
      Grdbuscar.ColWidth(2) = 1000
      Grdbuscar.ColWidth(3) = 2500
      Grdbuscar.ColWidth(4) = 1000
                          
      For i = 0 To Grdbuscar.Cols - 1
        Grdbuscar.Col = i
        Grdbuscar.CellFontBold = 1
      Next i
              
      fqry01 = "SELECT * " _
              & "FROM ophojemp " _
             & "WHERE henombre like '" & Txtcribus.text & "%'"
             
      Set Tbl1 = Dbs.Execute(fqry01)
    
           
      Do While Tbl1.EOF = False
      
         sNombreGerente = ""
         
         If Not IsNull(Tbl1!hegerent) Then
            sNombreGerente = Tbl1!hegerent
         End If
         
         entry = Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + sNombreGerente + Chr(9) + Tbl1!heestado
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
      
   If forma = "2" Then    'Detecta cual es la forma Vehiculo
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      s$ = "Codigo   |Conductor Activo |Num. Interno |Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
      Grdbuscar.ColWidth(3) = 1500
                          
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
        
        
      QRY = "SELECT hvplaca, hvconact, hvnumint, hvestado " _
            & "FROM ophojveh  " _
           & "WHERE hvplaca LIKE '" & Txtcribus.text & "%'"
         
      'qry = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE henombre LIKE '" & Txtcribus.Text & "%'"
              
      Set Tbl1 = Dbs.Execute(QRY)
           
      Do While Tbl1.EOF = False
         
         If IsNull(Tbl1(2)) Then
            snumint = 0
         Else
            snumint = Tbl1(2)
         End If
      
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + str(snumint) + Chr(9) + Tbl1(3)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
      
   If forma = "3" Then    'Detecta cual es la forma rutas
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      If forma1 = "14" Then
         s$ = "Codigo   |Ruta   |Estado"
      Else
        s$ = "Codigo   |Ruta   |Estado"
      End If
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
                          
      For i = 0 To Grdbuscar.Cols - 1
        Grdbuscar.Col = i
        Grdbuscar.CellFontBold = 1
      Next i
       
      If forma1 = "c3" Or forma1 = "c4" Or forma1 = "14" Or forma1 = "19" Or forma1 = "20" Or forma1 = "34" Or forma1 = "36" Then
        QRY = "SELECT rucodigo, rudestino, ruestado " _
              & "FROM opruta, OPRUTHOR   " _
             & "WHERE rudestino LIKE '" & Txtcribus.text & "%' " _
               & "AND rhcodrut = rucodigo " _
               & "AND rhestado = 'Activa' AND ruestado = 'Activa'  "
        
        If forma1 = "c3" Then
            QRY = QRY & "AND rhcodemp = '" & G_Codigo_Empresa & "' "
        ElseIf forma1 = "c4" Then
            QRY = QRY & "AND rhcodemp = '" & Crehoemp.Txtcodemp.text & "' "
        ElseIf forma1 = "14" Then
            QRY = QRY & "AND rhcodemp = '" & convempr.Txtcodemp.text & "' "
        ElseIf forma1 = "19" Then
            QRY = QRY & "AND rhcodemp = '" & G_Codigo_Empresa & "' "
        ElseIf forma1 = "20" Then
            QRY = QRY & "AND rhcodemp = '" & G_Codigo_Empresa & "' "
        ElseIf forma1 = "34" Then
            QRY = QRY & "AND rhcodemp = '" & G_Codigo_Empresa & "' "
        End If
        QRY = QRY & "GROUP BY rucodigo, rudestino, ruestado ORDER BY rudestino"
      
      Else
         QRY = "SELECT rucodigo, rudestino, ruestado " _
               & "FROM opruta  " _
              & "WHERE rudestino LIKE '" & Txtcribus.text & "%'" _
           & "ORDER BY rudestino"
      End If
      Set Tbl1 = Dbs.Execute(QRY)
      
      Do While Tbl1.EOF = False
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
      
   If forma = "4" Then           'Tipos de multa de las tablas de catalogo
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      s$ = "Codigo   |Descripcion      |Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                             
      QRY = "SELECT * " _
            & "FROM geparsis " _
           & "WHERE pscod LIKE '" & Txtcribus.text & "%' " _
             & "AND psmod = 'CAT' "
                        
      Set Tbl1 = Dbs.Execute(QRY)
            
      Do While Tbl1.EOF = False
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "7" Then    'Detecta cual es la forma Multas
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      s$ = "Num. Multa  |Num. Placa  |Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
                          
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
      
      respuesta = MsgBox("Dese Buscar Por Numero De Placa", 32 + vbYesNo, "Busqueda")
             
      If respuesta = vbYes Then
         QRY = "SELECT munumero, munumpla,  muestado " _
               & "FROM opmultas  " _
              & "WHERE munumpla LIKE '" & Txtcribus.text & "%' " _
                & "AND muestado = 'Pendiente' "
      Else
         QRY = "SELECT munumero, munumpla,  muestado " _
               & "FROM opmultas  " _
              & "WHERE munumero LIKE '" & Txtcribus.text & "%' " _
                & "AND muestado = 'Pendiente' "
      
      End If
      
      Set Tbl1 = Dbs.Execute(QRY)
     
      Do While Tbl1.EOF = False
         entry = str(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "10" Then
      Grdbuscar.Cols = 4
      Grdbuscar.SetFocus
      s$ = "Num. Factura   |Empresa      |Fecha   |Imprimio "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
      Grdbuscar.ColWidth(3) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                
      QRY = "SELECT cbnumfac, henombre, cbfecha, cbimprimio " _
            & "FROM opconblo, ophojemp " _
           & "WHERE cbnumfac LIKE '" & Txtcribus.text & "%' " _
             & "AND cbempresa = hecodigo "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Tbl1.EOF = False
         entry = CStr(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + CStr(Tbl1(2)) + Chr(9) + Tbl1(3)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "11" Then
      Grdbuscar.Cols = 4
      Grdbuscar.SetFocus
      s$ = "Consecutivo   |Placa    |Empresa   |Fecha "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1200
      Grdbuscar.ColWidth(1) = 1000
      Grdbuscar.ColWidth(2) = 2500
      Grdbuscar.ColWidth(3) = 1100
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                
      QRY = "SELECT alconsecutivo, alplaca, henombre, alfecha " _
            & "FROM opalcolemia, ophojemp " _
           & "WHERE alplaca LIKE '" & Txtcribus.text & "%' " _
             & "AND alempresa = hecodigo " _
        & "ORDER BY alconsecutivo, alplaca, henombre, alfecha "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Tbl1.EOF = False
         entry = CStr(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(9) + str(Tbl1(3))
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "12" Then
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      s$ = "Caja   |Consecutivo    |Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1500
      Grdbuscar.ColWidth(1) = 1500
      Grdbuscar.ColWidth(2) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                
      QRY = "SELECT canomcaj, caconcaj, caestado " _
            & "FROM cocaja " _
           & "WHERE canomcaj LIKE '" & Txtcribus.text & "%' " _
        & "ORDER BY canomcaj "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Tbl1.EOF = False
         entry = Tbl1(0) + Chr(9) + str(Tbl1(1)) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
   
   If forma = "13" Then
      Grdbuscar.Cols = 4
      Grdbuscar.SetFocus
      s$ = "Codigo|Banco   |Cuenta    |Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1500
      Grdbuscar.ColWidth(1) = 1500
      Grdbuscar.ColWidth(2) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                
      QRY = "SELECT bacodigo, banombre, bacuenta, baestado " _
            & "FROM cobancos " _
           & "WHERE banombre LIKE '" & Txtcribus.text & "%' " _
        & "ORDER BY banombre "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Tbl1.EOF = False
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(0) + Tbl1(3)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   ' convenios de mepresas
   If forma = "14" Then
      Grdbuscar.Cols = 5
      Grdbuscar.SetFocus
      s$ = "Número|Empresa Contratante|Empresa Contratada|Resolución|Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 2000
      Grdbuscar.ColWidth(2) = 2000
      Grdbuscar.ColWidth(3) = 1000
      Grdbuscar.ColWidth(4) = 1000
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
      
                
      QRY = "SELECT cenumero, henombre, ceempcon, ceresolu, ceestado " _
            & "FROM opconemp, ophojemp " _
           & "WHERE henombre LIKE '%" & Txtcribus.text & "%' " _
             & "AND hecodigo = cecodemp " _
             & "AND ceestado = 'A' " _
        & "ORDER BY ceestado, henombre "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Not Tbl1.EOF
         
         'qry = "SELECT henombre " _
               & "FROM ophojempcon  " _
              & "WHERE hecodigo = '" & Tbl1!ceempcon & "'"
          
         QRY = "SELECT henombre "
         If Len(Tbl1!ceempcon) = 4 Then
            QRY = QRY & "FROM ophojempcon  "
         Else
            QRY = QRY & "FROM ophojemp  "
         End If
         QRY = QRY & "WHERE hecodigo = '" & Tbl1!ceempcon & "'"
         
         
         Set Tbl2 = Dbs.Execute(QRY)
     
         If Not Tbl2.EOF Then
            'empcon = Tbl2!henombre
            empcon = Tbl2!henombre
         End If
         Tbl2.Close
         
         
         
         entry = CStr(Tbl1!cenumero) + Chr(9) + Tbl1!henombre + Chr(9) + empcon + Chr(9) + _
                Tbl1!ceresolu + Chr(9) + Tbl1!ceestado               '+ Chr(9) + _

               ' tbl1(4)
                
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "15" Then                'Los tipos de estados se encuentran en las tablas de catalogos
      Grdbuscar.Cols = 3
      Grdbuscar.SetFocus
      s$ = "Codigo   |Descripcion      |Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
                
      QRY = "SELECT dscoddet,dsdes,dsest " _
            & "FROM gesuptip, gedetsuptip " _
           & "WHERE dscoddet LIKE '" & Txtcribus.text & "%' " _
             & "AND dscodtip = stcodtip " _
             & "AND stdes = '" & Tabla_Catalogo & "' "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Tbl1.EOF = False
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
   If forma = "16" Then    'Detecta cual es la forma Vehiculo
      Grdbuscar.Cols = 5
      Grdbuscar.SetFocus
      s$ = "Codigo   |Nombre  |NIT  |Direccion  |Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 2000
      Grdbuscar.ColWidth(2) = 1000
      Grdbuscar.ColWidth(3) = 2500
      Grdbuscar.ColWidth(4) = 1000
                          
      For i = 0 To Grdbuscar.Cols - 1
        Grdbuscar.Col = i
        Grdbuscar.CellFontBold = 1
      Next i
       
     'qry = "SELECT hvplaca, hvconact, hvnumint, hvestado " _
           & "FROM ophojveh  " _
          & "WHERE hvplaca LIKE '" & Txtcribus.Text & "%'"
        
      QRY = "SELECT * " _
            & "FROM ophojemp " _
           & "WHERE henombre LIKE '%" & Txtcribus.text & "%' "
      
      If forma1 = "2" Or forma1 = "47" Then
        QRY = QRY & "AND heestado = 'Activa' "
      End If
                   
      QRY = QRY & " ORDER BY HENOMBRE"
                   
      Set Tbl1 = Dbs.Execute(QRY)
          
      Do While Tbl1.EOF = False
         entry = Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + Tbl1!hedirecc + Chr(9) + Tbl1!heestado
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
      
      If forma1 = "2" Or forma1 = "3" Or forma1 = "4" Or forma1 = "5" Or _
         forma1 = "6" Or forma1 = "48" Or forma1 = "60" Then
         QRY = "SELECT * " _
               & "FROM ophojempcon " _
              & "WHERE henombre LIKE '%" & Txtcribus.text & "%' ORDER BY HENOMBRE"
               
         Set Tbl1 = Dbs.Execute(QRY)
          
         Do While Tbl1.EOF = False
            entry = Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + Tbl1!hedirecc + Chr(9) + Tbl1!heestado
            Grdbuscar.AddItem entry            ' Agrega la entrada.
            Tbl1.MoveNext ' Pasa al campo siguiente.
            bandera = 1   ' Mensaje de busqueda
         Loop
         Tbl1.Close
      End If
   End If
  
   'entrega de fichas
   If forma = "17" Then
      Grdbuscar.Cols = 4
      Grdbuscar.SetFocus
      s$ = "Número|Placa|Empresa|Fecha"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 1500
      Grdbuscar.ColWidth(2) = 1500
      Grdbuscar.ColWidth(3) = 1500
       
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
      
                
      QRY = "SELECT efnumero, efplaca, henombre, effecha " _
            & "FROM opentfic, ophojemp " _
           & "WHERE efplaca LIKE '" & Txtcribus.text & "%' " _
             & "AND hecodigo = efcodemp " _
        & "ORDER BY efplaca "
                                 
      Set Tbl1 = Dbs.Execute(QRY)
             
      Do While Not Tbl1.EOF
         
         
         entry = CStr(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2) + Chr(9) + CStr(Tbl1(3))
                
                
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext                      ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
   End If
  
   If forma = "18" Then    'Consulta Conduces
       Btnescoger.Enabled = False
       Grdbuscar.Cols = 4
       Grdbuscar.SetFocus
       s$ = "Número   |Tasa de Uso|Fecha Sistema|Estado"
       Grdbuscar.FormatString = s$
       Grdbuscar.Row = 0
       Grdbuscar.Col = 0
       Grdbuscar.ColWidth(0) = 1000
       Grdbuscar.ColWidth(1) = 1500
       Grdbuscar.ColWidth(2) = 1500
       Grdbuscar.ColWidth(3) = 1000
       
       
       For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
        
       QRY = "SELECT vanumero, vanumcon, vafecsis, vaestado " _
             & "FROM covalida  " _
            & "WHERE vanumcon LIKE '" & Txtcribus.text & "%'"
          
       Set Tbl1 = Dbs.Execute(QRY)
       
       
            
       Do While Tbl1.EOF = False
          entry = CStr(Tbl1(0)) + Chr(9) + CStr(Tbl1(1)) + Chr(9) + Format(Tbl1(2), "MM/dd/yyyy HH:mm") + Chr(9) + Tbl1(3)
          Grdbuscar.AddItem entry            ' Agrega la entrada.
          Tbl1.MoveNext ' Pasa al campo siguiente.
          bandera = 1   ' Mensaje de busqueda
       Loop
       Tbl1.Close
   End If
   If forma = "19" Then    'Detecta cual es la forma Vehiculo
      Grdbuscar.Cols = 5
      Grdbuscar.SetFocus
      s$ = "Codigo   |Nombre  |NIT  |Direccion  |Estado"
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1000
      Grdbuscar.ColWidth(1) = 2000
      Grdbuscar.ColWidth(2) = 1000
      Grdbuscar.ColWidth(3) = 2500
      Grdbuscar.ColWidth(4) = 1000
                          
      For i = 0 To Grdbuscar.Cols - 1
        Grdbuscar.Col = i
        Grdbuscar.CellFontBold = 1
      Next i
       
     'qry = "SELECT hvplaca, hvconact, hvnumint, hvestado " _
           & "FROM ophojveh  " _
          & "WHERE hvplaca LIKE '" & Txtcribus.Text & "%'"
        
     QRY = "SELECT * " _
              & "FROM ophojempcon " _
             & "WHERE henombre LIKE '" & Txtcribus.text & "%' ORDER BY HENOMBRE"
             
     Set Tbl1 = Dbs.Execute(QRY)
    
           
      Do While Tbl1.EOF = False
         entry = Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + Tbl1!hedirecc + Chr(9) + Tbl1!heestado
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
      If forma1 = "1" Or forma1 = "5" Then
         QRY = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE henombre LIKE '" & Txtcribus.text & "%' "
         
         If forma1 = "2" Or forma1 = "47" Then
           QRY = QRY & "AND heestado = 'Activa' "
         End If
                      
         QRY = QRY & " ORDER BY HENOMBRE"
                      
         Set Tbl1 = Dbs.Execute(QRY)
             
         Do While Tbl1.EOF = False
            entry = Tbl1!hecodigo + Chr(9) + Tbl1!henombre + Chr(9) + Tbl1!henit + Chr(9) + Tbl1!hedirecc + Chr(9) + Tbl1!heestado
            Grdbuscar.AddItem entry            ' Agrega la entrada.
            Tbl1.MoveNext ' Pasa al campo siguiente.
            bandera = 1   ' Mensaje de busqueda
         Loop
         Tbl1.Close
      End If
   End If
   If forma = "20" Then                'Tipos de vehiculo de las tablas de catalogos
               
     QRY = "SELECT cocedula, conombres, coapellidos " _
           & "FROM coconductor " _
          & "WHERE conombres LIKE '" & Txtcribus.text & "%' "
                                
     
     Set Tbl1 = Dbs.Execute(QRY)
            
     Grdbuscar.Cols = 3
     Grdbuscar.SetFocus
     s$ = "Cédula   |Nombre      |Apellido "
     Grdbuscar.FormatString = s$
     Grdbuscar.Row = 0
     Grdbuscar.Col = 0
     Grdbuscar.ColWidth(0) = 1300
     Grdbuscar.ColWidth(1) = 1500
     Grdbuscar.ColWidth(2) = 1500
      
     For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
     Next i
            
     Do While Not Tbl1.EOF
        entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
        Grdbuscar.AddItem entry            ' Agrega la entrada.
        Tbl1.MoveNext                      ' Pasa al campo siguiente.
        bandera = 1   ' Mensaje de busqueda
     Loop
     Tbl1.Close
   End If
   
   If forma = "21" Then
      Dib_Grilla_Usuarios
      QRY = "SELECT uscodusu, usnom, usest " _
            & "FROM geususis " _
           & "WHERE usnom LIKE '" & Txtcribus.text & "%'"
            
      Set Tbl1 = Dbs.Execute(QRY)
     
      Do While Not Tbl1.EOF
         entry = Tbl1(0) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry
         Tbl1.MoveNext
        
         bandera = 1
      Loop
      Tbl1.Close
   End If
   
   If forma = "22" Then
      Dib_Grilla_Cupos
      QRY = "SELECT cpnumcup, henombre , cpestado, cptipo " _
            & "FROM cocupos INNER JOIN ophojemp ON cpempresa = hecodigo " _
           & "WHERE henombre LIKE '" & Txtcribus.text & "%' " _
           & "ORDER BY cpnumcup "
              
      Set Tbl1 = Dbs.Execute(QRY)
        
      Do While Not Tbl1.EOF
         sNombre = Tbl1!henombre
         
         If Tbl1!cptipo = "E" Then
            sTipo = "Empresa"
         Else
            sTipo = "Vehículo"
         End If
            
      
         entry = CStr(Tbl1!cpnumcup) + Chr(9) + sNombre + Chr(9) + sTipo + Chr(9) + Tbl1!cpestado
         Grdbuscar.AddItem entry
         Tbl1.MoveNext
         bandera = 1
      Loop
      Tbl1.Close
   End If
   
   If forma = "23" Then
      Dib_Grilla_Factura
      QRY = "SELECT cfnumfac, henombre, cfestado " _
            & "FROM cocupos, ophojemp, cofaccup " _
           & "WHERE henombre LIKE '" & Txtcribus.text & "%' " _
             & "AND cpnumcup = cfnumcup " _
             & "AND cpempresa = hecodigo "
              
      Set Tbl1 = Dbs.Execute(QRY)
        
      Do While Not Tbl1.EOF
         entry = CStr(Tbl1(0)) + Chr(9) + Tbl1(1) + Chr(9) + Tbl1(2)
         Grdbuscar.AddItem entry
         Tbl1.MoveNext
         bandera = 1
      Loop
      Tbl1.Close
   End If
   
   If forma = "24" Then                'persona comparedo electronico
               
     QRY = "SELECT cpcedula, cpnombres, cpapellidos " _
           & "FROM opcompapersona " _
          & "WHERE cpnombres LIKE '" & Txtcribus.text & "%' "
                                
     
     Set Tbl1 = Dbs.Execute(QRY)
            
     Grdbuscar.Cols = 3
     Grdbuscar.SetFocus
     s$ = "Cédula   |Nombre      |Apellido "
     Grdbuscar.FormatString = s$
     Grdbuscar.Row = 0
     Grdbuscar.Col = 0
     Grdbuscar.ColWidth(0) = 1300
     Grdbuscar.ColWidth(1) = 1500
     Grdbuscar.ColWidth(2) = 1500
      
     For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
     Next i
            
     Do While Not Tbl1.EOF
        entry = Tbl1!cpcedula + Chr(9) + Tbl1!cpnombres + Chr(9) + Tbl1!cpapellidos
        Grdbuscar.AddItem entry            ' Agrega la entrada.
        Tbl1.MoveNext                      ' Pasa al campo siguiente.
        bandera = 1   ' Mensaje de busqueda
     Loop
     Tbl1.Close
   End If
   
   If forma = "25" Then   'Concepto TAG - Sebastián Rondón 31 Octubre - 2022
   
        Grdbuscar.Cols = 6
      Grdbuscar.SetFocus
      s$ = "Consecutivo  |Num. Placa  |# TAG |Valor |Fecha Registro |Estado "
      Grdbuscar.FormatString = s$
      Grdbuscar.Row = 0
      Grdbuscar.Col = 0
      Grdbuscar.ColWidth(0) = 1300
      Grdbuscar.ColWidth(1) = 2500
      Grdbuscar.ColWidth(2) = 2500
      Grdbuscar.ColWidth(3) = 1500
      Grdbuscar.ColWidth(4) = 2500
      Grdbuscar.ColWidth(5) = 1500
                          
      For i = 0 To Grdbuscar.Cols - 1
         Grdbuscar.Col = i
         Grdbuscar.CellFontBold = 1
      Next i
      
      respuesta = MsgBox("Desea Buscar Por Numero De Placa", 32 + vbYesNo, "Busqueda")
      If respuesta = vbYes Then
         QRY = "SELECT tanumero, taplaca, tacodtag, tavalor, tafecreg, taestado " _
               & "FROM optag  " _
              & "WHERE taplaca LIKE '" & Txtcribus.text & "%' " _
              & "AND tafecpag IS NULL "
      Else
         QRY = "SELECT tanumero, taplaca, tacodtag, tavalor, tafecreg, taestado " _
               & "FROM optag  " _
              & "WHERE tacodtag LIKE '" & Txtcribus.text & "%' " _
              & "AND tafecpag IS NULL "
      End If
      
      Set Tbl1 = Dbs.Execute(QRY)
     
      Do While Tbl1.EOF = False
         entry = str(Tbl1!tanumero) + Chr(9) + Tbl1!taplaca + Chr(9) + Tbl1!tacodtag + Chr(9) + str(Tbl1!tavalor) + Chr(9) + str(Tbl1!tafecreg) + Chr(9) + Tbl1!taestado
         Grdbuscar.AddItem entry            ' Agrega la entrada.
         Tbl1.MoveNext ' Pasa al campo siguiente.
         bandera = 1   ' Mensaje de busqueda
      Loop
      Tbl1.Close
        
   End If
   
   If bandera = 0 Then
      MsgBox "No Se Encontro Ningun Registro Activo", 64, "Busqueda"
      Unload Me
   End If
   Exit Sub
      
Err_Btnbuscar_Click:
      MsgBox Err.Description, 16, "Error " & Err.Number

      Exit Sub
      Resume
End Sub

Private Sub Btnescoger_Click()
   If Grdbuscar.text <> "" And Grdbuscar.rows > 1 Then
      Y = Grdbuscar.Row
      Grdbuscar.Col = 0
     
      If forma = "16" Then
         If forma1 = "1" Then
           Empresas!Txtcodigo.Enabled = True
           Empresas!Txtcodigo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
           Vehiculo!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
           Conempresas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
           Grdbuscar.Col = 2
           Conempresas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
           Grdbuscar.Col = 1
           Conempresas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "6" Then
           Grdbuscar.Col = 3
           Conempresas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "7" Then
           Creruthor!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "8" Then
           Rutempresa!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "9" Then
           Conrutas!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "10" Then
           Regmultas!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "11" Then
           Conmultas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "12" Then
           Conmulcri!Txtcodemp.Enabled = True
           Conmulcri!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "13" Then
           Conrutgen!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "14" Then
           Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
           'Conduces
         ElseIf forma1 = "c2" Then
           Capconduce!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "15" Then
           Conconcriterio!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "16" Then
           Concontotal!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "17" Then
           Conbloques!Txtempresa.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "18" Then
           Consulbloques!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "19" Then
           Conalcolemia!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "20" Then
           convenios!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "25" Then
           Conempto!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "26" Then
           convehdi!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "27" Then
           conpasaj!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "28" Then
           contasdi!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "29" Then
           cosalpas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "30" Then
           convempr!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "31" Then
           convempr!txtempcon.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "32" Then
           pantalla!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "33" Then
           conconemp!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "34" Then
           covensal!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "35" Then
           entficha!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "36" Then
           conempco!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "37" Then
           Grdbuscar.Col = 0
           coemruse!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "38" Then
           Grdbuscar.Col = 0
           Genplano!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "39" Then
           Grdbuscar.Col = 0
           conlogfi!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "40" Then
           Grdbuscar.Col = 0
           convalid!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "41" Then
           Grdbuscar.Col = 0
           cosanmul!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "42" Then
           Grdbuscar.Col = 0
           Crehoemp!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "43" Then
           Grdbuscar.Col = 0
           Conparqu!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "44" Then
           Grdbuscar.Col = 0
           condespar!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "45" Then
           Grdbuscar.Col = 0
           conreser!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "46" Then
           Grdbuscar.Col = 0
           operempr.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "47" Then
           Grdbuscar.Col = 0
           Vehiculo.txtempcon.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "48" Then
           conempco!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "49" Then
           cocontra!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "50" Then
           conconve!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "51" Then
           concensur!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "52" Then
           conres315!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "53" Then
           conductor!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "54" Then
           cocupos!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "55" Then
           FrUsuEmpresa!Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "56" Then
           RptRodamiento.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "57" Then
           Rpttasanorte.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "58" Then
           reimprimir.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "59" Then
           rptmovimientos.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "60" Then
           confactu.TxtCodEmpContratada.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "61" Then
           confactu.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "62" Then
           rptcomelec.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "63" Then
           RutAdicionales.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "64" Then
           camservicio.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "65" Then
           rptmovimientos.TxtCodEmpVeh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "66" Then
           cocupos.TxtCodEmpVeh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "67" Then
            Regtag.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "68" Then
            Contagcriterio.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "69" Then
            coemgaf.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "70" Then
            coalcopruebas.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "71" Then
            coconfec.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "72" Then
            corecasaldo.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "73" Then
            cocondicional.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "74" Then
            ConConductores.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "75" Then
            ConVehiculos.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "76" Then
            Coconconcep.Txtcodemp.text = Trim(Grdbuscar.text)
         End If
         
      End If
      
      If forma = "2" Then
         If forma1 = "1" Then
            Vehiculo!Txtnumpla.Enabled = True
            Vehiculo!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            Regmultas!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Conmultas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            Conmulcri!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
            Grdbuscar.Col = 2
            Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "6" Then
            Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "7" Then
            Conalcolemia!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "8" Then
            detvehic!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "9" Then
            covensal!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "10" Then
            entficha!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "11" Then
            conlogfi!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "12" Then
            Deseduca!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "13" Then
            convalid!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "14" Then
           cosanmul!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "15" Then
           Conparqu!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "16" Then
           condespar!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "17" Then
           conreser!Txtnumpla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "18" Then
           conaudpa!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "19" Then
           conres315!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "20" Then
           RptRodamiento!txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "21" Then
           Rpttasanorte.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "22" Then
           reimprimir.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "23" Then
           rptcomelec.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "24" Then
            Regtag.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "25" Then
            coalcopruebas.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "26" Then
            coentsalprot.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "27" Then
            cocondicional.txtPlaca.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "28" Then
            ConVehiculos.txtPlaca.text = Trim(Grdbuscar.text)
         End If
      End If
     
      If forma = "3" Then
         If forma1 = "1" Then
            Crerutas!Txtcodigo.Enabled = True
            Crerutas!Txtcodigo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            Creruthor!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Rutempresa!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            Conrutas!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
            Conrutas!Txtcodruta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "6" Then
            Conrutgen!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "7" Then
            Conrutgen!Txtcodruta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "8" Then
            Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "c3" Then 'conduce
            Capconduce!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "9" Then
            Conconcriterio!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "10" Then
            Conbloques!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "11" Then
            Consulbloques!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "12" Then
            convenios!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "13" Then
            concrirut!Txtcodruta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "14" Then
            convempr!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "15" Then
            conprecond!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "16" Then
            pantalla!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "17" Then
            Conempto!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "18" Then
            convehto!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "19" Then
            Regmultas!Txtdestino.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "20" Then
            Regmultas!Txtdespla.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "21" Then
            conempco!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "22" Then
            Grdbuscar.Col = 0
            coemruse!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "23" Then
            Grdbuscar.Col = 0
            Genplano!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "c4" Then
            Crehoemp.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "24" Then
            Grdbuscar.Col = 0
            conconve!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "25" Then
            Grdbuscar.Col = 0
            concensur!Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "26" Then
            Grdbuscar.Col = 0
            conres315!TxtcodrutDes.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "27" Then
            Grdbuscar.Col = 0
            conres315!TxtcodrutOri.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "28" Then
            Grdbuscar.Col = 0
            RptRodamiento.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "29" Then
            Grdbuscar.Col = 0
            Rpttasanorte.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "30" Then
            Grdbuscar.Col = 0
            RutAdicionales.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "31" Then
            Grdbuscar.Col = 0
            RutAdicionales.TxtRutAdicional.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "32" Then
            Grdbuscar.Col = 0
            camservicio.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "33" Then
            Grdbuscar.Col = 0
            coalcopruebas.txtRuta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "34" Then
            Grdbuscar.Col = 0
            coalcopruebas.txtRuta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "35" Then
            Grdbuscar.Col = 0
            Coconconcep.Txtcodrut.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "36" Then
            Grdbuscar.Col = 0
            Coconconcep.Txtcodrut.text = Trim(Grdbuscar.text)
         End If
      End If
     
     If forma = "4" And forma1 = "1" Then
        Regmultas!Txtcodinf.text = Trim(Grdbuscar.text)
     ElseIf forma = "4" And forma1 = "2" Then
        Conmultas!Txtvalbus.text = Trim(Grdbuscar.text)
     ElseIf forma = "4" And forma1 = "3" Then
        cosanmul!Txtcodinf.text = Trim(Grdbuscar.text)
     End If
     
     
          
     
     
     If forma = "7" And forma1 = "1" Then
        Regmultas!Txtnumero.text = Trim(Grdbuscar.text)
     ElseIf forma = "7" And forma1 = "2" Then
        Canmultas!Txtnumero.text = Trim(Grdbuscar.text)
     ElseIf forma = "7" And forma1 = "3" Then
        Anumultas!Txtnumero.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "10" And forma1 = "1" Then
        Conbloques!Txtnumfac.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "11" And forma1 = "1" Then
        Grdbuscar.Col = 1
        Prueba!txtPlaca.text = Trim(Grdbuscar.text)
     ElseIf forma = "11" And forma1 = "2" Then
        Prueba!Txtconsecutivo.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "12" And forma1 = "1" Then
        Grdbuscar.Col = 0
        Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "2" Then
        Grdbuscar.Col = 0
        Concietur!txtcaja.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "3" Then
        Grdbuscar.Col = 0
        Concontotal!txtcaja.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "4" Then
        Grdbuscar.Col = 0
        conconsi!Txtvalbus.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "5" Then
        Grdbuscar.Col = 0
        comulcan!Txtvalbus.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "6" Then
        Grdbuscar.Col = 0
        contasdi!Txtvalbus.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "7" Then
        Grdbuscar.Col = 0
        convehdi!txtcaja.text = Trim(Grdbuscar.text)
     ElseIf forma = "12" And forma1 = "8" Then
        Grdbuscar.Col = 0
        conpasaj!txtcaja.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "13" And forma1 = "1" Then
        Grdbuscar.Col = 0
        Cuaturno!txtbancos.text = Trim(Grdbuscar.text)
     ElseIf forma = "13" And forma1 = "2" Then
        Grdbuscar.Col = 0
        conconsi!Txtvalbus.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "14" And forma1 = "1" Then
        Grdbuscar.Col = 0
        convempr!Txtnumero.text = Trim(Grdbuscar.text)
        'Grdbuscar.Col = 2
        'convempr!txtempcon.Text = Trim(Grdbuscar.Text)
        'Grdbuscar.Col = 4
        'convempr!txtcodrut.Text = Trim(Grdbuscar.Text)
        'Grdbuscar.Col = 6
        'convempr!txtcodser.Text = Trim(Grdbuscar.Text)
     End If
     
      If forma = "15" Then
         If forma1 = "1" Then
            Regmultas!Txttipveh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            Vehiculo!Txttipveh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Creruthor!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "c2" Then  'conduce
            Capconduce!Txttipveh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            Conconcriterio!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
            Conbloques!Txttipveh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "6" Then
            Consulbloques!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "7" Then
            convenios!txtcodser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "9" Then
            concrirut!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "10" Then
            conprecond!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "11" Then
            convempr!txtcodser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "12" Then
            Conempto!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "13" Then
            cosalpas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "14" Then
            Grdbuscar.Col = 0
            pantalla!Txtmodulo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "15" Then
            Regmultas!Txtnomaux.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "16" Then
            Conmultas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "17" Then
            Conconcriterio!Txtoritra.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "18" Then
            Conempto!Txtoritra.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "19" Then
            Conmultas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "20" Then
            Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "21" Then
            Grdbuscar.Col = 0
            pantalla!TxtModAlt.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "22" Then
            Grdbuscar.Col = 0
            pantalla!TxtModAlt1.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "23" Then
            Grdbuscar.Col = 0
            convehto!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "24" Then
            Grdbuscar.Col = 0
            Regmultas!Txtsitio.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "25" Then
            Grdbuscar.Col = 0
            Regmultas!Txtauxalt.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "26" Then
            Grdbuscar.Col = 0
            Empresas!Txtconvenio.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "27" Then
            Grdbuscar.Col = 0
            Crerutas!txtcoddep.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "28" Then
            Grdbuscar.Col = 0
            conempco!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "29" Then
            Grdbuscar.Col = 0
            coemruse!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "30" Then
            Grdbuscar.Col = 0
            Vehiculo.Txtcodcom.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "31" Then
            Grdbuscar.Col = 0
            covalida.Txtcodval.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "32" Then
            Grdbuscar.Col = 0
            convalid.Txtcodval.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "33" Then
            Grdbuscar.Col = 0
            Platafor.Txttipveh.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "35" Then
            opconsur!txtClase.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "36" Then
            Platafor.TxtZona.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "37" Then
             Crehoemp.Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "38" Then
             condespar.TxtTipo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "39" Then
            Grdbuscar.Col = 0
            Crerutas!TxtOperador.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "40" Then
            operempr!TxtOperador.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "41" Then
            Creruthor!TxtOperador.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "42" Then
            Grdbuscar.Col = 0
            EmprConv!Txtconvenio.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "43" Then
            Grdbuscar.Col = 0
            conconve!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "44" Then
            Grdbuscar.Col = 0
            conconve!Txtoritra.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "45" Then
            Grdbuscar.Col = 0
            concensur!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "46" Then
            Grdbuscar.Col = 0
            conres315!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "47" Then
            Grdbuscar.Col = 0
            confpres.Txtmodulo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "50" Then
            Grdbuscar.Col = 0
            FrUsuEmpresa.Txtcodemp.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "51" Then
            Grdbuscar.Col = 0
            RptRodamiento.Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "52" Then
            Grdbuscar.Col = 0
            Rpttasanorte.Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "53" Then
            confactu!Txtoritra.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "54" Then
            rptcomelec!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "55" Then
            rptcomelec!TxtFalta.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "56" Then
            rptcomelec!TxtComete.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "57" Then
            rptcomelec!TxtLugar.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "58" Then
            rptcomelec!TxtArea.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "59" Then
            rptcomelec!TxtControl.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "60" Then
            camservicio!Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "61" Then
            Regtag.TxtConcep.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "62" Then
            Contagcriterio.txtCodCon = Trim(Grdbuscar.text)
         ElseIf forma1 = "63" Then
            ConVehiculos.txtServicio.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "64" Then
            ConVehiculos.txtCombustible.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "65" Then
            ConVehiculos.txtClase.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "66" Then
            Grdbuscar.Col = 0
            Coconconcep.Txttipser.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "67" Then
            Grdbuscar.Col = 0
            Coconconcep.Txtoritra.text = Trim(Grdbuscar.text)
         End If
     End If
     
     If forma = "17" And forma1 = "1" Then
        Grdbuscar.Col = 0
        entficha!Txtnumero.text = Trim(Grdbuscar.text)
     End If
     If forma = "18" And forma1 = "1" Then
        Grdbuscar.Col = 1
        covalida!Txtnumcon.text = Trim(Grdbuscar.text)
     End If
     
     If forma = "19" Then
         If forma1 = "1" Then
            convempr!txtempcon.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            EmprConv.Txtcodigo.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Conempresas!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            cocontra!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
            conconve!Txtcodemp.text = Trim(Grdbuscar.text)
         End If
         
     End If
      If forma = "20" Then
         If forma1 = "1" Then
            Conconduces!Txtvalbus.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            conductor!txtCedula.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Regmultas!txtCedula.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            frmborrarhuella!txtCedula.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "5" Then
            Rpttasanorte.txtCedula.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "6" Then

         ElseIf forma1 = "7" Then
            'rptmovimientos.TxtCedula.Text = Trim(Grdbuscar.Text)
         ElseIf forma1 = "8" Then
            coemgaf.txtCedula.text = Trim(Grdbuscar.text)
        ElseIf forma1 = "9" Then
            coalcopruebas.txtCedula.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "10" Then
            ConConductores.txtCedula.text = Trim(Grdbuscar.text)
         End If
      End If
      If forma = "21" Then
         If forma1 = "1" Then
            Conmulcri.TxtFuncionario.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            FrUsuEmpresa.txtUserName.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "3" Then
            Rpttasanorte.TxtFuncSalida.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "4" Then
            Rpttasanorte.TxtFuncVentaTasa.text = Trim(Grdbuscar.text)
         End If
      End If
      
      If forma = "22" Then
         If forma1 = "1" Then
            cocupos.Txtnumcup.text = Trim(Grdbuscar.text)
         ElseIf forma1 = "2" Then
            Anucupos.Txtnumcup.text = Trim(Grdbuscar.text)
         
         ElseIf forma1 = "3" Then
            cofaccup.Txtnumcup.text = Trim(Grdbuscar.text)
         End If
      End If
      
      If forma = "23" Then
         If forma1 = "1" Then
            anufactu.Txtnumfac.text = Trim(Grdbuscar.text)
         End If
      End If
      If forma = "24" Then
         If forma1 = "1" Then
            rptcomelec.txtCedula.text = Trim(Grdbuscar.text)
         End If
      End If
      If forma = "25" Then
        If forma1 = "1" Then
            Regtag.Txtnumero.text = Trim(Grdbuscar.text)
        End If
        
      End If
     If forma = "c1" And forma1 = "c1" Then
       Btnescoger.Enabled = True
     End If
     Unload Me
  End If
End Sub

Private Sub Btnlimpiar_Click()
  Grdbuscar.rows = 1
  Grdbuscar.Cols = 2
  Grdbuscar.Clear
  Txtcribus.text = ""
  Txtcribus.SetFocus
End Sub

Private Sub Btnsalir_Click()
  Unload Me
End Sub

Private Sub GrdBuscar_DblClick()
  Btnescoger_Click
End Sub

Private Sub Txtcribus_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub
