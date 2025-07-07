VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form reimprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimprimir Tasa de Uso"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid GrdImprimir 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valores de Búsqueda:"
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
      TabIndex        =   11
      Top             =   1080
      Width           =   5895
      Begin VB.TextBox Txtcodemp 
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   2
         ToolTipText     =   "Código de la empresa (Precione doble click para ayuda)"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtPlaca 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Placa del vehículo (Precione doble click para ayuda)"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cód. Empresa:"
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
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Lblcodemp 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label LblPlaca 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   600
         Width           =   3135
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
      Left            =   6240
      Picture         =   "reimprimir.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar informacion en la base de datos"
      Top             =   1800
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
      Left            =   6240
      Picture         =   "reimprimir.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir de la forma actual"
      Top             =   4320
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
      Left            =   6240
      Picture         =   "reimprimir.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3480
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
      Left            =   6240
      Picture         =   "reimprimir.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir documento soporte"
      Top             =   2640
      Width           =   855
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
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5895
      Begin MSMask.MaskEdBox Mskfecini 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Fecha en la que empieze a buscar"
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
         Left            =   4200
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
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   1215
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
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   135
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   5160
   End
End
Attribute VB_Name = "reimprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim L_Busqueda As String

Private Function DIB_GRILLA()
     GrdImprimir.rows = 1
     GrdImprimir.Cols = 20
     GrdImprimir.Clear
     
     s$ = "Sel|Núm. Tasa |Caja|Placa|Empresa|Destino|Fecha|Hora|Núm. Int.|Tipo Servicio|O/T|Cedula 1|Conductor 1 |Cedula 2 |Conductor 2 |" _
        & "Turno|Vr Tasa|Vr Alcoholimetría|Funcionario|Cod Ruta "
     GrdImprimir.FormatString = s$
     GrdImprimir.Row = 0
     GrdImprimir.Col = 0
     GrdImprimir.ColWidth(0) = 500
     GrdImprimir.ColWidth(1) = 1300
     GrdImprimir.ColWidth(2) = 900
     GrdImprimir.ColWidth(3) = 1000
     GrdImprimir.ColWidth(4) = 2500
     GrdImprimir.ColWidth(5) = 2500
     GrdImprimir.ColWidth(6) = 1200
     GrdImprimir.ColWidth(7) = 800
     GrdImprimir.ColWidth(8) = 900
     GrdImprimir.ColWidth(9) = 1400
     GrdImprimir.ColWidth(10) = 900
     GrdImprimir.ColWidth(11) = 1200
     GrdImprimir.ColWidth(12) = 1800
     GrdImprimir.ColWidth(13) = 1200
     GrdImprimir.ColWidth(14) = 1800
     GrdImprimir.ColWidth(15) = 800
     GrdImprimir.ColWidth(16) = 1000
     GrdImprimir.ColWidth(17) = 1600
     GrdImprimir.ColWidth(18) = 1200
     GrdImprimir.ColWidth(19) = 0
     
     
     For i = 0 To GrdImprimir.Cols - 1
         GrdImprimir.Col = i
         GrdImprimir.CellFontBold = 1
         GrdImprimir.ColAlignment(i) = 0
     Next i
End Function


Private Sub Btnbuscar_Click()
   On Error GoTo Errores
   
   sBandera = "0"
   Btnimprimir.Enabled = False
   ProBar.Value = ProBar.Min
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
   
   If Trim(txtCodEmp.text) = "" And Trim(txtPlaca.text) = "" Then
      MsgBox "Debe de consultar por algun criterio de busqueda", 48, "Advertencia"
      txtCodEmp.SetFocus
      
      Exit Sub
   End If
   
   
   DIB_GRILLA
     
   sFecha1 = Mskfecini.text
   sFecha2 = Mskfecfin.text
   L_Busqueda = ""
   QRY = "SELECT conumero, conomcaj, coplaca, " _
               & "(CASE comodo " _
                  & "WHEN 'P' THEN  coconduc.cocodemp || '-' || henombre || '   T ' || coterminal || '     CUPOS' " _
                  & "ELSE coconduc.cocodemp || '-' || henombre || '   T ' || coterminal || '     EFECTIVO' " _
               & "END) as henombre,  cocodrut || '-' || rudestino as rudestino,  TO_CHAR(cofecsal,'MON/DD/YYYY') AS cofecsal, TO_CHAR(cohorsal,'HH24:mi') AS cohorsal, conumint, dsdes, " _
               & "(CASE cooritra " _
                  & "WHEN 'O' THEN 'Origen' " _
                  & "ELSE 'Transito' " _
                  & "END) AS cooritra," _
               & "(CASE " _
                  & "WHEN coconduc.cocedula IS NULL  THEN '' " _
                  & "WHEN coconduc.cocedula IS NOT NULL THEN  coconduc.cocedula " _
               & "END) as cocedula, " _
               & "(CASE " _
                  & "WHEN coconduc.cocedula IS NULL  THEN '' " _
                  & "WHEN coconduc.cocedula IS NOT NULL THEN (coconductor1.conombres || ' ' || coconductor1.coapellidos) " _
                  & "END) as nombreconductor1, "
              
   QRY = QRY & "(CASE " _
                  & "WHEN coconduc.cocedula1 IS NULL  THEN '' " _
                  & "WHEN coconduc.cocedula1 IS NOT NULL THEN  cocedula1 " _
               & "END) as cocedula1, " _
               & "(CASE " _
                  & "WHEN coconduc.cocedula1 IS NULL  THEN '' " _
                  & "WHEN coconduc.cocedula1 IS NOT NULL THEN (coconductor2.conombres || ' ' || coconductor2.coapellidos) " _
                  & "END) as nombreconductor2, " _
               & "coturno, covalor, covalalc, lower(cofuncio) AS cofuncio, cocodrut " _
         & "FROM coconduc LEFT OUTER JOIN coconductor coconductor1  ON  coconduc.cocedula = coconductor1.cocedula " _
                       & "LEFT OUTER JOIN coconductor coconductor2  ON  coconduc.cocedula1 = coconductor2.cocedula, " _
              & "opruta, ophojemp, gesuptip, gedetsuptip  " _
        & "WHERE cofecsal BETWEEN TO_DATE('" & sFecha1 & "','MM/DD/YYYY') " _
          & "AND TO_DATE('" & sFecha2 & "','MM/DD/YYYY') "
                       
   '************************ CUANDO SE BUSCAR POR CODIGO DE EMPRESA ************************************
   If Trim(txtCodEmp.text) <> "" Then
      QRY = QRY & "AND coconduc.cocodemp = '" & txtCodEmp.text & "' "
      L_Busqueda = txtCodEmp.text
   End If
   
   '************************ CUANDO SE BUSCA POR CODIGO DE RUTA ****************************************
   If Trim(txtPlaca.text) <> "" Then
      QRY = QRY & "AND coplaca = '" & txtPlaca.text & "' "
      
      If L_Busqueda = "" Then
         L_Busqueda = txtPlaca.text
      Else
         L_Busqueda = L_Busqueda & "-" & txtPlaca.text
      End If
   End If
   
   QRY = QRY & "AND coconduc.cocodrut = rucodigo " _
             & "AND coconduc.cocodemp = hecodigo " _
             & "AND dscoddet = coclase " _
             & "AND dscodtip = stcodtip " _
             & "AND stdes = 'TIPOS DE VEHICULO' " _
        & "ORDER BY cofecsal, cohorsal ASC "
   
   
   Set Tbl1 = Dbs.Execute(QRY)
   
   If Not Tbl1.EOF Then
      Tbl1.MoveLast
      Tbl1.MoveFirst
      ProBar.Max = Tbl1.RecordCount + 2
   End If
   sSeleccion = ""
   reimprimir.Enabled = False
   Do While Not Tbl1.EOF
      sCedula2 = ""
      sNombreConductor2 = ""
      If Not IsNull(Tbl1!cocedula1) Then
         sCedula2 = Tbl1!cocedula1
      End If
      
      If Not IsNull(Tbl1!nombreconductor2) Then
         sNombreConductor2 = Tbl1!nombreconductor2
      End If
      
      entry = sSeleccion + Chr(9) + CStr(Tbl1!conumero) + Chr(9) + Tbl1!conomcaj + Chr(9) + Tbl1!coplaca + Chr(9) + Tbl1!henombre + Chr(9) + _
               Tbl1!rudestino + Chr(9) + Tbl1!cofecsal + Chr(9) + Tbl1!cohorsal + Chr(9) + CStr(Tbl1!conumint) + Chr(9) + Tbl1!dsdes + Chr(9) + _
               Tbl1!Cooritra + Chr(9) + Tbl1!cocedula + Chr(9) + Tbl1!nombreconductor1 + Chr(9) + sCedula2 + Chr(9) + sNombreConductor2 + Chr(9) + _
               CStr(Tbl1!coturno) + Chr(9) + CStr(Tbl1!covalor) + Chr(9) + CStr(Tbl1!covalalc) + Chr(9) + Tbl1!cofuncio + Chr(9) + Tbl1!cocodrut
               
      GrdImprimir.AddItem entry
      ProBar.Value = ProBar.Value + 1
      Tbl1.MoveNext
      sBandera = "1"
   Loop
   Tbl1.Close
   reimprimir.Enabled = True
   If sBandera = "0" Then
      MsgBox "No se encontro ningun registro activo", 48, "Advertencia"
      Exit Sub
   End If
   Btnimprimir.Enabled = True
   ProBar.Value = ProBar.Max
   
   Exit Sub
Errores:
   reimprimir.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub

Private Sub Btnimprimir_Click()
   On Error GoTo Errores
   
   seleccion = "0"
   For i = 1 To GrdImprimir.rows - 1
      GrdImprimir.Row = i
      GrdImprimir.Col = 0
      
      If GrdImprimir.text = "->" Then
         seleccion = "1"
         Exit For
      End If
   Next i
       
   If seleccion = "0" Then
      MsgBox "No se ha seleccionado ningún registro", 48, "Advertencia"
      GrdImprimir.Row = 1
      GrdImprimir.Col = 0
      GrdImprimir.SetFocus
      Exit Sub
   End If
   
   G_Clase_Autorizacion = "REIMPTASALOTES"
   Relogin.Show 1
   If G_Aut_Vender <> "SI" Then
      Exit Sub
   End If
   
   
   ProBar.Value = ProBar.Min
   ProBar.Max = GrdImprimir.rows + 2
   reimprimir.Enabled = False
   sTasas = ""
   sBandera = 0
   For j = 1 To GrdImprimir.rows - 1
      GrdImprimir.Row = j
      GrdImprimir.Col = 0
      If GrdImprimir.TextMatrix(j, 0) = "->" Then
                      
         hora = GrdImprimir.TextMatrix(j, 7)
         presenprueba = "Si"
         If CDbl(GrdImprimir.TextMatrix(j, 17)) = 0 Then
            presenprueba = "No"
         End If
           
         presenprueba = presenprueba & "-" & GrdImprimir.TextMatrix(j, 10)
         
         sNombreEmpresa = GrdImprimir.TextMatrix(j, 4)
         
         If Len(GrdImprimir.TextMatrix(j, 1)) = 1 Then
            LCodBarras = "00" & GrdImprimir.TextMatrix(j, 1)
         ElseIf Len(GrdImprimir.TextMatrix(j, 1)) = 2 Then
            LCodBarras = "0" & GrdImprimir.TextMatrix(j, 1)
         Else
            LCodBarras = GrdImprimir.TextMatrix(j, 1)
         End If
         sNumeroBarras = "*" & LCodBarras & "*"
         s2Conductores = GrdImprimir.TextMatrix(j, 11) & "-" & GrdImprimir.TextMatrix(j, 13) ' "DESTINO NECESITA DOS CONDUCTORES"
         sPermanencia1 = "-"
         sPermanencia2 = "-"
         sPermanencia3 = "-"
      
         cero = "0"
         cedula1 = GrdImprimir.TextMatrix(j, 11)
         For i = 1 To 10 - Len(GrdImprimir.TextMatrix(j, 11))
            cedula1 = cero & cedula1
         Next i
         
         cedula2 = GrdImprimir.TextMatrix(j, 13)
         For i = 1 To 10 - Len(GrdImprimir.TextMatrix(j, 13))
            cedula2 = cero & cedula2
         Next i
         
         numcon = GrdImprimir.TextMatrix(j, 1)
         For i = 1 To 10 - Len(GrdImprimir.TextMatrix(j, 1))
            numcon = cero & numcon
         Next i
         
         sCodigoDestino = GrdImprimir.TextMatrix(j, 19)
         For i = 1 To 4 - Len(GrdImprimir.TextMatrix(j, 19))
            sCodigoDestino = cero & sCodigoDestino
         Next i
         
         LNumConBarra = GrdImprimir.TextMatrix(j, 3) & cedula1 & cedula2 & numcon & sCodigoDestino
         sNumeroTasa = GrdImprimir.TextMatrix(j, 1)
         sFechaExpedicion = GrdImprimir.TextMatrix(j, 6)
         sDestino = Replace(GrdImprimir.TextMatrix(j, 5), " ", "_")
         sValorTasa = Format(GrdImprimir.TextMatrix(j, 16), "#,##")
         sValorAlcohol = Format(GrdImprimir.TextMatrix(j, 17), "#,##")
         sValorTotal = Format(CDbl(GrdImprimir.TextMatrix(j, 16)) + CDbl(GrdImprimir.TextMatrix(j, 17)), "#,##")
         sCaja = GrdImprimir.TextMatrix(j, 2)
         
         sNombreTerminal = G_NomEmpresa
      
      
         sHoraExpedicion = hora
         sNumeroInterno = GrdImprimir.TextMatrix(j, 8)
         sPlaca = GrdImprimir.TextMatrix(j, 3)
         sEmpresaVehiculo = Replace(GrdImprimir.TextMatrix(j, 4), " ", "_")
         sNivelServicio = Replace(GrdImprimir.TextMatrix(j, 9), " ", "_")
      
         sUsuario = GrdImprimir.TextMatrix(j, 18)   'Usuario que registro la tasa de uso previamente
         s2Conductores = s2Conductores & "-COPIA-TU"
            
      
         sturno = GrdImprimir.TextMatrix(j, 15)   'turno que se regsitro la tasa de uso
         sPruebaAlcohol = presenprueba
         s2Conductores = Replace(s2Conductores, " ", "_")
         sPermanencias1 = Replace(sPermanencia1, " ", "_")
         sPermanencias2 = Replace(sPermanencia2, " ", "_")
         sPermanencias3 = Replace(sPermanencia3, " ", "_")
         sCodigoBarras = LNumConBarra
         
         If sValorAlcohol = "" Then
            sValorAlcohol = 0
         End If
         
        sTitulo = ""
        sTitNumAuto = "No_AUTORIZACION: "
        sNumAuto = ""
        sTitFecAuto = "FECHA_AUTORIZACION: "
        sFecAuto = ""
        
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'IMPFETU' "
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            If Tbl1!psval = "N" Then
            
                QRY1 = "SELECT canumaut, cafecaut, caprefijo " _
                        & "FROM cocaja WHERE canomcaj = '" & sCaja & "' "
                        
                Set Tbl2 = Dbs.Execute(QRY1)
                
                If Not Tbl2.EOF Then
                    sTitulo = "D.E.S/POS_" & Tbl2!caprefijo & " "
                    'sTitNumAuto = "No._AUTORIZACION: "
                    sNumAuto = Tbl2!canumaut
                    'sTitFecAuto = "FECHA_AUTORIZACION: "
                    sFecAuto = Mid(Tbl2!cafecaut, 1, 10)
                End If
            
            Else
                sTitulo = "Tasa_de_Uso_No."
                sNumAuto = "N/A"
                sFecAuto = "N/A"
            End If
        End If
        sTipoImpresion = "Z"
        
         scadena = sNumeroTasa & " " & sFechaExpedicion & " " & sHoraExpedicion & " " & sNumeroInterno & " " _
              & sPlaca & " " & sEmpresaVehiculo & " " & sDestino & " " & sNivelServicio & " " & sUsuario & " " & sCaja & " " _
              & sturno & " " & sValorTasa & " " & sValorAlcohol & " " & sValorTotal & " " & sPruebaAlcohol & " " & s2Conductores & " " _
              & sPermanencias1 & " " & sPermanencias2 & " " & sPermanencias3 & " " & sCodigoBarras & " " & sNumeroBarras & " " & sTipoImpresion & " " _
              & sTitulo & " " & sTitNumAuto & " " & sNumAuto & " " & sTitFecAuto & " " & sFecAuto
              
         'If sBandera > 0 Then
            Sleep 7000
         'End If
         s = Shell(G_Exe_Net_Impresion & " " & scadena, vbHide)
         
         'If sBandera > 0 Then
            Sleep 7000
         'End If
         sBandera = sBandera + 1
         If sTasas = "" Then
            sTasas = sNumeroTasa & "-" & sCaja
         Else
            sTasas = sTasas & ", " & sNumeroTasa & "-" & sCaja
         End If
      End If
      ProBar.Value = ProBar.Value + 1
   Next j
   reimprimir.Enabled = True
   ProBar.Value = ProBar.Max
   
   If sTasas <> "" Then
      sFecha = Fn_Hora_Servidor()
      
      Dbs.Execute " INSERT INTO gelogaud " _
               & "(lofecact, lofecnue, lofuncio, " _
               & "lofecha, loobser, lomodulo, lovalini, lovalfin ) " _
               & "VALUES (to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),  to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'" _
               & G_Login_Aut & "', to_date('" & sFecha & "', 'MM/dd/yyyy hh24:mi:ss'),'" & sTasas & "', 'IMPL', '" _
               & L_Busqueda & "', '" & L_Busqueda & "') "
      
   End If
   
   If Mid(G_DirImpBol, Len(G_DirImpBol), 1) = "\" Then
      sEliminar = G_DirImpBol & "Eliminar.BAT"
   Else
      sEliminar = G_DirImpBol & "\Eliminar.BAT"
   End If

   s = Shell(sEliminar, vbHide)
         
   Exit Sub
   
Errores:
   reimprimir.Enabled = True
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   Resume
End Sub

Private Sub Btnlimpiar_Click()
   L_Busqueda = ""
   Mskfecini.Mask = ""
   Mskfecini.text = ""
   Mskfecini.Mask = "##/##/####"
   Mskfecfin.Mask = ""
   Mskfecfin.text = ""
   Mskfecfin.Mask = "##/##/####"
   txtCodEmp.text = ""
   Lblcodemp.Caption = ""
   txtPlaca.text = ""
   lblPlaca.Caption = ""
   DIB_GRILLA
   Btnimprimir.Enabled = False
   ProBar.Value = ProBar.Min
   Mskfecini.SetFocus
End Sub

Private Sub Btnsalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   DIB_GRILLA
End Sub



Private Sub GrdImprimir_DblClick()
   On Error GoTo Errores
   If GrdImprimir.TextMatrix(GrdImprimir.Row, 0) = "" Then
      GrdImprimir.TextMatrix(GrdImprimir.Row, 0) = "->"
   ElseIf GrdImprimir.TextMatrix(GrdImprimir.Row, 0) = "->" Then
      GrdImprimir.TextMatrix(GrdImprimir.Row, 0) = ""
   End If
   
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
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
         Mskfecfin.SelStart = 0
         Mskfecfin.SelLength = Len(Mskfecfin.text)
         Exit Sub
      End If
  
      If Mskfecfin.ClipText <> "" Then
         If Not IsDate(Mskfecfin.text) Then
            MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
            Mskfecfin.SetFocus
            Mskfecfin.SelStart = 0
            Mskfecfin.SelLength = Len(Mskfecfin.text)
            Exit Sub
         End If
      End If
      
      If Mskfecini.ClipText <> "" Then
         If DateValue(Mskfecini.text) > DateValue(Mskfecfin.text) Then
            MsgBox "La Fecha Inicial No Puede Ser Mayor A La Fecha Final", 48, "Advertencia"
            Mskfecini.SetFocus
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
        Mskfecini.SelStart = 0
        Mskfecini.SelLength = Len(Mskfecini.text)
        Exit Sub
     End If
  
     If Mskfecini.ClipText <> "" Then
        If Not IsDate(Mskfecini.text) Then
           MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
           Mskfecini.SetFocus
           Mskfecini.SelStart = 0
           Mskfecini.SelLength = Len(Mskfecini.text)
           Exit Sub
        End If
     End If
  End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "58"
  txtCodEmp.text = ""
  Lblcodemp.Caption = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub

Private Sub Txtcodemp_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% = 8) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   If txtCodEmp.text <> "" Then
      fqry01 = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & txtCodEmp.text & "' "
              
      Set Ftb01 = Dbs.Execute(fqry01)
      
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
      Else
         MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
         forma = "16"
         forma1 = "58"
         buscar.Show 1
         txtCodEmp.SetFocus
      End If
   End If
End Sub

Private Sub TxtPlaca_DblClick()
   forma = "2"
   forma1 = "22"
   txtPlaca.text = ""
   lblPlaca.Caption = ""
   buscar.Show 1
   TxtPlaca_LostFocus
End Sub

Private Sub TxtPlaca_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Private Sub TxtPlaca_LostFocus()
   On Error GoTo Errores
   
   If Trim(txtPlaca.text) <> "" Then
      fqry01 = "SELECT hvconact " _
               & "FROM ophojveh " _
               & "WHERE hvplaca = '" & txtPlaca.text & "' "
         
        Set Ftb01 = Dbs.Execute(fqry01)
        
        If Not Ftb01.EOF Then
           lblPlaca.Caption = Ftb01!hvconact
           Ftb01.Close
        Else
           MsgBox "Placa De Vehiculo No Existe", 48, "Advertencia"
           forma = "2"
           forma1 = "22"
           buscar.Show 1
           txtPlaca.SetFocus
      End If
   End If
   Exit Sub
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
End Sub
