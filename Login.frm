VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form flogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   2985
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1763.636
   ScaleMode       =   0  'User
   ScaleWidth      =   3675.974
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc consulta 
      Height          =   330
      Left            =   360
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Login.frx":0000
      OLEDBString     =   $"Login.frx":008E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
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
      Left            =   2880
      Picture         =   "Login.frx":011C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir de la forma actual"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   1
      Top             =   255
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Entrar"
      Default         =   -1  'True
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
      Left            =   1920
      Picture         =   "Login.frx":055E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   645
      Width           =   1845
   End
   Begin VB.Label lblLabels 
      Caption         =   "Turno Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   2280
   End
   Begin VB.Label Lblturno 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario:"
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
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
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
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connectdb As String
Dim PRN
Dim i
Dim MnuAdmOpe As String
'Option Explicit

Private Sub Btnsalir_Click()
  Unload Me
End Sub



Private Sub cmdOK_Click()

  'Validacion de existencia de parametros para poder
  'iniciar la aplicacion
  'NUMFAC, HOY, FINDIA
  
  
   G_modulo = "CND"
   G_Numero_Comparendo_Antiguo = 0
   ContAlco = 0
   'G_Url_Web_Services = "http://190.85.249.35/swvalida/integra.asmx"
   G_Url_Web_Services = Fn_Parametros("URLINT315", 0)
   'valida parametros
   
   Dbs.Execute "DELETE FROM optempo1 WHERE tempc = '" & Pc & "' "
   
   Fecha = Mid(Fn_Parametros("HORASIS", 0), 1, 10)
   LFechaSistema = Format(Fecha, "MM/DD/YYYY")
   hoy = DateValue(Fecha)
   FnValor = Fn_Parametros("CONPRU", 0)
   G_Minutos_Minimos_Parqueadero = Fn_Parametros("MINMINPARQ", 0)
   'Validacion de los datos del usuario
   G_ValidaServicio = Fn_Parametros("VALNIVSER", 0)
   G_valida_rodamiento = Fn_Parametros("VALVENROD", 0)
   G_impresion = Fn_Parametros("PARIMPTASA", 0)
   G_valida_credito = Fn_Parametros("VALCREROD", 0)
   'wecc 08/21/2012 se arma la variable a buscar como paramatro porque ellos manejan _
   el terminal CENTRAL Y SUR y cada terminal tienen un codigo en integra.
   
   sTerminalIntegra = "TIN" & G_Tipo_Terminal
   G_Codigo_Terminal = Fn_Parametros(sTerminalIntegra, 0)
   
   txtPassword.text = Calcular_password(txtPassword.text)
  
'   Qry = "SELECT usnom " _
'      & "FROM geususis " _
'      & "WHERE uscodusu = '" & txtUserName.Text & "' " _
'      & "AND uscla = '" & txtPassword.Text & "' " _
'      & "AND usest = 'A' "
'
'
'
'   Set Tbl2 = Dbs.Execute(Qry)
'      If Tbl2.EOF = True Then
'      If CapsLock = True Then
'         MsgBox "El boton de las mayusculas esta encendido", 48, "Advertencia"
'
'      Else
'         MsgBox "CONTRASEÑA Incorrecta / USUARIO Bloqueado", 16, "Error de Conexión !"
'      End If
'      Unload flogin
'      Exit Sub
'  End If
'  Tbl2.Close

'*********************************************************************
    
    
    If G_Turno = "1" Then
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'HORAPETUR1'"
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
            horven = Format(Tbl1!psval, "hh:mm")
            Tbl1.Close
        Else
            MsgBox "No se encontro el parametro HORAPETUR1. " & Chr(13) & "No se puede iniciar la aplicación.", 48, "Advertencia"
            Unload flogin
            Exit Sub
        End If
        
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'HORASIS' "
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Not Tbl1.EOF Then
         horactual = Format(Tbl1!psval, "hh:mm")
         If CDate(horactual) < CDate(horven) Then
             MsgBox "No se puede abrir turno 1 antes de las 12:05 am. " & Chr(13) & "Proceso cierre automático ejecutandose", 48, "Advertencia"
             Unload flogin
             Exit Sub
         ElseIf (CDate(horactual) > CDate("21:00")) And (CDate(horactual) < CDate("23:59")) Then
             MsgBox "No se puede abrir turno 1 antes de las 12:05 am. " & Chr(13) & "Proceso cierre automático ejecutandose", 48, "Advertencia"
             Unload flogin
             Exit Sub
         End If
        End If
    End If
    
    If G_Turno = 1 Then
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'HORASIS' "
        Set Tbl1 = Dbs.Execute(QRY)
        
        hora1 = "23:50"
        hora2 = "23:59"
          
        If Not Tbl1.EOF Then
           horactual = Format(Tbl1!psval, "hh:mm")
           
           If CDate(horactual) >= CDate(hora1) And CDate(horactual) <= CDate(hora2) Then
              MsgBox "No se puede abrir turno 1 antes de las 12:05 am. " & Chr(13) & "Por favor intente nuevamente mas tarde", 48, "Advertencia"
               Unload flogin
               Exit Sub
           End If
        End If
    End If
   

   QRY = "SELECT usnom, usfeclimite, usest, uscla  " _
         & "FROM geususis " _
         & "WHERE uscodusu = '" & txtUserName.text & "' "
           
      
   Set Ftb01 = Dbs.Execute(QRY)
  
   If Ftb01.EOF Then
      MsgBox "USUARIO No existe", 48, "Advertencia"
      Unload flogin
      Exit Sub

   Else
      '*********************************************
      
      If Ftb01!usest = "I" Then
         MsgBox "USUARIO Bloqueado", 48, "Advertencia"
         Ftb01.Close
         Unload flogin
         Exit Sub
      End If
      
      If CapsLock = True Then
         MsgBox "El boton de las mayusculas esta encendido", 48, "Advertencia"
         txtPassword.SetFocus
         Exit Sub
      End If
      sFechaVencimientoContraseña = ""
      If Not IsNull(Ftb01!usfeclimite) Then
         sFechaVencimientoContraseña = Format(Ftb01!usfeclimite)
      Else
         sFechaVencimientoContraseña = LFechaSistema
      End If
      If sFechaVencimientoContraseña <> "" Then
         sDias = DateDiff("d", LFechaSistema, sFechaVencimientoContraseña)
         If sDias <= 3 And sDias > 0 Then
            MsgBox "En " & sDias & " días se bloquea el usuario, debe de cambiar la contraseña", 48, "Advertencia"
            
         ElseIf sDias <= 0 Then
            MsgBox "Tiene que cambiar la contraseña para poder ingresar al sistema", 48, "Advertencia"
            Ftb01.Close
            Login = txtUserName.text
            txtPassword.text = ""
            password.Show 1
            txtPassword.SetFocus
            Exit Sub
         End If
      End If
   
      sContrasenaDesincriptada = Ftb01!uscla
      
      If Trim(txtPassword.text) <> sContrasenaDesincriptada Then
         MsgBox "CONTRASEÑA incorrecta", 48, "Advertencia"
         Ftb01.Close
         Unload flogin
         Exit Sub
      End If
   End If
   Ftb01.Close
   
      'ABRE LA BASE DE DATOS DE SEGURIDAD
  On Error GoTo msgerr1
  
  'Determina el rol del usuario en el modulo
  QRY = "SELECT umcodrol " _
      & "FROM seusmorol " _
      & "WHERE umcodusu = '" & txtUserName.text & "' " _
      & "AND umcodmod = '" & G_modulo & "' "
      
  Set Tbl2 = Dbs.Execute(QRY)
  If Tbl2.EOF = True Then
    MsgBox "USUARIO No tiene ROL dentro del MODULO", 16, "Error de Conexión"
    Unload flogin
    Exit Sub
  Else
    G_rolusu = Tbl2![umcodrol]
  End If
  Tbl2.Close
  
  'Aplica perfil del rol
  If Not Aplicar_Perfil(main, "main") Then
    Unload flogin
    Exit Sub
  End If
  'se colocan en estado enable porque ellos no los trabajan.
  main.mnuconpen.Enabled = False
  main.mnuingconpen.Enabled = False
  
  'VUELVE ABRIR LA BASE DE DATOS DEL MODULO
  On Error GoTo msgerr2
       
  tmphoy = Format(hoy, "mmm dd yyyy")
  
  main.MnuConductor.Enabled = True
  
  MsgBox "BIENVENIDO AL SISTEMA." + Chr(13) + Chr(13) + "  Hoy es < " + tmphoy + " >", 64, "O.K."
  
  main.Mnulogin.Enabled = False
  main.Mnucambiopassword.Enabled = True
  Login = txtUserName.text
  'iniciar_variables
    
  sFecha = Format(Now, "mm/dd/yyyy hh:mm")
  QRY = " INSERT INTO gelogusu " _
          & "(lucodusu, lucodmod,lufecent) VALUES ('" _
          & txtUserName.text & "','CND',to_date('" & sFecha & "','MM/DD/YYYY hh24:mi'))"
  
  Dbs.Execute "commit"
         
  Unload flogin
    
   PRN = "PANTALLA"
  
  
  With main
    .Toolbar1.Buttons(1).Enabled = False
    
    If .Mnuhojvidveh.Enabled = True Then
       'vehiculos
       .Toolbar1.Buttons(3).Enabled = True
    End If
    
    If .Mnuhojvidemp.Enabled = True Then
       'empresa
       .Toolbar1.Buttons(4).Enabled = True
    End If
    If .Mnuapetur.Enabled = True Then
       'cierre turno
       .Toolbar1.Buttons(6).Enabled = True
    End If
    If G_cierredia = "S" Then
       If .Mnucietur.Enabled = True Then
          'cierre dia
          .Toolbar1.Buttons(7).Enabled = True
       End If
    Else
       .Mnucietur.Enabled = False
       .Toolbar1.Buttons(7).Enabled = False
    End If
    If .Mnuvencon.Enabled = True Then
       'venta conduce
       .Toolbar1.Buttons(8).Enabled = True
    End If
    If .Mnucancon.Enabled = True Then
       'anulacion copnduce
       .Toolbar1.Buttons(9).Enabled = True
    End If
    
    If .Mnuregmul.Enabled = True Then
       'anulacion copnduce
       .Toolbar1.Buttons(11).Enabled = True
    End If
    If .Mnucanmul.Enabled = True Then
       'anulacion copnduce
       .Toolbar1.Buttons(12).Enabled = True
    End If
    If .Mnuanumultas.Enabled = True Then
       'anulacion copnduce
       .Toolbar1.Buttons(13).Enabled = True
    End If
    
    If .mnuconsolidado.Enabled = True Then
       'anulacion copnduce
       .Toolbar1.Buttons(15).Enabled = True
    End If
  End With
  
  If MnuAdmOpe = "F" Then
     main.Mnuadmiopera.Visible = False
     main.Toolbar1.Buttons(2).Visible = False
     main.Toolbar1.Buttons(3).Visible = False
     main.Toolbar1.Buttons(4).Visible = False
     
     'Reportes EMGAF
     main.Mnurepemgaf.Visible = False
  ElseIf MnuAdmOpe = "V" Then
     main.Mnuadmiopera.Visible = True
     main.Toolbar1.Buttons(2).Visible = True
     main.Toolbar1.Buttons(3).Visible = True
     main.Toolbar1.Buttons(4).Visible = True
     
     'Reportes EMGAF
     main.Mnurepemgaf.Visible = True
     main.Mnurepemgaf.Enabled = True
  End If
  
'  If Mid(G_Tipo_Terminal, 1, 1) = "N" Or Mid(G_Tipo_Terminal, 1, 1) = "S" Then
'    main.Mnucanmul.Enabled = False
'  End If
  
  main.StatusBar1.Panels(1).text = "Usuario : " & Login
  main.StatusBar1.Panels(2).text = "PC : " & Pc
  main.StatusBar1.Panels(3).text = "Fecha Actual : " & tmphoy
  main.StatusBar1.Panels(4).text = "Turno : " & G_Turno
  main.StatusBar1.Panels(5).text = "PRN : " & PRN
  
  
'  If G_ImpCon = "Z" Then
'    NOMDIR = G_DirImpBol & "logo.BAT"
'
'    NOMBRE = Dir(NOMDIR)
'
'    If NOMBRE <> "" Then
'       Kill NOMDIR
'    End If
'
'
'    On Error Resume Next
'    Open NOMDIR For Append As #2
'
'    If Err Then
'       MsgBox Error$, 48
'       Close #2
'       Exit Sub
'    Else
'        a = "copy " & G_DirImpBol & "logo.txt PRN"
'        Print #2, a
'        Close #2
'     End If
'     Shell G_DirImpBol & "logo.BAT", vbHide
' End If
  
  
  
  Exit Sub
  
msgerr1:
   
   MsgBox "No se pudo Abrir Base de Datos de SEGURIDAD." _
         + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error" _
         + Chr(13) + Chr(13) + " ERROR >> " + Err.Description, 16, "Fatal Error"
   Unload flogin
   End

msgerr2:

   
   
   MsgBox "No se pudo Abrir Base de Datos de TASAS DE USO." _
         + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error" _
         + Chr(13) + Chr(13) + " ERROR >> " + Err.Description, 16, "Fatal Error"
   Unload flogin
   End
   Resume '
End Sub



Private Sub Form_Load()
   Dim TurIni As Double 'wecc para saber cual es el turno en que inicia la caja
   
   'Lee el archivo de configuraciones y valida
   'la definicion de las variables :
   'DBPATH      camino y archivo de base de datos
   'RPPATH      ruta para los reportes
   'DBSEC       ruta de seguridades
   'KEY         Numero de activacion para control pirata
   'PC          DONDE SE ENCUENTRA LA APLICACION
      
   On Error GoTo terminar
   
   DbPath = ""
   RpPath = ""
   numkey = ""
   Fnum = FreeFile
   Open "C:\Windows\System\CFCONBOG.CTA" For Input Shared As #Fnum
   Seek #Fnum, 1
   Do Until EOF(Fnum)
     Line Input #Fnum, linea
     x = InStr(1, linea, "=", 1)        'Busca la posicion del =
     varini = Mid(linea, 1, x - 1)      'Nombre de la variable
     valvar = Mid(linea, x + 1)         'Valor de la variable
     
     If varini = "STRING" Then
       DbPath = valvar                  'Ruta de la BD
     End If
     
     If varini = "RPPATH" Then
       RpPath = valvar                  'Ruta de los reportes
     End If
     
     If varini = "KEY" Then
       numkey = valvar                  'Ruta de los reportes
     End If
     
     If varini = "PC" Then
       Pc = valvar                      'Terminal
     End If
     
     If varini = "CIERRETURNO" Then  'permiso de cierre de turno
       cierreturno = valvar
     End If
     
     If varini = "DIRTEMP" Then
       DIRECCION = valvar
     End If
     
     If varini = "MANUAL" Then
       G_Manual = valvar
     End If
     
     If varini = "CIERREDIA" Then
       G_cierredia = valvar
     End If
     
     If varini = "DIRIMPCON" Then
       G_DirImpBol = UCase(valvar)                   'Ruta de la BD
     End If
     
     If varini = "IMPCON" Then
       G_ImpCon = valvar                   'Seguridad
     End If
     
     If varini = "MNUADMOP" Then
       MnuAdmOpe = valvar                  'Seguridad
     End If
     
     If varini = "IPUNIX" Then
        G_IpUnix = valvar                  'Ip Unix
     End If
     If varini = "CONEREPO" Then
        G_Conexion_reportes = valvar                   'Ip Unix
     End If
     If varini = "CIUDAD" Then
       CIUDAD = UCase(valvar)                  'Ruta de la BD
     End If
     If varini = "TIPOTERMINAL" Then
       G_Tipo_Terminal = UCase(valvar)                 'Ruta de la BD
     End If
     If varini = "VALIDAINSTANCIA" Then
       G_Valida_Instancia = Trim(UCase(valvar))                 'Ruta de la BD
     End If
     If varini = "HABLECCED" Then
       G_habLecCedula = Trim(UCase(valvar))                 'Ruta de la BD
     End If
     If varini = "PRINTERNAME" Then
         G_PrinterName1 = valvar
      End If
      If varini = "PRINTERDRIVER" Then
         G_PrinterDriver1 = valvar
      End If
      If varini = "PRINTERPORT" Then
         G_PrinterPort1 = valvar
      End If
      If varini = "EXENETIMPRESION" Then
         G_Exe_Net_Impresion = valvar
      End If
      
      If varini = "MANLECTOR" Then
        G_ManLector = valvar
     End If
     
     If varini = "DBPATHHUELLA" Then
         DbPathHuella = valvar
     End If
      
     If varini = "IMPPERMAMULTA" Then
         G_ImpEpsonZebra = UCase(valvar)
     End If
      
     If varini = "RUTAARCHIVOSTEMPORALES" Then
         G_Ruta_Archivos_Temporales = UCase(valvar)
     End If
     
     If varini = "INFLUENCIA" Then
         G_PermiteInfluencia = UCase(valvar)
     End If
     
     If varini = "LEERPEAJE" Then
         G_PermiteLeerPeaje = UCase(valvar)
     End If
'     If varini = "NUMIMPTAS" Then
'         G_NumImpTasa = valvar
'     End If

    G_CierreAutomatico = 0
      
   Loop
   Close #Fnum
   
   
   If DbPath = "" Then
      MsgBox "Variable STRING no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If RpPath = "" Then
      MsgBox "Variable RPPATH no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If CIUDAD = "" Then
      MsgBox "Variable CIUDAD no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   
   If Pc = "" Then
      MsgBox "Variable PC no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_DirImpBol = "" Then
      MsgBox "Variable DIRIMPCON no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   'El numero de la version es la suma de los valores ascii de cada una
   'de las letras que componen la sigla de la empresa (mayusculas)
   'mas la constante 2525
   
   If numkey = "" Or numkey <> "2878" Then
      MsgBox "Copia NO AUTORIZADA para ejecutarse" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
     
   'wecc 02/25/2002
   If cierreturno = "" Then
      MsgBox "Variable CIERRETURNO no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   'WECC 05/08/2002
   If DIRECCION = "" Then
      MsgBox "Variable DIRTEMP no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   'WECC 05 /14/2002
   If G_Manual = "" Then
      MsgBox "Variable MANUAL no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_cierredia = "" Then
      MsgBox "Variable CIERREDIA no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_ImpCon = "" Then
      MsgBox "Variable IMPCON no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If MnuAdmOpe = "" Then
      MsgBox "Variable MNUADMOP no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_IpUnix = "" Then
      MsgBox "Variable IPUNIX no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   
   If G_Conexion_reportes = "" Then
      MsgBox "Variable CONEREPO no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   If G_Tipo_Terminal = "" Then
      MsgBox "Variable TIPOTERMINAL no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   If G_Valida_Instancia = "" Then
      MsgBox "Variable VALIDAINSTANCIA no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   If Trim(G_habLecCedula) = "" Then
      MsgBox "Variable HABLECCED no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
   
    'Definicion de la primera impresora en un puerto paralelo
   If G_PrinterName1 = "" Then
      MsgBox "Variable PRINTERNAME no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   If G_PrinterDriver1 = "" Then
      MsgBox "Variable PRINTERDRIVER no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   If G_PrinterPort1 = "" Then
      MsgBox "Variable PRINTERPORT no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_ImpCon = "N" Then
      If G_Exe_Net_Impresion = "" Then
         MsgBox "Variable EXENETIMPRESION no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
         Unload flogin
         End
      End If
   End If
   
   If G_ManLector = "" Then
      MsgBox "Variable MANLECTOR no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If DbPathHuella = "" Then
      MsgBox "Variable DBPATHHUELLA no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_ImpEpsonZebra = "" Then
      MsgBox "Variable IMPPERMAMULTA no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   If G_ImpEpsonZebra <> "P" And G_ImpEpsonZebra <> "Z" Then
      MsgBox "Valor de variable IMPPERMAMULTA debe ser P o Z" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
      
   End If
   If G_Ruta_Archivos_Temporales = "" Then
      MsgBox "Valor de variable RUTAARCHIVOSTEMPORALES no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_PermiteInfluencia = "" Then
      MsgBox "Variable INFLUENCIA no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
   
   If G_PermiteLeerPeaje = "" Then
      MsgBox "Variable LEERPEAJE no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      Unload flogin
      End
   End If
     
'   If G_NumImpTasa = "" Then
'      MsgBox "Variable NUMIMPTAS no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
'      Unload flogin
'      End
'   End If
   
   If G_Valida_Instancia = "S" Then
      If App.PrevInstance = True Then
         MsgBox "La Aplicación Ya se Es Encuentra en Uso", 48, "Advertencia"
         Unload flogin
         End
      End If
   End If
   'validamos formato de fecha
    Fecha = Date
    Fecha1 = Format(Date, "MM/DD/YYYY")
    
    If Trim(Fecha) <> Trim(Fecha1) Then
        MsgBox "El Formato de Fecha es Incorrecto, el Formato " & _
               "de Fecha Corta es MM/DD/YYYY ó MM/DD/AAAA. ", 16, "Error "
        Unload flogin
        Exit Sub
    End If
    
    'validamos el separador de Decimales para el número
    numero = Rnd
    caracter = Mid(numero, 2, 1)
     
    If caracter <> "." Then
        MsgBox "El Formato de Número es Incorrecto." & _
               Chr(13) + " El Separador de Decimales es Punto(.)", 16, "Error "
    
        Unload flogin
        Exit Sub
    End If

    'validamos el separador de miles para el número
    numero = 1000
    numero = Format(numero, "#,##.0")
    caracter = "," 'Mid(numero, 2, 1)

    If caracter <> "," Then
        MsgBox "El Formato de Número es Incorrecto." & _
               Chr(13) + " El Separador de Miles es Coma(,)", 16, "Error "
    
        Unload flogin
        Exit Sub
    End If
   
   'cargo el string de conexión a el servidor de oracle
   'proveedor = dbpath
   'main.Label1.Caption = main.Label1.Caption & " " & G_Tipo_Terminal
   
   Set Dbs = New ADODB.Connection
   Dbs.CursorLocation = adUseClient
   Dbs.ConnectionString = DbPath  'proveedor '& nombrebasedatos
   Dbs.Open
   
   
   Call actualizacion
   'Abre la BASE DE DATOS para el sistema
   
   'wecc 03/29/2010 parametro para validar el horario al mmomento de vender la tasa de uso
   G_Valida_Horario = Fn_Parametros("VALHORARIO", 0)
   G_MinHorTU = Fn_Parametros("MINHORTU", 0)

   G_NomEmpresa = Fn_Parametros("NOMEMPRESA", 0)
   main.Label1.Caption = G_NomEmpresa & " - " & G_Tipo_Terminal & " - PROTECH"
   'wecc 04/11/2005
   'valida parametros
   G_Bloqueo_Tasa = Fn_Parametros("BLOQUETASA", 0)
   G_Validar_Ser = Fn_Parametros("VALIDARSER", 0)
   FnValor = Fn_Parametros("PORORI", 0)
   FnValor = Fn_Parametros("PORTRA", 0)
   FnValor = Fn_Parametros("ALCOTODOS", 0)
   FnValor = Fn_Parametros("ALCOORIG", 0)
   FnValor = Fn_Parametros("ALCOTRAN", 0)
   FnValor = Fn_Parametros("NUMCAR", 0)
   
   G_ValalcOri = Fn_Parametros("VALALCORI", 0)
   G_ValalcTra = Fn_Parametros("VALALCTRA", 0)
   G_ValalcInf = Fn_Parametros("VALALCINF", 0)
   G_NumSal = Fn_Parametros("NUMSAL", 0)
   G_CantRegist = Fn_Parametros("CANTREGIST", 0)

   'WECC 04/24/2002 *********************
   
   'Calculo
   'wecc 01/31/2002
   
   QRY = "SELECT canumtur, caturcaj, cacercaj, caturini " _
         & "FROM cocaja " _
        & "WHERE canomcaj = '" & Pc & "'"
     
   Set Tbl2 = Dbs.Execute(QRY)
    If Tbl2.EOF = False Then
     numturnos = Tbl2(0)
     G_maxtur = Tbl2(0)
     turactivo = Tbl2(1)
     cajacerrada = Tbl2(2)
     TurIni = Tbl2(3)
     Tbl2.Close
   Else
      QRY = "SELECT canumtur, caturcaj, cacercaj " _
            & "FROM cocaja " _
        & "WHERE canomcaj = 'CAJA1'"
         
      Set Tbl2 = Dbs.Execute(QRY)
      If Tbl2.EOF = False Then
         turactivo = Tbl2(1)
      End If
      Tbl2.Close
   End If
   
    If cierreturno = "S" Then
       If cajacerrada = "S" Then
         
         If G_cierredia = "N" Then
             'If turactivo = G_maxtur Then
             '    fqry01 = "SELECT * " _
                          & "FROM cocaja " _
                         & "WHERE caciedia = 'N' " _
                         & "AND canomcaj = '" & pc & "' " _
                           & "AND caestado = 'Activa' "
                 
             '    Set Tbl1 = Dbs.Execute(Qry)
             '         If tbl1.EOF = False Then
             '               MsgBox Chr(13) & "      No Se puede Iniciar Turno            " & Chr(13) & "         No Se Ha Cerrado Día" & Chr(13), 48, "Advertencia"
             '               Unload flogin
             '               Exit Sub
             '         End If
             '         tbl1.Close
             ' End If
         End If
         
         Dbs.Execute " UPDATE cocaja " _
                       & "SET cacercaj = 'N' " _
                     & "WHERE canomcaj = '" & Pc & "'"
      
         If Val(turactivo) < Val(numturnos) Then
           G_Turno = Val(turactivo) + 1
         Else
            'wecc 11/14/2002 se parcha porque el manejo del turnos
            'se cambio
           'G_turno = 1
           G_Turno = TurIni
         End If
         'wecc 11/14/2002 se parcha porque el manejo del turnos
         'se cambio
         'If G_turno = 1 Then
         If G_Turno = TurIni Then
            
            hora = Mid(Fn_Parametros("HORASIS", 0), 12, 15)
            
            hora = Format(hora, "hh:mm")
            Dbs.Execute "UPDATE cocaja " _
                         & "SET cahoraap = to_date('" & hora & "','HH24:MI:SS') " _
                       & "WHERE canomcaj = '" & Pc & "' "
         End If
         
         Dbs.Execute "UPDATE cocaja " _
                      & "SET caturcaj = '" & G_Turno & "' " _
                    & "WHERE canomcaj = '" & Pc & "' "
         
         LblTurno.Caption = G_Turno
         main.Mnuapetur.Enabled = True
       Else
         G_Turno = Val(turactivo)
         LblTurno.Caption = G_Turno
       End If
   Else
   'si no tiene derecho a cambio de turno
       G_Turno = Format(turactivo)
       LblTurno.Caption = G_Turno
   End If
   
   If G_Turno = "4" Then
      main.Mnucietur.Enabled = True
   Else
      main.Mnucietur.Enabled = False
   End If
    
    
    
   If Mid(G_DirImpBol, Len(G_DirImpBol), 1) = "\" Then
      NOMDIR = G_DirImpBol & "Eliminar.BAT"
   Else
      NOMDIR = G_DirImpBol & "\Eliminar.BAT"
   End If
    
   NOMBRE = Dir(NOMDIR)
   
   If NOMBRE <> "" Then
      Kill NOMDIR
   End If
   
   
   On Error Resume Next
   Open NOMDIR For Append As #2
   
   If Err Then
      MsgBox Error$, 48
      Close #2
      Exit Sub
   Else
      'C:\Users\Willia~1\AppData\Local\Temp\*.rpt /F
      If Mid(G_Ruta_Archivos_Temporales, Len(G_Ruta_Archivos_Temporales), 1) = "\" Then
         Print #2, "DEL " & G_Ruta_Archivos_Temporales & "*.rpt /F"
      Else
         Print #2, "DEL " & G_Ruta_Archivos_Temporales & "\*.rpt /F"
      End If
      
       Close #2
    End If
           
   Exit Sub
   
terminar:

   MsgBox Err.Description, 16, "Error " & Err.Number
   MsgBox "Archivo de Configuraciones " _
         + Chr(13) + "o Base de Datos Inexistentes" _
         + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"

   Unload flogin
   End

BaseServidor:
    G_BaseServidor = False
   Exit Sub
   Resume
    
End Sub

Private Sub txtPassword_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(LCase(Char))  ' mayuscula
  If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 97 And KeyAscii% <= 122) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub


