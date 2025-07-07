Attribute VB_Name = "funcion"
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function CapsLock() As Boolean
   CapsLock = GetKeyState(KeyCodeConstants.vbKeyCapital) = 1
End Function


Function Calcular_password(clave_usu As String) As String

  n_clave = 0
  For i = 1 To Len(clave_usu)
    valor = 0
    valor = Asc(Mid(clave_usu, i, 1))
    n_clave = n_clave + valor
  Next
  n_clave = 5000 - n_clave + Len(clave_usu)
  c_clave = Format(n_clave)
  Calcular_password = c_clave
  
End Function

Function Aplicar_Perfil(forma As Form, sforma As String) As Boolean
Dim tbl8 As Recordset
Dim dbs1 As Database
  
  On Error GoTo mensaje
  Qry1 = "SELECT eonomfrm, eonomobj, eoest " _
       & "FROM seestobj " _
       & "WHERE  eocodmod = '" & G_modulo & "'" _
       & "AND    eocodrol = '" & G_rolusu & "' " _
       & "AND    eonomfrm = '" & sforma & "' " _
       & "ORDER BY eocodmod, eocodrol, eonomfrm, eonomobj"
   Set Tbl2 = Dbs.Execute(Qry1)
   
   Do While Not Tbl2.EOF
  
            If Tbl2![eoest] = "Apagado" Then
              forma.Controls(Tbl2![eonomobj]).Enabled = False
            Else
              forma.Controls(Tbl2![eonomobj]).Enabled = True
            End If
            Tbl2.MoveNext
  Loop
  Tbl2.Close
  'dbs1.Close
  On Error GoTo 0
  Aplicar_Perfil = True
  Exit Function
  
mensaje:
  MsgBox "Problemas al Revisar SEGURIDAD.", 48, "Advertencia"
  Aplicar_Perfil = False

End Function
Function abrir(Control As Adodc, Cadena, tabla, servidor, basededatos)
If basededatos = "" Then
   Control.CommandType = adCmdTable
Else
   Control.CommandType = adCmdText
End If
   Control.ConnectionString = Cadena
    Control.RecordSource = tabla
   Control.Refresh
End Function
Function ABRIR_SQL(record_set As ADODB.Recordset, tabla, tipo, BASE_T)
    
    Set record_set = New ADODB.Recordset
    record_set.CursorType = adOpenDynamic
    record_set.LockType = adLockOptimistic
    record_set.Open tabla, BASE_T, , , tipo
    
    If record_set.State <> 1 Then
        MsgBox ("Tabla No se Pudo Abrir: " & tabla)
    End If
End Function


Public Sub Senalar_Texto(f As Form)
   On Error Resume Next
   With f.ActiveControl
      .SelStart = 0
      .SelLength = Len(.text)
   End With

End Sub

Public Function actualizacion()
'   On Error GoTo Errores
'
'      dbs.BeginTrans
'      On Error GoTo transac
'
'         On Error Resume Next
'         dbs.Execute "ALTER TABLE ophojveh " _
'            & "ADD HVBANO varchar2(5) NULL "
'
'         On Error Resume Next
'         dbs.Execute "ALTER TABLE ophojveh " _
'            & "ADD MUPASPLA NUMBER NULL "
'
'         On Error Resume Next
'         dbs.Execute "ALTER TABLE ophojveh " _
'            & "ADD HVCONDUCT NUMBER NULL "
'
'      dbs.CommitTrans
'   On Error Resume Next
'   Qry = "ALTER TABLE OPRUTA " _
'      & "ADD RUVALHOR VARCHAR2(1) DEFAULT 'N' "
'   Dbs.Execute Qry
'
'   Qry = "UPDATE OPRUTA " _
'      & "SET RUVALHOR = 'N' " _
'      & "WHERE RUVALHOR IS NULL "
'   Dbs.Execute Qry

      Exit Function
'
'Errores:
'   Exit Function
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Function
'
'transac:
'   dbs.RollbackTrans
'   MsgBox Err.Description, 16, "Error " & Err.Number
'   Exit Function

End Function

Public Function Tabular(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Sendkeys "{TAB}"
 End If
End Function

Function Fn_Parametros(ByVal Codparametro As String, ByVal sVerifica As Integer)
   On Error GoTo mensaje
   QRY = "SELECT psval " _
         & "FROM geparsis " _
        & "WHERE pscod = '" & Codparametro & "'"
     
   Set Tbl10 = Dbs.Execute(QRY)
   
   If Tbl10.EOF Then
      MsgBox "Falta Parámetro " & Codparametro & ". Imposible iniciar la Aplicación", 48, "Advertencia"
      
      If sVerifica = 0 Then
         End
      End If
      Exit Function
      
   Else
      Fn_Parametros = Tbl10![psval]
   End If
   Tbl10.Close
   Exit Function
  
mensaje:
   MsgBox Err.Description, 16 & "Error " & Err.Number
   Exit Function
End Function

Public Sub NumLetras(KeyAscii As Integer)
   Char = Chr(KeyAscii%)
   KeyAscii = Asc(UCase(Char))  ' mayuscula
   If (KeyAscii% >= 48 And KeyAscii% <= 57) Or (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Or (KeyAscii% = 32) Then
      Exit Sub
   Else
      KeyAscii% = 0
   End If
End Sub

Public Function Traer_Descripcion(sCodigo As String, sNomCatalogo As String) As String
   On Error GoTo Errores
   QRY = "SELECT  dsdes, dsest  " _
         & "FROM gesuptip, gedetsuptip " _
        & "WHERE dscoddet = '" & sCodigo & "' " _
          & "AND dscodtip = stcodtip " _
          & "AND stdes = '" & sNomCatalogo & "' "
    
   Set Tbl10 = Dbs.Execute(QRY)
   
   Traer_Descripcion = ""
   If Not Tbl10.EOF Then
      If Tbl10!dsest = "I" Then
         Tbl10.Close
         MsgBox "El detalle esta inactivo", 48, "Advertencia"
         Exit Function
      End If
      Traer_Descripcion = Tbl10!dsdes
   End If
   Tbl10.Close
   Exit Function
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Function
End Function

Public Function Fn_CalcularRangoHoras(ByVal tHora As String)
    
    L_HoraInicial = ""
    L_HoraFinal = ""
   
    If Len(Trim(G_MinHorTU)) > 0 Then
   
        sHorasIni = Mid(tHora, 1, 2)                        'Horas Inicial
        sMinutosIni = Mid(tHora, 4, 2)                      'Minutos Inicial
        
            sMinutos = sMinutosIni + Val(G_MinHorTU)
            sMinutosE = sMinutos Mod 60                     'Minutos por encima
            sHorasE = sHorasIni + (sMinutos \ 60)           'Horas por Encima
            
            sMinutos = sMinutosIni - Val(G_MinHorTU)
            If sMinutos >= 0 Then
                sMinutosD = sMinutos                        'Minutos por debajo
                sHorasD = sHorasIni                         'Horas por debajo
            Else
                sHorasD = sHorasIni
                sMinutos = sMinutos * (-1)
                sMinutosD = 60 - (sMinutos Mod 60)          'Minutos por Debajo
                sHorasD = sHorasD - (sMinutos \ 60) - 1     'Horas por Debajo
            End If
    End If
    
    If Len(sHorasE) = 1 Then sHorasE = "0" & sHorasE
    If Len(sMinutosE) = 1 Then sMinutosE = "0" & sMinutosE
    If Len(sHorasD) = 1 Then sHorasD = "0" & sHorasD
    If Len(sMinutosD) = 1 Then sMinutosD = "0" & sMinutosD
    
    L_HoraFinal = sHorasE & ":" & sMinutosE
    L_HoraInicial = sHorasD & ":" & sMinutosD
    
End Function

Public Function Fn_Equivalencia_Empresa(ByVal FnCodigo As String) As String
   On Error GoTo Fn_Errores
   
   QRY = "SELECT eeempdat, eeestado " _
         & "FROM opequemp " _
        & "WHERE eeempter = '" & FnCodigo & "' "
              
   Set Tbl10 = Dbs.Execute(QRY)
         
   If Not Tbl10.EOF Then
      Fn_Equivalencia_Empresa = Tbl10!eeempdat
   End If
   Tbl10.Close
   Exit Function
Fn_Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Function

Public Function Fn_Equivalencia_Ruta(ByVal FnCodigo As String) As String
   On Error GoTo Fn_Errores
   
   QRY = "SELECT errutdat " _
         & "FROM opequrut " _
        & "WHERE errutter = '" & FnCodigo & "' "
              
   Set Tbl10 = Dbs.Execute(QRY)
         
   If Not Tbl10.EOF Then
      Fn_Equivalencia_Ruta = Tbl10!errutdat
   End If
   Tbl10.Close
   Exit Function
Fn_Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Function

Public Function Fn_Equivalencia_Servicio(ByVal FnCodigo As String) As String
   On Error GoTo Fn_Errores
   
   QRY = "SELECT esserdat " _
         & "FROM opequser " _
        & "WHERE esserter = '" & FnCodigo & "' "
              
   Set Tbl10 = Dbs.Execute(QRY)
         
   If Not Tbl10.EOF Then
      Fn_Equivalencia_Servicio = Tbl10!esserdat
   End If
   Tbl10.Close
   Exit Function
Fn_Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
End Function

Public Function Traer_Nombre_Ruta(ByVal sCodigo As String) As String
   fqry01 = "SELECT rudestino " _
            & "FROM opruta " _
           & "WHERE rucodigo = '" & sCodigo & "' "
          
   Set Tbl10 = Dbs.Execute(fqry01)
   
   If Not Tbl10.EOF Then
      Traer_Nombre_Ruta = Tbl10!rudestino
      
   End If
   Tbl10.Close
End Function

Public Function Traer_Nombre_Conductor(ByVal sCodigo As String) As String
   QRY = "SELECT conombres, coapellidos " _
           & "FROM coconductor " _
          & "WHERE cocedula = '" & sCodigo & "'"
     
   Set Tbl10 = Dbs.Execute(QRY)
   
   If Not Tbl10.EOF Then
      Traer_Nombre_Conductor = Tbl10!conombres & " " & Tbl10!coapellidos
      
   End If
   Tbl10.Close
End Function

Public Sub Traer_Clase_Puestos(sPlaca As String)
   QRY = "SELECT hvnumpas, hvclase " _
           & "FROM ophojveh " _
          & "WHERE hvplaca = '" & sPlaca & "' "
     
   Set Tbl10 = Dbs.Execute(QRY)
   G_Clase_Vehiculo = ""
   G_Cantidad_Puestos_Vehiculo = 0
   If Not Tbl10.EOF Then
      If Not IsNull(Tbl10!hvclase) Then
         G_Clase_Vehiculo = Tbl10!hvclase
      End If
      G_Cantidad_Puestos_Vehiculo = Tbl10!hvnumpas
      
   End If
   Tbl10.Close
End Sub

Public Function Traer_Nombre_Funcionario(ByVal sLogin As String) As String
   QRY = "SELECT usnom " _
           & "FROM geususis " _
          & "WHERE uscodusu = '" & sLogin & "' "
     
   Set Tbl10 = Dbs.Execute(QRY)
   
   If Not Tbl10.EOF Then
      Traer_Nombre_Funcionario = Tbl10!usnom
      
   End If
   Tbl10.Close
End Function


Function Fn_Hora_Servidor() As String
   On Error GoTo Errores
   
   QRY = "SELECT sysdate " _
         & "FROM dual"
         
   Set Tbl10 = Dbs.Execute(QRY)
   
   If Not Tbl10.EOF Then
      Fn_Hora_Servidor = Format(Tbl10(0), "MM/dd/yyyy hh:mm:ss")
   Else
      Fn_Hora_Servidor = Format(Now, "MM/dd/yyyy hh:mm:ss")
   End If
   
   
   Exit Function
Errores:
   MsgBox Err.Description, 16, "Error " & Err.Number
   
End Function

Public Sub CargarFormatos()
    
    '     CARGANDO EL COLOR DEL RELOJ              '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLRELOJ' " _
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColReloj = Val(Tbl1!psval)
    Else
        ColReloj = &H8000000D
    End If
    Tbl1.Close
    
    '     CARGANDO EL COLOR DEL RELOJ              '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLMARQUE' " _
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColMarquesina = Val(Tbl1!psval)
    Else
        ColMarquesina = &H8000000D
    End If
    Tbl1.Close
    
    
    '     CARGANDO EL COLOR DEL TEXTO DE LA GRILLA         '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLTEXGRD' " _
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColTextoGrd = Val(Tbl1!psval)
    Else
        ColTextoGrd = 16777215
    End If
    Tbl1.Close
    
    '     CARGANDO EL COLOR 1 DE LA GRILLA         '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLFON1GRD' " _
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColFondo1 = Val(Tbl1!psval)
    Else
        ColFondo1 = 8388608
    End If
    Tbl1.Close
    
    '     CARGANDO EL COLOR 2 DE LA GRILLA         '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLFON2GRD' " _
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColFondo2 = Val(Tbl1!psval)
    Else
        ColFondo2 = 13074019
    End If
    Tbl1.Close
    
    '     CARGANDO EL TAMAÑO DE LETRA DE LA GRILLA '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'TAMLETRAGR' "
        
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        TamLetraGrilla = Val(Tbl1!psval)
    Else
        TamLetraGrilla = "16"
    End If
    Tbl1.Close
      '     CARGANDO EL COLOR DEL TITULO                     '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLTEXTIT' "
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColTextoTitulo = Val(Tbl1!psval)
    Else
        ColTextoTitulo = 16777215
    End If
    Tbl1.Close
    
      '     CARGANDO EL COLOR DEL TITULO                     '
    QRY = "SELECT PSVAL " _
        & "FROM GEPARSIS " _
        & "WHERE PSCOD = 'COLTITULO' "
        
    'ABRIR_SQL tbl1, qry, adCmdText, dbs
    Set Tbl1 = Dbs.Execute(QRY)
    If Not Tbl1.EOF Then
        ColTitulo = Val(Tbl1!psval)
    Else
        ColTitulo = 16777215
    End If
    Tbl1.Close
    
End Sub

Public Function Traer_Nombre_Empresa(ByVal fCodEmpresa As String) As String
   
   sLongitud = Len(fCodEmpresa)
   
   QRY = "SELECT henombre " _
               & "FROM  " _
               
   If sLongitud = 3 Then
      QRY = " ophojemp "
   Else
      QRY = " ophojempcon "
   End If
   QRY = QRY & "WHERE hecodigo = '" & fCodEmpresa & "'"
           
   Set Tbl10 = Dbs.Execute(QRY)
   Traer_Nombre_Empresa = ""
   If Not Tbl10.EOF Then
      Traer_Nombre_Empresa = Tbl10!henombre
   End If
   Tbl10.Close
      
End Function

Public Function Traer_Razon_Social_Empresa(ByVal fCodEmpresa As String) As String
   
      
   QRY = "SELECT cpvalor " _
         & "FROM opconcep " _
        & "WHERE cpcodemp = '" & fCodEmpresa & "' " _
          & "AND cpdescripcion = 'RAZSOCIAL' " _
     & "ORDER BY cpvalor "
          
   Traer_Razon_Social_Empresa = ""
   Set Ftb01 = Dbs.Execute(QRY)
   
   If Not Ftb01.EOF Then
      Traer_Razon_Social_Empresa = Ftb01!cpvalor
   End If
   Ftb01.Close
      
End Function

Public Function Traer_Nombre_Persona(ByVal sCodigo As String) As String
   QRY = "SELECT (cpnombres || ' ' || cpapellidos) AS cpnombres " _
           & "FROM opcompapersona " _
          & "WHERE cpcedula = '" & sCodigo & "'"
     
   Set Tbl10 = Dbs.Execute(QRY)
   Traer_Nombre_Persona = ""
   If Not Tbl10.EOF Then
      Traer_Nombre_Persona = Tbl10!cpnombres
      
   End If
   Tbl10.Close
End Function
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys text, wait
    Set WshShell = Nothing
End Sub

