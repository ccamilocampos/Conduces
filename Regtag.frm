VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Regtag 
   Caption         =   "Captura concepto TAG"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox Mskfecha 
      Height          =   300
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "Fecha en que se hizo el parte"
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   -2147483644
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
      Left            =   6120
      Picture         =   "Regtag.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
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
      Left            =   6120
      Picture         =   "Regtag.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Limpia la forma actual"
      Top             =   3960
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
      Left            =   6120
      Picture         =   "Regtag.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar información en la base de datos"
      Top             =   3120
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
      Left            =   6120
      Picture         =   "Regtag.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ingresa la empresa a la base de datos"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
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
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Frame Frame5 
         Caption         =   "Datos TAG"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   4440
         Width           =   2055
         Begin VB.TextBox TxtValTag 
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
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label11 
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
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   975
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
         Height          =   2175
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   2895
         Begin VB.TextBox TxtObservaciones 
            Height          =   1815
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.TextBox TxtConcep 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado del TAG"
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
         Height          =   1215
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   2895
         Begin VB.OptionButton OptEstado 
            Caption         =   "Inactivo"
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
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "Activo"
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
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "Pendiente"
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
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Sistema"
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
         Height          =   3255
         Left            =   3360
         TabIndex        =   22
         Top             =   1080
         Width           =   2295
         Begin VB.TextBox TxtCodTag 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox Txtnumero 
            Alignment       =   2  'Center
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
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   1935
         End
         Begin MSMask.MaskEdBox Mskhora 
            Height          =   300
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Hora en que se hizo el parte"
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483644
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm "
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "# TAG"
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
            TabIndex        =   32
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Hora:"
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
            TabIndex        =   25
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Registro:"
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
            TabIndex        =   24
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Número:"
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
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtCodEmp 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNumInt 
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtPlaca 
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblConcep 
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
         Left            =   1320
         TabIndex        =   28
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Concepto: "
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
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Lblcodemp 
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
         Left            =   3480
         TabIndex        =   21
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa:"
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
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Interno:"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Placa:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   7320
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   6000
      X2              =   6000
      Y1              =   0
      Y2              =   5760
   End
End
Attribute VB_Name = "Regtag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Val_Tag
Dim L_Consecutivo_Tag
Dim estado As String


Private Sub BtnBuscar_Click()
    forma = "25"
  forma1 = "1"
  Txtnumero.text = ""
  buscar.Show 1
  Txtnumero.Enabled = True
  Txtnumero.SetFocus
  Txtnumero_LostFocus
  Txtnumero.Enabled = False
End Sub

Private Sub BtnGrabar_Click()
    
    On Error GoTo Errores

    If txtPlaca.text = "" Then
        MsgBox "La placa no puede estar vacia.", 48, "Advertencia"
        Exit Sub
    End If
    If txtCodEmp.text = "" Then
        MsgBox "El código de la empresa no puede estar vacio.", 48, "Advertencia"
        Exit Sub
    End If
    If TxtConcep.text = "" Then
        MsgBox "Debe elegir al menos un concepto.", 48, "Advertencia"
        Exit Sub
    End If
    If TxtObservaciones.text = "" Then
        MsgBox "La observación no puede estar vacia.", 48, "Advertencia"
        Exit Sub
    End If
    If Txtnumero.text = "" Then
        MsgBox "El número de TAG no puede ser vacio.", 48, "Advertencia"
        Exit Sub
    End If
    If Mskfecha.text = "" Then
        MsgBox "La fecha no puede estar vacia.", 48, "Advertencia"
        Exit Sub
    End If
    If Mskhora.text = "" Then
        MsgBox "La hora no puede estar vacia.", 48, "Advertencia"
        Exit Sub
    End If
    If TxtValTag.text = "" Then
        MsgBox "El valor no puede estar vacio.", 48, "Advertencia"
        Exit Sub
    End If
    If TxtCodTag.text = "" Then
        MsgBox "El código del TAG no puede ir vacio.", 48, "Advertencia"
        Exit Sub
    End If
    
    respuesta = MsgBox("¿Desea Grabar El TAG?", 32 + vbYesNo, "GRABAR")
   
   If respuesta = vbNo Then
      Exit Sub
   End If
   
   If OptEstado(0).Value = True Then
        estado = "P"
   ElseIf OptEstado(1).Value = True Then
        estado = "A"
   ElseIf OptEstado(2).Value = True Then
        estado = "I"
   End If
   
   Caption = Mid(Btngrabar.Caption, 2, 10)
   
   If Caption = "Grabar" Then
         
        'Validamos primero que el vehículo no cuente con otro TAG activo
        If OptEstado(1).Value = True Then
            QRY = "SELECT * FROM optag " _
                & "WHERE taplaca = '" & txtPlaca.text & "' " _
                & "AND taestado = 'A' " _
                & "AND tafecpag IS NULL "
            Set Tbl1 = Dbs.Execute(QRY)
        
            If Tbl1.EOF = False Then
                MsgBox "El vehículo cuenta con otro tag activo. Código: " & Tbl1!tacodtag & " ", 48, "Advertencia"
                Exit Sub
            End If
            Tbl1.Close
        End If
        
        
        QRY = "SELECT psval FROM geparsis " _
                & "WHERE pscod = 'NUMTAG' "
        
        Set Tbl1 = Dbs.Execute(QRY)
        
        Consecutivo = Tbl1!psval + 1
        Fecha = Mskfecha.text
        hora = Format(Mskhora.text, "hh:mm")
        
        'Validamos los valores subtotales y del iva
        'Consultamos el valor del IVA
        QRY = "SELECT psval FROM geparsis WHERE pscod = 'IVA' "
        Set Tbl1 = Dbs.Execute(QRY)
        
        G_Iva = Tbl1!psval
        
'        IvaTag = TxtValTag.text / (1 + (G_Iva / 100))
'        IvaTag = IvaTag * (G_Iva / 100)
'
'        decimales = CCur(IvaTag) - Int(IvaTag)
'        If decimales > 0.5 Then
'          IvaTag = Int(IvaTag) + 1
'        Else
'          IvaTag = Int(IvaTag)
'        End If
'
'        ValSubTag = CDbl(TxtValTag.text) - CDbl(IvaTag)

        Base = TxtValTag.text / (1 + (G_Iva / 100))
        decimales = CCur(Base) - Int(Base)
        If decimales > 0.5 Then
            Base = Int(Base) + 1
        Else
            Base = Int(Base)
        End If
        
        IvaTag = CDbl(TxtValTag.text) - CDbl(Base)
        ValSubTag = CDbl(Base)
        
        If Tbl1.EOF = False Then
            
            Dbs.BeginTrans
            On Error GoTo transac
            
            'Insertamos el TAG
            fqry01 = "INSERT INTO optag " _
                    & "(tanumero, taplaca, tacodemp, tafecreg, tahorreg, tacodtag, taconcep, taobserv, tavalor, taestado, taterminal, tasubtot, taiva)" _
                    & "VALUES (" & Consecutivo & ", '" & txtPlaca.text & "', '" & txtCodEmp.text & "', TO_DATE('" & Fecha & "', 'MM/DD/YYYY'), TO_DATE('" & hora & "', 'hh24:mi:ss'), '" _
                    & TxtCodTag.text & "', '" & TxtConcep.text & "', '" & TxtObservaciones.text & "', " & TxtValTag.text & ", '" & estado & "', '" & Mid(G_Tipo_Terminal, 1, 1) & "', " & ValSubTag & ", " & IvaTag & ") "
            
            Dbs.Execute (fqry01)
            
            If OptEstado(1).Value = True Then
                'Actualizamos la hoja del vehículo con el número de TAG registrado
                fqry01 = "UPDATE ophojveh SET hvnumerotag = '" & TxtCodTag.text & "' " _
                        & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                        & "AND hvcodemp = '" & txtCodEmp.text & "' "
                        
                Dbs.Execute (fqry01)
            End If
            
            
            'Si todo es correcto actualizamos el consecutivo
            fqry01 = "UPDATE geparsis SET psval = '" & Txtnumero.text & "' " _
                        & "WHERE pscod = 'NUMTAG'"
            
            Dbs.Execute (fqry01)
            
            Dbs.CommitTrans
                On Error GoTo transac
            
            MsgBox "Se registro el TAG exitosamente.", 48, "Advertencia"
            Limpiar
        End If
   ElseIf Caption = "Actualizar" Then
        
        'Buscamos si no tiene un tag activo
        If OptEstado(1).Value = True Then
            QRY = "SELECT * FROM optag " _
                & "WHERE taplaca = '" & txtPlaca.text & "' " _
                & "AND tanumero <> " & Txtnumero.text & "" _
                & "AND taestado = 'A' "
            Set Tbl1 = Dbs.Execute(QRY)
        
            If Tbl1.EOF = False Then
                MsgBox "El vehículo cuenta con otro tag activo. Código: " & Tbl1!tacodtag & " ", 48, "Advertencia"
                Exit Sub
            End If
            Tbl1.Close
        End If
        
        Fecha = Mskfecha.text
        hora = Format(Mskhora.text, "hh:mm")
        
        Dbs.BeginTrans
        On Error GoTo transac
        
        fqry01 = "UPDATE optag SET " _
                    & "tafecreg = TO_DATE('" & Fecha & "', 'MM/DD/YYYY'), " _
                    & "tahorreg = TO_DATE('" & hora & "', 'hh24:mi:ss'), " _
                    & "taconcep = '" & TxtConcep.text & "', " _
                    & "tacodtag = '" & TxtCodTag.text & "', " _
                    & "taobserv = '" & TxtObservaciones.text & "', " _
                    & "tavalor = " & TxtValTag.text & ", " _
                    & "taestado = '" & estado & "' " _
                    & "WHERE tanumero = " & Txtnumero.text & " " _
                    & "AND taplaca = '" & txtPlaca.text & "' "
                    
        Dbs.Execute (fqry01)
        
        If OptEstado(1).Value = True Then
            QRY = "UPDATE ophojveh SET hvnumerotag = '" & TxtCodTag.text & "' " _
                    & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                    & "AND hvcodemp = '" & txtCodEmp.text & "' "
                    
            Dbs.Execute (QRY)
        End If
        
        
        Dbs.CommitTrans
            On Error GoTo transac

        MsgBox "Se registro el TAG exitosamente.", 48, "Advertencia"
        Limpiar
   End If
   
   Exit Sub
   
Errores:
    MsgBox Err.Description, 48, "Advertencia"
    Exit Sub
transac:
   Dbs.RollbackTrans
   MsgBox Err.Description, 16, "Error " & Err.Number
   Exit Sub
   
Resume

End Sub

Private Sub Btnlimpiar_Click()
    Limpiar
End Sub

Private Sub Limpiar()
    fqry01 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'HORASIS' "
          
   Set Ftb01 = Dbs.Execute(fqry01)
   
   If Not Ftb01.EOF Then
      Fecha = Mid(Ftb01!psval, 1, 10)
      hora = Mid(Ftb01!psval, 12, 5)
      Mskfecha.text = Fecha
      Mskhora.text = hora
   End If
   Ftb01.Close
   
   fqry02 = "SELECT psval " _
            & "FROM geparsis " _
           & "WHERE pscod = 'NUMTAG' "
 

   Set Ftb02 = Dbs.Execute(fqry02)
  
   If Not Ftb02.EOF Then
      Txtnumero.text = Ftb02!psval + 1
   End If
   Ftb02.Close
   txtPlaca.Enabled = True
   txtPlaca.text = ""
   TxtNumInt.Enabled = True
   TxtNumInt.text = ""
   txtCodEmp.Enabled = True
   txtCodEmp.text = ""
   TxtConcep.text = ""
   TxtCodTag.text = ""
   Lblcodemp.Caption = ""
   lblConcep.Caption = ""
   Btngrabar.Enabled = True
   Btngrabar.Caption = "&Grabar"
   OptEstado(0).Enabled = True
   OptEstado(0).Value = True
   OptEstado(1).Enabled = True
   OptEstado(2).Enabled = True
   txtCodEmp.Enabled = True
   TxtObservaciones.text = ""
   'Mskfecha.Mask = ""
   'Mskfecha.Text = ""
   'Mskfecha.Mask = "##/##/####"
   'Mskhora.Mask = ""
   'Mskhora.Text = ""
   'Mskhora.Mask = "##:##"
   txtPlaca.SetFocus
End Sub

Private Sub Btnsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo Errores
    
    QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'NUMTAG' "
            
    Set Ftb01 = Dbs.Execute(QRY)
    
    If Ftb01.EOF = False Then
        Txtnumero.text = CDbl(Ftb01!psval) + 1
    End If
    Ftb01.Close
    
    QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'VALTAG' "
            
    Set Ftb01 = Dbs.Execute(QRY)
    
    If Ftb01.EOF = False Then
        TxtValTag.text = Ftb01!psval
    End If
    Ftb01.Close
    
     QRY = "SELECT psval " _
           & "FROM geparsis " _
          & "WHERE pscod = 'HORASIS' "
   
    Set Ftb01 = Dbs.Execute(QRY)
    
    If Not Ftb01.EOF Then
       Fecha = Mid(Ftb01!psval, 1, 10)
       hora = Mid(Ftb01!psval, 12, 5)
       Mskfecha.text = Fecha
       Mskhora.text = hora
    End If
    Ftb01.Close
    
    OptEstado(0).Value = 1
    
    Exit Sub
    
Errores:
    MsgBox Err.Description, 48, "Advertencia"
    Exit Sub

End Sub

Private Sub Txtcodemp_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub TxtCodTag_GotFocus()
    Senalar_Texto Me
End Sub

Private Sub TxtCodTag_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
  End If
End Sub

Private Sub TxtConcep_DblClick()
  forma = "15"
  forma1 = "61"
  Tabla_Catalogo = "CONCEPTOS DE PAGO TAG"
  TxtConcep.text = ""
  buscar.Show 1
  TxtConcep_LostFocus
End Sub

'Private Sub TxtConcep_KeyPress(KeyAscii As Integer)
'  Char = Chr(KeyAscii%)
'  KeyAscii = Asc(UCase(Char))
'  If (KeyAscii% >= 65 And KeyAscii% <= 90) Or (KeyAscii% = 8) Then
'      Exit Sub
'  Else
'     KeyAscii% = 0
'  End If
'
'End Sub

Private Sub TxtConcep_LostFocus()
    If TxtConcep.text <> "" Then
     fqry01 = "SELECT dsdes " _
              & "FROM gesuptip, gedetsuptip " _
             & "WHERE dscoddet = '" & TxtConcep.text & "' " _
               & "AND dscodtip = stcodtip " _
               & "AND stdes = 'CONCEPTOS DE PAGO TAG' "
               
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        lblConcep.Caption = Ftb01!dsdes
        Ftb01.Close
     Else
        Ftb01.Close
        MsgBox "Tipo De Concepto No Existente", 48, "Advertencia"
        forma = "15"
        forma1 = "61"
        Tabla_Catalogo = "CONCEPTOS DE PAGO TAG"
        buscar.Show 1
        TxtConcep.SetFocus
     End If
  End If
End Sub

'Private Sub TxtPlaca_DblClick()
'  forma = "2"
'  forma1 = "24"
'  Txtnumero.Text = ""
'  buscar.Show 1
'  Txtnumero.Enabled = True
'  Txtnumero.SetFocus
'  Txtnumero_LostFocus
'End Sub

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
    If txtPlaca.text <> "" Then
          
     longitud = Len(txtPlaca.text)
  
     If longitud < 6 Then
        MsgBox "El Numero De Placa Le Falta Ingresar Datos", 48, "Advertencia"
        txtPlaca.SetFocus
        Exit Sub
     End If
     
     fqry01 = "SELECT * " _
              & "FROM ophojveh " _
              & "WHERE hvplaca = '" & txtPlaca.text & "' "
              
     Set Ftb01 = Dbs.Execute(fqry01)
     
     If Not Ftb01.EOF Then
        TxtNumInt.text = Ftb01!hvnumint
        txtCodEmp.text = Ftb01!hvcodemp
        'TxtTipVeh.Text = Ftb01!hvtipveh
        
        L_CodigoEmpresaConvenio = ""
        If Not IsNull(Ftb01!hvempcon) Then
            L_CodigoEmpresaConvenio = Ftb01!hvempcon
        End If
        
        Ftb01.Close
        txtcodemp_LostFocus
        'Txttipveh_LostFocus
        txtPlaca.Enabled = False
        TxtNumInt.Enabled = False
        txtCodEmp.Enabled = False
        'TxtTipVeh.Enabled = False
        
     Else
         MsgBox "El Vehiculo No Existe En La Base de Datos", 48, "Advertencia"
         txtPlaca.SelStart = 0
         txtPlaca.SelLength = Len(txtPlaca.text)
         txtPlaca.SetFocus
     End If
   End If
End Sub

Private Sub Txtnumint_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtCodEmp_DblClick()
  forma = "16"
  forma1 = "67"
  txtCodEmp.text = ""
  buscar.Show 1
  txtcodemp_LostFocus
End Sub



Private Sub txtcodemp_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub txtcodemp_LostFocus()
   If txtCodEmp.text <> "" Then
       
      fqry01 = "SELECT * " _
               & "FROM ophojemp " _
              & "WHERE hecodigo = '" & txtCodEmp.text & "'"
           
      Set Ftb01 = Dbs.Execute(fqry01)
     
      If Not Ftb01.EOF Then
         Lblcodemp.Caption = Ftb01!henombre
         Ftb01.Close
         If txtPlaca.text <> "" Then
            fqry01 = "SELECT * " _
                     & "FROM ophojveh " _
                    & "WHERE hvplaca = '" & txtPlaca.text & "' "
             
            Set Ftb01 = Dbs.Execute(fqry01)
     
            If Not Ftb01.EOF Then
               Ftb01.Close
               fqry01 = "SELECT * " _
                        & "FROM ophojveh " _
                       & "WHERE hvplaca = '" & txtPlaca.text & "' " _
                         & "AND hvcodemp = '" & txtCodEmp.text & "' "
              
              
               Set Ftb01 = Dbs.Execute(fqry01)
              
               If Ftb01.EOF Then
                  Ftb01.Close
                  MsgBox "Este Vehiculo No Pertenece A " & txtCodEmp.text, 48, "Advertencia"
                  Exit Sub
               Else
                  Ftb01.Close
               End If
            Else
               Ftb01.Close
            End If
        
         End If
      Else
         Ftb01.Close
         fqry01 = "SELECT henombre " _
                    & "FROM ophojempcon " _
                   & "WHERE hecodigo = '" & txtCodEmp.text & "'"
        
         Set Ftb01 = Dbs.Execute(fqry01)
   
         If Not Ftb01.EOF Then
            Lblcodemp.Caption = Ftb01!henombre
           
         Else
            MsgBox "Codigo De Empresa No Existe", 48, "Advertencia"
            forma = "16"
            forma1 = "10"
            buscar.Show 1
            txtCodEmp.Enabled = True
            txtCodEmp.SetFocus
         End If
         Ftb01.Close
        
     End If
  End If
End Sub

Private Sub TxtNumero_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Txtnumero_LostFocus()

    If Txtnumero.text <> "" Then
        
        QRY = "SELECT tanumero, taplaca, tacodemp, TO_CHAR(tafecreg, 'MM/DD/YYYY') AS tafecreg, " _
                & "TO_CHAR(tahorreg, 'hh24:mi') AS tahorreg, " _
                & "taconcep, taobserv, tavalor, taestado, tacodtag " _
                & "FROM optag " _
                & "WHERE tanumero = " & Txtnumero.text & " " _
                & "AND tafecpag IS NULL "
                
        Set Tbl1 = Dbs.Execute(QRY)
        
        If Tbl1.EOF = False Then
        
            'Asignamos los valores
            txtPlaca.text = Tbl1!taplaca
            TxtPlaca_LostFocus
            'TxtCodEmp.Text = Tbl1!tacodemp
            Mskfecha.text = Mid(Tbl1!tafecreg, 1, 10)
            Mskhora.text = Tbl1!tahorreg
            TxtConcep.text = Tbl1!taconcep
            TxtConcep_LostFocus
            TxtObservaciones.text = Tbl1!taobserv
            TxtValTag.text = Tbl1!tavalor
            If Tbl1!taestado = "P" Then
                OptEstado(0).Value = True
            ElseIf Tbl1!taestado = "A" Then
                OptEstado(1).Value = True
            ElseIf Tbl1!taestado = "I" Then
                OptEstado(2).Value = True
            End If
            TxtCodTag.text = Tbl1!tacodtag
            
            'Cambiamos el caption del botón
            Btngrabar.Caption = "&Actualizar"
            
            'Inhabilitamos algunos
            txtPlaca.Enabled = False
            txtCodEmp.Enabled = False
            TxtNumInt.Enabled = False
        Else
            MsgBox "Error al consultar.", 48, "Advertencia"
        End If
                
        
    End If

End Sub

Private Sub Mskfecha_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskfecha_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskfecha_LostFocus()
  If Mskfecha.ClipText <> "" Then
     If Not IsDate(Mskfecha.text) Then
         MsgBox "No es una FECHA VALIDA ", 48, "Advertencia"
         Mskfecha.SetFocus
         Exit Sub
     End If
   
     If Val(Left(Mskfecha.text, 2)) < 1 Or Val(Left(Mskfecha.text, 2)) > 12 Then
        MsgBox "No es un MES VALIDO", 48, "Advertencia"
        Mskfecha.SetFocus
        Exit Sub
     End If
     
     anomulta1 = Year(Mskfecha.text)
     anosistemas1 = Year(anosistemas)
     If anosistemas1 <> anomulta1 Then
         anomulta = Mskfecha.text
         anosistemas1 = DateAdd("m", -2, anosistemas)
         If CDate(anomulta) < CDate(anosistemas1) Then
            On Error Resume Next
            MsgBox "El año es diferente al actual", 48, "Advertencia"
            Mskfecha.SetFocus
            Exit Sub
         End If
     End If
     
     QRY = "SELECT psval FROM geparsis " _
            & "WHERE pscod = 'HORASIS' "
     Set Tbl1 = Dbs.Execute(QRY)
     If Tbl1.EOF = False Then
        Fecha = Mid(Tbl1!psval, 1, 10)
        If CDate(Fecha) < CDate(Mskfecha.text) Then
            On Error Resume Next
            MsgBox "La fecha del tag es mayor a la fecha del sistema", 48, "Advertencia"
            Mskfecha.SetFocus
            Exit Sub
        End If
     End If
     Tbl1.Close
'     If CDate(Mskfecsis.Text) < CDate(Mskfecha.Text) Then
'         On Error Resume Next
'         MsgBox "la fecha de la evasion es mayor a la fecha del sistema", 48, "Advertencia"
'         Mskfecha.SetFocus
'         Exit Sub
'      End If
     
  End If
End Sub

Private Sub Mskhora_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Mskhora_KeyPress(KeyAscii As Integer)
  If ((KeyAscii% >= 48) And (KeyAscii% <= 57)) Or (KeyAscii% = 8) Then
     Exit Sub
  Else
     KeyAscii% = 0
  End If
End Sub

Private Sub Mskhora_LostFocus()
  If Mskhora.ClipText <> "" Then
     longitud = Len(Mskhora.ClipText)
     If longitud = 4 Then
        If Mskhora.ClipText < "0000" Or Mskhora.ClipText > "2400" Then
           MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
           Mskhora.SetFocus
           Exit Sub
        End If
        Minutos = Mid(Mskhora.ClipText, 3)
        If Minutos > 60 Then
           MsgBox "El Valor De Minutos Que Coloco No Son Validos", 48, "Advertencia"
           Mskhora.SetFocus
           Exit Sub
        End If
     Else
        MsgBox "la hora Que Coloco No Es Valida", 48, "Advertencia"
        Mskhora.SetFocus
        Exit Sub
     End If
  End If
End Sub

Private Sub Txtobservaciones_GotFocus()
   Senalar_Texto Me
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii%)
  KeyAscii = Asc(UCase(Char))
  If KeyAscii% = 124 Or KeyAscii% = 39 Or KeyAscii% = 34 Then
     KeyAscii% = 0
  End If
End Sub






