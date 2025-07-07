VERSION 5.00
Begin VB.Form inicio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4515
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Inicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4410
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   3240
         Top             =   1800
      End
      Begin VB.Image Image2 
         Height          =   3165
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         OLEDragMode     =   1  'Automatic
         Picture         =   "Inicio.frx":27A2
         Top             =   240
         Width           =   2820
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "TERMINAL DE TRANSPORTE S.A."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblWarning 
         Caption         =   "Advertencia.  Este programa esta protegido bajo las leyes Colombianas de Derechos de Autor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Lblciudad 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versión 4.2  2.004"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6915
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Módulo de Tasas de Uso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3600
         TabIndex        =   3
         Top             =   2040
         Width           =   5220
      End
      Begin VB.Label lblCompany 
         Caption         =   "Consultores Tecnológicos Asociados Ltda."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "Licencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  On Error GoTo terminar
   Fnum = FreeFile
   Open "C:\Windows\System\CFCONBOG.CTA" For Input Shared As #Fnum
   'VBCFGOPE
   Seek #Fnum, 1
   Do Until EOF(Fnum)
     Line Input #Fnum, linea
     X = InStr(1, linea, "=", 1)        'Busca la posicion del =
     varini = Mid(linea, 1, X - 1)      'Nombre de la variable
     valvar = Mid(linea, X + 1)         'Valor de la variable
     
     If varini = "CIUDAD" Then
       CIUDAD = UCase(valvar)                  'Ruta de la BD
     End If
   Loop
   Close #Fnum
    
   If CIUDAD = "" Then
      MsgBox "Variable CIUDAD no encontrada" + Chr(13) + Chr(13) + " Imposible iniciar la Aplicación", 16, "Fatal Error"
      End
   End If
    
   'Lblciudad.Caption = "DE" & " " & CIUDAD & " " & "S.A."
    
   Exit Sub
    
terminar:
    
    MsgBox "Ocurrio el siguiente error al tratar de leer el arhcivo de configuraciones" & Chr(13) & Chr(13) & "Error >> " & Err.Description, vbExclamation, "Error al iniciar la aplicacion"
    End
End Sub

Private Sub Timer1_Timer()
    Unload Me
    main.Show
End Sub
