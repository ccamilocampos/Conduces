VERSION 5.00
Begin VB.Form AviRecMult 
   Caption         =   "Aviso recaudo de multas y permanencias"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNo 
      Caption         =   "&NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton btnSi 
      Caption         =   "&SI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "POR FAVOR RECAUDE EL VALOR DE LA MULTA/PERMANENCIA, YA QUE NO SE VERÁ REFLEJADO EN LA VENTA DE LA TASA DE USO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "AviRecMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNo_Click()
    G_Mensaje_Opc = False
    Unload Me
End Sub

Private Sub btnSi_Click()
    G_Mensaje_Opc = True
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = G_Mensaje_Pago
    
    If G_Mensaje_Tipo = "E" Then
        btnSi.Visible = True
        btnNo.Visible = True
    Else
        btnSi.Visible = False
        btnNo.Visible = False
    End If
End Sub
