VERSION 5.00
Begin VB.Form FormEntrar 
   Caption         =   "Entrar"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FormEntrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
If Len(txtnombre) = 0 Or Len(txtpass) = 0 Then
  MsgBox "Faltan datos por introducir"
Else
  If Len(txtnombre) > 50 Or Len(txtpass) > 8 Then
    MsgBox "Los datos son demasiado largos"
  Else
    FormCliente.Winsock.SendData Chr(2) & "E" & txtnombre & "|" & txtpass
    DoEvents 'Pide permiso para entrar
    Unload Me
  End If
End If
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Height = 3090
Width = 4800
End Sub
