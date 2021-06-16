VERSION 5.00
Begin VB.Form FormGrafico 
   Caption         =   "Gráfico de paquetes recibidos"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctgrafico 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   600
      ScaleHeight     =   3795
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "FormGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
FormMedia.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer

MsgBox ("Tpo=" & Tiempopaq)
For i = 0 To Tiempopaq
  MsgBox i & "=" & Paq(i)
Next
End Sub

Private Sub pctgrafico_Paint()
Dim i As Integer

For i = 0 To Tiempopaq - 1
  'pctgrafico.Line (i * 50, pctgrafico.Height - Paq(i) * 100)-((i + 1) * 50, pctgrafico.Height - Paq(i + 1) * 100), RGB(255, 0, 0)
  pctgrafico.Line (i * 50, Paq(i) * 100)-((i + 1) * 50, Paq(i + 1) * 100), RGB(255, 0, 0)
Next
End Sub
