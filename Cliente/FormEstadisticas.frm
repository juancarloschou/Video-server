VERSION 5.00
Begin VB.Form FormEstadisticas 
   Caption         =   "Estadisticas de la red"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   3750
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtancho 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtcalidad 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtperdidos 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtrecuperados 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtrecibidos 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ancho de banda"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Calidad recepcion"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Paquetes perdidos"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Paquetes recuperados"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Paquetes recibidos"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FormEstadisticas"
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
txtrecibidos = FormMedia.Media.ReceivedPackets
txtrecuperados = FormMedia.Media.RecoveredPackets
txtperdidos = FormMedia.Media.LostPackets
txtcalidad = FormMedia.Media.ReceptionQuality
txtancho = FormMedia.Media.Bandwidth
End Sub

