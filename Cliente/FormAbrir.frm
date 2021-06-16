VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormAbrir 
   Caption         =   "Abrir"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   ScaleHeight     =   2805
   ScaleWidth      =   4485
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexaminar 
      Caption         =   "Examinar..."
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtarchivo 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Abrir:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba el nombre de un archivo de audio o de una película (en Internet o en el equipo) y el reproductor lo abrirá."
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "FormAbrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
If txtarchivo = "" Then
  MsgBox "Falta el nombre del archivo"
Else
  FormMedia.Media.FileName = txtarchivo 'Abre el arch (si puede)
  Unload Me
End If
End Sub

Private Sub cmdcancelar_Click()
FormMedia.Show
Unload Me
End Sub

Private Sub cmdexaminar_Click()
'Pone los filtros para elegir tipos de archivos por extensiones
CommonDialog1.Filter = "Videos ASF|*.ASF|Videos AVI|*.AVI|Videos MPG|*.MPG;*.MPEG|" & _
  "Otros Videos (MOV,VOD,RA,RV)|*.MOV;*.VOD;*.RA;*.RV|Sonido (WAV,SND,AU,MID)|*.WAV;*.SND;*.AU;*.MID;*.MIDI|Todos los archivos|*.*"
CommonDialog1.ShowOpen 'Elige arch para abrir
txtarchivo = CommonDialog1.FileName
End Sub
