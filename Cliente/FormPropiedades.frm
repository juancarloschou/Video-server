VERSION 5.00
Begin VB.Form FormPropiedades 
   Caption         =   "Propiedades del video"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Videoclip"
      Height          =   2895
      Left            =   4440
      TabIndex        =   13
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtcarchivo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtctitulo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtcautor 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtccopyright 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtcclasificacion 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtcdescripcion 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Titulo"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Autor"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Copyright"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Clasificación"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de reproducción"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtdescripcion 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtclasificacion 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtcopyright 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtautor 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txttitulo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtarchivo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Clasificación"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Copyright"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Autor"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Titulo"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormPropiedades"
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
'Datos de Playlist
txtarchivo = FormMedia.Media.GetMediaInfoString(mpShowFilename)
txttitulo = FormMedia.Media.GetMediaInfoString(mpShowTitle)
txtautor = FormMedia.Media.GetMediaInfoString(mpShowAuthor)
txtcopyright = FormMedia.Media.GetMediaInfoString(mpShowCopyright)
txtclasificacion = FormMedia.Media.GetMediaInfoString(mpShowRating)
txtdescripcion = FormMedia.Media.GetMediaInfoString(mpShowDescription)

'Datos del Clip. ESTO QUITADO SE COLGABA
txtcarchivo = FormMedia.Media.GetMediaInfoString(mpClipFilename)
'txtctitulo = FormMedia.Media.GetMediaInfoString(mpClipTitle)
'txtcautor = FormMedia.Media.GetMediaInfoString(mpClipAuthor)
'txtccopyright = FormMedia.Media.GetMediaInfoString(mpClipCopyright)
txtcclasificacion = FormMedia.Media.GetMediaInfoString(mpClipRating)
txtcdescripcion = FormMedia.Media.GetMediaInfoString(mpClipDescription)

End Sub
