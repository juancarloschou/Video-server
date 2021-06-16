VERSION 5.00
Begin VB.Form FormLista 
   Caption         =   "Examinar los Videos"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9390
   Begin VB.TextBox txtTV 
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtmus 
      Height          =   285
      Left            =   4440
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtdep 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.FileListBox FileTV 
      Height          =   3015
      Left            =   6360
      TabIndex        =   8
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdAbrirTV 
      Caption         =   "Abrir archivo con el reproductor"
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.FileListBox FileMus 
      Height          =   3015
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdAbrirMus 
      Caption         =   "Abrir archivo con el reproductor"
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbrirDep 
      Caption         =   "Abrir archivo con el reproductor"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.FileListBox FileDep 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "FILTRO:"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "FILTRO:"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "FILTRO:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Televisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Música"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Deportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "FormLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbrirDep_Click()
If FileDep.ListIndex <> -1 Then
    FormMedia.Show
    FormMedia.Media.FileName = FileDep.Path & "\" & FileDep.FileName
End If
End Sub

Private Sub cmdAbrirMus_Click()
If FileMus.ListIndex <> -1 Then
    FormMedia.Show
    FormMedia.Media.FileName = FileMus.Path & "\" & FileMus.FileName
End If
End Sub

Private Sub cmdAbrirTV_Click()
If FileTV.ListIndex <> -1 Then
    FormMedia.Show
    FormMedia.Media.FileName = FileTV.Path & "\" & FileTV.FileName
End If
End Sub

Private Sub cmdaceptar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Height = 6600
Width = 9510

FileDep.Path = Path & "\Deportes"
FileMus.Path = Path & "\Musica"
FileTV.Path = Path & "\TV"

If Dep Then
    cmdAbrirDep.Enabled = True
Else
    cmdAbrirDep.Enabled = False
End If

If Mus Then
    cmdAbrirMus.Enabled = True
Else
    cmdAbrirMus.Enabled = False
End If

If TV Then
    cmdAbrirTV.Enabled = True
Else
    cmdAbrirTV.Enabled = False
End If

End Sub

Private Function EsFiltro(filtro As String) As Boolean
Dim ok As Boolean

ok = True
If InStr(1, filtro, ".", vbTextCompare) = 0 Then
    ok = False
End If
If Len(filtro) < 3 Then
    ok = False
End If
EsFiltro = ok
End Function

Private Sub txtdep_Change()
If EsFiltro(txtdep) Then
    FileDep.Pattern = txtdep
End If
End Sub

Private Sub txtmus_Change()
If EsFiltro(txtmus) Then
    FileMus.Pattern = txtmus
End If
End Sub

Private Sub txtTV_Change()
If EsFiltro(txtTV) Then
    FileTV.Pattern = txtTV
End If
End Sub
