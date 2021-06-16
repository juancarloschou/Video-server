VERSION 5.00
Begin VB.Form FormAlta 
   Caption         =   "Alta"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5070
   Begin VB.Frame Frame1 
      Caption         =   "Elige los tipos de videos que prefieras"
      Height          =   1815
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Width           =   4095
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkTV 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkmusica 
         Caption         =   "Check2"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkdeportes 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Total:"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "10 €"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "6 €"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "20 €"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "TV"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Música"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Deportes"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox txttarjeta 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nº Tarjeta crédito:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "FormAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub calculaTotal()
Dim total As Integer

total = 0
If chkdeportes.Value = 1 Then
    total = total + 20
End If
If chkmusica.Value = 1 Then
    total = total + 6
End If
If chkTV.Value = 1 Then
    total = total + 10
End If
txttotal = total
End Sub

Private Sub chkdeportes_Click()
calculaTotal
End Sub

Private Sub chkmusica_Click()
calculaTotal
End Sub

Private Sub chkTV_Click()
calculaTotal
End Sub

Private Sub cmdaceptar_Click()
Dim str As String

If Len(txtnombre) > 50 Or Len(txtnombre) = 0 Then
  MsgBox "Nombre incorrecto"
Else
  If Len(txtpass) > 8 Or Len(txtpass) = 0 Then
    MsgBox "Contraseña incorrecta"
  Else
    If Len(txttarjeta) <> 16 Or Not IsNumeric(txttarjeta) Then
      MsgBox "Número de tarjeta incorrecto"
    Else
      str = Chr(2) & "A" & txtnombre & "|" & txtpass & "|" & txttarjeta & "|"
      If chkdeportes.Value = 1 Then
        str = str & "S"
      Else
        str = str & "N"
      End If
      If chkmusica.Value = 1 Then
        str = str & "S"
      Else
        str = str & "N"
      End If
      If chkTV.Value = 1 Then
        str = str & "S"
      Else
        str = str & "N"
      End If
      
      FormCliente.Winsock.SendData str
      DoEvents 'Pide permiso para dar de alta
      Unload Me
    End If
  End If
End If
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Width = 5190
Height = 5535
End Sub
