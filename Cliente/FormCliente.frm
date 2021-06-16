VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm FormCliente 
   BackColor       =   &H8000000C&
   Caption         =   "Cliente"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9675
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnualta 
         Caption         =   "Darse de Alta"
      End
      Begin VB.Menu mnuentrar 
         Caption         =   "Entrar"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnureproductor 
      Caption         =   "&Reproductor"
   End
   Begin VB.Menu mnuvideos 
      Caption         =   "&Videos del servidor"
      Begin VB.Menu mnulista 
         Caption         =   "Lista"
      End
   End
End
Attribute VB_Name = "FormCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Conectado As Boolean 'Si el winsock esta conectado

Private Sub MDIForm_Load()
Conectado = False
mnulista.Enabled = False
End Sub

Private Sub mnualta_Click()
If Not Conectado Then
    'Conectarse con el servidor
    Winsock.RemoteHost = "LocalHost"
    Winsock.RemotePort = 6000
    Winsock.Connect
    Conectado = True
End If
FormAlta.Show
End Sub

Private Sub mnuentrar_Click()
If Not Conectado Then
    'Conectarse con el servidor
    Winsock.RemoteHost = "LocalHost"
    Winsock.RemotePort = 6000
    Winsock.Connect
    Conectado = True
End If
FormEntrar.Show
End Sub

Private Sub mnulista_Click()
FormLista.Show
End Sub

Private Sub mnureproductor_Click()
FormMedia.Show
End Sub

Private Sub mnusalir_Click() 'Cierra el Winsock y la Base de datos
Winsock.Close
Unload Me
End Sub

Private Sub Procesar_orden(orden As String)

If orden = "YES" Then 'Entrar
    MsgBox "Datos comprobados", , "Puede entrar"
    mnualta.Enabled = False
    mnuentrar.Enabled = False
    mnulista.Enabled = True
End If

If orden = "NO" Then 'No Entrar
    MsgBox "Datos incorrectos", , "Permiso denegado"
End If

If orden = "ALTASI" Then 'Alta
    MsgBox "Alta completada", , "Puede entrar"
    mnualta.Enabled = False
    mnuentrar.Enabled = False
    mnulista.Enabled = True
End If

If orden = "ALTANON" Then 'No Alta por Nombre
    MsgBox "Introduzca otro nombre", , "Datos duplicados"
End If

If orden = "ALTANOT" Then 'No Alta por Nº Tarjeta credito
    MsgBox "Introduzca otro Nº de tarjeta de crédito", , "Datos duplicados"
End If

If Left(orden, 5) = "TIPOS" Then 'Tipos de videos comprados y path servidor
    If Mid(orden, 6, 1) = "S" Then
        Dep = True
    Else
        Dep = False
    End If
    
    If Mid(orden, 7, 1) = "S" Then
        Mus = True
    Else
        Mus = False
    End If
    
    If Mid(orden, 8, 1) = "S" Then
        TV = True
    Else
        TV = False
    End If

    Path = Mid(orden, 9)
End If

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim txt As String

Winsock.GetData txt

If Left(txt, 1) = Chr(2) Then
    Procesar_orden Mid(txt, 2)
End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, "Error"
Winsock.Close
End Sub



