VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmservidor 
   BackColor       =   &H8000000C&
   Caption         =   "Servidor"
   ClientHeight    =   6930
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9705
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   480
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Clientes"
      Begin VB.Menu mnuclientes 
         Caption         =   "Lista de Clientes"
      End
      Begin VB.Menu mnuedicion 
         Caption         =   "Edición de Clientes"
      End
   End
   Begin VB.Menu mnusalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmservidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BD As ADODB.Connection 'Base de datos Videos.mdb de Access


Private Sub MDIForm_Load()
Dim i As Integer
    
Usuarios = 0
'Conectar Servidor
Winsock(0).LocalPort = 6000
Winsock(0).Listen
'Crea winsocks
For i = 1 To Maxusuarios
    Load Winsock(i)
Next
'Abre Base de datos
Set BD = New ADODB.Connection
BD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Videos.mdb;Mode=ReadWrite;Persist Security Info=False"
BD.Mode = adModeReadWrite
BD.Open
End Sub

Private Sub mnuclientes_Click()
frmListaClientes.Show
End Sub

Private Sub mnuedicion_Click()
frmEditar.Show
End Sub

Private Sub mnusalir_Click()
Dim i As Integer

For i = 0 To Maxusuarios 'desconecta winsocks
    If Winsock(i).State = sckConnected Then
        Winsock(i).Close
    End If
Next
BD.Close 'Cierra Base de datos
Unload Me
End Sub

Private Sub Procesar_orden(index As Integer, orden As String)
Dim Consulta As ADODB.Recordset 'Guarda resultados de las consultas
Dim separador As Integer
Dim i As Integer
Dim Nombre As String, Pass As String, Tarj As String, Bool As String, Tipos As String
Dim cons As String

'******************************************************
If Left(orden, 1) = "E" Then 'Entrar (usuario ya dado de alta). "E Nombre|Contraseña"
    
    separador = InStr(1, orden, "|", vbTextCompare)
    Set Consulta = BD.Execute("Select dep,mus,TV from clientes where (nombre='" & Mid(orden, 2, separador - 2) & "') and (pass='" & Mid(orden, separador + 1) & "')", , adCmdText)
    If Consulta.EOF Then
        Winsock(index).SendData Chr(2) & "NO" 'no existe el registro
        DoEvents
    Else
        Winsock(index).SendData Chr(2) & "YES" 'Acceso permitido
        DoEvents
        
        'Envia el path del servidor y tipos de videos comprados
        If Consulta!Dep Then
            Bool = "S"
        Else
            Bool = "N"
        End If
        If Consulta!Mus Then
            Bool = Bool & "S"
        Else
            Bool = Bool & "N"
        End If
        If Consulta!TV Then
            Bool = Bool & "S"
        Else
            Bool = Bool & "N"
        End If
        Winsock(index).SendData Chr(2) & "TIPOS" & Bool & App.Path
        DoEvents
    End If
End If

'******************************************************
If Left(orden, 1) = "A" Then 'Alta de usuario. "A Nombre|Contraseña|NºTarj|DEP MUS TV"

    separador = InStr(1, orden, "|", vbTextCompare)
    Nombre = Mid(orden, 2, separador - 2)
    
    i = separador + 1
    separador = InStr(i, orden, "|", vbTextCompare)
    Pass = Mid(orden, i, separador - i)
    
    i = separador + 1
    separador = InStr(i, orden, "|", vbTextCompare)
    Tarj = Mid(orden, i, separador - i)
    
    Set Consulta = BD.Execute("Select * from clientes where nombre='" & Nombre & "'", , adCmdText)
    If Not Consulta.EOF Then
        Winsock(index).SendData Chr(2) & "ALTANON" 'nombre repetido
        DoEvents
    Else
    
      Set Consulta = BD.Execute("Select * from clientes where tarjeta='" & Tarj & "'", , adCmdText)
      If Not Consulta.EOF Then
        Winsock(index).SendData Chr(2) & "ALTANOT" 'Nº Tarj repetido
        DoEvents
      Else
        'Hace el alta
        i = separador + 1
        If Mid(orden, i, 1) = "S" Then
            Bool = "',true,"
            Tipos = "S"
        Else
            Bool = "',false,"
            Tipos = "N"
        End If
        
        If Mid(orden, i + 1, 1) = "S" Then
          Bool = Bool & "true,"
          Tipos = Tipos & "S"
        Else
          Bool = Bool & "false,"
          Tipos = Tipos & "N"
        End If
        
        If Mid(orden, i + 2, 1) = "S" Then
          Bool = Bool & "true)"
          Tipos = Tipos & "S"
        Else
          Bool = Bool & "false)"
          Tipos = Tipos & "N"
        End If
        
        cons = "Insert into clientes (nombre,pass,tarjeta,dep,mus,TV) values('" & Nombre & "','" & Pass & "','" & Tarj & Bool
        'MsgBox cons
        BD.Execute cons, , adCmdText
    
        Winsock(index).SendData Chr(2) & "ALTASI" 'Se pudo hacer el alta
        DoEvents
      
        'Envia el path del servidor y tipos de videos comprados
        Winsock(index).SendData Chr(2) & "TIPOS" & Tipos & App.Path
        DoEvents
      End If
    End If
End If

End Sub

'Private Sub mnuvideos_Click()
'frmListaVideos.Show
'End Sub

Private Sub Winsock_ConnectionRequest(index As Integer, ByVal requestID As Long)
If Usuarios >= Maxusuarios Then
    MsgBox "No se puede conectar al servidor. Inténtelo más tarde"
Else
    Usuarios = Usuarios + 1
    Winsock(Usuarios).Close
    Winsock(Usuarios).Accept requestID
End If
End Sub

Private Sub Winsock_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim txt As String

Winsock(index).GetData txt
    
If Left(txt, 1) = Chr(2) Then
    Procesar_orden index, Mid(txt, 2)
End If
End Sub

Private Sub Winsock_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim i As Integer

MsgBox Description, vbCritical, "Error"
For i = 0 To Maxusuarios
  Winsock(i).Close
Next i
End Sub


