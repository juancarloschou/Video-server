VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormMedia 
   BackColor       =   &H00808080&
   Caption         =   "Reproductor de videos"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MediaPlayerCtl.MediaPlayer Media 
      Height          =   3765
      Left            =   240
      TabIndex        =   0
      Tag             =   "Manu"
      Top             =   240
      Width           =   4290
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   4194368
      DisplayForeColor=   49152
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -450
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuabrir 
         Caption         =   "Abrir..."
      End
      Begin VB.Menu mnucerrar 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu mnuguion 
         Caption         =   "-"
      End
      Begin VB.Menu mnuguardar 
         Caption         =   "Guardar como..."
      End
      Begin VB.Menu mnuguion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnupropiedades 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "&Ver"
      Begin VB.Menu mnuestandar 
         Caption         =   "Estandar"
      End
      Begin VB.Menu mnucompacta 
         Caption         =   "Compacta"
      End
      Begin VB.Menu mnuminima 
         Caption         =   "Minima"
      End
      Begin VB.Menu mnuguion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnupantalla 
         Caption         =   "Pantalla Completa"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuzoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnu50 
            Caption         =   "50%"
         End
         Begin VB.Menu mnu100 
            Caption         =   "100%"
         End
         Begin VB.Menu mnu200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuajustar 
            Caption         =   "Ajustar"
         End
      End
      Begin VB.Menu mnuguion2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuestadisticas 
         Caption         =   "Estadísticas"
      End
      Begin VB.Menu mnuvisible 
         Caption         =   "Siempre visible"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnureproducir 
      Caption         =   "&Reproducir"
      Begin VB.Menu mnuplay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnupause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuguion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvolumen 
         Caption         =   "Volumen"
         Begin VB.Menu mnusubir 
            Caption         =   "Subir"
         End
         Begin VB.Menu mnubajar 
            Caption         =   "Bajar"
         End
         Begin VB.Menu mnuguion4 
            Caption         =   "-"
         End
         Begin VB.Menu mnusilencio 
            Caption         =   "Silencio"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuir 
      Caption         =   "&Ir"
      Begin VB.Menu mnuguia 
         Caption         =   "Guía de medios"
      End
      Begin VB.Menu mnumusica 
         Caption         =   "Música"
      End
      Begin VB.Menu mnuRadio 
         Caption         =   "Radio"
      End
      Begin VB.Menu mnuwindows 
         Caption         =   "Windows media player"
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "&Ayuda"
      Index           =   1
      Begin VB.Menu mnuacerca 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu mnuayuda1 
         Caption         =   "Ayuda"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FormMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Videos ASF de ejemplo
'http://www.hyperdrive.co.uk/natsel/video/live/jbg_lo.asf
'http://www.jhepple.com/SampleMovies/niceday.asf
Option Explicit

'Declaración para usar ventanas siempre visibles
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declaración para abrir el internet explorer
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'SetWindowPos Flags:
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE


Private Sub Form_Load()
mnuvisible.Checked = False
mnusilencio.Checked = False
End Sub

Private Sub Form_Resize()
'Pone el fondo azul degradado
Dim i As Integer
Dim y As Integer

If WindowState <> vbMinimized Then 'no esta minimizado
    Cls
    AutoRedraw = True
    DrawStyle = 6
    DrawMode = 13
    DrawWidth = 2
    ScaleHeight = (256 * 2)
    y = 0
    For i = 0 To 255
      Line (0, y)-(Width, y + 2), RGB(0, 0, i), BF
      y = y + 2
    Next
End If

End Sub

Private Sub mnu100_Click()
Media.DisplaySize = mpDefaultSize
End Sub

Private Sub mnu200_Click()
Media.DisplaySize = mpDoubleSize
End Sub

Private Sub mnu50_Click()
Media.DisplaySize = mpHalfSize
End Sub

Private Sub mnuabrir_Click()
FormAbrir.Show
End Sub

Private Sub mnuacerca_Click()
Media.AboutBox
End Sub

Private Sub mnuajustar_Click()
Media.DisplaySize = mpFitToSize
'Ajusta al tamaño del control, no de la ventana
End Sub

Private Sub mnuayuda1_Click()
CommonDialog1.HelpFile = "mplayer2.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp 'Ayuda de Windows Media Player
End Sub

Private Sub mnubajar_Click()
Dim vol As Integer

vol = Media.Volume
vol = vol - 250
If vol < -5000 Then
    vol = -5000
End If
Media.Volume = vol
End Sub

Private Sub mnucerrar_Click()
Media.FileName = ""
End Sub

Private Sub mnucompacta_Click()
Media.ShowDisplay = False
Media.ShowStatusBar = True
Media.ShowTracker = True
End Sub

Private Sub mnuestadisticas_Click()
FormEstadisticas.Show
End Sub

Private Sub mnuestandar_Click()
Media.ShowDisplay = True
Media.ShowStatusBar = True
Media.ShowTracker = True
End Sub

Private Sub mnuguardar_Click()
Dim bData() As Byte  'Variable de datos
Dim iFile As Integer 'Variable FreeFile

CommonDialog1.ShowSave 'Eliges donde guardarlo
If CommonDialog1.FileName <> "" Then
    If Left(Media.FileName, 4) = "http" Or Left(Media.FileName, 3) = "ftp" Then
        'Baja el archivo de internet
        bData() = Inet1.OpenURL(Media.FileName, icByteArray) 'Baja el fichero
        iFile = FreeFile() 'Establece iFile a un archivo no utilizado
        Open CommonDialog1.FileName For Binary Access Write As #iFile
        Put #iFile, , bData() 'Llena el fich con datos
        Close #iFile
    Else 'Copia el archivo del disco o red
        FileCopy Media.FileName, CommonDialog1.FileName
    End If
    MsgBox "Archivo guardado"
Else
    MsgBox "Debes elegir un nombre y ruta para grabar el archivo"
End If
End Sub

Private Sub mnuguia_Click()
Dim i As Long
i = ShellExecute(hwnd, "Open", "http://windowsmedia.com/mg/home.asp", "", "", 1)
End Sub

Private Sub mnuminima_Click()
Media.ShowDisplay = False
Media.ShowStatusBar = False
Media.ShowTracker = False
End Sub

Private Sub mnumusica_Click()
Dim i As Long
i = ShellExecute(hwnd, "Open", "http://windowsmedia.com/mg/Music.asp", "", "", 1)
End Sub

Private Sub mnupantalla_Click()
Media.DisplaySize = mpFullScreen
End Sub

Private Sub mnupause_Click()
Media.Pause
End Sub

Private Sub mnuplay_Click()
Media.Play
End Sub

Private Sub mnupropiedades_Click()
FormPropiedades.Show
End Sub

Private Sub mnuRadio_Click()
Dim i As Long
i = ShellExecute(hwnd, "Open", "http://windowsmedia.msn.com/radiotuner/MyRadio.asp", "", "", 1)
End Sub

Private Sub mnusalir_Click()
Unload Me
End Sub

Private Sub mnusilencio_Click()
If Media.Mute Then
    mnusilencio.Checked = False
    Media.Mute = False
Else
    mnusilencio.Checked = True
    Media.Mute = True
End If
End Sub

Private Sub mnustop_Click()
Media.Stop
End Sub

Private Sub mnusubir_Click()
Dim vol As Integer

vol = Media.Volume
vol = vol + 250
If vol > 0 Then
  vol = 0
End If
Media.Volume = vol
End Sub

Private Sub mnuvisible_Click()
Dim i As Long

If mnuvisible.Checked Then
    mnuvisible.Checked = False
    'Pone ventana normal (-2)
    i = SetWindowPos(hwnd, -2, 0, 0, 0, 0, SWP_FLAGS)
Else
    mnuvisible.Checked = True
    'Pone ventana siempre visible (-1)
    i = SetWindowPos(hwnd, -1, 0, 0, 0, 0, SWP_FLAGS)
End If
End Sub

Private Sub mnuwindows_Click()
Dim i As Long

i = ShellExecute(hwnd, "Open", "http://www.microsoft.com/windows/windowsmedia/download/default.asp", "", "", 1)
End Sub
