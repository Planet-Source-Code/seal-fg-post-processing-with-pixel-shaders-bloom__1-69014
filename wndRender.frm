VERSION 5.00
Begin VB.Form wndRender 
   Caption         =   "Initializing..."
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   Icon            =   "wndRender.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   WindowState     =   2  'Maximized
   Begin VB.Timer tickFPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   4320
   End
End
Attribute VB_Name = "wndRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private iFPS As Long

Private lastX As Long
Private lastY As Long


Private Sub Form_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = Asc("P") Or KeyAscii = Asc("p") Then
  
    If shwPipeline = 0 Then
      shwPipeline = 1
    Else
      shwPipeline = 0
    End If
    
  End If
  
  If KeyAscii = Asc("H") Or KeyAscii = Asc("h") Then
  
    If shwHelp = 0 Then
      shwHelp = 1
    Else
      shwHelp = 0
    End If
    
  End If
  
  If KeyAscii = Asc("W") Or KeyAscii = Asc("w") Then
  
    If shwWalls = 0 Then
      shwWalls = 1
    Else
      shwWalls = 0
    End If
    
  End If
  
  If KeyAscii = Asc("R") Or KeyAscii = Asc("r") Then
  
    If effBright = 0 Then
      effBright = 1
    Else
      effBright = 0
    End If
    
  End If
  
  If KeyAscii = Asc("F") Or KeyAscii = Asc("f") Then
  
    If effFilter = 0 Then
      effFilter = 1
    Else
      effFilter = 0
    End If
    
  End If
  
  If KeyAscii = Asc("G") Or KeyAscii = Asc("g") Then
  
    If effGauss = 0 Then
      effGauss = 1
    Else
      effGauss = 0
    End If
    
  End If
  
  If KeyAscii = Asc("B") Or KeyAscii = Asc("b") Then
  
    If effBloom = 0 Then
      effBloom = 1
    Else
      effBloom = 0
    End If
    
  End If
  
  If KeyAscii = vbKeySpace Or KeyAscii = 13 Then Load wndSettings
  If KeyAscii = vbKeyEscape Then Unload Me
  
End Sub


Private Sub Form_Load()
  
  Show
  DoEvents
  
  Initialize
  
  iFPS = 0

  tickFPS.Enabled = True

  tickFPS_Timer

  Do
    Render
    iFPS = iFPS + 1
    DoEvents
  Loop

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = vbLeftButton Then
    camAlpha = camAlpha - (lastX - X) / 400
    camBeta = camBeta + (lastY - Y) / 400
    If camBeta >= Pi / 2 Then camBeta = Pi / 2 - 0.001
    If camBeta <= -Pi / 2 Then camBeta = -Pi / 2 + 0.001
  End If
  
  If Button = vbRightButton Then
    camDistance = camDistance + (lastY - Y) / (10000 - camDistance) * camDistance * 100
    If camDistance < 1 Then camDistance = 1
    If camDistance > 1000 Then camDistance = 1000
  End If
  
  If Button = vbRightButton + vbLeftButton Then
    camShift = camShift - (lastY - Y) / (100000 - camDistance) * camDistance * 100
    If camShift < -300 Then camShift = -300
    If camShift > 300 Then camShift = 300
  End If
  
  lastX = X
  lastY = Y

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Shutdown

End Sub


Private Sub tickFPS_Timer()

  Caption = "Rendering: " & ScaleWidth & " x " & ScaleHeight & " @ " & iFPS & " FPS - [Post-Processing: Bloom Effect]"
  iFPS = 0

End Sub
