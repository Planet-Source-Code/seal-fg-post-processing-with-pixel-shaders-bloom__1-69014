VERSION 5.00
Begin VB.Form wndSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effect Settings"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "wndSettigns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   Begin VB.CheckBox chkHelp 
      Caption         =   "Display Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkPipeline 
      Caption         =   "Display Bloom Rendering Pipeline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CheckBox chkWalls 
      Caption         =   "Render Walls Mesh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox chkBright 
      Caption         =   "Bright Pass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox chkBloom 
      Caption         =   "Entrie Bloom Effect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkGauss 
      Caption         =   "Gaussiasn Blur Pass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "x16 Anisotropic Filtering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox lstBuffer 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "wndSettigns.frx":038A
      Left            =   2160
      List            =   "wndSettigns.frx":03A6
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label infBuffer 
      AutoSize        =   -1  'True
      Caption         =   "Post-Processing Buffer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1965
   End
End
Attribute VB_Name = "wndSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnCancel_Click()
  
  Unload Me
  
End Sub


Private Sub btnOk_Click()
 
    Select Case lstBuffer.ListIndex
      Case 0
        effSampling = 2
      Case 1
        effSampling = 4
      Case 2
        effSampling = 6
      Case 3
        effSampling = 8
      Case 4
        effSampling = 10
      Case 5
        effSampling = 12
      Case 6
        effSampling = 14
      Case 7
        effSampling = 16
    End Select
    effFilter = chkFilter.Value
    effGauss = chkGauss.Value
    effBright = chkBright.Value
    effBloom = chkBloom.Value
    shwWalls = chkWalls.Value
    shwPipeline = chkPipeline.Value
    shwHelp = chkHelp.Value
 
    rtBrightPass.rtDestroy False
    rtGaussianBlur.rtDestroy False
    
    rtBrightPass.rtCreate Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)
    rtGaussianBlur.rtCreate Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)
    
    ppGaussBlur.memClear
    ppGaussBlur.objCreate5Tap Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)
    
    Unload Me

End Sub


Private Sub Form_Load()
  
  Move wndRender.Left + wndRender.Width - 1000 - wndSettings.Width, wndRender.Top + wndRender.Height - 1000 - wndSettings.Height
 
  lstBuffer.Text = "x1/" & effSampling & " Downsampling"
  chkFilter.Value = effFilter
  chkGauss.Value = effGauss
  chkBright.Value = effBright
  chkBloom.Value = effBloom
  chkWalls.Value = shwWalls
  chkPipeline.Value = shwPipeline
  chkHelp.Value = shwHelp
 
  Show
  
End Sub

