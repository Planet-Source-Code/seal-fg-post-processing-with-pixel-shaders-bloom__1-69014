VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Public objTexture As Direct3DTexture8

Private objLastRT As Direct3DSurface8
Private objLastDepth As Direct3DSurface8
Private objThisRT As Direct3DSurface8
Private objThisDepth As Direct3DSurface8


Public Sub rtCreate(Width As Long, Height As Long)
  
  Set objTexture = objD3DDev.CreateTexture(Width, Height, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
  
  Set objThisRT = objTexture.GetSurfaceLevel(0)
  Set objThisDepth = objD3DDev.CreateDepthStencilSurface(Width, Height, D3DFMT_D24S8, D3DMULTISAMPLE_NONE)

End Sub


Public Sub rtAquire()
  
  Set objLastRT = objD3DDev.GetRenderTarget
  Set objLastDepth = objD3DDev.GetDepthStencilSurface

End Sub


Public Sub rtDestroy(TrueRT As Boolean)
  
  If TrueRT Then
    Set objLastRT = Nothing
    Set objLastDepth = Nothing
  End If
  Set objTexture = Nothing
  Set objThisRT = Nothing
  Set objThisDepth = Nothing
  
End Sub


Public Sub rtEnable(rtSwitch As Boolean)
  
  If rtSwitch Then
    objD3DDev.SetRenderTarget objThisRT, objThisDepth, 0
  Else
    objD3DDev.SetRenderTarget objLastRT, objLastDepth, 0
  End If
  
End Sub
