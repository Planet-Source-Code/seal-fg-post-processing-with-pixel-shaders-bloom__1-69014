Attribute VB_Name = "mdlMain"

Option Explicit


Public Const Pi As Single = 3.14159265358979

Public confDevice As D3DPRESENT_PARAMETERS

Private objDX As DirectX8
Private objD3D As Direct3D8

Public objD3DDev As Direct3DDevice8
Public objD3Dhlp As D3DX8

Private txHelp As Direct3DTexture8
Private txOrig As Direct3DTexture8
Private txPass0 As Direct3DTexture8
Private txPass1 As Direct3DTexture8

Private mhWalls As clsMesh
Private txWalls As Direct3DTexture8

Private mhStatue As clsMesh
Private txStatue As Direct3DTexture8

Public camAlpha As Single
Public camBeta As Single
Public camDistance As Single
Public camShift As Single

Public psBrightPass As Long
Public psGaussianBlur As Long
Public psTextureBlend As Long

Private rtOriginalImage As clsRenderTarget
Public rtBrightPass As clsRenderTarget
Public rtGaussianBlur As clsRenderTarget

Private ppBrightPass As clsPostProcessing
Public ppGaussBlur As clsPostProcessing
Private ppFinalBlend As clsPostProcessing

Private ppPipelineView As clsPostProcessing

Public effSampling As Long
Public effFilter As Long
Public effGauss As Long
Public effBright As Long
Public effBloom As Long

Public shwHelp As Long
Public shwWalls As Long
Public shwPipeline As Long

Public Sub Initialize()

  On Error Resume Next

  Set objDX = New DirectX8
  Set objD3D = objDX.Direct3DCreate
  Set objD3Dhlp = New D3DX8
  
  Static confDisplay As D3DDISPLAYMODE
  objD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, confDisplay
  
  With confDevice
    .AutoDepthStencilFormat = D3DFMT_D24S8
    .BackBufferCount = 1
    .BackBufferFormat = confDisplay.Format
    .BackBufferHeight = wndRender.ScaleHeight
    .BackBufferWidth = wndRender.ScaleWidth
    .EnableAutoDepthStencil = 1
    .flags = 0
    .FullScreen_PresentationInterval = 0
    .FullScreen_RefreshRateInHz = 0
    .hDeviceWindow = wndRender.hWnd
    .MultiSampleType = D3DMULTISAMPLE_NONE
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    .Windowed = 1
  End With

  Set objD3DDev = objD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, confDevice.hDeviceWindow, D3DCREATE_HARDWARE_VERTEXPROCESSING, confDevice)
  If Not Err.Number = 0 Then
    Err.Clear
    MsgBox "Failed to create Direct3DDevice8. Application will now quit.", vbCritical Or vbOKOnly, "Error"
    Shutdown
  End If

  
  camDistance = 150
  camAlpha = 35 * Pi / 180
  camBeta = 15 * Pi / 180
  camShift = 50
  
  
  effSampling = 6
  effFilter = 1
  effGauss = 1
  effBright = 1
  effBloom = 1
  
  shwHelp = 1
  shwWalls = 1
  shwPipeline = 1
  
  
  psBrightPass = shCompile(App.Path & "\psh_BrightPass_ps.1.4.txt")
  psGaussianBlur = shCompile(App.Path & "\psh_GaussianBlur_ps.1.4.txt")
  psTextureBlend = shCompile(App.Path & "\psh_TextureBlend_ps.1.1.txt")


  Set txHelp = txLoad(App.Path & "\texHelp.png")
  Set txOrig = txLoad(App.Path & "\texOriginal.png")
  Set txPass0 = txLoad(App.Path & "\texPass0.png")
  Set txPass1 = txLoad(App.Path & "\texPass1.png")
  Set txWalls = txLoad(App.Path & "\texWalls.png")
  Set txStatue = txLoad(App.Path & "\texStatue.png")


  Set mhWalls = New clsMesh
  If Not mhWalls.objLoad(App.Path & "\objWalls.obj") Then
    mhWalls.memClear
    MsgBox "Failed to load mesh file: '" & App.Path & "\objWalls.obj" & "'.", vbCritical Or vbOKOnly, "Error"
  End If

  Set mhStatue = New clsMesh
  If Not mhStatue.objLoad(App.Path & "\objStatue.obj") Then
    mhStatue.memClear
    MsgBox "Failed to load mesh file: '" & App.Path & "\objStatue.obj" & "'.", vbCritical Or vbOKOnly, "Error"
  End If


  Set rtOriginalImage = New clsRenderTarget
  rtOriginalImage.rtAquire
  rtOriginalImage.rtCreate confDevice.BackBufferWidth, confDevice.BackBufferHeight
  
  Set rtBrightPass = New clsRenderTarget
  rtBrightPass.rtAquire
  rtBrightPass.rtCreate Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)

  Set rtGaussianBlur = New clsRenderTarget
  rtGaussianBlur.rtAquire
  rtGaussianBlur.rtCreate Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)

  Set ppBrightPass = New clsPostProcessing
  ppBrightPass.objCreate
  
  Set ppGaussBlur = New clsPostProcessing
  ppGaussBlur.objCreate5Tap Int(confDevice.BackBufferWidth / effSampling), Int(confDevice.BackBufferHeight / effSampling)

  Set ppFinalBlend = New clsPostProcessing
  ppFinalBlend.objCreate

  Set ppPipelineView = New clsPostProcessing

End Sub


Public Sub Render()

  On Error Resume Next

  With objD3DDev
    
    
    'pass0: render scene into full-size rt texture
    If effBloom = 1 Then rtOriginalImage.rtEnable True
    
    .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF3F3F3F, 1, 0
    .BeginScene
    
    
    Static camX As Single
    Static camY As Single
    Static camZ As Single
    camX = Sin(camAlpha) * Cos(camBeta) * camDistance
    camY = Sin(camBeta) * camDistance
    camZ = Cos(camAlpha) * Cos(camBeta) * camDistance
    
    
    Static matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, mkVec3f(camX, camY + camShift, camZ), mkVec3f(0, 0 + camShift, 0), mkVec3f(0, 1, 0)
    .SetTransform D3DTS_VIEW, matView
    
    Static matProjection As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProjection, 1, confDevice.BackBufferHeight / confDevice.BackBufferWidth, 1, 100000
    .SetTransform D3DTS_PROJECTION, matProjection
    
    
    Static iMap As Long
    For iMap = 0 To 4 Step 1
      .SetTextureStageState iMap, D3DTSS_TEXCOORDINDEX, iMap
      If effFilter = 1 Then
        .SetTextureStageState iMap, D3DTSS_MAXANISOTROPY, 16
        .SetTextureStageState iMap, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        .SetTextureStageState iMap, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        .SetTextureStageState iMap, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
      Else
        .SetTextureStageState iMap, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState iMap, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState iMap, D3DTSS_MIPFILTER, D3DTEXF_POINT
      End If
      .SetTextureStageState iMap, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP
      .SetTextureStageState iMap, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP
    Next iMap
    
    
    .SetRenderState D3DRS_LIGHTING, 0
    .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
  
    
    .SetPixelShader 0
  
  
    If shwWalls = 1 Then
      .SetTexture 0, txWalls
      If Not mhWalls.objRender Then
        MsgBox "Error rendering 'Walls' mesh. Application will now quit.", vbCritical Or vbOKOnly, "Error"
        Shutdown
      End If
    End If
  
    .SetTexture 0, txStatue
    If Not mhStatue.objRender Then
      MsgBox "Error rendering 'Statue' mesh. Application will now quit.", vbCritical Or vbOKOnly, "Error"
      Shutdown
    End If
  
  
    .EndScene
    
    If effBloom = 1 Then
      
      rtOriginalImage.rtEnable False
    
    
      'pass1: render original image into lower resolution rt texture with bright pass shader
      If effBright = 1 Then
        rtBrightPass.rtEnable True
        .BeginScene
        .SetPixelShader psBrightPass
        .SetTexture 0, rtOriginalImage.objTexture
        If Not ppBrightPass.objRender Then
          MsgBox "Error rendering 'Bright Pass' post-processing effect. Application will now quit.", vbCritical Or vbOKOnly, "Error"
          Shutdown
        End If
        .EndScene
        rtBrightPass.rtEnable False
      End If
    
    
      If effGauss = 1 Then
        
        'pass2: perform a gaussian blur on bright passed (or original - when b.pass disabled) rt texture
        'using 5 texture with UV coord shifting (.objCreate5Tap)
        rtGaussianBlur.rtEnable True
        .BeginScene
        .SetPixelShader psGaussianBlur
        For iMap = 0 To 4 Step 1
          If effBright = 1 Then
            .SetTexture iMap, rtBrightPass.objTexture
          Else
            .SetTexture iMap, rtOriginalImage.objTexture
          End If
        Next iMap
        If Not ppGaussBlur.objRender Then
          MsgBox "Error rendering 'Gaussian Blur' post-processing effect. Application will now quit.", vbCritical Or vbOKOnly, "Error"
          Shutdown
        End If
        .EndScene
        rtGaussianBlur.rtEnable False
        
      End If
    
    
      'pass3: blend original image with stretched back to full size gaussian output texture
      .BeginScene
      .SetPixelShader psTextureBlend
      If effGauss = 1 Then
        .SetTexture 0, rtGaussianBlur.objTexture
      Else
        If effBright = 1 Then
          .SetTexture 0, rtBrightPass.objTexture
        Else
          .SetTexture 0, rtOriginalImage.objTexture
        End If
      End If
      .SetTexture 1, rtOriginalImage.objTexture
      If Not ppFinalBlend.objRender Then
        MsgBox "Error rendering 'Final Blend' post-processing effect. Application will now quit.", vbCritical Or vbOKOnly, "Error"
        Shutdown
      End If
      
      
      If shwPipeline = 1 Then
        
        ppPipelineView.memClear
        ppPipelineView.objCreateUser -0.95, 0.95 - 0.35 * 0, -0.65, 0.65 - 0.35 * 0
        .SetPixelShader psTextureBlend
        .SetTexture 1, txOrig
        .SetTexture 0, rtOriginalImage.objTexture
        ppPipelineView.objRender
        
        If effBright = 1 Then
          ppPipelineView.memClear
          ppPipelineView.objCreateUser -0.95, 0.95 - 0.35 * 1, -0.65, 0.65 - 0.35 * 1
          .SetPixelShader psTextureBlend
          .SetTexture 1, txPass0
          .SetTexture 0, rtBrightPass.objTexture
          ppPipelineView.objRender
        End If
      
        If effGauss = 1 Then
          ppPipelineView.memClear
          If effBright = 1 Then
            ppPipelineView.objCreateUser -0.95, 0.95 - 0.35 * 2, -0.65, 0.65 - 0.35 * 2
          Else
            ppPipelineView.objCreateUser -0.95, 0.95 - 0.35 * 1, -0.65, 0.65 - 0.35 * 1
          End If
          .SetPixelShader psTextureBlend
          .SetTexture 1, txPass1
          .SetTexture 0, rtGaussianBlur.objTexture
          ppPipelineView.objRender
        End If
      
      End If
      
      
      .EndScene
    
    
    End If
    
    .SetTexture 1, Nothing
      
    If shwHelp = 1 Then
      ppPipelineView.memClear
      ppPipelineView.objCreateUser 0.3, -0.5, 1, -1
      .SetPixelShader 0
      .SetTexture 0, txHelp
      .SetRenderState D3DRS_ALPHABLENDENABLE, 1
      .SetRenderState D3DRS_SRCBLEND, 1
      .SetRenderState D3DRS_DESTBLEND, 3
      ppPipelineView.objRender
      .SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If
    
    .SetTexture 0, Nothing
    
    
    If Not .TestCooperativeLevel = 0 Then
      MsgBox "Cooperative level lost. Application will now quit.", vbCritical Or vbOKOnly, "Error"
      Shutdown
    Else
      .Present ByVal 0, ByVal 0, 0, ByVal 0
    End If
  End With


  If Not Err.Number = 0 Then
    Err.Clear
    MsgBox "Unexpected error occured in rendering pipeline. Application will now quit.", vbCritical Or vbOKOnly, "Error"
    Shutdown
  End If


End Sub


Public Sub Shutdown()

  On Error Resume Next

  Set txPass0 = Nothing
  Set txPass1 = Nothing
  Set txOrig = Nothing
  Set txHelp = Nothing

  ppPipelineView.memClear
  Set ppPipelineView = Nothing
  
  ppBrightPass.memClear
  Set ppBrightPass = Nothing
  
  ppGaussBlur.memClear
  Set ppGaussBlur = Nothing
  
  ppFinalBlend.memClear
  Set ppFinalBlend = Nothing

  rtOriginalImage.rtDestroy True
  rtBrightPass.rtDestroy True
  rtGaussianBlur.rtDestroy True
  
  Set rtOriginalImage = Nothing
  Set rtBrightPass = Nothing
  Set rtGaussianBlur = Nothing

  Set txWalls = Nothing

  mhWalls.memClear
  Set mhWalls = Nothing

  Set txStatue = Nothing

  mhStatue.memClear
  Set mhStatue = Nothing
  
  objD3DDev.DeletePixelShader psBrightPass
  objD3DDev.DeletePixelShader psGaussianBlur
  objD3DDev.DeletePixelShader psTextureBlend

  Set objD3DDev = Nothing
  Set objD3D = Nothing
  Set objDX = Nothing
  
  If Not Err.Number = 0 Then Err.Clear
  End

End Sub

