VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPostProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Private Type structVertex
  pos As D3DVECTOR
  tex0 As D3DVECTOR2
  tex1 As D3DVECTOR2
  tex2 As D3DVECTOR2
  tex3 As D3DVECTOR2
  tex4 As D3DVECTOR2
End Type

Private arrStream() As structVertex
Private vBuffer As Direct3DVertexBuffer8

Private Const vShader As Long = D3DFVF_XYZ Or D3DFVF_TEX5
Private Const vBytes As Long = 52


Public Function objCreate5Tap(Width As Long, Height As Long) As Boolean
 
  Static epsWidth As Single
  Static epsHeight As Single
  epsWidth = 0.5 / Width
  epsHeight = 0.5 / Height

  ReDim arrStream(0 To 5) As structVertex
  
  With arrStream(0)
    .pos.X = -1
    .pos.Y = 1
    .pos.Z = 0
    .tex0.X = 0 + epsWidth
    .tex0.Y = 0 + epsHeight
    .tex1.X = 0 + 2 * epsWidth
    .tex1.Y = 0 + 2 * epsHeight
    .tex2.X = 0
    .tex2.Y = 0 + 2 * epsHeight
    .tex3.X = 0
    .tex3.Y = 0
    .tex4.X = 0 + 2 * epsWidth
    .tex4.Y = 0
  End With
  
  With arrStream(1)
    .pos.X = 1
    .pos.Y = 1
    .pos.Z = 0
    .tex0.X = 1 + epsWidth
    .tex0.Y = 0 + epsHeight
    .tex1.X = 1 + 2 * epsWidth
    .tex1.Y = 0 + 2 * epsHeight
    .tex2.X = 1
    .tex2.Y = 0 + 2 * epsHeight
    .tex3.X = 1
    .tex3.Y = 0
    .tex4.X = 1 + 2 * epsWidth
    .tex4.Y = 0
  End With
  
  With arrStream(2)
    .pos.X = -1
    .pos.Y = -1
    .pos.Z = 0
    .tex0.X = 0 + epsWidth
    .tex0.Y = 1 + epsHeight
    .tex1.X = 0 + 2 * epsWidth
    .tex1.Y = 1 + 2 * epsHeight
    .tex2.X = 0
    .tex2.Y = 1 + 2 * epsHeight
    .tex3.X = 0
    .tex3.Y = 1
    .tex4.X = 0 + 2 * epsWidth
    .tex4.Y = 1
  End With
  
  arrStream(3) = arrStream(1)
  
  With arrStream(4)
    .pos.X = 1
    .pos.Y = -1
    .pos.Z = 0
    .tex0.X = 1 + epsWidth
    .tex0.Y = 1 + epsHeight
    .tex1.X = 1 + 2 * epsWidth
    .tex1.Y = 1 + 2 * epsHeight
    .tex2.X = 1
    .tex2.Y = 1 + 2 * epsHeight
    .tex3.X = 1
    .tex3.Y = 1
    .tex4.X = 1 + 2 * epsWidth
    .tex4.Y = 1
  End With
  
  arrStream(5) = arrStream(2)
  
  Set vBuffer = objD3DDev.CreateVertexBuffer(vBytes * 2 * 3, 0, vShader, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData vBuffer, 0, vBytes * 2 * 3, 0, arrStream(0)
  objCreate5Tap = True

End Function

    
Public Function objCreate() As Boolean

  ReDim arrStream(0 To 5) As structVertex
  
  With arrStream(0)
    .pos.X = -1
    .pos.Y = 1
    .pos.Z = 0
    .tex0.X = 0
    .tex0.Y = 0
    .tex1 = .tex0
  End With
  
  With arrStream(1)
    .pos.X = 1
    .pos.Y = 1
    .pos.Z = 0
    .tex0.X = 1
    .tex0.Y = 0
    .tex1 = .tex0
  End With
  
  With arrStream(2)
    .pos.X = -1
    .pos.Y = -1
    .pos.Z = 0
    .tex0.X = 0
    .tex0.Y = 1
    .tex1 = .tex0
  End With
  
  arrStream(3) = arrStream(1)
  
  With arrStream(4)
    .pos.X = 1
    .pos.Y = -1
    .pos.Z = 0
    .tex0.X = 1
    .tex0.Y = 1
    .tex1 = .tex0
  End With
  
  arrStream(5) = arrStream(2)
  
  Set vBuffer = objD3DDev.CreateVertexBuffer(vBytes * 2 * 3, 0, vShader, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData vBuffer, 0, vBytes * 2 * 3, 0, arrStream(0)
  objCreate = True

End Function


Public Function objCreateUser(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Boolean

  ReDim arrStream(0 To 5) As structVertex
  
  With arrStream(0)
    .pos.X = X1
    .pos.Y = Y1
    .pos.Z = 0
    .tex0.X = 0
    .tex0.Y = 0
    .tex1 = .tex0
  End With
  
  With arrStream(1)
    .pos.X = X2
    .pos.Y = Y1
    .pos.Z = 0
    .tex0.X = 1
    .tex0.Y = 0
    .tex1 = .tex0
  End With
  
  With arrStream(2)
    .pos.X = X1
    .pos.Y = Y2
    .pos.Z = 0
    .tex0.X = 0
    .tex0.Y = 1
    .tex1 = .tex0
  End With
  
  arrStream(3) = arrStream(1)
  
  With arrStream(4)
    .pos.X = X2
    .pos.Y = Y2
    .pos.Z = 0
    .tex0.X = 1
    .tex0.Y = 1
    .tex1 = .tex0
  End With
  
  arrStream(5) = arrStream(2)
  
  Set vBuffer = objD3DDev.CreateVertexBuffer(vBytes * 2 * 3, 0, vShader, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData vBuffer, 0, vBytes * 2 * 3, 0, arrStream(0)
  objCreateUser = True

End Function


Public Function objRender() As Boolean

  On Error Resume Next

  With objD3DDev
    
    Static matWorld As D3DMATRIX
    Static matView As D3DMATRIX
    Static matProjection As D3DMATRIX
    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity matView
    D3DXMatrixIdentity matProjection
    .SetTransform D3DTS_WORLD, matWorld
    .SetTransform D3DTS_VIEW, matView
    .SetTransform D3DTS_PROJECTION, matProjection

    .SetVertexShader vShader
    .SetStreamSource 0, vBuffer, vBytes
    .SetRenderState D3DRS_ZENABLE, 0
    .DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    .SetRenderState D3DRS_ZENABLE, 1
    
  End With
  
  If Not Err.Number = 0 Then
    Err.Clear
    objRender = False
  Else
    objRender = True
  End If
  
End Function


Public Function memClear() As Boolean

  Erase arrStream()
  Set vBuffer = Nothing
  
  memClear = True
  
End Function


