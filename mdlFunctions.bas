Attribute VB_Name = "mdlFunctions"

Option Explicit


Public Function mkVec3f(X As Single, Y As Single, Z As Single) As D3DVECTOR
  With mkVec3f
    .X = X
    .Y = Y
    .Z = Z
  End With
End Function


Public Function shCompile(fName As String) As Long

  On Error Resume Next
  shCompile = 0
  
  Static shArray() As Long
  Static shLength As Long
  Static shCode As D3DXBuffer

  Set shCode = objD3Dhlp.AssembleShaderFromFile(fName, 0, vbNullString, Nothing)
  shLength = shCode.GetBufferSize() / 4
  
  If Not Err.Number = 0 Then
    Err.Clear
    Set shCode = Nothing
    MsgBox "Could not assemble pixel shader: '" & fName & "'.", vbCritical Or vbOKOnly, "Error"
  Else
  
    ReDim shArray(shLength - 1) As Long
    objD3Dhlp.BufferGetData shCode, 0, 4, shLength, shArray(0)
    
    shCompile = objD3DDev.CreatePixelShader(shArray(0))
    
    If Not Err.Number = 0 Or shCompile = 0 Then
      Err.Clear
      Set shCode = Nothing
      shCompile = 0
      MsgBox "Pixel shader was sucessfully assembled, but failed to create." & vbCrLf & fName, vbCritical Or vbOKOnly, "Error"
    End If
  
  End If

End Function


Public Function txLoad(fName As String) As Direct3DTexture8

  On Error Resume Next

  Set txLoad = objD3Dhlp.CreateTextureFromFileEx(objD3DDev, fName, -1, -1, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
  If Not Err.Number = 0 Then
    Err.Clear
    Set txLoad = Nothing
    MsgBox "Failed to load texture map file: '" & fName, vbCritical Or vbOKOnly, "Error"
  End If

End Function
