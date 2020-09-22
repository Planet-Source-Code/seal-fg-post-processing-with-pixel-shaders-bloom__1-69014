VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 0


Private Type structVertex
  pos As D3DVECTOR
  tex As D3DVECTOR2
End Type

Private arrStream() As structVertex
Private vBuffer As Direct3DVertexBuffer8

Private Const vShader As Long = D3DFVF_XYZ Or D3DFVF_TEX1
Private Const vBytes As Long = 20


Private Type structFace
  iPos(0 To 2) As Long
  iTex(0 To 2) As Long
End Type


Public iVertex As Long
Private arrVertex() As D3DVECTOR

Public iTexture As Long
Private arrTexture() As D3DVECTOR2

Public iFace As Long
Private arrFace() As structFace


Public Function objLoad(fName As String) As Boolean

  On Error Resume Next
  Err.Clear
  objLoad = False

  If Len(Dir(fName)) > 0 Then
  
    Static fIndex As Long
    fIndex = FreeFile
    Open fName For Input As #fIndex
    If Not Err.Number = 0 Then
      MsgBox "appEffect.clsMesh.objLoad()::File was found, but failed to open.", vbExclamation Or vbOKOnly, "Error"
      Err.Clear
    Else
    
      memClear
      Static strData As String
      Do While Not EOF(fIndex)
        Line Input #fIndex, strData
      
        Select Case UCase(Left(strData, 2))
          
          Case "V" & Chr(32)
            If iVertex = 0 Then
              ReDim arrVertex(0) As D3DVECTOR
            Else
              ReDim Preserve arrVertex(iVertex) As D3DVECTOR
              If Not Err.Number = 0 Then
                MsgBox "appEffect.clsMesh.objLoad()::Failed to reallocate memory for vertex data.", vbExclamation Or vbOKOnly, "Error"
                Err.Clear
                Exit Function
              End If
            End If
            strData = Right(strData, Len(strData) - 2)
            Static posSpace1 As Long
            Static posSpace2 As Long
            posSpace1 = InStr(1, strData, Chr(32), vbBinaryCompare)
            posSpace2 = InStr(posSpace1 + 1, strData, Chr(32), vbBinaryCompare)
            With arrVertex(iVertex)
              .X = Val(Left(strData, posSpace1 - 1))
              .Y = Val(Mid(strData, posSpace1 + 1, posSpace2 - posSpace1 - 1))
              .Z = Val(Right(strData, Len(strData) - posSpace2))
            End With
            iVertex = iVertex + 1
            If Not Err.Number = 0 Then
              MsgBox "appEffect.clsMesh.objLoad()::Incorrect source vertex data format.", vbExclamation Or vbOKOnly, "Error"
              Err.Clear
              Exit Function
            End If
          
          Case "VT"
            If iTexture = 0 Then
              ReDim arrTexture(0) As D3DVECTOR2
            Else
              ReDim Preserve arrTexture(iTexture) As D3DVECTOR2
              If Not Err.Number = 0 Then
                MsgBox "appEffect.clsMesh.objLoad()::Failed to reallocate memory for texture coords data.", vbExclamation Or vbOKOnly, "Error"
                Err.Clear
                Exit Function
              End If
            End If
            strData = Right(strData, Len(strData) - 3)
            Static texSpace As Long
            texSpace = InStr(1, strData, Chr(32), vbBinaryCompare)
            With arrTexture(iTexture)
              .X = Val(Left(strData, texSpace - 1))
              .Y = Val(Right(strData, Len(strData) - texSpace))
            End With
            iTexture = iTexture + 1
            If Not Err.Number = 0 Then
              MsgBox "appEffect.clsMesh.objLoad()::Incorrect source texture coords data format.", vbExclamation Or vbOKOnly, "Error"
              Err.Clear
              Exit Function
            End If
          
          Case "F" & Chr(32)
            If iFace = 0 Then
              ReDim arrFace(0) As structFace
            Else
              ReDim Preserve arrFace(iFace) As structFace
              If Not Err.Number = 0 Then
                MsgBox "appEffect.clsMesh.objLoad()::Failed to reallocate memory for face data.", vbExclamation Or vbOKOnly, "Error"
                Err.Clear
                Exit Function
              End If
            End If
            strData = Right(strData, Len(strData) - 2)
            Static triSpace1 As Long
            Static triSpace2 As Long
            Static fDat As String
            Static fSpace As Long
            triSpace1 = InStr(1, strData, Chr(32), vbBinaryCompare)
            triSpace2 = InStr(triSpace1 + 1, strData, Chr(32), vbBinaryCompare)
            With arrFace(iFace)
              fDat = Left(strData, triSpace1 - 1)
              fSpace = InStr(1, fDat, "/", vbBinaryCompare)
              .iPos(0) = Val(Left(fDat, fSpace - 1)) - 1
              .iTex(0) = Val(Right(fDat, Len(fDat) - fSpace)) - 1
              fDat = Mid(strData, triSpace1 + 1, triSpace2 - triSpace1 - 1)
              fSpace = InStr(1, fDat, "/", vbBinaryCompare)
              .iPos(1) = Val(Left(fDat, fSpace - 1)) - 1
              .iTex(1) = Val(Right(fDat, Len(fDat) - fSpace)) - 1
              fDat = Right(strData, Len(strData) - triSpace2)
              fSpace = InStr(1, fDat, "/", vbBinaryCompare)
              .iPos(2) = Val(Left(fDat, fSpace - 1)) - 1
              .iTex(2) = Val(Right(fDat, Len(fDat) - fSpace)) - 1
            End With
            iFace = iFace + 1
            If Not Err.Number = 0 Then
              MsgBox "appEffect.clsMesh.objLoad()::Incorrect source face data format.", vbExclamation Or vbOKOnly, "Error"
              Err.Clear
              Exit Function
            End If
          
        End Select
        
      Loop
      Close #fIndex
      
      ReDim arrStream(iFace * 3 - 1) As structVertex
      If Not Err.Number = 0 Then
        MsgBox "appEffect.clsMesh.objLoad()::Can not allocate memory for vertex data stream.", vbExclamation Or vbOKOnly, "Error"
        Err.Clear
        Exit Function
      End If
      Static nFace As Long
      Static nVertex As Long
      Static nPoint As Long
      nVertex = 0
      For nFace = 0 To iFace - 1 Step 1
        For nPoint = 0 To 2 Step 1
          With arrStream(nVertex + nPoint)
            .pos = arrVertex(arrFace(nFace).iPos(nPoint))
            .tex = arrTexture(arrFace(nFace).iTex(nPoint))
            .tex.Y = 1 - .tex.Y
          End With
        Next nPoint
        nVertex = nVertex + 3
      Next nFace
      If Not Err.Number = 0 Then
        MsgBox "appEffect.clsMesh.objLoad()::Face data array has element(s) with out of range indexes.", vbExclamation Or vbOKOnly, "Error"
        Err.Clear
        Exit Function
      End If

      Set vBuffer = objD3DDev.CreateVertexBuffer(vBytes * iFace * 3, 0, vShader, D3DPOOL_DEFAULT)
      D3DVertexBuffer8SetData vBuffer, 0, vBytes * iFace * 3, 0, arrStream(0)
      If Not Err.Number = 0 Then
        MsgBox "appEffect.clsMesh.objLoad()::Error in creating D3DVertexBuffer.", vbExclamation Or vbOKOnly, "Error"
        Err.Clear
      Else
        objLoad = True
      End If
    
    End If
  
  Else
    MsgBox "appEffect.clsMesh.objLoad()::File was not found.", vbExclamation Or vbOKOnly, "Error"
  End If

End Function


Public Function objRender() As Boolean

  On Error Resume Next

  With objD3DDev
    If iFace > 0 Then
      .SetVertexShader vShader
      .SetStreamSource 0, vBuffer, vBytes
      .DrawPrimitive D3DPT_TRIANGLELIST, 0, iFace
    End If
  End With
  
  If Not Err.Number = 0 Then
    Err.Clear
    objRender = False
  Else
    objRender = True
  End If
  
End Function


Public Function memClear() As Boolean

  iVertex = 0
  Erase arrVertex()
  iTexture = 0
  Erase arrTexture()
  iFace = 0
  Erase arrFace()
  
  Erase arrStream()
  Set vBuffer = Nothing
  
  memClear = True
  
End Function

