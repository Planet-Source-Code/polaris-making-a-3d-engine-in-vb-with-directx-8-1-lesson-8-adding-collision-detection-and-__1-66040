VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_SkyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name:
'     Developer/Programmer: Johna
'    ----------------------------------------------------------
'     Module Name: cQUEST3D_SkyBox
'     Module File: Johna_SkyBOX.cls
'     Module Type: Class
'     Module Description:
'
'  THIS MODULE CONTAINS VERY USEFUL AND HARD HAND CODED METHODS, IF YOU
'   WANT TO USE THEM, GIVE CREDITS TO MY PERSON (Polaris),johna_pop@yahoo.fr
'
'
'    ----------------------------------------------------------
'     © johna 2002 for information and authaurization johna_pop@yahoo.fr
'    ----------------------------------------------------------
'   ============================================================

Option Explicit
Public Enum QUEST3D_SkyBOXPosition
    QUEST3D_SKY_LEFT = &H0
    QUEST3D_SKY_RIGHT = &H1
    QUEST3D_SKY_TOP = &H2
    QUEST3D_SKY_DOWN = &H3
    QUEST3D_SKY_BACK = &H4
    QUEST3D_SKY_FRONT = &H5

End Enum

Private Type LITVERTEX
    Position As D3DVECTOR

    Tu As Single            'texture map coordinate
    TV As Single            'texture map coordinate
End Type

' custom FVF, which describes our custom vertex structure
Private Const Lit_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_TEX1)

Dim SKY_TEX() As Long
Private SkyScale As D3DVECTOR

'Private Type Tblend
'    SRC_BLEND As CONST_D3DBLEND
'    DEST_BLEND As CONST_D3DBLEND
'End Type

Private Blend As QUEST3D_TEXTURE_BLEND

Private Type Tsky_Layer
    HI As Single
    TEXT_ID As Long
    Rotation As D3DVECTOR
    Trans As D3DVECTOR
    Speed As Long
    Size As D3DVECTOR2

    Blend As QUEST3D_TEXTURE_BLEND
    VertexBUFF As Direct3DVertexBuffer8
    Layed_ID As Long

End Type

'for vertex
Private Type PARTICLEVERTEX
    V As D3DVECTOR

    color As Long

    'specular As Long
    Specular As Long
    TEX As D3DVECTOR2
End Type

Private Const FVF_PARTICLEVERTEX = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1
'(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Private LAYER_VERT(4) As PARTICLEVERTEX
Private LAYER_VERTID(4) As PARTICLEVERTEX

Private NumLAYA As Long
Private LAYA() As Tsky_Layer

Private Type VerTEX_PARAM
    vPosition As D3DVECTOR
    Vscal As D3DVECTOR
    Vrotate As D3DVECTOR
    VertexBUFF(5) As Direct3DVertexBuffer8
    SKY_FRONT As Long
    SKY_BACK As Long
    SKY_LEFT As Long
    SKY_RIGHT As Long
    SKY_Down As Long
    SKY_TOP As Long
    LAYER_VERT(3) As LITVERTEX
    'CUBE As cQUEST3D_RenderToSurface
    IS_CUBE As Boolean
End Type

Private SKY As VerTEX_PARAM

Private matView As D3DMATRIX, matViewSave As D3DMATRIX, hr As Long
Private matWorld As D3DMATRIX
'private v As LITVERTEX
'private 20 As Long
'private i As Integer
Private Clamp1 As Long
Private Clamp2 As Long

'----------------------------------------
'Name: Add_SkyBOX
'Object: Add
'Event: SkyBOX
'Description:
'----------------------------------------
Sub Add_SkyBOX(Position As QUEST3D_SkyBOXPosition, TextureFile As String)

    Add_SkyBOXEX Position, obj_D3DX.CreateTextureFromFileEx(obj_Device, TextureFile, 0, 0, MIP_LEVELS, 0, CFG.BK_FMT, D3DPOOL_MANAGED, Data.TEXTURE_FILTER, Data.TEXTURE_MIPMAPFILTER, 0, ByVal 0, ByVal 0)

End Sub

'Function Add_SkyBoxFromCubeTexture(TextureFile As String) As Boolean
'
'  If FileiS_valid(TextureFile) Then
'   Set Sky.CUBE = New cQUEST3D_RenderToSurface
'  Add_SkyBoxFromCubeTexture = Sky.CUBE.Init_CreateCubeMapFromFile(TextureFile)
'  Sky.IS_CUBE = Add_SkyBoxFromCubeTexture
'
'
'
'
'
'
'
'   Dim V(0 To 3) As LITVERTEX
'  Dim VertexSizeInBytes As Long
'
'    VertexSizeInBytes = Len(V(0))
'
'    ' The front of the sky
'    V(0).Position = Vector(-1#, -1#, 1#)
'    V(1).Position = Vector(-1#, 1#, 1#)
'    V(2).Position = Vector(1#, -1#, 1#)
'    V(3).Position = Vector(1#, 1#, 1#)
'    V(0).Tu = 0
'    V(0).TV = 0
'    V(1).Tu = 0
'    V(1).TV = 0
'    V(2).Tu = 0
'    V(2).TV = 0
'    V(3).Tu = 0
'    V(3).TV = 0
'
'    Set Sky.VertexBUFF(0) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(0), 0, VertexSizeInBytes * 4, 0, V(0)
'
'    ' The back of the sky
'    V(0).Position = Vector(1#, -1#, -1#)
'    V(1).Position = Vector(1#, 1#, -1#)
'    V(2).Position = Vector(-1#, -1#, -1#)
'    V(3).Position = Vector(-1#, 1#, -1#)
'    V(0).Tu = 0
'    V(0).TV = 0
'    V(1).Tu = 0
'    V(1).TV = 0
'    V(2).Tu = 0
'    V(2).TV = 0
'    V(3).Tu = 0
'    V(3).TV = 0
'
'    'TOTAL_MEM_USED = TOTAL_MEM_USED + VertexSizeInBytes * 4 * 6
'
'    Set Sky.VertexBUFF(1) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(1), 0, VertexSizeInBytes * 4, 0, V(0)
'
'    ' The left of the sky
'    V(0).Position = Vector(-1#, -1#, -1#)
'    V(1).Position = Vector(-1#, 1#, -1#)
'    V(2).Position = Vector(-1#, -1#, 1#)
'    V(3).Position = Vector(-1#, 1#, 1#)
'    V(0).Tu = 0
'    V(0).TV = 0
'    V(1).Tu = 0
'    V(1).TV = 0
'    V(2).Tu = 0
'    V(2).TV = 0
'    V(3).Tu = 0
'    V(3).TV = 0
'
'    Set Sky.VertexBUFF(2) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(2), 0, VertexSizeInBytes * 4, 0, V(0)
'
'    ' The right of the sky
'    V(0).Position = Vector(1#, -1#, 1#)
'    V(1).Position = Vector(1#, 1#, 1#)
'    V(2).Position = Vector(1#, -1#, -1#)
'    V(3).Position = Vector(1#, 1#, -1#)
'    V(0).Tu = 0
'    V(0).TV = 0
'    V(1).Tu = 0
'    V(1).TV = 0
'    V(2).Tu = 0
'    V(2).TV = 0
'    V(3).Tu = 0
'    V(3).TV = 0
'
'    Set Sky.VertexBUFF(3) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(3), 0, VertexSizeInBytes * 4, 0, V(0)
'
'    ' The top of the sky
'    V(0).Position = Vector(-1#, 1#, 1#)
'    V(1).Position = Vector(-1#, 1#, -1#)
'    V(2).Position = Vector(1#, 1#, 1#)
'    V(3).Position = Vector(1#, 1#, -1#)
'    'V(0).Tu = 0: V(0).TV =0
'    'V(1).Tu = 0: V(1).TV = 0
'    'V(2).Tu =0: V(2).TV =0
'    'V(3).Tu =0: V(3).TV = 0
'
'    V(0).TV = 0
'    V(0).Tu = 0
'    V(1).TV = 0
'    V(1).Tu = 0
'    V(2).TV = 0
'    V(2).Tu = 0
'    V(3).TV = 0
'    V(3).Tu = 0
'
'    CopyMemory Sky.LAYER_VERT(0), V(0), VertexSizeInBytes * 4
'
'    Set Sky.VertexBUFF(4) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(4), 0, VertexSizeInBytes * 4, 0, V(0)
'
'    ' The bottom of the sky
'    V(0).Position = Vector(-1#, -1#, -1#)
'    V(1).Position = Vector(-1#, -1#, 1#)
'    V(2).Position = Vector(1#, -1#, -1#)
'    V(3).Position = Vector(1#, -1#, 1#)
'    'V(0).Tu = 0: V(0).TV =0
'    'V(1).Tu = 0: V(1).TV = 0
'    'V(2).Tu =0: V(2).TV =0
'    'V(3).Tu =0: V(3).TV = 0
'
'    V(0).TV = 0
'    V(0).Tu = 0
'    V(1).TV = 0
'    V(1).Tu = 0
'    V(2).TV = 0
'    V(2).Tu = 0
'    V(3).TV = 0
'    V(3).Tu = 0
'
'    Set Sky.VertexBUFF(5) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
    '        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    D3DVertexBuffer8SetData Sky.VertexBUFF(5), 0, VertexSizeInBytes * 4, 0, V(0)
'
'
'
'  End If
'
'End Function
'
'----------------------------------------
'Name: Add_SkyBOXEX
'Object: Add
'Event: SkyBOXEX
'Description:
'----------------------------------------
Sub Add_SkyBOXEX(Position As QUEST3D_SkyBOXPosition, Texture As Direct3DBaseTexture8)

  Dim TID As Long

    TID = Add_TextureToPoolEX(Texture)

    If Position = QUEST3D_SKY_FRONT Then _
       SKY.SKY_FRONT = TID

    If Position = QUEST3D_SKY_BACK Then _
       SKY.SKY_BACK = TID

    If Position = QUEST3D_SKY_LEFT Then _
       SKY.SKY_LEFT = TID

    If Position = QUEST3D_SKY_RIGHT Then _
       SKY.SKY_RIGHT = TID

    If Position = QUEST3D_SKY_DOWN Then _
       SKY.SKY_Down = TID

    If Position = QUEST3D_SKY_TOP Then _
       SKY.SKY_TOP = TID

End Sub

'Dim PIX As New cQUEST3D_Shader
'Dim Ver As New cQUEST3D_Shader
'
'

Sub Add_SkyLayer(TextureFile As String, Optional ByVal srcBlend As CONST_D3DBLEND = D3DBLEND_SRCALPHA, Optional ByVal DestBlend As CONST_D3DBLEND = D3DBLEND_INVSRCALPHA)

  '
  '    MySprite.Blend.SRC_BLEND = srcBlend
  '    MySprite.Blend.DEST_BLEND = DestBlend
  '

  Dim V(0 To 3) As LITVERTEX
  Dim VertexSizeInBytes As Long

    VertexSizeInBytes = Len(V(0))

    NumLAYA = NumLAYA + 1
    ReDim Preserve LAYA(NumLAYA - 1)

    LAYA(NumLAYA - 1).Blend = Blend
    LAYA(NumLAYA - 1).HI = 600
    LAYA(NumLAYA - 1).Size = Vector2D(8000, 8000)
    LAYA(NumLAYA - 1).Layed_ID = NumLAYA

    LAYA(NumLAYA - 1).Blend.SRC_BLEND = srcBlend
    LAYA(NumLAYA - 1).Blend.DEST_BLEND = DestBlend

    '1# - LAYA(NumLAYA - 1).Layed_ID * 0.001
    ' The top of the sky
    V(0).Position = Vector(-1#, 1#, 1#)
    V(1).Position = Vector(-1#, 1#, -1#)
    V(2).Position = Vector(1#, 1#, 1#)
    V(3).Position = Vector(1#, 1#, -1#)
    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0

    V(0).TV = 0
    V(0).Tu = 0
    V(1).TV = 0
    V(1).Tu = 1
    V(2).TV = 1
    V(2).Tu = 0
    V(3).TV = 1
    V(3).Tu = 1

    'CopyMemory Sky.LAYER_VERT(0), V(0), VertexSizeInBytes * 4

    Set LAYA(NumLAYA - 1).VertexBUFF = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData LAYA(NumLAYA - 1).VertexBUFF, 0, VertexSizeInBytes * 4, 0, V(0)

    'Dim BUFF(64 * 64 * 3) As Byte
    '
    'Perlin_Texture 64, 64, BUFF(0)

    LAYA(NumLAYA - 1).TEXT_ID = Add_TextureToPool(TextureFile) 'Add_TextureFromMemory(BUFF, 32, 32)

    'Dim i As Long
    'For i = 0 To 3
    '  LAYER_VERT(i).v.y = LAYA(NumLAYA - 1).HI
    '  LAYER_VERT(i).v.x = LAYER_VERTID(i).v.x * LAYA(NumLAYA - 1).Size.x
    '  LAYER_VERT(i).v.z = LAYER_VERTID(i).v.z * LAYA(NumLAYA - 1).Size.y
    '
    '
    'Next i

End Sub

Private Sub Class_Initialize()

    Blend.SRC_BLEND = D3DBLEND_ONE
    Blend.DEST_BLEND = D3DBLEND_ONE
    '
    ' LAYER_VERT(0).TEX = Vector2D(0, 1)
    ' LAYER_VERT(1).TEX = Vector2D(0, 0)
    ' LAYER_VERT(2).TEX = Vector2D(1, 1)
    ' LAYER_VERT(3).TEX = Vector2D(1, 0)
    '
    '  LAYER_VERT(0).Color = &HFFFFFFFF
    '  LAYER_VERT(1).Color = &HFFFFFFFF
    '  LAYER_VERT(2).Color = &HFFEEEE
    '  LAYER_VERT(3).Color = &HFFFFFFFF
    '
    '    LAYER_VERT(0).v = Vector(-1, 0, -1)
    '    LAYER_VERT(1).v = Vector(-1, 0, 1)
    '
    '    LAYER_VERT(2).v = Vector(1, 0, -1)
    '    LAYER_VERT(3).v = Vector(1, 0, 1)
    '

    CopyMemory LAYER_VERTID(0), LAYER_VERT(0), Len(LAYER_VERT(0)) * 4

End Sub

'----------------------------------------
'Name: CreateSKY
'Description:
'----------------------------------------
Private Sub CreateSKY()

  Dim V(0 To 3) As LITVERTEX
  Dim VertexSizeInBytes As Long

    VertexSizeInBytes = Len(V(0))

    ' The front of the sky
    V(0).Position = Vector(-1#, -1#, 1#)
    V(1).Position = Vector(-1#, 1#, 1#)
    V(2).Position = Vector(1#, -1#, 1#)
    V(3).Position = Vector(1#, 1#, 1#)
    V(0).Tu = 0
    V(0).TV = 1
    V(1).Tu = 0
    V(1).TV = 0
    V(2).Tu = 1
    V(2).TV = 1
    V(3).Tu = 1
    V(3).TV = 0

    Set SKY.VertexBUFF(0) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(0), 0, VertexSizeInBytes * 4, 0, V(0)

    ' The back of the sky
    V(0).Position = Vector(1#, -1#, -1#)
    V(1).Position = Vector(1#, 1#, -1#)
    V(2).Position = Vector(-1#, -1#, -1#)
    V(3).Position = Vector(-1#, 1#, -1#)
    V(0).Tu = 0
    V(0).TV = 1
    V(1).Tu = 0
    V(1).TV = 0
    V(2).Tu = 1
    V(2).TV = 1
    V(3).Tu = 1
    V(3).TV = 0

    'TOTAL_MEM_USED = TOTAL_MEM_USED + VertexSizeInBytes * 4 * 6

    Set SKY.VertexBUFF(1) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(1), 0, VertexSizeInBytes * 4, 0, V(0)

    ' The left of the sky
    V(0).Position = Vector(-1#, -1#, -1#)
    V(1).Position = Vector(-1#, 1#, -1#)
    V(2).Position = Vector(-1#, -1#, 1#)
    V(3).Position = Vector(-1#, 1#, 1#)
    V(0).Tu = 0
    V(0).TV = 1
    V(1).Tu = 0
    V(1).TV = 0
    V(2).Tu = 1
    V(2).TV = 1
    V(3).Tu = 1
    V(3).TV = 0

    Set SKY.VertexBUFF(2) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(2), 0, VertexSizeInBytes * 4, 0, V(0)

    ' The right of the sky
    V(0).Position = Vector(1#, -1#, 1#)
    V(1).Position = Vector(1#, 1#, 1#)
    V(2).Position = Vector(1#, -1#, -1#)
    V(3).Position = Vector(1#, 1#, -1#)
    V(0).Tu = 0
    V(0).TV = 1
    V(1).Tu = 0
    V(1).TV = 0
    V(2).Tu = 1
    V(2).TV = 1
    V(3).Tu = 1
    V(3).TV = 0

    Set SKY.VertexBUFF(3) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(3), 0, VertexSizeInBytes * 4, 0, V(0)

    ' The top of the sky
    V(0).Position = Vector(-1#, 1#, 1#)
    V(1).Position = Vector(-1#, 1#, -1#)
    V(2).Position = Vector(1#, 1#, 1#)
    V(3).Position = Vector(1#, 1#, -1#)
    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0

    V(0).TV = 0
    V(0).Tu = 0
    V(1).TV = 0
    V(1).Tu = 1
    V(2).TV = 1
    V(2).Tu = 0
    V(3).TV = 1
    V(3).Tu = 1

    CopyMemory SKY.LAYER_VERT(0), V(0), VertexSizeInBytes * 4

    Set SKY.VertexBUFF(4) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(4), 0, VertexSizeInBytes * 4, 0, V(0)

    ' The bottom of the sky
    V(0).Position = Vector(-1#, -1#, -1#)
    V(1).Position = Vector(-1#, -1#, 1#)
    V(2).Position = Vector(1#, -1#, -1#)
    V(3).Position = Vector(1#, -1#, 1#)
    'V(0).Tu = 0: V(0).TV = 1
    'V(1).Tu = 0: V(1).TV = 0
    'V(2).Tu = 1: V(2).TV = 1
    'V(3).Tu = 1: V(3).TV = 0

    V(0).TV = 1
    V(0).Tu = 1
    V(1).TV = 1
    V(1).Tu = 0
    V(2).TV = 0
    V(2).Tu = 1
    V(3).TV = 0
    V(3).Tu = 0

    Set SKY.VertexBUFF(5) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
        0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SKY.VertexBUFF(5), 0, VertexSizeInBytes * 4, 0, V(0)

    '    Add_SkyLayer ""
    '
    '    LAYA(0).Text_ID = Add_TextureToPool(App.path + "\cloudsTurb256.tga")

End Sub

'----------------------------------------
'Name: Init_SkyBox
'Object: Init
'Event: SkyBox
'Description:
'----------------------------------------
Sub Init_SkyBox()

    ReDim SKY_TEX(5)
    SKY.Vscal = Vector(1, 1, 1)
    CreateSKY

    'Dim b As Boolean
    '
    'Ver.Add_Declaration QUEST3D_SHADERCODE_STREAM
    'Ver.Add_Declaration QUEST3D_SHADERCODE_POSITION
    'Ver.Add_Declaration QUEST3D_SHADERCODE_TEXCOORD0
    'Ver.Add_Declaration QUEST3D_SHADERCODE_END
    '
    'b = Ver.Init_VertexShaderFromFile(App.path + "\shaders\clouds.vsh")
    'b = PIX.Init_PixelShaderFromFile(App.path + "\shaders\clouds.psh")
    '

End Sub

''----------------------------------------
''Name: RenderSky
''Description:
''----------------------------------------
'
'Private Sub RenderCUBE()
'
'
'
'    obj_Device.SetVertexShader Lit_CUSTOMVERTEX
'    ' Draw the triangles in the vertex buffer
'    ' Note we are now using a triangle strip of vertices
'    ' instead of a triangle list
'    '20 = Len(v)
'
'    'LpGLOBAL_QUEST3D.Set_light 1
'    'get default
'
'    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING
'    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ZFUNC
'
'    LpGLOBAL_QUEST3D.Push_TextureRenderstate 0, D3DTSS_COLORARG1
'    LpGLOBAL_QUEST3D.Push_TextureRenderstate 0, D3DTSS_COLOROP
'
'    Clamp1 = obj_Device.GetTextureStageState(0, D3DTSS_ADDRESSU)
'    Clamp2 = obj_Device.GetTextureStageState(0, D3DTSS_ADDRESSV)
'
'    obj_Device.SetRenderState D3DRS_LIGHTING, 0
'
''    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP)
''    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP)
'
'    Call obj_Device.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
'    Call obj_Device.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)
'
'    Call obj_Device.SetRenderState(D3DRS_ZFUNC, D3DCMP_ALWAYS)
'
'
'   obj_Device.SetTexture 0, Sky.CUBE.Get_CubeTexture
'
'   Call obj_Device.SetTextureStageState(0, D3DTSS_TEXCOORDINDEX, _
    '                                   D3DTSS_TCI_CAMERASPACENORMAL Or 1)
'
'
'    matWorld = CreateMatrix(Sky.Vrotate, Sky.Vscal, Vector(0, 0, 0))
'
'    obj_Device.GetTransform D3DTS_VIEW, matViewSave
'    matView = matViewSave
'    matView.m41 = 0
'    matView.m42 = 0
'    matView.m43 = 0
'    obj_Device.SetTransform D3DTS_VIEW, matView
'
'    obj_Device.SetTransform D3DTS_WORLD, matWorld
'
'
'                                    'Render the front
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_BACK)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(0), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'    'Render the Back
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_FRONT)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(1), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'    'Render the Left
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_LEFT)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(2), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'    'Render the Right
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_RIGHT)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(3), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'    'Render the Top
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_TOP)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(4), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'
'       'Render the down
'    'obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(Sky.SKY_Down)
'    obj_Device.SetStreamSource 0, Sky.VertexBUFF(5), 20
'    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
'
'
'
'
'
' 'restore view,Uv
'    Call obj_Device.SetTextureStageState(0, D3DTSS_TEXCOORDINDEX, 0)
'
'    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSU, Clamp1)
'    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSV, Clamp2)
'    obj_Device.SetTransform D3DTS_VIEW, matViewSave
'    'LpGLOBAL_QUEST3D.CLearMATRIX
'    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING
'    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ZFUNC
'
'    LpGLOBAL_QUEST3D.Pop_TextureRenderstate 0, D3DTSS_COLORARG1
'    LpGLOBAL_QUEST3D.Pop_TextureRenderstate 0, D3DTSS_COLOROP
'
'End Sub
Sub RenderSky()

  Dim I As Integer

    '   If Sky.IS_CUBE Then
    '     RenderCUBE
    '     Exit Sub
    '
    '   End If

    'Dim MatZ As D3DMATRIX
    'Dim Clock As Long

    ' Static Fz As Single
    '
    ' Fz = Fz + 1 / 360
    ' Dim TexMat As D3DMATRIX

    obj_Device.SetVertexShader Lit_CUSTOMVERTEX
    ' Draw the triangles in the vertex buffer
    ' Note we are now using a triangle strip of vertices
    ' instead of a triangle list
    '20 = Len(v)

    'LpGLOBAL_QUEST3D.Set_light 1
    'get default

    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ZFUNC

    LpGLOBAL_QUEST3D.Push_TextureRenderstate 0, D3DTSS_COLORARG1
    LpGLOBAL_QUEST3D.Push_TextureRenderstate 0, D3DTSS_COLOROP

    Clamp1 = obj_Device.GetTextureStageState(0, D3DTSS_ADDRESSU)
    Clamp2 = obj_Device.GetTextureStageState(0, D3DTSS_ADDRESSV)

    obj_Device.SetRenderState D3DRS_LIGHTING, 0

    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP)
    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP)

    Call obj_Device.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
    Call obj_Device.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SELECTARG1)

    Call obj_Device.SetRenderState(D3DRS_ZFUNC, D3DCMP_ALWAYS)

    ' Disable the Zbuffer and render the sky

    'D3DXMatrixScaling matWorld, Sky.Vscal.x, Sky.Vscal.y, Sky.Vscal.z
    'Static uu
    'uu = PI / 500 + uu
    matWorld = Matrix_GetEX(SKY.Vscal, SKY.Vrotate, Vector(0, 0, 0))

    obj_Device.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0
    matView.m42 = 0
    matView.m43 = 0
    obj_Device.SetTransform D3DTS_VIEW, matView

    obj_Device.SetTransform D3DTS_WORLD, matWorld

  Dim Po As QUEST3D_SkyBOXPosition
    Po = QUEST3D_SKY_FRONT

    'Render the front
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_BACK)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(0), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'Render the Back
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_FRONT)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(1), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'Render the Left
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_LEFT)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(2), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'Render the Right
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_RIGHT)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(3), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'Render the Top
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_TOP)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(4), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'Render the down
    obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(SKY.SKY_Down)
    obj_Device.SetStreamSource 0, SKY.VertexBUFF(5), 20
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    'draw layers

    Call LpGLOBAL_QUEST3D.Push_TextureRenderstate(0, D3DTSS_ALPHAOP)
    Call LpGLOBAL_QUEST3D.Push_TextureRenderstate(0, D3DTSS_ALPHAARG1)
    Call LpGLOBAL_QUEST3D.Push_TextureRenderstate(0, D3DTSS_ALPHAARG2)

    Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE)
    Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)

    obj_Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1

    '    For I = 0 To NumLAYA - 1
    '
    '        matWorld = CreateMatrix(Sky.Vrotate, Vector(Sky.Vscal.x, Sky.Vscal.y, Sky.Vscal.z), Vector(0, -(I + 1) * 0.5, 0))
    '
    '         obj_Device.SetRenderState D3DRS_SRCBLEND, Blend.SRC_BLEND 'LAYA(I).Blend.SRC_BLEND
    '        obj_Device.SetRenderState D3DRS_DESTBLEND, Blend.DEST_BLEND 'LAYA(I).Blend.DEST_BLEND
    '
    '         obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(LAYA(I).TEXT_ID)
    '        obj_Device.SetStreamSource 0, LAYA(I).VertexBUFF, 20
    '        obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    '
    '    Next I
    '
    obj_Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0

    Call LpGLOBAL_QUEST3D.Pop_TextureRenderstate(0, D3DTSS_ALPHAOP)
    Call LpGLOBAL_QUEST3D.Pop_TextureRenderstate(0, D3DTSS_ALPHAARG1)
    Call LpGLOBAL_QUEST3D.Pop_TextureRenderstate(0, D3DTSS_ALPHAARG2)

    Add_Tri 16
    Add_Verti 16 * 3

    'restore view,Uv
    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSU, Clamp1)
    Call obj_Device.SetTextureStageState(0, D3DTSS_ADDRESSV, Clamp2)
    obj_Device.SetTransform D3DTS_VIEW, matViewSave
    'LpGLOBAL_QUEST3D.CLearMATRIX
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ZFUNC

    LpGLOBAL_QUEST3D.Pop_TextureRenderstate 0, D3DTSS_COLORARG1
    LpGLOBAL_QUEST3D.Pop_TextureRenderstate 0, D3DTSS_COLOROP

Exit Sub

    '
    '   matWorld = CreateMatrix(Sky.Vrotate, Vector(Sky.Vscal.x, Sky.Vscal.y, Sky.Vscal.z), Vector(0, 0, 0))
    '
    'obj_Device.SetTransform D3DTS_WORLD, matWorld
    '
    '
    '
    '    'Render the Top
    '      obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(LAYA(0).Text_ID)
    '      obj_Device.SetTexture 1, myTEXPOOL.POOL_texture(LAYA(0).Text_ID)
    '        obj_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky.LAYER_VERT(0), 20

    ''render layer
    'Dim C0P As D3DVECTOR4
    'Dim C1P As D3DVECTOR4
    'Dim C2P As D3DVECTOR4
    'Dim C3P As D3DVECTOR4
    'Dim C4P As D3DVECTOR4
    'Dim C5P As D3DVECTOR4
    'Dim C6P As D3DVECTOR4
    'Dim Mat As D3DMATRIX
    '
    'Dim C0 As D3DVECTOR4
    '
    'LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ZENABLE
    'LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_CULLMODE
    'LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING
    '
    'LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHABLENDENABLE
    '
    'LpGLOBAL_QUEST3D.Push_PixelShader
    '
    '
    '
    '
    '
    '    Call obj_Device.SetRenderState(D3DRS_ZENABLE, False)
    '    Call obj_Device.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    '    Call obj_Device.SetRenderState(D3DRS_LIGHTING, False)
    '
    '    Call obj_Device.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    '    Call obj_Device.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
    '    Call obj_Device.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    '
    '
    '   obj_Device.SetTextureStageState 1, D3DTSS_TEXCOORDINDEX, 0
    '
    'obj_Device.SetPixelShader PIX.Get_PixelShader
    '
    'obj_Device.SetVertexShader Ver.Get_VertexShader
    '
    '       C0 = Vector4(Cos(Timer) * 0.0000514, Timer * 0.00515, 0#, 0#)
    '
    '
    '  C0P = Vector4(0#, 0.5, 1#, 0.75)
    '         C1P = Vector4(0.07, 0.07, 0#, 0#)
    '         C2P = Vector4(1#, 0.7, 0.5, 1#)
    '         C3P = Vector4(0.2, 0.2, 0#, 0#)
    '         C4P = Vector4(1#, 0.7, 0.5, 1#)
    '         C5P = Vector4(0.7, 0#, 0.7, 1#)
    '         C6P = Vector4(0#, 0#, 0.7, 1#)
    '
    '        Call obj_Device.SetPixelShaderConstant(0, C0P, 1)
    '        Call obj_Device.SetPixelShaderConstant(1, C1P, 1)
    '        Call obj_Device.SetPixelShaderConstant(2, C2P, 1)
    '        Call obj_Device.SetPixelShaderConstant(3, C3P, 1)
    '        Call obj_Device.SetPixelShaderConstant(4, C4P, 1)
    '        Call obj_Device.SetPixelShaderConstant(5, C5P, 1)
    '        Call obj_Device.SetPixelShaderConstant(6, C6P, 1)
    '
    'Call obj_Device.SetVertexShaderConstant(0, C0, 1)
    '
    '
    '
    '  matWorld = CreateMatrix(Sky.Vrotate, Vector(Sky.Vscal.x, Sky.Vscal.y - 1, Sky.Vscal.z), Vector(0, 0, 0))
    '
    ''obj_Device.SetTransform D3DTS_WORLD, matWorld
    '
    '
    'Call D3DXMatrixMultiply(Mat, matViewSave, Data.MatProjec)
    '        Call D3DXMatrixMultiply(Mat, matWorld, Mat)
    '        Call D3DXMatrixTranspose(Mat, Mat)
    '        Call obj_Device.SetVertexShaderConstant(4, Mat, 4)
    '
    '
    '
    '
    '    'Render the Top
    ''      obj_Device.SetTexture 0, myTEXPOOL.POOL_texture(LAYA(0).Text_ID)
    ''      obj_Device.SetTexture 1, myTEXPOOL.POOL_texture(LAYA(0).Text_ID)
    ''
    '      obj_Device.SetStreamSource 0, Sky.VertexBUFF(4), 20
    '      obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    '      'obj_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Sky.LAYER_VERT(0), 20
    '
    'LpGLOBAL_QUEST3D.Pop_PixelShader
    '
    'LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ZENABLE
    'LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_CULLMODE
    'LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING
    '
    'LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHABLENDENABLE
    '
    '
    '
    'Exit Sub
    '

End Sub

Sub Set_Rotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    SKY.Vrotate.x = x
    SKY.Vrotate.y = y
    SKY.Vrotate.z = z
    'ComputeBOX

End Sub

'----------------------------------------
'Name: Set_Scale
'Object: Set
'Event: Scale
'Description:
'----------------------------------------
Sub Set_Scale(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    SKY.Vscal.x = x
    SKY.Vscal.y = y
    SKY.Vscal.z = z

End Sub

Sub Set_TextureCoordonates(Position As QUEST3D_SkyBOXPosition, ByVal Vertex0_Tu As Single, ByVal Vertex0_Tv As Single, Vertex1_Tu As Single, ByVal Vertex1_Tv As Single, Vertex2_Tu As Single, ByVal Vertex2_Tv As Single, Vertex3_Tu As Single, ByVal Vertex3_Tv As Single)

  Dim V(0 To 3) As LITVERTEX
  Dim VertexSizeInBytes As Long

    VertexSizeInBytes = Len(V(0))

    If Position = QUEST3D_SKY_FRONT Then
        ' The front of the sky
        V(0).Position = Vector(-1#, -1#, 1#)
        V(1).Position = Vector(-1#, 1#, 1#)
        V(2).Position = Vector(1#, -1#, 1#)
        V(3).Position = Vector(1#, 1#, 1#)
        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(0) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(0), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The back of the sky
    If Position = QUEST3D_SKY_BACK Then

        V(0).Position = Vector(1#, -1#, -1#)
        V(1).Position = Vector(1#, 1#, -1#)
        V(2).Position = Vector(-1#, -1#, -1#)
        V(3).Position = Vector(-1#, 1#, -1#)
        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(1) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(1), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The left of the sky
    If Position = QUEST3D_SKY_LEFT Then

        V(0).Position = Vector(-1#, -1#, -1#)
        V(1).Position = Vector(-1#, 1#, -1#)
        V(2).Position = Vector(-1#, -1#, 1#)
        V(3).Position = Vector(-1#, 1#, 1#)
        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(2) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(2), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The right of the sky
    If Position = QUEST3D_SKY_RIGHT Then

        V(0).Position = Vector(1#, -1#, 1#)
        V(1).Position = Vector(1#, 1#, 1#)
        V(2).Position = Vector(1#, -1#, -1#)
        V(3).Position = Vector(1#, 1#, -1#)
        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(3) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(3), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

    ' The top of the sky
    If Position = QUEST3D_SKY_TOP Then

        V(0).Position = Vector(-1#, 1#, 1#)
        V(1).Position = Vector(-1#, 1#, -1#)
        V(2).Position = Vector(1#, 1#, 1#)
        V(3).Position = Vector(1#, 1#, -1#)
        'V(0).Tu = 0: V(0).TV = 1
        'V(1).Tu = 0: V(1).TV = 0
        'V(2).Tu = 1: V(2).TV = 1
        'V(3).Tu = 1: V(3).TV = 0

        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(4) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(4), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

    ' The bottom of the sky
    If Position = QUEST3D_SKY_DOWN Then

        V(0).Position = Vector(-1#, -1#, -1#)
        V(1).Position = Vector(-1#, -1#, 1#)
        V(2).Position = Vector(1#, -1#, -1#)
        V(3).Position = Vector(1#, -1#, 1#)
        V(0).Tu = Vertex0_Tu
        V(0).TV = Vertex0_Tv
        V(1).Tu = Vertex1_Tu
        V(1).TV = Vertex1_Tv
        V(2).Tu = Vertex2_Tu
        V(2).TV = Vertex2_Tv
        V(3).Tu = Vertex3_Tu
        V(3).TV = Vertex3_Tv

        Set SKY.VertexBUFF(5) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(5), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

End Sub

Sub Set_TextureRotateCoordonates(Position As QUEST3D_SkyBOXPosition, ByVal Angle As Single)

  Dim Vertex0_Tu As Single, Vertex0_Tv As Single, Vertex1_Tu As Single, Vertex1_Tv As Single, Vertex2_Tu As Single, Vertex2_Tv As Single, Vertex3_Tu As Single, Vertex3_Tv As Single
  Dim V(0 To 3) As LITVERTEX
  Dim VertexSizeInBytes As Long

    ' The front of the sky
    If Position = QUEST3D_SKY_FRONT Then
        V(0).Tu = 0
        V(0).TV = 1
        V(1).Tu = 0
        V(1).TV = 0
        V(2).Tu = 1
        V(2).TV = 1
        V(3).Tu = 1
        V(3).TV = 0
        V(0).Position = Vector(-1#, -1#, 1#)
        V(1).Position = Vector(-1#, 1#, 1#)
        V(2).Position = Vector(1#, -1#, 1#)
        V(3).Position = Vector(1#, 1#, 1#)
    End If

    ' The back of the sky
    If Position = QUEST3D_SKY_BACK Then
        V(0).Tu = 0
        V(0).TV = 1
        V(1).Tu = 0
        V(1).TV = 0
        V(2).Tu = 1
        V(2).TV = 1
        V(3).Tu = 1
        V(3).TV = 0

        V(0).Position = Vector(1#, -1#, -1#)
        V(1).Position = Vector(1#, 1#, -1#)
        V(2).Position = Vector(-1#, -1#, -1#)
        V(3).Position = Vector(-1#, 1#, -1#)
    End If

    ' The left of the sky
    If Position = QUEST3D_SKY_LEFT Then
        V(0).Tu = 0
        V(0).TV = 1
        V(1).Tu = 0
        V(1).TV = 0
        V(2).Tu = 1
        V(2).TV = 1
        V(3).Tu = 1
        V(3).TV = 0
        V(0).Position = Vector(-1#, -1#, -1#)
        V(1).Position = Vector(-1#, 1#, -1#)
        V(2).Position = Vector(-1#, -1#, 1#)
        V(3).Position = Vector(-1#, 1#, 1#)
    End If

    ' The right of the sky
    If Position = QUEST3D_SKY_RIGHT Then
        V(0).Tu = 0
        V(0).TV = 1
        V(1).Tu = 0
        V(1).TV = 0
        V(2).Tu = 1
        V(2).TV = 1
        V(3).Tu = 1
        V(3).TV = 0

        V(0).Position = Vector(1#, -1#, 1#)
        V(1).Position = Vector(1#, 1#, 1#)
        V(2).Position = Vector(1#, -1#, -1#)
        V(3).Position = Vector(1#, 1#, -1#)
    End If

    ' The top of the sky
    If Position = QUEST3D_SKY_TOP Then

        V(0).TV = 0
        V(0).Tu = 0
        V(1).TV = 0
        V(1).Tu = 1
        V(2).TV = 1
        V(2).Tu = 0
        V(3).TV = 1
        V(3).Tu = 1

        V(0).Position = Vector(-1#, 1#, 1#)
        V(1).Position = Vector(-1#, 1#, -1#)
        V(2).Position = Vector(1#, 1#, 1#)
        V(3).Position = Vector(1#, 1#, -1#)
    End If

    ' The bottom of the sky
    If Position = QUEST3D_SKY_DOWN Then

        V(0).TV = 1
        V(0).Tu = 1
        V(1).TV = 1
        V(1).Tu = 0
        V(2).TV = 0
        V(2).Tu = 1
        V(3).TV = 0
        V(3).Tu = 0

        V(0).Position = Vector(-1#, -1#, -1#)
        V(1).Position = Vector(-1#, -1#, 1#)
        V(2).Position = Vector(1#, -1#, -1#)
        V(3).Position = Vector(1#, -1#, 1#)
    End If
    
    '===================================================
    'In fact we rotate vertex via matrix transformation
    '
    '====================================================
    

  Dim V1 As D3DVECTOR
  Dim V2 As D3DVECTOR
  Dim Matrix As D3DMATRIX

    If Position = QUEST3D_SKY_TOP Then

        Matrix = Matrix_Get(1, 1, 1, 0, Angle, 0, 0, 0, 0)

      ElseIf Position = QUEST3D_SKY_FRONT Then

        Matrix = Matrix_Get(1, 1, 1, 0, 0, Angle, 0, 0, 0)
        
    
      ElseIf Position = QUEST3D_SKY_BACK Then

        Matrix = Matrix_Get(1, 1, 1, 0, 0, Angle, 0, 0, 0)
    
       ElseIf Position = QUEST3D_SKY_LEFT Or Position = QUEST3D_SKY_RIGHT Then

        Matrix = Matrix_Get(1, 1, 1, Angle, 0, 0, 0, 0, 0)
        

    End If

    V1 = V(0).Position
    D3DXVec3TransformCoord V(0).Position, V1, Matrix

    V1 = V(1).Position
    D3DXVec3TransformCoord V(1).Position, V1, Matrix

    V1 = V(2).Position
    D3DXVec3TransformCoord V(2).Position, V1, Matrix

    V1 = V(3).Position
    D3DXVec3TransformCoord V(3).Position, V1, Matrix

    VertexSizeInBytes = Len(V(0))

    If Position = QUEST3D_SKY_FRONT Then
        ' The front of the sky

        Set SKY.VertexBUFF(0) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(0), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The back of the sky
    If Position = QUEST3D_SKY_BACK Then

        Set SKY.VertexBUFF(1) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(1), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The left of the sky
    If Position = QUEST3D_SKY_LEFT Then

        Set SKY.VertexBUFF(2) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(2), 0, VertexSizeInBytes * 4, 0, V(0)
    End If

    ' The right of the sky
    If Position = QUEST3D_SKY_RIGHT Then

        Set SKY.VertexBUFF(3) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(3), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

    ' The top of the sky
    If Position = QUEST3D_SKY_TOP Then

        Set SKY.VertexBUFF(4) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(4), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

    ' The bottom of the sky
    If Position = QUEST3D_SKY_DOWN Then

        Set SKY.VertexBUFF(5) = obj_Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
            0, Lit_CUSTOMVERTEX, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData SKY.VertexBUFF(5), 0, VertexSizeInBytes * 4, 0, V(0)

    End If

End Sub

'----------------------------------------
'Name: Let VScale
'Description:
'----------------------------------------
Public Property Let VScale(vNew As D3DVECTOR)

    SKY.Vscal = vNew

End Property

'----------------------------------------
'Name: Get VScale
'Description:
'----------------------------------------
Public Property Get VScale() As D3DVECTOR

    VScale = SKY.Vscal

End Property
