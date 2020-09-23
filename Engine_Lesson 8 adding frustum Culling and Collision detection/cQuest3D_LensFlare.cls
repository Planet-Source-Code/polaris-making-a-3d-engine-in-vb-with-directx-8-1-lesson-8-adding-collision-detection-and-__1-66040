VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_LensFlare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name:
'     Developer/Programmer: Johna
'    ----------------------------------------------------------
'     Module Name: cQUEST3D_LensFlare
'     Module File: QUEST3D_lensflare.cls
'     Module Type: Class
'     Module Description:
'    ----------------------------------------------------------
'     © johna 2002
'    ----------------------------------------------------------
'   ============================================================

'===========================================
' QUEST3D Lenflare class version 1.00
' created 26-may 2002 by Polaris
' (c) 2002 QUEST3DEngine version 1.04
'
'
'
'Please do not copy or modify this code without my autozation
'if you do it, ask me and give me credits
'because this code is the most advanced Lensflare routines
'i have ever created and the most advanced in VB
'===========================================

Option Explicit
Private Type Lens_VERTEX
    POS As D3DVECTOR
    color As Long
    Diffuse As Long
    Tu As Single
    TV As Single

End Type

Enum EN_LENSFSOURCE_STATUS
    STATIC_SOURCE = &HCDFCD
    CAMERA_RELATIVE = &HFCF

End Enum

Private Type TLens
    Vertex(3) As QUEST3D_LVERTEX
    Position As Single
    TextureIndex As Long
    Size As Single

    VertexBUFF As Direct3DVertexBuffer8

End Type

'Private Type QUEST3D_TEXTURE_BLEND
'    SRC_BLEND As CONST_D3DBLEND
'    DEST_BLEND As CONST_D3DBLEND
'End Type

Enum QUEST3D_LENSFLARE_BLENDMODE
    SUN_BLENDMODE = &H101
    FLARES_BLENDMODE = &H1011
    BURNOUT_BLENDMODE = &H1101

End Enum

Private Type TFlare
    Flares() As TLens
    Num_Lens As Integer

    SUN_pos As D3DVECTOR
    SUN_FromCAM As D3DVECTOR
    BURNOUT(3) As QUEST3D_TLVERTEX
    BURNOUT_VertexBUFF As Direct3DVertexBuffer8
    
    Status As EN_LENSFSOURCE_STATUS
    DrawBurnScreen As Boolean
    ColorR As Byte
    ColorG As Byte
    ColorB As Byte
    Sun_Blend As QUEST3D_TEXTURE_BLEND
    Flares_Blend As QUEST3D_TEXTURE_BLEND
    BurnOut_Blend As QUEST3D_TEXTURE_BLEND
    Sun_animated As Boolean
    Sun_RotSpeed As Single
    Min_VisibleAngle As Single
    Speed As Single

End Type

Dim SUN(3) As QUEST3D_LVERTEX
Dim SUN_VertexBUFF As Direct3DVertexBuffer8
Dim SUN_TEXID As Long
Dim MyFlare As TFlare
Dim SunSize As Single
Dim SunPos As D3DVECTOR
Dim SunWI As Single
Dim SunHi As Single

'----------------------------------------
'Name: Add_Lens
'Object: Add
'Event: Lens
'Description:
'----------------------------------------
Sub Add_Lens(Texture As String, Index As Long, Size As Single, POS As Single, Basecolor As Long, Highcolor As Long)

    CreateFlare Index, Size, POS, Basecolor, Highcolor, Texture
    MyFlare.Num_Lens = Index

End Sub

'----------------------------------------
'Name: Add_SunTEX
'Object: Add
'Event: SunTEX
'Description:
'----------------------------------------
Sub Add_SunTEX(Texture As String, POS As D3DVECTOR, Optional ByVal Size As Single = 100)

    Call Init_Sun(Texture, POS, Size)
    SunSize = Size
    SunPos = POS

End Sub

Private Sub Class_Initialize()

    MyFlare.ColorR = 255
    MyFlare.ColorG = 255
    MyFlare.ColorB = 255

    MyFlare.Sun_RotSpeed = QUEST3D_PI / 200

    MyFlare.DrawBurnScreen = True

    MyFlare.Status = CAMERA_RELATIVE

    MyFlare.Sun_Blend.SRC_BLEND = D3DBLEND_SRCCOLOR
    MyFlare.Sun_Blend.DEST_BLEND = D3DBLEND_ONE

    MyFlare.Flares_Blend.SRC_BLEND = D3DBLEND_SRCCOLOR
    MyFlare.Flares_Blend.DEST_BLEND = D3DBLEND_ONE

    MyFlare.BurnOut_Blend.SRC_BLEND = D3DBLEND_SRCALPHA
    MyFlare.BurnOut_Blend.DEST_BLEND = D3DBLEND_INVSRCALPHA

    MyFlare.Min_VisibleAngle = 0.85

    MyFlare.Speed = 1

End Sub

'----------------------------------------
'Name: CreateFlare
'Description:
'----------------------------------------
Private Sub CreateFlare(Index As Long, Size As Single, flare_position As Single, Basecolor As Long, highlightCOLOR As Long, Texture As String)

    ReDim Preserve MyFlare.Flares(Index + 1)

    MyFlare.Flares(Index).Vertex(1) = Make_LVertex(-Size, -Size, 0, Basecolor, 0, 0, 0)
    MyFlare.Flares(Index).Vertex(0) = Make_LVertex(Size, -Size, 0, highlightCOLOR, 0, 1, 0)
    MyFlare.Flares(Index).Vertex(3) = Make_LVertex(-Size, Size, 0, highlightCOLOR, 0, 0, 1)
    MyFlare.Flares(Index).Vertex(2) = Make_LVertex(Size, Size, 0, Basecolor, 0, 1, 1)
    MyFlare.Flares(Index).Position = flare_position
    MyFlare.Flares(Index).Size = Size

    'we create the VERTEX BUFFER
    Create_VertexBufferFromLVERTEX MyFlare.Flares(Index).VertexBUFF, MyFlare.Flares(Index).Vertex

    MyFlare.Flares(Index).TextureIndex = Add_TextureToPool(Texture)

End Sub

'----------------------------------------
'Name: Make_TLVertex
'Description:
'----------------------------------------

Function Get_MinimumVisibilityAngle() As Single

    Get_MinimumVisibilityAngle = MyFlare.Min_VisibleAngle

End Function

Function Get_SunPosition() As D3DVECTOR

    Get_SunPosition = MyFlare.SUN_pos

End Function

Function Get_SunSize() As Single

    Get_SunSize = SunSize

End Function

'----------------------------------------
'Name: GetAlphaTEX
'Description:
'----------------------------------------

'----------------------------------------
'Name: Init_Sun
'Object: Init
'Event: Sun
'Description:
'----------------------------------------
Private Sub Init_Sun(SunTEX As String, POS As D3DVECTOR, Optional ByVal Size = 100)

  Dim Position As D3DVECTOR

    Position = Vector(-Size, -Size, 0)
    SUN(0).x = Position.x
    SUN(0).y = Position.y
    SUN(0).z = Position.z

    Position = Vector(Size, -Size, 0)
    SUN(1).x = Position.x
    SUN(1).y = Position.y
    SUN(1).z = Position.z

    Position = Vector(-Size, Size, 0)
    SUN(2).x = Position.x
    SUN(2).y = Position.y
    SUN(2).z = Position.z

    Position = Vector(Size, Size, 0)
    SUN(3).x = Position.x
    SUN(3).y = Position.y
    SUN(3).z = Position.z

    SUN(0).color = Make_ColorRGBAEx(0.9, 0.9, 0.9, 0)
    SUN(1).color = SUN(0).color
    SUN(2).color = SUN(0).color
    SUN(3).color = SUN(0).color

    SUN(0).Specular = SUN(0).color
    SUN(1).Specular = SUN(0).color
    SUN(2).Specular = SUN(0).color
    SUN(3).Specular = SUN(0).color

    SUN(0).Tu = 0
    SUN(0).TV = 0
    SUN(1).Tu = 1
    SUN(1).TV = 0
    SUN(2).Tu = 0
    SUN(2).TV = 1
    SUN(3).Tu = 1
    SUN(3).TV = 1

    Create_VertexBufferFromLVERTEX SUN_VertexBUFF, SUN

    MyFlare.SUN_pos = POS

    SUN_TEXID = Add_TextureToPool(SunTEX)  'GetAlphaTEX(SunTEX)
    'Set SUN_TEX = GetAlphaTEX(SunTEX)

    D3DXVec3Subtract MyFlare.SUN_FromCAM, MyFlare.SUN_pos, Data.EYES.EYE

  Dim VP As D3DVIEWPORT8

    obj_Device.GetViewport VP

    SunWI = VP.Width
    SunHi = VP.Height

    'create our burnout screen effect
    MyFlare.BURNOUT(0) = Make_TLVertex(0#, 0#, 0#, 1, Make_ColorRGBAEx(1, 1, 1, 1), 1, 0, 0)
    MyFlare.BURNOUT(1) = Make_TLVertex(VP.Width, 0#, 0#, 1, Make_ColorRGBAEx(1, 1, 1, 1), 0, 0, 0)
    MyFlare.BURNOUT(2) = Make_TLVertex(0#, VP.Height, 0#, 1, Make_ColorRGBAEx(1, 1, 1, 1), 0, 0, 0)
    MyFlare.BURNOUT(3) = Make_TLVertex(VP.Width, VP.Height, 0#, 1, Make_ColorRGBAEx(1, 1, 1, 1), 0, 0, 0)

   'Create_VertexBufferFromTLVERTEX MyFlare.BURNOUT_VertexBUFF, MyFlare.BURNOUT
   
End Sub

'----------------------------------------
'Name: Update
'Description:
'----------------------------------------
Sub Render()

  'save state

    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ALPHABLENDENABLE
    LpGLOBAL_QUEST3D.Push_Renderstate D3DRS_ALPHATESTENABLE
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_DESTBLEND
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_DESTBLEND
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ZENABLE
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_ZWRITEENABLE
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_CULLMODE

    'save texture state
    LpGLOBAL_QUEST3D.Push_AllTextureRenderstate

  Dim AddV As D3DVECTOR

  Dim anglevector As D3DVECTOR
  Dim Angle As Single
  Dim z As Integer
  Static zzdeg As Single
  Dim x

    'D3DXVec3Scale AddV, data.EYES.EYE, 10.4
    If MyFlare.Status = CAMERA_RELATIVE Then D3DXVec3Add MyFlare.SUN_pos, Data.EYES.EYE, MyFlare.SUN_FromCAM

  Dim vt As D3DMATRIX
  Dim fl As Single
    Call D3DXMatrixLookAtLH(vt, MyFlare.SUN_pos, Data.EYES.EYE, Vector(0, 1, 0))

    Call D3DXMatrixInverse(vt, fl, vt)

    'draw sun
    LpGLOBAL_QUEST3D.Set_EngineLight 1
    obj_Device.SetRenderState D3DRS_ZENABLE, 0
    obj_Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    obj_Device.SetRenderState D3DRS_ZWRITEENABLE, 0
    Call obj_Device.SetRenderState(D3DRS_SRCBLEND, MyFlare.Sun_Blend.SRC_BLEND)
    Call obj_Device.SetRenderState(D3DRS_DESTBLEND, MyFlare.Sun_Blend.DEST_BLEND)

    If MyFlare.Sun_animated Then
        zzdeg = zzdeg + (MyFlare.Sun_RotSpeed * LpGLOBAL_QUEST3D.Get_TimePassed) * MyFlare.Speed
        If zzdeg > QUEST3D_PI * 100 Then zzdeg = 0

  Dim MatZ As D3DMATRIX
        D3DXMatrixRotationAxis MatZ, Vector(0, 0, 1), zzdeg

        D3DXMatrixMultiply vt, MatZ, vt
    End If

    LpGLOBAL_QUEST3D.Set_EngineTextureByID SUN_TEXID, 0
    obj_Device.SetTransform D3DTS_WORLD, vt

    '    obj_Device.SetVertexShader QUEST3D_FVFLVERTEX
    '    obj_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SUN(0), Len(SUN(0))
    '
    obj_Device.SetVertexShader QUEST3D_FVFLVERTEX
    obj_Device.SetStreamSource 0, SUN_VertexBUFF, Len(SUN(0))
    obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    Add_Tri 2
    Add_Verti 4

    D3DXVec3Subtract anglevector, MyFlare.SUN_pos, Data.EYES.Dest_at
    D3DXVec3Normalize anglevector, anglevector

  Dim Mat As D3DMATRIX

    obj_Device.GetTransform D3DTS_VIEW, Mat

    Angle = D3DXVec3Dot(Vector(Mat.m13, Mat.m23, Mat.m33), anglevector)

    If (Angle > MyFlare.Min_VisibleAngle) Then

        Call obj_Device.GetTransform(D3DTS_VIEW, vt)
  Dim newpos As D3DMATRIX
  Dim lookatvector As D3DVECTOR
  Dim Distance As Single

        Call D3DXMatrixInverse(vt, fl, vt)

        Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1)
        Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)

        Call obj_Device.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
        Call obj_Device.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_DIFFUSE)
        Call obj_Device.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
        Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1)
        Call obj_Device.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)

        Call obj_Device.SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        Call obj_Device.SetRenderState(D3DRS_SRCBLEND, MyFlare.Flares_Blend.SRC_BLEND)
        Call obj_Device.SetRenderState(D3DRS_DESTBLEND, MyFlare.Flares_Blend.DEST_BLEND)

        Call obj_Device.SetRenderState(D3DRS_LIGHTING, 0)
        Call obj_Device.SetRenderState(D3DRS_ZENABLE, 0)
        Call obj_Device.SetRenderState(D3DRS_ZWRITEENABLE, 0)

        Call obj_Device.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)

        While z <= MyFlare.Num_Lens

            Call D3DXMatrixIdentity(newpos)
            Call D3DXMatrixTranslation(newpos, 0, 0, 50)
            Call D3DXMatrixMultiply(newpos, newpos, vt)

            D3DXVec3Subtract lookatvector, Vector(MyFlare.SUN_pos.x, MyFlare.SUN_pos.y, MyFlare.SUN_pos.z), Vector(newpos.m41, newpos.m42, newpos.m43)

            Distance = Vector_Distance(Vector(MyFlare.SUN_pos.x, MyFlare.SUN_pos.y, MyFlare.SUN_pos.z), Vector(newpos.m41, newpos.m42, newpos.m43))

            Call D3DXVec3Normalize(lookatvector, lookatvector)

            newpos.m41 = newpos.m41 + (lookatvector.x * MyFlare.Flares(z).Position)
            newpos.m42 = newpos.m42 + (lookatvector.y * MyFlare.Flares(z).Position)
            newpos.m43 = newpos.m43 + (lookatvector.z * MyFlare.Flares(z).Position)

            LpGLOBAL_QUEST3D.Set_EngineTextureByID MyFlare.Flares(z).TextureIndex, 0
            obj_Device.SetTransform D3DTS_WORLD, newpos

            obj_Device.SetVertexShader QUEST3D_FVFLVERTEX
            obj_Device.SetStreamSource 0, MyFlare.Flares(z).VertexBUFF, Len(MyFlare.Flares(z).Vertex(0))
            obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

            'obj_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MyFlare.Flares(z).Vertex(0), Len(SUN(0))

            Add_Verti 4
            Add_Tri 2

            z = z + 1
        Wend
    End If

    If MyFlare.DrawBurnScreen Then
        If (Angle > 0.85) Then  'If we are looking into the sun at all calculate the alpha and render the screen burnout polygon
  Dim Alpha
            Alpha = (Abs(Angle) - 0.9) * 1880 / 255
            'Alpha = Abs(Alpha)
            'Debug.Print Alpha
            Alpha = Alpha * 255
            If Alpha > 255 Then Alpha = 255

            If (Alpha < 0#) Then Alpha = 0#
            For x = 0 To 3

                MyFlare.BURNOUT(x).color = Make_ColorRGBA(MyFlare.ColorR, MyFlare.ColorG, MyFlare.ColorB, Alpha)
            Next x

            Call obj_Device.SetRenderState(D3DRS_SRCBLEND, MyFlare.BurnOut_Blend.SRC_BLEND)
            Call obj_Device.SetRenderState(D3DRS_DESTBLEND, MyFlare.BurnOut_Blend.DEST_BLEND)
            
            obj_Device.SetTexture 0, Nothing
            obj_Device.SetVertexShader QUEST3D_FVFTLVERTEX
            
            obj_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MyFlare.BURNOUT(0), Len(MyFlare.BURNOUT(0))
'            obj_Device.SetStreamSource 0, MyFlare.BURNOUT_VertexBUFF, Len(MyFlare.BURNOUT(0))
'            obj_Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
            
            Add_Tri 2
            Add_Verti 6

        End If
    End If

    'recall state

    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ALPHABLENDENABLE
    LpGLOBAL_QUEST3D.Pop_Renderstate D3DRS_ALPHATESTENABLE
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_DESTBLEND
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_DESTBLEND
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ZENABLE
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_ZWRITEENABLE
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_CULLMODE

    'save texture state
    LpGLOBAL_QUEST3D.Pop_AllTextureRenderstate

End Sub

'----------------------------------------
'Name: RLVertex
'Description:
'----------------------------------------

Sub Set_LenflareBlend(Blend As QUEST3D_LENSFLARE_BLENDMODE, SRC_BLEND As CONST_D3DBLEND, DEST_BLEND As CONST_D3DBLEND)

    If Blend = SUN_BLENDMODE Then
        MyFlare.Sun_Blend.SRC_BLEND = SRC_BLEND
        MyFlare.Sun_Blend.DEST_BLEND = DEST_BLEND
      ElseIf Blend = FLARES_BLENDMODE Then
        MyFlare.Flares_Blend.SRC_BLEND = SRC_BLEND
        MyFlare.Flares_Blend.DEST_BLEND = DEST_BLEND
      ElseIf Blend = BURNOUT_BLENDMODE Then
        MyFlare.BurnOut_Blend.SRC_BLEND = SRC_BLEND
        MyFlare.BurnOut_Blend.DEST_BLEND = DEST_BLEND
    End If

End Sub

'Private lpQUEST3D As QUEST3DX

'----------------------------------------
'Name: Set_LensPositionStatus
'Object: Set
'Event: LensPositionStatus
'Description:
'----------------------------------------
Sub Set_LensPositionStatus(Status As EN_LENSFSOURCE_STATUS)

    MyFlare.Status = Status

End Sub

Sub Set_MinimumVisibilityAngle(Optional ByVal Angle As Single = 0.85)

    MyFlare.Min_VisibleAngle = Angle

End Sub

Sub Set_SunAnimation(Enable As Boolean)

    MyFlare.Sun_animated = Enable

End Sub

Sub Set_SunAnimationSpeed(Optional Speed As Single = 1)

    MyFlare.Speed = Speed

End Sub

Sub Set_SunburningEffectColor(R As Byte, g As Byte, b As Byte)

    MyFlare.ColorR = R
    MyFlare.ColorG = g
    MyFlare.ColorB = b

End Sub

Sub Set_SunburningScreen(Enable As Boolean)

    MyFlare.DrawBurnScreen = Enable

End Sub

'----------------------------------------
'Name: Set_SunPosition
'Object: Set
'Event: SunPosition
'Description:
'----------------------------------------
Sub Set_SunPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    MyFlare.SUN_pos = Vector(x, y, z)
    SunPos = Vector(x, y, z)

End Sub

'----------------------------------------
'Name: Set_SunSize
'Object: Set
'Event: SunSize
'Description:
'----------------------------------------
Sub Set_SunSize(ByVal Size As Single)

  Dim Position As D3DVECTOR

    Position = Vector(-Size, -Size, 0)
    SUN(0).x = Position.x
    SUN(0).y = Position.y
    SUN(0).z = Position.z

    Position = Vector(Size, -Size, 0)
    SUN(1).x = Position.x
    SUN(1).y = Position.y
    SUN(1).z = Position.z

    Position = Vector(-Size, Size, 0)
    SUN(2).x = Position.x
    SUN(2).y = Position.y
    SUN(2).z = Position.z

    Position = Vector(Size, Size, 0)
    SUN(3).x = Position.x
    SUN(3).y = Position.y
    SUN(3).z = Position.z

   
    Create_VertexBufferFromLVERTEX SUN_VertexBUFF, SUN

End Sub

Sub Set_SunColor(ByVal color As Long)

    SUN(0).color = color
    SUN(1).color = SUN(0).color
    SUN(2).color = SUN(0).color
    SUN(3).color = SUN(0).color

End Sub

Sub Free()

  Dim I As Integer

    For I = 0 To MyFlare.Num_Lens - 1
        Set MyFlare.Flares(I).VertexBUFF = Nothing
    Next I
    
    Set SUN_VertexBUFF = Nothing

    ReDim MyFlare.Flares(0)
    Erase MyFlare.Flares

End Sub
