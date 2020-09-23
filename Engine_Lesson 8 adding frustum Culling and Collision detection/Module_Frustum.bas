Attribute VB_Name = "Module_Frustum"
'==========================================================================================
'
'   This module contains all useful methods to compute sphere and box visibility
'
'  SetUpFrustum() compute all frustum pyramid planes (6 planes)
'
'  3 kind of tests are provided
'
'    Cube versus camera planes
'    Sphere versus Cam
'    Box Versus cam
'
'
'To see how this code can be used go at cQuest3D_Mesh class
'we compute if bounding box or sphere are visible before rendering polys
'
'
'
'=========================================================================================





Option Explicit

Private BoundingBOX(35) As QUEST3D_VERTEXCOLORED3D

Public Sub SetUpFrustum()

  Dim clip As D3DMATRIX
  Dim matView As D3DMATRIX
  Dim matProj As D3DMATRIX, J As Single

    obj_Device.GetTransform D3DTS_VIEW, matView
    obj_Device.GetTransform D3DTS_PROJECTION, matProj

    D3DXMatrixMultiply clip, matView, matProj

    clip.m11 = matView.m11 * matProj.m11 + matView.m12 * matProj.m21 + matView.m13 * matProj.m31 + matView.m14 * matProj.m41
    clip.m12 = matView.m11 * matProj.m12 + matView.m12 * matProj.m22 + matView.m13 * matProj.m32 + matView.m14 * matProj.m42
    clip.m13 = matView.m11 * matProj.m13 + matView.m12 * matProj.m23 + matView.m13 * matProj.m33 + matView.m14 * matProj.m43
    clip.m14 = matView.m11 * matProj.m14 + matView.m12 * matProj.m24 + matView.m13 * matProj.m34 + matView.m14 * matProj.m44

    clip.m21 = matView.m21 * matProj.m11 + matView.m22 * matProj.m21 + matView.m23 * matProj.m31 + matView.m24 * matProj.m41
    clip.m22 = matView.m21 * matProj.m12 + matView.m22 * matProj.m22 + matView.m23 * matProj.m32 + matView.m24 * matProj.m42
    clip.m23 = matView.m21 * matProj.m13 + matView.m22 * matProj.m23 + matView.m23 * matProj.m33 + matView.m24 * matProj.m43
    clip.m24 = matView.m21 * matProj.m14 + matView.m22 * matProj.m24 + matView.m23 * matProj.m34 + matView.m24 * matProj.m44

    clip.m31 = matView.m31 * matProj.m11 + matView.m32 * matProj.m21 + matView.m33 * matProj.m31 + matView.m34 * matProj.m41
    clip.m32 = matView.m31 * matProj.m12 + matView.m32 * matProj.m22 + matView.m33 * matProj.m32 + matView.m34 * matProj.m42
    clip.m33 = matView.m31 * matProj.m13 + matView.m32 * matProj.m23 + matView.m33 * matProj.m33 + matView.m34 * matProj.m43
    clip.m34 = matView.m31 * matProj.m14 + matView.m32 * matProj.m24 + matView.m33 * matProj.m34 + matView.m34 * matProj.m44

    clip.m41 = matView.m41 * matProj.m11 + matView.m42 * matProj.m21 + matView.m43 * matProj.m31 + matView.m44 * matProj.m41
    clip.m42 = matView.m41 * matProj.m12 + matView.m42 * matProj.m22 + matView.m43 * matProj.m32 + matView.m44 * matProj.m42
    clip.m43 = matView.m41 * matProj.m13 + matView.m42 * matProj.m23 + matView.m43 * matProj.m33 + matView.m44 * matProj.m43
    clip.m44 = matView.m41 * matProj.m14 + matView.m42 * matProj.m24 + matView.m43 * matProj.m34 + matView.m44 * matProj.m44

    'Right
    lpFRUST.PLANE(QUEST3D_RIGHT).A = clip.m14 - clip.m11
    lpFRUST.PLANE(QUEST3D_RIGHT).B = clip.m24 - clip.m21
    lpFRUST.PLANE(QUEST3D_RIGHT).c = clip.m34 - clip.m31
    lpFRUST.PLANE(QUEST3D_RIGHT).d = clip.m44 - clip.m41
    NormalizePlane lpFRUST.PLANE(), QUEST3D_RIGHT
    'Left
    lpFRUST.PLANE(QUEST3D_LEFT).A = clip.m14 + clip.m11
    lpFRUST.PLANE(QUEST3D_LEFT).B = clip.m24 + clip.m21
    lpFRUST.PLANE(QUEST3D_LEFT).c = clip.m34 + clip.m31
    lpFRUST.PLANE(QUEST3D_LEFT).d = clip.m44 + clip.m41
    NormalizePlane lpFRUST.PLANE(), QUEST3D_LEFT
    'Bottom
    lpFRUST.PLANE(QUEST3D_BOTTOM).A = clip.m14 + clip.m12
    lpFRUST.PLANE(QUEST3D_BOTTOM).B = clip.m24 + clip.m22
    lpFRUST.PLANE(QUEST3D_BOTTOM).c = clip.m34 + clip.m32
    lpFRUST.PLANE(QUEST3D_BOTTOM).d = clip.m44 + clip.m42
    NormalizePlane lpFRUST.PLANE(), QUEST3D_BOTTOM
    'Top
    lpFRUST.PLANE(QUEST3D_TOP).A = clip.m14 - clip.m12
    lpFRUST.PLANE(QUEST3D_TOP).B = clip.m24 - clip.m22
    lpFRUST.PLANE(QUEST3D_TOP).c = clip.m34 - clip.m32
    lpFRUST.PLANE(QUEST3D_TOP).d = clip.m44 - clip.m42
    NormalizePlane lpFRUST.PLANE(), QUEST3D_TOP
    'Back
    lpFRUST.PLANE(QUEST3D_BACK).A = clip.m14 - clip.m13
    lpFRUST.PLANE(QUEST3D_BACK).B = clip.m24 - clip.m23
    lpFRUST.PLANE(QUEST3D_BACK).c = clip.m34 - clip.m33
    lpFRUST.PLANE(QUEST3D_BACK).d = clip.m44 - clip.m43
    NormalizePlane lpFRUST.PLANE(), QUEST3D_BACK
    'Front
    lpFRUST.PLANE(QUEST3D_FRONT).A = clip.m14 + clip.m13
    lpFRUST.PLANE(QUEST3D_FRONT).B = clip.m24 + clip.m23
    lpFRUST.PLANE(QUEST3D_FRONT).c = clip.m34 + clip.m33
    lpFRUST.PLANE(QUEST3D_FRONT).d = clip.m44 + clip.m43
    NormalizePlane lpFRUST.PLANE(), QUEST3D_FRONT

    Data.FRUSTUM_HASCHANGED = False

End Sub

Private Function NormalizePlane(Frust() As D3DPLANE, ByVal Side As Long)

  Dim magnitude  As Single

    magnitude = Sqr(Frust(Side).A * Frust(Side).A + _
                Frust(Side).B * Frust(Side).B + _
                Frust(Side).c * Frust(Side).c)

    'If magnitude = 0 Then magnitude = 0.00001

    'Then we divide the plane's values by it's magnitude.
    'This makes it easier to work with.
    Frust(Side).A = Frust(Side).A / magnitude
    Frust(Side).B = Frust(Side).B / magnitude
    Frust(Side).c = Frust(Side).c / magnitude

    Frust(Side).d = Frust(Side).d / magnitude

End Function

'this routine is adapted from Gametutorials.com
Function Check_CubeInFrustum(m_Frustum As QUEST3D_FRUSTUM, x As Single, y As Single, z As Single, Size As Single) As Boolean

  Dim I As Integer

    For I = 0 To 5

        If (m_Frustum.PLANE(I).A * (x - Size) + m_Frustum.PLANE(I).B * (y - Size) + m_Frustum.PLANE(I).c * (z - Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x + Size) + m_Frustum.PLANE(I).B * (y - Size) + m_Frustum.PLANE(I).c * (z - Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x - Size) + m_Frustum.PLANE(I).B * (y + Size) + m_Frustum.PLANE(I).c * (z - Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x + Size) + m_Frustum.PLANE(I).B * (y + Size) + m_Frustum.PLANE(I).c * (z - Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x - Size) + m_Frustum.PLANE(I).B * (y - Size) + m_Frustum.PLANE(I).c * (z + Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x + Size) + m_Frustum.PLANE(I).B * (y - Size) + m_Frustum.PLANE(I).c * (z + Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x - Size) + m_Frustum.PLANE(I).B * (y + Size) + m_Frustum.PLANE(I).c * (z + Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * (x + Size) + m_Frustum.PLANE(I).B * (y + Size) + m_Frustum.PLANE(I).c * (z + Size) + m_Frustum.PLANE(I).d > 0) Then _
           GoTo CONTINUE

        Exit Function
CONTINUE:
    Next I

    Check_CubeInFrustum = True

End Function

'this routine is adapted from Gametutorials.com
Function Check_PointInFrustum(m_Frustum As QUEST3D_FRUSTUM, PointToTest As D3DVECTOR) As Boolean

  Dim I As Integer

    For I = 0 To 5

        If (m_Frustum.PLANE(I).A * PointToTest.x + m_Frustum.PLANE(I).B * PointToTest.y + m_Frustum.PLANE(I).c * PointToTest.z + m_Frustum.PLANE(I).d <= 0) Then

            'The point was behind a side, so it ISN'T in the frustum
            Check_PointInFrustum = False
            Exit Function
        End If

    Next I

    Check_PointInFrustum = True

End Function

'this routine is adapted from Gametutorials.com
Function Check_SphereInFrustum(m_Frustum As QUEST3D_FRUSTUM, SphereCenter As D3DVECTOR, ByVal Radius As Single) As Boolean

  Dim I As Integer

    For I = 0 To 5

        If (m_Frustum.PLANE(I).A * SphereCenter.x + m_Frustum.PLANE(I).B * SphereCenter.y + m_Frustum.PLANE(I).c * SphereCenter.z + m_Frustum.PLANE(I).d <= -Radius) Then

            ''// The distance was greater than the radius so the sphere is outside of the frustum
            Check_SphereInFrustum = False
            Exit Function
        End If

    Next I

    Check_SphereInFrustum = True

End Function

Function Check_SphereInFrustumEx(m_Frustum As QUEST3D_FRUSTUM, Sphere As QUEST3D_SPHERE) As Boolean

  Dim I As Integer


    For I = 0 To 5

        If (m_Frustum.PLANE(I).A * Sphere.SphereCenter.x + m_Frustum.PLANE(I).B * Sphere.SphereCenter.y + m_Frustum.PLANE(I).c * Sphere.SphereCenter.z + m_Frustum.PLANE(I).d <= -Sphere.Radius) Then

            ''// The distance was greater than the radius so the sphere is outside of the frustum
            Check_SphereInFrustumEx = False
            Exit Function
        End If

    Next I

    Check_SphereInFrustumEx = True

End Function

Function Check_BoxInFrustum(m_Frustum As QUEST3D_FRUSTUM, BoxMIN As D3DVECTOR, BoxMax As D3DVECTOR) As Boolean

  Dim I As Integer

    For I = 0 To 5

        If (m_Frustum.PLANE(I).A * BoxMIN.x + m_Frustum.PLANE(I).B * BoxMIN.y + m_Frustum.PLANE(I).c * BoxMIN.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMax.x + m_Frustum.PLANE(I).B * BoxMIN.y + m_Frustum.PLANE(I).c * BoxMIN.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMIN.x + m_Frustum.PLANE(I).B * BoxMax.y + m_Frustum.PLANE(I).c * BoxMIN.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMax.x + m_Frustum.PLANE(I).B * BoxMax.y + m_Frustum.PLANE(I).c * BoxMIN.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMIN.x + m_Frustum.PLANE(I).B * BoxMIN.y + m_Frustum.PLANE(I).c * BoxMax.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMax.x + m_Frustum.PLANE(I).B * BoxMIN.y + m_Frustum.PLANE(I).c * BoxMax.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMIN.x + m_Frustum.PLANE(I).B * BoxMax.y + m_Frustum.PLANE(I).c * BoxMax.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE
        If (m_Frustum.PLANE(I).A * BoxMax.x + m_Frustum.PLANE(I).B * BoxMax.y + m_Frustum.PLANE(I).c * BoxMax.z + m_Frustum.PLANE(I).d > 0) Then GoTo CONTINUE

       Exit Function
CONTINUE:
Next I

    Check_BoxInFrustum = True

End Function


Function Check_BoxInFrustumEx(m_Frustum As QUEST3D_FRUSTUM, BOX As QUEST3D_BOX) As Boolean
  
     Check_BoxInFrustumEx = Check_BoxInFrustum(m_Frustum, BOX.Vmin, BOX.Vmax)

End Function


Sub Render_Box(BoxMIN As D3DVECTOR, BoxMax As D3DVECTOR, Optional ByVal color As Long = &HFFFFFFFF)

    'front
    BoundingBOX(0) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(1) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(2) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(3) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(4) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(5) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    'back
    BoundingBOX(6) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(7) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(8) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(9) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(10) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(11) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    'left
    BoundingBOX(12) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(13) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(14) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(15) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(16) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(17) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    'right
    BoundingBOX(18) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(19) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(20) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(21) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(22) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(23) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    'top
    BoundingBOX(24) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(25) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(26) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(27) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMax.z, color)
    BoundingBOX(28) = Make_Vertex3D(BoxMax.x, BoxMax.y, BoxMIN.z, color)
    BoundingBOX(29) = Make_Vertex3D(BoxMIN.x, BoxMax.y, BoxMIN.z, color)
    'bottom
    BoundingBOX(30) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(31) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMax.z, color)
    BoundingBOX(32) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(33) = Make_Vertex3D(BoxMIN.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(34) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMIN.z, color)
    BoundingBOX(35) = Make_Vertex3D(BoxMax.x, BoxMIN.y, BoxMax.z, color)

    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_CULLMODE
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_FILLMODE
    LpGLOBAL_QUEST3D.Push_Renderstate QUEST3DRS_LIGHTING
    
    

    obj_Device.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    obj_Device.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
    obj_Device.SetRenderState D3DRS_LIGHTING, 0

    obj_Device.SetTransform D3DTS_WORLD, ID_MATRIX

    obj_Device.SetVertexShader QUEST3D_FVFVERTEXCOLORED3D
    obj_Device.SetTexture 0, Nothing
    
    obj_Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, 12, BoundingBOX(0), Len(BoundingBOX(0))

    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_CULLMODE
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_FILLMODE
    LpGLOBAL_QUEST3D.Pop_Renderstate QUEST3DRS_LIGHTING

    Add_Verti 36
    Add_Tri 12

End Sub







