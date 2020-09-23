Attribute VB_Name = "Module_Collision"

'==========================================================================================
'
'   This module contains all useful methods to compute collision
'
'
'To see how this code can be used go at cQuest3D_Mesh class
'we compute fix and sliding collision
'
'VERY IMPORTANT SOME PART OF THA CODE IS ADAPTED
'HUGELY FROM Gametutorials.com sample code
'DigiBen you are a great guy
'=========================================================================================

'box

Option Explicit
Public Type QUEST3D_BOX
    Vmin As D3DVECTOR
    Vmax As D3DVECTOR
End Type

Public Type QUEST3D_DYNAMIC_BOX
    BOX_IDENTITY As QUEST3D_BOX
    BOX_TRANSFORMED As QUEST3D_BOX
End Type

Public Type QUEST3D_SPHERE
    SphereCenter As D3DVECTOR
    Radius As Single
End Type

Public Type QUEST3D_DYNAMIC_SPHERE
    SPHERE_IDENTITY As QUEST3D_SPHERE
    SPHERE_TRANSFORMED As QUEST3D_SPHERE
End Type

Public Type QUEST3D_LINE
    LineStart As D3DVECTOR
    LineEnd As D3DVECTOR
End Type

Public Type QUEST3D_RAY
    RayStart As D3DVECTOR
    RayDirection As D3DVECTOR
End Type

Public Enum QUEST3D_INTERSECTIONSTATUS
    QUEST3D_INTERSECTION_BEHIND = 0
    QUEST3D_INTERSECTION_INTERSECTS = 1
    QUEST3D_INTERSECTION_FRONT = 2

End Enum

Function Check_PointInBox(vPoint As D3DVECTOR, BOX As QUEST3D_BOX) As Boolean

    If ((vPoint.x >= BOX.Vmin.x) And _
       (vPoint.x <= BOX.Vmax.x) And _
       (vPoint.y >= BOX.Vmin.y) And _
       (vPoint.y <= BOX.Vmax.y) And _
       (vPoint.z >= BOX.Vmin.z) And _
       (vPoint.z <= BOX.Vmax.z)) Then

        Check_PointInBox = True
    End If

    '    If vPoint.x < BOX.Vmin.x Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    If vPoint.y < BOX.Vmin.y Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    If vPoint.z < BOX.Vmin.z Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    If vPoint.x > BOX.Vmax.x Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    If vPoint.y > BOX.Vmax.y Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    If vPoint.z > BOX.Vmax.z Then
    '        Check_PointInBox = False
    '        Exit Function
    '    End If
    '
    '    Check_PointInBox = True

End Function

Function Check_PointInSphere(vPoint As D3DVECTOR, Sphere As QUEST3D_SPHERE) As Boolean

  Dim TempVect As D3DVECTOR
  Dim D As Single

    TempVect.x = vPoint.x - Sphere.SphereCenter.x
    TempVect.y = vPoint.y - Sphere.SphereCenter.y
    TempVect.z = vPoint.z - Sphere.SphereCenter.z

    D = Vector_Magnitude(TempVect)

    If (D <= Sphere.Radius) Then
        Check_PointInSphere = True
      Else
        Check_PointInSphere = False
    End If

End Function

Function Check_CollisionRaySphere(Center As D3DVECTOR, Radius As Single, RayPosition As D3DVECTOR, RayDirection As D3DVECTOR) As Boolean

    Check_CollisionRaySphere = obj_D3DX.SphereBoundProbe(Center, Radius, RayPosition, RayDirection)

End Function

Function Check_CollisionRaySphereEx(Ray As QUEST3D_RAY, Sphere As QUEST3D_SPHERE) As Boolean

    Check_CollisionRaySphereEx = obj_D3DX.SphereBoundProbe(Sphere.SphereCenter, Sphere.Radius, Ray.RayStart, Ray.RayDirection)

End Function

' returns the addition of 2 vectors
Function Vector_Add(A As D3DVECTOR, B As D3DVECTOR) As D3DVECTOR

    Vector_Add.x = A.x + B.x
    Vector_Add.y = A.y + B.y
    Vector_Add.z = A.z + B.z

End Function

Function Vector_Scale(A As D3DVECTOR, ByVal ScaleValue As Single) As D3DVECTOR

    Vector_Scale.x = A.x * ScaleValue
    Vector_Scale.y = A.y * ScaleValue
    Vector_Scale.z = A.z * ScaleValue

End Function

' returns the subraction of one vector from another
Function Vector_Subtract(A As D3DVECTOR, B As D3DVECTOR) As D3DVECTOR

    Vector_Subtract.x = A.x - B.x
    Vector_Subtract.y = A.y - B.y
    Vector_Subtract.z = A.z - B.z

End Function

' puts the subtraction of two vectors into a destination vector
Sub Vector_SubtractEx(Dest As D3DVECTOR, A As D3DVECTOR, B As D3DVECTOR)

    Dest.x = A.x - B.x
    Dest.y = A.y - B.y
    Dest.z = A.z - B.z

End Sub

'Public Function Vector_Distance(VA As D3DVECTOR, VB As D3DVECTOR) As Single
'
'    Vector_Distance = Sqr((VB.x - VA.x) ^ 2 + (VB.y - VA.y) ^ 2 + (VB.z - VA.z) ^ 2)
'
'End Function

' puts the cross product of two vectors in a destination vector
Sub Vector_CrossProduct(Dest As D3DVECTOR, A As D3DVECTOR, B As D3DVECTOR)

    Dest.x = A.y * B.z - A.z * B.y
    Dest.y = A.z * B.x - A.x * B.z
    Dest.z = A.x * B.y - A.y * B.x

End Sub

' returns the dot product of two vectors
Function Vector_DotProduct(A As D3DVECTOR, B As D3DVECTOR) As Single

    Vector_DotProduct = A.x * B.x + A.y * B.y + A.z * B.z

End Function

'=================================
' VectorNormalize
'=================================
' creates a vector of length 1 in the same direction
'
Sub Vector_Normalize(Dest As D3DVECTOR)

    On Local Error Resume Next
      Dim L As Double
        L = Dest.x * Dest.x + Dest.y * Dest.y + Dest.z * Dest.z
        L = Sqr(L)
        If L = 0 Then
            Dest.x = 0
            Dest.y = 0
            Dest.z = 0
            Exit Sub
        End If
        Dest.x = Dest.x / L
        Dest.y = Dest.y / L
        Dest.z = Dest.z / L

End Sub

Function Vector_CalculateNormal(p0 As D3DVECTOR, p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR

  '//0. Any variables

  Dim vNorm As D3DVECTOR
  Dim v01 As D3DVECTOR
  Dim v02 As D3DVECTOR

    '//1. Subtract vectors

    Vector_SubtractEx v01, p1, p0
    Vector_SubtractEx v02, p2, p0

    '//2. Perform Cross Product

    D3DXVec3Cross vNorm, v01, v02

    '//3. Normalize
    D3DXVec3Normalize vNorm, vNorm

    '//4. Return normal
    Vector_CalculateNormal = vNorm

End Function

Function Vector_Cross(p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR

  'The X value for the vector is:  (V1.y * V2.z) - (V1.z * V2.y)                                                    'Get the X value

    Vector_Cross.x = ((p1.y * p2.z) - (p1.z * p2.y))

    'The Y value for the vector is:  (V1.z * V2.x) - (V1.x * V2.z)
    Vector_Cross.y = ((p1.z * p2.x) - (p1.x * p2.z))

    'The Z value for the vector is:  (V1.x * V2.y) - (V1.y * V2.x)
    Vector_Cross.z = ((p1.x * p2.y) - (p1.y * p2.x))

End Function

Function Vector_Magnitude(vNormal As D3DVECTOR)

    Vector_Magnitude = Sqr((vNormal.x * vNormal.x) + (vNormal.y * vNormal.y) + (vNormal.z * vNormal.z))

End Function

Function Vector_Normal(vPoly() As D3DVECTOR) As D3DVECTOR

  Dim V1 As D3DVECTOR
  Dim V2 As D3DVECTOR
  Dim vNormal As D3DVECTOR

    V1.x = vPoly(2).x - vPoly(0).x
    V1.y = vPoly(2).y - vPoly(0).y
    V1.z = vPoly(2).z - vPoly(0).z

    V2.x = vPoly(1).x - vPoly(0).x
    V2.y = vPoly(1).y - vPoly(0).y
    V2.z = vPoly(1).z - vPoly(0).z

    vNormal = Vector_Cross(V1, V2)
    Vector_Normalize vNormal
    Vector_Normal = vNormal

End Function

Function Vector_ClosestPointOnLine(Va As D3DVECTOR, Vb As D3DVECTOR, vPoint As D3DVECTOR) As D3DVECTOR

  Dim V1 As D3DVECTOR
  Dim V2 As D3DVECTOR
  Dim V3 As D3DVECTOR
  Dim vClosestPoint As D3DVECTOR
  Dim T As Single
  Dim D As Single

    'Create the vector from end point vA to our point vPoint.
    V1 = Vector_Subtract(vPoint, Va)

    'Create a normalized direction vector from end point vA to end point vB

    D3DXVec3Normalize V2, Vector_Subtract(Vb, Va)

    'Use the distance formula to find the distance of the line segment (or magnitude)
    D = Vector_Distance(Va, Vb)

    'Using the dot product, we project the V1 onto the vector V2.
    'This essentially gives us the distance from our projected vector from vA.
    T = D3DXVec3Dot(V2, V1)

    'If our projected distance from vA, "t", is less than or equal to 0, it must
    'be closest to the end point vA.  We want to Vector_ClosestPointOnLine=this end point.
    If (T <= 0) Then _
       Vector_ClosestPointOnLine = Va

    'If our projected distance from vA, "t", is greater than or equal to the magnitude
    'or distance of the line segment, it must be closest to the end point vB.  So, Vector_ClosestPointOnLine=vB.
    If (T >= D) Then _
       Vector_ClosestPointOnLine = Vb

    'Here we create a vector that is of length t and in the direction of V2
    V3.x = V2.x * T
    V3.y = V2.y * T
    V3.z = V2.z * T

    'To find the closest point on the line segment, we just add V3 to the original
    'end point vA.
    vClosestPoint.x = Va.x + V3.x
    vClosestPoint.y = Va.y + V3.y
    vClosestPoint.z = Va.z + V3.z

    'Vector_ClosestPointOnLine=the closest point on the line segment
    Vector_ClosestPointOnLine = vClosestPoint

End Function

Function Vector_PlaneDistance(vNormal As D3DVECTOR, vPoint As D3DVECTOR) As Single

    Vector_PlaneDistance = -((vNormal.x * vPoint.x) + (vNormal.y * vPoint.y) + (vNormal.z * vPoint.z))

End Function

Function Vector_IntersectedPlane(vPoly() As D3DVECTOR, vLine() As D3DVECTOR, vNormal As D3DVECTOR, ByRef originDistance As Single) As Boolean

  Dim distance1  As Single, distance2  As Single                ' The distances from the 2 points of the line from the plane

    vNormal = Vector_Normal(vPoly)                            ' We need to get the normal of our plane to go any further

    ' Let's find the distance our plane is from the origin.  We can find this value
    ' from the normal to the plane (polygon) and any point that lies on that plane (Any vertex)
    originDistance = Vector_PlaneDistance(vNormal, vPoly(0))

    ' Get the distance from point1 from the plane using: Ax + By + Cz + D = (The distance from the plane)

    distance1 = ((vNormal.x * vLine(0).x) + _
                (vNormal.y * vLine(0).y) + _
                (vNormal.z * vLine(0).z)) + originDistance    ' Cz + D

    ' Get the distance from point2 from the plane using Ax + By + Cz + D = (The distance from the plane)

    distance2 = ((vNormal.x * vLine(1).x) + _
                (vNormal.y * vLine(1).y) + _
                (vNormal.z * vLine(1).z)) + originDistance    ' Cz + D

    ' Now that we have 2 distances from the plane, if we times them together we either
    ' get a positive or negative number.  If it's a negative number, that means we collided!
    ' This is because the 2 points must be on either side of the plane (IE. -1 * 1 = -1).

    If (distance1 * distance2 >= 0) Then _
       Vector_IntersectedPlane = False                      ' Vector_IntersectedPlane=false if each point has the same sign.  -1 and 1 would mean each point is on either side of the plane.  -1 -2 or 3 4 wouldn't...

    Vector_IntersectedPlane = True                          ' The line intersected the plane, Vector_IntersectedPlane=TRUE

End Function

Function Vector_AngleBetweenVectors(Vector1 As D3DVECTOR, Vector2 As D3DVECTOR) As Single

  ' Get the dot product of the vectors

  Dim DotProduct As Single
  Dim vectorsMagnitude As Single
  Dim Angle As Single

    DotProduct = D3DXVec3Dot(Vector1, Vector2)

    ' Get the product of both of the vectors magnitudes
    vectorsMagnitude = Vector_Magnitude(Vector1) * Vector_Magnitude(Vector2)

    ' Get the angle in radians between the 2 vectors
    Angle = ArcCos(DotProduct / vectorsMagnitude)

    ' Here we make sure that the angle is not a -1E+27 number, which means indefinate
    If (Angle < -1E+27) Then _
       Vector_AngleBetweenVectors = 0

    ' Return the angle in radians
    Vector_AngleBetweenVectors = Angle

End Function

Function Vector_IntersectionPoint(vNormal As D3DVECTOR, vLine() As D3DVECTOR, distance As Single) As D3DVECTOR

  Dim vPoint As D3DVECTOR, vLineDir As D3DVECTOR                  ' Variables to hold the point and the line's direction
  Dim Numerator As Single
  Dim Denominator As Single
  Dim dist As Single

    ' 1)  First we need to get the vector of our line, Then normalize it so it's a length of 1
    vLineDir = Vector_Subtract(vLine(1), vLine(0))      ' Get the Vector of the line
    ' Normalize the lines vector
    D3DXVec3Normalize vLineDir, vLineDir

    ' 2) Use the plane equation (distance = Ax + By + Cz + D) to find the
    ' distance from one of our points to the plane.
    Numerator = -(vNormal.x * vLine(0).x + _
                vNormal.y * vLine(0).y + _
                vNormal.z * vLine(0).z + distance)

    ' 3) If we take the dot product between our line vector and the normal of the polygon,
    Denominator = D3DXVec3Dot(vNormal, vLineDir)       ' Get the dot product of the line's vector and the normal of the plane

    ' Since we are using division, we need to make sure we don't get a divide by zero error
    ' If we do get a 0, that means that there are INFINATE points because the the line is
    ' on the plane (the normal is perpendicular to the line - (Normal.Vector = 0)).
    ' In this case, we should just return any point on the line.

    If (Denominator = 0) Then _
       Vector_IntersectionPoint = vLine(0)                      ' Return an arbitrary point on the line

    dist = Numerator / Denominator             ' Divide to get the multiplying (percentage) factor

    ' Now, like we said above, we times the dist by the vector, then add our arbitrary point.
    vPoint.x = (vLine(0).x + (vLineDir.x * dist))
    vPoint.y = (vLine(0).y + (vLineDir.y * dist))
    vPoint.z = (vLine(0).z + (vLineDir.z * dist))

    Vector_IntersectionPoint = vPoint                            ' Return the intersection point

End Function

Function Vector_InsidePolygon(vIntersection As D3DVECTOR, Poly() As D3DVECTOR, ByVal verticeCount As Long) As Boolean

  Dim I As Long
  Dim Angle As Single
  Dim Va As D3DVECTOR
  Dim Vb As D3DVECTOR

    For I = 0 To verticeCount - 1      ' Go in a circle to each vertex and get the angle between

        Va = Vector_Subtract(Poly(I), vIntersection)            ' Subtract the intersection point from the current vertex
        ' Subtract the point from the next vertex
        Vb = Vector_Subtract(Poly((I + 1) Mod verticeCount), vIntersection)

        Angle = Angle + Vector_AngleBetweenVectors(Va, Vb) ' Find the angle between the 2 vectors and add them all up as we go along
    Next I
    ' If the angle is greater than 2 PI, (360 degrees)
    If (Angle >= (0.99 * (2# * QUEST3D_PI))) Then
        Vector_InsidePolygon = True                          ' The point is inside of the polygon
        Exit Function
    End If

    Vector_InsidePolygon = False                             ' If you get here, it obviously wasn't inside the polygon, so Vector_InsidePolygon=FALSE

End Function

Function Vector_IntersectedPolygon(vPoly() As D3DVECTOR, vLine() As D3DVECTOR, ByVal verticeCount As Long) As Boolean

  Dim vNormal As D3DVECTOR, vIntersection As D3DVECTOR
  Dim originDistance As Single

    ' First, make sure our line intersects the plane
    ' Reference   ' Reference
    If (Not Vector_IntersectedPlane(vPoly, vLine, vNormal, originDistance)) Then
        Vector_IntersectedPolygon = False
        Exit Function

    End If

    ' Now that we have our normal and distance passed back from IntersectedPlane(),
    ' we can use it to calculate the intersection point.
    vIntersection = Vector_IntersectionPoint(vNormal, vLine, originDistance)

    ' Now that we have the intersection point, we need to test if it's inside the polygon.
    If (Vector_InsidePolygon(vIntersection, vPoly, verticeCount)) Then
        Vector_IntersectedPolygon = True
        Exit Function

    End If
    Vector_IntersectedPolygon = False                             ' There was no collision, so return false

End Function

Function Vector_ClassifySphere(vCenter As D3DVECTOR, _
                  vNormal As D3DVECTOR, vPoint As D3DVECTOR, ByVal Radius As Single, ByRef distance As Single) As QUEST3D_INTERSECTIONSTATUS

  Dim D As Single

    ' First we need to find the distance our polygon plane is from the origin.
    D = Vector_PlaneDistance(vNormal, vPoint)

    ' Here we use the famous distance formula to find the distance the center point
    ' of the sphere is from the polygon's plane.
    distance = (vNormal.x * vCenter.x + vNormal.y * vCenter.y + vNormal.z * vCenter.z + D)

    ' If the absolute value of the distance we just found is less than the radius,
    ' the sphere intersected the plane.
    If (Abs(distance) < Radius) Then
        Vector_ClassifySphere = QUEST3D_INTERSECTION_INTERSECTS
        Exit Function

        ' Else, if the distance is greater than or equal to the radius, the sphere is
        ' completely in FRONT of the plane.
      ElseIf (distance >= Radius) Then
        Vector_ClassifySphere = QUEST3D_INTERSECTION_FRONT
        Exit Function
    End If

    ' If the sphere isn't intersecting or in FRONT of the plane, it must be BEHIND
    Vector_ClassifySphere = QUEST3D_INTERSECTION_BEHIND

End Function

Function Vector_GetCollisionOffset(vNormal As D3DVECTOR, ByVal Radius As Single, ByVal distance As Single) As D3DVECTOR

  Dim vOffset As D3DVECTOR
  Dim distanceOver As Single

    ' If our distance is greater than zero, we are in front of the polygon
    If (distance > 0) Then

        ' Find the distance that our sphere is overlapping the plane, then
        ' find the direction vector to move our sphere.
        distanceOver = Radius - distance
        vOffset.x = vNormal.x * distanceOver
        vOffset.y = vNormal.y * distanceOver
        vOffset.z = vNormal.z * distanceOver

      Else
        ' Find the distance that our sphere is overlapping the plane, then
        ' find the direction vector to move our sphere.
        distanceOver = Radius + distance
        vOffset.x = vNormal.x * -distanceOver
        vOffset.y = vNormal.y * -distanceOver
        vOffset.z = vNormal.z * -distanceOver
    End If

    ' Return the offset we need to move back to not be intersecting the polygon.
    Vector_GetCollisionOffset = vOffset

End Function

Function Vector_EdgeSphereCollision(vCenter As D3DVECTOR, _
                  vPolygon() As D3DVECTOR, ByVal vertexCount As Long, ByVal Radius As Single) As Boolean

  Dim vPoint As D3DVECTOR
  Dim I As Long
  Dim distance As Single

    ' This function takes in the sphere's center, the polygon's vertices, the vertex count
    ' and the radius of the sphere.  We will Vector_EdgeSphereCollision=true from this function if the sphere
    ' is intersecting any of the edges of the polygon.

    ' Go through all of the vertices in the polygon
    For I = 0 To vertexCount - 1

        ' This returns the closest point on the current edge to the center of the sphere.
        vPoint = Vector_ClosestPointOnLine(vPolygon(I), vPolygon((I + 1) Mod vertexCount), vCenter)

        ' Now, we want to calculate the distance between the closest point and the center
        distance = Vector_Distance(vPoint, vCenter)

        ' If the distance is less than the radius, there must be a collision so Vector_EdgeSphereCollision=true
        If (distance < Radius) Then
            Vector_EdgeSphereCollision = True
            Exit Function
        End If
    Next I

    ' The was no intersection of the sphere and the edges of the polygon
    Vector_EdgeSphereCollision = False

End Function

Function Check_CollisionSphereVERTEX2(SphereCenter As D3DVECTOR, ByVal Radius As Single, VERT() As QUEST3D_VERTEX2, ByVal NumVERT As Long) As Boolean

  Dim vTriangle(2) As D3DVECTOR
  Dim I As Long
  Dim distance As Single
  Dim classification As QUEST3D_INTERSECTIONSTATUS
  Dim vOffset As D3DVECTOR
  Dim vNormal As D3DVECTOR
  Dim vIntersection As D3DVECTOR

    For I = 0 To NumVERT - 1 Step 3

        'Store of the current triangle we testing
        vTriangle(0) = VERT(I).Position
        vTriangle(1) = VERT(I + 1).Position
        vTriangle(2) = VERT(I + 2).Position

        '1) STEP ONE - Finding the sphere's classification

        'We want the normal to the current polygon being checked
        vNormal = Vector_Normal(vTriangle)

        'This will store the distance our sphere is from the plane
        distance = 0

        'This is where we determine if the sphere is in FRONT, BEHIND, or INTERSECTS the plane
        classification = Vector_ClassifySphere(SphereCenter, vNormal, vTriangle(0), Radius, distance)

        'If the sphere intersects the polygon's plane, then we need to check further
        If (classification = QUEST3D_INTERSECTION_INTERSECTS) Then

            '2) STEP TWO - Finding the psuedo intersection point on the plane

            'Now we want to project the sphere's center onto the triangle's plane

            vOffset.x = vNormal.x * distance
            vOffset.y = vNormal.y * distance
            vOffset.z = vNormal.z * distance

            'Once we have the offset to the plane, we just subtract it from the center
            'of the sphere.  "vIntersection" is now a point that lies on the plane of the triangle.
            vIntersection.x = SphereCenter.x - vOffset.x
            vIntersection.y = SphereCenter.y - vOffset.y
            vIntersection.z = SphereCenter.z - vOffset.z

            '3) STEP THREE - Check if the intersection point is inside the triangles perimeter

            If (Vector_InsidePolygon(vIntersection, vTriangle, 3) Or _
               Vector_EdgeSphereCollision(SphereCenter, vTriangle, 3, Radius / 2)) Then

                Check_CollisionSphereVERTEX2 = True
                Exit Function

            End If
        End If
    Next I

End Function

Function Check_CollisionSphereVERTEX2Sliding(SphereCenter As D3DVECTOR, ByVal Radius As Single, VERT() As QUEST3D_VERTEX2, ByVal NumVERT As Long, DestVector As D3DVECTOR) As Boolean

  Dim vTriangle(2) As D3DVECTOR
  Dim I As Long
  Dim distance As Single
  Dim classification As QUEST3D_INTERSECTIONSTATUS
  Dim vOffset As D3DVECTOR
  Dim vNormal As D3DVECTOR
  Dim vIntersection As D3DVECTOR

    For I = 0 To NumVERT - 1 Step 3

        'Store of the current triangle we testing
        vTriangle(0) = VERT(I).Position
        vTriangle(1) = VERT(I + 1).Position
        vTriangle(2) = VERT(I + 2).Position

        '1) STEP ONE - Finding the sphere's classification

        'We want the normal to the current polygon being checked
        vNormal = Vector_Normal(vTriangle)

        'This will store the distance our sphere is from the plane
        distance = 0

        'This is where we determine if the sphere is in FRONT, BEHIND, or INTERSECTS the plane
        classification = Vector_ClassifySphere(SphereCenter, vNormal, vTriangle(0), Radius, distance)

        'If the sphere intersects the polygon's plane, then we need to check further
        If (classification = QUEST3D_INTERSECTION_INTERSECTS) Then

            '2) STEP TWO - Finding the psuedo intersection point on the plane

            'Now we want to project the sphere's center onto the triangle's plane

            vOffset.x = vNormal.x * distance
            vOffset.y = vNormal.y * distance
            vOffset.z = vNormal.z * distance

            'Once we have the offset to the plane, we just subtract it from the center
            'of the sphere.  "vIntersection" is now a point that lies on the plane of the triangle.
            vIntersection.x = SphereCenter.x - vOffset.x
            vIntersection.y = SphereCenter.y - vOffset.y
            vIntersection.z = SphereCenter.z - vOffset.z

            '3) STEP THREE - Check if the intersection point is inside the triangles perimeter

            If (Vector_InsidePolygon(vIntersection, vTriangle, 3) Or _
               Vector_EdgeSphereCollision(SphereCenter, vTriangle, 3, Radius / 2)) Then

                vOffset = Vector_GetCollisionOffset(vNormal, Radius, distance)

                'Now that we have the offset, we want to ADD it to the position and
                'view vector in our camera.  This pushes us back off of the plane.  We

                DestVector.x = SphereCenter.x + vOffset.x
                DestVector.y = SphereCenter.y + vOffset.y
                DestVector.z = SphereCenter.z + vOffset.z

                Check_CollisionSphereVERTEX2Sliding = True
                Exit Function

            End If
        End If
    Next I

End Function

'many collision methods can be added over this foundation
