Attribute VB_Name = "Module_Util"
Option Explicit
'==============================================================================================================
'
'  THIS MODULE CONTAINS VERY USEFUL AND HARD HAND CODED METHODS, IF YOU
'   WANT TO USE THEM, GIVE CREDITS TO MY PERSON (Polaris),johna_pop@yahoo.fr
'
'
'
'
'
'       In this module we define all useful methods
'
'================================================================================================

'for file
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Const INVALID_HANDLE_VALUE = -1

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Function Make_Vertex2D(ByVal Xpos As Single, ByVal Ypos As Single, ByVal color As Long) As QUEST3D_VERTEXCOLORED2D

    With Make_Vertex2D
        .color = color
        .Position.z = 0.5
        .Rhw = 1
        .Position.x = Xpos
        .Position.y = Ypos

    End With

End Function

Function Make_Vertex3D(ByVal Xpos As Single, ByVal Ypos As Single, ByVal Zpos As Single, ByVal color As Long) As QUEST3D_VERTEXCOLORED3D

    With Make_Vertex3D
        .color = color
        .Position.z = Zpos
        .Position.x = Xpos
        .Position.y = Ypos

    End With

End Function

Function Make_LVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal color As Long, ByVal Specular As Long, ByVal Tu As Single, _
                  ByVal TV As Single) As QUEST3D_LVERTEX

    With Make_LVertex
        .Specular = Specular
        .Tu = Tu
        .TV = TV
        .x = x
        .y = y
        .z = z
        .color = color
    End With

End Function

Function Make_TLVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal Rhw As Single, _
                  ByVal color As Long, ByVal Specular As Long, ByVal Tu As Single, _
                  ByVal TV As Single) As QUEST3D_TLVERTEX

    Make_TLVertex.x = x
    Make_TLVertex.y = y
    Make_TLVertex.z = z
    Make_TLVertex.Rhw = Rhw
    Make_TLVertex.color = color
    Make_TLVertex.Specular = Specular
    Make_TLVertex.Tu = Tu
    Make_TLVertex.TV = TV

End Function

Function Make_ColorRGB(ByVal RedChanel As Byte, ByVal GreenChanel As Byte, ByVal BlueChanel As Byte) As Long

    Make_ColorRGB = D3DColorXRGB(RedChanel, GreenChanel, BlueChanel)

End Function

Function Make_ColorRGBA(ByVal RedChanel As Byte, ByVal GreenChanel As Byte, ByVal BlueChanel As Byte, Optional ByVal Alpha As Byte = 255) As Long

    Make_ColorRGBA = D3DColorRGBA(RedChanel, GreenChanel, BlueChanel, Alpha)

End Function

Function Make_ColorRGBAEx(ByVal RedChanel As Single, ByVal GreenChanel As Single, ByVal BlueChanel As Single, Optional ByVal Alpha As Single = 1) As Long

    Make_ColorRGBAEx = D3DColorRGBA(RedChanel * 255, GreenChanel * 255, BlueChanel * 255, Alpha * 255)

End Function

Function LONGtoD3DCOLORVALUE(ByVal color As Long) As D3DCOLORVALUE

  Dim A As Long, R As Long, g As Long, B As Long

    If color < 0 Then
        A = ((color And (&H7F000000)) / (2 ^ 24)) Or &H80&
      Else
        A = color / (2 ^ 24)
    End If
    R = (color And &HFF0000) / (2 ^ 16)
    g = (color And &HFF00&) / (2 ^ 8)
    B = (color And &HFF&)

    LONGtoD3DCOLORVALUE.A = A / 255
    LONGtoD3DCOLORVALUE.R = R / 255
    LONGtoD3DCOLORVALUE.g = g / 255
    LONGtoD3DCOLORVALUE.B = B / 255

End Function

'============================================================================
'
'MATRIX methods
'
'
'=============================================================================
Function Matrix_Get(ByVal Xscal As Single, ByVal Yscal As Single, ByVal Zscal As Single, ByVal Xrot As Single, ByVal Yrot As Single, ByVal Zrot As Single, ByVal Xmov As Single, ByVal Ymov As Single, ByVal Zmov As Single) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_Get = MatZ

End Function

Function Matrix_Ret(MatRet As D3DMATRIX, ByVal Xscal As Single, ByVal Yscal As Single, ByVal Zscal As Single, ByVal Xrot As Single, ByVal Yrot As Single, ByVal Zrot As Single, ByVal Xmov As Single, ByVal Ymov As Single, ByVal Zmov As Single) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    MatRet = MatZ

End Function

'----------------------------------------
'Name: Matrix_GetEX
'Object: Matrix
'Event: GetEX
'----------------------------------------
'----------------------------------------
'Name: Matrix_GetEX
'Object: Matrix
'Event: GetEX
'Description:
'----------------------------------------
Function Matrix_GetEX(Vscal As D3DVECTOR, vRot As D3DVECTOR, Vtrans As D3DVECTOR) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Vscal.x, Vscal.y, Vscal.z)
    Call MRotate(ROTz, vRot.x, vRot.y, vRot.z)
    Call D3DXMatrixTranslation(MOVz, Vtrans.x, Vtrans.y, Vtrans.z)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_GetEX = MatZ

End Function

Sub MRotate(DestMat As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single)

  Dim MatX As D3DMATRIX
  Dim MatY As D3DMATRIX
  Dim MatZ As D3DMATRIX
  Dim MatTemp As D3DMATRIX

    D3DXMatrixIdentity MatTemp
    D3DXMatrixIdentity MatX
    D3DXMatrixIdentity MatY
    D3DXMatrixIdentity MatZ

    D3DXMatrixRotationX MatX, nValueX
    D3DXMatrixRotationY MatY, nValueY
    D3DXMatrixRotationZ MatZ, nValueZ

    D3DXMatrixMultiply MatTemp, MatX, MatY
    D3DXMatrixMultiply MatTemp, MatTemp, MatZ

    DestMat = MatTemp

End Sub

'========================================================================
'
'Vector
'====================================================================

Function Vector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR

    Vector.x = x
    Vector.y = y
    Vector.z = z

End Function

Function Vector2D(ByVal x As Single, ByVal y As Single) As D3DVECTOR2

    Vector2D.x = x
    Vector2D.y = y

End Function

Function Vector_Compare(V1 As D3DVECTOR, V2 As D3DVECTOR, ByVal Tolerance As Double) As Boolean

    Vector_Compare = True

    If (Abs(V2.x - V1.x) > Tolerance) Then _
       Vector_Compare = False
    If (Abs(V2.y - V1.y) > Tolerance) Then _
       Vector_Compare = False
    If (Abs(V2.z - V1.z) > Tolerance) Then _
       Vector_Compare = False

End Function

Function Vector_Distance(vPoint1 As D3DVECTOR, vPoint2 As D3DVECTOR) As Single

  ' Distance = sqrt(  (P2.x - P1.x)^2 + (P2.y - P1.y)^2 + (P2.z - P1.z)^2 )

    Vector_Distance = Sqr((vPoint2.x - vPoint1.x) * (vPoint2.x - vPoint1.x) + _
                      (vPoint2.y - vPoint1.y) * (vPoint2.y - vPoint1.y) + _
                      (vPoint2.z - vPoint1.z) * (vPoint2.z - vPoint1.z))

End Function

Function RotationToDirection(ByVal pitch As Single, ByVal yaw As Single) As D3DVECTOR

    RotationToDirection.x = -Sin(yaw) * Cos(pitch)
    RotationToDirection.y = Sin(pitch)
    RotationToDirection.z = Cos(pitch) * Cos(yaw)

End Function

Sub GetRotationFromTO(LookFrom As D3DVECTOR, LookAt As D3DVECTOR, ByRef LookRotation As D3DVECTOR)

  Dim currentVelocity As D3DVECTOR

    D3DXVec3Subtract currentVelocity, LookAt, LookFrom

    LookRotation.x = Atn(currentVelocity.y / _
                     Sqr(currentVelocity.z * currentVelocity.z _
                     + currentVelocity.x * currentVelocity.x))

    LookRotation.y = -ArcTan(currentVelocity.x, currentVelocity.z)

    LookRotation.z = 0

End Sub

Function ArcTan(ByVal y As Single, ByVal x As Single)

  Dim Azimuth As Single
  Dim EPSILON As Double

    EPSILON = 0.000000000001

    Azimuth = 3.14159265358979 / 2

    If (Abs(x) > EPSILON) Then

        Azimuth = Atn(y / x)

        If (x < 0#) Then
            Azimuth = Azimuth + 3.14159265358979
          ElseIf (y < 0#) Then
            Azimuth = Azimuth + 2 * 3.14159265358979
        End If

      ElseIf (y < 0#) Then
        Azimuth = Azimuth * 3#

    End If

    ArcTan = Azimuth

End Function


Function ArcCos(ByVal x As Single) As Single
 
    If Abs(x) < 1 Then
        ArcCos = Atn(-x / Sqr(1 - x * x)) + QUEST3D_PI / 2
    ElseIf x = 1 Then
        ArcCos = 0
    ElseIf x = -1 Then
        ArcCos = QUEST3D_PI
    End If

  
End Function

Public Function FloatToDWord(ByVal f As Single) As Long

    DXCopyMemory FloatToDWord, f, 4

End Function

'================================================================================
'File manipulation
'
'
'===================================================================================
Function Get_fileName(LongNAME As String) As String

  Dim St As String
  Dim ZZ As Integer, I As Integer
  Dim EXT As Integer
  Dim ST1 As String

    For I = Len(LongNAME) To 1 Step -1

        If Mid$(LongNAME, I, 1) = "." Then Exit For ':( Expand Structure
    Next I

    EXT = Len(LongNAME) - I

    ZZ = Len(LongNAME)

    For I = ZZ To 1 Step -1
        If Mid$(LongNAME, I, 1) = "\" Or Mid$(LongNAME, I, 1) = "/" Then Exit For ':( Expand Structure
    Next I

    ST1 = Right$(LongNAME, ZZ - I)

    On Error GoTo ooF
    If InStr(LongNAME, ".") > 0 Then Get_fileName = Left$(ST1, Len(ST1) - EXT - 1) ':( Expand Structure
    If InStr(LongNAME, ".") < 1 Then Get_fileName = Left$(ST1, Len(ST1)) ':( Expand Structure

Exit Function

ooF:
    Get_fileName = ST1

End Function

'----------------------------------------
'Name: Get_fileNameEX
'Object: GET
'Event: fileNameEX
'Description:
'----------------------------------------
Function Get_fileNameEX(LongNAME As String) As String

  Dim St As String
  Dim ZZ As Integer, I As Integer
  Dim EXT As Integer
  Dim ST1 As String

    For I = Len(LongNAME) To 1 Step -1

        If Mid$(LongNAME, I, 1) = "." Then Exit For ':( Expand Structure
    Next I

    EXT = Len(LongNAME) - I

    ZZ = Len(LongNAME)

    For I = ZZ To 1 Step -1
        If Mid$(LongNAME, I, 1) = "\" Or Mid$(LongNAME, I, 1) = "/" Then Exit For ':( Expand Structure
    Next I

    ST1 = Right$(LongNAME, ZZ - I)

    On Error GoTo ooF
    If InStr(LongNAME, ".") > 0 Then Get_fileNameEX = Left$(ST1, Len(ST1))   ':( Expand Structure
    If InStr(LongNAME, ".") < 1 Then Get_fileNameEX = Left$(ST1, Len(ST1)) ':( Expand Structure

Exit Function

ooF:
    Get_fileNameEX = ST1

End Function

'----------------------------------------
'Name: Get_LastpathName
'Object: GET
'Event: LastpathName
'----------------------------------------
'----------------------------------------
'Name: Get_LastpathName
'Object: GET
'Event: LastpathName
'Description:
'----------------------------------------
Function Get_LastpathName(LongNAME As String) As String

  Dim St As String
  Dim ST1 As String
  Dim ZZ As Integer, I As Integer, J As Integer, POS1 As Integer, POS2 As Integer

    ZZ = Len(LongNAME)

    For I = ZZ To 1 Step -1
        J = J + 1
        If Mid$(LongNAME, I, 1) = "\" Or Mid$(LongNAME, I, 1) = "/" Then
            If POS1 = 0 Then
                POS1 = I
                GoTo Nexta
            End If
            If POS1 > 0 And POS2 = 0 Then POS2 = I
        End If
Nexta:

        If I = 1 Or POS2 > 0 Then
            If POS2 = 0 Then POS2 = 1
        End If

    Next I

    ST1 = Mid$(LongNAME, POS2 + 1, POS1 - POS2 - 1)
    Get_LastpathName = ST1

End Function

Function Get_LastpathNameEX(LongNAME As String) As String

  Dim St As String
  Dim ST1 As String
  Dim ZZ As Integer, I As Integer, J As Integer, POS1 As Integer, POS2 As Integer

    ZZ = Len(LongNAME)

    For I = ZZ To 1 Step -1
        J = J + 1
        If Mid$(LongNAME, I, 1) = "\" Or Mid$(LongNAME, I, 1) = "/" Then
            If POS1 = 0 Then
                POS1 = I
                GoTo Nexta
            End If
            If POS1 > 0 And POS2 = 0 Then POS2 = I
        End If
Nexta:

        If I = 1 Or POS2 > 0 Then
            If POS2 = 0 Then POS2 = 1
        End If

    Next I

    If POS1 - 1 > 0 Then
        ST1 = Left$(LongNAME, POS1 - 1)
        Get_LastpathNameEX = ST1
    End If

End Function

Function Get_pathName(LongNAME As String) As String

  Dim St As String
  Dim ZZ As Integer, I As Integer
  Dim ST1 As String

    ZZ = Len(LongNAME)

    For I = ZZ To 1 Step -1
        If Mid$(LongNAME, I, 1) = "\" Or Mid$(LongNAME, I, 1) = "/" Then Exit For ':( Expand Structure
    Next I

    ST1 = Left$(LongNAME, I - 1)
    Get_pathName = ST1

End Function

Function FileIs_Valid(ByVal Filename As String) As Boolean

  Dim WFD As WIN32_FIND_DATA ':( Duplicated Name
  Dim hFile As Long ':( Duplicated Name
  Dim fn As String

    If Right$(Filename, 1) <> Chr$(0) Then
        fn = Filename & Chr$(0)
      Else
        fn = Filename
    End If
    hFile = FindFirstFile(Filename, WFD)
    FileIs_Valid = (hFile <> INVALID_HANDLE_VALUE)
    FindClose hFile

End Function

'==============================================================================
'MESH and vertex routines
'
'
'=================================================================================

Sub Fill_Mesh2(VERTZ() As QUEST3D_VERTEX2, MESHES() As D3DXMesh, NumberCreated As Integer)

  Dim I As Long, J As Long, K As Long, L As Long, Num_SUB As Integer

  Dim TotVERT As Long
  Dim B As Boolean
  Dim N As Single
  Dim R As Long
  Dim VV() As QUEST3D_VERTEX2
  Dim MaxVV As Long

    MaxVV = 32766

    On Error GoTo EXIT_ERROR
    TotVERT = UBound(VERTZ) - LBound(VERTZ) + 1

    Num_SUB = TotVERT / MaxVV + 1

    'NumberCreated = Num_SUB
    NumberCreated = 0

    If Num_SUB < 1 Then Num_SUB = 1

    For I = 1 To Num_SUB

        If I = Num_SUB Then

            N = TotVERT Mod MaxVV

            ReDim VV(N - 1)

            J = TotVERT - N

            CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, N

            DXmeshFromVERT2 MESHES(I - 1), VV

            If Not (MESHES(I - 1) Is Nothing) Then NumberCreated = NumberCreated + 1

          Else

            N = MaxVV

            ReDim VV(N - 1)

            J = I * MaxVV - 1

            CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, (I - 1) * MaxVV, N

            DXmeshFromVERT2 MESHES(I - 1), VV

            If Not (MESHES(I - 1) Is Nothing) Then NumberCreated = NumberCreated + 1

        End If

    Next I

EXIT_ERROR:

End Sub

Sub Fill_Mesh(VERTZ() As QUEST3D_VERTEX2, m1 As D3DXMesh, _
        M2 As D3DXMesh, M3 As D3DXMesh, M4 As D3DXMesh, M5 As D3DXMesh, _
        m6 As D3DXMesh, _
        M7 As D3DXMesh, M8 As D3DXMesh, M9 As D3DXMesh, M10 As D3DXMesh, OK() As Boolean)

  Dim I As Long, J As Long, K As Long, L As Long

  Dim TotVERT As Long
  Dim B As Boolean
  Dim N As Single
  Dim R As Long
  Dim VV() As QUEST3D_VERTEX2
  Dim MaxVV As Long

    MaxVV = 32766

    On Error GoTo EXIT_ERROR
    TotVERT = UBound(VERTZ) - LBound(VERTZ) + 1

    If TotVERT = 0 Then Exit Sub

    N = TotVERT / MaxVV

    If N <= 1 Then
        DXmeshFromVERT2 m1, VERTZ
        OK(0) = 1
        '        obj_D3DX.ComputeNormals m1
        GoTo OKK
      ElseIf N <= 2 And N > 1 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1
        '        obj_D3DX.ComputeNormals m1

        'mesh 2
        I = MaxVV
        ReDim VV(TotVERT - MaxVV - 1)

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV

        '
        DXmeshFromVERT2 M2, VV
        OK(1) = 1

      ElseIf N <= 3 And N > 2 Then
        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(TotVERT - MaxVV * 2 - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 2

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        GoTo OKK

      ElseIf N <= 4 And N > 3 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(TotVERT - MaxVV * 3 - 1)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 3

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        GoTo OKK

      ElseIf N > 4 And N <= 5 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(TotVERT - MaxVV * 4 - 1)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 4

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        GoTo OKK

      ElseIf N > 5 And N <= 6 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV - 1)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(MaxVV - 1)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        'mesh 6
        ReDim VV(TotVERT - MaxVV * 5 - 1)
        I = MaxVV * 5

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 5

        DXmeshFromVERT2 m6, VV
        OK(5) = 1

        GoTo OKK

      ElseIf N > 6 And N <= 7 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(MaxVV)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        'mesh 6
        ReDim VV(MaxVV)
        I = MaxVV * 5

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 m6, VV
        OK(5) = 1

        'mesh 7
        ReDim VV(TotVERT - MaxVV * 6 - 1)
        I = MaxVV * 6

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 6

        DXmeshFromVERT2 M7, VV
        OK(6) = 1

        GoTo OKK

      ElseIf N > 7 And N <= 8 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(MaxVV)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        'mesh 6
        ReDim VV(MaxVV)
        I = MaxVV * 5

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 m6, VV
        OK(5) = 1

        'mesh 7
        ReDim VV(MaxVV)
        I = MaxVV * 6

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M7, VV
        OK(6) = 1

        'mesh 8
        ReDim VV(TotVERT - MaxVV * 7 - 1)
        I = MaxVV * 7

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 7

        DXmeshFromVERT2 M8, VV
        OK(7) = 1

        GoTo OKK

      ElseIf N > 8 And N <= 9 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(MaxVV)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        'mesh 6
        ReDim VV(MaxVV)
        I = MaxVV * 5

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 m6, VV
        OK(5) = 1

        'mesh 7
        ReDim VV(MaxVV)
        I = MaxVV * 6

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M7, VV
        OK(6) = 1

        'mesh 8
        ReDim VV(MaxVV)
        I = MaxVV * 7

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M8, VV
        OK(7) = 1

        'mesh 9
        ReDim VV(TotVERT - MaxVV * 8 - 1)
        I = MaxVV * 8

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 8

        DXmeshFromVERT2 M9, VV
        OK(8) = 1

        GoTo OKK

      ElseIf N > 9 And N <= 10 Then

        'mesh 1
        ReDim VV(MaxVV - 1)

        J = 0
        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, J, MaxVV

        '
        DXmeshFromVERT2 m1, VV
        OK(0) = 1

        'mesh 2
        ReDim VV(MaxVV - 1)
        I = MaxVV

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M2, VV
        OK(1) = 1

        'mesh 3
        ReDim VV(MaxVV - 1)
        I = MaxVV * 2

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M3, VV
        OK(2) = 1

        'mesh 4
        ReDim VV(MaxVV)
        I = MaxVV * 3

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M4, VV
        OK(3) = 1

        'mesh 5
        ReDim VV(MaxVV)
        I = MaxVV * 4

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M5, VV
        OK(4) = 1

        'mesh 6
        ReDim VV(MaxVV)
        I = MaxVV * 5

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 m6, VV
        OK(5) = 1

        'mesh 7
        ReDim VV(MaxVV)
        I = MaxVV * 6

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M7, VV
        OK(6) = 1

        'mesh 8
        ReDim VV(MaxVV)
        I = MaxVV * 7

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M8, VV
        OK(7) = 1

        'mesh 9
        ReDim VV(MaxVV)
        I = MaxVV * 8

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, MaxVV

        DXmeshFromVERT2 M9, VV
        OK(8) = 1

        'mesh 10
        ReDim VV(TotVERT - MaxVV * 9 - 1)
        I = MaxVV * 9

        CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2 VV, VERTZ, 0, I, TotVERT - MaxVV * 9

        DXmeshFromVERT2 M10, VV
        OK(9) = 1

        GoTo OKK

    End If

OKK:

Exit Sub

    'vertex size error
EXIT_ERROR:

End Sub

Sub DXmeshFromVERT(LpMESH As D3DXMesh, VERT() As QUEST3D_VERTEX)

  Dim obj_Dinput() As Integer
  Dim I, K, N As Long

    N = UBound(VERT) - LBound(VERT) + 1

    ReDim obj_Dinput(N - 1)

    '        For K = 0 To n - 3 Step 3
    '            obj_Dinput(K) = K
    '            obj_Dinput(K + 1) = K + 1
    '            obj_Dinput(K + 2) = K + 2
    '
    '        Next K

    QUEST3D_CopyIndiceOrder obj_Dinput(), N

    CreateDxMeshFromVERTEX LpMESH, VERT(), obj_Dinput, QUEST3D_FVFVERTEX

End Sub

Sub DXmeshFromVERT2(LpMESH As D3DXMesh, VERT() As QUEST3D_VERTEX2)

  Dim obj_Dinput() As Integer
  Dim I, K, N As Long

    N = UBound(VERT) - LBound(VERT) + 1

    ReDim obj_Dinput(N - 1)

    '        For K = 0 To n - 3 Step 3
    '            obj_Dinput(K) = K
    '            obj_Dinput(K + 1) = K + 1
    '            obj_Dinput(K + 2) = K + 2
    '
    '        Next K

    QUEST3D_CopyIndiceOrder obj_Dinput, N

    CreateDxMeshFromVERTEX2 LpMESH, VERT(), obj_Dinput, QUEST3D_FVFVERTEX2

End Sub

Sub QUEST3D_CopyIndiceOrder(ByRef IndexArray() As Integer, ByVal Num As Long)

  Dim I As Long

    I = 0
    While I < Num
        IndexArray(I) = I
        I = I + 1
    Wend

End Sub

Sub CreateDxMeshFromVERTEX(LpMESH As D3DXMesh, VERT() As QUEST3D_VERTEX, Indices() As Integer, ByVal FVF As Long)

  Dim NV As Long
  Dim NI As Long

    NV = UBound(VERT) + 1
    NI = UBound(Indices) + 1

    On Error Resume Next
        Set LpMESH = obj_D3DX.CreateMeshFVF(NI / 3, NV, 0, FVF, obj_Device)

        'fill Vertex and indice BUFFERS

        D3DVertexBuffer8SetData LpMESH.GetVertexBuffer, 0, NV * Len(VERT(0)), 0, VERT(0)
        D3DIndexBuffer8SetData LpMESH.GetIndexBuffer, 0, NI * 2#, 0, Indices(0)

End Sub

Sub CreateDxMeshFromVERTEX2(LpMESH As D3DXMesh, VERT() As QUEST3D_VERTEX2, Indices() As Integer, ByVal FVF As Long)

  Dim NV As Long
  Dim NI As Long

    NV = UBound(VERT) + 1
    NI = UBound(Indices) + 1

    Set LpMESH = obj_D3DX.CreateMeshFVF(NV / 3, NV, 0, FVF, obj_Device)

    'fill Vertex and indice BUFFERS

    D3DVertexBuffer8SetData LpMESH.GetVertexBuffer, 0, NV * Len(VERT(0)), 0, VERT(0)
    D3DIndexBuffer8SetData LpMESH.GetIndexBuffer, 0, NI * 2#, 0, Indices(0)

End Sub

Sub CreateDxMeshFromVERTEXeX(LpMESH As D3DXMesh, VERT() As QUEST3D_VERTEX, Indices() As Long, ByVal FVF As Long)

  Dim NV As Long
  Dim NI As Long

    NV = UBound(VERT) + 1
    NI = UBound(Indices) + 1

    On Error Resume Next
        Set LpMESH = obj_D3DX.CreateMeshFVF(NI / 3, NV, 0, FVF, obj_Device)

        'fill Vertex and indice BUFFERS

        D3DVertexBuffer8SetData LpMESH.GetVertexBuffer, 0, NV * Len(VERT(0)), 0, VERT(0)
        D3DIndexBuffer8SetData LpMESH.GetIndexBuffer, 0, NI * 4#, 0, Indices(0)

End Sub

Sub CopyQUEST3D_VERTEX2toQUEST3D_VERTEX2(Dest() As QUEST3D_VERTEX2, Source() As QUEST3D_VERTEX2, ByVal DestStart As Long, SourceStart, ByVal Num As Long)

  Dim V As QUEST3D_VERTEX2

    DXCopyMemory Dest(DestStart), Source(SourceStart), Len(V) * Num

End Sub

Sub Create_VertexBufferFromLVERTEX(RetBuffer As Direct3DVertexBuffer8, VERTEX_Array() As QUEST3D_LVERTEX)

  Dim VertexSizeInBytes As Long
  Dim NUM_VERT As Long

    NUM_VERT = UBound(VERTEX_Array) - LBound(VERTEX_Array) + 1

    VertexSizeInBytes = Len(VERTEX_Array(LBound(VERTEX_Array)))

    Set RetBuffer = obj_Device.CreateVertexBuffer(VertexSizeInBytes * NUM_VERT, _
        0, QUEST3D_FVFLVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData RetBuffer, 0, VertexSizeInBytes * NUM_VERT, 0, VERTEX_Array(LBound(VERTEX_Array))

End Sub

Sub Create_VertexBufferFromTLVERTEX(RetBuffer As Direct3DVertexBuffer8, VERTEX_Array() As QUEST3D_TLVERTEX)

  Dim VertexSizeInBytes As Long
  Dim NUM_VERT As Long

    NUM_VERT = UBound(VERTEX_Array) - LBound(VERTEX_Array) + 1

    VertexSizeInBytes = Len(VERTEX_Array(LBound(VERTEX_Array)))

    Set RetBuffer = obj_Device.CreateVertexBuffer(VertexSizeInBytes * NUM_VERT, _
        0, QUEST3D_FVFTLVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData RetBuffer, 0, VertexSizeInBytes * NUM_VERT, 0, VERTEX_Array(LBound(VERTEX_Array))

End Sub

Sub Add_Tri(ByVal T As Long)

    Data.Total_TriangleRENDERED = Data.Total_TriangleRENDERED + T

End Sub

Sub Add_Verti(ByVal T As Long)

    Data.Total_VerticeRENDERED = Data.Total_VerticeRENDERED + T

End Sub

' extract major/minor from version cap
Function D3DSHADER_VERSION_MAJOR(version As Long) As Long

    D3DSHADER_VERSION_MAJOR = (((version) \ 8) And &HFF&)

End Function

Function D3DSHADER_VERSION_MINOR(version As Long) As Long

    D3DSHADER_VERSION_MINOR = (((version)) And &HFF&)

End Function

'vertex shader version token
Function D3DVS_VERSION(Major As Long, Minor As Long) As Long

    D3DVS_VERSION = (&HFFFE0000 Or ((Major) * 2 ^ 8) Or (Minor))

End Function

