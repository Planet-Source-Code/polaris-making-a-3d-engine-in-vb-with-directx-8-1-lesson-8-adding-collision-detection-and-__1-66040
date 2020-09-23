VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Engine_Lesson 7 Sky System and LensFlare"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'WELCOME Engine_Lesson 8 adding frustum Culling and Collision detection
'_____________________________________________________________________________________
'-------------------------------------------------------------------------------------
'
'===================================================================================
'Welcome to this Step by step Quest to a 3D Engine programming
'this tutorial will show you how to design a simple 3D
'engine, Next tutorials will show how to add other engine objet
'like Camera,Mesh and Object Polygon
'
'This tutorial 8: adding frustum Culling and Collision detection
'
'It shows you how to
'  - Add an advanced collision system with collision response
'  - use view frustum datas in order to check if an object is visible or not
'  - We use advanced settings to improve rendering speed and quality
'
'
'

'How to read this code
'
'   - Form1: is the engine code in action
'   - frmEnum have code for Device anumeration
'   - Module_definitions will hold all engine objets definitions and types
'   - Module_Util will hold all Vector,matrix,Color math stuff
'   - Module_Frustum for visibility testing
'   - Module_Input for input device
'   - Module_TextureManagement
'   - Module_Collision for collision
'
'   - cQuest3D_Core is our first object, it defines Main entry of the engine
'   - cQuest3D_Mesh 3D Mesh Class, to generate procedurally 3D Scene
'   - cQuest3D_Input Define Input class (Keyboard,Mouse,Joystick)
'   - cQuest_Camera handle Camera
'   - cQuest3D_SkyBox 'for sky and sky animation
'   - cQuest3D_LensFlare for lensflare effect

'
'Good coding
'
'Vote if you want the sequel!!
'
'==================================================================================

Option Explicit

'we use the engine here
'we declare an objet
Dim QUEST As cQuest3D_Core

'mesh
Dim MyMESH1 As cQuest3D_Mesh
Dim MyMESH2 As cQuest3D_Mesh

'camera
Dim KAMERA As cQuest3D_Camera
'we use these vars to do
'frame based camera animation
Const ROTATION_SPEED As Single = 1
Const PLAYER_SPEED As Single = 500
'see GetInput()

'Input
Dim KEY As cQuest3D_Input
Dim ShowInfo As Boolean

'for lensflare

Dim LENSFLARE As cQuest3D_LensFlare

'for skybox
Dim SKY As cQuest3D_SkyBox


Private Sub Form_Load()

    InitEngine

End Sub

Sub InitEngine()

  'we allocate memory here

    Set QUEST = New cQuest3D_Core

    '        'we initialize the engine
'            If QUEST.Init_Dialogue(Me.hwnd) = False Then
'             MsgBox "Sorry there was an error"
'             End
'            End If

    If QUEST.Init(Me.hwnd) = False Then
        MsgBox "Sorry there was an error"
        End
    End If

    Me.Refresh
    Me.Show

    'init default state for the engine
    QUEST.Set_EngineAmbientColor Make_ColorRGB(255, 255, 255) 'white

    QUEST.Set_BackbufferClearColor Make_ColorRGB(100, 100, 100)

    'we want to clear zbuffer and backbuffer
    QUEST.Set_EngineClearRenderTarget True
    'fill mode
    QUEST.Set_EngineFillMode QUEST3D_FILL_SOLID

    'shade mode
    QUEST.Set_EngineShadeMode QUEST3D_SHADE_GOURAUD
    'choose from:
    '    QUEST3D_FILTER_POINT = 0
    '    QUEST3D_FILTER_BILINEAR = 1
    '    QUEST3D_FILTER_TRILINEAR = 2
    '    QUEST3D_FILTER_ANISOTROPIC = 3
    ' Best is QUEST3D_FILTER_ANISOTROPIC
    QUEST.Set_EngineTextureFilter QUEST3D_FILTER_TRILINEAR

    'here we define mimap filter
    'what is a mipmap?
    'A mipmap is reduce sized textures that allow
    'texture mapping on small polygon size,it allow to preserve drawing details
    'we choos the best filter to apply to these mipmaps
    'because the engine we are building tells to Direct3D to generate mipmaps
    'for each texture we use DEFAULT filtering method
    QUEST.Set_EngineTextureMipMapFilter QUEST3D_TEXTURE_DEFAULT

    'prepare input

    Set KEY = New cQuest3D_Input
    'we force input devices creation
    KEY.ReCreateInputDevices
    'we force device polling
    KEY.ReCreateInputDevices

    ShowInfo = True

    'init Camera
    Set KAMERA = New cQuest3D_Camera
    'we set the 6DOF style
    KAMERA.Set_CameraStyle FREE_6DOF
    'we use Left Hand perpective projection
    KAMERA.Set_CameraProjectionType PT_PERSPECTIVE_LH
    'we use A field of view where near=10 and far=10000, Angle=45 degree
    KAMERA.Set_ViewFrustum 10, 20000, 45 * QUEST3D_RAD
    'initial camera position=0,100,0 looking at 0,100,100
    KAMERA.Set_camera Vector(-100, 200, 0), Vector(0, 200, 100)
    'we update the camera
    KAMERA.Update

    'we prapare geometry
    PrepareGeometry
    'we call game loop
    GameLoop

End Sub

'==================================================================================
'
'In this sub we procedurally generate
'geometry here, the floor and cylinders
'
'==================================================================================

Sub PrepareGeometry()

  Dim Texture_ID As Long
  Dim I As Long

    'here we allocate memory for mesh
    Set MyMESH1 = New cQuest3D_Mesh

    '1st we add textures
    Texture_ID = MyMESH1.Add_Texture(App.Path + "\Data\castle_m04.jpg") '(ID 1)

    '2nd we ad vertices,polygons
    'MyMESH1.Add_WallFloor Vector(-9000, -1, -9000), Vector(9000, -1, 9000), 10, 10, 0 '0 means we used fisrt textures added
'
'    'we add randomly center cylinders
    For I = 1 To 10 '80 can be changed to 850 max ,This Engine can handle 120 000 Polygons Max per Sub mesh

        MyMESH1.Add_Cilynder Vector((Rnd - Rnd) * 10000, -1, (Rnd - Rnd) * 10000), 100 + (Rnd - Rnd) * 50 + 50, 500 + (Rnd - Rnd) * 200 + 100, 10 + (Rnd) * 20, Texture_ID '0 means 1st texture added

    Next I

    '3rd we Build the mesh
    'all information for fast rendering will be
    'computed
    MyMESH1.BuildMesh
    

    

    'now we define a more complex scene

    'here we allocate memory for mesh
    Set MyMESH2 = New cQuest3D_Mesh

    MyMESH2.Add_Texture (App.Path + "\Data\Relief_8.jpg")          '0
    MyMESH2.Add_Texture (App.Path + "\Data\facade2.jpg")  '1

    MyMESH2.Add_Texture App.Path + "\Data\street.JPG"     '2
    MyMESH2.Add_Texture App.Path + "\Data\street2.JPG"    '3
    MyMESH2.Add_Texture (App.Path + "\Data\kerb2.jpg")    '4
    MyMESH2.Add_Texture (App.Path + "\Data\kerb.jpg")     '5

    MyMESH2.Add_Texture (App.Path + "\Data\square.JPG")   '6

    MyMESH2.Add_Texture App.Path + "\Data\cement.JPG"     '7

    MyMESH2.Add_Texture App.Path + "\Data\road_t03.jpg"   '8

    MyMESH2.Add_Texture App.Path + "\Data\Asfalto1.bmp"   '9
    MyMESH2.Add_Texture App.Path + "\Data\pierres.JPG"    '10

    MyMESH2.Add_Texture App.Path + "\Data\win2.JPG"       '11
    MyMESH2.Add_Texture App.Path + "\Data\windows.JPG"       '12

    MyMESH2.Add_Texture App.Path + "\Data\StoreSd.BMP"     '13

    'city 1 floor
    MyMESH2.Add_WallFloor Vector(-5000, -1, -5000), Vector(5000, -1, 5000), 10, 10, 0

    'city 2 floor
    MyMESH2.Add_WallFloor Vector(-5000, -1, 5000), Vector(5000, -1, 10000), 10, 10, 7

    '=======NEW CODE======='
    'add building 1
    MyMESH2.Add_Box Vector(-500, 0, 0), Vector(0, 500, 500), 1, 1, 1, 1, 1, 1

    'add building 2
    MyMESH2.Add_Box Vector(-500, 0, -2000), Vector(0, 500, -1500), 1, 1, 1, 1, 1, 1

    'add building 3
    MyMESH2.Add_Box Vector(601, 0, 0), Vector(1051, 800, 500), 11, 11, 11, 11, 11, 11

    'add building 4
    MyMESH2.Add_Box Vector(601, 0, -2000), Vector(1051, 900, -1500), 12, 12, 12, 12, 12, 12

    'add building 4
    MyMESH2.Add_Box Vector(601, 0, 2000), Vector(1051, 900, 2500), 13, 13, 13, 13, 13, 13

    'draw the road  segment of 1000 from -5000 to 5000
    MyMESH2.Add_WallFloor Vector(50, 0, -5000), Vector(550, 0, 5000), 1, 2, 2 'the 3rd texture passed to the mesh class

    'draw the north and south roads
    'north part
    MyMESH2.Add_WallFloor Vector(50, 0, 5000), Vector(550, 0, 5500), 1, 1, 6
    MyMESH2.Add_WallFloor Vector(550, 0, 5000), Vector(5000, 0, 5500), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(-5000, 0, 5000), Vector(50, 0, 5500), 5, 1, 8
    'south part
    MyMESH2.Add_WallFloor Vector(550, 0, -5500), Vector(5000, 0, -5000), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(-5000, 0, -5500), Vector(50, 0, -5000), 5, 1, 8
    MyMESH2.Add_WallFloor Vector(50, 0, -5500), Vector(550, 0, -5000), 1, 1, 6

    'make the pavements
    MyMESH2.Add_WallFloor Vector(0, 10, -5000), Vector(50, 10, 5000), 1, 10, 5
    MyMESH2.Add_WallFloor Vector(550, 10, -5000), Vector(600, 10, 5000), 1, 10, 4
    MyMESH2.Add_WallLeft Vector(50, 0, -5000), Vector(50, 10, 5000), 10, 0.25, 9
    MyMESH2.Add_WallRight Vector(550, 0, -5000), Vector(550, 10, 5000), 18, 0.25, 9

    'the south west pavements
    MyMESH2.Add_WallBack Vector(-5000, -1, -5000), Vector(50, 10, -5000), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, -4950), Vector(0, 10, -4950), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, -5000), Vector(0, 10, -4950), 18, 1, 9

    MyMESH2.Add_WallBack Vector(-5000, -1, -5500), Vector(50, 10, -5500), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, -5450), Vector(0, 10, -5450), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, -5500), Vector(0, 10, -5450), 18, 1, 9

    'the south east pavements
    MyMESH2.Add_WallBack Vector(550, -1, -5000), Vector(5000, 10, -5000), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, -4950), Vector(5000, 10, -4950), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(600, 10, -5000), Vector(5000, 10, -4950), 18, 1, 9

    MyMESH2.Add_WallBack Vector(550, -1, -5500), Vector(5000, 10, -5500), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, -5450), Vector(5000, 10, -5450), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(500, 10, -5500), Vector(5000, 10, -5450), 18, 1, 9
    'big wall
    MyMESH2.Add_WallFront Vector(-5000, -1, -5501), Vector(5000, 500, -5501), 18, 1, 10

    'the north west pavements
    MyMESH2.Add_WallBack Vector(-5000, -1, 4950), Vector(0, 10, 4950), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(-5000, -1, 5000), Vector(50, 10, 5000), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, 4950), Vector(0, 10, 5000), 18, 1, 9

    MyMESH2.Add_WallFront Vector(-5000, -1, 5550), Vector(50, 10, 5550), 18, 0.5, 9
    MyMESH2.Add_WallBack Vector(-5000, -1, 5500), Vector(0, 10, 5500), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(-5000, 10, 5500), Vector(0, 10, 5550), 18, 1, 9

    'the north east pavements
    MyMESH2.Add_WallBack Vector(550, -1, 4950), Vector(5000, 10, 4950), 18, 0.5, 9
    MyMESH2.Add_WallFront Vector(550, -1, 5000), Vector(5000, 10, 5000), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(600, 10, 4950), Vector(5000, 10, 5000), 18, 1, 9

    MyMESH2.Add_WallFront Vector(550, -1, 5550), Vector(5000, 10, 5550), 18, 0.5, 9
    MyMESH2.Add_WallBack Vector(550, -1, 5500), Vector(5000, 10, 5500), 18, 0.5, 9
    MyMESH2.Add_WallFloor Vector(500, 10, 5500), Vector(5000, 10, 5550), 18, 1, 9

    'add big wall on the south to limit city extension
    MyMESH2.Add_WallBack Vector(-5000, -1, 5551), Vector(5000, 1500, 5551), 18, 1, 10


   
    'then we build our mesh
    MyMESH2.BuildMesh
    
    
FX:
    
    'we prepare sky
    'classic skybox code

    Set SKY = New cQuest3D_SkyBox
    SKY.Init_SkyBox


    SKY.Add_SkyBOX QUEST3D_SKY_LEFT, App.Path + "\Data\Sky\" + "cloudy_noon_LF.jpg"
    SKY.Add_SkyBOX QUEST3D_SKY_RIGHT, App.Path + "\Data\Sky\" + "cloudy_noon_RT.jpg"
    'because front and back are inversed in our sky system we switch them
    'back becomes front and vice verca
    SKY.Add_SkyBOX QUEST3D_SKY_BACK, App.Path + "\Data\Sky\" + "cloudy_noon_FR.jpg"
    SKY.Add_SkyBOX QUEST3D_SKY_FRONT, App.Path + "\Data\Sky\" + "cloudy_noon_BK.jpg"
    
    SKY.Add_SkyBOX QUEST3D_SKY_DOWN, App.Path + "\Data\Sky\" + "cloudy_noon_DN.jpg"
    SKY.Add_SkyBOX QUEST3D_SKY_TOP, App.Path + "\Data\Sky\" + "cloudy_noon_UP.jpg"
    
    SKY.Set_TextureRotateCoordonates QUEST3D_SKY_TOP, QUEST3D_RAD * (-90)
    
    SKY.Set_Scale 11000, 11000, 11000
    
    
    
    'we create the LENSFLARE flare
     '============NEW CODE========
    ' setting up the LENSFLAREflare
    '===========================

    'allocate memory for our LENSFLAREflare class
    Set LENSFLARE = New cQuest3D_LensFlare
    'set our SunGlow textureFile,Position, and size
    LENSFLARE.Add_SunTEX App.Path + "\Data" + "\flare\sunHalo_Color.tga", Vector(-2000, 5000, -8900), 2000

    'configurate our cameraLENSFLAREburningout color default=white &HFFFFFFFF
    LENSFLARE.Set_SunburningEffectColor 2, 255, 255
    LENSFLARE.Set_SunburningScreen 0 'deactivate burning screen out effect

    'specify that we wil use a static position for our LENSFLARElare
    'this allow to simultate Ray of LENSFLARE from any Static lightSource
    'In this demo we will use a static source

    'uncomment the above lign to set the LENSFLARE flare at a static
    'position
    'LENSFLARE.Set_LENSFLAREPositionStatus STATIC_SOURCE

    'activate randomizer
    Randomize Timer

    'we add 6 LENSFLARE Spark.....TextureFile,index,size,position,Color,SpecularColor
    'TextureFile =any TGA,BMP,JPG file
    'index       =the Index for the LENSFLARESpark
    'size        =the size of our spark
    'position    =the position the sun pos is 100, middle distance=50
    'Color       =main Color for the spark
    'highlight   =color for the outter of the circle

    'use random color=Make_ColorRGBAEx(Rnd, Rnd, Rnd, .4)

    'we add 14 flares textures
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\far1.tga", 0, 25, 200, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.9), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)

    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare1.tga", 1, 20, 180, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.9), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare2.tga", 2, 17, 199, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.9), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare3.tga", 3, 13, 190, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.9), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare4.tga", 4, 10, 100, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.7), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare5.tga", 5, 9, 80, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.7), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare6.tga", 6, 5, 68, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.7), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)

    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare7.tga", 7, 4, 50, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.85), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare8.tga", 8, 3, 40, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.5), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare9.tga", 9, 5, 13, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.5), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.6)

    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare10.tga", 10, 2, -10, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.5), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare11.tga", 11, 1, -17, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.4), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8)
    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare12.tga", 12, 0.8, -20, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.8), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.9)

    LENSFLARE.Add_Lens App.Path + "\Data" + "\flare\Flare13.tga", 13, 1.4, -30, Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.4), Make_ColorRGBAEx(Rnd, Rnd, Rnd, 0.4)

    'this next line allows to specify the blending mode
    LENSFLARE.Set_LenflareBlend FLARES_BLENDMODE, D3DBLEND_SRCALPHA, D3DBLEND_ONE
    'make sun animation
    LENSFLARE.Set_SunAnimation True
    
    'set rotation speed default speed pi/10000
    'LENSFLARE.Set_SunAnimationSpeed 0.5
    'uncomment if you want player view to be over brighten
'    LENSFLARE.Set_SunburningEffectColor 255, 255, 255
'    LENSFLARE.Set_SunburningScreen True

End Sub

Sub GameLoop()
Dim V As D3DVECTOR

    Do

        'Now we use camera
        GetInput
        'here we check if camera collides with meshes if so we just
        'slide over the polygon plane
        If MyMESH2.Check_SphereCollisionSliding(KAMERA.Get_Position, V, 80) Then
          KAMERA.Set_Position V
          KAMERA.Update
        
        End If
        
          If MyMESH1.Check_SphereCollisionSliding(KAMERA.Get_Position, V, 80) Then
          KAMERA.Set_Position V
          KAMERA.Update
        
        End If
       
        'change the clear color randomely
        If QUEST.Get_KeyPressedApi(vbKeySpace) Then QUEST.Set_BackbufferClearColor (D3DColorXRGB(Rnd * 255, Rnd * 255, Rnd * 255))
        'we begin 3D
        QUEST.Begin3D
        
        'render the sky firts
       
        SKY.RenderSky

        MyMESH1.Render
        MyMESH2.Render
       
       LENSFLARE.Render

        'draw FPS
        QUEST.Draw_Text "FPS=" + CStr(QUEST.Get_FramesPerSeconde), 1, 10, &HFFFFFFFF

        If ShowInfo Then
            QUEST.Draw_Text "Polygon =" + CStr(QUEST.Get_NumberOfPolygonDrawn), 1, 25, &HFFFFFFFF
            QUEST.Draw_Text "Press ESC key to quit", 1, 40, &HFFFFFF00

            QUEST.Draw_Text "Press F1 to switch in FREE camera mode,F2 to FPS camera mode", 1, 55, &HFF00FF00

            QUEST.Draw_Text "Use Mouse to Rotate camera,Arrow Keyboard to move camera", 1, 85, &HFFFFFFFF

            If KAMERA.Get_CameraStyle = FPS_STYLE Then
                QUEST.Draw_Text "Current camera mode=", 1, 115, &HFFFFFFFF
                QUEST.Draw_Text "FIRST PERSON SHOOTER STYLE", 170, 115, &HFFFF0001

            Else
                QUEST.Draw_Text "Current camera mode=", 1, 115, &HFFFFFFFF
                QUEST.Draw_Text "6 DEGREE OF FREEDOM STYLE", 170, 115, &HFFFF0001

            End If

            QUEST.Draw_Text "Press 'H' to hide info,'S' to show info", 1, 135, &HFFFFFFFF

        End If
        
        

        'we close 3D Drawing
        QUEST.End3D
        DoEvents

        If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_ESCAPE) Then Call CloseGame
    Loop

End Sub

'in this sub we used
'advanced camera movements
Sub GetInput()

  'for informations printed to the screen

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_F1) Then KAMERA.Set_CameraStyle FREE_6DOF
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_F2) Then KAMERA.Set_CameraStyle FPS_STYLE

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_S) Then ShowInfo = True
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_H) Then ShowInfo = False

    'we use Get_TimePassed to move camera in X unit comparativelly to the
    'time passed, so are in Time based animation

    'strafe
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_LEFT) Then _
       KAMERA.Strafe_Left QUEST.Get_TimePassed * PLAYER_SPEED

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_RIGHT) Then _
       KAMERA.Strafe_Right QUEST.Get_TimePassed * PLAYER_SPEED

    'move forward and backward
    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_UP) Then _
       KAMERA.Move_Forward QUEST.Get_TimePassed * PLAYER_SPEED

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_RCONTROL) Then _
       KAMERA.Move_Forward QUEST.Get_TimePassed * PLAYER_SPEED * 8

    If KEY.Get_KeyBoardKeyPressed(QUEST3D_KEY_DOWN) Then _
       KAMERA.Move_Backward QUEST.Get_TimePassed * PLAYER_SPEED

    'here we use automated Camera rotation via Mouse Input
    '1st param=mouse speed default=0.001
    '2nd param=invert mouse default=false
    '3rd param=center mouse default=false
    'Rotate camera
    KAMERA.RotateByMouse 0.001, False, False

    KAMERA.Update

End Sub

'we quit game here
Sub CloseGame()

    MyMESH1.Free
    MyMESH2.Free
    QUEST.FreeEngine
    LENSFLARE.Free
    
    Set MyMESH1 = Nothing
    Set MyMESH2 = Nothing
    Set QUEST = Nothing
    Set KAMERA = Nothing
    Set KEY = Nothing
    Set LENSFLARE = Nothing
   
    

    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseGame

End Sub
