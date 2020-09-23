VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2508
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'
'           DIRECT3D TUTORIAL #4 - DIRECT3D INTRODUCTION
'
'                        BY SIMON PRICE
'
'********************************************************************

' This is the simplist sort of Direct3D application you can possibly
' make. There is no attempt to perform any error handling whatsoever,
' for simplicity. Therefore you run this code at your own risk. If
' you find a bug please tell me. Please vote for this code on
' www.planet-source-code.com/vb

' Simon Price, Email : Si@VBgames.co.uk, Website : www.VBgames.co.uk

' the main DirectX7 object - created using the New keyword
Dim DX As New DirectX7
' the main DirectDraw object
Dim DDRAW As DirectDraw7
' this describes surfaces
Dim SurfDesc As DDSURFACEDESC2
' the primary surface - represents the screen
Dim Primary As DirectDrawSurface7
' the backbuffer surface - used for drawing on
Dim Backbuffer As DirectDrawSurface7
' the clip - to contain drawing in just the forms window
Dim Clipper As DirectDrawClipper
' the source and destination rectangles
Dim DestRect As RECT
Dim SrcRect As RECT
' the main Direct3D object
Dim D3D As Direct3D7
' the rendering device
Dim D3Ddevice As Direct3DDevice7
' the viewport rectangle for the rendering device
Dim Viewport(0) As D3DRECT
' the viewport description
Dim VPdesc As D3DVIEWPORT7
' the vertices for the triangle
Dim Vertex(0 To 2) As D3DVERTEX
' the material for out polygon surface
Dim Material As D3DMATERIAL7
' matrices we will need to describe transformations
Dim matWorld As D3DMATRIX
Dim matView  As D3DMATRIX
Dim matProj As D3DMATRIX
Dim matSpin As D3DMATRIX

' this tells the program when to end
Dim EndNow As Boolean
' this is used to rotate the triangle
Dim Counter As Long




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'          DirectDrawInit - Initializes DirectDraw objects
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function DirectDrawInit() As Long
' create the directdraw object
Set DDRAW = DX.DirectDrawCreate("")
' set the cooperative level, we only need normal
DDRAW.SetCooperativeLevel hWnd, DDSCL_NORMAL

' set the properties of the primary surface
SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
' create the primary surface
Set Primary = DDRAW.CreateSurface(SurfDesc)

' set up the backbuffer surface (which will be where we render the 3D view)
SurfDesc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
' use the size of the form to determine the size of the render target
' and viewport rectangle
DX.GetWindowRect hWnd, DestRect
' set the dimensions of the surface description
SurfDesc.lWidth = DestRect.Right - DestRect.Left
SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
' create the backbuffer surface
Set Backbuffer = DDRAW.CreateSurface(SurfDesc)

' cache the size of the render target for later use
With SrcRect
        .Left = 0: .Top = 0
        .Bottom = SurfDesc.lHeight
        .Right = SurfDesc.lWidth
End With

' create a DirectDrawClipper and attach it to the primary surface.
Set Clipper = DDRAW.CreateClipper(0)
Clipper.SetHWnd hWnd
Primary.SetClipper Clipper

' report any errors
DirectDrawInit = Err.Number
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'          Direct3DInit - Initializes Direct3D objects
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Direct3DInit() As Long
' create the direct3d object
Set D3D = DDRAW.GetDirect3D

' create the rendering device - we are using software emulation only
Set D3Ddevice = D3D.CreateDevice("IID_IDirect3DRGBDevice", Backbuffer)

' set the viewport rectangle.
VPdesc.lWidth = DestRect.Right - DestRect.Left
VPdesc.lHeight = DestRect.Bottom - DestRect.Top
VPdesc.minz = 0
VPdesc.maxz = 1
D3Ddevice.SetViewport VPdesc

' cache the viewport rectangle for later use
With Viewport(0)
    .X1 = 0: .Y1 = 0
    .X2 = VPdesc.lWidth
    .Y2 = VPdesc.lHeight
End With
    
' enable ambient lighting
D3Ddevice.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(1, 1, 1, 1)
' disable culling
D3Ddevice.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_NONE

' set the material to a red color
Material.Ambient.r = 1
Material.Ambient.g = 0
Material.Ambient.b = 0
D3Ddevice.SetMaterial Material

' the world matrix - all polygons in world space are transformed by this matrix
DX.IdentityMatrix matWorld
D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matWorld
' the view matrix - basically the camera position is at -3
' (although it's really just making the whole world at +3)
DX.IdentityMatrix matView
DX.ViewMatrix matView, MakeVector(0, 0, -3), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
' the projection matrix - decides how the 3D scene is projected onto the 2D surface
DX.IdentityMatrix matProj
DX.ProjectionMatrix matProj, 1, 1000, 3.14 / 2
D3Ddevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj

' report errors
Direct3DInit = Err.Number
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'          CreateTriangle - Makes a polygon from our vertices
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CreateTriangle()
' fill in the vertex positions - we don't need to worry about the normals
' or texture coordinates for this tutorial
DX.CreateD3DVertex -1, -1, 0, 0, 0, 0, 0, 0, Vertex(0)
DX.CreateD3DVertex 0, 1, 0, 0, 0, 0, 0, 0, Vertex(1)
DX.CreateD3DVertex 1, -1, 0, 0, 0, 0, 0, 0, Vertex(2)
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'          MakeVector - Makes a vector from x,y & z coords
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function MakeVector(x As Single, y As Single, z As Single) As D3DVECTOR
' copy x, y and z into the return value
With MakeVector
    .x = x
    .y = y
    .z = z
End With
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'             MainLoop - Continually renders the scene
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub MainLoop()
Do While EndNow = False
    ' increase the counter
    Counter = Counter + 1
    
    ' clear the viewport with a green color
    D3Ddevice.Clear 1, Viewport(), D3DCLEAR_TARGET, vbGreen, 0, 0
    ' begin the scene, render the triangle, then end the scene
    D3Ddevice.BeginScene
    D3Ddevice.DrawPrimitive D3DPT_TRIANGLELIST, D3DFVF_VERTEX, Vertex(0), 3, D3DDP_DEFAULT
    D3Ddevice.EndScene
    
    ' rotate the matrix
    DX.RotateYMatrix matSpin, Counter / 360
    ' set the new world transform matrix
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matSpin
    
    ' copy the backbuffer to the screen
    DX.GetWindowRect hWnd, DestRect
    Primary.Blt DestRect, Backbuffer, SrcRect, DDBLT_WAIT
    
    ' look for window messages - we need to know when the escape key is pressed
    DoEvents
Loop
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        The Form_Load event - Calls the other procedures
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
' show the form
Show
' call the DirectDrawInit function and exit if it fails
If DirectDrawInit() <> DD_OK Then Unload Me
' call the Direct3DInit function and exit if it fails
If Direct3DInit() <> DD_OK Then Unload Me
' create the triangle
CreateTriangle
' call the main rendering loop
MainLoop
' end the program
Unload Me
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'           The Form_Unload event - Stops the main loop
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Unload(Cancel As Integer)
EndNow = True
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        The Form_KeyDown event - Also stops the main loop
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' end program if escape is pressed
If KeyCode = vbKeyEscape Then EndNow = True
End Sub

