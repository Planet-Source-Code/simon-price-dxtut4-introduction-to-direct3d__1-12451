<div align="center">

## DXTUT4 \- Introduction to Direct3D\!


</div>

### Description

<p>This tutorial <b> WILL</b> teach you how to get started with using <font color="#FF0000"> Direct3D Immediate Mode from Visual

Basic</font>. It includes background knowledge, definitions, explanations, a sample program to download, and exercises for you to practice on. I have spent hours, wrong, days planning, writing, testing and re-reading this so that it's almost a work of art. Seriously though, you will learn alot. I recommend a very basic

knowledge of DirectDraw, but this is not required, and a fairly good general programming ability, since only DirectX terms will be explained in detail.

If you think that this has helped you, interested you, or changed your whole

life (OK maybe not), <font color="#FF0000">please vote and/or give feedback

because I value your opinions</font>. Especially if you think this was a bad

tutorial, please tell me why and I will try to fix it.</p>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-10-29 17:10:10
**By**             |[Simon Price](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/simon-price.md)
**Level**          |Intermediate
**User Rating**    |4.7 (160 globes from 34 users)
**Compatibility**  |VB 6\.0
**Category**       |[DirectX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/directx__1-44.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD111851112000\.zip](https://github.com/Planet-Source-Code/simon-price-dxtut4-introduction-to-direct3d__1-12451/archive/master.zip)





### Source Code

<h1 align="center"><u><b><a name="lesson">An Introduction To Direct3D</a></b></u></h1>
<h4 align="center"><u>By <a href="mailto:Si@VBgames.co.uk">Simon Price</a></u></h4>
<h3 align="left"><u>Tutorial Breakdown</u></h3>
<p align="left">This tutorial will consist of the following steps :</p>
<ul>
 <li>
  <p align="left"><a href="#explain">Explanation</a> of what Direct3D does and how you can use it
  from Visual Basic</li>
 <li>
  <p align="left"><a href="#demonstate">Definitions</a> of all the objects, types and enumerations you
  will need to know to get started</li>
 <li>
  <p align="left"><a href="#imitate">Example</a> source code with heavy commenting</li>
 <li>
  <p align="left"><a href="#practice">Summary</a> of what you have learnt</li>
 <li>
  <p align="left"><a href="#bingo">Exercises</a> to make you remember it all</li>
</ul>
<h3 align="left"><u><a name="explain">Direct3D Overview</a></u></h3>
<p align="left">Direct3D is a part of DirectX. This tutorial is specific to
Direct3D 7, so you will need DirectX 7.0 or higher if you are planning to use
what you learn here. DirectX has a component called DirectDraw, which is used to
perform graphics functions at a lower level that Windows GDI. If you have never
used DirectDraw before, I suggest you look at my tutorial &quot;An Introduction
To DirectDraw&quot;, available on this site, or <a href="http://www.VBgames.co.uk">my
website</a>. Direct3D (D3D) has two main parts - Immediate Mode and Retained
Mode. This tutorial deals with Immediate Mode only. Immediate Mode (IM) is built
on top of DirectDraw. That means it uses DirectDraw to place graphics on the
screen, or in memory. D3D Retained Mode (RM) is built on top of D3D IM.
Therefore, D3D RM is not as efficient as D3D IM. This is why I have chosen to
learn D3D IM. However, I do not claim that one is better than the other, just
that IM is faster and RM is easier to learn and create applications very quickly
with. If you learn IM, heavy vector mathematics and slow development is involved
but you will be rewarded with more power and control. The choice is yours. If
you still want to learn IM, then read on.</p>
<p align="left">Direct3D has a job - to give programmers a common interface for
all 3D devices. In English - no matter what computer your application runs on,
whether it has a Voodoo Mega Wicked 10000 3D accelerator or a Omega Budget 256
Color Economy VGA card, you still use the same objects to program with. It means
that you don't have to learn about how every graphics card works for your
application to work on every computer. Direct3D also provides software
emulation. This means that if half your users have hardware acceleration, and
half don't, you can use hardware if available and then fall back to using
Direct3D software emulation if the hardware is not available. Of course software
emulation is alot slower.</p>
<p align="left">It's time to start Visual Basic! Create a new project and called
it something imaginative like &quot;D3Dintro.vbp&quot;. Next, click Project -&gt; References
and a dialog box will show a list of references your project uses. If you have
installed the DirectX7 For Visual Basic type library, scroll down to it and
check the check box next to it. Click OK to add the reference. Now Visual Basic
knows every single class, type and enumeration you need to use DirectX7. If you
do not have the DirectX 7 For Visual Basic Type Library, you can download it
from <a href="http://www.microsoft.com">www.microsoft.com</a> .</p>
<h3 align="left"><u><a name="demonstate">Get on with the programming!</a></u></h3>
<p align="left">Here are the declarations you will need for the tutorial
programs, with a short explanation as to what they are all about. First the
objects followed by the types.</p>
<ul>
 <li>
  <p align="left"><b>DirectX7</b> - this is the great big daddy of them all!
  It is from the DirectX7 object that you will create all the other objects,
  including DirectDraw and Direct3D. Note the use of the New keyword, meaning
  that your application puts aside the memory to create a new instance of this
  object.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> DX <font color="#0000FF">As</font> <font color="#0000FF">New</font> DirectX7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>DirectDraw7</b> - this is the base of all the graphics
  functionality that DirectX provides, including Direct3D7. Note the omission
  of the New keyword, since you do not create this object, but DirectX does.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> DDRAW <font color="#0000FF">As</font> DirectDraw7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>DirectDrawSurface7</b> - this is an object created by
  DirectDraw to represent a piece of memory. You will need a primary and
  backbuffer surface. The primary surface represents the actual graphics on
  the screen, the backbuffer is a surface to draw our whole image onto before
  we copy it to the primary surface.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> Primary <font color="#0000FF">As</font> DirectDrawSurface7</pre>
</div>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> Backbuffer <font color="#0000FF">As</font> DirectDrawSurface7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>DirectDrawClipper</b> - this is used to clip areas,
  meaning that if you try draw outside the clipping boundaries, nothing will
  be drawn. This is useful in Windows so that you don't make a mess all over
  bits of screen that don't belong to your application.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> Clipper <font color="#0000FF">As</font> DirectDrawClipper</pre>
</div>
<ul>
 <li>
  <p align="left"><b>Direct3D7</b> - this is based upon DirectDraw. It
  provides all the 3D functionality you will need.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> D3D <font color="#0000FF">As</font> Direct3D7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>Direct3DDevice7 </b>- this is the rendering device. You
  use it to control the states and parameters of Direct3D, and to send drawing
  commands to draw (usually) triangles.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> D3Ddevice <font color="#0000FF">As</font> Direct3DDevice7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>RECT </b>- this describes a rectangle, and DirectDraw
  uses it to copy rectangular pieces of pictures around. Here we need two,
  they are just cached for regular use in the program.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> SrcRect <font color="#0000FF">As</font> RECT</pre>
</div>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> DestRect <font color="#0000FF">As</font> RECT</pre>
</div>
<ul>
 <li>
  <p align="left"><b>D3DRECT</b> - this is similar to the RECT type used with
  DirectDraw. We will use it in clearing operations. You will always need to
  declare it as an array, even if you only need one of them.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> Viewport(0) <font color="#0000FF">As</font> D3DRECT</pre>
</div>
<div align="left">
 <ul>
  <li>
   <p align="left"><b>DDSURFACEDESC2 - </b>this describes a DirectDrawSurface,
   so we can ask DirectDraw to create a surface with the properties we need.</li>
 </ul>
 <div align="left">
  <pre align="left"><font color="#0000FF">Dim</font> SurfDesc as DDSURFACEDESC2</pre>
 </div>
</div>
<ul>
 <li>
  <p align="left"><b>D3DVIEWPORT7</b> - this describes the way in which
  Direct3D transforms a 3D scene to represent it on a 2D surface.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> VPdesc <font color="#0000FF">As</font> D3DVIEWPORT7</pre>
</div>
<ul>
 <li>
  <p align="left"><b>D3DVERTEX</b> - this type holds all the information we need to
     create a vertex. We are going to create a triangle so we need an array
     of 3.
 </li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> Vertex(0 to 2) as D3DVERTEX</pre>
</div>
<ul>
 <li>
  <p align="left"><b>D3DMATRIX</b> - this holds 16 values which are used for
  any and every translation in 3D. With a matrix, you can translate, rotate
  and scale. We will need four in this tutorial, the world, view, projection
  and spin matrices.</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Dim</font> matWorld <font color="#0000FF">As</font> D3DMATRIX
<font color="#0000FF">Dim</font> matView <font color="#0000FF">As</font> D3DMATRIX
<font color="#0000FF">Dim</font> matProj <font color="#0000FF">As</font> D3DMATRIX
<font color="#0000FF">Dim</font> matSpin <font color="#0000FF">As</font> D3DMATRIX</pre>
</div>
 <p align="left">You will also need to declare two other variables:</p>
<div align="left">
 <pre align="left"><font color="#008000">' this tells the program when to end
</font><font color="#0000FF">Dim</font> EndNow <font color="#0000FF">As</font> <font color="#0000FF">Boolean</font>
<font color="#008000">' this is used to rotate the triangle
</font><font color="#0000FF">Dim</font> Counter <font color="#0000FF">As</font> <font color="#0000FF">Long</font>
</pre>
</div>
<h3 align="left"><u><a name="imitate">Initiation of DirectDraw and Direct3D</a></u></h3>
<p align="left">Now we have declared all the objects we need, we need to call
some of their methods to make them do something. We will also use the variables
to send information to DirectX. Since Direct3D is built upon DirectDraw, we will
need to initialize the DirectDraw objects before Direct3D.</p>
<h4 align="left"><b>The DirectDrawInit Function</b></h4>
<p align="left">We will create a function that creates the DirectDraw object,
sets the cooperative level, sets up the primary and backbuffer surfaces for
our graphics functions to work on, and finally creates a clipper to restrict
drawing to just the application window. Note then when we create the backbuffer
surface, we pass the DDSCAPS_3DDEVICE flag to tell DirectDraw that we are going
to use it as a 3D rendering target.</p>
<div align="left">
 <pre align="left"><font color="#0000FF">Function</font> DirectDrawInit() <font color="#0000FF">As</font> <font color="#0000FF">Long</font>
<font color="#008000">' create the directdraw object
</font><font color="#0000FF">Set</font> DDRAW = DX.DirectDrawCreate(&quot;&quot;)
<font color="#008000">' set the cooperative level, we only need normal
</font>DDRAW.SetCooperativeLevel hWnd, DDSCL_NORMAL
<font color="#008000">' set the properties of the primary surface
</font>SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
<font color="#008000">' create the primary surface
</font><font color="#0000FF">Set</font> Primary = DDRAW.CreateSurface(SurfDesc)
<font color="#008000">' set up the backbuffer surface (which will be where we render the 3D view)
</font>SurfDesc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
<font color="#008000">' use the size of the form to determine the size of the render target
' and viewport rectangle
</font>DX.GetWindowRect hWnd, DestRect
<font color="#008000">' set the dimensions of the surface description
</font>SurfDesc.lWidth = DestRect.Right - DestRect.Left
SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
<font color="#008000">' create the backbuffer surface
</font><font color="#0000FF">Set</font> Backbuffer = DDRAW.CreateSurface(SurfDesc)
<font color="#008000">' cache the size of the render target for later use
</font><font color="#0000FF">With</font> SrcRect
    .Left = 0: .Top = 0
    .Bottom = SurfDesc.lHeight
    .Right = SurfDesc.lWidth
<font color="#0000FF">End</font> <font color="#0000FF">With</font>
<font color="#008000">' create a DirectDrawClipper and attach it to the primary surface.
</font><font color="#0000FF">Set</font> Clipper = DDRAW.CreateClipper(0)
Clipper.SetHWnd hWnd
Primary.SetClipper Clipper
<font color="#008000">' report any errors
</font>DirectDrawInit = Err.Number
<font color="#0000FF">End</font> <font color="#0000FF">Function</font>
</pre>
</div>
 <h4 align="left">The Direct3DInit Function</h4>
 <p align="left">Now we need to initialize all our Direct3D objects. In this
 function, we need to create Direct3D, a rendering device (something that does
 the drawing for us), a material (defines the appearance of polygons), and
 several matrices. The rendering device can be some hardware device like a 3D
 accelerator card, or software emulation. For this tutorial, we will use
 software emulation for simplicity. The matrices are :</p>
<ul>
 <li>
  <p align="left">The world matrix - all objects in world space are
  transformed by this matrix</li>
 <li>
  <p align="left">The view matrix - sets the position of the camera</li>
 <li>
  <p align="left">The projection matrix - defines how Direct3D projects the 3D
  scene onto the 2D surface</li>
</ul>
<div align="left">
 <pre align="left"><font color="#0000FF">Function</font> Direct3DInit() <font color="#0000FF">As</font> <font color="#0000FF">Long</font>
<font color="#008000">' create the direct3d object
</font><font color="#0000FF">Set</font> D3D = DDRAW.GetDirect3D
<font color="#008000">' create the rendering device - we are using software emulation only
</font><font color="#0000FF">Set</font> D3Ddevice = D3D.CreateDevice(&quot;IID_IDirect3DRGBDevice&quot;, Backbuffer)
<font color="#008000">' set the viewport rectangle.
</font>VPdesc.lWidth = DestRect.Right - DestRect.Left
VPdesc.lHeight = DestRect.Bottom - DestRect.Top
VPdesc.minz = 0
VPdesc.maxz = 1
D3Ddevice.SetViewport VPdesc
<font color="#008000">' cache the viewport rectangle for later use
</font><font color="#0000FF">With</font> Viewport(0)
  .X1 = 0: .Y1 = 0
  .X2 = VPdesc.lWidth
  .Y2 = VPdesc.lHeight
<font color="#0000FF">End</font> <font color="#0000FF">With</font>
<font color="#008000">' enable ambient lighting
</font>D3Ddevice.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(1, 1, 1, 1)
<font color="#008000">' disable culling
</font>D3Ddevice.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_NONE
<font color="#008000">' set the material to a red color
</font>Material.Ambient.r = 1
Material.Ambient.g = 0
Material.Ambient.b = 0
D3Ddevice.SetMaterial Material
<font color="#008000">' the world matrix - all polygons in world space are transformed by this matrix
</font>DX.IdentityMatrix matWorld
D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matWorld
<font color="#008000">' the view matrix - basically the camera position is at -3
' (although it's really just making the whole world at +3)
</font>DX.IdentityMatrix matView
DX.ViewMatrix matView, MakeVector(0, 0, -3), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
<font color="#008000">' the projection matrix - decides how the 3D scene is projected onto the 2D surface
</font>DX.IdentityMatrix matProj
DX.ProjectionMatrix matProj, 1, 1000, 3.14 / 2
D3Ddevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj
<font color="#008000">' report errors
</font>Direct3DInit = Err.Number
<font color="#0000FF">End</font> <font color="#0000FF">Function</font>
</pre>
</div>
<h4 align="left">The MakeVector Function</h4>
<p align="left">If you're still alert and haven't become totally confused yet,
you will be saying &quot;hey Simon, you called a MakeVector function - what's
that all about? The MakeVector function is very similar to the
DX.CreateD3DVertex (see later) function - it just saves us alot of typing by
copying values into the D3DVECTOR type. So we need to create the MakeVector
function for the Direct3DInit function to work.</p>
<div align="left">
 <pre align="left"><font color="#0000FF">Function</font> MakeVector(x <font color="#0000FF">As</font> <font color="#0000FF">Single</font>, y <font color="#0000FF">As</font> <font color="#0000FF">Single</font>, z <font color="#0000FF">As</font> <font color="#0000FF">Single</font>) <font color="#0000FF">As</font> D3DVECTOR
<font color="#008000">' copy x, y and z into the return value
</font><font color="#0000FF">With</font> MakeVector
  .x = x
  .y = y
  .z = z
<font color="#0000FF">End</font> <font color="#0000FF">With</font>
<font color="#0000FF">End</font> <font color="#0000FF">Function</font></pre>
</div>
<h3 align="left"><u>Creating The Scene</u></h3>
<p align="left">We need to supply triangles for Direct3D to render. Therefore we
should declare some vertices to make the triangle from. For simplicity, we will
render just one triangle which means we need only 3 vertices (one for each
corner). We could fill in the data separately for each field of the type
D3DVERTEX, but it's much shorter to use a function of the DirectX object that
does this for you in one line of code.</p>
<h4 align="left">The CreateTriangle Sub</h4>
<p align="left">This procedure takes the already declare vertices and forms them
into a triangle shape. In a D3DVERTEX, there are three pieces of data - the
position (x,y,z), the normal (nx,ny,nz) and the texture coordinates (tu,tv). We
only need to use the position in this tutorial. The normal of a triangle is
concerned with lighting, which we aren't using. The texture coordinates are for,
well, textures - which we aren't using either.</p>
<div align="left">
 <pre align="left"><font color="#0000FF">Sub</font> CreateTriangle()
<font color="#008000">' fill in the vertex positions - we don't need to worry about the normals
' or texture coordinates for this tutorial
</font>DX.CreateD3DVertex -1, 0, 0, 0, 0, 0, 0, 0, Vertex(0)
DX.CreateD3DVertex 0, 2, 0, 0, 0, 0, 0, 0, Vertex(1)
DX.CreateD3DVertex 1, 0, 0, 0, 0, 0, 0, 0, Vertex(2)
<font color="#0000FF">End</font> <font color="#0000FF">Sub</font>
</pre>
</div>
<h3 align="left"><u>The Main Program Loop</u></h3>
<p align="left">OK that's enough loading and initializing to last me a lifetime!
But once you've learnt it, it will get easier and you can always reuse your
code. Now we move onto the main program loop. This is a loop where we clear the
backbuffer, draw the polygon, copy the backbuffer to the screen and then move
the polygon before we draw the next frame. Don't be surprised if this loop runs
at over 100 frames per second - after all, it's just one polygon. In a real
world application, you may want to render thousands per frame. On with the show:</p>
<div align="left">
 <pre align="left"><font color="#0000FF">Sub</font> MainLoop()
<font color="#0000FF">Do</font> <font color="#0000FF">While</font> EndNow = False
<font color="#008000">  ' increase the counter
</font>  Counter = Counter + 1
<font color="#008000">  ' clear the viewport with a green color
</font>  D3Ddevice.Clear 1, Viewport(), D3DCLEAR_TARGET, vbGreen, 0, 0
<font color="#008000">  ' begin the scene, render the triangle, then end the scene
</font>  D3Ddevice.BeginScene
  D3Ddevice.DrawPrimitive D3DPT_TRIANGLELIST, D3DFVF_VERTEX, Vertex(0), 3, D3DDP_DEFAULT
  D3Ddevice.EndScene
<font color="#008000">  ' rotate the matrix
</font>  DX.RotateYMatrix matSpin, Counter / 360
<font color="#008000">  ' set the new world transform matrix
</font>  D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matSpin
<font color="#008000">  ' copy the backbuffer to the screen
</font>  DX.GetWindowRect hWnd, DestRect
  Primary.Blt DestRect, Backbuffer, SrcRect, DDBLT_WAIT
<font color="#008000">  ' look for window messages - we need to know when the escape key is pressed
</font>  DoEvents
<font color="#0000FF">Loop</font>
<font color="#0000FF">End</font> <font color="#0000FF">Sub</font>
</pre>
</div>
<h3 align="left"><u>Getting It Together</u></h3>
<p align="left">If you run your program now, nowt will happen at all. This is
because you have created a load of procedures but you haven't called them from
anywhere. This is when you will need to put some code into the Form_Load event,
to do initiation and then the main loop. We will check the return values of the
initiation functions, and if they report errors we will end the program.</p>
<h4 align="left">The Form_Load Event</h4>
<div align="left">
 <pre align="left"><font color="#0000FF">Private</font> <font color="#0000FF">Sub</font> Form_Load()
<font color="#008000">' show the form
</font>Show
<font color="#008000">' call the DirectDrawInit function and exit if it fails
</font><font color="#0000FF">If</font> DirectDrawInit() &lt;&gt; DD_OK <font color="#0000FF">Then</font> <font color="#0000FF">Unload Me</font>
<font color="#008000">' call the Direct3DInit function and exit if it fails
</font><font color="#0000FF">If</font> Direct3DInit() &lt;&gt; DD_OK <font color="#0000FF">Then</font> <font color="#0000FF">Unload Me</font>
<font color="#008000">' create the triangle
</font>CreateTriangle
<font color="#008000">' call the main rendering loop
</font>MainLoop
<font color="#008000">' end the program
</font><font color="#0000FF">Unload Me</font>
<font color="#0000FF">End</font> <font color="#0000FF">Sub</font></pre>
</div>
<h4 align="left">The Form_Unload and Form_KeyDown Events</h4>
<p align="left">There is one more thing to do - end the program! The main loop
is exited if the EndNow variable is set to true - so that's all we need to do.
We can also end the program if the escape key is pressed, by putting the same
code in the Form_KeyDown event.</p>
<div align="left">
 <pre align="left"><font color="#0000FF">Private</font> <font color="#0000FF">Sub</font> Form_Unload(Cancel <font color="#0000FF">As</font> Integer)
EndNow = True
<font color="#0000FF">End</font> <font color="#0000FF">Sub</font>
</pre>
</div>
<div align="left">
 <pre align="left"><font color="#0000FF">Private</font> <font color="#0000FF">Sub</font> Form_KeyDown(KeyCode <font color="#0000FF">As</font> Integer, Shift <font color="#0000FF">As</font> Integer)
<font color="#008000">' end program if escape is pressed
</font><font color="#0000FF">If</font> KeyCode = vbKeyEscape <font color="#0000FF">Then</font> EndNow = True
<font color="#0000FF">End</font> <font color="#0000FF">Sub</font>
</pre>
</div>
<h3 align="left"><u>Run The Program</u></h3>
<p align="left">Run the program. If you've typed it correctly (or just used my
example code), you will see the form has a spinning triangle painted on it. You
can even resize the form and the picture will resize to the form size. When you
close the form or press escape, the program ends.</p>
<h3><u><a name="practice">Summary</a></u></h3>
<p>In this tutorial, we have :</p>
<ul>
 <li>Learnt how to set up DirectDraw surfaces for Direct3D.</li>
 <li>Set up Direct3D, telling it to render on a DirectDraw surface</li>
 <li>Create a very basic geometric shape</li>
 <li>Render a triangle and change the world matrix to move spin the world</li>
</ul>
<p>There are many bad points to the program you have created, although I have
made the program in this way to make it as simple as possible.</p>
<ul>
 <li>All the variables were global - in my opinion you should restrict access
  to each variable as much as possible. I made them all global for this
  tutorial so I could explain each one at the beginning</li>
 <li>Very little error handling was done. In a real application, we would find
  the cause of the error, attempt to fix it, and if that's not possible we
  would tell the user why, rather than ending immediately.</li>
 <li>We used software rendering only. What we should do is find out what sort
  of hardware the user has, and make our program adapt to either make maximum
  use of the hardware, or fall back onto just software if no hardware is
  available.</li>
 <li>And I'm sure the critics amongst you will think of more.</li>
</ul>
<h3><u><a name="bingo">Exercises</a></u></h3>
<p>You can only learn something if you actually practice doing it. So here I
have some features which you can add to the program yourself. Come on, be a
little creative and start making your own 3D graphics!</p>
<ul>
 <li>That triangle is boring! It's even looks 2D! Use more vertices to make
  another shape - a cube, a pyramid, a sphere if you're smart enough -
  whatever you like!</li>
 <li>Make a frame counter, so that you know how fast the program is running. I
  bet it goes at over 100 FPS!</li>
 <li>Change the colors to something you like.</li>
 <li>Explore more Direct3D functions, meddle with the code, make it your
  program. I don't want here any complaints that this tutorial was boring -
  it's up to you to make it interesting!</li>
</ul>
<p>I hope I've set you along the exciting journey towards creating Direct3D
graphics from Visual Basic. This tutorial has taken me <b>ALOT</b> of time and
effort - I had to write code, make comments, write a tutorial, get it as
accurate as possible. I would appreciate in return:</p>
<ul>
 <li><b><u>Please vote for me</u></b> - Whether you think this tutorial was
  good or bad, I want to know about it.</li>
 <li><b><u>Please give me some feedback</u></b> - Tell me why you voted the
  score that you did.</li>
 <li><b><u>Please visit my website</u></b> - If you liked this then you'll want
  to visit my website to see more of my programs and tutorials. The URL is <a href="http://www.VBgames.co.uk">www.VBgames.co.uk</a>
 </li>
 <li><b><u>Please give me $30000 to write a book</u></b> - OK only joking.</li>
</ul>
<p>Tutorial by <b>Simon Price</b>, you can email me at <a href="mailto:Si@VBgames.co.uk">Si@VBgames.co.uk</a>
</p>
<p align="left"><a href="#lesson">Back To Top</a></p>

