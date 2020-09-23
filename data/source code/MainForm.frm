VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CylinderView by Simon Price"
   ClientHeight    =   5376
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4536
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Panoramic Image"
      Height          =   1692
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4332
      Begin VB.FileListBox File 
         Height          =   840
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4092
      End
      Begin VB.Label Label3 
         Caption         =   "Select an image file. These must be bitmaps places in the \data\maps directory."
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   4092
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   1680
      TabIndex        =   7
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number of polygons"
      Height          =   1572
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4332
      Begin VB.HScrollBar PolyScroll 
         Height          =   252
         LargeChange     =   50
         Left            =   120
         Max             =   508
         Min             =   20
         SmallChange     =   10
         TabIndex        =   4
         Top             =   240
         Value           =   508
         Width           =   4092
      End
      Begin VB.Label Label2 
         Caption         =   $"MainForm.frx":0000
         Height          =   612
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   4092
      End
      Begin VB.Label lblPolyCount 
         Alignment       =   2  'Center
         Caption         =   "Polygon Count = 508"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4092
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Renderer"
      Height          =   1452
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4332
      Begin VB.CheckBox chkHardware 
         Caption         =   "Hardware Acceleration"
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   4092
      End
      Begin VB.Label Label1 
         Caption         =   $"MainForm.frx":008D
         Height          =   852
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4092
      End
   End
   Begin VB.Timer FPStimer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' the main directx object
Private DX As New DirectX7

' the main directdraw object
Private DDRAW As DirectDraw7
Private Primary As DirectDrawSurface7 ' primary surface
Private BackBuffer As DirectDrawSurface7 ' back buffer
Private SurfDesc As DDSURFACEDESC2 ' surface desription
Private SurfCaps As DDSCAPS2 ' surface capabilities

' the main direct3d object
Private D3D As Direct3D7
Private D3Ddev As Direct3DDevice7 ' rendering device
Private VPdesc As D3DVIEWPORT7 ' viewport description
Private Viewport(0) As D3DRECT ' viewport
Private RenderTarget As DirectDrawSurface7 ' rendering surface
Private Material As D3DMATERIAL7 ' the polygon surface material
Private Texture As DirectDrawSurface7  ' the texture we are using
Private CameraDir As D3DVECTOR ' direction of camera

' the polygon type - holds polygon info
Private Type tPolygon
    v() As D3DVERTEX ' each will have a number of vertices
End Type
' declare polygons
Private Polygon() As tPolygon

' the main directinput object
Private DINPUT As DirectInput
Private DIdev As DirectInputDevice ' device (mouse)
Private DIstate As DIMOUSESTATE ' mouse state

' look for windows messages (allows for faster loops!)
'Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

' camera rotation speed
Private Const CAMERA_SPEED = 0.001

' for counting the FPS
Private FPS As Long
Private Frames As Long

' signals the user has pressed escape
Private EndNow As Boolean



' load directdraw, direct3d and directinput and returns true if successful
Function Init(Optional UseHardware As Boolean = False) As Boolean
On Error Resume Next

' create directdraw
Set DDRAW = DX.DirectDrawCreate("")
' set coop level
DDRAW.SetCooperativeLevel hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
' set display mode
DDRAW.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
If Err.Number <> DD_OK Then Exit Function

' create primary surface
With SurfDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    .lFlags = .lFlags Or DDSD_BACKBUFFERCOUNT
    .lBackBufferCount = 1
End With
Set Primary = DDRAW.CreateSurface(SurfDesc)
If Err.Number <> DD_OK Then Exit Function

' create backbuffer surface
With SurfDesc
    SurfCaps.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(SurfCaps)
End With
BackBuffer.SetForeColor vbRed
If Err.Number <> DD_OK Then Exit Function

' create render target
With SurfDesc
    .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    If UseHardware Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_VIDEOMEMORY
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_SYSTEMMEMORY
    End If
    .lWidth = 640
    .lHeight = 480
End With
Set RenderTarget = DDRAW.CreateSurface(SurfDesc)
If Err.Number <> DD_OK Then Exit Function

' get d3d and rendering device
Dim Guid As String
If UseHardware Then Guid = "IID_IDirect3DHALDevice" Else Guid = "IID_IDirect3DRGBDevice"
Set D3D = DDRAW.GetDirect3D
Set D3Ddev = D3D.CreateDevice(Guid, RenderTarget)
If Err.Number <> DD_OK Then Exit Function

' set viewport
VPdesc.lWidth = 640
VPdesc.lHeight = 480
VPdesc.minz = 0.01
VPdesc.maxz = 1000000
D3Ddev.SetViewport VPdesc
With Viewport(0)
    .X1 = 0: .Y1 = 0
    .X2 = 640
    .Y2 = 480
End With
If Err.Number <> DD_OK Then Exit Function

' set projection matrix
Dim matProj As D3DMATRIX
DX.IdentityMatrix matProj
DX.ProjectionMatrix matProj, 0.01, 100000, PI / 3
matProj.rc11 = matProj.rc11 * 480 / 640
D3Ddev.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj
If Err.Number <> DD_OK Then Exit Function

' set material
Material.Ambient = modMath.MakeD3DCOLORVALUE(1, 1, 1, 1)
Material.diffuse = modMath.MakeD3DCOLORVALUE(1, 1, 1, 1)
D3Ddev.SetMaterial Material

' set rendering options
D3Ddev.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_NONE
'D3Ddev.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_WIREFRAME
D3Ddev.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(1, 1, 1, 1)

' get directinput
Set DINPUT = DX.DirectInputCreate
' get mouse device
Set DIdev = DINPUT.CreateDevice("GUID_SysMouse")
DIdev.SetCommonDataFormat DIFORMAT_MOUSE
DIdev.SetCooperativeLevel hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
DIdev.Acquire

' report success/failure
If Err.Number = DD_OK Then Init = True
End Function

' unloads stuff in reverse order of what was loaded
Sub CleanUp()
On Error Resume Next

' get rid of mouse
DIdev.Unacquire
' set di stuff to nothing
Set DIdev = Nothing
Set DINPUT = Nothing

' set d3d stuff to nothing
Set RenderTarget = Nothing
Set D3Ddev = Nothing
Set D3D = Nothing

' ste ddraw stuff to nothing
Set BackBuffer = Nothing
Set Primary = Nothing
Set DDRAW = Nothing
End Sub

Private Sub cmdOK_Click()
If File.Filename = "" Then
    MsgBox "Please select an image file first!"
    Exit Sub
End If
' load stuff, unload and report on failure
If Init(chkHardware.Value) Then
    ' load map
    If LoadTexture(File.Path & "\" & File.Filename) = False Then Debug.Print "tex err"
    ' create sphere
    'MakeSphere modMath.MakeVector(100, 100, 100), 254
    ' create cylinder
    MakeCylinder modMath.MakeVector(100, 200, 100), PolyScroll.Value \ 2
    ' start main rendering loop
    MainLoop
Else
    MsgBox "ERROR : Could not initialize DirectX!", vbCritical, "ERROR!"
End If
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        ' end program
        EndNow = True
End Select
End Sub

Sub MainLoop()
On Error Resume Next
Dim i As Long
Dim Rotation As D3DVECTOR
Do
    DoEvents
    ' get mouse movements
    DIdev.GetDeviceStateMouse DIstate
    Rotation.x = DIstate.y * CAMERA_SPEED
    Rotation.y = DIstate.x * CAMERA_SPEED
    ' rotate camera
    RotateCamera Rotation
    ' clear render target
    'D3Ddev.Clear 1, Viewport(), D3DCLEAR_TARGET, vbBlack, 0, 0
    ' render each polygon
    For i = 0 To UBound(Polygon)
        With Polygon(i)
            D3Ddev.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, .v(0), UBound(.v) + 1, D3DDP_DEFAULT
        End With
    Next
    ' copy to backbuffer
    BackBuffer.BltFast 0, 0, RenderTarget, MakeRect(0, 0, 640, 480), DDBLTFAST_WAIT
    ' write frame rate on backbuffer
    BackBuffer.DrawText 0, 0, "FPS = " & FPS, True
    ' flip into view
    Primary.Flip Nothing, DDFLIP_WAIT
    ' increase frame counter
    Frames = Frames + 1
    If EndNow = True Then Exit Do
Loop
End Sub

Private Sub Form_Load()
' set file path
File.Path = App.Path & "\data\maps\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' we need to cleanup dx stuff
CleanUp
End Sub

' loads a bitmap onto the texture surface and reports success/failure
Function LoadTexture(Filename As String) As Boolean
On Error Resume Next
Dim i As Long
Dim IsFound As Boolean
Dim TextureEnum As Direct3DEnumPixelFormats
' set description flags
SurfDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT Or DDSD_TEXTURESTAGE
' get texture formats
Set TextureEnum = D3Ddev.GetTextureFormatsEnum()
' check each one until a suitable one is found
For i = 1 To TextureEnum.GetCount()
    IsFound = True
    TextureEnum.GetItem i, SurfDesc.ddpfPixelFormat
    With SurfDesc.ddpfPixelFormat
        If .lFlags And (DDPF_LUMINANCE Or DDPF_BUMPLUMINANCE Or DDPF_BUMPDUDV) Then IsFound = False
        If .lFourCC <> 0 Then IsFound = False
        If .lFlags And DDPF_ALPHAPIXELS Then IsFound = False
        If .lRGBBitCount <> 16 Then IsFound = False
    End With
    If IsFound Then Exit For
Next i
' if we still haven't found a texture format, we failed
If Not IsFound Then Exit Function
' create texture surface
SurfDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
SurfDesc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
SurfDesc.lWidth = 0
SurfDesc.lHeight = 0
SurfDesc.lTextureStage = 0
Set Texture = DDRAW.CreateSurfaceFromFile(Filename, SurfDesc)
D3Ddev.SetTexture 0, Texture
' report if successful
If Err.Number = DD_OK Then LoadTexture = True
End Function

Sub MakeCylinder(Size As D3DVECTOR, Quality As Long)
On Error Resume Next
Dim tempPoly As tPolygon
Dim i As Long
MakeCircle tempPoly, Size.x, Quality
ReDim Polygon(0 To Quality)
For i = 0 To UBound(Polygon)
    With Polygon(i)
        ReDim Polygon(i).v(0 To 3)
        .v(0) = tempPoly.v(i + 1)
        .v(0).y = Size.y / 2
        .v(0).tv = 0
        .v(1) = tempPoly.v(i + 2)
        .v(1).y = Size.y / 2
        .v(1).tv = 0
        .v(2) = tempPoly.v(i + 1)
        .v(2).y = -Size.y / 2
        .v(2).tv = 1
        .v(3) = tempPoly.v(i + 2)
        .v(3).y = -Size.y / 2
        .v(3).tv = 1
    End With
Next
End Sub

' makes a circle
Private Sub MakeCircle(ThePoly As tPolygon, Radius As Single, Quality As Long)
Dim i As Byte
Dim DiffAngle As Double
Dim CurAngle As Double
Dim DiffTu As Double
Dim CurTu As Single
With ThePoly
    ReDim .v(0 To Quality + 1)
    SetVertexWON .v(0), 0, 0, 0, 0, 0
    DiffAngle = 2 * PI / Quality
    DiffTu = 1 / Quality
    For i = 1 To Quality
        SetVertexWON .v(i), Sin(CurAngle) * Radius, 0, Cos(CurAngle) * Radius, CurTu, 1
        CurAngle = CurAngle + DiffAngle
        CurTu = CurTu + DiffTu
    Next
    CurAngle = PIby2
    CurTu = 1
    SetVertexWON .v(i), Sin(CurAngle) * Radius, 0, Cos(CurAngle) * Radius, CurTu, 1
End With
End Sub

' makes a sphere
Sub MakeSphere(Size As D3DVECTOR, Quality As Long)
On Error Resume Next
Dim i As Long
Dim i2 As Long
Dim ti As Integer
Dim ti2 As Integer
Dim NumSides As Long
Dim Angle As Double
Dim stepAngle As Double
Dim DiffTv As Double
Dim CurTv As Double
Dim r As Single
Dim N As D3DVECTOR
Dim tempPoly() As tPolygon
NumSides = Quality - 1
ReDim tempPoly(0 To NumSides)
stepAngle = PIby2 / NumSides
DiffTv = 2 / NumSides
For i = 0 To NumSides
    r = Size.x
    MakeRegularPoly tempPoly(i), r, r, Quality
    N.x = Angle
    RotatePoly tempPoly(i), N
    Angle = Angle + stepAngle
Next
ReDim Polygon(0 To NumSides)
For i = 0 To NumSides
    With Polygon(i)
        If i = NumSides Then
            ti = -1
        Else
            ti = i
        End If
        ReDim .v(0 To NumSides + 1)
        For i2 = 0 To NumSides \ 2
            .v(i2 * 2) = tempPoly(i).v(i2 + 1)
            .v(i2 * 2 + 1) = tempPoly(ti + 1).v(i2 + 1)
        Next
        .v(NumSides + 1) = tempPoly(0).v(NumSides \ 2 + 2)
        For i2 = 0 To NumSides
            If i2 Mod 2 = 0 Then
                .v(i2).tv = CurTv
            Else
                .v(i2).tv = CurTv + DiffTv
            End If
        Next
        CurTv = CurTv + DiffTv
    End With
Next
End Sub

' makes a polygon of any number of sides
Private Sub MakeRegularPoly(ThePoly As tPolygon, w As Single, h As Single, NumSides As Long)
Dim i As Byte
Dim DiffAngle As Double
Dim CurAngle As Double
Dim DiffTu As Double
Dim CurTu As Single
w = w / 2
h = h / 2
With ThePoly
    ReDim .v(0 To NumSides + 1)
    SetVertexWON .v(0), 0, 0, 0, 0, 0
    DiffAngle = 2 * PI / NumSides
    DiffTu = 1 / NumSides
    For i = 1 To NumSides
         SetVertexWON .v(i), Sin(CurAngle) * w, Cos(CurAngle) * h, 0, CurTu, 1
         CurAngle = CurAngle + DiffAngle
         CurTu = CurTu + DiffTu
    Next
    CurAngle = PIby2
    CurTu = 1
    SetVertexWON .v(i), Sin(CurAngle) * w, Cos(CurAngle) * h, 0, CurTu, 1
End With
End Sub

' rotates the polygon around the world origin
Private Sub RotatePoly(ThePoly As tPolygon, N As D3DVECTOR)
Dim C As D3DVECTOR
Dim C2 As D3DVECTOR
Dim vec As D3DVECTOR
Dim i As Byte
With ThePoly
For i = 0 To UBound(.v)
    modMath.CopyVertex2Vec .v(i), vec
    vec = modMath.RotateYVectorAroundVector(vec, N.x, C)
    vec = modMath.RotateXVectorAroundVector(vec, N.y, C)
    vec = modMath.RotateZVectorAroundVector(vec, N.Z, C)
    modMath.CopyVec2Vertex vec, .v(i)
Next
End With
End Sub

' sets a vertex without a normal
Sub SetVertexWON(v As D3DVERTEX, x As Single, y As Single, Z As Single, tu As Single, tv As Single)
With v
   .x = x
   .y = y
   .Z = Z
   .tu = tu
   .tv = tv
End With
End Sub

Sub RotateCamera(Rotation As D3DVECTOR)
On Error Resume Next
' direction vector
Dim N As D3DVECTOR
N = CameraDir
' rotate on x-axis
If Rotation.x Then
    N.x = N.x + Rotation.x
    If N.x > PIby2 Then N.x = N.x - PIby2
    If N.x < 0 Then N.x = N.x + PIby2
End If
' rotate on y-axis
If Rotation.y Then
    N.y = N.y + Rotation.y
    If N.y > PIby2 Then N.y = N.y - PIby2
    If N.y < 0 Then N.y = N.y + PIby2
End If
' rotate on z-axis
If Rotation.Z Then
    N.Z = N.Z + Rotation.Z
    If N.Z > PIby2 Then N.Z = N.Z - PIby2
    If N.Z < 0 Then N.Z = N.Z + PIby2
End If
Select Case N.x
    Case 0.25 To PIby2 - 0.25
' don't rotate camera
    Case Else
        CameraDir = N
End Select

' matrices
Dim matView As D3DMATRIX
Dim matTemp As D3DMATRIX
Dim matRot As D3DMATRIX

DX.IdentityMatrix matView
matView.rc41 = vec.x
matView.rc42 = vec.y
matView.rc43 = vec.Z

DX.IdentityMatrix matRot
DX.RotateYMatrix matTemp, N.y
DX.MatrixMultiply matRot, matRot, matTemp
DX.RotateXMatrix matTemp, N.x
DX.MatrixMultiply matRot, matRot, matTemp
DX.RotateZMatrix matTemp, N.Z
DX.MatrixMultiply matRot, matRot, matTemp

' combine matrices and set camera transformation
DX.MatrixMultiply matView, matView, matRot
D3Ddev.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub

Private Sub FPStimer_Timer()
FPS = Frames
Frames = 0
End Sub

Private Sub PolyScroll_Change()
lblPolyCount = "Polygon Count = " & PolyScroll
End Sub
