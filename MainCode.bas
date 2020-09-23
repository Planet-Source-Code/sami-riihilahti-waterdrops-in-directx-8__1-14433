Attribute VB_Name = "MainCode"
Option Explicit

'We are using gettickcount to calculate the exact FPS
'And for animation timing
Declare Function GetTickCount Lib "kernel32" () As Long

Const WaterDropCount = 200 'Change this to render less/more waterdrops
Const ScreenWidth = 800
Const ScreenHeight = 600

'Flexible vertex format (transformed and lit vertices)
Const FVF_SPRITE = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Transformed and lit vertex structure (same as D3DTLVERTEX)
Private Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Public Type XEngine 'This is the main 3D-engine structure to keep variables simple
    DX As New DirectX8 'Main directX -object (in case you didn't know) :)
    D3DX As New D3DX8 'X-helper object
    D3D As Direct3D8 'Main Direct3D
    D3Ddevice As Direct3DDevice8 'Rendering device
    dInput As DirectInput8 'DirectInput from keyboard
    dInputDevice As DirectInputDevice8 'DirectInput Device
    framecount As Long 'Frame rate counter (just a counter, not current framerate)
    AFrameCount As Long 'Average framerate count (just a counter, not current framerate)
    RunGame As Boolean 'If set to false, the game should exit
    KeyState As DIKEYBOARDSTATE 'Array to store keyboard status
    BackGround As Direct3DSurface8 'Background surface
    WaterDrop As Direct3DTexture8 'WaterDrop texture
End Type
Public Type XWaterDrop
    Xposition As Integer 'x
    Yposition As Integer 'y
    AnimTime As Long 'Animation timing
    MaxSize As Single 'max drop size
End Type
Public Type XTLmesh
    WaterDrop(3) As TLVERTEX
End Type

Public WaterDrop() As XWaterDrop
Public TLmesh As XTLmesh
Public TruePath As String 'Fixes root path error
Public RetSuccess As Boolean
Public FPS As Integer

'Drawing text
Public MainFont As D3DXFont 'This is the main font object
Public MainFontDesc As IFont 'We use this temporarily to setup the font
Public TextRect As RECT 'Text draw place
Public fnt As New StdFont 'Used to describe and setup the font

Public Engine As XEngine

Public Sub PreSetWaterDrops(WaterDrops As Integer, areaWidth As Integer, areaHeight As Integer)
    
    Dim wdCount As Integer
    Dim cTime As Long
    
    'Preset WaterDrop TLmesh (uv-mapping and rhw stays unchanged)
    With TLmesh.WaterDrop(0)
        .tu = 0: .tv = 1
        .rhw = 1
    End With
    With TLmesh.WaterDrop(1)
        .tu = 0: .tv = 0
        .rhw = 1
    End With
    With TLmesh.WaterDrop(2)
        .tu = 1: .tv = 1
        .rhw = 1
    End With
    With TLmesh.WaterDrop(3)
        .tu = 1: .tv = 0
        .rhw = 1
    End With
    
    'Make room for waterdrop array
    ReDim WaterDrop(WaterDrops - 1)
    
    cTime = GetTickCount 'animation start time timer
    
    'Randomize positions and sizes
    For wdCount = 0 To WaterDrops - 1
        With WaterDrop(wdCount)
            '.Size = 0 'Already set to 0
            .MaxSize = Rnd + 0.1 'From 0.1 to 1.1
            .AnimTime = Int(Rnd * 5000) + cTime 'max 5 second pause before first drop
            .Xposition = (Rnd * (areaWidth - 64)) + 32
            .Yposition = (Rnd * (areaHeight - 64)) + 32
        End With
    Next
    
End Sub
Public Sub CleanDirectX()

    'Clean all surfaces
    Set Engine.BackGround = Nothing
    Set Engine.WaterDrop = Nothing
    
    'Clean DirectX
    If Not (Engine.D3D Is Nothing) Then Set Engine.D3D = Nothing
    If Not (Engine.D3Ddevice Is Nothing) Then Set Engine.D3Ddevice = Nothing
    If Not (Engine.DX Is Nothing) Then Set Engine.DX = Nothing
End Sub

Public Sub Main()

'Root path error fix
If Len(App.Path) = 3 Then
    TruePath = Left(App.Path, 2)
Else:
    TruePath = App.Path
End If

RetSuccess = InitDirectX
If Not RetSuccess Then
    MsgBox ("DirectX initialization failed")
    End
End If

RetSuccess = InitKeyboard
If Not RetSuccess Then
    MsgBox ("Keyboard initialization failed")
    End
End If

RetSuccess = InitD3D(Form_main.hwnd, ScreenWidth, ScreenHeight) 'Initialize d3d
If Not RetSuccess Then
    MsgBox ("Direct3D initialization failed")
    End
End If

Form_main.Show 'Show main form
DoEvents

StartMainScreen  'Start main screen

End Sub

Public Sub StartMainScreen()
    
    'Setup font for rendering text
    fnt.Name = "Arial"
    fnt.Size = 10
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = Engine.D3DX.CreateFont(Engine.D3Ddevice, MainFontDesc.hFont)
    
    TextRect.Top = 5
    TextRect.Left = 5
    TextRect.bottom = 20
    TextRect.Right = 50

    'Load start screen image
    'We could use Image surfaces from system memory, but they are about 10x slower
    'We also could use textures, but some video-cards supports textures only
    'up to 256x256 sizes. (We are using 800x600 image)
    
    Set Engine.BackGround = Engine.D3Ddevice.CreateRenderTarget(ScreenWidth _
        , ScreenHeight, D3DFMT_X8R8G8B8, D3DMULTISAMPLE_NONE, True)
    Engine.D3DX.LoadSurfaceFromFile Engine.BackGround, ByVal 0&, ByVal 0&, TruePath & "\images\Background.jpg", ByVal 0&, D3DX_FILTER_NONE, 0, ByVal 0&
    
    'Load waterdrop image to a texture
    Set Engine.WaterDrop = Engine.D3DX.CreateTextureFromFileEx(Engine.D3Ddevice _
        , TruePath & "\images\wDrop.bmp", 64, 64, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
        D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0&, ByVal 0&)
    
    'Preset the rendering options
     With Engine.D3Ddevice
        .BeginScene
        .SetVertexShader FVF_SPRITE
        'Set the main screen background texture on the device
        .SetTexture 0, Engine.WaterDrop
        .EndScene
    End With
    
    ' Turn off the zbuffer (we dont need it)
    Engine.D3Ddevice.SetRenderState D3DRS_ZENABLE, 0
    ' Turn on lighting
    Engine.D3Ddevice.SetRenderState D3DRS_LIGHTING, 1
    ' Turn on full ambient light to white
    Engine.D3Ddevice.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
    Engine.D3Ddevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID 'Normal fillmode
    Engine.D3Ddevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD 'Set smooth polygons
    Engine.D3Ddevice.SetRenderState D3DRS_CLIPPING, False 'In this point, we don't need clipping
    Engine.D3Ddevice.SetRenderState D3DRS_COLORVERTEX, True 'We are using colored vertexs
    Engine.D3Ddevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    Engine.D3Ddevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    'Setup transparency mode for waterdrops
    Engine.D3Ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    Engine.D3Ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR

    'Preset water drops
    PreSetWaterDrops WaterDropCount, ScreenWidth, ScreenHeight
    
    Dim framecount As Long
    Dim frametimer As Long
    frametimer = GetTickCount
    
    Do
        RenderMainScreen
        framecount = framecount + 1
        If GetTickCount >= frametimer + 1000 Then
            FPS = framecount
            framecount = 0
            frametimer = GetTickCount
        End If
        DoEvents
        'Get keyboard strokes
        Engine.dInputDevice.GetDeviceStateKeyboard Engine.KeyState
    Loop Until (Engine.KeyState.Key(DIK_ESCAPE) And &H80) 'Escape key pressed
    
    'Exit
    CleanDirectX
    Unload Form_main
    End
    
End Sub

Public Sub RenderMainScreen()
    
    Dim wdCount As Integer 'Water drop count
    Dim cTime As Long
    Dim DropSize As Single
    Dim xmin As Single
    Dim xmax As Single
    Dim ymin As Single
    Dim ymax As Single
    Dim AnimSize As Single
    Dim DropColor As Long
    
    'Clear screen to a black color
    Engine.D3Ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1, 0
    
    'Set alpha transparency off while rendering background
    Engine.D3Ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    
    'Copy background image to back buffer
    Engine.D3Ddevice.CopyRects Engine.BackGround, _
        ByVal 0&, 0, Engine.D3Ddevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), ByVal 0&
    
    'Set alpha transparency on for waterdrops
    Engine.D3Ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    cTime = GetTickCount 'Timer for animations
    
    'Begin the scene for '3D'-rendering
    Engine.D3Ddevice.BeginScene

    For wdCount = 0 To WaterDropCount - 1
        If WaterDrop(wdCount).AnimTime <= cTime - 600 Then
        
                'Randomize new waterdrop if previous has animated fully
                With WaterDrop(wdCount)
                    .AnimTime = cTime + (Rnd * 1000) + 600 'Max 1 second pause before blitting again
                    .MaxSize = Rnd + 0.1 'From 0.1 to 1.1
                    .Xposition = (Rnd * (800 - 64)) + 32
                    .Yposition = (Rnd * (600 - 64)) + 32
                End With
        
        ElseIf WaterDrop(wdCount).AnimTime <= cTime Then
        
            'Time to make animation (this part can be quite confusing, but you don't really have to understand it)
            AnimSize = 600 - (cTime - WaterDrop(wdCount).AnimTime)
            If AnimSize <= 1 Then AnimSize = 1
            DropColor = D3DColorXRGB(AnimSize / 2, AnimSize / 2, AnimSize / 2)
            AnimSize = (600 - AnimSize) / 600
            xmin = WaterDrop(wdCount).Xposition - (32 * AnimSize) * WaterDrop(wdCount).MaxSize
            xmax = xmin + (64 * AnimSize) * WaterDrop(wdCount).MaxSize
            ymin = WaterDrop(wdCount).Yposition - (32 * AnimSize) * WaterDrop(wdCount).MaxSize
            ymax = ymin + (64 * AnimSize) * WaterDrop(wdCount).MaxSize
            
            With TLmesh.WaterDrop(0)
                .x = xmin: .y = ymax
                .color = DropColor
            End With
            With TLmesh.WaterDrop(1)
                .x = xmin: .y = ymin
                .color = DropColor
            End With
            With TLmesh.WaterDrop(2)
                .x = xmax: .y = ymax
                .color = DropColor
            End With
            With TLmesh.WaterDrop(3)
                .x = xmax: .y = ymin
                .color = DropColor
            End With
            
            'Draw the 2 polygons that makes the waterdrop
            Engine.D3Ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLmesh.WaterDrop(0), Len(TLmesh.WaterDrop(0))
        End If
    Next
    
    'Render FPS-text
    Engine.D3DX.DrawText MainFont, &HFFFFFFFF, FPS & " FPS", TextRect, DT_TOP Or DT_LEFT
    
    'End the scene
    Engine.D3Ddevice.EndScene
        
    'Present the backbuffer to the screen
    Engine.D3Ddevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&

End Sub

