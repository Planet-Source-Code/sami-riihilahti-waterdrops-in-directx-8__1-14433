Attribute VB_Name = "Inits"
''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Initializations
''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Function InitDirectX() As Boolean
    On Local Error Resume Next
    Set Engine.DX = New DirectX8
    If Engine.DX Is Nothing Then Exit Function 'DirectX not created properly
    
    InitDirectX = True
End Function

Public Function InitD3D(hwnd As Long, ScreenWidth As Integer, ScreenHeight As Integer) As Boolean
    On Local Error Resume Next
    
    ' Create the D3D object
    ' Exit if we can't create it
    Set Engine.D3D = Engine.DX.Direct3DCreate()
    If Engine.D3D Is Nothing Then Exit Function
    
    ' Get The current Display Mode
    Dim mode As D3DDISPLAYMODE
    Engine.D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
         
    ' Using 16 bit z buffer
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = False 'Not windowed
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY 'We also could render scene in v-sync and so on..
    d3dpp.BackBufferFormat = mode.Format ' Create backbuffer with same format as desktop
    d3dpp.BackBufferCount = 1 ' We could set this, but it lowers/stabilizes framerate
    d3dpp.BackBufferHeight = ScreenHeight
    d3dpp.BackBufferWidth = ScreenWidth
    d3dpp.EnableAutoDepthStencil = 0 ' We dont need z-buffer for this
    'd3dpp.AutoDepthStencilFormat = D3DFMT_D16 ' We could use z-buffer, but we dont
    d3dpp.hDeviceWindow = hwnd ' This is nessecary for some cases

    ' Create the D3DDevice
    Set Engine.D3Ddevice = Engine.D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, _
        D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'Check for errors
    If Engine.D3Ddevice Is Nothing Then Exit Function

    InitD3D = True
End Function


Public Function InitKeyboard() As Boolean
    On Local Error Resume Next
    
    Set Engine.dInput = Engine.DX.DirectInputCreate()
    Set Engine.dInputDevice = Engine.dInput.CreateDevice("GUID_SysKeyboard")
    If Engine.dInputDevice Is Nothing Then Exit Function
    
    Engine.dInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    Engine.dInputDevice.SetCooperativeLevel Form_main.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    Engine.dInputDevice.Acquire
    
    If Err.Number = 0 Then InitKeyboard = True
End Function


