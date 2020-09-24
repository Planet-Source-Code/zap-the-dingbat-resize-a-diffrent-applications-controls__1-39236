Attribute VB_Name = "Module1"
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const GWL_STYLE = (-16)

Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const SWP_DRAWFRAME = &H20

Const WS_THICKFRAME = &H40000
Const WS_BORDER = &H800000

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal Window As Long, ByVal Buffer As String, ByVal BufferLength As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Const GW_CHILD = 5
Const GW_HWNDNEXT = 2

Const GWL_EXSTYLE = (-20)
Const WS_EX_WINDOWEDGE = &H100
Const WS_EX_CLIENTEDGE = &H200
Const WS_EX_STATICEDGE = &H20000

Public hSelected As Long

Sub main()

    'clear nodes
    frmMain.treWindows.Nodes.Clear
    'get desktop hWnd
    Dim hDesktop As Long
    hDesktop = GetDesktopWindow
    'add desktop node
    addWindowNode hDesktop
    'populate child windows
    populateNode hDesktop

    frmMain.Show
End Sub

'add a node
Sub addWindowNode(hWindow As Long, Optional hParent As Long = 0)
    
    Dim strNodeKey As String
    Dim strNodeText As String
    
    strNodeKey = KeyFromhWnd(hWindow)
    strNodeText = Hex(hWindow) & " " & windowClass(hWindow) & " '" & windowText(hWindow) & "'"
    
    Debug.Print "Made " & hWindow, strNodeKey
    
    'add node
    If hParent = 0 Then
        'desktop
        frmMain.treWindows.Nodes.Add , tvwFirst, strNodeKey, strNodeText, 1
    Else
        'normal
        frmMain.treWindows.Nodes.Add KeyFromhWnd(hParent), tvwChild, strNodeKey, strNodeText, 1
    End If
    
End Sub

'add children to an existing node
Sub populateNode(hWindow As Long)

    Dim hChild As Long
    
    'get first child handle
    hChild = GetWindow(hWindow, GW_CHILD)
    
    'for each child
    Do Until hChild = 0
        
        'if the window is visible and hidden checkbox is checked
        If frmMain.mnuHidden.Checked Or IsWindowVisible(hChild) Then
            'add node
            addWindowNode hChild, hWindow
            'call once more
            populateNode hChild
        End If
        
        'move next
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
    
End Sub

'on window select
Sub selectWindow(hWindow As Long)

    Debug.Print "Selected " & hWindow
    If Not hWindow = hSelected Then
        'FlashWindow
        FlashWindow hWindow, 1
        'update selected windw
        hSelected = hWindow
        
        'if hWindow is being resized, check menu items
        lStyle = GetWindowLong(hWindow, GWL_STYLE)
        If lStyle = (lStyle Or WS_THICKFRAME) Then
            frmMain.mnuResize.Checked = True
            frmMain.mnuResizeAll.Checked = False
        Else
            frmMain.mnuResize.Checked = False
            frmMain.mnuResizeAll.Checked = False
        End If
    End If
    
End Sub
'derives hWnd from unique string
Function hWndFromKey(strKey As String) As Long
    hWndFromKey = Val("&h" & Mid(strKey, 2)) 'hex to dec
End Function
'derives unique string from Hwnd
Function KeyFromhWnd(hWindow As Long) As String
    KeyFromhWnd = "h" & Hex(hWindow) '"h" AND dec to hex
End Function
'get window class name
Function windowClass(hWindow As Long) As String
    Dim strBuffer As String * 255
    Dim lngLength As Long
    lngLength = GetClassName(hWindow, strBuffer, 255)
    windowClass = Left(strBuffer, lngLength)
End Function
'get window title
Function windowText(hWindow As Long) As String
    Dim strBuffer As String * 255
    Dim lngLength As Long
    lngLength = GetWindowText(hWindow, strBuffer, 255)
    windowText = Left(strBuffer, lngLength)
End Function

'recursive window handleing loops
Sub resizeAll(hwnd)
    Dim hChild As Long
    'get first child handle
    hChild = GetWindow(hwnd, GW_CHILD)
    'for each child
    Do Until hChild = 0
        'add node
        If IsWindowVisible(hChild) Then resizeWindow (hChild)
        'call this function again for the child
        resizeAll hChild
        'move next
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
End Sub

Sub fixAll(hwnd)
    Dim hChild As Long
    'get first child handle
    hChild = GetWindow(hwnd, GW_CHILD)
    
    'for each child
    Do Until hChild = 0
        'add node
        If IsWindowVisible(hChild) Then fixWindow (hChild)
        'call this function again for the child
        fixAll hChild
        'move next
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
End Sub

'single window handleing
Sub resizeWindow(hWindow As Long)

    Dim lStyle As Long
    Dim hParent As Long
    
    hParent = GetParent(hWindow)
    
    Debug.Print "Resizeing " & hWindow, hParent
    'check for parent
    
    If hParent = 0 Then
        MsgBox "Couldn't find parent window"
    Else
        lStyle = GetWindowLong(hWindow, GWL_STYLE)
        'add WS_THICKFRAME
        lStyle = lStyle Or WS_THICKFRAME
        lStyle = SetWindowLong(hWindow, GWL_STYLE, lStyle)
        'makes the window redraw
        SetWindowPos hWindow, hParent, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    End If
    
End Sub

Sub fixWindow(hWindow As Long)

    Dim lStyle As Long
    Dim hParent As Long
    
    hParent = GetParent(hWindow)
    
    Debug.Print "Fixing " & hWindow, hParent
    
    'check for parent
    If hParent = 0 Then
        MsgBox "Couldn't find parent window"
    Else
        lStyle = GetWindowLong(hWindow, GWL_STYLE)
        'remove WS_THICKFRAME
        lStyle = lStyle And Not WS_THICKFRAME
        lStyle = SetWindowLong(hWindow, GWL_STYLE, lStyle)
        'makes the window redraw
        SetWindowPos hWindow, hParent, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    End If
End Sub
