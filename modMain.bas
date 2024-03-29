Attribute VB_Name = "modMain"
'Hi, welcome to THE code, to a code you surely looked for - REAL C++ CONTROLS!!!
'The special thing on this code is that you only need the classname of the
'control you want to create. You can get it using the Microsoft Spy++.
'Now have fun with this code and if you like it, PLEASE VOTE FOR ME!!!

'This is the Module that does all the stuff,
'the Form is only a "container" for the controls

'*****************************
'* API Declarations          *
'*****************************
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

'*****************************
'* Type Declarations         *
'*****************************
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'*****************************
'* Consts                    *
'*****************************
Const ES_MULTILINE = &H4&
Const WS_CHILD = &H40000000
Const WS_EX_CLIENTEDGE = &H200&
Const WM_DESTROY = &H2
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const GWL_WNDPROC = (-4)
Const SW_SHOWNORMAL = 1
Const MB_OK = &H0&
Const MB_ICONEXCLAMATION = &H30&
Const OptionStyle = 1442840585
Const CheckStyle = 1442906115

'*****************************
'* Vars                      *
'*****************************
Public ControlHwnd As Long 'The hWnd of the created Control
Public gButOldProc As Long 'To save the old WindowProc for the button
Public ButtonHwnd As Long 'The hWnd of the Button (for the Procedures)
Public TextBoxHwnd As Long 'The hWnd of the TextBox (to get the text)

'The sub that creates a control
Private Function CreateControl(cClassName As String, cWindowText As String, XPos As Long, YPos As Long, cWidth As Long, cHeight As Long, Optional cStyle As Long) As Long
    'If the control is a TextBox there are special settings
    If cClassName = "Edit" Then
         ControlHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, cClassName, cWindowText, WS_CHILD Or ES_MULTILINE, XPos, YPos, cWidth, cHeight, frmControls.hwnd, 0&, App.hInstance, 0&)
    'Else there are no special settings
    Else:
         ControlHwnd& = CreateWindowEx(0&, cClassName, cWindowText, WS_CHILD Or cStyle, XPos, YPos, cWidth, cHeight, frmControls.hwnd, 0&, App.hInstance, 0&)
    End If
    'Show the Control because it´s invisible
    Call ShowWindow(ControlHwnd&, SW_SHOWNORMAL)
    CreateControl = ControlHwnd&
End Function

'This is needed that you can use Control Procedures
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim strTemp As String
    Select Case uMsg&
        Case WM_DESTROY:
        'Call the WM_QUIT message
        Call PostQuitMessage(0&)
    End Select
    'Call the standard window procedure
    WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)
End Function

'To get the adress of a procedure using "AddressOf"
Public Function GetAddress(ByVal lngAddr As Long) As Long
   GetAddress = lngAddr&
End Function

Public Sub Main()
    'Create the controls you want using CreateControl - and have fun!!!
    TextBoxHwnd& = CreateControl("Edit", "This is a REAL C++ TextBox supporting multiline", 10, 10, 200, 70)
    Call CreateControl("Static", "This is a C++ Label, to see that the controls are C++ controls use Spy++, look at the classname and compare it with the classname of controls in Windows programs", 220, 10, 300, 80)
    Call CreateControl("Button", "C++ Button - Resizable!", 10, 100, 200, 50, 4020000)
    ButtonHwnd& = CreateControl("Button", "Another Button - click it!", 230, 100, 200, 50)
    'Get the address of the standard button procedure and save it in "gButOldProc"
    gButOldProc& = GetWindowLong(ButtonHwnd&, GWL_WNDPROC)
    'Use GWL_WNDPROC to save the adress of the procedure for the button
    'You have to do this for every control you want to have a procedure
    Call SetWindowLong(ButtonHwnd&, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc))
    Call CreateControl("msctls_hotkey32", "Hotkey control", 10, 198, 100, 25)
    Call CreateControl("Static", "<", 115, 202, 20, 20)
    Call CreateControl("Static", "This is the cool hotkey control. Select it, and press any hotkeys you want.", 130, 195, 250, 50)
    'CheckBoxes and OptionButtons are also Buttons.
    'Use the constants CheckStyle and OptionStyle as style to get
    'Checks and Options.
    'To get the styles of windows controls, use Spy++ and convert the
    'the number Spy++ gives you to decimal using this function:
    'Val("&H" & Style)
    Call CreateControl("Button", "Option1", 10, 160, 100, 30, OptionStyle)
    Call CreateControl("Button", "Option2", 110, 160, 100, 30, OptionStyle)
    Call CreateControl("Button", "Option3", 210, 160, 100, 30, OptionStyle)
    Call CreateControl("Button", "Check", 310, 160, 100, 30, CheckStyle)
    frmControls.Show
End Sub

'This is the procedure that is called when you click the button
Public Function ButtonWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
    Case WM_LBUTTONUP:
        'Left button is up (user clicked the Button)
        'Use "WM_LBUTTONDOWN"
        TextBoxText = Get_Text_Of_Control(TextBoxHwnd)
        MsgText = "The text of the TextBox control:" & vbCrLf & Trim(TextBoxText)
        'Use the MessageBox API to show a message box
        'with the text of the TextBox
        Call MessageBox(gHwnd&, MsgText, App.Title, MB_OK Or MB_ICONEXCLAMATION)
    End Select
    'Call the standard window proc
    ButtonWndProc = CallWindowProc(gButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
End Function

'Use this function to get the actual text of created controls
Function Get_Text_Of_Control(ByVal cHwnd As Long) As String
    Dim ControlText As String
    ControlText = Space(254)
    'Use the GetWindowText API to get the actual text of the control
    GetWindowText cHwnd, ControlText, 254
    Get_Text_Of_Control = Trim(ControlText)
End Function
