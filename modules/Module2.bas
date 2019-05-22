Attribute VB_Name = "Module2"
Option Compare Database
Option Explicit

' Default duration in milliseconds
Private Const cElapse As Long = 10000

Private Declare Function GetActiveWindow Lib "user32" () _
    As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () _
    As Long

Private Declare Function GetWindowLongA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) _
    As Long

Private Declare Function GetWindowTextA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) _
    As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) _
    As Long

Private Declare Function MessageBoxA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal wType As Long) _
    As Long

Private Declare Function PostMessageA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) _
    As Long

Private Declare Function SetFocus Lib "user32" ( _
    ByVal hwnd As Long) _
    As Long

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) _
    As Long

Private Declare Function SetWindowsHookExA Lib "user32" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hMod As Long, _
    ByVal dwThreadId As Long) _
    As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long) _
    As Long

Private Const GWL_HINSTANCE = (-6)
Private Const HCBT_ACTIVATE = 5
Private Const HCBT_SETFOCUS = 9
Private Const NV_CLOSEMSGBOX As Long = &H5000&
Private Const WH_CBT = 5
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_TIMER = &H113

Private hDlgMsgBox As Long
Private hHook As Long
Private hTimerID As Long
Private hWndApp As Long
Private hWndMsgBox As Long


Private Sub TimerClear()

    If hWndApp <> 0 Then
        KillTimer hWndApp, _
            NV_CLOSEMSGBOX
        hDlgMsgBox = 0
        hHook = 0
        hTimerID = 0
        hWndApp = 0
        hWndMsgBox = 0
    End If

End Sub


Public Function MsgBoxHookProc( _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) _
    As Long

    Select Case uMsg
        Case HCBT_SETFOCUS
            hDlgMsgBox = wParam
        Case HCBT_ACTIVATE
            hWndMsgBox = wParam
            UnhookWindowsHookEx hHook
    End Select

End Function


Public Function TimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long) _
    As Long

    Select Case uMsg
        Case WM_TIMER
            If idEvent = NV_CLOSEMSGBOX Then
                If hWndMsgBox <> 0 Then
                    If hDlgMsgBox <> 0 Then
                        SetFocus hDlgMsgBox
                        DoEvents
                        PostMessageA hDlgMsgBox, _
                            WM_LBUTTONDOWN, _
                            0, _
                            ByVal 0&
                        PostMessageA hDlgMsgBox, _
                            WM_LBUTTONUP, _
                            0, _
                            ByVal 0&
                    End If
                    TimerClear
                End If
            End If
        Case Else
    End Select

End Function


Public Function TMsgBox( _
    ByVal Prompt, _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal Title, _
    Optional ByVal Elapse) _
    As VbMsgBoxResult

' For Access 97, use:
' Public Function TMsgBox( _
      ByVal Prompt, _
      Optional ByVal Buttons As Long, _
      Optional ByVal Title, _
      Optional ByVal Elapse) As Long

    Dim hMod As Long
    Dim hThreadId As Long
    Dim lTitle As Long
    Dim sTitle As String

    hWndApp = GetActiveWindow

    hMod = GetWindowLongA(hWndApp, _
        GWL_HINSTANCE)
    hThreadId = GetCurrentThreadId()

    If IsMissing(Title) = True Then
        sTitle = String(255, 0)
        lTitle = GetWindowTextA(Application.hWndAccessApp, _
                                sTitle, 255)
        If lTitle > 0 Then sTitle = Left(sTitle, lTitle)
    Else
        sTitle = Title
    End If

    If IsMissing(Elapse) = True Then
        Elapse = cElapse
    Else
        Elapse = CLng(Elapse)
    End If

    ' For Access 2000/2002/2003/2007
    hHook = SetWindowsHookExA(WH_CBT, _
        AddressOf MsgBoxHookProc, _
        hMod, hThreadId)
    If Elapse > 0 Then
        hTimerID = SetTimer(hWndApp, _
            NV_CLOSEMSGBOX, _
            Elapse, _
            AddressOf TimerProc)
    End If
    ' For Access 97
    '   You will need to download the code for the AddrOf
    '   function from Trigeminal Software at:
    '   http://www.trigeminal.com/lang/1033/codes.asp?ItemID=19#19
    '   Comment out the code under Access 2000/2002/2003/2007
    '   and Uncomment the following lines:
    'hHook = SetWindowsHookExA(WH_CBT, _
        AddrOf("MsgBoxHookProc"), _
        hMod, hThreadId)
    'If Elapse > 0 Then
    '    hTimerID = SetTimer(hWndApp, _
            NV_CLOSEMSGBOX, _
            Elapse, _
            AddrOf("TimerProc"))
    'End If

    TMsgBox = MessageBoxA(hWndApp, _
        Prompt, _
        sTitle, _
        Buttons)

    TimerClear

End Function

