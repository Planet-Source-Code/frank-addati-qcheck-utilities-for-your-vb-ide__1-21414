Attribute VB_Name = "basMain"
Option Explicit

    '-------------------------
    'Used for Shell_NotifyIcon
    '-------------------------
    'User defined type required by Shell_NotifyIcon API call
    Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
    End Type
    Public nid As NOTIFYICONDATA
    
    '-------------------------
    'Used for Shell_NotifyIcon
    '-------------------------
    'Constants required by Shell_NotifyIcon API call:
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDOWN = &H201     'Button down
    Public Const WM_LBUTTONUP = &H202       'Button up
    Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
    Public Const WM_RBUTTONDOWN = &H204     'Button down
    Public Const WM_RBUTTONUP = &H205       'Button up
    Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
    
    '-------------------------
    'Used for Shell_NotifyIcon
    '-------------------------
    Public Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
    Public Declare Function Shell_NotifyIcon Lib "shell32" _
        Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
        
    '---------------------
    'Used for SetWindowPos
    '---------------------
    Public glngRetWinPos As Long
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = 1
    Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

    '---------------------
    'Used for SetWindowPos
    '---------------------
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal h&, ByVal hb&, ByVal X&, ByVal Y& _
        , ByVal cx&, ByVal cy&, ByVal f&) As Integer
    
    '-------------
    'Used for Word
    '-------------
    Public mobjWord97 As Word.Application
    Public gblnIsMainRunning As Boolean
    Public gobjSetting As clsSetting

    '-----------------------------
    'Used for GetPrevWindowForMain
    '-----------------------------
    Public gstrPrevAppForMain As String
    Private mblnIsMain As Boolean
    
    '-----------------------------
    'Used for GetPrevWindowForQPad
    '-----------------------------
    Public gstrPrevAppForQPad As String
    Private mblnIsQpad As Boolean
    
    '------------------------------------------------------
    'Used for GetPrevWindowForMain and GetPrevWindowForQPad
    '------------------------------------------------------
    Private Declare Function EnumWindows Lib "user32" _
        (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" _
        (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
        
Private Sub Main()
    
    Const cstrProc As String = "Main"
    On Error GoTo ErrHandler

    'Only one instance allowed in the system tray
    If Not App.PrevInstance Then
        Set gobjSetting = New clsSetting
        frmWait.Show vbModal
        If frmWait.blnWord97Loaded Then Load frmMain
    End If
    
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError("basMain", cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Public Sub GetPrevWindowForMain(frm As Form)
                 
    Call EnumWindows(AddressOf EnumWindowsProcForMain, frm.hwnd)
'    Debug.Print "---"
    
End Sub

Private Function EnumWindowsProcForMain _
    (ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    Static WindowText As String
    Static nRet As Long
    
    If IsWindowVisible(hwnd) Then
        
        WindowText = Space$(256)
        nRet = GetWindowText(hwnd, WindowText, Len(WindowText))
        
        If nRet Then
            WindowText = Left$(WindowText, nRet)
'            Debug.Print WindowText
            
            If WindowText = App.Title Then
                mblnIsMain = True
            End If
            
            If WindowText <> App.Title And mblnIsMain = True Then
                gstrPrevAppForMain = WindowText
                mblnIsMain = False
                
            End If
        End If
    End If
    
    EnumWindowsProcForMain = True
    
End Function

Public Sub GetPrevWindowForQPad(frm As Form)
                 
    Call EnumWindows(AddressOf EnumWindowsProcForQPad, frm.hwnd)
'    Debug.Print "---"
    
End Sub

Private Function EnumWindowsProcForQPad _
    (ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    Static WindowText As String
    Static nRet As Long
    
    If IsWindowVisible(hwnd) Then
        
        WindowText = Space$(256)
        nRet = GetWindowText(hwnd, WindowText, Len(WindowText))
        
        If nRet Then
            WindowText = Left$(WindowText, nRet)
'            Debug.Print WindowText
            
            If WindowText = frmQPad.Caption Then
                mblnIsQpad = True
            End If
            
            If WindowText <> frmQPad.Caption _
                And WindowText <> App.Title _
                And mblnIsQpad = True Then
                    gstrPrevAppForQPad = WindowText
                    mblnIsQpad = False
            
            End If
        End If
    End If
    
    EnumWindowsProcForQPad = True
    
End Function

Public Sub LogError(rstrMod As String, rstrProc As String _
    , rintErrNum As Integer, rstrErrDesc As String)
' Used for Global Error Checking.

    Dim X As Integer
    On Error Resume Next
    
    X = FreeFile
    Open App.Path & "\QCheck.Err" For Append As #X
    Write #X, rstrMod, rstrProc, rintErrNum, rstrErrDesc, Now, vbCrLf
    Close #X
    MsgBox "This application has encountered a problem:" & vbCrLf & vbCrLf _
        & "Message: " & rstrErrDesc & " [" & rintErrNum & "]" & vbCrLf & vbCrLf _
        & "Module: " & rstrMod & " [" & rstrProc & "]" & vbCrLf & vbCrLf _
        & "Suggestion: If this problem persists, report the message information " _
        & "to Technical Support.", vbCritical
  
End Sub
