VERSION 5.00
Begin VB.Form frmQPad 
   Caption         =   "Qpad - [No Window Selected]"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4695
   Icon            =   "frmQPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4695
   Begin VB.TextBox txtPad 
      Height          =   1755
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuSpelling 
      Caption         =   "&Spelling"
   End
   Begin VB.Menu mnuSpellGramm 
      Caption         =   "+&Grammar"
   End
   Begin VB.Menu mnuThesaurus 
      Caption         =   "&Thesaurus"
   End
   Begin VB.Menu mnuSet 
      Caption         =   "S&et"
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptOnTop 
         Caption         =   "&Always on Top"
      End
      Begin VB.Menu mnuOptReturn 
         Caption         =   "Sen&d on 'Enter'"
      End
   End
   Begin VB.Menu mnuOk 
      Caption         =   "O&k"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmQPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private mstrAppNameSet As String
    
Private Sub Form_Activate()
'The form_Activate event will only run when the user clicks
'QPad on the main menu (It will not run when the form is already opened).

    Const cstrProc As String = "Form_Activate"
    On Error GoTo ErrHandler
    
    Me.mnuOptReturn.Checked = gobjSetting.QPadReturn
    Me.mnuOptOnTop.Checked = gobjSetting.QPadOnTop
    CheckQpadOnTop
    
    'Set to the first previous available window
    mnuSet_Click

ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere

End Sub

Private Sub Form_Load()

    Const cstrProc As String = "Form_Load"
    On Error GoTo ErrHandler
    
    With gobjSetting
        Me.Move .QPadPL, .QPadPT, .QPadPW, .QPadPH
    End With
    
    gblnIsMainRunning = True
    
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Const cstrProc As String = "Form_QueryUnload"
    On Error GoTo ErrHandler
    
    If Me.WindowState <> vbMinimized Then
        With gobjSetting
            .QPadOnTop = Me.mnuOptOnTop.Checked
            .QPadReturn = Me.mnuOptReturn.Checked
            .QPadPH = Me.Height
            .QPadPL = Me.Left
            .QPadPT = Me.Top
            .QPadPW = Me.Width
        End With
    End If
                
    gblnIsMainRunning = False

ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Private Sub Form_Resize()

    Const cstrProc As String = "Form_Resize"
    On Error GoTo ErrHandler
    
    If Me.WindowState <> vbMinimized Then
        'Form minimum size allowed
        If Me.Width < 4815 Then
            Me.Width = 4815
        End If
        If Me.Height < 950 Then
            Me.Height = 950
        End If
        'Adjust the textbox's size
        With txtPad
            .Width = Me.Width - 125
            .Height = Me.Height - 690
        End With
    End If
    
ExitHere:
    Exit Sub
    
ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Private Sub txtPad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And mnuOptReturn.Checked = True Then
        KeyAscii = 0
        mnuOk_Click
    End If

End Sub

Private Sub mnuSpelling_Click()
    
    RunOption "SPELLING"
    
End Sub

Private Sub mnuSpellGramm_Click()
    
    RunOption "GRAMMAR"
    
End Sub

Private Sub mnuThesaurus_Click()
    
    RunOption "THESAURUS"
        
End Sub

Private Sub mnuSet_Click()

    GetPrevWindowForQPad Me
    mstrAppNameSet = gstrPrevAppForQPad
    Me.Caption = "Qpad - [" & mstrAppNameSet & "]"
            
End Sub

Private Sub mnuOptOnTop_Click()
        
    Me.mnuOptOnTop.Checked = Not Me.mnuOptOnTop.Checked
    CheckQpadOnTop
    
End Sub

Private Sub mnuOptReturn_Click()

    Me.mnuOptReturn.Checked = Not Me.mnuOptReturn.Checked
    
End Sub

Private Sub mnuOk_Click()

    Const cstrProc As String = "mnuOk_Click"
    On Error GoTo ErrHandler
    
    With txtPad
        Clipboard.Clear
        Clipboard.SetText .Text
        
        AppActivate mstrAppNameSet
        DoEvents
        SendKeys "^v", True
        SendKeys "{ENTER}", True
        
        .Text = ""
        .SetFocus
    End With
        
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere

End Sub

Private Sub mnuExit_Click()

    Const cstrProc As String = "mnuExit_Click"
    On Error GoTo ErrHandler
    
    Unload Me
    
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere

End Sub

Private Sub RunOption(strOption As String)

    Dim strText As String
    Dim blnAllText As Boolean
    Const cstrProc As String = "RunOption"
    On Error GoTo ErrHandler
         
    If txtPad.SelText = "" Then
        strText = txtPad.Text
        blnAllText = True
    Else
        strText = txtPad.SelText
        blnAllText = False
    End If

    With txtPad
        strText = frmMain.Word97Do(strOption, strText)
        If blnAllText Then
            .Text = strText
            .SelStart = Len(strText)
        Else
            .SelText = strText
        End If
    End With
    
    SetFocus

ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere

End Sub

Private Sub CheckQpadOnTop()

    If Me.mnuOptOnTop.Checked = True Then
        glngRetWinPos = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        glngRetWinPos = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If

End Sub
