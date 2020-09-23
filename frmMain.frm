VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About QCheck"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      Height          =   3555
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":0442
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   5400
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmMain.frx":0458
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 1999 Frank Addati  Freeware Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   810
      TabIndex        =   3
      Top             =   390
      Width           =   3510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "QuickCheck For Word 97/2K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   870
      TabIndex        =   2
      Top             =   90
      Width           =   3390
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Picture         =   "frmMain.frx":0570
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuPopSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopSpelling 
         Caption         =   "&Spelling"
      End
      Begin VB.Menu mnuPopSpellGramm 
         Caption         =   "+&Grammar"
      End
      Begin VB.Menu mnuPopThesaurus 
         Caption         =   "&Thesaurus"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopQPad 
         Caption         =   "&Qpad..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopLineNumberAdd 
         Caption         =   "A&dd Line Numbers"
      End
      Begin VB.Menu mnuPopLineNumberRemove 
         Caption         =   "Re&move Line Numbers"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuPopExit 
         Caption         =   "&Remove from Tray"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private mstrMode As String
    
Private Sub Form_Load()

    Const cstrProc As String = "Form_Load"
    On Error GoTo ErrHandler
        
    'The form must be fully visible before calling Shell_NotifyIcon
    Me.WindowState = vbMinimized
    Me.Refresh
    'Setting the System Tray
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "QCheck" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    'Setting the timer
    With Timer1
        .Enabled = False
        .Interval = 10
    End With
    
    txtHelp.Text = "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf
       
    txtHelp.Text = txtHelp.Text & "[ DESCRIPTION ]" & vbCrLf _
       & "This utility provides Spelling, Grammar and Thesaurus capabilities to most " _
       & "Windows applications. Simply highlight a word or " _
       & "sentence then right-click the QCHECK icon from the System Tray and select one " _
       & "of these three options:" & vbCrLf _
       & "1. Spelling (also accessible by left-clicking the icon)" & vbCrLf _
       & "2. +Grammar" & vbCrLf _
       & "3. Thesaurus" & vbCrLf & vbCrLf _
       & "A simple text editor can be found by clicking Qpad from the popup menu. " _
       & "From here you can check spelling and grammar without highlighting the text. " _
       & "In addition by clicking the 'OK' menu option the corrected text will be " _
       & "sent to the application the editor was hooked onto " _
       & "(Do this by clicking the 'Set' menu option)" & vbCrLf & vbCrLf & vbCrLf
    
    txtHelp.Text = txtHelp.Text & "[ WHAT YOU NEED TO RUN IT ]" & vbCrLf _
       & "1. Visual Basic 6 runtime (Msvbvm60.dll)" & vbCrLf _
       & "2. Microsoft Word 97 or Word 2000" _
       & vbCrLf & vbCrLf & vbCrLf
    
    txtHelp.Text = txtHelp.Text & "[ ABOUT QCHECK ]" & vbCrLf _
       & "I created QCHECK primarily for myself. Code editors in programming " _
       & "languages like Visual Basic, Access, SQL Server etc., " _
       & "don't provide a spell checker, so sometimes I was forced to " _
       & "cut and paste text back and forward from Microsoft Word. " _
       & "By using QCHECK now I can access the power of Word quickly and painlessly." _
       & vbCrLf & vbCrLf _
       & "This program is Freeware. There are no special limitations " _
       & "in the program's functionality." & vbCrLf & vbCrLf _
       & "Send comments, suggestions, bugs, ... to:" & vbCrLf _
       & "frankaddati@21century.com.au" & vbCrLf _
       & "http://www.21century.com.au/frankaddati" & vbCrLf & vbCrLf & vbCrLf
    
    txtHelp.Text = txtHelp.Text & "[ DISCLAIMER OF WARRANTY ]" & vbCrLf _
       & "THIS SOFTWARE IS PROVIDED <AS IS> AND WITHOUT WARRANTIES AS TO PERFORMANCE OR " _
       & "MERCHANTABILITY OR ANY OTHER WARRANTIES WHETHER EXPRESSED OR IMPLIED.  " _
       & "BECAUSE OF THE VARIOUS HARDWARE AND SOFTWARE ENVIRONMENTS INTO WHICH QCHECK " _
       & "MAY BE PUT, NO WARRANTY OF FITNESS FOR A PARTICULAR PURPOSE IS OFFERED.  " _
       & "THE AUTHOR ASSUMES NO LIABILITY FOR DAMAGES, DIRECT OR CONSEQUENTIAL, " _
       & "WHICH MAY RESULT FROM THE USE OF QCHECK."
       
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
    
    Static bRunning As Boolean
    Dim Result As Long
    Dim msg As Long
    Const cstrProc As String = "Form_MouseMove"
    On Error GoTo ErrHandler
    
    If gblnIsMainRunning = True Then Exit Sub
    
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Result = SetForegroundWindow(Me.hwnd)
            mstrMode = "SPELLING"
            Timer1.Enabled = True
            
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mnuPopSys
    End Select
            
ExitHere:
    Exit Sub

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Sub

Private Sub Form_Resize()

    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this removes the icon from the system tray

    Shell_NotifyIcon NIM_DELETE, nid
    
End Sub

Private Sub Timer1_Timer()
    
    Dim strTemp As String
    Dim frmOpen As Form
    Const cstrProc As String = "Timer1_Timer"
    On Error GoTo ErrHandler

    Timer1.Enabled = False
    
    Select Case mstrMode
        Case "SPELLING", "GRAMMAR"
            gblnIsMainRunning = True
            With Clipboard
                .Clear
                AutoCopy
                strTemp = Word97Do(mstrMode, .GetText)
                .Clear
                .SetText strTemp
                AutoPaste
            End With
            gblnIsMainRunning = False
            
        Case "THESAURUS"
            gblnIsMainRunning = True
            With Clipboard
                AutoCopy
                strTemp = Word97Do(mstrMode, .GetText)
                .Clear
                .SetText strTemp
                AutoPaste
            End With
            gblnIsMainRunning = False
            
        Case "QPAD"
            frmQPad.Show
            
        Case "LINENUMBER_ADD"
            gblnIsMainRunning = True
            With Clipboard
                AutoCopy
                strTemp = LineNumberAdd(.GetText)
                .Clear
                .SetText strTemp
                AutoPaste
            End With
            gblnIsMainRunning = False
            
        Case "LINENUMBER_REMOVE"
            gblnIsMainRunning = True
            With Clipboard
                AutoCopy
                strTemp = LineNumberRemove(.GetText)
                .Clear
                .SetText strTemp
                AutoPaste
            End With
            gblnIsMainRunning = False
                
        Case "UNLOAD"
            If mobjWord97.Documents.Count = 0 Then mobjWord97.Quit
            Set mobjWord97 = Nothing
            For Each frmOpen In Forms
                Unload frmOpen
            Next
    End Select

ExitHere:
    Exit Sub

ErrHandler:
    If Err = 462 Then 'When Word is already closed by another application
        Resume Next
    Else
        Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
        Resume ExitHere
    End If
         
End Sub

Private Sub cmdOk_Click()
    
    Me.WindowState = vbMinimized
    gblnIsMainRunning = False
    
End Sub

Private Sub mnuPopSpelling_Click()

    mstrMode = "SPELLING"
    Timer1.Enabled = True

End Sub

Private Sub mnuPopSpellGramm_Click()
    
    mstrMode = "GRAMMAR"
    Timer1.Enabled = True

End Sub

Private Sub mnuPopThesaurus_Click()
    
    mstrMode = "THESAURUS"
    Timer1.Enabled = True
    
End Sub

Private Sub mnuPopQPad_Click()

    mstrMode = "QPAD"
    Timer1.Enabled = True

End Sub

Private Sub mnuPopLineNumberAdd_Click()

    mstrMode = "LINENUMBER_ADD"
    Timer1.Enabled = True

End Sub

Private Sub mnuPopLineNumberRemove_Click()

    mstrMode = "LINENUMBER_REMOVE"
    Timer1.Enabled = True

End Sub

Private Sub mnuPopAbout_Click()

    Dim Result As Long

    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    
    gblnIsMainRunning = True
    
End Sub

Private Sub mnuPopExit_Click()
    
    mstrMode = "UNLOAD"
    Timer1.Enabled = True
    
End Sub

Private Function TextFixed(strText As String) As String
    
    Dim lngPos As Long
    Const cstrProc As String = "TextFixed"
    On Error GoTo ErrHandler
    
    'Copy the selection and remove the last carriage return added by Word
    strText = Left(strText, Len(strText) - 1)
    
    'Parse the whole string and replace a single CR with CRLF
    lngPos = 1
    
    Do While lngPos >= 0
        lngPos = InStr(lngPos, strText, Chr(13))
        
        If lngPos > 0 Then
            strText = Left(strText, lngPos - 1) & Chr(13) + Chr(10) _
                    & Mid(strText, lngPos + 1)
            lngPos = lngPos + 2
            
        Else
            Exit Do
        End If
    Loop
    
    TextFixed = strText
    
ExitHere:
    Exit Function

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Function

Private Sub AutoCopy()

    Const cstrProc As String = "AutoCopy"
    On Error GoTo ErrHandler
    
    'Find previous window and activate it
    GetPrevWindowForMain Me
    AppActivate gstrPrevAppForMain
    'Copy the selected text
    DoEvents
    SendKeys "^c", True
    'Bring the focus back to QCheck
    AppActivate App.Title
    
ExitHere:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 5
            Resume ExitHere
        Case Else
            Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
            Resume ExitHere
    End Select
    
End Sub

Private Sub AutoPaste()

    Const cstrProc As String = "AutoPaste"
    On Error GoTo ErrHandler
    
    'Activate the window, which we have stored in the global variable
    'and paste the text from the clipboard
    AppActivate gstrPrevAppForMain
    DoEvents
    SendKeys "^v", True
    
ExitHere:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 5
            Resume ExitHere
        Case Else
            Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
            Resume ExitHere
    End Select
        
End Sub

Private Function LineNumberAdd(strInfectedCode As String) As String
    
    Dim lngLineNumber As Long
    Dim lngKeywordPos As Long
    Dim strParagraph As String
    Dim lngReturnPos As Long
    Dim intLineNumber As Integer
    Dim strParagraphClean As String
    Dim strCleaned As String
    Dim blnContinuationLine As Boolean
    Const cstrProc As String = "LineNumberAdd"
    On Error GoTo ErrHandler
    
    lngReturnPos = 1
    
    Do While lngReturnPos > 0
        lngReturnPos = InStr(strInfectedCode, vbCrLf)
        strParagraph = Mid(strInfectedCode, 1, lngReturnPos + 1)
        strInfectedCode = Mid(strInfectedCode, lngReturnPos + 2)
        
        If Not blnContinuationLine Then
        
            If InStr(strParagraph, "Function ") > 0 _
            Or InStr(strParagraph, "Sub ") > 0 _
            Or InStr(strParagraph, "Event ") > 0 _
            Or InStr(strParagraph, " Get ") > 0 _
            Or InStr(strParagraph, " Let ") > 0 _
            Or Trim(strParagraph) = vbCrLf _
            Or Trim(strParagraph) = "" _
            Or IsNumeric(Left(strParagraph, 1)) _
            Or Left(Trim(strParagraph), 1) = "'" Then
                strParagraphClean = strParagraph
            Else
                lngLineNumber = lngLineNumber + 10
                strParagraphClean = lngLineNumber & ":" & strParagraph
            End If
            
        Else
            strParagraphClean = strParagraph
        End If
        
        strCleaned = strCleaned & strParagraphClean
        
        If Right(strParagraphClean, 3) = "_" & vbCrLf Then
            blnContinuationLine = True
        Else
            blnContinuationLine = False
        End If
    Loop
    
ExitHere:
    LineNumberAdd = strCleaned
    Exit Function

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Function

Private Function LineNumberRemove(strInfectedCode As String) As String

    Dim strParagraph As String
    Dim lngReturnPos As Long
    Dim intLineNumber As Integer
    Dim strParagraphClean As String
    Dim strCleaned As String
    Dim blnContinuationLine As Boolean
    Const cstrProc As String = "LineNumberRemove"
    On Error GoTo ErrHandler
    
    lngReturnPos = 1
    
    Do While lngReturnPos > 0
        lngReturnPos = InStr(strInfectedCode, vbCrLf)
        strParagraph = Mid(strInfectedCode, 1, lngReturnPos + 1)
        strInfectedCode = Mid(strInfectedCode, lngReturnPos + 2)
        
        If Not blnContinuationLine Then
        
            If IsNumeric(Left(strParagraph, 1)) Then
                intLineNumber = Val(strParagraph)
                strParagraphClean = Replace(strParagraph, intLineNumber, "", , 1)
                
                If Left(strParagraphClean, 1) = ":" Then
                    strParagraphClean = Replace(strParagraphClean, ":", "", , 1)
                End If
                
            Else
                strParagraphClean = strParagraph
            End If
            
        Else
            strParagraphClean = strParagraph
        End If
        
        strCleaned = strCleaned & strParagraphClean
        
        If Right(strParagraphClean, 3) = "_" & vbCrLf Then
            blnContinuationLine = True
        Else
            blnContinuationLine = False
        End If
    Loop
    
ExitHere:
    LineNumberRemove = strCleaned
    Exit Function

ErrHandler:
    Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
    Resume ExitHere
    
End Function

Public Function Word97Do(strMode As String, strText As String) As String
    
    Dim objdoc As Word.Document
    Const cstrProc As String = "Word97Do"
    On Error GoTo ErrHandler
        
    'Create a dummy document and paste the text from the clipboard
    Set objdoc = mobjWord97.Documents.Add
    With objdoc
        .Range.InsertBefore (strText)
        'Must activate Word in order to have the Spelling dialog box visible
        'Also make sure is minimised
        With mobjWord97
            .WindowState = wdWindowStateMinimize
            .Visible = True
            .Activate
        End With
        
        'Run the appropriate feature
        Select Case strMode
        
            Case "SPELLING", "GRAMMAR" 'Run Spelling
                If .SpellingErrors.Count > 0 Then
                    .CheckSpelling
                Else
                    mobjWord97.Visible = False
                    MsgBox "No spelling errors found!", vbInformation ' _
                        ', "QCheck - [" & gstrPrevAppForMain & "]"
                End If
                'Run Grammar
                If strMode = "GRAMMAR" Then
                    If .GrammaticalErrors.Count > 0 Then
                        .CheckGrammar
                    Else
                        MsgBox "No grammar errors found!", vbInformation '_
                           ' , "QCheck - [" & gstrPrevAppForMain & "]"
                    End If
                End If
                
            Case "THESAURUS" 'Run Thesaurus
                .Range.CheckSynonyms
            
        End Select
        
        'Compensate Word interpretation of CRLF combination
        Word97Do = (TextFixed(.Range.Text))
        'Close the document but keep Word in memory(better performance)
        .Close SaveChanges:=wdDoNotSaveChanges
        If mobjWord97.Documents.Count = 0 Then
            mobjWord97.Visible = False
        Else
            mobjWord97.Visible = True
            mobjWord97.WindowState = wdWindowStateNormal
        End If
        
    End With

ExitHere:
    Exit Function

ErrHandler:
    'Could not open macro storage.
    'this happens when trying to add a doc to Word.
    'Until the user fixes this problem (due to insufficient memory, space or
    'faulty normal.dot template)we must quit, otherwise the system will create
    'one million Word applications in an infinite loop.
    If Err = 5981 Then
        Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
        End
        
    'This normally happens when the user open a Word document from a shortcut
    '(automatically uses our instance of Word) and later he closes Word app.
    'Err numbers fro NT and 95
    ElseIf Err = -2147023174 Or 462 Then
        frmWait.Show vbModal
        If frmWait.blnWord97Loaded Then Resume
        
    Else
        Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
        Resume ExitHere
    End If
    
End Function
