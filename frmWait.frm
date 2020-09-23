VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   0  'None
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "QCheck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   555
      TabIndex        =   1
      Top             =   150
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "© 1999 Frank Addati"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   540
      Width           =   1545
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmWait.frx":000C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loading Microsoft Word Please Wait..."
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1950
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Public blnWord97Loaded As Boolean

Private Sub Form_Activate()
    
    Dim intRet As Integer
    Dim lngLoop As Long
    Dim blnPause As Boolean
    Const cstrProc As String = "Form_Activate"
    On Error GoTo ErrHandler
    
    blnWord97Loaded = False
    Set mobjWord97 = New Word.Application
    mobjWord97.WindowState = wdWindowStateMinimize
    mobjWord97.Visible = False
    blnWord97Loaded = True
    
ExitHere:
    Unload Me
    Exit Sub

ErrHandler:
    Select Case Err
        Case -2147023179, -2147023174, 462 'No enough time to load Word
            Set mobjWord97 = Nothing
            'Pause for a while then try again
            If Not blnPause Then
                For lngLoop = 0 To 50000
                Next
                blnPause = True
                Resume
            'After second tentative give up and let the user to do it
            Else
                intRet = MsgBox("Your System is busy at the moment and " _
                    & "Microsoft Word couldn't be loaded." _
                    & vbCrLf & vbCrLf _
                    & "Would you like to try again?", vbYesNo + vbExclamation)
                
                If intRet = vbYes Then
                    Resume
                
                Else
                    blnWord97Loaded = False
                    Resume ExitHere
                End If
            End If
            
        Case 429 'Word97 doesn't exist
            MsgBox "Microsoft Word 97 is not installed on this computer." _
                , vbCritical
                blnWord97Loaded = False
                Resume ExitHere
                
        Case Else
            Call LogError(Me.Name, cstrProc, Err.Number, Err.Description)
            Resume ExitHere
    End Select
    
End Sub
