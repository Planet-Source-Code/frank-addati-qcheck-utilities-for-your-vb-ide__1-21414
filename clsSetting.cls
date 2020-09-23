VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
    Private mblnQPadOnTop As Boolean
    Private mblnQPadReturn As Boolean
    Private mlngQPadPH As Long
    Private mlngQPadPW As Long
    Private mlngQPadPT As Long
    Private mlngQPadPL As Long
    
Private Sub Class_Initialize()
    
    mblnQPadOnTop = GetSetting("AFE", "QCheck", "QPadOnTop", True)
    mblnQPadReturn = GetSetting("AFE", "QCheck", "QPadReturn", True)
    mlngQPadPH = GetSetting("AFE", "QCheck", "QPadPH", 3000)
    mlngQPadPW = GetSetting("AFE", "QCheck", "QPadPW", 3000)
    mlngQPadPT = GetSetting("AFE", "QCheck", "QPadPT", 2500)
    mlngQPadPL = GetSetting("AFE", "QCheck", "QPadPL", 2500)
    
End Sub

Private Sub Class_Terminate()

    SaveSetting "AFE", "QCheck", "QPadOnTop", mblnQPadOnTop
    SaveSetting "AFE", "QCheck", "QPadReturn", mblnQPadReturn
    SaveSetting "AFE", "QCheck", "QPadPH", mlngQPadPH
    SaveSetting "AFE", "QCheck", "QPadPW", mlngQPadPW
    SaveSetting "AFE", "QCheck", "QPadPT", mlngQPadPT
    SaveSetting "AFE", "QCheck", "QPadPL", mlngQPadPL
    
End Sub

Public Property Get QPadOnTop() As Boolean
  
    QPadOnTop = mblnQPadOnTop
      
End Property

Public Property Let QPadOnTop(ByVal vData As Boolean)
  
    mblnQPadOnTop = vData
    
  
End Property

Public Property Get QPadReturn() As Boolean
  
    QPadReturn = mblnQPadReturn
  
End Property

Public Property Let QPadReturn(ByVal vData As Boolean)
  
    mblnQPadReturn = vData
  
End Property

Public Property Get QPadPH() As Long
    
    QPadPH = mlngQPadPH
    
End Property

Public Property Let QPadPH(ByVal vData As Long)
    
    mlngQPadPH = vData
    
End Property

Public Property Get QPadPW() As Long
    
    QPadPW = mlngQPadPW
    
End Property

Public Property Let QPadPW(ByVal vData As Long)
    
    mlngQPadPW = vData
    
End Property

Public Property Get QPadPT() As Long
    
    QPadPT = mlngQPadPT
    
End Property

Public Property Let QPadPT(ByVal vData As Long)
    
    mlngQPadPT = vData
    
End Property

Public Property Get QPadPL() As Long
    
    QPadPL = mlngQPadPL
    
End Property

Public Property Let QPadPL(ByVal vData As Long)
    
    mlngQPadPL = vData
    
End Property
