VERSION 5.00
Begin VB.UserControl WB 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "WB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim webView As Long

Sub StopLoading()
    If IsInit Then wkeStopLoading webView
End Sub

Sub GoForward()
    If IsInit Then wkeGoForward webView
End Sub

Sub EditorSelectAll()
    If IsInit Then wkeEditorSelectAll webView
End Sub

Sub Goback()
    If IsInit Then wkeGoBack webView
End Sub

Public Property Get Title() As String
    If IsInit Then
        Title = Utf8ToString(wkeGetTitle(webView))
    Else
        Title = ""
    End If
End Property

Public Property Get Url() As String
    If IsInit Then
        Url = Utf8ToString(wkeGetUrl(webView))
    Else
        Url = ""
    End If
End Property


Public Property Get IsInit() As Boolean
    IsInit = (webView <> 0)
End Property

Private Sub UserControl_Initialize()
    If Not IsInit Then wkeInitialize
End Sub

Private Sub UserControl_Show()
    If App.LogMode = 0 Then Exit Sub
    If webView = 0 Then
        wkeInitialize
        'webView = wkeCreateWebView()
        webView = wkeCreateWebWindow(WKE_WINDOW_TYPE_CONTROL, UserControl.HWND, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
        'lsMb.Add webView, Me
        
        '此函数触发在非UI线程，编译出来的程序到现在的时候会崩溃，只能将这个函数弃用
        'mbOnDownloadInBlinkThread webView, AddressOf mbDownloadInBlinkThreadCallback
        wkeOnDownload2 webView, AddressOf wkeDownload2Callback
        wkeOnWindowDestroy webView, AddressOf wkeWindowDestroyCallback
        wkeOnWindowClosing webView, AddressOf wkeWindowClosingCallback
    
    End If
    ShowWindow True
    LoadURL "https://yasuo.360.cn/"
End Sub

Sub LoadURL(ByVal Url As String)
    wkeLoadURL webView, Url
End Sub

Private Sub UserControl_Resize()
    If Not IsInit Then Exit Sub
    Resize UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub


Private Sub ShowWindow(ByVal show As Boolean)
    wkeShowWindow webView, show
End Sub

Sub Resize(ByVal width As Long, ByVal height As Long)
    If Not IsInit Then Exit Sub
    wkeResize webView, width, height
End Sub

Sub C()
    'wkeShowWindow webView, False
    'wkeSetHandle webView, 0
    
    'webView = 0
        
       
    '获取当前页面的cookie
    'AppendLog Utf8ToString(wkeGetCookie(webView))
    
    
    'Debug.Print StringToUtf8(r)
End Sub

Private Sub UserControl_Terminate()
    wkeDestroyWebWindow webView
    wkeShutdown
    webView = 0
End Sub
