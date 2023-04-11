Attribute VB_Name = "wkeModule"
'typedef enum _wkeWindowType {
'    WKE_WINDOW_TYPE_POPUP,
'    WKE_WINDOW_TYPE_TRANSPARENT,
'    WKE_WINDOW_TYPE_CONTROL
'} wkeWindowType;
Enum wkeWindowType
    WKE_WINDOW_TYPE_POPUP
    WKE_WINDOW_TYPE_TRANSPARENT
    WKE_WINDOW_TYPE_CONTROL
End Enum

'typedef enum _wkeLoadingResult {
'    WKE_LOADING_SUCCEEDED,
'    WKE_LOADING_FAILED,
'    WKE_LOADING_CANCELED
'} wkeLoadingResult;
Enum wkeLoadingResult
    WKE_LOADING_SUCCEEDED
    WKE_LOADING_FAILED
    WKE_LOADING_CANCELED
End Enum

'typedef enum _wkeDownloadOpt {
'    kWkeDownloadOptCancel,
'    kWkeDownloadOptCacheData,
'} wkeDownloadOpt;
Enum wkeDownloadOpt
    kWkeDownloadOptCancel
    kWkeDownloadOptCacheData
End Enum

'typedef struct _wkeNetJobDataBind {
'    void* param;
'    wkeNetJobDataRecvCallback recvCallback;
'    wkeNetJobDataFinishCallback finishCallback;
'}wkeNetJobDataBind;

Type wkeNetJobDataBind
    param As Long
    recvCallback As Long
    finishCallback As Long
End Type

'typedef enum _wkeProxyType {
'    WKE_PROXY_NONE,
'    WKE_PROXY_HTTP,
'    WKE_PROXY_SOCKS4,
'    WKE_PROXY_SOCKS4A,
'    WKE_PROXY_SOCKS5,
'    WKE_PROXY_SOCKS5HOSTNAME
'} wkeProxyType;
Enum wkeProxyType
    WKE_PROXY_NONE
    WKE_PROXY_HTTP
    WKE_PROXY_SOCKS4
    WKE_PROXY_SOCKS4A
    WKE_PROXY_SOCKS5
    WKE_PROXY_SOCKS5HOSTNAME
End Enum

'typedef struct _wkeProxy {
'    wkeProxyType type;
'    char hostname[100];
'    unsigned short port;
'    char username[50];
'    char password[50];
'} wkeProxy;
Type wkeProxy
    type As wkeProxyType
    hostname As String
    port As Long
    username As String
    password As String
End Type

'WKE_EXTERN_C __declspec(dllexport) void WKE_CALL_TYPE wkeInitialize();
Public Declare Sub wkeInitialize Lib "node.dll" ()

'ITERATOR0(bool, wkeIsInitialize, "")
Public Declare Function wkeIsInitialize Lib "node.dll" () As Boolean

'ITERATOR0(void, wkeShutdown, "")
Public Declare Sub wkeShutdown Lib "node.dll" ()

'ITERATOR6(wkeWebView, wkeCreateWebWindow, wkeWindowType type, HWND parent, int x, int y, int width, int height, "")
Public Declare Function wkeCreateWebWindow _
               Lib "node.dll" (ByVal wkeWindowType As Long, _
                               ByVal parent As Long, _
                               ByVal x As Long, _
                               ByVal y As Long, _
                               ByVal width As Long, _
                               ByVal height As Long) As Long

'ITERATOR0(wkeWebView, wkeCreateWebView, "")
Public Declare Function wkeCreateWebView Lib "node.dll" () As Long

'ITERATOR2(void, wkeSetHandle, wkeWebView webView, HWND wnd, "")
Public Declare Sub wkeSetHandle Lib "node.dll" (ByVal webView As Long, wnd As Long)

'ITERATOR1(void, wkeDestroyWebWindow, wkeWebView webWindow, "")
Public Declare Sub wkeDestroyWebWindow Lib "node.dll" (ByVal webView As Long)
                             
'ITERATOR2(void, wkeShowWindow, wkeWebView webWindow, bool show, "")
Public Declare Sub wkeShowWindow _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal show As Boolean)

'ITERATOR2(void, wkeLoadURL, wkeWebView webView, const utf8* url, "")
Public Declare Sub wkeLoadURL _
               Lib "node.dll" (ByVal mbWebView As Long, _
                               ByVal url As String)

'ITERATOR3(void, wkeResize, wkeWebView webView, int w, int h, "") \
Public Declare Sub wkeResize _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal width As Long, _
                               ByVal height As Long)
                             
'ITERATOR1(void, wkeStopLoading, wkeWebView webView, "")
Public Declare Sub wkeStopLoading Lib "node.dll" (ByVal webView As Long)

'ITERATOR1(BOOL, wkeGoForward, wkeWebView webView, "")
Public Declare Sub wkeGoForward Lib "node.dll" (ByVal webView As Long)

'ITERATOR1(BOOL, wkeGoBack, wkeWebView webView, "")
Public Declare Sub wkeGoBack Lib "node.dll" (ByVal webView As Long)

'ITERATOR1(void, wkeEditorSelectAll, wkeWebView webView, "")
Public Declare Sub wkeEditorSelectAll Lib "node.dll" (ByVal webView As Long)

'ITERATOR1(wkeWebFrameHandle, wkeWebFrameGetMainFrame, wkeWebView webView, "")
Public Declare Function wkeWebFrameGetMainFrame _
               Lib "node.dll" (ByVal webView As Long) As Long
                          
'ITERATOR1(const utf8*, wkeGetTitle, wkeWebView webView, "")
Public Declare Function wkeGetTitle Lib "node.dll" (ByVal webView As Long) As Long

'ITERATOR1(const utf8*, wkeGetURL, wkeWebView webView, "")
Public Declare Function wkeGetUrl Lib "node.dll" (ByVal wkeWebView As Long) As Long

'ITERATOR3(void, wkeOnDownload2, wkeWebView webView, wkeDownload2Callback callback, void* param, "")
Public Declare Sub wkeOnDownload2 _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal callback As Long, _
                               Optional ByVal param As Long = 0&)
                               
'ITERATOR1(const utf8*, wkeGetCookie, wkeWebView webView, "")
Public Declare Function wkeGetCookie Lib "node.dll" (ByVal webView As Long) As Long

'ITERATOR1(const wchar_t*, wkeGetCookieW, wkeWebView webView, "")
Public Declare Function wkeGetCookieW Lib "node.dll" (ByVal webView As Long) As String

'ITERATOR1(wkePostBodyElements*, wkeNetGetPostBody, wkeNetJob jobPtr, "")
Public Declare Function wkeNetGetPostBody Lib "node.dll" (jobPtr As Long) As Long

'ITERATOR1(wkeRequestType, wkeNetGetRequestMethod, wkeNetJob jobPtr, "")
Public Declare Function wkeNetGetRequestMethod Lib "node.dll" (jobPtr As Long) As Long

'ITERATOR3(void, wkeOnWindowDestroy, wkeWebView webWindow, wkeWindowDestroyCallback callback, void* param, "")
Public Declare Sub wkeOnWindowDestroy _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal callback As Long, _
                               Optional ByVal param As Long = 0&)
                           
'ITERATOR3(void, wkeOnWindowClosing, wkeWebView webWindow, wkeWindowClosingCallback callback, void* param, "") \
Public Declare Sub wkeOnWindowClosing _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal callback As Long, _
                               Optional ByVal param As Long = 0&)
                            
    

Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal length As Long)
Public Declare Sub CopyMemory2 _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                      ByVal Source As Long, _
                                      ByVal length As Long)

'ITERATOR2(void, wkeSetViewProxy, wkeWebView webView, wkeProxy *proxy, "")
Public Declare Sub wkeSetViewProxy _
               Lib "node.dll" (ByVal webView As Long, _
                               ByVal proxy As Long)

