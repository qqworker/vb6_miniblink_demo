Attribute VB_Name = "CallbackModule"
Option Explicit

Public dataBindCallback As wkeNetJobDataBind


'typedef void(WKE_CALL_TYPE*wkeNetJobDataRecvCallback)(void* ptr, wkeNetJob job, const char* data, int length);
Sub wkeNetJobDataRecvCallback(ByVal ptr As Long, job As Long, data As Long, length As Long)
    AppendLog "wkeNetJobDataRecvCallback"
    'AppendLog "Job：" & CStr(job)
End Sub


'typedef void(WKE_CALL_TYPE*wkeNetJobDataFinishCallback)(void* ptr, wkeNetJob job, wkeLoadingResult result);
Sub wkeNetJobDataFinishCallback(ByVal ptr As Long, job As Long, result As wkeLoadingResult)
    MsgBox "下载完成"
End Sub

'typedef wkeDownloadOpt(WKE_CALL_TYPE*wkeDownload2Callback)(
'    wkeWebView webView,
'    void* param,
'    size_t expectedContentLength,
'    const char* url,
'    const char* mime,
'    const char* disposition,
'    wkeNetJob job,
'    wkeNetJobDataBind* dataBind);
Function wkeDownload2Callback(ByVal webView As Long, _
                                         ByVal params As Long, _
                                         ByVal expectedContentLength As Long, _
                                         ByVal url As Long, _
                                         ByVal mime As Long, _
                                         ByVal disposition As Long, _
                                         ByVal job As Long, _
                                         ByVal dataBind As Long) As Long
    Static t As wkeNetJobDataBind
    
    t.param = 0
    t.finishCallback = FuncAddressOf(AddressOf wkeNetJobDataFinishCallback)
    t.recvCallback = FuncAddressOf(AddressOf wkeNetJobDataRecvCallback)
    CopyMemory2 dataBind, VarPtr(t), LenB(t)
    
    wkeDownload2Callback = 1
End Function


'typedef void(WKE_CALL_TYPE*wkeWindowDestroyCallback)(wkeWebView webWindow, void* param);
'不加这个方法，IDE会容易崩溃
Sub wkeWindowDestroyCallback(ByVal webView As Long, ByVal params As Long)
    AppendLog "wkeWindowDestroyCallback"
End Sub

'typedef bool(WKE_CALL_TYPE*wkeWindowClosingCallback)(wkeWebView webWindow, void* param);
Sub wkeWindowClosingCallback(ByVal webView As Long, ByVal params As Long)
    AppendLog "wkeWindowClosingCallback"
End Sub

