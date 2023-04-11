#### 介绍

MiniBlink是国内著名的浏览器专家龙泉扫地僧开发的chrome内核的浏览器组件。

项目首页：http://www.miniblink.net/index.html

Miniblink压缩后仅几M左右的体积，只需一个dll，通过纯C接口，数行代码即可集成到各种软件，并且完美支持WinXP、NpAPI

但是由于Miniblink的免费版的node.dll文件导出的方式是__cdecl，VB6导入函数目前VB6的实现方式只看到有这个由国人写的Ocx组件：

https://github.com/imxcstar/vb6-miniblink-SBrowser

@imxcstar 确实是一位VB大牛，至少我没想到能用这种方式调用。

我开这个项目的目的，是希望把Miniblink的最后开源的版本，把node.dll 改为__stdcall导出。

扫地僧的代码很友好，修改起来非常简单：

1、wke\wkedefine.h<br />
修改#define WKE_CALL_TYPE __cdecl为#define WKE_CALL_TYPE __stdcall

2、miniblink项目增加一个预定义文件 miniblink.def

3、wke\wke2.cpp文件，把
```
void wkeUtilRelasePrintPdfDatas(const wkePdfDatas* datas)
{
    for (int i = 0; i < datas->count; ++i) {	
        free((void *)(datas->datas[i]));		
    }
    free((void *)(datas->sizes));	
    free((void *)(datas->datas));	
    delete datas;	
}

const wkePdfDatas* wkeUtilPrintToPdf(wkeWebView webView, wkeWebFrameHandle frameId, const wkePrintSettings* settings)
{
    content::WebPage* webPage = webView->webPage();	
    blink::WebFrame* webFrame = webPage->getWebFrameFromFrameId(wke::CWebView::wkeWebFrameHandleToFrameId(webPage, frameId));	
    return wke::printToPdf(webView, webFrame, settings);	
}

const wkeMemBuf* wkePrintToBitmap(wkeWebView webView, wkeWebFrameHandle frameId, const wkeScreenshotSettings* settings)
{
    return wke::printToBitmap(webView, settings);	
}
```
这三个函数名后面加个2

4、依次编译这些项目
```
libcurl.lib 
harfbuzz.lib 
libxml.lib 
libjpeg.lib 
libpng.lib 
openssl.lib 
ots.lib 
skia.lib 
zlib.lib 
wolfssl.lib 
v8_5_7_1.lib 
v8_5_7.lib 
node.lib
```
从而生成支持stdcall的node.dll





