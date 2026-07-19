// GCAW.cpp : 定义应用程序的入口点。
//

#include "framework.h"
#include "GCAW.h"
#include "server.h"
#include <windows.h>
#include <wrl.h>
#include <WebView2.h>
#include <ShlObj.h>

using namespace Microsoft::WRL;

ComPtr<ICoreWebView2Controller> g_controller;
ComPtr<ICoreWebView2> g_webview;

HWND g_hWnd;

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_SIZE:
    {
        if (g_controller)
        {
            RECT rc;
            GetClientRect(hwnd, &rc);
            g_controller->put_Bounds(rc);
        }
        return 0;
    }

    case WM_DESTROY:
        StopServer();
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void InitWebView()
{
    WCHAR tempPath[MAX_PATH];

    // 获取系统临时目录
    GetTempPathW(
        MAX_PATH,
        tempPath
    );

    // 添加自己的子目录
    wcscat_s(
        tempPath,
        L"GCAW"
    );

    // 创建目录
    CreateDirectoryW(
        tempPath,
        nullptr
    );


    CreateCoreWebView2EnvironmentWithOptions(
        nullptr,
        tempPath,
        nullptr,

        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [](HRESULT result,
                ICoreWebView2Environment* env) -> HRESULT
            {
                env->CreateCoreWebView2Controller(
                    g_hWnd,

                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [](HRESULT result,
                            ICoreWebView2Controller* controller) -> HRESULT
                        {
                            if (controller == nullptr)
                                return E_FAIL;

                            g_controller = controller;

                            g_controller->get_CoreWebView2(&g_webview);

                            ComPtr<ICoreWebView2Settings> settings;
                            g_webview->get_Settings(&settings);
                            settings->put_AreDefaultContextMenusEnabled(FALSE);
                            //settings->put_IsZoomControlEnabled(FALSE);
                            settings->put_AreDevToolsEnabled(FALSE);

                            RECT rc;
                            GetClientRect(g_hWnd, &rc);
                            g_controller->put_Bounds(rc);

                            EventRegistrationToken navToken;
                            g_webview->add_NavigationCompleted(
                                Callback<ICoreWebView2NavigationCompletedEventHandler>(
                                    [](ICoreWebView2* sender,
                                        ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT
                                    {
                                        sender->ExecuteScript(
                                            L"document.cookie='token=fixed-token; Max-Age=31536000; Path=/; SameSite=Lax';",
                                            nullptr);
                                        return S_OK;
                                    }).Get(),
                                        &navToken);
                            g_webview->Navigate(L"http://127.0.0.2:8964");
                            g_webview->ExecuteScript(
                                LR"(document.cookie ="token=fixed-token; "
                                "Max-Age=31536000; " "Path=/; ""SameSite=Lax"; )",
                                nullptr
                            );


                            return S_OK;
                        }).Get());

                return S_OK;
            }).Get());
}

int WINAPI WinMain(HINSTANCE hInst,
    HINSTANCE,
    LPSTR,
    int nCmdShow)
{
    StartServer();
    WNDCLASS wc = {};

    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = L"GCAW";
    wc.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_GCAW));

    SetProcessDpiAwarenessContext(
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    );

    RegisterClass(&wc);

    g_hWnd = CreateWindowEx(
        0,
        wc.lpszClassName,
        L"成绩查看分析工具（GCAW客户端）v1.0.0 （若未登录请按Ctrl+R）",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        1200,
        800,
        nullptr,
        nullptr,
        hInst,
        nullptr);

    ShowWindow(g_hWnd, nCmdShow);

    InitWebView();

    MSG msg;

    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}