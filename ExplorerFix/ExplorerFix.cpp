// ExplorerFix.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"



#include <windows.h>
#include <shellapi.h>
#include <stdio.h>
#include <string.h>
#include <time.h>

// --- 必须在 commctrl.h 之前定义 _WIN32_IE ---
#define _WIN32_IE 0x0500
#include <commctrl.h>

// --- 如果 still 报错，手动定义这些常量 (备用方案) ---
#ifndef ICC_STANDARD_CLASSES
#define ICC_STANDARD_CLASSES 0x00000004
#endif
#ifndef ICC_WIN95_CLASSES
#define ICC_WIN95_CLASSES 0x00000001
#endif

// --- 启用视觉样式的 pragma ---
#pragma comment(linker, "\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// --- 定义常量 ---
#define IDM_TRAY_SHOW   1001
#define IDM_TRAY_EXIT   1002
#define IDT_TIMER_FIX   2001
#define WM_TRAYICON     (WM_USER + 1)
#define MAX_LOG_LINES   100
#define LOG_LINE_LENGTH 256
#define IDC_EDIT_LOG    100

// 单实例互斥量名称 (全局唯一)
#define MUTEX_NAME "ExplorerFix_SingleInstance_Mutex"

// --- 全局变量 ---
HWND g_hWnd = NULL;
HWND g_hEdit = NULL;
HINSTANCE g_hInst = NULL;
NOTIFYICONDATA g_nid;
HBRUSH g_hBlackBrush = NULL;
HFONT g_hLogFont = NULL;
HICON g_hMainIcon = NULL;
HANDLE g_hMutex = NULL;

// 日志缓冲区
char g_szLogBuffer[MAX_LOG_LINES][LOG_LINE_LENGTH];
int g_iLogCount = 0;
int g_iLogStart = 0;

// 要移除的后缀 (确保文件保存为 ANSI 编码)
const char* g_szSuffix = " - 文件资源管理器";

// --- 函数声明 ---
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void AddTrayIcon(HWND hwnd);
void RemoveTrayIcon();
void ShowContextMenu(HWND hwnd);
void AddLog(const char* format, ...);
void UpdateLogDisplay();
void CenterWindow(HWND hwnd);
HICON LoadProgramIcon();
void InitCommonControlsEx();
BOOL CheckSingleInstance();
void ActivateExistingInstance();

// --- 检查单实例 ---
BOOL CheckSingleInstance()
{
    g_hMutex = CreateMutex(NULL, TRUE, MUTEX_NAME);
    
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
        return FALSE;
    }
    
    return TRUE;
}

// --- 激活已存在的实例窗口 ---
void ActivateExistingInstance()
{
    HWND hExistingWnd = FindWindow("ExplorerFixClass", "资源管理器标题修正工具");
    
    if (hExistingWnd)
    {
        if (IsIconic(hExistingWnd))
        {
            ShowWindow(hExistingWnd, SW_RESTORE);
        }
        
        ShowWindow(hExistingWnd, SW_SHOW);
        SetForegroundWindow(hExistingWnd);
        FlashWindow(hExistingWnd, TRUE);
    }
}

// --- 初始化通用控件 ---
void InitCommonControlsEx()
{
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
    InitCommonControlsEx(&icex);
}

// --- 加载程序自己的图标 ---
HICON LoadProgramIcon()
{
    HICON hIcon = NULL;
    
    hIcon = LoadIcon(g_hInst, MAKEINTRESOURCE(1));
    if (hIcon) return hIcon;
    
    hIcon = ExtractIcon(g_hInst, "ExplorerFix.exe", 0);
    if (hIcon) return hIcon;
    
    hIcon = (HICON)LoadImage(g_hInst, MAKEINTRESOURCE(1), IMAGE_ICON, 
                              GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), 0);
    if (hIcon) return hIcon;
    
    return LoadIcon(NULL, IDI_APPLICATION);
}

// --- 添加日志 ---
void AddLog(const char* format, ...)
{
    char szLine[LOG_LINE_LENGTH];
    char szTemp[LOG_LINE_LENGTH];
    char szFormat[LOG_LINE_LENGTH];
    int i;
    
    time_t now = time(NULL);
    struct tm* t = localtime(&now);
    sprintf(szTemp, "[%02d:%02d:%02d] ", t->tm_hour, t->tm_min, t->tm_sec);
    
    strcpy(szFormat, szTemp);
    strcat(szFormat, format);
    
    va_list args;
    va_start(args, format);
    vsprintf(szLine, szFormat, args);
    va_end(args);
    
    int iIndex = (g_iLogStart + g_iLogCount) % MAX_LOG_LINES;
    if (g_iLogCount < MAX_LOG_LINES)
    {
        g_iLogCount++;
    }
    else
    {
        g_iLogStart = (g_iLogStart + 1) % MAX_LOG_LINES;
    }
    iIndex = (g_iLogStart + g_iLogCount - 1) % MAX_LOG_LINES;
    strncpy(g_szLogBuffer[iIndex], szLine, LOG_LINE_LENGTH - 1);
    g_szLogBuffer[iIndex][LOG_LINE_LENGTH - 1] = '\0';
    
    UpdateLogDisplay();
}

// --- 更新日志显示 ---
void UpdateLogDisplay()
{
    int i;
    
    if (!g_hEdit) return;
    
    char szFullText[MAX_LOG_LINES * LOG_LINE_LENGTH];
    szFullText[0] = '\0';
    
    for (i = 0; i < g_iLogCount; i++)
    {
        int iIndex = (g_iLogStart + i) % MAX_LOG_LINES;
        strcat(szFullText, g_szLogBuffer[iIndex]);
        strcat(szFullText, "\r\n");
    }
    
    SetWindowText(g_hEdit, szFullText);
    
    SendMessage(g_hEdit, EM_SETSEL, -1, -1);
    SendMessage(g_hEdit, EM_SCROLLCARET, 0, 0);
}

// --- 窗口居中 ---
void CenterWindow(HWND hwnd)
{
    RECT rcWindow, rcScreen;
    int nWidth, nHeight, nX, nY;
    
    GetWindowRect(hwnd, &rcWindow);
    SystemParametersInfo(SPI_GETWORKAREA, 0, &rcScreen, 0);
    
    nWidth = rcWindow.right - rcWindow.left;
    nHeight = rcWindow.bottom - rcWindow.top;
    
    nX = rcScreen.left + (rcScreen.right - rcScreen.left - nWidth) / 2;
    nY = rcScreen.top + (rcScreen.bottom - rcScreen.top - nHeight) / 2;
    
    SetWindowPos(hwnd, NULL, nX, nY, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
}

// --- 枚举窗口的回调函数 ---
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
    char szTitle[1024];
    int nLen = GetWindowText(hwnd, szTitle, 1024);
    
    if (nLen > 0)
    {
        int nSuffixLen = strlen(g_szSuffix);
        if (nLen >= nSuffixLen)
        {
            if (strncmp(szTitle + nLen - nSuffixLen, g_szSuffix, nSuffixLen) == 0)
            {
                char szOldTitle[1024];
                strcpy(szOldTitle, szTitle);
                
                szTitle[nLen - nSuffixLen] = '\0';
                
                if (SetWindowText(hwnd, szTitle))
                {
                    AddLog("[OK] 已修正窗口标题");
                    AddLog("     句柄：0x%08X", (unsigned int)hwnd);
                    AddLog("     原：%s", szOldTitle);
                    AddLog("     新：%s", szTitle);
                    AddLog("-------------------------------------------");
                }
            }
        }
    }
    return TRUE;
}

// --- 主窗口过程 ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    int i;
    
    switch (msg)
    {
    case WM_CREATE:
        {
            g_hBlackBrush = CreateSolidBrush(RGB(0, 0, 0));
            
            g_hLogFont = CreateFont(
                -14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, "Courier New"
            );
            
            g_hEdit = CreateWindowEx(
                WS_EX_CLIENTEDGE,
                "EDIT",
                "",
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | 
                ES_READONLY | ES_AUTOVSCROLL,
                5, 5, 0, 0,
                hwnd,
                (HMENU)IDC_EDIT_LOG,
                g_hInst,
                NULL
            );
            
            if (g_hEdit)
            {
                SendMessage(g_hEdit, WM_SETFONT, (WPARAM)g_hLogFont, TRUE);
            }
            
            SetTimer(hwnd, IDT_TIMER_FIX, 1000, NULL);
            AddTrayIcon(hwnd);
            
            AddLog("===========================================");
            AddLog("  资源管理器标题修正工具已启动");
            AddLog("  按右键托盘图标可退出程序");
            AddLog("===========================================");
            AddLog("");
        }
        break;

    case WM_SIZE:
        if (g_hEdit)
        {
            MoveWindow(g_hEdit, 5, 5, LOWORD(lParam) - 10, HIWORD(lParam) - 10, TRUE);
        }
        break;

    case WM_TIMER:
        if (wParam == IDT_TIMER_FIX)
        {
            EnumWindows(EnumWindowsProc, 0);
        }
        break;

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
        {
            HDC hdc = (HDC)wParam;
            int iCtrlID = GetDlgCtrlID((HWND)lParam);
            
            if (iCtrlID == IDC_EDIT_LOG)
            {
                SetTextColor(hdc, RGB(0, 255, 0));
                SetBkColor(hdc, RGB(0, 0, 0));
                return (LRESULT)g_hBlackBrush;
            }
        }
        break;

    case WM_TRAYICON:
        switch (lParam)
        {
        case WM_LBUTTONDBLCLK:
            ShowWindow(hwnd, SW_SHOW);
            SetForegroundWindow(hwnd);
            break;
        case WM_RBUTTONUP:
            ShowContextMenu(hwnd);
            break;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDM_TRAY_SHOW:
            ShowWindow(hwnd, SW_SHOW);
            SetForegroundWindow(hwnd);
            break;
        case IDM_TRAY_EXIT:
            DestroyWindow(hwnd);
            break;
        }
        break;

    case WM_CLOSE:
        ShowWindow(hwnd, SW_HIDE);
        return 0;

    case WM_DESTROY:
        AddLog("\n程序正在退出...");
        KillTimer(hwnd, IDT_TIMER_FIX);
        RemoveTrayIcon();
        
        if (g_hBlackBrush) DeleteObject(g_hBlackBrush);
        if (g_hLogFont) DeleteObject(g_hLogFont);
        if (g_hMainIcon) DestroyIcon(g_hMainIcon);
        
        if (g_hMutex)
        {
            ReleaseMutex(g_hMutex);
            CloseHandle(g_hMutex);
            g_hMutex = NULL;
        }
        
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// --- 托盘图标操作 ---
void AddTrayIcon(HWND hwnd)
{
    memset(&g_nid, 0, sizeof(NOTIFYICONDATA));
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    
    g_hMainIcon = LoadProgramIcon();
    g_nid.hIcon = g_hMainIcon;
    
    strcpy(g_nid.szTip, "资源管理器标题修正工具 (双击显示)");
    
    Shell_NotifyIcon(NIM_ADD, &g_nid);
}

void RemoveTrayIcon()
{
    Shell_NotifyIcon(NIM_DELETE, &g_nid);
}

void ShowContextMenu(HWND hwnd)
{
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING, IDM_TRAY_SHOW, "显示主窗口");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, IDM_TRAY_EXIT, "退出程序");
    
    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_RIGHTALIGN | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, NULL);
    PostMessage(hwnd, WM_NULL, 0, 0);
    
    DestroyMenu(hMenu);
}

// --- 程序入口 ---
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    WNDCLASSEX wc;
    MSG msg;
    HWND hwnd;
    int i;

    g_hInst = hInstance;
    
    // 检查单实例
    if (!CheckSingleInstance())
    {
        ActivateExistingInstance();
        MessageBox(NULL, "程序已经在运行中！\n\n已激活已有窗口。", "提示", MB_ICONINFORMATION | MB_OK);
        return 0;
    }
    
    // 初始化通用控件
    InitCommonControlsEx();

    g_hMainIcon = LoadProgramIcon();

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = g_hMainIcon;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = "ExplorerFixClass";
    wc.hIconSm = g_hMainIcon;

    if (!RegisterClassEx(&wc))
    {
        MessageBox(NULL, "窗口注册失败！", "错误", MB_ICONERROR);
        return 0;
    }

    hwnd = CreateWindowEx(
        0,
        "ExplorerFixClass",
        "资源管理器标题修正工具",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT,
        500, 400,
        NULL, NULL, hInstance, NULL
    );

    if (!hwnd)
    {
        MessageBox(NULL, "窗口创建失败！", "错误", MB_ICONERROR);
        return 0;
    }

    g_hWnd = hwnd;
    CenterWindow(hwnd);
    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return msg.wParam;
}