#pragma once
#include <windows.h>
HANDLE g_hJob = nullptr;
PROCESS_INFORMATION g_pi = {};
bool StartServer()
{
    WCHAR exePath[MAX_PATH];
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);

    WCHAR* p = wcsrchr(exePath, L'\\');
    if (p)
        wcscpy_s(p + 1, MAX_PATH - (p - exePath) - 1, L"server.exe");

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    g_hJob = CreateJobObjectW(nullptr, nullptr);
    if (!g_hJob)
        return false;

    JOBOBJECT_EXTENDED_LIMIT_INFORMATION info = {};
    info.BasicLimitInformation.LimitFlags =
        JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

    SetInformationJobObject(
        g_hJob,
        JobObjectExtendedLimitInformation,
        &info,
        sizeof(info));

    BOOL ok = CreateProcessW(
        exePath,
        nullptr,
        nullptr,
        nullptr,
        FALSE,
        CREATE_NO_WINDOW,
        nullptr,
        nullptr,
        &si,
        &g_pi);

    if (!ok)
    {
        CloseHandle(g_hJob);
        g_hJob = nullptr;
        return false;
    }

    AssignProcessToJobObject(g_hJob, g_pi.hProcess);

    return true;
}
void StopServer()
{
    if (g_pi.hThread)
    {
        CloseHandle(g_pi.hThread);
        g_pi.hThread = nullptr;
    }

    if (g_pi.hProcess)
    {
        CloseHandle(g_pi.hProcess);
        g_pi.hProcess = nullptr;
    }

    if (g_hJob)
    {
        CloseHandle(g_hJob);   // ×Ô¶¯½áÊø server.exe
        g_hJob = nullptr;
    }
}