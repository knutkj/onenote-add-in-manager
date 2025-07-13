#pragma once

#include <windows.h>
#include <string>

// Log file name constant
const wchar_t* const LOGFILE = L"MinimalComServer.log";

// Utility function to get the full path of the current DLL module
std::wstring GetDllModulePath();

// Utility function to log to the log file
void Log(const wchar_t* message);