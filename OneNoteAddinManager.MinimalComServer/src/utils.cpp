#include "utils.h"
#include <fstream>
#include <sstream>
#include <chrono>
#include <iomanip>

extern HMODULE g_hModule; // Defined in MinimalComServer.cpp

// Utility function to get the full path of the current DLL module.
std::wstring GetDllModulePath() {
    wchar_t modulePath[MAX_PATH];
    if (g_hModule) {
        GetModuleFileNameW(g_hModule, modulePath, MAX_PATH);
        return std::wstring(modulePath);
    }
    
    return std::wstring();
}

// Utility function to log to the log file.
void Log(const wchar_t* message) {
    // Get the full path to the DLL.
    std::wstring modulePath = GetDllModulePath();
    std::wstring logDir;
    size_t lastSlash = modulePath.find_last_of(L"\\/");
    if (lastSlash != std::wstring::npos) {
        // Extract the directory part from the DLL path.
        logDir = modulePath.substr(0, lastSlash + 1);
    } else {
        // If not found, use an empty string (current working directory).
        logDir = L"";
    }

    // Build the full log file path.
    std::wstring logPath = logDir + LOGFILE;

    // Open the log file for appending.
    std::wofstream logFile(logPath, std::ios::app);
    if (logFile.is_open()) {
        // Get current time for timestamp.
        auto now = std::chrono::system_clock::now();
        std::time_t now_c = std::chrono::system_clock::to_time_t(now);
        struct tm timeinfo;
        localtime_s(&timeinfo, &now_c);
        std::wstringstream ts;

        // Format timestamp as [YYYY-MM-DD HH:MM:SS]:
        ts << L"[" << std::put_time(&timeinfo, L"%Y-%m-%d %H:%M:%S") << L"]: ";

        // Write timestamp and message to log file.
        logFile << ts.str() << message << std::endl;

        // Close the log file.
        logFile.close();
    }
}