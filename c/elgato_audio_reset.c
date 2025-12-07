/*
 * elgato_reset.c - Complete Elgato Audio Reset Tool
 * Kills audio processes, restarts audio services, relaunches WaveLink/StreamDeck,
 * and sets audio defaults.
 * 
 * Version: 0.9.5
 * Compile: cl /O2 elgato_reset.c
 */

#define APP_VERSION L"0.9.5"

#define INITGUID
#define COBJMACROS
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <tlhelp32.h>
#include <psapi.h>
#include <mmdeviceapi.h>
#include <propsys.h>
#include <functiondiscoverykeys_devpkey.h>
#include <endpointvolume.h>
#include <objbase.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdi32.lib")

/* Use Windows subsystem to hide console window */
#pragma comment(linker, "/SUBSYSTEM:WINDOWS /ENTRY:mainCRTStartup")

/* ==========================================================================
 * CONFIGURATION - These can be overridden by config.txt in the same folder
 * 
 * If config.txt exists, it will be read. Otherwise these defaults are used.
 * The installer auto-detects your current audio settings and creates config.txt.
 * ==========================================================================*/

/* Default Output Device - Used for general audio (games, music, videos)
 * This is the "Default Device" in Windows Sound -> Playback tab */
static WCHAR g_playbackDefault[256] = L"System (Elgato Virtual Audio)";

/* Communications Output Device - Used for voice chat apps (Discord, Teams, etc.)
 * This is the "Default Communication Device" in Windows Sound -> Playback tab */
static WCHAR g_playbackComm[256] = L"Voice Chat (Elgato Virtual Audio)";

/* Default Input Device - Used for general recording
 * This is the "Default Device" in Windows Sound -> Recording tab */
static WCHAR g_recordDefault[256] = L"Microphone (Razer Kraken V4 2.4 - Chat)";

/* Communications Input Device - Used for voice chat apps (Discord, Teams, etc.)
 * This is the "Default Communication Device" in Windows Sound -> Recording tab */
static WCHAR g_recordComm[256] = L"Microphone (Razer Kraken V4 2.4 - Chat)";

/* Additional devices to unmute and set to 100% volume after reset */
static const WCHAR* OUTPUT_RAZER_CHAT = L"Speakers (Razer Kraken V4 2.4 - Chat)";
static const WCHAR* OUTPUT_RAZER_GAME = L"Speakers (Razer Kraken V4 2.4 - Game)";

/* Timing settings (seconds) */
static int MAX_SERVICE_WAIT = 30;   /* Max time to wait for audio services */
static int MAX_DEVICE_WAIT = 60;    /* Max time to wait for Elgato devices */
static int POLL_INTERVAL = 2;       /* How often to check during waits */

/* ========== Globals ========== */
static char g_logPath[MAX_PATH] = {0};
static FILE* g_logFile = NULL;
static char g_waveLinkPath[MAX_PATH] = {0};
static char g_waveLinkSEPath[MAX_PATH] = {0};
static char g_streamDeckPath[MAX_PATH] = {0};

/* Saved volume levels for safe reset */
static float g_savedPlaybackVolume = -1.0f;  /* -1 means not saved */
static float g_savedCommVolume = -1.0f;

/* ========== Audio GUIDs (manually defined) ========== */
DEFINE_GUID(MY_CLSID_MMDeviceEnumerator, 0xBCDE0395, 0xE52F, 0x467C, 0x8E, 0x3D, 0xC4, 0x57, 0x92, 0x91, 0x69, 0x2E);
DEFINE_GUID(MY_IID_IMMDeviceEnumerator, 0xA95664D2, 0x9614, 0x4F35, 0xA7, 0x46, 0xDE, 0x8D, 0xB6, 0x36, 0x17, 0xE6);
DEFINE_GUID(MY_IID_IAudioEndpointVolume, 0x5CDF2C82, 0x841E, 0x4546, 0x97, 0x22, 0x0C, 0xF7, 0x40, 0x78, 0x22, 0x9A);
/* IPolicyConfig (undocumented) */
DEFINE_GUID(CLSID_PolicyConfigClient, 0x870AF99C, 0x171D, 0x4F9E, 0xAF, 0x0D, 0xE6, 0x3D, 0xF4, 0x0C, 0x2B, 0xC9);
DEFINE_GUID(IID_IPolicyConfig, 0xF8679F50, 0x850A, 0x41CF, 0x9C, 0x72, 0x43, 0x0F, 0x29, 0x02, 0x90, 0xC8);

typedef struct IPolicyConfig IPolicyConfig;
typedef struct IPolicyConfigVtbl {
    /* IUnknown */
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(IPolicyConfig*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(IPolicyConfig*);
    ULONG (STDMETHODCALLTYPE *Release)(IPolicyConfig*);
    /* IPolicyConfig methods we don't use */
    HRESULT (STDMETHODCALLTYPE *GetMixFormat)(IPolicyConfig*, LPCWSTR, void**);
    HRESULT (STDMETHODCALLTYPE *GetDeviceFormat)(IPolicyConfig*, LPCWSTR, int, void**);
    HRESULT (STDMETHODCALLTYPE *ResetDeviceFormat)(IPolicyConfig*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *SetDeviceFormat)(IPolicyConfig*, LPCWSTR, void*, void*);
    HRESULT (STDMETHODCALLTYPE *GetProcessingPeriod)(IPolicyConfig*, LPCWSTR, int, void*, void*);
    HRESULT (STDMETHODCALLTYPE *SetProcessingPeriod)(IPolicyConfig*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *GetShareMode)(IPolicyConfig*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *SetShareMode)(IPolicyConfig*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *GetPropertyValue)(IPolicyConfig*, LPCWSTR, const PROPERTYKEY*, PROPVARIANT*);
    HRESULT (STDMETHODCALLTYPE *SetPropertyValue)(IPolicyConfig*, LPCWSTR, const PROPERTYKEY*, const PROPVARIANT*);
    HRESULT (STDMETHODCALLTYPE *SetDefaultEndpoint)(IPolicyConfig*, LPCWSTR, ERole);
    HRESULT (STDMETHODCALLTYPE *SetEndpointVisibility)(IPolicyConfig*, LPCWSTR, int);
} IPolicyConfigVtbl;
struct IPolicyConfig { IPolicyConfigVtbl* lpVtbl; };

/* ========== GUI Control IDs ========== */
#define ID_COMBO_PLAYBACK_DEFAULT  101
#define ID_COMBO_RECORD_DEFAULT    102
#define ID_COMBO_PLAYBACK_COMM     103
#define ID_COMBO_RECORD_COMM       104
#define ID_BUTTON_SAVE             105
#define ID_HEADER                  106  /* Main header (blue) */
#define ID_BUTTON_BROWSE           107
#define ID_EDIT_FOLDER             108
#define ID_BUTTON_RUN              109
#define ID_CHECK_RUN_BACKGROUND    110
#define ID_CHECK_SHOW_NOTIFY       111
#define ID_LABEL_START             150  /* Title labels start at 150 */
#define ID_DESC_START              200  /* Description labels start at 200 */

/* ========== System Tray ========== */
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_ICON 1001
#define ID_TRAY_OPEN 2001
#define ID_TRAY_RUN  2002
#define ID_TRAY_EXIT 2003
static NOTIFYICONDATAW g_nid = {0};
static HWND g_trayHwnd = NULL;
static const wchar_t* g_trayStatus = L"Starting...";
static int g_trayAction = 0;  /* 0=none, 1=open config, 2=run reset, 3=exit */

/* ========== Device Lists for GUI ========== */
#define MAX_DEVICES 32
static WCHAR g_playbackDeviceNames[MAX_DEVICES][256];
static WCHAR g_recordDeviceNames[MAX_DEVICES][256];
static int g_playbackDeviceCount = 0;
static int g_recordDeviceCount = 0;
static char g_exeDir[MAX_PATH] = {0};
static char g_installDir[MAX_PATH] = {0};  /* User-selected install folder */
static char g_currentExePath[MAX_PATH] = {0};  /* Current exe location */
static int g_shouldRun = 0;  /* Flag to indicate if we should run reset after GUI */
static int g_configExists = 0;  /* Flag to track if config file exists */
static int g_runInBackground = 0;  /* If true, run silently without GUI */
static int g_showNotification = 1;  /* If true, show notification on completion */

/* Saved config values (for comparison with current Windows settings) */
static WCHAR g_savedPlaybackDefault[256] = {0};
static WCHAR g_savedPlaybackComm[256] = {0};
static WCHAR g_savedRecordDefault[256] = {0};
static WCHAR g_savedRecordComm[256] = {0};

/* Current Windows audio settings */
static WCHAR g_currentPlaybackDefault[256] = {0};
static WCHAR g_currentPlaybackComm[256] = {0};
static WCHAR g_currentRecordDefault[256] = {0};
static WCHAR g_currentRecordComm[256] = {0};

/* ========== Config File Path Helper ========== */
static void getConfigPath(char* configPath, size_t len) {
    snprintf(configPath, len, "%s\\config.txt", g_exeDir);
}

/* ========== Config File Loading ========== */
static int loadConfig(const char* exePath) {
    char configPath[MAX_PATH];
    
    /* Store the full exe path */
    strncpy(g_currentExePath, exePath, MAX_PATH);
    
    strncpy(g_exeDir, exePath, MAX_PATH);
    char* lastSlash = strrchr(g_exeDir, '\\');
    if (lastSlash) *lastSlash = '\0';
    
    getConfigPath(configPath, MAX_PATH);
    
    FILE* f = fopen(configPath, "r");
    if (!f) {
        /* No config file */
        return 0;
    }
    
    char line[512];
    while (fgets(line, sizeof(line), f)) {
        /* Remove newline */
        char* nl = strchr(line, '\n');
        if (nl) *nl = '\0';
        char* cr = strchr(line, '\r');
        if (cr) *cr = '\0';
        
        /* Skip comments and empty lines */
        if (line[0] == '#' || line[0] == '\0') continue;
        
        /* Parse KEY=VALUE */
        char* eq = strchr(line, '=');
        if (!eq) continue;
        
        *eq = '\0';
        char* key = line;
        char* value = eq + 1;
        
        /* Convert value to wide string and store */
        if (strcmp(key, "PLAYBACK_DEFAULT") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_playbackDefault, 256);
            wcsncpy(g_savedPlaybackDefault, g_playbackDefault, 256);
        } else if (strcmp(key, "PLAYBACK_COMM") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_playbackComm, 256);
            wcsncpy(g_savedPlaybackComm, g_playbackComm, 256);
        } else if (strcmp(key, "RECORD_DEFAULT") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_recordDefault, 256);
            wcsncpy(g_savedRecordDefault, g_recordDefault, 256);
        } else if (strcmp(key, "RECORD_COMM") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_recordComm, 256);
            wcsncpy(g_savedRecordComm, g_recordComm, 256);
        } else if (strcmp(key, "RUN_IN_BACKGROUND") == 0) {
            g_runInBackground = (strcmp(value, "1") == 0 || _stricmp(value, "true") == 0);
        } else if (strcmp(key, "SHOW_NOTIFICATION") == 0) {
            g_showNotification = (strcmp(value, "1") == 0 || _stricmp(value, "true") == 0);
        }
    }
    
    fclose(f);
    return 1;
}

/* ========== Move/Copy Files to Install Folder ========== */
static int moveToInstallDir(void) {
    if (g_installDir[0] == '\0') return 0;
    
    /* Check if install dir is different from current exe location */
    if (_stricmp(g_installDir, g_exeDir) == 0) {
        /* Same directory, nothing to do */
        return 1;
    }
    
    /* Create install directory if it doesn't exist */
    CreateDirectoryA(g_installDir, NULL);
    
    /* Build paths */
    char srcExePath[MAX_PATH], destExePath[MAX_PATH];
    char srcConfigPath[MAX_PATH], destConfigPath[MAX_PATH];
    char srcLogDir[MAX_PATH], destLogDir[MAX_PATH];
    
    snprintf(srcExePath, MAX_PATH, "%s\\elgato_audio_reset.exe", g_exeDir);
    snprintf(destExePath, MAX_PATH, "%s\\elgato_audio_reset.exe", g_installDir);
    snprintf(srcConfigPath, MAX_PATH, "%s\\config.txt", g_exeDir);
    snprintf(destConfigPath, MAX_PATH, "%s\\config.txt", g_installDir);
    snprintf(srcLogDir, MAX_PATH, "%s\\logs", g_exeDir);
    snprintf(destLogDir, MAX_PATH, "%s\\logs", g_installDir);
    
    /* Copy the exe */
    CopyFileA(g_currentExePath, destExePath, FALSE);
    
    /* Move config if it exists in old location */
    if (GetFileAttributesA(srcConfigPath) != INVALID_FILE_ATTRIBUTES) {
        MoveFileExA(srcConfigPath, destConfigPath, MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED);
    }
    
    /* Move logs folder if it exists */
    if (GetFileAttributesA(srcLogDir) != INVALID_FILE_ATTRIBUTES) {
        /* Create dest logs folder */
        CreateDirectoryA(destLogDir, NULL);
        
        /* Find and move all log files */
        WIN32_FIND_DATAA fd;
        char searchPath[MAX_PATH];
        snprintf(searchPath, MAX_PATH, "%s\\*.log", srcLogDir);
        HANDLE hFind = FindFirstFileA(searchPath, &fd);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                char srcFile[MAX_PATH], destFile[MAX_PATH];
                snprintf(srcFile, MAX_PATH, "%s\\%s", srcLogDir, fd.cFileName);
                snprintf(destFile, MAX_PATH, "%s\\%s", destLogDir, fd.cFileName);
                MoveFileExA(srcFile, destFile, MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED);
            } while (FindNextFileA(hFind, &fd));
            FindClose(hFind);
        }
        
        /* Try to remove old logs folder (will fail if not empty, which is fine) */
        RemoveDirectoryA(srcLogDir);
    }
    
    /* Update g_exeDir to the new install location */
    strncpy(g_exeDir, g_installDir, MAX_PATH);
    
    return 1;
}

/* ========== Save Config File ========== */
static void saveConfig(void) {
    /* First, move files to install dir if it changed */
    if (g_installDir[0] != '\0') {
        moveToInstallDir();
        /* Use install dir for config */
        strncpy(g_exeDir, g_installDir, MAX_PATH);
    }
    
    /* Create logs subfolder */
    char logDir[MAX_PATH];
    snprintf(logDir, MAX_PATH, "%s\\logs", g_exeDir);
    CreateDirectoryA(logDir, NULL);
    
    char configPath[MAX_PATH];
    getConfigPath(configPath, MAX_PATH);
    
    FILE* f = fopen(configPath, "w");
    if (!f) return;
    
    SYSTEMTIME st;
    GetLocalTime(&st);
    
    fprintf(f, "# Elgato Audio Reset Configuration\n");
    fprintf(f, "# Created on %04d-%02d-%02d %02d:%02d:%02d\n",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    fprintf(f, "# Delete this file to show the device selection GUI again\n\n");
    
    char buf[512];
    WideCharToMultiByte(CP_UTF8, 0, g_playbackDefault, -1, buf, sizeof(buf), NULL, NULL);
    fprintf(f, "PLAYBACK_DEFAULT=%s\n", buf);
    WideCharToMultiByte(CP_UTF8, 0, g_playbackComm, -1, buf, sizeof(buf), NULL, NULL);
    fprintf(f, "PLAYBACK_COMM=%s\n", buf);
    WideCharToMultiByte(CP_UTF8, 0, g_recordDefault, -1, buf, sizeof(buf), NULL, NULL);
    fprintf(f, "RECORD_DEFAULT=%s\n", buf);
    WideCharToMultiByte(CP_UTF8, 0, g_recordComm, -1, buf, sizeof(buf), NULL, NULL);
    fprintf(f, "RECORD_COMM=%s\n", buf);
    fprintf(f, "RUN_IN_BACKGROUND=%d\n", g_runInBackground ? 1 : 0);
    fprintf(f, "SHOW_NOTIFICATION=%d\n", g_showNotification ? 1 : 0);
    
    fclose(f);
}

/* ========== Enumerate Audio Devices for GUI ========== */
static void enumerateDevicesForGUI(void) {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) return;
    
    IMMDeviceEnumerator* pEnum = NULL;
    hr = CoCreateInstance(&MY_CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                          &MY_IID_IMMDeviceEnumerator, (void**)&pEnum);
    if (FAILED(hr)) { CoUninitialize(); return; }
    
    /* Get playback devices */
    IMMDeviceCollection* pCol = NULL;
    hr = IMMDeviceEnumerator_EnumAudioEndpoints(pEnum, eRender, DEVICE_STATE_ACTIVE, &pCol);
    if (SUCCEEDED(hr)) {
        UINT count = 0;
        IMMDeviceCollection_GetCount(pCol, &count);
        g_playbackDeviceCount = 0;
        for (UINT i = 0; i < count && g_playbackDeviceCount < MAX_DEVICES; i++) {
            IMMDevice* pDev = NULL;
            if (SUCCEEDED(IMMDeviceCollection_Item(pCol, i, &pDev))) {
                IPropertyStore* pStore = NULL;
                if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDev, STGM_READ, &pStore))) {
                    PROPVARIANT pv;
                    PropVariantInit(&pv);
                    if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                        wcsncpy(g_playbackDeviceNames[g_playbackDeviceCount], pv.pwszVal, 255);
                        g_playbackDeviceCount++;
                        PropVariantClear(&pv);
                    }
                    IPropertyStore_Release(pStore);
                }
                IMMDevice_Release(pDev);
            }
        }
        IMMDeviceCollection_Release(pCol);
    }
    
    /* Get recording devices */
    hr = IMMDeviceEnumerator_EnumAudioEndpoints(pEnum, eCapture, DEVICE_STATE_ACTIVE, &pCol);
    if (SUCCEEDED(hr)) {
        UINT count = 0;
        IMMDeviceCollection_GetCount(pCol, &count);
        g_recordDeviceCount = 0;
        for (UINT i = 0; i < count && g_recordDeviceCount < MAX_DEVICES; i++) {
            IMMDevice* pDev = NULL;
            if (SUCCEEDED(IMMDeviceCollection_Item(pCol, i, &pDev))) {
                IPropertyStore* pStore = NULL;
                if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDev, STGM_READ, &pStore))) {
                    PROPVARIANT pv;
                    PropVariantInit(&pv);
                    if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                        wcsncpy(g_recordDeviceNames[g_recordDeviceCount], pv.pwszVal, 255);
                        g_recordDeviceCount++;
                        PropVariantClear(&pv);
                    }
                    IPropertyStore_Release(pStore);
                }
                IMMDevice_Release(pDev);
            }
        }
        IMMDeviceCollection_Release(pCol);
    }
    
    /* Get current defaults */
    IMMDevice* pDefault = NULL;
    if (SUCCEEDED(IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eRender, eConsole, &pDefault))) {
        IPropertyStore* pStore = NULL;
        if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDefault, STGM_READ, &pStore))) {
            PROPVARIANT pv;
            PropVariantInit(&pv);
            if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                wcsncpy(g_currentPlaybackDefault, pv.pwszVal, 255);
                /* Only set as default if no config loaded */
                if (g_savedPlaybackDefault[0] == L'\0') {
                    wcsncpy(g_playbackDefault, pv.pwszVal, 255);
                }
                PropVariantClear(&pv);
            }
            IPropertyStore_Release(pStore);
        }
        IMMDevice_Release(pDefault);
    }
    if (SUCCEEDED(IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eRender, eCommunications, &pDefault))) {
        IPropertyStore* pStore = NULL;
        if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDefault, STGM_READ, &pStore))) {
            PROPVARIANT pv;
            PropVariantInit(&pv);
            if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                wcsncpy(g_currentPlaybackComm, pv.pwszVal, 255);
                if (g_savedPlaybackComm[0] == L'\0') {
                    wcsncpy(g_playbackComm, pv.pwszVal, 255);
                }
                PropVariantClear(&pv);
            }
            IPropertyStore_Release(pStore);
        }
        IMMDevice_Release(pDefault);
    }
    if (SUCCEEDED(IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eCapture, eConsole, &pDefault))) {
        IPropertyStore* pStore = NULL;
        if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDefault, STGM_READ, &pStore))) {
            PROPVARIANT pv;
            PropVariantInit(&pv);
            if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                wcsncpy(g_currentRecordDefault, pv.pwszVal, 255);
                if (g_savedRecordDefault[0] == L'\0') {
                    wcsncpy(g_recordDefault, pv.pwszVal, 255);
                }
                PropVariantClear(&pv);
            }
            IPropertyStore_Release(pStore);
        }
        IMMDevice_Release(pDefault);
    }
    if (SUCCEEDED(IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eCapture, eCommunications, &pDefault))) {
        IPropertyStore* pStore = NULL;
        if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDefault, STGM_READ, &pStore))) {
            PROPVARIANT pv;
            PropVariantInit(&pv);
            if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                wcsncpy(g_currentRecordComm, pv.pwszVal, 255);
                if (g_savedRecordComm[0] == L'\0') {
                    wcsncpy(g_recordComm, pv.pwszVal, 255);
                }
                PropVariantClear(&pv);
            }
            IPropertyStore_Release(pStore);
        }
        IMMDevice_Release(pDefault);
    }
    
    IMMDeviceEnumerator_Release(pEnum);
    CoUninitialize();
}

/* ========== GUI Dialog Procedure ========== */
static LRESULT CALLBACK ConfigDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hComboPlaybackDefault, hComboRecordDefault, hComboPlaybackComm, hComboRecordComm;
    static HWND hEditFolder, hButtonSave, hButtonRun, hButtonBrowse;
    static HBRUSH hBrushBg, hBrushEdit;
    static HFONT hFont, hFontBold, hFontSmall;
    
    switch (msg) {
        case WM_CREATE: {
            /* Dark theme colors */
            hBrushBg = CreateSolidBrush(RGB(30, 30, 30));
            hBrushEdit = CreateSolidBrush(RGB(45, 45, 45));
            hFont = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                               DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            hFontBold = CreateFontW(20, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                   OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            hFontSmall = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                    DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            
            int yPos = 15;
            
            /* Header */
            HWND hHeader = CreateWindowW(L"STATIC", L"Device Selection",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 25, hwnd, (HMENU)ID_HEADER, NULL, NULL);
            SendMessage(hHeader, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            yPos += 30;
            
            /* Subheader */
            HWND hSubheader = CreateWindowW(L"STATIC", 
                L"Select your audio devices and choose where to install.",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_DESC_START), NULL, NULL);
            SendMessage(hSubheader, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            yPos += 30;
            
            /* Install Folder */
            HWND hLabelFolder = CreateWindowW(L"STATIC", L"Install Folder",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_LABEL_START), NULL, NULL);
            SendMessage(hLabelFolder, WM_SETFONT, (WPARAM)hFont, TRUE);
            HWND hDescFolder = CreateWindowW(L"STATIC", L"Change this to move everything to a new folder",
                WS_CHILD | WS_VISIBLE, 20, yPos + 18, 460, 18, hwnd, (HMENU)(ID_DESC_START + 5), NULL, NULL);
            SendMessage(hDescFolder, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            hEditFolder = CreateWindowW(L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 20, yPos + 40, 370, 25,
                hwnd, (HMENU)ID_EDIT_FOLDER, NULL, NULL);
            SendMessage(hEditFolder, WM_SETFONT, (WPARAM)hFont, TRUE);
            hButtonBrowse = CreateWindowW(L"BUTTON", L"Browse...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 400, yPos + 38, 80, 28,
                hwnd, (HMENU)ID_BUTTON_BROWSE, NULL, NULL);
            SendMessage(hButtonBrowse, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 75;
            
            /* Default Playback Device */
            HWND hLabel1 = CreateWindowW(L"STATIC", L"Default Playback Device",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_LABEL_START + 1), NULL, NULL);
            SendMessage(hLabel1, WM_SETFONT, (WPARAM)hFont, TRUE);
            HWND hDesc1 = CreateWindowW(L"STATIC", L"Where your system audio plays (games, music, videos, notifications)",
                WS_CHILD | WS_VISIBLE, 20, yPos + 18, 460, 18, hwnd, (HMENU)(ID_DESC_START + 1), NULL, NULL);
            SendMessage(hDesc1, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            hComboPlaybackDefault = CreateWindowW(L"COMBOBOX", NULL,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL, 20, yPos + 40, 460, 200,
                hwnd, (HMENU)ID_COMBO_PLAYBACK_DEFAULT, NULL, NULL);
            SendMessage(hComboPlaybackDefault, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 75;
            
            /* Default Recording Device */
            HWND hLabel2 = CreateWindowW(L"STATIC", L"Default Recording Device",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_LABEL_START + 2), NULL, NULL);
            SendMessage(hLabel2, WM_SETFONT, (WPARAM)hFont, TRUE);
            HWND hDesc2 = CreateWindowW(L"STATIC", L"Default microphone for apps that don't specify one",
                WS_CHILD | WS_VISIBLE, 20, yPos + 18, 460, 18, hwnd, (HMENU)(ID_DESC_START + 2), NULL, NULL);
            SendMessage(hDesc2, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            hComboRecordDefault = CreateWindowW(L"COMBOBOX", NULL,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL, 20, yPos + 40, 460, 200,
                hwnd, (HMENU)ID_COMBO_RECORD_DEFAULT, NULL, NULL);
            SendMessage(hComboRecordDefault, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 75;
            
            /* Communications Playback Device */
            HWND hLabel3 = CreateWindowW(L"STATIC", L"Communications Playback Device",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_LABEL_START + 3), NULL, NULL);
            SendMessage(hLabel3, WM_SETFONT, (WPARAM)hFont, TRUE);
            HWND hDesc3 = CreateWindowW(L"STATIC", L"Where you hear voice chat (Discord, Teams, Zoom, etc.)",
                WS_CHILD | WS_VISIBLE, 20, yPos + 18, 460, 18, hwnd, (HMENU)(ID_DESC_START + 3), NULL, NULL);
            SendMessage(hDesc3, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            hComboPlaybackComm = CreateWindowW(L"COMBOBOX", NULL,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL, 20, yPos + 40, 460, 200,
                hwnd, (HMENU)ID_COMBO_PLAYBACK_COMM, NULL, NULL);
            SendMessage(hComboPlaybackComm, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 75;
            
            /* Communications Recording Device */
            HWND hLabel4 = CreateWindowW(L"STATIC", L"Communications Recording Device",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 20, hwnd, (HMENU)(ID_LABEL_START + 4), NULL, NULL);
            SendMessage(hLabel4, WM_SETFONT, (WPARAM)hFont, TRUE);
            HWND hDesc4 = CreateWindowW(L"STATIC", L"Microphone used for voice chat (Discord, Teams, Zoom, etc.)",
                WS_CHILD | WS_VISIBLE, 20, yPos + 18, 460, 18, hwnd, (HMENU)(ID_DESC_START + 4), NULL, NULL);
            SendMessage(hDesc4, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            hComboRecordComm = CreateWindowW(L"COMBOBOX", NULL,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL, 20, yPos + 40, 460, 200,
                hwnd, (HMENU)ID_COMBO_RECORD_COMM, NULL, NULL);
            SendMessage(hComboRecordComm, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 75;
            
            /* Checkboxes row - two checkboxes side by side */
            HWND hCheckRunBackground = CreateWindowW(L"BUTTON", L"Run in background",
                WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 20, yPos, 200, 20,
                hwnd, (HMENU)ID_CHECK_RUN_BACKGROUND, NULL, NULL);
            SendMessage(hCheckRunBackground, WM_SETFONT, (WPARAM)hFont, TRUE);
            if (g_runInBackground) {
                SendMessage(hCheckRunBackground, BM_SETCHECK, BST_CHECKED, 0);
            }
            
            HWND hCheckShowNotify = CreateWindowW(L"BUTTON", L"Show completion notification",
                WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 225, yPos, 250, 20,
                hwnd, (HMENU)ID_CHECK_SHOW_NOTIFY, NULL, NULL);
            SendMessage(hCheckShowNotify, WM_SETFONT, (WPARAM)hFont, TRUE);
            if (g_showNotification) {
                SendMessage(hCheckShowNotify, BM_SETCHECK, BST_CHECKED, 0);
            }
            yPos += 40;
            
            /* Buttons centered with gap between them */
            /* Window is 520px wide, buttons are 120px each, 20px gap = 260px total */
            /* Center: (520 - 260) / 2 = 130px from left */
            
            /* Save button - saves config only */
            hButtonSave = CreateWindowW(L"BUTTON", L"Save",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 130, yPos, 120, 35,
                hwnd, (HMENU)ID_BUTTON_SAVE, NULL, NULL);
            SendMessage(hButtonSave, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
            /* Run button - disabled until config exists */
            hButtonRun = CreateWindowW(L"BUTTON", L"Fix Audio",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | (g_configExists ? 0 : WS_DISABLED), 
                270, yPos, 120, 35,
                hwnd, (HMENU)ID_BUTTON_RUN, NULL, NULL);
            SendMessage(hButtonRun, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
            /* Populate combos with playback devices */
            for (int i = 0; i < g_playbackDeviceCount; i++) {
                SendMessageW(hComboPlaybackDefault, CB_ADDSTRING, 0, (LPARAM)g_playbackDeviceNames[i]);
                SendMessageW(hComboPlaybackComm, CB_ADDSTRING, 0, (LPARAM)g_playbackDeviceNames[i]);
            }
            /* Populate combos with recording devices */
            for (int i = 0; i < g_recordDeviceCount; i++) {
                SendMessageW(hComboRecordDefault, CB_ADDSTRING, 0, (LPARAM)g_recordDeviceNames[i]);
                SendMessageW(hComboRecordComm, CB_ADDSTRING, 0, (LPARAM)g_recordDeviceNames[i]);
            }
            
            /* Select current defaults */
            int idx;
            idx = (int)SendMessageW(hComboPlaybackDefault, CB_FINDSTRINGEXACT, -1, (LPARAM)g_playbackDefault);
            SendMessage(hComboPlaybackDefault, CB_SETCURSEL, idx >= 0 ? idx : 0, 0);
            idx = (int)SendMessageW(hComboPlaybackComm, CB_FINDSTRINGEXACT, -1, (LPARAM)g_playbackComm);
            SendMessage(hComboPlaybackComm, CB_SETCURSEL, idx >= 0 ? idx : 0, 0);
            idx = (int)SendMessageW(hComboRecordDefault, CB_FINDSTRINGEXACT, -1, (LPARAM)g_recordDefault);
            SendMessage(hComboRecordDefault, CB_SETCURSEL, idx >= 0 ? idx : 0, 0);
            idx = (int)SendMessageW(hComboRecordComm, CB_FINDSTRINGEXACT, -1, (LPARAM)g_recordComm);
            SendMessage(hComboRecordComm, CB_SETCURSEL, idx >= 0 ? idx : 0, 0);
            
            /* Always show current exe directory in the folder field */
            /* If install dir was set via env var, use that; otherwise use exe dir */
            if (g_installDir[0] == '\0' && g_exeDir[0] != '\0') {
                strncpy(g_installDir, g_exeDir, MAX_PATH);
            }
            if (g_installDir[0] != '\0') {
                WCHAR wPath[MAX_PATH];
                MultiByteToWideChar(CP_UTF8, 0, g_installDir, -1, wPath, MAX_PATH);
                SetWindowTextW(hEditFolder, wPath);
            } else if (g_exeDir[0] != '\0') {
                WCHAR wPath[MAX_PATH];
                MultiByteToWideChar(CP_UTF8, 0, g_exeDir, -1, wPath, MAX_PATH);
                SetWindowTextW(hEditFolder, wPath);
                strncpy(g_installDir, g_exeDir, MAX_PATH);
            }
            
            return 0;
        }
        
        case WM_CTLCOLORSTATIC: {
            HDC hdc = (HDC)wParam;
            HWND hCtrl = (HWND)lParam;
            int ctrlId = GetDlgCtrlID(hCtrl);
            SetBkColor(hdc, RGB(30, 30, 30));
            /* Title labels (ID 150-199) get blue text */
            if (ctrlId == ID_HEADER || (ctrlId >= ID_LABEL_START && ctrlId < ID_DESC_START)) {
                SetTextColor(hdc, RGB(100, 200, 255));
            }
            /* Description labels (ID >= 200) get gray text */
            else if (ctrlId >= ID_DESC_START) {
                SetTextColor(hdc, RGB(140, 140, 140));
            }
            /* Everything else (header, etc.) gets white */
            else {
                SetTextColor(hdc, RGB(255, 255, 255));
            }
            return (LRESULT)hBrushBg;
        }
        
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX: {
            HDC hdc = (HDC)wParam;
            SetBkColor(hdc, RGB(45, 45, 45));
            SetTextColor(hdc, RGB(255, 255, 255));
            return (LRESULT)hBrushEdit;
        }
        
        case WM_COMMAND:
            if (LOWORD(wParam) == ID_BUTTON_BROWSE) {
                /* Show folder browser dialog */
                BROWSEINFOW bi = {0};
                bi.hwndOwner = hwnd;
                bi.lpszTitle = L"Select Install Folder (optional)";
                bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
                
                LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
                if (pidl) {
                    WCHAR wPath[MAX_PATH];
                    if (SHGetPathFromIDListW(pidl, wPath)) {
                        /* Store in global */
                        WideCharToMultiByte(CP_UTF8, 0, wPath, -1, g_installDir, MAX_PATH, NULL, NULL);
                        /* Update edit box */
                        SetWindowTextW(hEditFolder, wPath);
                    }
                    CoTaskMemFree(pidl);
                }
                return 0;
            }
            else if (LOWORD(wParam) == ID_BUTTON_SAVE) {
                /* Save button - saves config, enables Run, but doesn't close */
                int idx = (int)SendMessage(hComboPlaybackDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackDefault, CB_GETLBTEXT, idx, (LPARAM)g_playbackDefault);
                
                idx = (int)SendMessage(hComboPlaybackComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackComm, CB_GETLBTEXT, idx, (LPARAM)g_playbackComm);
                
                idx = (int)SendMessage(hComboRecordDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordDefault, CB_GETLBTEXT, idx, (LPARAM)g_recordDefault);
                
                idx = (int)SendMessage(hComboRecordComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordComm, CB_GETLBTEXT, idx, (LPARAM)g_recordComm);
                
                /* Get install folder from edit box */
                WCHAR wInstallDir[MAX_PATH];
                GetWindowTextW(hEditFolder, wInstallDir, MAX_PATH);
                WideCharToMultiByte(CP_UTF8, 0, wInstallDir, -1, g_installDir, MAX_PATH, NULL, NULL);
                
                /* Get checkbox states */
                HWND hCheckBg = GetDlgItem(hwnd, ID_CHECK_RUN_BACKGROUND);
                g_runInBackground = (SendMessage(hCheckBg, BM_GETCHECK, 0, 0) == BST_CHECKED);
                HWND hCheckNotify = GetDlgItem(hwnd, ID_CHECK_SHOW_NOTIFY);
                g_showNotification = (SendMessage(hCheckNotify, BM_GETCHECK, 0, 0) == BST_CHECKED);
                
                /* Save config (moves files to install folder if path changed) */
                saveConfig();
                g_configExists = 1;
                
                /* Enable the Run button now that config exists */
                EnableWindow(hButtonRun, TRUE);
                
                /* Show success message with exe path */
                char exePath[MAX_PATH];
                if (g_installDir[0] != '\0') {
                    snprintf(exePath, MAX_PATH, "%s\\elgato_audio_reset.exe", g_installDir);
                } else {
                    strncpy(exePath, g_currentExePath, MAX_PATH);
                }
                
                char msg[1024];
                snprintf(msg, sizeof(msg), 
                    "Configuration saved! Run via the \"Fix Audio\" button.\n\n"
                    "Executable saved:\n%s\n\n"
                    "You can assign the executable to a hotkey or macro button. "
                    "If you set the config to run in the background with no notification, "
                    "your audio will reset silently within a few seconds on key press.",
                    exePath);
                MessageBoxA(hwnd, msg, "Elgato Audio Reset", MB_OK | MB_ICONINFORMATION);
                
                /* Don't close - user can click Run or X */
            }
            else if (LOWORD(wParam) == ID_BUTTON_RUN) {
                /* Run button - saves config and runs the reset */
                int idx = (int)SendMessage(hComboPlaybackDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackDefault, CB_GETLBTEXT, idx, (LPARAM)g_playbackDefault);
                
                idx = (int)SendMessage(hComboPlaybackComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackComm, CB_GETLBTEXT, idx, (LPARAM)g_playbackComm);
                
                idx = (int)SendMessage(hComboRecordDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordDefault, CB_GETLBTEXT, idx, (LPARAM)g_recordDefault);
                
                idx = (int)SendMessage(hComboRecordComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordComm, CB_GETLBTEXT, idx, (LPARAM)g_recordComm);
                
                /* Get install folder from edit box */
                WCHAR wInstallDir[MAX_PATH];
                GetWindowTextW(hEditFolder, wInstallDir, MAX_PATH);
                WideCharToMultiByte(CP_UTF8, 0, wInstallDir, -1, g_installDir, MAX_PATH, NULL, NULL);
                
                /* Get checkbox states */
                HWND hCheckBg = GetDlgItem(hwnd, ID_CHECK_RUN_BACKGROUND);
                g_runInBackground = (SendMessage(hCheckBg, BM_GETCHECK, 0, 0) == BST_CHECKED);
                HWND hCheckNotify = GetDlgItem(hwnd, ID_CHECK_SHOW_NOTIFY);
                g_showNotification = (SendMessage(hCheckNotify, BM_GETCHECK, 0, 0) == BST_CHECKED);
                
                /* Save config (moves files to install folder if path changed) */
                saveConfig();
                
                /* Set flag to run the reset after GUI closes */
                g_shouldRun = 1;
                
                DestroyWindow(hwnd);
            }
            return 0;
        
        case WM_ERASEBKGND: {
            HDC hdc = (HDC)wParam;
            RECT rc;
            GetClientRect(hwnd, &rc);
            FillRect(hdc, &rc, hBrushBg);
            return 1;
        }
        
        case WM_CLOSE:
            /* User closed via X - just close the window */
            DestroyWindow(hwnd);
            return 0;
        
        case WM_DESTROY:
            DeleteObject(hBrushBg);
            DeleteObject(hBrushEdit);
            DeleteObject(hFont);
            DeleteObject(hFontBold);
            DeleteObject(hFontSmall);
            PostQuitMessage(0);
            return 0;
    }
    
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

/* ========== Check for Config Mismatch ========== */
static int hasConfigMismatch(void) {
    if (g_savedPlaybackDefault[0] == L'\0') return 0; /* No saved config */
    
    if (wcscmp(g_savedPlaybackDefault, g_currentPlaybackDefault) != 0) return 1;
    if (wcscmp(g_savedPlaybackComm, g_currentPlaybackComm) != 0) return 1;
    if (wcscmp(g_savedRecordDefault, g_currentRecordDefault) != 0) return 1;
    if (wcscmp(g_savedRecordComm, g_currentRecordComm) != 0) return 1;
    
    return 0;
}

/* ========== Mismatch Dialog IDs ========== */
#define ID_MISMATCH_RESTORE  501
#define ID_MISMATCH_KEEP     502
#define ID_MISMATCH_TITLE    503
#define ID_MISMATCH_SAVED_HDR    504
#define ID_MISMATCH_CURRENT_HDR  505
#define ID_MISMATCH_LABEL_BASE   510  /* 510-517 for blue labels */

static int g_mismatchResult = 0;

/* ========== Mismatch Dialog Procedure ========== */
static LRESULT CALLBACK MismatchDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HFONT hFont, hFontBold, hFontDevice;
    static HBRUSH hBrushBg;
    
    switch (msg) {
        case WM_CREATE: {
            hBrushBg = CreateSolidBrush(RGB(30, 30, 30));
            hFont = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                               DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            hFontBold = CreateFontW(18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                   OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            hFontDevice = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                     OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                     DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            
            int yPos = 15;
            int colWidth = 380;
            int leftCol = 20;
            int rightCol = 420;
            int labelHeight = 18;
            int deviceHeight = 18;
            int rowSpacing = 12;
            
            /* Title */
            HWND hTitle = CreateWindowW(L"STATIC", 
                L"Your saved audio configuration differs from current Windows settings.",
                WS_CHILD | WS_VISIBLE | SS_CENTER, 20, yPos, 760, 20, hwnd, (HMENU)ID_MISMATCH_TITLE, NULL, NULL);
            SendMessage(hTitle, WM_SETFONT, (WPARAM)hFont, TRUE);
            yPos += 40;
            
            /* Column headers */
            HWND hSavedHeader = CreateWindowW(L"STATIC", L"Saved Configuration",
                WS_CHILD | WS_VISIBLE | SS_CENTER, leftCol, yPos, colWidth, 24, hwnd, (HMENU)ID_MISMATCH_SAVED_HDR, NULL, NULL);
            SendMessage(hSavedHeader, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
            HWND hCurrentHeader = CreateWindowW(L"STATIC", L"Current Windows Settings",
                WS_CHILD | WS_VISIBLE | SS_CENTER, rightCol, yPos, colWidth, 24, hwnd, (HMENU)ID_MISMATCH_CURRENT_HDR, NULL, NULL);
            SendMessage(hCurrentHeader, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            yPos += 35;
            
            /* Row 1: Playback Default */
            CreateWindowW(L"STATIC", L"Playback Default:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 0), NULL, NULL);
            CreateWindowW(L"STATIC", L"Playback Default:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 1), NULL, NULL);
            yPos += labelHeight;
            
            HWND hSavedPD = CreateWindowW(L"STATIC", g_savedPlaybackDefault,
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hSavedPD, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            HWND hCurrentPD = CreateWindowW(L"STATIC", g_currentPlaybackDefault,
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hCurrentPD, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            yPos += deviceHeight + rowSpacing;
            
            /* Row 2: Playback Comms */
            CreateWindowW(L"STATIC", L"Playback Comms:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 2), NULL, NULL);
            CreateWindowW(L"STATIC", L"Playback Comms:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 3), NULL, NULL);
            yPos += labelHeight;
            
            HWND hSavedPC = CreateWindowW(L"STATIC", g_savedPlaybackComm,
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hSavedPC, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            HWND hCurrentPC = CreateWindowW(L"STATIC", g_currentPlaybackComm,
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hCurrentPC, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            yPos += deviceHeight + rowSpacing;
            
            /* Row 3: Recording Default */
            CreateWindowW(L"STATIC", L"Recording Default:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 4), NULL, NULL);
            CreateWindowW(L"STATIC", L"Recording Default:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 5), NULL, NULL);
            yPos += labelHeight;
            
            HWND hSavedRD = CreateWindowW(L"STATIC", g_savedRecordDefault,
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hSavedRD, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            HWND hCurrentRD = CreateWindowW(L"STATIC", g_currentRecordDefault,
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hCurrentRD, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            yPos += deviceHeight + rowSpacing;
            
            /* Row 4: Recording Comms */
            CreateWindowW(L"STATIC", L"Recording Comms:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 6), NULL, NULL);
            CreateWindowW(L"STATIC", L"Recording Comms:",
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, labelHeight, hwnd, (HMENU)(ID_MISMATCH_LABEL_BASE + 7), NULL, NULL);
            yPos += labelHeight;
            
            HWND hSavedRC = CreateWindowW(L"STATIC", g_savedRecordComm,
                WS_CHILD | WS_VISIBLE | SS_LEFT, leftCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hSavedRC, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            HWND hCurrentRC = CreateWindowW(L"STATIC", g_currentRecordComm,
                WS_CHILD | WS_VISIBLE | SS_LEFT, rightCol, yPos, colWidth, deviceHeight, hwnd, NULL, NULL, NULL);
            SendMessage(hCurrentRC, WM_SETFONT, (WPARAM)hFontDevice, TRUE);
            yPos += deviceHeight + 20;
            
            /* Buttons - centered in dialog */
            HWND hBtnRestore = CreateWindowW(L"BUTTON", L"Restore Saved",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 220, yPos, 150, 35,
                hwnd, (HMENU)ID_MISMATCH_RESTORE, NULL, NULL);
            SendMessage(hBtnRestore, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
            HWND hBtnKeep = CreateWindowW(L"BUTTON", L"Keep Current",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 430, yPos, 150, 35,
                hwnd, (HMENU)ID_MISMATCH_KEEP, NULL, NULL);
            SendMessage(hBtnKeep, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
            /* Apply font to blue labels */
            for (int i = 0; i < 8; i++) {
                HWND hLabel = GetDlgItem(hwnd, ID_MISMATCH_LABEL_BASE + i);
                if (hLabel) SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
            }
            
            return 0;
        }
        
        case WM_CTLCOLORSTATIC: {
            HDC hdc = (HDC)wParam;
            HWND hCtrl = (HWND)lParam;
            int ctrlId = GetDlgCtrlID(hCtrl);
            
            SetBkColor(hdc, RGB(30, 30, 30));
            
            /* Blue for headers and subheading labels */
            if (ctrlId == ID_MISMATCH_TITLE ||
                ctrlId == ID_MISMATCH_SAVED_HDR ||
                ctrlId == ID_MISMATCH_CURRENT_HDR ||
                (ctrlId >= ID_MISMATCH_LABEL_BASE && ctrlId <= ID_MISMATCH_LABEL_BASE + 7)) {
                SetTextColor(hdc, RGB(100, 200, 255));
            } else {
                /* White for device names */
                SetTextColor(hdc, RGB(255, 255, 255));
            }
            return (LRESULT)hBrushBg;
        }
        
        case WM_COMMAND:
            if (LOWORD(wParam) == ID_MISMATCH_RESTORE) {
                g_mismatchResult = 1;
                DestroyWindow(hwnd);
            } else if (LOWORD(wParam) == ID_MISMATCH_KEEP) {
                g_mismatchResult = 0;
                DestroyWindow(hwnd);
            }
            return 0;
        
        case WM_ERASEBKGND: {
            HDC hdc = (HDC)wParam;
            RECT rc;
            GetClientRect(hwnd, &rc);
            FillRect(hdc, &rc, hBrushBg);
            return 1;
        }
        
        case WM_CLOSE:
            g_mismatchResult = 0; /* Default to keep current on close */
            DestroyWindow(hwnd);
            return 0;
        
        case WM_DESTROY:
            DeleteObject(hBrushBg);
            DeleteObject(hFont);
            DeleteObject(hFontBold);
            DeleteObject(hFontDevice);
            PostQuitMessage(0);
            return 0;
    }
    
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

/* ========== Show Mismatch Dialog ========== */
/* Returns 1 if user chose "Restore Saved", 0 if "Keep Current" */
static int showMismatchDialog(void) {
    /* Register window class */
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = MismatchDlgProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"ElgatoMismatchDlg";
    RegisterClassW(&wc);
    
    /* Calculate center position */
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int winW = 820, winH = 380;
    int x = (screenW - winW) / 2;
    int y = (screenH - winH) / 2;
    
    /* Create window */
    HWND hwnd = CreateWindowExW(0, L"ElgatoMismatchDlg", L"Audio Configuration Mismatch",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x, y, winW, winH, NULL, NULL, GetModuleHandle(NULL), NULL);
    
    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);
    
    /* Message loop */
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    if (g_mismatchResult == 1) {
        /* User chose to restore saved defaults */
        wcsncpy(g_playbackDefault, g_savedPlaybackDefault, 256);
        wcsncpy(g_playbackComm, g_savedPlaybackComm, 256);
        wcsncpy(g_recordDefault, g_savedRecordDefault, 256);
        wcsncpy(g_recordComm, g_savedRecordComm, 256);
    } else {
        /* User chose to keep current Windows settings */
        wcsncpy(g_playbackDefault, g_currentPlaybackDefault, 256);
        wcsncpy(g_playbackComm, g_currentPlaybackComm, 256);
        wcsncpy(g_recordDefault, g_currentRecordDefault, 256);
        wcsncpy(g_recordComm, g_currentRecordComm, 256);
    }
    
    return g_mismatchResult;
}

/* ========== Show Config GUI ========== */
static void showConfigGUI(void) {
    /* Enumerate devices first */
    enumerateDevicesForGUI();
    
    /* Check for mismatch between saved config and current Windows settings */
    if (g_configExists && hasConfigMismatch()) {
        showMismatchDialog();
    }
    
    /* Register window class */
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = ConfigDlgProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"ElgatoResetConfig";
    RegisterClassW(&wc);
    
    /* Calculate center position */
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int winW = 520, winH = 600;
    int x = (screenW - winW) / 2;
    int y = (screenH - winH) / 2;
    
    /* Create window */
    WCHAR windowTitle[64];
    swprintf(windowTitle, 64, L"Elgato Audio Reset v%s", APP_VERSION);
    HWND hwnd = CreateWindowExW(0, L"ElgatoResetConfig", windowTitle,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x, y, winW, winH, NULL, NULL, GetModuleHandle(NULL), NULL);
    
    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);
    
    /* Message loop */
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

/* ========== Logging ========== */
static void logMsg(const char* fmt, ...) {
    va_list args;
    char buf[1024];
    
    va_start(args, fmt);
    vsnprintf(buf, sizeof(buf), fmt, args);
    va_end(args);
    
    printf("%s", buf);
    fflush(stdout);
    
    if (g_logFile) {
        fprintf(g_logFile, "%s", buf);
        fflush(g_logFile);
    }
}

static void initLog(const char* exePath) {
    SYSTEMTIME st;
    GetLocalTime(&st);
    
    char dir[MAX_PATH];
    char logDir[MAX_PATH];
    strncpy(dir, exePath, MAX_PATH);
    char* lastSlash = strrchr(dir, '\\');
    if (lastSlash) *lastSlash = '\0';
    
    /* Create logs subfolder */
    snprintf(logDir, MAX_PATH, "%s\\logs", dir);
    CreateDirectoryA(logDir, NULL);
    
    /* Month names for readable dates */
    const char* months[] = {"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
    
    /* Format: logs/ElgatoReset_05-Dec-2025_12-17-22.log */
    snprintf(g_logPath, MAX_PATH, "%s\\ElgatoReset_%02d-%s-%04d_%02d-%02d-%02d.log",
             logDir, st.wDay, months[st.wMonth], st.wYear, st.wHour, st.wMinute, st.wSecond);
    
    g_logFile = fopen(g_logPath, "w");
}

/* ========== Protected Processes ========== */
static const char* g_protected[] = {
    "svchost.exe", "audiodg.exe", "System", "Idle", "dwm.exe", "explorer.exe",
    "csrss.exe", "wininit.exe", "services.exe", "lsass.exe", "smss.exe",
    "winlogon.exe", "fontdrvhost.exe", "sihost.exe", "taskhostw.exe",
    "RuntimeBroker.exe", "ShellExperienceHost.exe", "SearchHost.exe",
    "ctfmon.exe", "conhost.exe", "dllhost.exe", "powershell.exe", "cmd.exe",
    "Code.exe", "devenv.exe", "elgato_reset.exe",
    NULL
};

static int isProtected(const char* name) {
    for (int i = 0; g_protected[i]; i++) {
        if (_stricmp(name, g_protected[i]) == 0) return 1;
    }
    return 0;
}

static int isElgatoProcess(const char* name) {
    return (strstr(name, "WaveLink") || strstr(name, "StreamDeck") || strstr(name, "Elgato"));
}

/* ========== Registry Path Discovery ========== */
static int findInstallPath(const char* appName, char* outPath, size_t outLen) {
    const char* regPaths[] = {
        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
        "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
        NULL
    };
    
    for (int r = 0; regPaths[r]; r++) {
        HKEY hKey;
        if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, regPaths[r], 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
            char subKeyName[256];
            DWORD subKeyLen;
            for (DWORD i = 0; ; i++) {
                subKeyLen = sizeof(subKeyName);
                if (RegEnumKeyExA(hKey, i, subKeyName, &subKeyLen, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
                    break;
                
                HKEY hSubKey;
                char fullPath[512];
                snprintf(fullPath, sizeof(fullPath), "%s\\%s", regPaths[r], subKeyName);
                
                if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, fullPath, 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
                    char displayName[256] = {0};
                    DWORD displayLen = sizeof(displayName);
                    RegQueryValueExA(hSubKey, "DisplayName", NULL, NULL, (LPBYTE)displayName, &displayLen);
                    
                    if (strstr(displayName, appName)) {
                        char installLoc[MAX_PATH] = {0};
                        DWORD installLen = sizeof(installLoc);
                        if (RegQueryValueExA(hSubKey, "InstallLocation", NULL, NULL, (LPBYTE)installLoc, &installLen) == ERROR_SUCCESS
                            && installLoc[0] != '\0') {  /* Make sure it's not empty! */
                            strncpy(outPath, installLoc, outLen);
                            RegCloseKey(hSubKey);
                            RegCloseKey(hKey);
                            return 1;
                        }
                    }
                    RegCloseKey(hSubKey);
                }
            }
            RegCloseKey(hKey);
        }
    }
    return 0;
}

static void discoverPaths(void) {
    char installPath[MAX_PATH];
    
    /* WaveLink */
    if (findInstallPath("Wave Link", installPath, MAX_PATH)) {
        snprintf(g_waveLinkPath, MAX_PATH, "%s\\WaveLink.exe", installPath);
        snprintf(g_waveLinkSEPath, MAX_PATH, "%s\\WaveLinkSE.exe", installPath);
    }
    if (!g_waveLinkPath[0] && GetFileAttributesA("C:\\Program Files\\Elgato\\WaveLink\\WaveLink.exe") != INVALID_FILE_ATTRIBUTES) {
        strcpy(g_waveLinkPath, "C:\\Program Files\\Elgato\\WaveLink\\WaveLink.exe");
        strcpy(g_waveLinkSEPath, "C:\\Program Files\\Elgato\\WaveLink\\WaveLinkSE.exe");
    }
    
    /* StreamDeck */
    if (findInstallPath("Stream Deck", installPath, MAX_PATH)) {
        snprintf(g_streamDeckPath, MAX_PATH, "%s\\StreamDeck.exe", installPath);
    }
    if (!g_streamDeckPath[0] && GetFileAttributesA("C:\\Program Files\\Elgato\\StreamDeck\\StreamDeck.exe") != INVALID_FILE_ATTRIBUTES) {
        strcpy(g_streamDeckPath, "C:\\Program Files\\Elgato\\StreamDeck\\StreamDeck.exe");
    }
    
    logMsg("[i] Discovered paths:\n");
    logMsg("    WaveLink: %s\n", g_waveLinkPath[0] ? g_waveLinkPath : "NOT FOUND");
    logMsg("    WaveLinkSE: %s\n", g_waveLinkSEPath[0] ? g_waveLinkSEPath : "NOT FOUND");
    logMsg("    StreamDeck: %s\n", g_streamDeckPath[0] ? g_streamDeckPath : "NOT FOUND");
}

/* ========== Process Functions ========== */
static void killElgatoProcesses(void) {
    logMsg("[i] Discovering and killing Elgato processes...\n");
    
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return;
    
    PROCESSENTRY32 pe;
    pe.dwSize = sizeof(pe);
    int killed = 0;
    
    if (Process32First(hSnap, &pe)) {
        do {
            if (!isProtected(pe.szExeFile) && isElgatoProcess(pe.szExeFile)) {
                HANDLE hProc = OpenProcess(PROCESS_TERMINATE, FALSE, pe.th32ProcessID);
                if (hProc) {
                    if (TerminateProcess(hProc, 1)) {
                        logMsg("    [+] Killed: %s (PID %lu)\n", pe.szExeFile, pe.th32ProcessID);
                        killed++;
                    }
                    CloseHandle(hProc);
                }
            }
        } while (Process32Next(hSnap, &pe));
    }
    CloseHandle(hSnap);
    
    if (killed == 0) {
        logMsg("    [i] No Elgato processes found to kill.\n");
    }
    
    /* Wait for processes to fully exit */
    Sleep(1000);
}

/* ========== Service Control ========== */
static int controlService(const char* svcName, int start) {
    SC_HANDLE hSCM = OpenSCManagerA(NULL, NULL, SC_MANAGER_CONNECT);
    if (!hSCM) return 0;
    
    SC_HANDLE hSvc = OpenServiceA(hSCM, svcName, SERVICE_START | SERVICE_STOP | SERVICE_QUERY_STATUS);
    if (!hSvc) {
        CloseServiceHandle(hSCM);
        return 0;
    }
    
    SERVICE_STATUS status;
    int result = 0;
    
    if (start) {
        result = StartServiceA(hSvc, 0, NULL);
    } else {
        result = ControlService(hSvc, SERVICE_CONTROL_STOP, &status);
    }
    
    CloseServiceHandle(hSvc);
    CloseServiceHandle(hSCM);
    return result || GetLastError() == ERROR_SERVICE_ALREADY_RUNNING;
}

static int isServiceRunning(const char* svcName) {
    SC_HANDLE hSCM = OpenSCManagerA(NULL, NULL, SC_MANAGER_CONNECT);
    if (!hSCM) return 0;
    
    SC_HANDLE hSvc = OpenServiceA(hSCM, svcName, SERVICE_QUERY_STATUS);
    if (!hSvc) {
        CloseServiceHandle(hSCM);
        return 0;
    }
    
    SERVICE_STATUS status;
    int running = 0;
    if (QueryServiceStatus(hSvc, &status)) {
        running = (status.dwCurrentState == SERVICE_RUNNING);
    }
    
    CloseServiceHandle(hSvc);
    CloseServiceHandle(hSCM);
    return running;
}

static void restartAudioServices(void) {
    logMsg("[i] Restarting audio services...\n");
    
    controlService("audiosrv", 0);
    controlService("AudioEndpointBuilder", 0);
    Sleep(1000);
    
    controlService("AudioEndpointBuilder", 1);
    controlService("audiosrv", 1);
    
    /* Wait for services to start */
    logMsg("[i] Waiting for audio services");
    for (int i = 0; i < MAX_SERVICE_WAIT; i++) {
        if (isServiceRunning("audiosrv") && isServiceRunning("AudioEndpointBuilder")) {
            logMsg("\n[+] Audio services running after %d sec.\n", i + 1);
            return;
        }
        logMsg(".");
        fflush(stdout);
        Sleep(1000);
    }
    logMsg("\n[!] WARNING: Audio services may not be running!\n");
}

/* ========== Launch Applications ========== */
/* Callback to find and minimize windows by process ID */
static DWORD g_targetPid = 0;
static BOOL CALLBACK minimizeWindowCallback(HWND hwnd, LPARAM lParam) {
    DWORD windowPid = 0;
    GetWindowThreadProcessId(hwnd, &windowPid);
    if (windowPid == g_targetPid && IsWindowVisible(hwnd)) {
        ShowWindow(hwnd, SW_MINIMIZE);
    }
    return TRUE; /* Continue enumeration */
}

static void minimizeProcessWindows(const char* exeName) {
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return;
    
    PROCESSENTRY32 pe;
    pe.dwSize = sizeof(pe);
    
    if (Process32First(hSnap, &pe)) {
        do {
            if (_stricmp(pe.szExeFile, exeName) == 0) {
                g_targetPid = pe.th32ProcessID;
                EnumWindows(minimizeWindowCallback, 0);
            }
        } while (Process32Next(hSnap, &pe));
    }
    CloseHandle(hSnap);
}

static int isProcessRunning(const char* exeName) {
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return 0;
    
    PROCESSENTRY32 pe;
    pe.dwSize = sizeof(pe);
    int found = 0;
    
    if (Process32First(hSnap, &pe)) {
        do {
            if (_stricmp(pe.szExeFile, exeName) == 0) {
                found = 1;
                break;
            }
        } while (Process32Next(hSnap, &pe));
    }
    CloseHandle(hSnap);
    return found;
}

static void launchApp(const char* path, const char* exeName, const char* friendlyName) {
    if (!path[0] || GetFileAttributesA(path) == INVALID_FILE_ATTRIBUTES) {
        logMsg("[!] %s not found.\n", friendlyName);
        return;
    }
    
    logMsg("[i] Starting %s (minimized)...\n", friendlyName);
    
    STARTUPINFOA si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_SHOWMINIMIZED;
    
    char cmdLine[MAX_PATH + 32];
    strcpy(cmdLine, path);
    
    if (CreateProcessA(NULL, cmdLine, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
        
        /* Wait for process to appear */
        logMsg("[i] Waiting for %s", friendlyName);
        for (int i = 0; i < 10; i++) {
            Sleep(2000);
            if (isProcessRunning(exeName)) {
                logMsg("\n[+] %s detected.\n", friendlyName);
                return;
            }
            logMsg(".");
            fflush(stdout);
        }
        logMsg("\n[!] %s may not have started properly.\n", friendlyName);
    } else {
        logMsg("[!] Failed to start %s (Error %lu)\n", friendlyName, GetLastError());
    }
}

static void waitForElgatoDevices(void) {
    logMsg("[i] Waiting for Elgato virtual devices");
    
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
        logMsg("\n[!] COM initialization failed.\n");
        return;
    }
    
    for (int elapsed = 0; elapsed < MAX_DEVICE_WAIT; elapsed += POLL_INTERVAL) {
        IMMDeviceEnumerator* pEnum = NULL;
        hr = CoCreateInstance(&MY_CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                              &MY_IID_IMMDeviceEnumerator, (void**)&pEnum);
        if (SUCCEEDED(hr)) {
            IMMDeviceCollection* pCol = NULL;
            hr = IMMDeviceEnumerator_EnumAudioEndpoints(pEnum, eRender, DEVICE_STATE_ACTIVE, &pCol);
            if (SUCCEEDED(hr)) {
                UINT count = 0;
                IMMDeviceCollection_GetCount(pCol, &count);
                
                int elgatoCount = 0;
                for (UINT i = 0; i < count; i++) {
                    IMMDevice* pDev = NULL;
                    if (SUCCEEDED(IMMDeviceCollection_Item(pCol, i, &pDev))) {
                        IPropertyStore* pStore = NULL;
                        if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDev, STGM_READ, &pStore))) {
                            PROPVARIANT pv;
                            PropVariantInit(&pv);
                            if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                                if (wcsstr(pv.pwszVal, L"Elgato")) {
                                    elgatoCount++;
                                }
                                PropVariantClear(&pv);
                            }
                            IPropertyStore_Release(pStore);
                        }
                        IMMDevice_Release(pDev);
                    }
                }
                IMMDeviceCollection_Release(pCol);
                
                if (elgatoCount >= 2) {
                    logMsg("\n[+] Elgato virtual devices ready (%d sec).\n", elapsed);
                    IMMDeviceEnumerator_Release(pEnum);
                    CoUninitialize();
                    return;
                }
            }
            IMMDeviceEnumerator_Release(pEnum);
        }
        
        logMsg(".");
        fflush(stdout);
        Sleep(POLL_INTERVAL * 1000);
    }
    
    logMsg("\n[!] Elgato devices not detected - proceeding anyway.\n");
    CoUninitialize();
}

/* ========== Audio Default Setting ========== */
static IMMDevice* findDeviceByName(IMMDeviceEnumerator* pEnum, const WCHAR* name, EDataFlow dataFlow) {
    IMMDeviceCollection* pCol = NULL;
    if (FAILED(IMMDeviceEnumerator_EnumAudioEndpoints(pEnum, dataFlow, DEVICE_STATE_ACTIVE, &pCol))) return NULL;
    
    UINT count = 0;
    IMMDeviceCollection_GetCount(pCol, &count);
    
    IMMDevice* result = NULL;
    for (UINT i = 0; i < count && !result; i++) {
        IMMDevice* pDev = NULL;
        if (SUCCEEDED(IMMDeviceCollection_Item(pCol, i, &pDev))) {
            IPropertyStore* pStore = NULL;
            if (SUCCEEDED(IMMDevice_OpenPropertyStore(pDev, STGM_READ, &pStore))) {
                PROPVARIANT pv;
                PropVariantInit(&pv);
                if (SUCCEEDED(IPropertyStore_GetValue(pStore, &PKEY_Device_FriendlyName, &pv)) && pv.pwszVal) {
                    if (_wcsicmp(pv.pwszVal, name) == 0) {
                        result = pDev;
                        pDev = NULL; /* Don't release, we're returning it */
                    }
                    PropVariantClear(&pv);
                }
                IPropertyStore_Release(pStore);
            }
            if (pDev) IMMDevice_Release(pDev);
        }
    }
    IMMDeviceCollection_Release(pCol);
    return result;
}

static int setDefaultDevice(IMMDeviceEnumerator* pEnum, IPolicyConfig* pPolicy, 
                            const WCHAR* name, EDataFlow dataFlow, ERole role) {
    IMMDevice* pDev = findDeviceByName(pEnum, name, dataFlow);
    if (!pDev) return 0;
    
    LPWSTR devId = NULL;
    IMMDevice_GetId(pDev, &devId);
    
    int ok = 0;
    if (devId) {
        ok = SUCCEEDED(pPolicy->lpVtbl->SetDefaultEndpoint(pPolicy, devId, role));
        CoTaskMemFree(devId);
    }
    
    IMMDevice_Release(pDev);
    return ok;
}

static int unmuteDevice(IMMDeviceEnumerator* pEnum, const WCHAR* name, EDataFlow dataFlow) {
    IMMDevice* pDev = findDeviceByName(pEnum, name, dataFlow);
    if (!pDev) return 0;
    
    IAudioEndpointVolume* pVol = NULL;
    HRESULT hr = IMMDevice_Activate(pDev, &MY_IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (void**)&pVol);
    
    int ok = 0;
    if (SUCCEEDED(hr) && pVol) {
        IAudioEndpointVolume_SetMute(pVol, FALSE, NULL);
        IAudioEndpointVolume_SetMasterVolumeLevelScalar(pVol, 1.0f, NULL);
        IAudioEndpointVolume_Release(pVol);
        ok = 1;
    }
    
    IMMDevice_Release(pDev);
    return ok;
}

/* Get volume of default playback device */
static float getDefaultPlaybackVolume(void) {
    float volume = -1.0f;
    
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) return volume;
    
    IMMDeviceEnumerator* pEnum = NULL;
    hr = CoCreateInstance(&MY_CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                          &MY_IID_IMMDeviceEnumerator, (void**)&pEnum);
    if (FAILED(hr)) {
        CoUninitialize();
        return volume;
    }
    
    IMMDevice* pDev = NULL;
    hr = IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eRender, eConsole, &pDev);
    if (SUCCEEDED(hr) && pDev) {
        IAudioEndpointVolume* pVol = NULL;
        hr = IMMDevice_Activate(pDev, &MY_IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (void**)&pVol);
        if (SUCCEEDED(hr) && pVol) {
            IAudioEndpointVolume_GetMasterVolumeLevelScalar(pVol, &volume);
            IAudioEndpointVolume_Release(pVol);
        }
        IMMDevice_Release(pDev);
    }
    
    IMMDeviceEnumerator_Release(pEnum);
    CoUninitialize();
    return volume;
}

/* Set volume of default playback device */
static void setDefaultPlaybackVolume(float volume) {
    if (volume < 0.0f || volume > 1.0f) return;
    
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) return;
    
    IMMDeviceEnumerator* pEnum = NULL;
    hr = CoCreateInstance(&MY_CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                          &MY_IID_IMMDeviceEnumerator, (void**)&pEnum);
    if (FAILED(hr)) {
        CoUninitialize();
        return;
    }
    
    IMMDevice* pDev = NULL;
    hr = IMMDeviceEnumerator_GetDefaultAudioEndpoint(pEnum, eRender, eConsole, &pDev);
    if (SUCCEEDED(hr) && pDev) {
        IAudioEndpointVolume* pVol = NULL;
        hr = IMMDevice_Activate(pDev, &MY_IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (void**)&pVol);
        if (SUCCEEDED(hr) && pVol) {
            IAudioEndpointVolume_SetMasterVolumeLevelScalar(pVol, volume, NULL);
            IAudioEndpointVolume_Release(pVol);
        }
        IMMDevice_Release(pDev);
    }
    
    IMMDeviceEnumerator_Release(pEnum);
    CoUninitialize();
}

/* Save current volume, set to safe level */
static void saveAndLowerVolume(void) {
    g_savedPlaybackVolume = getDefaultPlaybackVolume();
    if (g_savedPlaybackVolume >= 0) {
        logMsg("[i] Saved volume: %.0f%%, lowering to 20%% for safety\n", g_savedPlaybackVolume * 100);
        setDefaultPlaybackVolume(0.20f);
    }
}

/* Restore previously saved volume */
static void restoreVolume(void) {
    if (g_savedPlaybackVolume >= 0) {
        logMsg("[i] Restoring volume to %.0f%%\n", g_savedPlaybackVolume * 100);
        setDefaultPlaybackVolume(g_savedPlaybackVolume);
        g_savedPlaybackVolume = -1.0f;
    }
}

static void setAudioDefaults(void) {
    logMsg("[i] Setting audio defaults and volumes...\n");
    
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
        logMsg("[!] COM initialization failed.\n");
        return;
    }
    
    IMMDeviceEnumerator* pEnum = NULL;
    hr = CoCreateInstance(&MY_CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
                          &MY_IID_IMMDeviceEnumerator, (void**)&pEnum);
    if (FAILED(hr)) {
        logMsg("[!] Failed to create device enumerator.\n");
        CoUninitialize();
        return;
    }
    
    IPolicyConfig* pPolicy = NULL;
    hr = CoCreateInstance(&CLSID_PolicyConfigClient, NULL, CLSCTX_ALL,
                          &IID_IPolicyConfig, (void**)&pPolicy);
    if (FAILED(hr)) {
        logMsg("[!] Failed to create policy config client.\n");
        IMMDeviceEnumerator_Release(pEnum);
        CoUninitialize();
        return;
    }
    
    /* Set defaults: eRender = 0, eCapture = 1; eConsole = 0, eCommunications = 2 */
    if (setDefaultDevice(pEnum, pPolicy, g_playbackDefault, eRender, eConsole)) {
        logMsg("    [+] Playback default: %ls\n", g_playbackDefault);
    } else {
        logMsg("    [!] Playback default not found: %ls\n", g_playbackDefault);
    }
    
    if (setDefaultDevice(pEnum, pPolicy, g_playbackComm, eRender, eCommunications)) {
        logMsg("    [+] Playback comms: %ls\n", g_playbackComm);
    } else {
        logMsg("    [!] Playback comms not found: %ls\n", g_playbackComm);
    }
    
    if (setDefaultDevice(pEnum, pPolicy, g_recordDefault, eCapture, eConsole)) {
        logMsg("    [+] Recording default: %ls\n", g_recordDefault);
    } else {
        logMsg("    [!] Recording default not found: %ls\n", g_recordDefault);
    }
    
    if (setDefaultDevice(pEnum, pPolicy, g_recordComm, eCapture, eCommunications)) {
        logMsg("    [+] Recording comms: %ls\n", g_recordComm);
    } else {
        logMsg("    [!] Recording comms not found: %ls\n", g_recordComm);
    }
    
    /* Unmute and set volume */
    unmuteDevice(pEnum, OUTPUT_RAZER_CHAT, eRender);
    unmuteDevice(pEnum, OUTPUT_RAZER_GAME, eRender);
    unmuteDevice(pEnum, g_recordDefault, eCapture);
    
    pPolicy->lpVtbl->Release(pPolicy);
    IMMDeviceEnumerator_Release(pEnum);
    CoUninitialize();
    
    logMsg("[+] Audio defaults configured.\n");
}

/* ========== Check Admin ========== */
static int isAdmin(void) {
    BOOL isAdmin = FALSE;
    PSID adminGroup = NULL;
    SID_IDENTIFIER_AUTHORITY ntAuth = SECURITY_NT_AUTHORITY;
    
    if (AllocateAndInitializeSid(&ntAuth, 2, SECURITY_BUILTIN_DOMAIN_RID,
                                  DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &adminGroup)) {
        CheckTokenMembership(NULL, adminGroup, &isAdmin);
        FreeSid(adminGroup);
    }
    return isAdmin;
}

static void elevateAndRestart(const char* exePath) {
    SHELLEXECUTEINFOA sei = {0};
    sei.cbSize = sizeof(sei);
    sei.lpVerb = "runas";
    sei.lpFile = exePath;
    sei.nShow = SW_SHOWNORMAL;
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    
    if (ShellExecuteExA(&sei)) {
        WaitForSingleObject(sei.hProcess, INFINITE);
        CloseHandle(sei.hProcess);
    }
}

/* ========== System Tray Functions ========== */
static LRESULT CALLBACK TrayWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_TRAYICON) {
        if (lParam == WM_LBUTTONUP) {
            /* Left click - open config GUI */
            g_trayAction = 1;
        } else if (lParam == WM_RBUTTONUP) {
            /* Right click - show context menu */
            POINT pt;
            GetCursorPos(&pt);
            
            HMENU hMenu = CreatePopupMenu();
            AppendMenuW(hMenu, MF_STRING, ID_TRAY_OPEN, L"Open Settings");
            AppendMenuW(hMenu, MF_STRING, ID_TRAY_RUN, L"Run Reset");
            AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
            AppendMenuW(hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
            
            /* Required for menu to work properly */
            SetForegroundWindow(hwnd);
            
            UINT cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, 
                                       pt.x, pt.y, 0, hwnd, NULL);
            DestroyMenu(hMenu);
            
            if (cmd == ID_TRAY_OPEN) {
                g_trayAction = 1;  /* Open config */
            } else if (cmd == ID_TRAY_RUN) {
                g_trayAction = 2;  /* Run reset */
            } else if (cmd == ID_TRAY_EXIT) {
                g_trayAction = 3;  /* Exit */
            }
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void initTrayIcon(void) {
    /* Create hidden window for tray messages */
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = TrayWndProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.lpszClassName = L"ElgatoResetTray";
    RegisterClassW(&wc);
    
    g_trayHwnd = CreateWindowW(L"ElgatoResetTray", L"", 0, 0, 0, 0, 0, 
                                HWND_MESSAGE, NULL, wc.hInstance, NULL);
    
    /* Set up notification icon */
    g_nid.cbSize = sizeof(NOTIFYICONDATAW);
    g_nid.hWnd = g_trayHwnd;
    g_nid.uID = ID_TRAY_ICON;
    g_nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcscpy(g_nid.szTip, L"Elgato Audio Reset - Starting...");
    
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

static void updateTrayStatus(const wchar_t* status) {
    g_trayStatus = status;
    swprintf(g_nid.szTip, 64, L"Elgato Reset - %s", status);
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

static void removeTrayIcon(void) {
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
    if (g_trayHwnd) {
        DestroyWindow(g_trayHwnd);
        g_trayHwnd = NULL;
    }
}

/* Sleep while pumping messages (for tray icon responsiveness) */
static void sleepWithMessages(DWORD ms) {
    DWORD start = GetTickCount();
    while (GetTickCount() - start < ms) {
        MSG msg;
        while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        Sleep(50);  /* Small sleep between message checks */
    }
}

/* ========== Main ========== */
int main(int argc, char* argv[]) {
    char exePath[MAX_PATH];
    GetModuleFileNameA(NULL, exePath, MAX_PATH);
    
    /* Check for install dir environment variable (set by PowerShell installer) */
    char* envInstallDir = getenv("ELGATO_INSTALL_DIR");
    if (envInstallDir && envInstallDir[0]) {
        strncpy(g_installDir, envInstallDir, MAX_PATH);
    }
    
    /* Load config file - track if it exists for Run button state */
    g_configExists = loadConfig(exePath);
    if (!g_configExists || !g_runInBackground) {
        /* First run, config deleted, or user wants to see GUI */
        showConfigGUI();
        
        /* If user didn't click Fix Audio, exit without running reset */
        if (!g_shouldRun) {
            return 0;
        }
    } else {
        /* Running in background - show system tray icon */
        initTrayIcon();
    }
    
    /* Initialize log */
    initLog(exePath);
    
    SYSTEMTIME st;
    GetLocalTime(&st);
    logMsg("===== Elgato Reset %02d/%02d/%04d %02d:%02d:%02d =====\n",
           st.wDay, st.wMonth, st.wYear, st.wHour, st.wMinute, st.wSecond);
    
    /* Discover paths */
    discoverPaths();
    
    /* Save current volume and lower to safe level before reset */
    saveAndLowerVolume();
    
    /* Step 1: Kill Elgato processes */
    if (g_trayHwnd) updateTrayStatus(L"Stopping processes...");
    killElgatoProcesses();
    
    /* Step 2: Restart audio services */
    if (g_trayHwnd) updateTrayStatus(L"Restarting audio...");
    restartAudioServices();
    
    /* Step 3: Launch WaveLink */
    if (g_trayHwnd) updateTrayStatus(L"Starting WaveLink...");
    if (g_waveLinkSEPath[0] && GetFileAttributesA(g_waveLinkSEPath) != INVALID_FILE_ATTRIBUTES) {
        launchApp(g_waveLinkSEPath, "WaveLinkSE.exe", "WaveLinkSE");
    }
    if (g_waveLinkPath[0]) {
        launchApp(g_waveLinkPath, "WaveLink.exe", "WaveLink");
    }
    
    /* Step 4: Wait for Elgato devices */
    if (g_trayHwnd) updateTrayStatus(L"Waiting for devices...");
    waitForElgatoDevices();
    
    /* Step 5: Launch StreamDeck */
    if (g_streamDeckPath[0]) {
        if (g_trayHwnd) updateTrayStatus(L"Starting StreamDeck...");
        launchApp(g_streamDeckPath, "StreamDeck.exe", "StreamDeck");
        logMsg("[i] Waiting for StreamDeck to fully initialize...\n");
        
        /* Wait for StreamDeck window to appear (up to 30 seconds) */
        for (int i = 0; i < 30; i++) {
            if (g_trayHwnd) sleepWithMessages(1000); else Sleep(1000);
            HWND hwnd = FindWindowA(NULL, "Stream Deck");
            if (hwnd) {
                if (g_trayHwnd) sleepWithMessages(2000); else Sleep(2000);
                break;
            }
        }
        
        minimizeProcessWindows("StreamDeck.exe");
        logMsg("[i] StreamDeck minimized.\n");
    }
    
    /* Step 6: Set audio defaults */
    if (g_trayHwnd) updateTrayStatus(L"Setting audio defaults...");
    if (g_trayHwnd) sleepWithMessages(2000); else Sleep(2000);
    setAudioDefaults();
    
    /* Restore original volume */
    restoreVolume();
    
    /* Done */
    logMsg("\n[+] Reset complete!\n");
    logMsg("[i] Log saved to:\n    %s\n", g_logPath);
    
    if (g_logFile) fclose(g_logFile);
    
    /* Remove tray icon if shown */
    if (g_trayHwnd) {
        updateTrayStatus(L"Complete!");
        sleepWithMessages(500); /* Brief moment to show "Complete!" */
        removeTrayIcon();
    }
    
    /* Show completion notification if enabled */
    if (g_showNotification) {
        MessageBoxA(NULL, "Elgato Audio Reset Complete", "Elgato Audio Reset", MB_OK | MB_ICONINFORMATION);
    }
    
    /* Handle tray action if user clicked during reset */
    if (g_trayAction == 1) {
        /* Open config */
        showConfigGUI();
    } else if (g_trayAction == 3) {
        /* Exit was requested - already exiting */
    }
    /* g_trayAction == 2 (Run) doesn't need handling - reset already ran */
    
    return 0;
}
