/*
 * elgato_reset.c - Complete Elgato Audio Reset Tool
 * Kills audio processes, restarts audio services, relaunches WaveLink/StreamDeck,
 * and sets audio defaults.
 * 
 * Compile: cl /O2 elgato_reset.c
 */

#define INITGUID
#define COBJMACROS
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shellapi.h>
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
#define ID_LABEL_START             150  /* Title labels start at 150 */
#define ID_DESC_START              200  /* Description labels start at 200 */

/* ========== Device Lists for GUI ========== */
#define MAX_DEVICES 32
static WCHAR g_playbackDeviceNames[MAX_DEVICES][256];
static WCHAR g_recordDeviceNames[MAX_DEVICES][256];
static int g_playbackDeviceCount = 0;
static int g_recordDeviceCount = 0;
static char g_exeDir[MAX_PATH] = {0};

/* ========== Config File Path Helper ========== */
static void getConfigPath(char* configPath, size_t len) {
    snprintf(configPath, len, "%s\\config.txt", g_exeDir);
}

/* ========== Config File Loading ========== */
static int loadConfig(const char* exePath) {
    char configPath[MAX_PATH];
    
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
        } else if (strcmp(key, "PLAYBACK_COMM") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_playbackComm, 256);
        } else if (strcmp(key, "RECORD_DEFAULT") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_recordDefault, 256);
        } else if (strcmp(key, "RECORD_COMM") == 0) {
            MultiByteToWideChar(CP_UTF8, 0, value, -1, g_recordComm, 256);
        }
    }
    
    fclose(f);
    return 1;
}

/* ========== Save Config File ========== */
static void saveConfig(void) {
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
                wcsncpy(g_playbackDefault, pv.pwszVal, 255);
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
                wcsncpy(g_playbackComm, pv.pwszVal, 255);
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
                wcsncpy(g_recordDefault, pv.pwszVal, 255);
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
                wcsncpy(g_recordComm, pv.pwszVal, 255);
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
                L"These are your current default audio devices. No changes are needed\nif you're happy with your current input/output configuration.",
                WS_CHILD | WS_VISIBLE, 20, yPos, 460, 40, hwnd, (HMENU)(ID_DESC_START), NULL, NULL);
            SendMessage(hSubheader, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
            yPos += 55;
            
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
            yPos += 85;
            
            /* Save button */
            HWND hButton = CreateWindowW(L"BUTTON", L"Save",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 160, yPos, 180, 35,
                hwnd, (HMENU)ID_BUTTON_SAVE, NULL, NULL);
            SendMessage(hButton, WM_SETFONT, (WPARAM)hFontBold, TRUE);
            
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
            if (LOWORD(wParam) == ID_BUTTON_SAVE) {
                /* Get selections */
                int idx = (int)SendMessage(hComboPlaybackDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackDefault, CB_GETLBTEXT, idx, (LPARAM)g_playbackDefault);
                
                idx = (int)SendMessage(hComboPlaybackComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboPlaybackComm, CB_GETLBTEXT, idx, (LPARAM)g_playbackComm);
                
                idx = (int)SendMessage(hComboRecordDefault, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordDefault, CB_GETLBTEXT, idx, (LPARAM)g_recordDefault);
                
                idx = (int)SendMessage(hComboRecordComm, CB_GETCURSEL, 0, 0);
                if (idx >= 0) SendMessageW(hComboRecordComm, CB_GETLBTEXT, idx, (LPARAM)g_recordComm);
                
                /* Save config */
                saveConfig();
                
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
            /* User closed without saving - exit app */
            ExitProcess(0);
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

/* ========== Show Config GUI ========== */
static void showConfigGUI(void) {
    /* Enumerate devices first */
    enumerateDevicesForGUI();
    
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
    int winW = 520, winH = 500;
    int x = (screenW - winW) / 2;
    int y = (screenH - winH) / 2;
    
    /* Create window */
    HWND hwnd = CreateWindowExW(0, L"ElgatoResetConfig", L"Elgato Audio Reset",
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

/* ========== Main ========== */
int main(int argc, char* argv[]) {
    char exePath[MAX_PATH];
    GetModuleFileNameA(NULL, exePath, MAX_PATH);
    
    /* Set console title */
    SetConsoleTitleA("Elgato Audio Reset");
    
    /* Unbuffered output */
    setvbuf(stdout, NULL, _IONBF, 0);
    
    /* Load config file - if it doesn't exist, show GUI */
    int configExists = loadConfig(exePath);
    if (!configExists) {
        /* First run or config deleted - show device selection GUI */
        showConfigGUI();
    }
    
    /* Initialize log */
    initLog(exePath);
    
    SYSTEMTIME st;
    GetLocalTime(&st);
    logMsg("===== Elgato Reset %02d/%02d/%04d %02d:%02d:%02d =====\n",
           st.wDay, st.wMonth, st.wYear, st.wHour, st.wMinute, st.wSecond);
    
    /* Discover paths */
    discoverPaths();
    
    /* Step 1: Kill Elgato processes */
    killElgatoProcesses();
    
    /* Step 2: Restart audio services */
    restartAudioServices();
    
    /* Step 3: Launch WaveLink */
    if (g_waveLinkSEPath[0] && GetFileAttributesA(g_waveLinkSEPath) != INVALID_FILE_ATTRIBUTES) {
        launchApp(g_waveLinkSEPath, "WaveLinkSE.exe", "WaveLinkSE");
    }
    if (g_waveLinkPath[0]) {
        launchApp(g_waveLinkPath, "WaveLink.exe", "WaveLink");
    }
    
    /* Step 4: Wait for Elgato devices */
    waitForElgatoDevices();
    
    /* Step 5: Launch StreamDeck */
    if (g_streamDeckPath[0]) {
        launchApp(g_streamDeckPath, "StreamDeck.exe", "StreamDeck");
        logMsg("[i] Waiting for StreamDeck to fully initialize...\n");
        Sleep(5000); /* StreamDeck needs extra time to fully load */
        minimizeProcessWindows("StreamDeck.exe");
        logMsg("[i] StreamDeck minimized.\n");
    }
    
    /* Step 6: Set audio defaults */
    Sleep(2000); /* Let everything settle */
    setAudioDefaults();
    
    /* Done */
    logMsg("\n[+] Reset complete!\n");
    logMsg("[i] Log saved to:\n    %s\n", g_logPath);
    
    if (g_logFile) fclose(g_logFile);
    
    return 0;
}
