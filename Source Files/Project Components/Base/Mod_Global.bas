Attribute VB_Name = "Mod_Global"
Rem -> ***|*********************************************************************|***|
Rem -> ***|                                                                     |***|
Rem -> ***|                      ______                __                       |***|
Rem -> ***|                     | ____ \              |  |                      |***|
Rem -> ***|                     | |   \ \             |  |                      |***|
Rem -> ***|                     | |    \ \            |  |                      |***|
Rem -> ***|                     | |    / /            |  |                      |***|
Rem -> ***|                     | |___/ /     __      |  |______                |***|
Rem -> ***|                     |______/     (__)     |_________|               |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   _______________________________________________________________   |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   Author       : John Fawzy (Dominator Legend)                      |***|
Rem -> ***|   Email        : Dominator_Legand@Yahoo.com                         |***|
Rem -> ***|   Date         : 21/3/2006                                          |***|
Rem -> ***|   Copyrights   : Some of these function not written by me,          |***|
Rem -> ***|                  However, Contents of code must be intact without   |***|
Rem -> ***|                  Change, If this work will used for commercial      |***|
Rem -> ***|                  Purpose please inform me, if you like this code    |***|
Rem -> ***|                  Please Rate It, Thanks                             |***|
Rem -> ***|                                                                     |***|
Rem -> ***|*********************************************************************|***|
Option Explicit
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __              ___            ___          |###'
'###|        | ____ \              |  |             \  \          /  /          |###'
'###|        | |   \ \             |  |              \  \        /  /           |###'
'###|        | |    \ \            |  |               \  \      /  /            |###'
'###|        | |    / /            |  |                \  \    /  /             |###'
'###|        | |___/ /     __      |  |______    __     \  \__/  /              |###'
'###|        |______/     (__)     |_________|  (__)     \______/               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|  Global Variables   |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Global WindowsEnabled           As Boolean
        Global KeyHook                  As Long             'Rem -> Handle to the hook
        Global HookEnabled              As Boolean          'Rem -> Flag
        Global Hook                     As KBDHOOKSTRUCT    'Rem -> The hook object
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______|              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |______               |###'
'###|        | |    / /            |  |                 |  ______|              |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________|              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Enum      |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> registry root directory constants
        Enum RegistryHives
            HKEY_CLASSES_ROOT = &H80000000
            HKEY_CURRENT_CONFIG = &H80000005
            HKEY_CURRENT_USER = &H80000001
            HKEY_DYN_DATA = &H80000006
            HKEY_LOCAL_MACHINE = &H80000002
            HKEY_PERFORMANCE_DATA = &H80000004
            HKEY_USERS = &H80000003
        End Enum
        Enum RegistryLongTypes
            REG_BINARY = 3              ' Free form binary
            REG_DWORD = 4               ' 32-bit number
            REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
            REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
        End Enum
        Rem -> registry key constants
        Enum RegistryKeyAccess
            KEY_CREATE_LINK = &H20
            KEY_CREATE_SUB_KEY = &H4
            KEY_ENUMERATE_SUB_KEYS = &H8
            KEY_EVENT = &H1    '  Event contains key event record
            KEY_NOTIFY = &H10
            KEY_QUERY_VALUE = &H1
            KEY_SET_VALUE = &H2
            READ_CONTROL = &H20000
            STANDARD_RIGHTS_ALL = &H1F0000
            STANDARD_RIGHTS_REQUIRED = &HF0000
            SYNCHRONIZE = &H100000
            STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
            STANDARD_RIGHTS_READ = (READ_CONTROL)
            STANDARD_RIGHTS_WRITE = (READ_CONTROL)
            KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL + KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK) And (Not SYNCHRONIZE))
            KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
            KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
            KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
        End Enum
        Rem -> error codes returned
        Enum RegistryErrorCodes
            ERROR_ACCESS_DENIED = 5&
            ERROR_INVALID_PARAMETER = 87    '  dderror
            ERROR_MORE_DATA = 234           '  dderror
            ERROR_NO_MORE_ITEMS = 259
            ERROR_SUCCESS = 0&
        End Enum
        Rem -> The different privilages that can be set/unset
        Enum EnumNTSettings
            Rem -> items that can be disabled on the Lock Screen
            CHANGE_PASSWORD = 0
            LOCK_WORKSTATION = 1
            REGISTRY_TOOLS = 2
            TASK_MGR = 3
            Rem -> the tabs on the Display Properties dialog box
            DISP_APPEARANCE_PAGE = 4
            DISP_BACKGROUND_PAGE = 5
            DISP_CPL = 6
            DISP_SCREENSAVER = 7
            DISP_SETTINGS = 8
        End Enum
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______)              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |                     |###'
'###|        | |    / /            |  |                 | |                     |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________)              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Const     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Global Const SWP_HIDEWINDOW     As String = &H80
        Global Const SWP_SHOWWINDOW     As String = &H40
        Global Const SW_MINIMIZE        As Integer = 6
        Global Const VK_Escape          As String = &H1B    'Rem -> The Escape key
        Global Const VK_LCtrl           As String = &HA2    'Rem -> The Left CTRL key
        Global Const VK_RCtrl           As String = &HA3    'Rem -> Thr Right CTRL key
        Global Const VK_LWinKey         As String = &H5B    'Rem -> The Window's key on the left of the keyboard
        Global Const VK_RWinKey         As String = &H5C    'Rem -> The Window's key on the right of the keyboard
        Global Const VK_Tab             As String = &H9     'Rem -> The Alt-Tab
        Global Const KEYEVENTF_KEYUP    As String = &H2     'Rem -> The Realease left windows key
        Global Const WH_KeyBoard        As String = 13&     'Rem -> Tells Windows what we want to hook
        Global Const HC_Action          As String = 0&      'Rem -> Message received by our hook function
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                  ______________          |###'
'###|        | ____ \              |  |                |_____    _____|         |###'
'###|        | |   \ \             |  |                      |  |               |###'
'###|        | |    \ \            |  |                      |  |               |###'
'###|        | |    / /            |  |                      |  |               |###'
'###|        | |___/ /     __      |  |______    __          |  |               |###'
'###|        |______/     (__)     |_________|  (__)         |__|               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|     Global Type     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Type TagInitCommonControlsEx
            LngSize                                 As Long
            LngICC                                  As Long
        End Type
        Type OSVERSIONINFO
            dwOSVersionInfoSize                     As Long
            dwMajorVersion                          As Long
            dwMinorVersion                          As Long
            dwBuildNumber                           As Long
            dwPlatformId                            As Long
            szCSDVersion                            As String * 128
        End Type
        Type KBDHOOKSTRUCT
            VKCode                                  As Long
            ScanCode                                As Long
            Flags                                   As Long
            Time                                    As Long
            DwExtraInfo                             As Long
        End Type
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                      ______              |###'
'###|        | ____ \              |  |                    /  __  \             |###'
'###|        | |   \ \             |  |                   /  /  \  \            |###'
'###|        | |    \ \            |  |                  /  /____\  \           |###'
'###|        | |    / /            |  |                 /  /______\  \          |###'
'###|        | |___/ /     __      |  |______    __    /  /        \  \         |###'
'###|        |______/     (__)     |_________|  (__)  /__/          \__\        |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global API       |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Declare Function InitCommonControlsEx Lib _
                                "comctl32.dll" ( _
                                iccex As TagInitCommonControlsEx) As _
                                Boolean
        Rem -> Finds the first window in the queue with the caption matching
        Rem -> The specified null termimated string.
        Declare Function FindWindow Lib "user32" _
                                Alias "FindWindowA" _
                                (ByVal lpClassName As String, _
                                ByVal lpWindowName As String) _
                                As Long
        Rem -> Set the position of the specified window
        Declare Function SetWindowPos Lib "user32" _
                                (ByVal HWnd As Long, _
                                 ByVal hWndInsertAfter As Long, _
                                 ByVal X As Long, _
                                 ByVal Y As Long, _
                                 ByVal cx As Long, _
                                 ByVal cy As Long, _
                                 ByVal wFlags As Long) _
                                 As Long
        Rem -> The EnumProcesses function retrieves the process identifier for each
        Rem -> Process object in the system.
        Declare Function EnumProcesses Lib "PSAPI.DLL" _
                                (ByRef lpidProcess As Long, _
                                ByVal cb As Long, _
                                ByRef cbNeeded As Long) _
                                As Long
        Rem -> Opens an existing process object
        Declare Function OpenProcess Lib "kernel32.dll" _
                                (ByVal dwDesiredAccessas As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwProcId As Long) _
                                As Long
        Rem -> The EnumProcessModules function retrieves a handle for each module in
        Rem -> The specified process.
        Declare Function EnumProcessModules Lib "PSAPI.DLL" _
                                (ByVal hProcess As Long, _
                                ByRef lphModule As Long, _
                                ByVal cb As Long, _
                                ByRef cbNeeded As Long) _
                                As Long
        Rem -> The GetModuleFileName function retrieves the full path and filename for
        Rem -> the executable file containing the specified module.
        Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" _
                                (ByVal hProcess As Long, _
                                ByVal hModule As Long, _
                                ByVal ModuleName As String, _
                                ByVal nSize As Long) _
                                As Long
        Rem -> Destroyes the specified process
        Declare Function TerminateProcess Lib "kernel32" _
                                (ByVal hProcess As Long, _
                                ByVal uExitCode As Long) _
                                As Long
        Declare Function GetCurrentProcess Lib _
                                "kernel32" () _
                                As Long
        Rem -> This function closes an open object handle
        Declare Function CloseHandle Lib "kernel32.dll" _
                                (ByVal Handle As Long) _
                                As Long
        Rem -> get information about the current operating system
        Declare Function GetVersionEx Lib "kernel32" _
                                Alias "GetVersionExA" _
                                (ByRef lpVersionInformation As OSVERSIONINFO) _
                                As Long
        Declare Function ShowWindow Lib "user32" _
                                (ByVal HWnd As Long, _
                                ByVal nCmdShow As Long) _
                                As Long
        Declare Function FindWindowEx Lib "user32" _
                                Alias "FindWindowExA" _
                                (ByVal hWnd1 As Long, _
                                ByVal hWnd2 As Long, _
                                ByVal lpsz1 As String, _
                                ByVal lpsz2 As String) _
                                As Long
       Declare Function GetWindowThreadProcessId Lib "user32" _
                                (ByVal HWnd As Long, _
                                lpdwProcessId As Long) _
                                As Long
        Declare Sub keybd_event Lib "user32" _
                                (ByVal bVk As Byte, _
                                ByVal bScan As Byte, _
                                ByVal dwFlags As Long, _
                                ByVal DwExtraInfo As Long)
        Rem -> used to set a hotkey
        Declare Function RegisterHotKey Lib "user32" _
                                (ByVal HWnd As Long, _
                                ByVal id As Long, _
                                ByVal fsModifiers As Long, _
                                ByVal vk As Long) _
                                As Long
        Rem -> remove the specified registered hotkey for the specified window
        Declare Function UnregisterHotKey Lib "user32" _
                                (ByVal HWnd As Long, _
                                ByVal id As Long) _
                                As Long
        Rem -> create a new registry key
        Declare Function RegCreateKey Lib "advapi32.dll" _
                                Alias "RegCreateKeyA" _
                                (ByVal HKey As Long, _
                                ByVal lpSubKey As String, _
                                phkResult As Long) _
                                As Long
        Rem -> return or set system information
        Declare Function SystemParametersInfo Lib "user32" _
                                Alias "SystemParametersInfoA" _
                                (ByVal uAction As Long, _
                                ByVal uParam As Long, _
                                ByRef lpvParam As Any, _
                                ByVal fuWinIni As Long) _
                                As Long
        Rem -> close an open registry key
        Declare Function RegCloseKey Lib "advapi32.dll" _
                                (ByVal HKey As Long) _
                                As Long
        Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                                Alias "RegOpenKeyExA" _
                                (ByVal HKey As Long, _
                                ByVal lpSubKey As String, _
                                ByVal ulOptions As Long, _
                                ByVal samDesired As Long, _
                                phkResult As Long) _
                                As Long
        Declare Function RegSetValueEx Lib "advapi32.dll" _
                                Alias "RegSetValueExA" _
                                (ByVal HKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                lpData As Any, _
                                ByVal cbData As Long) _
                                As Long
        Declare Function EnumWindows Lib "user32.dll" _
                                (ByVal lpEnumFunc As Long, _
                                ByVal lParam As Long) _
                                As Long
        Declare Function GetWindowTextLength Lib "user32.dll" _
                                Alias "GetWindowTextLengthA" _
                                (ByVal HWnd As Long) _
                                As Long
        Declare Function GetWindowText Lib "user32.dll" _
                                Alias "GetWindowTextA" _
                                (ByVal HWnd As Long, _
                                ByVal lpString As String, _
                                ByVal nMaxCount As Long) _
                                As Long
        Declare Function EnableWindow Lib "user32.dll" _
                                (ByVal HWnd As Long, _
                                ByVal fEnable As Long) _
                                As Long
        Declare Function GetTickCount Lib "kernel32" () _
                                As Long
        Declare Function SetWindowsHookEx Lib "user32" _
                                Alias "SetWindowsHookExA" _
                                (ByVal idHook As Long, _
                                ByVal lpfn As Long, _
                                ByVal hmod As Long, _
                                ByVal dwThreadId As Long) _
                                As Long
        Declare Function UnhookWindowsHookEx Lib "user32" _
                                (ByVal hHook As Long) _
                                As Long
        Declare Function CallNextHookEx Lib "user32" _
                                (ByVal hHook As Long, _
                                ByVal nCode As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) _
                                As Long
        Declare Function GetAsyncKeyState Lib "user32" _
                                (ByVal vKey As Long) _
                                As Integer
        Declare Sub CopyMemory Lib "kernel32" _
                                Alias "RtlMoveMemory" _
                                (pDest As Any, _
                                pSource As Any, _
                                ByVal cb As Long)
        Declare Function IsWindowVisible Lib "user32" _
                                (ByVal HWnd As Long) _
                                As Long

