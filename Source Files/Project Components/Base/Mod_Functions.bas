Attribute VB_Name = "Mod_Functions"
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
Rem -> *************************************************************************************************************************************************
Rem -> WndProc Function Used To Enum Process
Public Function WndProc(ByVal HWnd As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim ProcessHWND                 As Long
    Dim BufferLen                   As Long
    Dim BufferText                  As String
    ProcessHWND = HWnd
    BufferLen = GetWindowTextLength(ProcessHWND) + 1
    BufferText = Space(BufferLen)
    GetWindowText ProcessHWND, BufferText, BufferLen
    BufferText = Left(BufferText, Len(BufferText) - 1)
    If (BufferText = "") Or (HWnd = Frm_Main.HWnd) Or (BufferText = Frm_Main.Caption) Or (BufferText = App.EXEName) Or (BufferText = App.Title) Or InStr(1, BufferText, "Locker - Microsoft Visual Basic") Then GoTo NextProcess
    Call EnableWindow(HWnd, WindowsEnabled)
    If (IsWindowVisible(HWnd)) Then Call ShowWindow(HWnd, SW_MINIMIZE)
NextProcess:
    WndProc = 1
End Function
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper sub procedure that will enable or disable the windows.
Public Sub EnableWindows(ByVal Enabled As Boolean, ByVal OwnerHwnd As Long)
    Call WindowsInterfaceController(Enabled)
    Call AutoRestartShell(Enabled)
    Call AltTabController(Enabled, OwnerHwnd)
    Call AntiTaskManagerController(Enabled)
    Call NTController(CHANGE_PASSWORD, Enabled)
    Call NTController(LOCK_WORKSTATION, Enabled)
    Call ShellController(Enabled)
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper function that will enable or disable the windows taskbar components.
Function WindowsInterfaceController(Enabled As Boolean)
    Dim HWnd As Long
    HWnd = FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString)      'Rem -> Get the handel of system tray to show or hide
    ShowWindow HWnd, IIf(Enabled, 5, 0)
    HWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)            'Rem -> Get the handel of progman to show or hide
    ShowWindow HWnd, IIf(Enabled, 5, 0)
End Function
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper function that will Minimize all windows.
Function MinimizeAllWindows()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Simulate that the user press (LWin key + D) = Show Desktop even our windows
    Call keybd_event(VK_LWinKey, 0, 0, 0)
    Call keybd_event(Asc("M"), 0, 0, 0): Call keybd_event(Asc("M"), 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(Asc("D"), 0, 0, 0): Call keybd_event(Asc("D"), 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_LWinKey, 0, KEYEVENTF_KEYUP, 0)
End Function
Rem -> *************************************************************************************************************************************************
Rem -> This will prevent the task manager from appearing
Rem -> when the user presses Ctrl+Alt+Del. This will also disable the
Rem -> Alt+Tab and the Ctrl+Esc key combinations.
Public Sub AntiTaskManagerController(Enabled As Boolean)
    On Error Resume Next
    If IsWinNT Then
        Call NTController(TASK_MGR, Enabled)                                'Rem -> control the task manager in registry
        If Enabled Then
            Close #1                                                        'Rem -> Close the Taskmgr.exe, so we can run task manager normally
        Else
            Dim TMHwnd              As Long
            Dim ProcID              As Long
            Dim ProcessName         As Long
            Dim RetVal              As Long
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> If task manager opened we closed it first as we can't open an opened file ;)
            TMHwnd = FindWindow("#32770", "Windows Task Manager")           'Rem -> Find the HWnd of task manager
            RetVal = GetWindowThreadProcessId(TMHwnd, ProcID)               'Rem -> Find the process id
            ProcessName = OpenProcess(&H1F0FFF, 0&, ProcID)                 'Rem -> Open the process
            RetVal = TerminateProcess(ProcessName, 0&)                      'Rem -> Terminate it back
            Open Environ("WinDir") & "\System32\Taskmgr.exe" For Input Lock Read Write As #1
        End If
    Else
        SystemParametersInfo 97, Enabled, Enabled, 0
    End If
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This will enable or disable the windows task manager. Please note that
Rem -> this procedure does not work on any Non-NT based system (win 9x)
Public Sub NTController(ByVal EnmPrivilage As EnumNTSettings, ByVal Enabled As Boolean)
    If Not IsWinNT Then Exit Sub
    Dim Command         As String   'holds the Value to open
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Get the text to for the registry value for the selected setting
    Select Case EnmPrivilage
        Case CHANGE_PASSWORD: Command = "DisableChangePassword"         'Rem -> Don't allow pasword change
        Case LOCK_WORKSTATION: Command = "DisableLockWorkStation"       'Rem -> Disabling locking of workstation
        Case REGISTRY_TOOLS: Command = "DisableRegistryTools"           'Rem -> Cancel the register tools, like regedit
        Case TASK_MGR: Command = "DisableTaskMgr"                       'Rem -> Cancel task manager
        Case DISP_APPEARANCE_PAGE: Command = "NoDispAppearancePage"     'Rem -> No Display properties page
        Case DISP_BACKGROUND_PAGE: Command = "NoDispBackgroundPage"     'Rem -> No Background properties page
        Case DISP_CPL: Command = "NoDispCPL"                            'Rem -> Don't Display CPLs
        Case DISP_SCREENSAVER: Command = "NoDispScrSavPage"             'Rem -> No Screen saver any more
        Case DISP_SETTINGS: Command = "NoDispSettingsPage"              'Rem -> No setting page any more
        Case Else: Exit Sub
    End Select
    If IsWinNT Then
        Call CreateRegLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", Command, Not Enabled)
        If IsW2000 Then Call CreateRegLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\System", Command, Not Enabled)
    End If
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This will enable or disable the Alt+Tab functionality for windows. The
Rem -> hWnd parameter is needed, because Alt+Tab must be re-directed to a window
Rem -> instead of the operating system. The parameter is also needed to remove the
Rem -> functionality.
Public Sub AltTabController(Enabled As Boolean, Optional ByVal OwnerHwnd As Long)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Holds the numeric/ascii value for the hotkey
    Rem -> If the code is compiled to an APP, this needs to be between 0 and 49151
    Rem -> If the code is compiled to a DLL, this nees to be between 49152 and 65535
    Const HOT_KEY       As Long = 9
    Static IntHotkeyId  As Integer  'holds the windows id for the hotkey
    Static HWnd         As Long     'holds a handle to the desktop window
    Dim RetVal          As Long     'holds any returned error value from an api call
    Dim blnOld          As Boolean  'holds whether or not the screensaver was already running
    Rem -> ~~~~~~~~~~~~~~~~~~~
    Rem -> Turn on/off Alt+Tab
    If Enabled Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Turn on the Alt+Tab functionality for windows
        If IsWinNT Then
            If IntHotkeyId <> 0 Then
                RetVal = UnregisterHotKey(HWnd, IntHotkeyId)
                If RetVal <> False Then
                    HWnd = 0
                    IntHotkeyId = 0
                End If
            End If
        Else
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> On non-NT based systems, we can fool the computer into disabling
            Rem -> Alt+Tab by telling it a screen saver is running
            RetVal = SystemParametersInfo(97, True, Enabled, 0)
        End If
    Else
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Turn off the Alt+Tab functionality for windows
        If IsWinNT Then
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> We cannot register a hotkey unless a handel to a window was specified
            If OwnerHwnd = 0 Then
                Exit Sub
            End If
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Remove any active hotkey
            If IntHotkeyId <> 0 Then
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> hotkey already active - try to disable
                RetVal = UnregisterHotKey(HWnd, IntHotkeyId)
                If RetVal <> False Then
                    HWnd = 0
                    IntHotkeyId = 0
                Else
                    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    Rem -> unable to remove hotkey or invalid hWnd
                    Exit Sub
                End If
            End If
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> get a handle to the desktop window
            HWnd = OwnerHwnd
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Register the hotkey to disable Alt+Tab
            RetVal = RegisterHotKey(HWnd, IntHotkeyId, &H1, HOT_KEY)
        Else
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> On non-NT based systems, we can fool the computer into disabling
            Rem -> Alt+Tab by telling it a screen saver is running
            RetVal = SystemParametersInfo(97, False, Enabled, 0)
        End If
    End If
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper sub procedure that help to controling the shell.
Public Sub ShellController(Enabled As Boolean)
    Rem -> create/destroy the windows shell
    If Enabled Then
        Rem -> Restart the explorer shell
        Call DestroyShell
        Call CreateShell
    Else
        Rem -> Stop all threads and instances of the shell from running
        Call DestroyShell
    End If
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This will stop the explorer.exe process from running and if this is a
Rem -> win NT based system, it will also temperorily disable the auto restart
Rem -> function
Public Sub DestroyShell()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> If this is windows NT, make sure that the shell will start automatically
    Rem -> should it unexpectadly shutdown
    Call AutoRestartShell(False)
    Call KillProcess("Explorer.exe")
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This will create the explorer.exe process, by lunching it from windows path.
Public Sub CreateShell()
    Rem -> NOTE: This will create the task bar if it has been destroyed
    Dim StrExplorerPath As String       'holds the complete file path to Exploere.exe
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> If this is windows NT, make sure that the shell will start automatically
    Rem -> should it unexpectadly shutdown
    Call AutoRestartShell(True)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> get the path to "explorer.exe" and run it to create the task bar
    StrExplorerPath = Environ("WinDir") & "\Explorer.exe"
    Shell StrExplorerPath
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This will stop the specified executable file from running if it is active.
Public Sub KillProcess(ByVal StrExeName As String)
    Dim RetVal                  As Long     'holds any returned error value from an api call
    Dim StrProcessName          As String
    Dim LngCbSize               As Long     'Specifies the size, In bytes, of the lpidProcess array
    Dim LngCbSizeReturned       As Long     'Receives the number of bytes returned
    Dim LngNumElements          As Long
    Dim LngProcessIds()         As Long
    Dim LngCbSize2              As Long
    Dim LngModules(1 To 200)    As Long
    Dim StrModuleName           As String
    Dim LngSize                 As Long
    Dim HWndProcess             As Long
    Dim LngCounter              As Long
    Dim StrProcName             As String
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Make sure something was passed
    If Trim(StrExeName) = "" Then
        Exit Sub
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Really needs To be 16, but Loop will increment prior to calling API
    LngCbSize = 8
    LngCbSizeReturned = 96
    Do While LngCbSize <= LngCbSizeReturned
        Rem -> ~~~~~~~~~~~~~~~~~
        Rem -> Increment lngSize
        LngCbSize = LngCbSize * 2
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Allocate Memory for Array
        ReDim LngProcessIds(LngCbSize / 4)
        Rem -> ~~~~~~~~~~~~~~~~
        Rem -> Get Process ID's
        RetVal = EnumProcesses(LngProcessIds(1), LngCbSize, LngCbSizeReturned)
    Loop
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Count number of processes returned
    LngNumElements = LngCbSizeReturned / 4
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Loop through each process
    For LngCounter = 1 To LngNumElements
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Get a handle to the Process and Open it
        HWndProcess = OpenProcess(&H1F0FFF, 0, LngProcessIds(LngCounter))
        If HWndProcess <> 0 Then
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Get an array of the module handles for the specified process
            RetVal = EnumProcessModules(HWndProcess, LngModules(1), 200, LngCbSize2)
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> If the Module Array is retrieved, Get the ModuleFileName
            If RetVal <> 0 Then
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Prepare buffer to hold module name
                StrModuleName = Space(260)
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Must be set prior to calling API
                LngSize = 500
                Rem -> ~~~~~~~~~~~~~~~~
                Rem -> Get Process Name
                RetVal = GetModuleFileNameExA(HWndProcess, LngModules(1), StrModuleName, LngSize)
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Remove trailing spaces
                StrProcessName = Left(StrModuleName, RetVal)
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Check for Matching Upper case result
                StrProcessName = UCase$(Trim$(StrProcessName))
                Rem -> ~~~~~~~~~~~~~~~~
                Rem -> Get Process Name
                StrProcName = Mid(StrProcessName, InStrRev(StrProcessName, "\") + 1, Len(StrProcessName) - InStrRev(StrProcessName, "\"))
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Check if this is the proces we need to terminate it or not.
                If (UCase(Trim(StrProcName)) = Trim(UCase(StrExeName))) Then
                    Rem -> ~~~~~~~~~~~~~~~~~~~~~
                    Rem -> Terminate the process
                    RetVal = TerminateProcess(HWndProcess, 0)
                End If
            End If
        End If
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Close the handle to this process
        RetVal = CloseHandle(HWndProcess)
    Next
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This sub will idecate wherever we need to auto restart explorer when we close it
Public Sub AutoRestartShell(ByVal Enabled As Boolean)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> This will turn on/off whether or not the windows shell restarts if it is
    Rem -> Shutdown or not. This only works on NT based systems
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> If this is not an NT machine, this won't work
    If Not IsWinNT Then Exit Sub
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> set the value to enable or disable the specified setting
    Call CreateRegLong(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "AutoRestartShell", Abs(Enabled))
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> Is the current platform NT??
Public Function IsWinNT() As Boolean
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Detect if the program is running under an NT based system (NT, 2000, XP)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> holds the operating system information
    Dim OSInfo    As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> get version information
    GetVersionEx OSInfo
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> return True if the test of windows NT is positive
    Rem -> 2 mean that OS is NT.
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
Rem -> *************************************************************************************************************************************************
Rem -> This will only return True if the version returned by the registry
Rem -> value CurrentVersion is 5.0, So its windows 2000
Public Function IsW2000() As Boolean
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Detect if the program is running under an NT based system (NT, 2000, XP)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> holds the operating system information
    Dim OSInfo    As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> get version information
    GetVersionEx OSInfo
    If (OSInfo.dwMajorVersion & "." & OSInfo.dwMinorVersion) = "5.0" Then: IsW2000 = True: Else: IsW2000 = False
End Function
Rem -> *************************************************************************************************************************************************
Rem -> This will only return True if the version returned by the registry
Rem -> value CurrentVersion is 5.1, So its Win-XP
Public Function IsWinXP() As Boolean
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Detect if the program is running under an NT based system (NT, 2000, XP)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> holds the operating system information
    Dim OSInfo    As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> get version information
    GetVersionEx OSInfo
    If (OSInfo.dwMajorVersion & "." & OSInfo.dwMinorVersion) = "5.1" Then: IsWinXP = True: Else: IsWinXP = False
End Function
Rem -> *************************************************************************************************************************************************
Rem -> This will create value in the registry of the specified type And value data
Public Sub CreateRegLong(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, ByVal strValueName As String, ByVal LngData As Long, Optional ByVal EnmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)
    Dim HKey        As Long     'Rem -> Holds a pointer to an open registry key
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Make sure the registry value exists
    Call CreateSubKey(EnmHive, StrSubKey)
    Rem -> ~~~~~~~~~~~~~~~
    Rem -> Open the subkey
    HKey = GetSubKeyHandle(EnmHive, StrSubKey, KEY_ALL_ACCESS)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Create the registry value
    RegSetValueEx HKey, strValueName, 0, EnmType, LngData, 4
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~
    Rem -> close the registry key
    RegCloseKey HKey
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This procedure will create a sub key in the specified header key.
Public Sub CreateSubKey(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String)
    Dim HKey        As Long     'Rem -> holds the handle to the created key
    Rem -> ~~~~~~~~~~~~~~
    Rem -> Create the key
    RegCreateKey EnmHive, StrSubKey & Chr(0), HKey
    Rem -> ~~~~~~~~~~~~~
    Rem -> Close the key
    RegCloseKey HKey
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> This function returns a handle to the specified registry key
Private Function GetSubKeyHandle(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, Optional ByVal EnmAccess As RegistryKeyAccess = KEY_READ) As Long
    Dim HKey        As Long     'Rem -> holds the handle to the specified key
    Dim RetVal      As Long     'Rem -> holds any returned error value from an api call
    Rem -> ~~~~~~~~~~~~~~~~~~~~~
    Rem -> Open the registry key
    RetVal = RegOpenKeyEx(EnmHive, StrSubKey, 0, EnmAccess, HKey)
    If RetVal <> ERROR_SUCCESS Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~
        Rem -> Could not create key
        HKey = 0
    End If
    Rem -> ~~~~~~~~~~~~
    Rem -> Return value
    GetSubKeyHandle = HKey
End Function
Rem -> *************************************************************************************************************************************************
Rem -> #################################################################################################################################################
Rem ->
Rem -> The following functions are provided by "DMS-Meshal", http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=62020&lngWId=1&txtForceRefresh=321200613175949213
Rem -> Thanks a lot for sharing it, Also thanks to the original source that provide that matrial to him, as he said ;)
Rem ->
Rem -> #################################################################################################################################################
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper function which represent wproc, which called by windows, in this function
Rem -> We detect which key is presed and enable or disable it, you may disable the
Rem -> Entire keyboard by using Exit Function on the first line.
Public Function KBWProc(ByVal KeyCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Check if the code passed in is the one we're concerned about (means a key was pressed)
    If KeyCode = HC_Action Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Copy the information into our hook object
        Call CopyMemory(Hook, ByVal lParam, Len(Hook))
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Intercept the Alt + TAB
        If (Hook.VKCode = VK_Tab) Then
            KBWProc = 1
            Exit Function
        End If
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Intercept the Left Ctrl + Escape
        If (Hook.VKCode = VK_Escape) And (Hook.Flags = 0) And CBool((GetAsyncKeyState(VK_LCtrl) And &H8000)) Then
            KBWProc = 1
            Exit Function
        End If
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Intercept the Right Ctrl + Escape
        If (Hook.VKCode = VK_Escape) And (Hook.Flags = 0) And CBool((GetAsyncKeyState(VK_RCtrl) And &H8000)) Then
            KBWProc = 1
            Exit Function
        End If
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Intercept the left Window's key
        If (Hook.VKCode = VK_LWinKey) And (Hook.Flags = 1) Then
            KBWProc = 1
            Exit Function
        End If
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Intercept the right Window's key
        If (Hook.VKCode = VK_RWinKey) And (Hook.Flags = 1) Then
            KBWProc = 1
            Exit Function
        End If
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Handle any other messages as usual
    KBWProc = CallNextHookEx(KeyHook, KeyCode, wParam, lParam)
End Function
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper function that hook the keyboard to disable some of hotkeys
Public Function SetHook()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Hook into the keyboard process
    If KeyHook = 0 Then KeyHook = SetWindowsHookEx(WH_KeyBoard, AddressOf KBWProc, App.hInstance, 0&)
    If KeyHook = 0 Then
        MsgBox "Failed to install Keyboard Hook"
        HookEnabled = False
        Exit Function
    Else
        HookEnabled = True
    End If
End Function
Rem -> *************************************************************************************************************************************************
Rem -> Wrapper function that unhook the keyboard to prevent app. crash
Public Function UnSetHook()
    If HookEnabled Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Unhook from the keyboard process so we don't crash when we try to exit our program
        Call UnhookWindowsHookEx(KeyHook)
    End If
    HookEnabled = False
End Function
