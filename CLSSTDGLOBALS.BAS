Attribute VB_Name = "clsStdGlobals"
' API Declarations

Declare Function SQLSetEnvAttr Lib "odbc32.dll" (ByVal henv As Long, _
        ByVal MyAttribute As Long, ByVal Value As Long, ByVal StringLength As _
            Long) As Long

'CONNECTION POOLING CONSTANTS
' env attribute
Public Const SQL_ATTR_CONNECTION_POOLING = 201
' values for SQL_ATTR_CONNECTION_POOLING:
Public Const SQL_CP_ONE_PER_DRIVER = 1
Public Const SQL_IS_INTEGER = -6
Public Const SQL_CP_OFF = 0


' Help declarations
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Integer
Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FINDER = &HB
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const HELP_INDEX = &H3            '  Display index
Public Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2             '  Terminate help
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Public Const HELP_SETWINPOS = &H203&
Public Const HELPMSGSTRING = "commdlg_help"

' Registry Declarations
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
 
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
 
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
 
Public Const KEY_ALL_ACCESS = &HF003F      'Combination of KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, KEY_CREATE_SUB_KEY, KEY_CREATE_LINK, and KEY_SET_VALUE access.
Public Const KEY_CREATE_LINK = &H20        'Permission to create a symbolic link.
Public Const KEY_CREATE_SUB_KEY = &H4      'Permission to create subkeys.
Public Const KEY_ENUMERATE_SUB_KEYS = &H8  'Permission to enumerate subkeys.
Public Const KEY_EXECUTE = &H20019         'Permission for read access.
Public Const KEY_NOTIFY = &H10             'Permission for change notification.
Public Const KEY_QUERY_VALUE = &H1         'Permission to query subkey data.
Public Const KEY_READ = &H20019            'Combination of KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY access.
Public Const KEY_SET_VALUE = &H2           'Permission to set subkey data.
Public Const KEY_WRITE = &H20006           'Combination of KEY_SET_VALUE and KEY_CREATE_SUB_KEY access.

' constants for dwFlags - can be used in conjunction with each other
' see API documentation for their interpretation
Public Const STARTF_USESHOWWINDOW = &H1&
Public Const STARTF_USESIZE = &H2&
Public Const STARTF_USEPOSITION = &H4&
Public Const STARTF_USECOUNTCHARS = &H8&
Public Const STARTF_USEFILLATTRIBUTE = &H10&
Public Const STARTF_RUNFULLSCREEN = &H20&          ' ignored for non-x86 platforms
Public Const STARTF_FORCEONFEEDBACK = &H40&
Public Const STARTF_FORCEOFFFEEDBACK = &H80&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USEHOTKEY = &H200&

Public Const REGKEY_ONLINE_SHOPPING = "OnlineShopping"

  ' *************** for sheriden Bug ***************** added in 09/09/97  John atabak
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
        
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, _
        ByVal lpData As String, lpcbData As Long) As Long
        
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, _
        lpData As Long, lpcbData As Long) As Long
        
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, _
        ByVal lpData As Long, lpcbData As Long) As Long

' Crystal report API declarations
Declare Function PEOpenEngine Lib "crpe32.dll" () As Integer

Declare Function PEOpenPrintJob Lib "crpe32.dll" (ByVal RptName As String) As Integer

Declare Function PEOutputToWindow Lib "crpe32.dll" (ByVal printJob As Integer, _
        ByVal Title As String, ByVal Left As Long, ByVal Top As Long, _
        ByVal Width As Long, ByVal Height As Long, ByVal Style As Long, _
        ByVal PWindow As Long) As Integer

Declare Function PEStartPrintJob Lib "crpe32.dll" (ByVal printJob As Integer, _
        ByVal WaitOrNot As Integer) As Integer
                 
Declare Sub PEClosePrintJob Lib "crpe32.dll" (ByVal printJob As Integer)

Declare Sub PECloseWindow Lib "crpe32.dll" (ByVal printJob As Integer)
' End of Crystal report API declarations

Declare Function DestroyWindow& Lib "user32" (ByVal hwnd As Long)

' Process Threads API

   Public Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type
 
   Public Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type
 
   Public Const NORMAL_PRIORITY_CLASS = &H20&
   Public Const INFINITE = -1&
   
Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
        hHandle As Long, ByVal dwMilliseconds As Long) As Long
 
Declare Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
        lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long
 
Declare Function CloseHandle Lib "kernel32" (ByVal _
        hObject As Long) As Long
 
Declare Function TerminateProcess Lib "kernel32" (ByVal _
        hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function GetExitCodeProcess Lib "kernel32" (ByVal _
        hProcess As Long, lpExitCode As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
        lpClassName As String, ByVal lpWindowName As String) As Long



' Environment API
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName& Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)

' GetActiveWindow handle
Declare Function GetActiveWindow Lib "user32.dll" () As Long
Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean


' SetWindowPos API
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos API Constants
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40


'
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' !!! BEWARE NOT TO ACTIVATE THIS WHILE RUNNING A PROGRAM UNDER VB DEV. MODE
' !!! IT WILL ALSO KILL OFF VB DEV. BECAUSE VB DOES NOT EXECUTE THE PROGRAM
' !!! IN AN INDEPENDANT THREAD.
'
' Exit a process with a return code.
'
Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'
'FDM error handling constants
'
Public Const NSS_RESUME_NEXT = 4
Public Const NSS_RESUME = 5
Public Const NSS_RAISE_ERROR = 6
Public Const NSS_TERMINATE = 7


'cursor related API functions
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Declare Function ClientToScreen& Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const SC_MAXIMIZE = &HF030
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0





