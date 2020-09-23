Attribute VB_Name = "modMiscAPI"
'//////////////////////////////////////////////////////////
'// DisableX Code
Public Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As Long) As Long

Public Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
    
Public Declare Function DrawMenuBar Lib "user32" _
    (ByVal hwnd As Long) As Long

Public Const MF_BYPOSITION = &H400&

Public Enum curs
    http = 1
    folder = 2
    index = 3
    none = 0
End Enum

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
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0




Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long


Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long


Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long


Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_SHARE = 26

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1



'OpenFile
Public Const OFS_MAXPATHNAME = 128
Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private cfo As OFSTRUCT
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Const OF_EXIST = &H4000

Function ClearUpString(st As String) As String
st = Trim(st)
Do While (InStr(1, st, vbCr) > 0) Or (InStr(1, st, vbLf) > 0) Or (InStr(1, st, vbCrLf) > 0)
If Len(Trim(st)) = 0 Then ClearUpString = st: Exit Function
If Right$(st, 1) = vbCr Then
st = Left$(st, Len(st) - 1)
End If
If Len(Trim(st)) = 0 Then ClearUpString = st: Exit Function
If Right$(st, 1) = vbLf Then
st = Left$(st, Len(st) - 1)
End If
If Len(Trim(st)) = 0 Then ClearUpString = st: Exit Function
If Right$(st, 1) = vbNullString Then
st = Left$(st, Len(st) - 1)
End If
If Len(Trim(st)) = 0 Then ClearUpString = st: Exit Function
If Right$(st, 1) = vbCrLf Then
st = Left$(st, Len(st) - 1)
End If
Loop
ClearUpString = st
End Function



Function CreateWebLocation(cFile As String) As String
Dim cFileNew As String

cFileNew = Replace(cFile, " ", "%20")
cFileNew = "file://" & cFileNew
CreateWebLocation = cFileNew


End Function

Sub OpenFolder(path As String)
ret& = ShellExecute(CLng(0), "open", path, vbNullString, App.path, 1)
If ret > 32 Then
Exit Sub
Else
    Select Case ret
        Case 0: r$ = "The operating system is out of memory or resources."
        Case ERROR_FILE_NOT_FOUND: r$ = "The specified file was not found."
        Case ERROR_PATH_NOT_FOUND: r$ = "The specified path was not found."
        Case ERROR_BAD_FORMAT: r$ = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
        Case SE_ERR_ACCESSDENIED: r$ = "The operating system denied access to the specified file. "
        Case SE_ERR_ASSOCINCOMPLETE: r$ = "The file name association is incomplete or invalid."
        Case SE_ERR_DDEBUSY: r$ = "The DDE transaction could not be completed because other DDE transactions were being processed."
        Case SE_ERR_DDEFAIL: r$ = "The DDE transaction failed."
        Case SE_ERR_DDETIMEOUT: r$ = "The DDE transaction could not be completed because the request timed out."
        Case SE_ERR_DLLNOTFOUND: r$ = "The specified dynamic-link library was not found. "
        Case SE_ERR_FNF: r$ = "The specified file was not found. "
        Case SE_ERR_NOASSOC: r$ = "There is no application associated with the given file name extension."
        Case SE_ERR_PNF: r$ = "The specified path was not found."
        Case SE_ERR_OOM: r$ = "There was not enough memory to complete the operation."
        Case SE_ERR_SHARE: r$ = "A sharing violation occurred."
    End Select
    MsgBox r$, vbCritical, "ShellExecute error"
End If

        

End Sub

Sub OpenIndex(ind As Long, cData As Data)
cData.Recordset.FindFirst "Indexid=" & ind

End Sub

Sub openURL(url As String)
If Left$(url, 7) = "http://" Then
ret& = ShellExecute(CLng(0), "open", url, vbNullString, vbNullString, 1)
If ret > 32 Then
Exit Sub
Else
    Select Case ret
        Case 0: r$ = "The operating system is out of memory or resources."
        Case ERROR_FILE_NOT_FOUND: r$ = "The specified file was not found."
        Case ERROR_PATH_NOT_FOUND: r$ = "The specified path was not found."
        Case ERROR_BAD_FORMAT: r$ = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
        Case SE_ERR_ACCESSDENIED: r$ = "The operating system denied access to the specified file. "
        Case SE_ERR_ASSOCINCOMPLETE: r$ = "The file name association is incomplete or invalid."
        Case SE_ERR_DDEBUSY: r$ = "The DDE transaction could not be completed because other DDE transactions were being processed."
        Case SE_ERR_DDEFAIL: r$ = "The DDE transaction failed."
        Case SE_ERR_DDETIMEOUT: r$ = "The DDE transaction could not be completed because the request timed out."
        Case SE_ERR_DLLNOTFOUND: r$ = "The specified dynamic-link library was not found. "
        Case SE_ERR_FNF: r$ = "The specified file was not found. "
        Case SE_ERR_NOASSOC: r$ = "There is no application associated with the given file name extension."
        Case SE_ERR_PNF: r$ = "The specified path was not found."
        Case SE_ERR_OOM: r$ = "There was not enough memory to complete the operation."
        Case SE_ERR_SHARE: r$ = "A sharing violation occurred."
    End Select
    MsgBox r$, vbCritical, "Shellexecute error"
End If

Else
MsgBox "Not a valid URL"
End If

End Sub

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
    ' Determine the size and type of data to
    '     be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5


    Select Case lType
        ' For strings
        Case REG_SZ:
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)


        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        Else
            vValue = Empty
        End If
        ' For DWORDS
        Case REG_DWORD:
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
        Case Else
        'all other data types not supported
        lrc = -1
    End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function


Function GetBrowseCursorResID() As Integer
Dim DefaultBrowser As Variant
    DefaultBrowser = CStr(QueryValue(HKEY_CLASSES_ROOT, "http\shell\open\ddeexec\Application", ""))
Select Case DefaultBrowser
    Case "IExplore"
    GetBrowseCursorResID = 103
    Case "NSShell"
    GetBrowseCursorResID = 104
    
    Case Else
    GetBrowseCursorResID = 105
    End Select
    
End Function

Public Function QueryValue(RootKey As Long, sKeyName As String, sValueName As String) As Variant
    'With this procedure, a call of:
    'varX = QueryValue("TestKey\SubKey1", "S
    '     tringValue")
    'will return the current setting of the
    '     "StringValue" value, and assumes that "S
    '     tringValue" exists in the "TestKey\SubKe
    '     y1" key.
    'If the Value that you query does not ex
    '     ist then QueryValue will return an error
    '     code of 2 - 'ERROR_BADKEY'.
    Dim lRetVal As Long 'result of the API functions
    Dim hKey As Long 'handle of opened key
    Dim vValue As Variant 'setting of queried value
    lRetVal = RegOpenKeyEx(RootKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function

Function HiWord(dw As Long) As Integer

    If dw And &H80000000 Then
        HiWord = (dw \ 65535) - 1
    Else
        HiWord = dw \ 65535
    End If
End Function

Function LoWord(dw As Long) As Integer

    If dw And &H8000& Then
        LoWord = &H8000 Or (dw And &H7FFF&)
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Public Sub DisableX(fHwnd As Long)
    Dim hMenu As Long
    Dim nCount As Long
    hMenu = GetSystemMenu(fHwnd, 0)
    nCount = GetMenuItemCount(hMenu)
    
ret& = RemoveMenu(hMenu, nCount - 1, MF_BYPOSITION)
ret& = RemoveMenu(hMenu, nCount - 2, MF_BYPOSITION)
    
    
    DrawMenuBar fHwnd
End Sub

Public Sub EnableX(fHwnd)
    Dim hMenu As Long
    Dim nCount As Long
    hMenu = GetSystemMenu(fHwnd, 1)
    nCount = GetMenuItemCount(hMenu)

    DrawMenuBar fHwnd
End Sub

