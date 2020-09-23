Attribute VB_Name = "modFolderBrowse"
'This module contains all the declarations to use the
'Windows 95 Shell API to use the browse for folders
'dialog box.  To use the browse for folders dialog box,
'please call the BrowseForFolders function using the
'syntax: stringFolderPath=BrowseForFolders(Hwnd,TitleOfDialog)
'

Option Explicit
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Const MAX_PATH = 260
Public Const BIF_RETURNONLYFSDIRS = &H1    ' For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2    ' For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20             ' insist on valid result (or CANCEL)

Public Const BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long



Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_EDITBOX + BIF_STATUSTEXT
        
    End With
'CopyMemory udtBI.lpfnCallback, AddressOf BrowseCallbackProc, Len(udtBI.lpfnCallback)
'Form1.Caption = udtBI.lpfnCallback
    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     '
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath

End Function
