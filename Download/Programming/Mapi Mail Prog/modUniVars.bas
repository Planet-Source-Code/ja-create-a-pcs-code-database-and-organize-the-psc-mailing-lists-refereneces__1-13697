Attribute VB_Name = "modUniVars"
Public UniVarDate As Date
Public UniVarPrompt As Boolean
Public UniVarmsgIDTostartFrom As Long
Public Enum PromptRespone
    yes = 1
    No = 0
    YesForAll = -1
End Enum

Public pr As PromptRespone

Public LastDateToStore As Date
Public LastMEssageIDtoStore As Long

Function FormatTime(nSeconds As Long) As String
Dim nHours As Integer, nMinutes As Integer, cSeconds As Integer
Dim nTMP As Long

nHours = nSeconds / 3600
nTMP = nSeconds Mod 3600
nMinutes = nTMP / 60
nTMP = nTMP Mod 60
FormatTime = Format(nHours, "00") & ":" & Format(nMinutes, "00") & ":" & Format(nTMP, "00")




End Function


Function GetPromptReponce() As PromptRespone
Dialog.Show 1
GetPromptReponce = pr
End Function


