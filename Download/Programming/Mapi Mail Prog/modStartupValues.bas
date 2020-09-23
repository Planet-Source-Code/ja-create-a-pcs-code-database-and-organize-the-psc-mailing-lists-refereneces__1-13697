Attribute VB_Name = "modStartupValues"
Public LoadSQLatStartup As Boolean
Public StartUpSQL As String
Public TitleColor As Long
Public TitleBold As Boolean
Public TitleItalics As Boolean
Public TitleUnderline As Boolean
Public TitleAllingment As Integer

Public TextColor As Long
Public TextBold As Boolean
Public TextItalics As Boolean
Public TextUnderline As Boolean
Public TextAllingment As Integer

Public FieldColor As Long
Public FieldBold As Boolean
Public FieldItalics As Boolean
Public FieldUnderline As Boolean
Public FieldAllingment As Integer

Public HttpColor As Long
Public HttpBold As Boolean
Public HttpItalics As Boolean
Public HttpUnderline As Boolean
Public HttpAllingment As Integer

Public CommentColor As Long
Public CommentBold As Boolean
Public CommentItalics As Boolean
Public CommentUnderline As Boolean
Public CommentAllingment As Integer

Public rtBackColor As Long

Function GetRegString(Key As String, DefValue As String) As String
GetRegString = GetSetting("PSC Database", "Startup", Key, DefValue)

End Function

Sub LoadStartUpSettings()
'Misc
LoadSQLatStartup = CBool(GetRegString("LoadSQLatStartup", "0"))
StartUpSQL = GetRegString("StartupSQL", "Select PSC1.* FROM PSC1 order by IndexID ASC,datesumbitted DESC,Accessed DESC")
'Title in search box
TitleColor = Val(GetRegString("TitleColor", "&H0000FF"))
TitleBold = CBool(GetRegString("TitleBold", "-1"))
TitleItalics = CBool(GetRegString("TitleItalics", "0"))
TitleUnderline = CBool(GetRegString("TitleUnderline", "-1"))
TitleAllingment = CInt(GetRegString("TitleAlign", CStr(rtfCenter)))
'Text in search box
TextColor = Val(GetRegString("TextColor", "&H000000"))
TextBold = CBool(GetRegString("TextBold", "0"))
TextItalics = CBool(GetRegString("TextItalics", "0"))
TextUnderline = CBool(GetRegString("TextUnderline", "0"))
TextAllingment = CInt(GetRegString("TextAlign", CStr(rtfLeft)))
'Field in search box
FieldColor = Val(GetRegString("FieldColor", "&H000000"))
FieldBold = CBool(GetRegString("FieldBold", "-1"))
FieldItalics = CBool(GetRegString("FieldItalics", "0"))
FieldUnderline = CBool(GetRegString("FieldUnderline", "0"))
FieldAllingment = CInt(GetRegString("FieldAlign", CStr(rtfLeft)))
'Links in search box
HttpColor = Val(GetRegString("HttpColor", "&HFF0000"))
HttpBold = CBool(GetRegString("HttpBold", "-1"))
HttpItalics = CBool(GetRegString("HttpItalics", "0"))
HttpUnderline = CBool(GetRegString("HttpUnderline", "-1"))
HttpAllingment = CInt(GetRegString("HttpAlign", CStr(rtfLeft)))
'Coments in search box
CommentColor = Val(GetRegString("CommentColor", "&H000000"))
CommentBold = CBool(GetRegString("CommentBold", "0"))
CommentItalics = CBool(GetRegString("CommentItalics", "-1"))
CommentUnderline = CBool(GetRegString("CommentUnderline", "-1"))
CommentAllingment = Val(GetRegString("CommentAlign", CStr(rtfLeft)))
'Searchbox back color
rtBackColor = Val(GetRegString("BackColor", "12648447"))
End Sub


Sub PutRegString(Key As String, Value As String)
SaveSetting "PSC Database", "Startup", Key, Value

End Sub


