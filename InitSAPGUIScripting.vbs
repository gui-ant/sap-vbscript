' Globals
Dim strLibDir   : strLibDir = ".\lib"
Dim objFileSys  : Set objFileSys = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "ClassSAPGUIScripting.vbs")).ReadAll()

Dim SAPGUIScripting
set SAPGUIScripting = New ClassSAPGUIScripting

Sub Engine_CreateSession(ByRef Session)
  WScript.Echo "Session created"
  SAPGUIScripting.AttachSession Session
  SAPGUIScripting.Waiting = 0
End Sub


