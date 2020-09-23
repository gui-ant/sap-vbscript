' Globals
Dim strLibDir   : strLibDir = ".\lib"
Dim objFileSys  : Set objFileSys = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "ClassSapGuiScripting.vbs")).ReadAll()

Dim SapGuiScripting
Set SapGuiScripting = New ClassSapGuiScripting

Sub Engine_CreateSession(ByRef Session)
  WScript.Echo "Session created"
  SapGuiScripting.AttachSession Session
  SapGuiScripting.Waiting = 0
End Sub


