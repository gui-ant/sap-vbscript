' Globals
Dim strLibDir   : strLibDir = ".\lib"
Dim objFileSys  : Set objFileSys = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "ClassSapGuiScripting.vbs")).ReadAll()
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "ClassSapSession.vbs")).ReadAll()

Dim SapGuiScripting
Set SapGuiScripting = New ClassSapGuiScripting




