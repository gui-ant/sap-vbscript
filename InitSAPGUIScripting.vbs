' Globals
Dim LIB_DIR       : LIB_DIR = ".\lib"
Dim FILE_SYS_OBJ  : Set FILE_SYS_OBJ = CreateObject("Scripting.FileSystemObject")
Dim cls_sap

Include "ClassSAPGUIScripting.vbs"

WScript.Quit SAPGUIScripting()

Function Include(file)
  Dim filePath : filePath = FILE_SYS_OBJ.BuildPath(LIB_DIR, file)
  ExecuteGlobal FILE_SYS_OBJ.OpenTextFile(filePath).ReadAll()
End Function ' Include

Sub Engine_CreateSession(ByVal Session)
  cls_sap.AttachSession Session
  WScript.Echo "Session created"
End Sub

Function SAPGUIScripting()
  Set cls_sap = New ClassSAPGUIScripting
  WScript.ConnectObject cls_sap.ScriptingEngine,  "Engine_"
  
  set SAPGUIScripting = cls_sap
End Function ' SAPGUIScripting

