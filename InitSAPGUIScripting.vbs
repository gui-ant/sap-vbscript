' Globals
Dim LIB_DIR       : LIB_DIR = ".\lib"
Dim FILE_SYS_OBJ  : Set FILE_SYS_OBJ = CreateObject("Scripting.FileSystemObject")

Function Include(file)
  Dim filePath : filePath = FILE_SYS_OBJ.BuildPath(LIB_DIR, file)
  ExecuteGlobal FILE_SYS_OBJ.OpenTextFile(filePath).ReadAll()
End Function ' Include

Function Echo(message)
  WScript.echo message
End Function ' Echo

Dim cls_sap

Function Main()

  Set cls_sap = New ClassSAPGUIScripting
  
  cls_sap.SetConnectionParams "epr.sig.defesa.pt", "00", "110"
  cls_sap.SetUserParams "D0402214", "GfA0a7"
  'cls_sap.Attach

  Main = 1 ' can't call this a success
End Function ' Main

Sub Engine_CreateSession(ByVal Session)
  WScript.Echo "Session created"
End Sub

Include "ClassSAPGUIScripting.vbs"

WScript.Quit Main()


