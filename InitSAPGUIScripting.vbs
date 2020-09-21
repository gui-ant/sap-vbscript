' Globals
Dim gsLibDir : gsLibDir = ".\lib"
Dim goFS     : Set goFS = CreateObject("Scripting.FileSystemObject")

' LibraryInclude
ExecuteGlobal goFS.OpenTextFile(goFS.BuildPath(gsLibDir, "ClassSAPGUIScripting.vbs")).ReadAll()

WScript.Quit main()

Function main()
  Dim cls_sap : Set cls_sap = New ClassSAPGUIScripting
  cls_sap.SetConnectionParams "epr.sig.defesa.pt", "00", "110"
  cls_sap.SetUserParams "D0402214", "GfA0a7"
  cls_sap.Attach
  msgbox "Terminado"
  main = 1 ' can't call this a success
End Function ' main