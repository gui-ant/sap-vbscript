Dim objFileSys  : Set objFileSys = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "InitSAPGUIScripting.vbs")).ReadAll()

WScript.Quit Main()

Function Main()
  set sap = SAPGUIScripting

  sap.SetConnectionParams "epr.sig.defesa.pt", "00", "110"
  sap.SetUserParams "D0402214", "GfA0a7"
  sap.Attach

  set ses0 = sap.GetSession(0)
  ses0.StartTransaction "fb01"

  sap.CreateNewSession
  
  set ses1 = sap.GetSession(1)
  'ses1.StartTransaction "fbl3n"

  Main = 1
End Function