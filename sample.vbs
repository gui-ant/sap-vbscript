Dim objFileSys  : Set objFileSys = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal objFileSys.OpenTextFile(objFileSys.BuildPath(strLibDir, "InitSAPGUIScripting.vbs")).ReadAll()

WScript.Quit Main()

Function Main()
  set sap = SAPGUIScripting

  ' These parameters can be found on the connection properties window, from the SAP Logon screen
  sap.SetConnectionParams "<AppServer>", "<InstanceNr>", "<SystemID>"
  sap.SetUserParams "<Username>", "<Password>"
  sap.Attach

  set ses0 = sap.GetAvailableSession
  ses0.StartTransaction "<transactionCode>"

  set ses1 = sap.CreateNewSession
  ses1.StartTransaction "<transactionCode>"

End Function