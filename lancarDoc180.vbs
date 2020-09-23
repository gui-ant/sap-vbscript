  
  set sap = SAPGUIScripting
  
  sap.SetConnectionParams "epr.sig.defesa.pt", "00", "110"
  sap.SetUserParams "D0402214", "GfA0a7"
  sap.Attach
  
  set ses0 = sap.GetSession(0)
  set ses1 = sap.GetSession(0)
  