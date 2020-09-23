' Globals
Dim strLibDir       : strLibDir = ".\lib"
Dim objFSys  : Set objFSys = CreateObject("Scripting.FileSystemObject")
Dim cls_sap

Function Main()
  strPw = Password( "Please enter your password:" )
  WScript.Echo "Your password is: " & strPw
  
  'Set cls_sap = New ClassSAPGUIScripting
  
  'cls_sap.SetConnectionParams "epr.sig.defesa.pt", "00", "110"
  'cls_sap.SetUserParams "D0402214", "GfA0a7"
  'cls_sap.Attach

  Main = 1 ' can't call this a success
End Function ' Main

Sub Engine_CreateSession(ByVal Session)
  WScript.Echo "Session created"
End Sub

Include "ClassSAPGUIScripting.vbs"

WScript.Quit Main()

Function Include(file)
  ExecuteGlobal objFSys.OpenTextFile(objFSys.BuildPath(strLibDir, file)).ReadAll()
End Function ' Include

Function Password( myPrompt )
' This function hides a password while it is being typed.
' myPrompt is the text prompting the user to type a password.
' The function clears the prompt and returns the typed password.
' This code is based on Microsoft TechNet ScriptCenter "Mask Command Line Passwords"
' http://www.microsoft.com/technet/scriptcenter/scripts/default.mspx?mfr=true

    ' Standard housekeeping
    Dim objPassword

    ' Use ScriptPW.dll by creating an object
    Set objPassword = CreateObject( "ScriptPW.Password" )

    ' Display the prompt text
    WScript.StdOut.Write myPrompt

    ' Return the typed password
    Password = objPassword.GetPassword()

    ' Clear prompt
    WScript.StdOut.Write String( Len( myPrompt ), Chr( 8 ) ) _
                       & Space( Len( myPrompt ) ) _
                       & String( Len( myPrompt ), Chr( 8 ) )
End Function