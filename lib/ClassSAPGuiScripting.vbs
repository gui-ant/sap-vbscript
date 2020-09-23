Const SAP_LOGON_PATH = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
Const SESSION_LIMIT = 6
Const DEFAULT_TRANSACTION_NAME = "SESSION_MANAGER"

Class ClassSAPGUIScripting
    
    Private objSapGui
    Private objScriptingEngine
    Private objConnection

    Private lstSessions

    Private SAP_SERVER
    Private SAP_INSTANCE
    Private SAP_SID

    Private SAP_USER
    Private SAP_PASS

    Private Waiting

    Private Sub Class_Initialize()

        Set lstSessions = CreateObject("System.Collections.ArrayList")
        Waiting = 0

        on error resume next
        
        set objSapGui = StartOrGetApplication
        if objSapGui is Nothing then Exit Sub
        
        set objScriptingEngine = GetScriptingEngine
        if objScriptingEngine is Nothing then Exit Sub
        
        on error goto 0
    End Sub

    Private Function StartOrGetApplication()
        Set obj = GetObject("SAPGUI")
        if err.number <> 0 then
            if StartApplication then
                set StartOrGetApplication = GetObject("SAPGUI")
                err.clear
                Exit Function
            Else
                WScript.echo "The application could not be started."
                Exit Function
            End If
        end if
        set StartOrGetApplication = Nothing
    End Function

    Private Function GetScriptingEngine()
        Set obj = objSapGui.GetScriptingEngine

        if err.number <> 0 then
            WScript.echo "SAP Scripting Engine not found or it is disabled"
            Set GetScriptingEngine = Nothing
            Exit Function
        end if
        set GetScriptingEngine = obj
    End Function

    Private Function StartOrGetConnection()
        Set conn = GetActiveConnection 
        if conn is Nothing then 
            Set conn = objScriptingEngine.OpenConnectionByConnectionString("/SAP_CODEPAGE=" & SAP_SID & "0 /FULLMENU " & SAP_SERVER & " " & SAP_INSTANCE & " /3 /UPDOWNLOAD_CP=2")
        End If
        Set StartOrGetConnection = conn
    End Function

    Private Function StartApplication(path)
        if path = "" then path = SAP_LOGON_PATH

        Dim WSHShell : Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run path, 1, false
        WScript.echo "Initiating a new SAPGUI instance..."
        
        attempts = 0
        Do Until WSHShell.AppActivate("SAP Logon ")
            WScript.Sleep 500
            attempts = attempts + 1
            if attempts = 10 then 
                WScript.echo "It was not possible to intanciate after " & attempts & " attempts."
                StartApplication = False
                Exit Function
            End If
        Loop
        Set WSHShell = Nothing
        WScript.echo "An new SAP GUI instance was initiated."
        StartApplication = True
    End Function

    Sub Attach()
        Set objConnection = StartOrGetConnection
        If Not IsUserLoggedIn then Login
        AttachAvailableSessions
    End Sub

    Private Sub AppWait()
        Waiting = 1
        Do While (Waiting = 1)
            WScript.Sleep(100)
        Loop
    End Sub

    'Private Sub AttachSessions()
    '    For Each conn In objScriptingEngine.Connections
    '        If InStr(1, conn.ConnectionString, server & " " & instance) > 0 And InStr(1, conn.ConnectionString, "/SAP_CODEPAGE=" & SID) > 0 Then
    '            If conn.Sessions.Count > 0 Then
    '                For Each sess In conn.Sessions
    '                    If sess.Info.user = user And sess.Info.Transaction = "SESSION_MANAGER" And i < SESSION_LIMIT) Then
    '                        Set controlledSessions(i) = S.GetSession
    '                        i = i + 1
    '                    ElseIf sess.Info.user = "" Then
    '                        conn.CloseConnection
    '                    End If
    '                Next
    '            End If
    '        End If
    '    Next
    'End Sub
    
    Private Function GetActiveConnection() 
        For Each conn In objScriptingEngine.Connections
            If ConnectionHasParameters(conn, SAP_SERVER, SAP_INSTANCE, SAP_SID) Then
                Set GetActiveConnection = conn
                Exit Function
            End If
        Next
        Set GetActiveConnection = Nothing
    End Function

    Private Sub AttachAvailableSessions() 
        For Each sess In objConnection.Sessions
            If sess.Info.Transaction = DEFAULT_TRANSACTION_NAME Then
               lstSessions.add sess
            End If
        Next
    End Sub

    Private Function ConnectionHasParameters(conn, server, instance, SID)
        HasSameServerAndInstance = InStr(1, conn.ConnectionString, server & " " & instance) > 0
        HasSameSID = InStr(1, conn.ConnectionString, sap & "/SAP_CODEPAGE=" & SID) > 0
        
        ConnectionHasParameters = (HasSameServerAndInstance And HasSameSID)
    End Function

    Public Function Login(user, pass)
        objConnection.Sessions(0).FindById("wnd[0]/usr/txtRSYST-BNAME").Text = SAP_USER
        objConnection.Sessions(0).FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = SAP_PASS
        objConnection.Sessions(0).FindById("wnd[0]").SendVKey 0
    End Function
    
    Public Sub AttachSession(session)
        lstSessions.add session
    End Sub

    Public Sub SetConnectionParams(server, instance, SID)
        SetServer = server
        SetInstance = instance
        SetSID = SID
    End Sub

    Private Property Let SetServer(server)
        SAP_SERVER = server
    End Property

    Private Property Let SetInstance(instance)
        SAP_INSTANCE = instance
    End Property

    Private Property Let SetSID(id)
        SAP_SID = id
    End Property

    Public Sub SetUserParams(user, password)
        SetUser = user
        SetPassword = password
    End Sub

    Private Property Let SetUser(user)
        SAP_USER = user
    End Property

    Public Property Get User()
        User = SAP_USER
    End Property

    Public Property Get ScriptingEngine()
        ScriptingEngine = objScriptingEngine
    End Property

    Private Property Let SetPassword(pass)
        SAP_PASS = pass
    End Property

    Public Property Get GetSession(index)
        Set GetSession = lstSessions(index)
    End Property

    Private Sub Class_Terminate()
        set objSapGui = Nothing
        set objScriptingEngine = Nothing
        Waiting = 0
    End Sub
End Class
