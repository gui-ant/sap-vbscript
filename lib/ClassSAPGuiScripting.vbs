
Class ClassSapGuiScripting
    Public SAP_LOGON_PATH
    Public SESSION_LIMIT
    Public DEFAULT_TRANSACTION_NAME
    
    Private objSapGui
    Private objScriptingEngine
    Private objConnection

    Private lstSessions

    Private SAP_SERVER
    Private SAP_INSTANCE
    Private SAP_SID

    Private SAP_USER
    Private SAP_PASS

    Private DECIMAL_SEPARATOR
    Private THOUSANDS_SEPARATOR

    Public Waiting

    Private Sub Class_Initialize()
        SAP_LOGON_PATH = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
        SESSION_LIMIT = 6
        DEFAULT_TRANSACTION_NAME = "SESSION_MANAGER"

        Set lstSessions = CreateObject("System.Collections.ArrayList")
        Waiting = 0

        on error resume next
        set objSapGui = StartOrGetApplication
        if objSapGui is Nothing then Exit Sub

        set objScriptingEngine = GetScriptingEngine
        if objScriptingEngine is Nothing then Exit Sub
        
        WScript.ConnectObject objScriptingEngine, "Engine_"

        on error goto 0
    End Sub

    Private Function StartOrGetApplication()
        on error resume next 
        Set obj = GetObject("SAPGUI")
        if err.number <> 0 then
            if StartApplication then
                err.clear
            Else
                set StartOrGetApplication = Nothing
                WScript.echo "The application could not be started."
                Exit Function
            End If
        end if
        set StartOrGetApplication = GetObject("SAPGUI")
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
        
        If path = "" or isNull(path) then path = SAP_LOGON_PATH
        
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
    
    Private Function IsUserLoggedIn()
        If objConnection.Sessions(0).Info.User = SAP_USER then IsUserLoggedIn = true
    End Function
    
    Private Sub AppWait()
        Waiting = 1
        Do While (Waiting = 1)
            WScript.Sleep(100)
        Loop
    End Sub
    
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
               AttachSession sess
            End If
        Next

        ' If theres no available sessions, creates a new one
        if lstSessions.count = 0 then
            If objConnection.Sessions.Count < SESSION_LIMIT then
                CreateNewSession        
            End If
        End If
    End Sub

    Public Function CreateNewSession()
        For Each sess In objConnection.Sessions
            If sess.Info.Transaction = DEFAULT_TRANSACTION_NAME Then
                set CreateNewSession = sess
                Exit Function
            End If
        Next
        objConnection.Sessions(0).CreateSession
        AppWait
        set CreateNewSession = lstSessions(lstSessions.count - 1)
    End Function
    
    Private Function ConnectionHasParameters(conn, server, instance, SID)
        HasSameServerAndInstance = InStr(1, conn.ConnectionString, server & " " & instance) > 0
        HasSameSID = InStr(1, conn.ConnectionString, sap & "/SAP_CODEPAGE=" & SID) > 0
        
        ConnectionHasParameters = (HasSameServerAndInstance And HasSameSID)
    End Function

    Public Function Login
        Set sess = new ClassSAPSession
        sess.Attach objConnection.Sessions(0)
        AttachSession sess
        sess.GetElement("txtRSYST-BNAME").Text = SAP_USER
        sess..GetElement("pwdRSYST-BCODE").Text = SAP_PASS
        sess.Window(0).SendVKey 0
        sess.StartTransaction "ZSU3"
        sess.SelectElement("tabsTABSTRIP1/tabpDEFA")
        DECIMAL_SEPARATOR   = Mid(sess.GetElement("tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").text, 10, 1)
        THOUSANDS_SEPARATOR = Mid(sess.GetElement("tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").text, 6, 1)
        sess.GoToMenu
    End Function
    
    Public Sub AttachSession(session)
        lstSessions.add session
    End Sub

    Public Function GetAvailableSession()
        Set GetAvailableSession = lstSessions(0)
    End Function
    
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
        Set ScriptingEngine = objScriptingEngine
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
        set objConnection = Nothing
        Waiting = 0
    End Sub

    Sub Engine_CreateSession(ByRef Session)
        WScript.Echo "Session created"
        SapGuiScripting.AttachSession Session
        SapGuiScripting.Waiting = 0
    End Sub
End Class
