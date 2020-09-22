Class ClassSAPGUIScripting
    Private objSapGui
    Private objScriptingEngine
    Private SapApp
    Private SapConn
    Private SapSession

    Private SAP_SERVER
    Private SAP_INSTANCE
    Private SAP_SID
    Private SAP_USER
    Private SAP_PASS

    Private WSHShell
    Private Waiting

    Private Sub Class_Initialize()
        
        Waiting = 0
        Set WSHShell = CreateObject("WScript.Shell")
        
        on error resume next
        Set objSapGui = GetObject("SAPGUI")
        if err.number <> 0 then
            WSHShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""", 1, false
            WScript.echo "SAP GUI instance not found. Initiating a new SAPGUI instance..."
            
            attempts = 0
            Do Until WSHShell.AppActivate("SAP Logon ")
                WScript.Sleep 500
                attempts = attempts + 1
                if attempts = 10 then 
                    WScript.echo "It was not possible to intanciate, after " & attempts & " attempts."
                    Exit sub
                End If
            Loop

            Set objSapGui = GetObject("SAPGUI")
            WScript.echo "An new SAP GUI instance was initiated."
            err.clear
        end if

        Set objScriptingEngine = objSapGui.GetScriptingEngine

        if err.number <> 0 then
            WScript.echo "SAP Scripting Engine not found"
            err.clear
        end if

        WScript.ConnectObject objScriptingEngine,  "Engine_"
        
        If err.number <> 0 then
            WScript.Echo err.Number & ": " & err.description
            Set WSHShell = Nothing
            Exit Sub
        End if        
        

        'Set objSapGui = Nothing

        'Set WSHShell = CreateObject("WScript.Shell")
        'Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    
    'Do Until WSHShell.AppActivate("SAP Logon ")
    '    Application.Wait Now + TimeValue("0:00:01")
    'Loop
    'Set WSHShell = Nothing
    
    'Dim opt
    'Set SapApp = Nothing
    '
    'opt = MsgBox("Para executar processos 'SAP Gui Scripting', terá de proceder à ativação da funcionalidade. " & _
    '                "Para tal, deverá aceder a 'Ajustar layout local (Alt+F12)' -> 'Opções...' -> 'Acessibilidade & Scripting' -> 'Scripting' -> Selecionar 'Ativar scripting'." & vbNewLine & _
    '                " Tentar novamente?", vbOKCancel, "Scripting desativado")
    'If opt = vbCancel Then Exit Sub

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

    Private Property Let SetUser(user)
        SAP_USER = user
    End Property

    Private Property Let SetPassword(pass)
        SAP_PASS = pass
    End Property

    Public Sub SetUserParams(user, password)
        SetUser = user
        SetPassword = password
    End Sub

    Sub Attach()

        Set SapConn = GetActiveConnection(SAP_SERVER, SAP_INSTANCE, SAP_SID, SAP_USER)
        
        ' Verifica se existe conexão ativa, ou seja, se exite sessão com login
        If SapConn Is Nothing Then
            Set SapConn = objScriptingEngine.OpenConnectionByConnectionString("/SAP_CODEPAGE=" & SAP_SID & "0 /FULLMENU " & SAP_SERVER & " " & SAP_INSTANCE & " /3 /UPDOWNLOAD_CP=2")
            
            AppWait
            
            Set SapSession = SapConn.Sessions(0)
            SapSession.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = SAP_USER
            If SAP_PASS <> "" Then
                SapSession.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = SAP_PASS
                SapSession.FindById("wnd[0]").SendVKey 0
            End If
        Else
            AttachSessions
        End If
            
    End Sub

    Private Sub AppWait()
        Do While (Waiting = 1)
            WScript.Sleep(100)
        Loop
    End Sub

    Private Sub AttachSessions()
        For Each SapConn In objScriptingEngine.Connections
            
            If InStr(1, SapConn.ConnectionString, server & " " & instance) > 0 And InStr(1, SapConn.ConnectionString, sap & "/SAP_CODEPAGE=" & SID) > 0 Then
                If SapConn.Sessions.Count > 0 Then
                    For Each sess In SapConn.Sessions
                        Set SapSession = sess
                        With sess.Info
                            If sess.Info.user = user And sess.Info.Transaction = "SESSION_MANAGER" And i < CInt(GetConfig("SESSION_LIMIT").Values(1)) Then
                                'Dim s As New clsSapSession
                                's.SetSession sapSession
                                Set controlledSessions(i) = S.GetSession
                                i = i + 1
                            ElseIf sess.Info.user = "" Then
                                conn.CloseConnection
                            End If
                        End With
                    Next
                End If
            End If
        Next
    End Sub

    Private Function GetActiveConnection(server, instance, SID, user) 
    
        For Each conn In objScriptingEngine.Connections
            If ConnectionHasParameters(server, instance, SID, conn) Then
                For Each sess In conn.Sessions
                    If sess.Info.user = user Then
                        Set GetActiveConnection = conn
                        Exit Function
                    End If
                Next
            End If
        Next
        Set GetActiveConnection = Nothing
    End Function


    Private Function ConnectionHasParameters(server, instance, SID, conn)
        HasSameServerAndInstance = InStr(1, conn.ConnectionString, server & " " & instance) > 0
        HasSameSID = InStr(1, conn.ConnectionString, sap & "/SAP_CODEPAGE=" & SID) > 0
        
        ConnectionHasParameters = HasSameServerAndInstance And HasSameSID
    End Function


    Public Property Get GetUser()
        GetUser = SAP_USER
    End Property

    Public Property Get GetSession()
        Set GetSession = SapSession
    End Property

    Public Function Login(user, pass)
        
        If user <> "" Then SetUser = user
        If pass <> "" Then SetPass = pass

        If 0 <> "" Then
            Set SAPCon = objScriptingEngine.Children(0)
            Set SapSession = SAPCon.Children(0)
            
        Else
            Debug.Print "No user"
            'Set SAPCon = objScriptingEngine.Children(0)
            'SAPLogon
        End If
    End Function


End Class
