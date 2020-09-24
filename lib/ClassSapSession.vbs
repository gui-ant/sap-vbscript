Class ClassSapSession
    
    Private objSession
    Public DEFAULT_TRANSACTION_NAME

    Private Sub Class_Initialize()
        
    End Sub

    Public Sub Login(user, pass)
        WScript.Echo "Logging in " & user & "..."
        GetElement("txtRSYST-BNAME",0).Text = user
        GetElement("pwdRSYST-BCODE",0).Text = pass
        GetWindow(0).SendVKey 0
        DEFAULT_TRANSACTION_NAME = GetTransaction
    End Sub

    Public Sub StartTransaction(code)
        objSession.StartTransaction code
    End Sub

    Public Property Get GetTransaction()
        GetTransaction = objSession.Info.Transaction
    End Property
    
    Public Property Get GetWindow(index)
        set GetWindow = objSession.FindById("wnd[" & index &"]")
    End Property
    
    Public Property Get GetUserArea(wnd)
        if isNull(wnd) then wnd = 0
        set GetUserArea = GetWindow(wnd).FindById("usr")
    End Property
    
    Public Property Get GetStatusBar(wnd)
        Set GetStatusBar = GetWindow(wnd).FindById("sbar")
    End Property
    
    Public Property Get GetSbarMsgType(wnd)
        GetSbarMsgType = GetStatusBar(wnd).MessageType
    End Property

    Public Property Get GetObject()
        Set GetObject = objSession
    End Property

    Public Function GetElement(id, wnd)
        if isNull(wnd) then wnd = 0
        Set GetElement = GetUserArea(wnd).FindById(id)
    End Function

    Public Sub GoToMenu()
        StartTransaction DEFAULT_TRANSACTION_NAME
    End Sub

    Public Sub SelectElement(id, wnd)
        GetElement(id, wnd).select
    End Sub
    
    Public Sub PressEnter(times, wnd)
        if isNull(times) then times = 1
        t = 0
        Do While t < times
            GetWindow(wnd).SendVKey 0
            IgnoreWarnings(wnd)
            t = t + 1
        Loop
    End Sub

    Sub IgnoreWarnings(wnd)
        Do While GetSbarMsgType(wnd) = "W"
            PressEnter 1, wnd
        Loop
    End Sub
   
    Sub SetValue(ByVal field, wnd, value)
        If value < 0 Then value = Abs(value)
        GetElement(field ,wnd).text = value
    End Sub
    
    Private Property Get GetToolbar(wnd, tbar)
        Set GetToolbar = GetWindow(wnd).FindById("/tbar[" & tbar & "]/")
    End Property

    Sub PressToolbarBtn(buttonID, wnd, tbar)
        WScript.echo "Button pressed: " & GetToolbar(wnd, tbar).FindById(buttonID).Id
        GetToolbar(wnd, tbar).FindById(buttonID).press
    End Sub

    Sub Handle(ByRef session)
        Set objSession = session
    End Sub
    
    Private Sub Class_Terminate()
        set objSession = Nothing
    End Sub

End Class
