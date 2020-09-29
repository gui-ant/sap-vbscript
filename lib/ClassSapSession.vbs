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
        StartTransaction "!"
    End Sub

    Public Sub SelectElement(id, wnd)
        GetElement(id, wnd).select
    End Sub
    
    Public Sub PressEnter(wnd)
        GetWindow(wnd).SendVKey 0
        IgnoreWarnings
    End Sub

    Sub IgnoreWarnings()
        Do While GetSbarMsgType(0) = "W"
            PressEnter 0
        Loop
    End Sub
   
    Sub SetValue(ByVal field, wnd, value)
        if Not IsDate(value) and IsNumeric(value) then
            If value < 0 Then value = Abs(value)
        End If
        GetElement(field ,wnd).text = value
    End Sub
    
    Private Property Get GetToolbar(wnd, tbar)
        Set GetToolbar = GetWindow(wnd).FindById("tbar[" & tbar & "]")
    End Property

    Sub PressToolbarBtn(buttonID, wnd, tbar)
        WScript.echo "Button pressed: " & GetToolbar(wnd, tbar).FindById(buttonID).Id
        GetToolbar(wnd, tbar).FindById(buttonID).press
    End Sub

    Sub PressButton(buttonID, wnd)
        GetElement(buttonID, wnd).Press
    End Sub
    
    Sub SelectMenu(wnd, menu0, menu1)
        GetWindow(wnd).FindById("mbar/menu[" & menu0 & "]/menu[" & menu1 & "]").Select
    End Sub

    Sub Handle(ByRef session)
        Set objSession = session
    End Sub
    
    Private Sub Class_Terminate()
        set objSession = Nothing
    End Sub

    
End Class
