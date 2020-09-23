Class ClassSAPSession
    
    Private objSession

    Private Sub Class_Initialize()

    End Sub

    Private Function StartTransaction(code)
        objSession.StartTransaction code
    End Function

    Public Property Get GetTransaction()
        CurrentTransaction = objSession.Info.Transaction
    End Property
    
    Public Function Get()
        set Get = objSession
    End Function
    
    Public Property Window(index)
        set Window = objSession.FindById("wnd[" & index &"]")
    End Property
    
    Public Property UserArea(wnd)
        if isNull(wnd) then wnd = 0
        set UserArea = Window(wnd).FindById("usr")
    End Property

    Public Function GetElement(id, wnd)
        if isNull(wnd) then wnd = 0
        set GetElement = UserArea(wnd).FindById(id)
    End Function

    Public Sub SelectElement(id, wnd)
        GetElement(id,wnd).select
    End Sub

    Sub Attach(ByRef session)
        Set objSession = session
    End Sub
    
    Private Sub Class_Terminate()
        set objSession = Nothing
    End Sub

End Class
