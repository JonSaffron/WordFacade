Public Class Application
    Private ReadOnly _app As Object

    Public Sub New()
        Try
            Dim typePowerPoint As Type = Type.GetTypeFromProgID("Word.Application")
            Me._app = Activator.CreateInstance(typePowerPoint)
        Catch ex As Exception
            Throw New InvalidOperationException("It was not possible to start Word - " & ex.Message, ex)
        End Try
    End Sub

    Friend Sub New(ByVal _application As Object)
        Me._app = _application
    End Sub

    Public Property Visible() As Boolean
        Get
            Return Me._app.Visible
        End Get
        Set(ByVal value As Boolean)
            Me._app.visible = value
        End Set
    End Property

    Public ReadOnly Property Documents() As Documents
        Get
            Return New Documents(Me._app.Documents)
        End Get
    End Property

    Public ReadOnly Property Version() As String
        Get
            Return Me._app.Version
        End Get
    End Property

    Public Sub Quit()
        Call Me._app.Quit()
    End Sub

    Public Sub Quit(ByVal savechanges As WdSaveOptions)
        Call Me._app.Quit(savechanges)
    End Sub

    Public Sub Run(ByVal macroname As String)
        Call Me._app.Run(macroname)
    End Sub

    Public Sub Run(ByVal macroname As String, ByVal varg1 As Object)
        Call Me._app.Run(macroname, varg1)
    End Sub

    Public ReadOnly Property Options() As Options
        Get
            Return New Options(Me._app.Options)
        End Get
    End Property

    Public Property ScreenUpdating() As Boolean
        Get
            Return Me._app.ScreenUpdating
        End Get
        Set(ByVal value As Boolean)
            Me._app.ScreenUpdating = value
        End Set
    End Property

    Public Property DisplayAlerts() As WdAlertLevel
        Get
            Return Me._app.DisplayAlerts
        End Get
        Set(ByVal value As WdAlertLevel)
            Me._app.DisplayAlerts = value
        End Set
    End Property

    Public Property WindowState() As WdWindowState
        Get
            Return Me._app.WindowState
        End Get
        Set(ByVal value As WdWindowState)
            Me._app.WindowState = value
        End Set
    End Property

    Public ReadOnly Property ActiveWindow() As Window
        Get
            Return New Window(Me._app.ActiveWindow)
        End Get
    End Property

    Public ReadOnly Property Selection() As Selection
        Get
            Return New Selection(Me._app.Selection)
        End Get
    End Property

    Public Function CentimetersToPoints(ByVal centimeters As Single) As Single
        Return Me._app.CentimetersToPoints(centimeters)
    End Function

    Public ReadOnly Property COMAddIns() As COMAddIns
        Get
            Return New COMAddIns(Me._app.COMAddIns)
        End Get
    End Property
End Class
