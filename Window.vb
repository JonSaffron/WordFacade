Public Class Window
    Private _window As Object

    Friend Sub New(ByVal window As Object)
        Me._window = window
    End Sub

    Public ReadOnly Property ActivePane() As Pane
        Get
            Return New Pane(Me._window.ActivePane)
        End Get
    End Property

    Public ReadOnly Property View() As View
        Get
            Return New View(Me._window.View)
        End Get
    End Property

    Public ReadOnly Property Selection() As Selection
        Get
            Return New Selection(Me._window.Selection)
        End Get
    End Property
End Class
