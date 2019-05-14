Public Class Pane
    Private _pane As Object

    Friend Sub New(ByVal pane As Object)
        Me._pane = pane
    End Sub

    Public ReadOnly Property View() As View
        Get
            Return New View(Me._pane.View)
        End Get
    End Property

    Public ReadOnly Property Selection() As Selection
        Get
            Return New Selection(Me._pane.Selection)
        End Get
    End Property
End Class
