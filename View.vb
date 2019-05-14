Public Class View
    Private _view As Object

    Friend Sub New(ByVal view As Object)
        Me._view = view
    End Sub

    Public Property Type() As WdViewType
        Get
            Return Me._view.Type
        End Get
        Set(ByVal value As WdViewType)
            Me._view.Type = value
        End Set
    End Property

    Public Property SeekView() As WdSeekView
        Get
            Return Me._view.SeekView
        End Get
        Set(ByVal value As WdSeekView)
            Me._view.SeekView = value
        End Set
    End Property

    Public ReadOnly Property Zoom() As Zoom
        Get
            Return New Zoom(Me._view.Zoom)
        End Get
    End Property
End Class
