Public Class Zoom
    Private _zoom As Object

    Friend Sub New(ByVal zoom As Object)
        Me._zoom = zoom
    End Sub

    Public Property Percentage() As Integer
        Get
            Return Me._zoom.Percentage
        End Get
        Set(ByVal value As Integer)
            Me._zoom.Percentage = value
        End Set
    End Property
End Class
