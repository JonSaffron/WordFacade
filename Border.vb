Public Class Border
    Private ReadOnly _border As Object

    Friend Sub New(ByVal border As Object)
        Me._border = border
    End Sub

    Public Property Color() As WdColor
        Get
            Return Me._border.Color
        End Get
        Set(ByVal value As WdColor)
            Me._border.Color = value
        End Set
    End Property

    Public Property ColorIndex() As WdColorIndex
        Get
            Return Me._border.ColorIndex
        End Get
        Set(ByVal value As WdColorIndex)
            Me._border.ColorIndex = value
        End Set
    End Property

    Public Property LineStyle() As WdLineStyle
        Get
            Return Me._border.LineStyle
        End Get
        Set(ByVal value As WdLineStyle)
            Me._border.LineStyle = value
        End Set
    End Property

    Public Property LineWidth() As WdLineWidth
        Get
            Return Me._border.LineWidth
        End Get
        Set(ByVal value As WdLineWidth)
            Me._border.LineWidth = value
        End Set
    End Property
End Class
