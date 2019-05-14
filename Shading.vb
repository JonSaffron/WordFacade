Public Class Shading
    Private _shading As Object

    Friend Sub New(ByVal shading As Object)
        Me._shading = shading
    End Sub

    Public Property BackgroundPatternColor() As WdColor
        Get
            Return Me._shading.BackgroundPatternColor
        End Get
        Set(ByVal value As WdColor)
            Me._shading.BackgroundPatternColor = value
        End Set
    End Property

    Public Property BackgroundPatternColorIndex() As WdColorIndex
        Get
            Return Me._shading.BackgroundPatternColorIndex
        End Get
        Set(ByVal value As WdColorIndex)
            Me._shading.BackgroundPatternColorIndex = BackgroundPatternColorIndex
        End Set
    End Property

    Public Property ForegroundPatternColor() As WdColor
        Get
            Return Me._shading.ForegroundPatternColor
        End Get
        Set(ByVal value As WdColor)
            Me._shading.ForegroundPatternColor = value
        End Set
    End Property

    Public Property ForegroundPatternColorIndex() As WdColorIndex
        Get
            Return Me._shading.ForegroundPatternColorIndex
        End Get
        Set(ByVal value As WdColorIndex)
            Me._shading.ForegroundPatternColorIndex = value
        End Set
    End Property

    Public Property Texture() As WdTextureIndex
        Get
            Return Me._shading.Texture
        End Get
        Set(ByVal value As WdTextureIndex)
            Me._shading.Texture = value
        End Set
    End Property
End Class
