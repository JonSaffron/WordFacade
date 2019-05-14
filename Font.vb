Public Class Font
    Private ReadOnly _font As Object

    Friend Sub New(ByVal font As Object)
        Me._font = font
    End Sub

    Friend ReadOnly Property underlyingComObject() As Object
        Get
            Return Me._font
        End Get
    End Property

    Public Property Name() As String
        Get
            Return Me._font.Name
        End Get
        Set(ByVal value As String)
            Me._font.Name = value
        End Set
    End Property

    Public Property Size() As Single
        Get
            Return Me._font.Size
        End Get
        Set(ByVal value As Single)
            Me._font.Size = value
        End Set
    End Property

    Public Property Color() As WdColor
        Get
            Return Me._font.Color
        End Get
        Set(ByVal value As WdColor)
            Me._font.Color = value
        End Set
    End Property

    Public Property Italic() As WdConstants
        Get
            Return Me._font.Italic
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Italic = value
        End Set
    End Property

    Public Property Bold() As WdConstants
        Get
            Return Me._font.Bold
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Bold = value
        End Set
    End Property

    Public Property Underline() As WdUnderline
        Get
            Return Me._font.Underline
        End Get
        Set(ByVal value As WdUnderline)
            Me._font.Underline = value
        End Set
    End Property

    Public Property UnderlineColor() As WdColor
        Get
            Return Me._font.UnderlineColor
        End Get
        Set(ByVal value As WdColor)
            Me._font.UnderlineColor = value
        End Set
    End Property

    Public Property StrikeThrough() As WdConstants
        Get
            Return Me._font.StrikeThrough
        End Get
        Set(ByVal value As WdConstants)
            Me._font.StrikeThrough = value
        End Set
    End Property

    Public Property DoubleStrikeThrough() As WdConstants
        Get
            Return Me._font.DoubleStrikeThrough
        End Get
        Set(ByVal value As WdConstants)
            Me._font.DoubleStrikeThrough = value
        End Set
    End Property

    Public Property Outline() As WdConstants
        Get
            Return Me._font.Outline
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Outline = value
        End Set
    End Property

    Public Property Emboss() As WdConstants
        Get
            Return Me._font.Emboss
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Emboss = value
        End Set
    End Property

    Public Property Shadow() As WdConstants
        Get
            Return Me._font.Shadow
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Shadow = value
        End Set
    End Property

    Public Property Hidden() As WdConstants
        Get
            Return Me._font.Hidden
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Hidden = value
        End Set
    End Property

    Public Property SmallCaps() As WdConstants
        Get
            Return Me._font.SmallCaps
        End Get
        Set(ByVal value As WdConstants)
            Me._font.SmallCaps = value
        End Set
    End Property

    Public Property AllCaps() As WdConstants
        Get
            Return Me._font.AllCaps
        End Get
        Set(ByVal value As WdConstants)
            Me._font.AllCaps = value
        End Set
    End Property

    Public Property Engrave() As WdConstants
        Get
            Return Me._font.Engrave
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Engrave = value
        End Set
    End Property

    Public Property Superscript() As WdConstants
        Get
            Return Me._font.Superscript
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Superscript = value
        End Set
    End Property

    Public Property Subscript() As WdConstants
        Get
            Return Me._font.Subscript
        End Get
        Set(ByVal value As WdConstants)
            Me._font.Subscript = value
        End Set
    End Property

    Public Property Scaling() As Integer
        Get
            Return Me._font.Scaling
        End Get
        Set(ByVal value As Integer)
            Me._font.Scaling = 100
        End Set
    End Property

    Public Property Kerning() As Single
        Get
            Return Me._font.Kerning
        End Get
        Set(ByVal value As Single)
            Me._font.Kerning = value
        End Set
    End Property

    Public Property Animation() As WdAnimation
        Get
            Return Me._font.Animation
        End Get
        Set(ByVal value As WdAnimation)
            Me._font.Animation = value
        End Set
    End Property
End Class
