Public Class Style
    Private _style As Object

    Friend Sub New(ByVal style As Object)
        Me._style = style
    End Sub

    Friend ReadOnly Property underlyingComObject() As Object
        Get
            Return Me._style
        End Get
    End Property

    Public Property NameLocal() As String
        Get
            Return Me._style.NameLocal
        End Get
        Set(ByVal value As String)
            Me._style.NameLocal = value
        End Set
    End Property

    Public Property AutomaticallyUpdate() As Boolean
        Get
            Return Me._style.AutomaticallyUpdate
        End Get
        Set(ByVal value As Boolean)
            Me._style.AutomaticallyUpdate = value
        End Set
    End Property

    Public Property BaseStyle() As Style
        Get
            Return New Style(Me._style.BaseStyle)
        End Get
        Set(ByVal value As Style)
            Me._style.BaseStyle = value.underlyingComObject.NameLocal
        End Set
    End Property

    Public Sub BaseStyleSet(ByVal StyleName As String)
        Me._style.BaseStyle = StyleName
    End Sub

    Public Sub BaseStyleSet(ByVal StyleNumber As Integer)
        Me._style.BaseStyle = StyleNumber
    End Sub

    Public Sub BaseStyleSet(ByVal BuiltInStyle As WdBuiltInStyle)
        Me._style.BaseStyle = BuiltInStyle
    End Sub

    Public Property NextParagraphStyle() As Style
        Get
            Return New Style(Me._style.NextParagraphStyle)
        End Get
        Set(ByVal value As Style)
            Me._style.NextParagraphStyle = value.underlyingComObject.NameLocal
        End Set
    End Property

    Public Sub NextParagraphStyleSet(ByVal StyleName As String)
        Me._style.NextParagraphStyle = StyleName
    End Sub

    Public Sub NextParagraphStyleSet(ByVal StyleNumber As Integer)
        Me._style.NextParagraphStyle = StyleNumber
    End Sub

    Public Sub NextParagraphStyleSet(ByVal BuiltInStyle As WdBuiltInStyle)
        Me._style.NextParagraphStyle = BuiltInStyle
    End Sub

    Public Property Font() As Font
        Get
            Return New Font(Me._style.Font)
        End Get
        Set(ByVal value As Font)
            Me._style.Font = value.underlyingComObject
        End Set
    End Property

    Public Property ParagraphFormat() As ParagraphFormat
        Get
            Return New ParagraphFormat(Me._style.ParagraphFormat)
        End Get
        Set(ByVal value As ParagraphFormat)
            Me._style.ParagraphFormat = value.underlyingComObject
        End Set
    End Property
End Class
