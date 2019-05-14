Public Class Options
    Private _options As Object

    Friend Sub New(ByVal options As Object)
        Me._options = options
    End Sub

    Public Property Pagination() As Boolean
        Get
            Return Me._options.Pagination
        End Get
        Set(ByVal value As Boolean)
            Me._options.Pagination = value
        End Set
    End Property

    Public Property CheckSpellingAsYouType() As Boolean
        Get
            Return Me._options.CheckSpellingAsYouType
        End Get
        Set(ByVal value As Boolean)
            Me._options.CheckSpellingAsYouType = value
        End Set
    End Property

    Public Property CheckGrammarAsYouType() As Boolean
        Get
            Return Me._options.CheckGrammarAsYouType
        End Get
        Set(ByVal value As Boolean)
            Me._options.CheckGrammarAsYouType = value
        End Set
    End Property
End Class
