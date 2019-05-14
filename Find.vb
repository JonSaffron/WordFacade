Public Class Find
    Private ReadOnly _find As Object

    Friend Sub New(ByVal find As Object)
        Me._find = find
    End Sub

    Public ReadOnly Property Replacement() As Replacement
        Get
            Return New Replacement(Me._find.Replacement)
        End Get
    End Property

    Public Property Text() As String
        Get
            Return Me._find.Text
        End Get
        Set(ByVal value As String)
            Me._find.Text = value
        End Set
    End Property

    Public Property Forward() As Boolean
        Get
            Return Me._find.Forward
        End Get
        Set(ByVal value As Boolean)
            Me._find.Forward = value
        End Set
    End Property

    Public Property Wrap() As WdFindWrap
        Get
            Return Me._find.Wrap
        End Get
        Set(ByVal value As WdFindWrap)
            Me._find.Wrap = value
        End Set
    End Property

    Public Property Format() As Boolean
        Get
            Return Me._find.Format
        End Get
        Set(ByVal value As Boolean)
            Me._find.Format = value
        End Set
    End Property

    Public Property MatchCase() As Boolean
        Get
            Return Me._find.MatchCase
        End Get
        Set(ByVal value As Boolean)
            Me._find.MatchCase = value
        End Set
    End Property

    Public Property MatchWholeWord() As Boolean
        Get
            Return Me._find.MatchWholeWord
        End Get
        Set(ByVal value As Boolean)
            Me._find.MatchWholeWord = value
        End Set
    End Property

    Public Property MatchWildcards() As Boolean
        Get
            Return Me._find.MatchWildcards
        End Get
        Set(ByVal value As Boolean)
            Me._find.MatchWildcards = value
        End Set
    End Property

    Public Property MatchSoundsLike() As Boolean
        Get
            Return Me._find.MatchSoundsLike
        End Get
        Set(ByVal value As Boolean)
            Me._find.MatchSoundsLike = value
        End Set
    End Property

    Public Property MatchAllWordForms() As Boolean
        Get
            Return Me._find.MatchAllWordForms
        End Get
        Set(ByVal value As Boolean)
            Me._find.MatchAllWordForms = value
        End Set
    End Property

    Public Sub Execute()
        Call Me._find.Execute()
    End Sub

    Public Sub Execute(ByVal FindText As String)
        Call Me._find.Execute(FindText)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean)
        Call Me._find.Execute(FindText, MatchCase)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean, ByVal Forward As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean, ByVal Forward As Boolean, ByVal Wrap As WdFindWrap)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean, ByVal Forward As Boolean, ByVal Wrap As WdFindWrap, ByVal Format As Boolean)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean, ByVal Forward As Boolean, ByVal Wrap As WdFindWrap, ByVal Format As Boolean, ByVal ReplaceWith As String)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith)
    End Sub

    Public Sub Execute(ByVal FindText As String, ByVal MatchCase As Boolean, ByVal MatchWholeWord As Boolean, ByVal MatchWildcards As Boolean, ByVal MatchSoundsLike As Boolean, ByVal MatchAllWordForms As Boolean, ByVal Forward As Boolean, ByVal Wrap As WdFindWrap, ByVal Format As Boolean, ByVal ReplaceWith As String, ByVal Replace As WdReplace)
        Call Me._find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
    End Sub
End Class
