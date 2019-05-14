Public Class Bookmark
    Private ReadOnly _bookmark As Object

    Friend Sub New(ByVal bookmark As Object)
        Me._bookmark = bookmark
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Return Me._bookmark.Name
        End Get
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._bookmark.Range)
        End Get
    End Property

    Public Property Start() As Integer
        Get
            Return Me._bookmark.Start
        End Get
        Set(ByVal value As Integer)
            Me._bookmark.Start = value
        End Set
    End Property

    Public Property [End]() As Integer
        Get
            Return Me._bookmark.End
        End Get
        Set(ByVal value As Integer)
            Me._bookmark.End = value
        End Set
    End Property

    Public Sub Delete()
        Call Me._bookmark.Delete()
    End Sub

    Public ReadOnly Property Parent() As Document
        Get
            Return New Document(Me._bookmark.Parent)
        End Get
    End Property
End Class
