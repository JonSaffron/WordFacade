Public Class HeaderFooter
    Private ReadOnly _headerfooter As Object

    Friend Sub New(ByVal headerfooter As Object)
        Me._headerfooter = headerfooter
    End Sub

    Public Property Exists() As Boolean
        Get
            Return Me._headerfooter.Exists
        End Get
        Set(ByVal value As Boolean)
            Me._headerfooter.Exists = value
        End Set
    End Property

    Public ReadOnly Property Index() As WdHeaderFooterIndex
        Get
            Return Me._headerfooter.Index
        End Get
    End Property

    Public ReadOnly Property IsHeader() As Boolean
        Get
            Return Me._headerfooter.IsHeader
        End Get
    End Property

    Public Property LinkToPrevious() As Boolean
        Get
            Return Me._headerfooter.LinkToPrevious
        End Get
        Set(ByVal value As Boolean)
            Me._headerfooter.LinkToPrevious = value
        End Set
    End Property

    ' Public Property PageNumbers() as PageNumbers

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._headerfooter.Range)
        End Get
    End Property

    ' Public Property Shapes() as Shapes
End Class
