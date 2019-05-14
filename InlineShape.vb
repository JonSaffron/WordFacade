Public Class InlineShape
    Private _inlineshape As Object

    Friend Sub New(ByVal inlineshape As Object)
        Me._inlineshape = inlineshape
    End Sub

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._inlineshape.Range)
        End Get
    End Property
End Class
