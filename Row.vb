Public Class Row
    Private _row As Object

    Friend Sub New(ByVal row As Object)
        Me._row = row
    End Sub

    Friend ReadOnly Property underlyingobject() As Object
        Get
            Return Me._row
        End Get
    End Property

    Public Sub Delete()
        Call Me._row.Delete()
    End Sub

    Public ReadOnly Property Shading() As Shading
        Get
            Return New Shading(Me._row.Shading)
        End Get
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._row.Range)
        End Get
    End Property

    Public ReadOnly Property Cells() As Cells
        Get
            Return New Cells(Me._row.Cells)
        End Get
    End Property

    Public Property HeadingFormat() As WdConstants
        Get
            Return Me._row.HeadingFormat
        End Get
        Set(ByVal value As WdConstants)
            Me._row.HeadingFormat = value
        End Set
    End Property
End Class
