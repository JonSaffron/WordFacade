Public Class Column
    Private ReadOnly _column As Object

    Friend Sub New(ByVal column As Object)
        Me._column = column
    End Sub

    Friend ReadOnly Property underlyingobject() As Object
        Get
            Return Me._column
        End Get
    End Property

    Public Sub Delete()
        Call Me._column.Delete()
    End Sub

    Public ReadOnly Property Shading() As Shading
        Get
            Return New Shading(Me._column.Shading)
        End Get
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._column.Range)
        End Get
    End Property

    Public ReadOnly Property Cells() As Cells
        Get
            Return New Cells(Me._column.Cells)
        End Get
    End Property

    Public Sub [Select]()
        Call Me._column.Select()
    End Sub

    Public Property Width() As Single
        Get
            Return Me._column.Width
        End Get
        Set(ByVal value As Single)
            Me._column.Width = value
        End Set
    End Property
End Class
