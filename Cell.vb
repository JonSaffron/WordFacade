Public Class Cell
    Private ReadOnly _cell As Object

    Friend Sub New(ByVal cell As Object)
        Me._cell = cell
    End Sub

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._cell.Range)
        End Get
    End Property

    Public ReadOnly Property Application() As Application
        Get
            Return New Application(Me._cell.Application)
        End Get
    End Property
End Class
