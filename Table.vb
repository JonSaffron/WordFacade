Public Class Table
    Private _table As Object

    Friend Sub New(ByVal table As Object)
        Me._table = table
    End Sub

    Public ReadOnly Property Columns() As Columns
        Get
            Return New Columns(Me._table.Columns)
        End Get
    End Property

    Public ReadOnly Property Rows() As Rows
        Get
            Return New Rows(Me._table.Rows)
        End Get
    End Property

    Public Function Cell(ByVal row As Integer, ByVal column As Integer) As Cell
        Return New Cell(Me._table.Cell(Row, Column))
    End Function

    Public Property AllowAutoFit() As Boolean
        Get
            Return Me._table.AllowAutoFit
        End Get
        Set(ByVal value As Boolean)
            Me._table.AllowAutoFit = value
        End Set
    End Property

    Public Property PreferredWidthType() As WdPreferredWidthType
        Get
            Return Me._table.PreferredWidthType
        End Get
        Set(ByVal value As WdPreferredWidthType)
            Me._table.PreferredWidthType = value
        End Set
    End Property

    Public Property PreferredWidth() As Single
        Get
            Return Me._table.PreferredWidth
        End Get
        Set(ByVal value As Single)
            Me._table.PreferredWidth = value
        End Set
    End Property

    Public Property TopPadding() As Single
        Get
            Return Me._table.TopPadding
        End Get
        Set(ByVal value As Single)
            Me._table.TopPadding = value
        End Set
    End Property

    Public Property BottomPadding() As Single
        Get
            Return Me._table.BottomPadding
        End Get
        Set(ByVal value As Single)
            Me._table.BottomPadding = value
        End Set
    End Property

    Public Property LeftPadding() As Single
        Get
            Return Me._table.LeftPadding
        End Get
        Set(ByVal value As Single)
            Me._table.LeftPadding = value
        End Set
    End Property

    Public Property RightPadding() As Single
        Get
            Return Me._table.RightPadding
        End Get
        Set(ByVal value As Single)
            Me._table.RightPadding = value
        End Set
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._table.Range)
        End Get
    End Property

    Public ReadOnly Property Borders() As Borders
        Get
            Return New Borders(Me._table.Borders)
        End Get
    End Property

    Public ReadOnly Property Application() As Application
        Get
            Return New Application(Me._table.Application)
        End Get
    End Property

    Public Sub AutoFitBehavior(ByVal Behavior As WdAutoFitBehavior)
        Call Me._table.AutoFitBehavior(Behavior)
    End Sub
End Class
