Public Class TabStop
    Private _tabstop As Object

    Friend Sub New(ByVal tabstop As Object)
        Me._tabstop = tabstop
    End Sub

    Public Property Alignment() As WdTabAlignment
        Get
            Return Me._tabstop.Alignment
        End Get
        Set(ByVal value As WdTabAlignment)
            Me._tabstop.Alignment = value
        End Set
    End Property

    Public ReadOnly Property CustomTab() As Boolean
        Get
            Return Me._tabstop.CustomTab
        End Get
    End Property

    Public Property Leader() As WdTabLeader
        Get
            Return Me._tabstop.Leader
        End Get
        Set(ByVal value As WdTabLeader)
            Me._tabstop.Leader = value
        End Set
    End Property

    Public ReadOnly Property [Next]() As TabStop
        Get
            Return New TabStop(Me._tabstop.[Next])
        End Get
    End Property

    Public Property Position() As Single
        Get
            Return Me._tabstop.Position
        End Get
        Set(ByVal value As Single)
            Me._tabstop.Position = value
        End Set
    End Property

    Public ReadOnly Property Previous() As TabStop
        Get
            Return New TabStop(Me._tabstop.Previous)
        End Get
    End Property

    Public Sub Clear()
        Call Me._tabstop.Clear()
    End Sub
End Class
