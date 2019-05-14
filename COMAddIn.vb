Public Class COMAddIn
    Private _COMAddIn As Object

    Friend Sub New(ByVal COMAddIn As Object)
        Me._COMAddIn = COMAddIn
    End Sub

    Public Property Connect() As Boolean
        Get
            Return Me._COMAddIn.Connect
        End Get
        Set(ByVal value As Boolean)
            Me._COMAddIn.Connect = value
        End Set
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return Me._COMAddIn.Description
        End Get
    End Property

    Public ReadOnly Property Guid() As Guid
        Get
            Dim g As String = Me._COMAddIn.Guid
            Return New Guid(g)
        End Get
    End Property

    Public ReadOnly Property ProdId() As String
        Get
            Return Me._COMAddIn.ProgId
        End Get
    End Property
End Class
