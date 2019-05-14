Public Class Section
    Private _section As Object

    Friend Sub New(ByVal section As Object)
        Me._section = Section
    End Sub

    Public ReadOnly Property Borders() As Borders
        Get
            Return New Borders(Me._section.Borders)
        End Get
    End Property

    Public ReadOnly Property Footers() As HeadersFooters
        Get
            Return New HeadersFooters(Me._section.Footers)
        End Get
    End Property

    Public ReadOnly Property Headers() As HeadersFooters
        Get
            Return New HeadersFooters(Me._section.Headers)
        End Get
    End Property

    Public ReadOnly Property Index() As Integer
        Get
            Return Me._section.Index
        End Get
    End Property

    Public ReadOnly Property PageSetup() As PageSetup
        Get
            Return New PageSetup(Me._section.PageSetup)
        End Get
    End Property

    Public Property ProtectedForForms() As Boolean
        Get
            Return Me._section.ProtectedForForms
        End Get
        Set(ByVal value As Boolean)
            Me._section.ProtectedForForms = value
        End Set
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._section.Range)
        End Get
    End Property
End Class
