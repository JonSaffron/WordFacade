Public Class Replacement
    Private _replacement As Object

    Friend Sub New(ByVal replacement As Object)
        Me._replacement = replacement
    End Sub

    Public Sub ClearFormatting()
        Call Me._replacement.ClearFormatting()
    End Sub

    Public Property Text() As String
        Get
            Return Me._replacement.Text
        End Get
        Set(ByVal value As String)
            Me._replacement.Text = value
        End Set
    End Property

    Public Property NoProofing() As WdConstants
        Get
            Return Me._replacement.NoProofing
        End Get
        Set(ByVal value As WdConstants)
            Me._replacement.NoProofing = value
        End Set
    End Property
End Class
