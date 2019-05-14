Public Class ListFormat
    Private _listformat As Object

    Friend Sub New(ByVal listformat As Object)
        Me._listformat = listformat
    End Sub

    Public Sub ApplyBulletDefault()
        Call Me._listformat.ApplyBulletDefault()
    End Sub

    Public Sub ApplyNumberDefault()
        Call Me._listformat.ApplyNumberDefault()
    End Sub
End Class
