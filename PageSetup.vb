Public Class PageSetup
    Private _pagesetup As Object

    Friend Sub New(ByVal pagesetup As Object)
        Me._pagesetup = pagesetup
    End Sub

    Public Property Orientation() As WdOrientation
        Get
            Return Me._pagesetup.Orientation
        End Get
        Set(ByVal value As WdOrientation)
            Me._pagesetup.Orientation = value
        End Set
    End Property

    Public Property TopMargin() As Single
        Get
            Return Me._pagesetup.TopMargin
        End Get
        Set(ByVal value As Single)
            Me._pagesetup.TopMargin = value
        End Set
    End Property

    Public Property BottomMargin() As Single
        Get
            Return Me._pagesetup.BottomMargin
        End Get
        Set(ByVal value As Single)
            Me._pagesetup.BottomMargin = value
        End Set
    End Property

    Public Property LeftMargin() As Single
        Get
            Return Me._pagesetup.LeftMargin
        End Get
        Set(ByVal value As Single)
            Me._pagesetup.LeftMargin = value
        End Set
    End Property

    Public Property RightMargin() As Single
        Get
            Return Me._pagesetup.RightMargin
        End Get
        Set(ByVal value As Single)
            Me._pagesetup.RightMargin = value
        End Set
    End Property
End Class
