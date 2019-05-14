Public Class Range
    Private _range As Object

    Friend Sub New(ByVal range As Object)
        Me._range = range
    End Sub

    Friend ReadOnly Property underlyingobject()
        Get
            Return Me._range
        End Get
    End Property

    Public Sub Delete()
        Call Me._range.Delete()
    End Sub

    Public Sub InsertAfter(ByVal text As String)
        Call Me._range.InsertAfter(text)
    End Sub

    Public Property Bold() As WdConstants
        Get
            Return Me._range.Bold
        End Get
        Set(ByVal value As WdConstants)
            Me._range.Bold = value
        End Set
    End Property

    Public Property Italic() As WdConstants
        Get
            Return Me._range.Italic
        End Get
        Set(ByVal value As WdConstants)
            Me._range.Italic = value
        End Set
    End Property

    Public Property Underline() As WdUnderline
        Get
            Return Me._range.Underline
        End Get
        Set(ByVal value As WdUnderline)
            Me._range.Underline = value
        End Set
    End Property

    Public Sub [Select]()
        Call Me._range.Select()
    End Sub

    Public Sub InsertParagraphAfter()
        Call Me._range.InsertParagraphAfter()
    End Sub

    Public ReadOnly Property ListFormat() As ListFormat
        Get
            Return New ListFormat(Me._range.ListFormat)
        End Get
    End Property

    Public Property Start() As Integer
        Get
            Return Me._range.Start
        End Get
        Set(ByVal value As Integer)
            Me._range.Start = value
        End Set
    End Property

    Public Property [End]() As Integer
        Get
            Return Me._range.End
        End Get
        Set(ByVal value As Integer)
            Me._range.End = value
        End Set
    End Property

    Public ReadOnly Property Parent() As Document
        Get
            Return New Document(Me._range.Parent)
        End Get
    End Property

    Public ReadOnly Property Tables() As Tables
        Get
            Return New Tables(Me._range.Tables)
        End Get
    End Property

    Public Property Text() As String
        Get
            Return Me._range.Text
        End Get
        Set(ByVal value As String)
            Me._range.Text = value
        End Set
    End Property

    Public ReadOnly Property ParagraphFormat() As ParagraphFormat
        Get
            Return New ParagraphFormat(Me._range.ParagraphFormat)
        End Get
    End Property

    Public ReadOnly Property Cells() As Cells
        Get
            Return New Cells(Me._range.Cells)
        End Get
    End Property

    Public ReadOnly Property Font() As Font
        Get
            Return New Font(Me._range.Font)
        End Get
    End Property

    Public ReadOnly Property Fields() As Fields
        Get
            Return New Fields(Me._range.Fields)
        End Get
    End Property

    Public Function ConvertToTable() As Table
        Return New Table(Me._range.ConvertToTable())
    End Function

    Public Function ConvertToTable(ByVal Separator As String) As Table
        Return New Table(Me._range.ConvertToTable(Separator))
    End Function

    Public Function ConvertToTable(ByVal Separator As WdTableFieldSeparator) As Table
        Return New Table(Me._range.ConvertToTable(Separator))
    End Function

    Public Sub Collapse()
        Call Me._range.Collapse()
    End Sub

    Public Sub Collapse(ByVal Direction As WdCollapseDirection)
        Call Me._range.Collapse(Direction)
    End Sub

    Public ReadOnly Property Find() As Find
        Get
            Return New Find(Me._range.Find)
        End Get
    End Property
End Class
