Public Class Document
    Private _document As Object

    Friend Sub New(ByVal document As Object)
        Me._document = document
    End Sub

    Public Sub Save()
        Call Me._document.Save()
    End Sub

    <Obsolete("Use override specifying fileformat. In Word 2007 you should ensure that the extension is appropriate to the fileformat.")> _
    Public Sub SaveAs(ByVal FileName As String)
        Call Me._document.SaveAs(filename)
    End Sub

    Public Sub SaveAs(ByVal FileName As String, ByVal FileFormat As WdSaveFormat)
        Call Me._document.SaveAs(FileName, FileFormat)
    End Sub

    Public Sub Close()
        Call Me._document.Close()
    End Sub

    Public Property Saved() As Boolean
        Get
            Return Me._document.Saved
        End Get
        Set(ByVal value As Boolean)
            Me._document.Saved = value
        End Set
    End Property

    Public ReadOnly Property Bookmarks() As Bookmarks
        Get
            Return New Bookmarks(Me._document.Bookmarks)
        End Get
    End Property

    Public Function Range(ByVal start As Integer, ByVal [end] As Integer) As Range
        Return New Range(Me._document.Range(start, [end]))
    End Function

    Public ReadOnly Property InlineShapes() As InlineShapes
        Get
            Return New InlineShapes(Me._document.InlineShapes)
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return Me._document.Name
        End Get
    End Property

    Public Sub Repaginate()
        Call Me._document.Repaginate()
    End Sub

    Public ReadOnly Property ActiveWindow() As Window
        Get
            Return New Window(Me._document.ActiveWindow)
        End Get
    End Property

    Public ReadOnly Property Tables() As Tables
        Get
            Return New Tables(Me._document.Tables)
        End Get
    End Property

    Public Property ShowSpellingErrors() As Boolean
        Get
            Return Me._document.ShowSpellingErrors
        End Get
        Set(ByVal value As Boolean)
            Me._document.ShowSpellingErrors = value
        End Set
    End Property

    Public Property ShowGrammaticalErrors() As Boolean
        Get
            Return Me._document.ShowGrammaticalErrors
        End Get
        Set(ByVal value As Boolean)
            Me._document.ShowGrammaticalErrors = value
        End Set
    End Property

    Public ReadOnly Property PageSetup() As PageSetup
        Get
            Return New PageSetup(Me._document.PageSetup)
        End Get
    End Property

    Public ReadOnly Property Application() As Application
        Get
            Return New Application(Me._document.Application)
        End Get
    End Property

    Public ReadOnly Property BuiltInDocumentProperties() As DocumentProperties
        Get
            Return New DocumentProperties(Me._document.BuiltInDocumentProperties)
        End Get
    End Property

    Public ReadOnly Property CustomDocumentProperties() As DocumentProperties
        Get
            Return New DocumentProperties(Me._document.CustomDocumentProperties)
        End Get
    End Property

    Public ReadOnly Property Sections() As Sections
        Get
            Return New Sections(Me._document.Sections)
        End Get
    End Property

    Public ReadOnly Property FullName() As String
        Get
            Return Me._document.FullName
        End Get
    End Property

    Public ReadOnly Property Path() As String
        Get
            Return Me._document.Path
        End Get
    End Property

    Public ReadOnly Property SaveFormat() As WdSaveFormat
        Get
            Return Me._document.FileFormat
        End Get
    End Property

    Public ReadOnly Property Content() As Range
        Get
            Return New Range(Me._document.Content)
        End Get
    End Property

    Public Sub UndoClear()
        Call Me._document.UndoClear()
    End Sub

    Public ReadOnly Property Styles() As Styles
        Get
            Return New Styles(Me._document.Styles)
        End Get
    End Property
End Class
