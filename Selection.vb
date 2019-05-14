Public Class Selection
    Private _selection As Object

    Friend Sub New(ByVal selection As Object)
        Me._selection = selection
    End Sub

    Public Function HomeKey() As Integer
        Return Me._selection.HomeKey()
    End Function

    Public Function HomeKey(ByVal unit As WdUnits) As Integer
        Return Me._selection.HomeKey(unit)
    End Function

    Public Function HomeKey(ByVal unit As WdUnits, ByVal extend As WdMovementType) As Integer
        Return Me._selection.HomeKey(unit, extend)
    End Function

    Public ReadOnly Property ParagraphFormat() As ParagraphFormat
        Get
            Return New ParagraphFormat(Me._selection.ParagraphFormat)
        End Get
    End Property

    Public ReadOnly Property Font() As font
        Get
            Return New Font(Me._selection.Font)
        End Get
    End Property

    Public Sub TypeText(ByVal text As String)
        Call Me._selection.TypeText(text)
    End Sub

    Public Sub TypeParagraph()
        Call Me._selection.TypeParagraph()
    End Sub

    Public Function EndKey() As Integer
        Return Me._selection.EndKey
    End Function

    Public Function EndKey(ByVal unit As WdUnits)
        Return Me._selection.EndKey(unit)
    End Function

    Public Function EndKey(ByVal unit As WdUnits, ByVal extend As WdMovementType)
        Return Me._selection.EndKey(unit, extend)
    End Function

    Public Sub InsertBreak()
        Call Me._selection.InsertBreak()
    End Sub

    Public Sub InsertBreak(ByVal type As WdBreakType)
        Call Me._selection.InsertBreak(type)
    End Sub

    Public ReadOnly Property Application() As Application
        Get
            Return New Application(Me._selection.Application)
        End Get
    End Property

    Public ReadOnly Property Range() As Range
        Get
            Return New Range(Me._selection.Range)
        End Get
    End Property

    Public ReadOnly Property Cells() As Cells
        Get
            Return New Cells(Me._selection.Cells)
        End Get
    End Property

    Public ReadOnly Property Find() As Find
        Get
            Return New Find(Me._selection.Find)
        End Get
    End Property

    Public Property Style() As Style
        Get
            Return New Style(Me._selection.Style)
        End Get
        Set(ByVal value As Style)
            Me._selection.Style = value.underlyingComObject
        End Set
    End Property

    Public Sub StyleSet(ByVal StyleName As String)
        Me._selection.Style = StyleName
    End Sub

    Public Sub StyleSet(ByVal StyleNumber As Integer)
        Me._selection.Style = StyleNumber
    End Sub

    Public Sub StyleSet(ByVal BuiltInStyle As WdBuiltInStyle)
        Me._selection.Style = BuiltInStyle
    End Sub
End Class
