Public Class ParagraphFormat
    Private ReadOnly _paragraphformat As Object

    Friend Sub New(ByVal paragraphformat As Object)
        Me._paragraphformat = paragraphformat
    End Sub

    Friend ReadOnly Property underlyingComObject() As Object
        Get
            Return Me._paragraphformat
        End Get
    End Property

    Public Property Alignment() As WdParagraphAlignment
        Get
            Return Me._paragraphformat.Alignment
        End Get
        Set(ByVal value As WdParagraphAlignment)
            Me._paragraphformat.Alignment = value
        End Set
    End Property

    Public ReadOnly Property TabStops() As TabStops
        Get
            Return New TabStops(Me._paragraphformat.TabStops)
        End Get
    End Property

    Public Property SpaceBefore() As Single
        Get
            Return Me._paragraphformat.SpaceBefore
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.SpaceBefore = value
        End Set
    End Property

    Public Property SpaceAfter() As Single
        Get
            Return Me._paragraphformat.SpaceAfter
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.SpaceAfter = value
        End Set
    End Property

    Public Property LeftIndent() As Single
        Get
            Return Me._paragraphformat.LeftIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.LeftIndent = value
        End Set
    End Property

    Public Property RightIndent() As Single
        Get
            Return Me._paragraphformat.RightIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.RightIndent = value
        End Set
    End Property

    Public Property SpaceBeforeAuto() As WdConstants
        Get
            Return Me._paragraphformat.SpaceBeforeAuto
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.SpaceBeforeAuto = value
        End Set
    End Property

    Public Property SpaceAfterAuto() As WdConstants
        Get
            Return Me._paragraphformat.SpaceAfterAuto
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.SpaceAfterAuto = value
        End Set
    End Property

    Public Property LineSpacingRule() As WdLineSpacing
        Get
            Return Me._paragraphformat.LineSpacingRule
        End Get
        Set(ByVal value As WdLineSpacing)
            Me._paragraphformat.LineSpacingRule = value
        End Set
    End Property

    Public Property WidowControl() As WdConstants
        Get
            Return Me._paragraphformat.WidowControl
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.WidowControl = value
        End Set
    End Property

    Public Property KeepWithNext() As WdConstants
        Get
            Return Me._paragraphformat.KeepWithNext
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.KeepWithNext = value
        End Set
    End Property

    Public Property KeepTogether() As WdConstants
        Get
            Return Me._paragraphformat.KeepTogether
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.KeepTogether = value
        End Set
    End Property

    Public Property PageBreakBefore() As WdConstants
        Get
            Return Me._paragraphformat.PageBreakBefore
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.PageBreakBefore = value
        End Set
    End Property

    Public Property NoLineNumber() As WdConstants
        Get
            Return Me._paragraphformat.NoLineNumber
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.NoLineNumber = value
        End Set
    End Property

    Public Property Hyphenation() As WdConstants
        Get
            Return Me._paragraphformat.Hyphenation
        End Get
        Set(ByVal value As WdConstants)
            Me._paragraphformat.Hyphenation = value
        End Set
    End Property

    Public Property FirstLineIndent() As Single
        Get
            Return Me._paragraphformat.FirstLineIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.FirstLineIndent = value
        End Set
    End Property

    Public Property OutlineLevel() As WdOutlineLevel
        Get
            Return Me._paragraphformat.OutlineLevel
        End Get
        Set(ByVal value As WdOutlineLevel)
            Me._paragraphformat.OutlineLevel = value
        End Set
    End Property

    Public Property CharacterUnitLeftIndent() As Single
        Get
            Return Me._paragraphformat.CharacterUnitLeftIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.CharacterUnitLeftIndent = value
        End Set
    End Property

    Public Property CharacterUnitRightIndent() As Single
        Get
            Return Me._paragraphformat.CharacterUnitRightIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.CharacterUnitRightIndent = value
        End Set
    End Property

    Public Property CharacterUnitFirstLineIndent() As Single
        Get
            Return Me._paragraphformat.CharacterUnitFirstLineIndent
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.CharacterUnitFirstLineIndent = value
        End Set
    End Property

    Public Property LineUnitBefore() As Single
        Get
            Return Me._paragraphformat.LineUnitBefore
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.LineUnitBefore = value
        End Set
    End Property

    Public Property LineUnitAfter() As Single
        Get
            Return Me._paragraphformat.LineUnitAfter
        End Get
        Set(ByVal value As Single)
            Me._paragraphformat.LineUnitAfter = value
        End Set
    End Property

End Class
