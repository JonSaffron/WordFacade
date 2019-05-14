' InlineShapes is a 1 based collection of object InlineShape
Public Class InlineShapes
    Implements IEnumerable(Of InlineShape)
    Implements IEnumerator(Of InlineShape)
    Implements IDisposable

    Private _inlineshapes As Object

    Friend Sub New(ByVal inlineshapes As Object)
        Me._inlineshapes = inlineshapes
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._inlineshapes.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As InlineShape
        Get
            Return New InlineShape(Me._inlineshapes.item(index))
        End Get
    End Property

    Public Function AddPicture(ByVal filename As String, ByVal linktofile As Boolean, ByVal savewithdocument As Boolean, ByVal range As Range) As InlineShape
        Return New InlineShape(Me._inlineshapes.AddPicture(filename, linktofile, savewithdocument, range.underlyingobject))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfInlineShape() As System.Collections.Generic.IEnumerator(Of InlineShape) Implements System.Collections.Generic.IEnumerable(Of InlineShape).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfInlineShape() As InlineShape Implements System.Collections.Generic.IEnumerator(Of InlineShape).Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        Me._enumeratorPostion += 1
        Return (Me._enumeratorPostion <= Me.Count)
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
        Me._enumeratorPostion = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Private _isDisposed As Boolean = False

    Protected Overridable Sub Dispose(ByVal isSafeToDisposeManagedResources As Boolean)
        If Not Me._isDisposed Then
            If isSafeToDisposeManagedResources Then
                ' no managed resources to free
            End If

            ' no shared unmanaged resources to free
        End If
        Me._isDisposed = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Call Me.Dispose(True)
        Call GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
