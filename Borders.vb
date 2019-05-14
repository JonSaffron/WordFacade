' Borders is a 1 based collection of object Border
Public Class Borders
    Implements IEnumerable(Of Border)
    Implements IEnumerator(Of Border)
    Implements IDisposable

    Private ReadOnly _borders As Object

    Friend Sub New(ByVal borders As Object)
        Me._borders = borders
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._borders.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As WdBorderType) As Border
        Get
            Return New Border(Me._borders.Item(index))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfBorder() As System.Collections.Generic.IEnumerator(Of Border) Implements System.Collections.Generic.IEnumerable(Of Border).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfBorder() As Border Implements System.Collections.Generic.IEnumerator(Of Border).Current
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
