' TabStops is a 1 based collection of object TabStop
Public Class TabStops
    Implements IEnumerable(Of TabStop)
    Implements IEnumerator(Of TabStop)
    Implements IDisposable

    Private _tabstops As Object

    Friend Sub New(ByVal tabstops As Object)
        Me._tabstops = tabstops
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._tabstops.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As TabStop
        Get
            Return New TabStop(Me._tabstops.Item(index))
        End Get
    End Property

    Public Function Add(ByVal position As Single)
        Return New TabStop(Me._tabstops.Add(position))
    End Function

    Public Function Add(ByVal position As Single, ByVal alignment As WdTabAlignment)
        Return New TabStop(Me._tabstops.Add(position, alignment))
    End Function

    Public Function Add(ByVal position As Single, ByVal alignment As WdTabAlignment, ByVal leader As WdTabLeader)
        Return New TabStop(Me._tabstops.Add(position, alignment, leader))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfTabStop() As System.Collections.Generic.IEnumerator(Of TabStop) Implements System.Collections.Generic.IEnumerable(Of TabStop).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfTabStop() As TabStop Implements System.Collections.Generic.IEnumerator(Of TabStop).Current
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
