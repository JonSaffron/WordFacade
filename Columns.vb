' Columns is a 1 based collection of object Column
Public Class Columns
    Implements IEnumerable(Of Column)
    Implements IEnumerator(Of Column)
    Implements IDisposable

    Private ReadOnly _columns As Object

    Friend Sub New(ByVal columns As Object)
        Me._columns = columns
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._columns.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Column
        Get
            Return New Column(Me._columns.Item(index))
        End Get
    End Property

    Public Sub AutoFit()
        Call Me._columns.AutoFit()
    End Sub

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfColumn() As System.Collections.Generic.IEnumerator(Of Column) Implements System.Collections.Generic.IEnumerable(Of Column).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfColumn() As Column Implements System.Collections.Generic.IEnumerator(Of Column).Current
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
