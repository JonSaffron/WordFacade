' Cells is a 1 based collection of object Cell
Public Class Cells
    Implements IEnumerable(Of Cell)
    Implements IEnumerator(Of Cell)
    Implements IDisposable

    Private ReadOnly _cells As Object

    Friend Sub New(ByVal cells As Object)
        Me._cells = cells
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._cells.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Cell
        Get
            Return New Cell(Me._cells.Item(index))
        End Get
    End Property

    Public Property VerticalAlignment() As WdCellVerticalAlignment
        Get
            Return Me._cells.VerticalAlignment
        End Get
        Set(ByVal value As WdCellVerticalAlignment)
            Me._cells.VerticalAlignment = value
        End Set
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfCell() As System.Collections.Generic.IEnumerator(Of Cell) Implements System.Collections.Generic.IEnumerable(Of Cell).GetEnumerator
        Me._collection = New List(Of Cell)(Me.Count)
        For i As Integer = 1 To Me.Count
            Call Me._collection.Add(Nothing)
        Next
        'For Each cell As Object In Me._cells
        '    Me._collection.Add(New Cell(cell))
        'Next
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return GetEnumeratorOfCell()
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer
    Private _collection As System.Collections.Generic.List(Of Cell)

    Public ReadOnly Property CurrentOfCell() As Cell Implements System.Collections.Generic.IEnumerator(Of Cell).Current
        Get
            Dim Result As Cell = Me._collection.Item(Me._enumeratorPostion)
            If Result Is Nothing Then
                Result = New Cell(Me._cells.Item(Me._enumeratorPostion + 1))
                Me._collection.Item(Me._enumeratorPostion) = Result
            End If
            Return Result


            'Return New Cell(Me._collection.Item(Me._enumeratorPostion))
        End Get
    End Property

    Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
        Get
            Return Me.CurrentOfCell()
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        Me._enumeratorPostion += 1
        Return (Me._enumeratorPostion < Me._collection.Count)
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
        Me._enumeratorPostion = -1
        Me._collection = Nothing
    End Sub
#End Region

#Region " IDisposable Support "
    Private _isDisposed As Boolean = False

    Protected Overridable Sub Dispose(ByVal isSafeToDisposeManagedResources As Boolean)
        If Not Me._isDisposed Then
            If isSafeToDisposeManagedResources Then
                Me._collection = Nothing
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
