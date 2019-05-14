' Tables is a 1 based collection of object Table
Public Class Tables
    Implements IEnumerable(Of Table)
    Implements IEnumerator(Of Table)
    Implements IDisposable

    Private _tables As Object

    Friend Sub New(ByVal tables As Object)
        Me._tables = tables
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._tables.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Table
        Get
            Return New Table(Me._tables.Item(index))
        End Get
    End Property

    Public Function Add(ByVal range As Range, ByVal countofrows As Integer, ByVal countofcolumns As Integer) As Table
        Return New Table(Me._tables.Add(range.underlyingobject, countofrows, countofcolumns))
    End Function

    Public Function Add(ByVal range As Range, ByVal countofrows As Integer, ByVal countofcolumns As Integer, ByVal defaulttablebehaviour As WdDefaultTableBehavior) As Table
        Return New Table(Me._tables.Add(range.underlyingobject, countofrows, countofcolumns, defaulttablebehaviour))
    End Function

    Public Function Add(ByVal range As Range, ByVal countofrows As Integer, ByVal countofcolumns As Integer, ByVal defaulttablebehaviour As WdDefaultTableBehavior, ByVal autofitbehaviour As WdAutoFitBehavior) As Table
        Return New Table(Me._tables.Add(range.underlyingobject, countofrows, countofcolumns, defaulttablebehaviour, autofitbehaviour))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfTable() As System.Collections.Generic.IEnumerator(Of Table) Implements System.Collections.Generic.IEnumerable(Of Table).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfTable() As Table Implements System.Collections.Generic.IEnumerator(Of Table).Current
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
