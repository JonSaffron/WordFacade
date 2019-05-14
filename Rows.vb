' Rows is a 1 based collection of object Row
Public Class Rows
    Implements IEnumerable(Of Row)
    Implements IEnumerator(Of Row)
    Implements IDisposable

    Private _rows As Object

    Friend Sub New(ByVal rows As Object)
        Me._rows = rows
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._rows.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Row
        Get
            Return New Row(Me._rows.Item(index))
        End Get
    End Property

    Public Function Add() As Row
        Return New Row(Me._rows.Add())
    End Function

    Public Function Add(ByVal beforeRow As Row) As Row
        Return New Row(Me._rows.Add(beforeRow.underlyingobject))
    End Function

    Public Property LeftIndent() As Single
        Get
            Return Me._rows.LeftIndent
        End Get
        Set(ByVal value As Single)
            Me._rows.LeftIndent = value
        End Set
    End Property

    Public Property AllowBreakAcrossPages() As WdConstants
        Get
            Return Me._rows.AllowBreakAcrossPages
        End Get
        Set(ByVal value As WdConstants)
            Me._rows.AllowBreakAcrossPages = value
        End Set
    End Property

    Public ReadOnly Property Shading() As Shading
        Get
            Return New Shading(Me._rows.Shading)
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfRow() As System.Collections.Generic.IEnumerator(Of Row) Implements System.Collections.Generic.IEnumerable(Of Row).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfRow() As Row Implements System.Collections.Generic.IEnumerator(Of Row).Current
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
