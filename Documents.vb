' Documents is a 1 based collection of object Document
Public Class Documents
    Implements IEnumerable(Of Document)
    Implements IEnumerator(Of Document)
    Implements IDisposable

    Private _documents As Object

    Friend Sub New(ByVal documents As Object)
        Me._documents = documents
        Call Me.Reset()
    End Sub

    Public Function Open(ByVal filename As String) As Document
        Return New Document(Me._documents.Open(filename))
    End Function

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._documents.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Document
        Get
            Return New Document(Me._documents.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal documentname As String) As Document
        Get
            Return New Document(Me._documents.Item(documentname))
        End Get
    End Property

    Public Function Add() As Document
        Return New Document(Me._documents.Add())
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfDocument() As System.Collections.Generic.IEnumerator(Of Document) Implements System.Collections.Generic.IEnumerable(Of Document).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfDocument() As Document Implements System.Collections.Generic.IEnumerator(Of Document).Current
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
