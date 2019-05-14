' Sections is a 1 based collection of object Section
Public Class Sections
    Implements IEnumerable(Of Section)
    Implements IEnumerator(Of Section)
    Implements IDisposable

    Private _sections As Object

    Friend Sub New(ByVal sections As Object)
        Me._sections = sections
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._sections.Count
        End Get
    End Property

    Public ReadOnly Property First() As Section
        Get
            Return New Section(Me._sections.First)
        End Get
    End Property

    Public ReadOnly Property Last() As Section
        Get
            Return New Section(Me._sections.Last)
        End Get
    End Property

    Public ReadOnly Property PageSetup() As PageSetup
        Get
            Return New PageSetup(Me._sections.PageSetup)
        End Get
    End Property

    Public Function Add() As Section
        Return New Section(Me._sections.Add())
    End Function

    Public Function Add(ByVal Range As Range) As Section
        Return New Section(Me._sections.Add(Range.underlyingobject))
    End Function

    Public Function Add(ByVal Range As Range, ByVal Start As WdSectionStart) As Section
        Return New Section(Me._sections.add(Range.underlyingobject, Start))
    End Function

    Default Public ReadOnly Property Item(ByVal index As Integer) As Section
        Get
            Return New Section(Me._sections.Item(index))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfSection() As System.Collections.Generic.IEnumerator(Of Section) Implements System.Collections.Generic.IEnumerable(Of Section).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfSection() As Section Implements System.Collections.Generic.IEnumerator(Of Section).Current
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
