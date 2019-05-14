﻿Public Class COMAddIns
    Implements IEnumerable(Of COMAddIn)
    Implements IEnumerator(Of COMAddIn)
    Implements IDisposable

    Private ReadOnly _COMAddIns As Object

    Friend Sub New(ByVal COMAddIns As Object)
        Me._COMAddIns = COMAddIns
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Return Me._COMAddIns.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As COMAddIn
        Get
            Return New COMAddIn(Me._COMAddIns.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal progId As String) As COMAddIn
        Get
            Return New COMAddIn(Me._COMAddIns.Item(progId))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfCOMAddIn() As System.Collections.Generic.IEnumerator(Of COMAddIn) Implements System.Collections.Generic.IEnumerable(Of COMAddIn).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfCOMAddIn() As COMAddIn Implements System.Collections.Generic.IEnumerator(Of COMAddIn).Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        Me._enumeratorPosition += 1
        Return (Me._enumeratorPosition <= Me.Count)
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
        Me._enumeratorPosition = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Private _isDisposed As Boolean = False

    ' IDisposable
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
        Call Dispose(True)
        Call GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
