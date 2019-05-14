﻿Public Class DocumentProperty
    Private ReadOnly _documentproperty As Object

    Friend Sub New(ByVal documentproperty As Object)
        Me._documentproperty = documentproperty
    End Sub

    Public Property Name() As String
        Get
            Return Me._documentproperty.Name
        End Get
        Set(ByVal value As String)
            Me._documentproperty.Name = value
        End Set
    End Property

    Public Property Type() As MsoDocProperties
        Get
            Return Me._documentproperty.Type
        End Get
        Set(ByVal value As MsoDocProperties)
            Me._documentproperty.Type = value
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return Me._documentproperty.Value
        End Get
        Set(ByVal value As Object)
            Me._documentproperty.Value = value
        End Set
    End Property

    Public Sub Delete()
        Call Me._documentproperty.Delete()
    End Sub

    Public Property LinkToContent() As Boolean
        Get
            Return Me._documentproperty.LinkToContent
        End Get
        Set(ByVal value As Boolean)
            Me._documentproperty.LinkToContent = value
        End Set
    End Property

    Public Property LinkSource() As String
        Get
            Return Me._documentproperty.LinkSource
        End Get
        Set(ByVal value As String)
            Me._documentproperty.LinkSource = value
        End Set
    End Property
End Class
