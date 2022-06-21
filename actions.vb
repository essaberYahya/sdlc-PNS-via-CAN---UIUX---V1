Imports System.Collections.Generic

Public Class actions
    Dim procname As String
    Dim procDescription As String
    Dim procSerial As String
    Dim procState As Integer
    Dim procCreatedAt As DateTime
    Public listActions As New ArrayList()

    Public Sub New(procname As String, procDescription As String, procSerial As String, procState As Integer, procCreatedAt As DateTime)
        'Me.procname = procname.Substring(0, 1).ToUpper() + procname.Substring(1).ToLower()
        Me.procname = procname.ToUpper()
        'Me.procDescription = procDescription.Substring(0, 1).ToUpper() + procDescription.Substring(1).ToLower()
        Me.procDescription = procDescription.ToUpper()
        Me.procSerial = procSerial
        Me.procState = procState
        Me.procCreatedAt = procCreatedAt
    End Sub

    Public Property Procname1 As String
        Get
            Return procname
        End Get
        Set(value As String)
            procname = value
        End Set
    End Property

    Public Property ProcDescription1 As String
        Get
            Return procDescription
        End Get
        Set(value As String)
            procDescription = value
        End Set
    End Property

    Public Property ProcSerial1 As String
        Get
            Return procSerial
        End Get
        Set(value As String)
            procSerial = value
        End Set
    End Property

    Public Property ProcState1 As Integer
        Get
            Return procState
        End Get
        Set(value As Integer)
            procState = value
        End Set
    End Property

    Public Property ProcCreatedAt1 As Date
        Get
            Return procCreatedAt
        End Get
        Set(value As Date)
            procCreatedAt = value
        End Set
    End Property

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim actions = TryCast(obj, actions)
        Return actions IsNot Nothing AndAlso
               procname = actions.procname AndAlso
               procDescription = actions.procDescription AndAlso
               procSerial = actions.procSerial AndAlso
               procState = actions.procState AndAlso
               procCreatedAt = actions.procCreatedAt AndAlso
               EqualityComparer(Of Array).Default.Equals(listActions, actions.listActions)
    End Function



End Class
