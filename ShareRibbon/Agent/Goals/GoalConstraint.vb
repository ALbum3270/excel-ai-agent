Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Agent.Goals

    Public NotInheritable Class GoalConstraint
        Private ReadOnly _id As String
        Private ReadOnly _statement As String
        Private ReadOnly _kind As String
        Private ReadOnly _sourceClauseIds As ReadOnlyCollection(Of String)
        Private ReadOnly _required As Boolean

        Friend Sub New(id As String,
                       statement As String,
                       kind As String,
                       sourceClauseIds As IEnumerable(Of String),
                       required As Boolean)
            _id = id
            _statement = statement
            _kind = kind
            _sourceClauseIds = New List(Of String)(If(sourceClauseIds, New List(Of String)())).AsReadOnly()
            _required = required
        End Sub

        Public ReadOnly Property Id As String
            Get
                Return _id
            End Get
        End Property

        Public ReadOnly Property Statement As String
            Get
                Return _statement
            End Get
        End Property

        Public ReadOnly Property Kind As String
            Get
                Return _kind
            End Get
        End Property

        Public ReadOnly Property SourceClauseIds As IReadOnlyList(Of String)
            Get
                Return _sourceClauseIds
            End Get
        End Property

        Public ReadOnly Property Required As Boolean
            Get
                Return _required
            End Get
        End Property
    End Class

    Public Class CandidateGoalConstraint
        Public Property Id As String = ""
        Public Property Statement As String = ""
        Public Property Kind As String = "policy"
        Public Property SourceClauseIds As New List(Of String)()
        Public Property Required As Boolean = True
    End Class

End Namespace
