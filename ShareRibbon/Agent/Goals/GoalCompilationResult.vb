Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Agent.Goals

    Public NotInheritable Class GoalCoverageMapEntry
        Private ReadOnly _sourceClauseId As String
        Private ReadOnly _criterionIds As ReadOnlyCollection(Of String)

        Public Sub New(sourceClauseId As String, criterionIds As IEnumerable(Of String))
            _sourceClauseId = If(sourceClauseId, "")
            _criterionIds = New List(Of String)(If(criterionIds, New List(Of String)())).AsReadOnly()
        End Sub

        Public ReadOnly Property SourceClauseId As String
            Get
                Return _sourceClauseId
            End Get
        End Property

        Public ReadOnly Property CriterionIds As IReadOnlyList(Of String)
            Get
                Return _criterionIds
            End Get
        End Property
    End Class

    Public NotInheritable Class GoalCompilationResult
        Private ReadOnly _candidate As CandidateGoalContract
        Private ReadOnly _coverageMap As ReadOnlyCollection(Of GoalCoverageMapEntry)
        Private ReadOnly _unresolvedClauses As ReadOnlyCollection(Of String)
        Private ReadOnly _assumptions As ReadOnlyCollection(Of String)
        Private ReadOnly _diagnostics As ReadOnlyCollection(Of String)
        Private ReadOnly _requiresClarification As Boolean

        Public Sub New(candidate As CandidateGoalContract,
                       coverageMap As IEnumerable(Of GoalCoverageMapEntry),
                       unresolvedClauses As IEnumerable(Of String),
                       assumptions As IEnumerable(Of String),
                       diagnostics As IEnumerable(Of String),
                       requiresClarification As Boolean)
            _candidate = candidate
            _coverageMap = New List(Of GoalCoverageMapEntry)(If(coverageMap, New List(Of GoalCoverageMapEntry)())).AsReadOnly()
            _unresolvedClauses = New List(Of String)(If(unresolvedClauses, New List(Of String)())).AsReadOnly()
            _assumptions = New List(Of String)(If(assumptions, New List(Of String)())).AsReadOnly()
            _diagnostics = New List(Of String)(If(diagnostics, New List(Of String)())).AsReadOnly()
            _requiresClarification = requiresClarification
        End Sub

        Public ReadOnly Property Candidate As CandidateGoalContract
            Get
                Return _candidate
            End Get
        End Property

        Public ReadOnly Property CoverageMap As IReadOnlyList(Of GoalCoverageMapEntry)
            Get
                Return _coverageMap
            End Get
        End Property

        Public ReadOnly Property UnresolvedClauses As IReadOnlyList(Of String)
            Get
                Return _unresolvedClauses
            End Get
        End Property

        Public ReadOnly Property Assumptions As IReadOnlyList(Of String)
            Get
                Return _assumptions
            End Get
        End Property

        Public ReadOnly Property Diagnostics As IReadOnlyList(Of String)
            Get
                Return _diagnostics
            End Get
        End Property

        Public ReadOnly Property RequiresClarification As Boolean
            Get
                Return _requiresClarification
            End Get
        End Property
    End Class

End Namespace
