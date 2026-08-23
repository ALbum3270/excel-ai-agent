Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Agent.Goals

    ''' <summary>
    ''' Mutable interpretation candidate.  This type is intentionally distinct from the
    ''' frozen authoritative goal.
    ''' </summary>
    Public Class CandidateGoalContract
        Public Property RawUserRequest As String = ""
        Public Property SourceClauses As New List(Of CandidateGoalSourceClause)()
        Public Property Criteria As New List(Of CandidateGoalCriterion)()
        Public Property Constraints As New List(Of CandidateGoalConstraint)()
        Public Property RequiredCapabilities As New List(Of String)()
    End Class

    ''' <summary>
    ''' Authoritative user-goal semantics.  The interface has no mutation or replacement
    ''' operations; all collections are defensive read-only copies created by the freezer.
    ''' </summary>
    Public NotInheritable Class GoalContract
        Private ReadOnly _goalId As String
        Private ReadOnly _rawUserRequest As String
        Private ReadOnly _sourceClauses As ReadOnlyCollection(Of GoalSourceClause)
        Private ReadOnly _criteria As ReadOnlyCollection(Of GoalCriterion)
        Private ReadOnly _constraints As ReadOnlyCollection(Of GoalConstraint)
        Private ReadOnly _requiredCapabilities As ReadOnlyCollection(Of String)
        Private ReadOnly _contractHash As String
        Private ReadOnly _semanticHash As String

        Friend Sub New(goalId As String,
                       rawUserRequest As String,
                       sourceClauses As IEnumerable(Of GoalSourceClause),
                       criteria As IEnumerable(Of GoalCriterion),
                       constraints As IEnumerable(Of GoalConstraint),
                       requiredCapabilities As IEnumerable(Of String),
                       contractHash As String,
                       semanticHash As String)
            _goalId = goalId
            _rawUserRequest = rawUserRequest
            _sourceClauses = New List(Of GoalSourceClause)(If(sourceClauses, New List(Of GoalSourceClause)())).AsReadOnly()
            _criteria = New List(Of GoalCriterion)(If(criteria, New List(Of GoalCriterion)())).AsReadOnly()
            _constraints = New List(Of GoalConstraint)(If(constraints, New List(Of GoalConstraint)())).AsReadOnly()
            _requiredCapabilities = New List(Of String)(If(requiredCapabilities, New List(Of String)())).AsReadOnly()
            _contractHash = contractHash
            _semanticHash = semanticHash
        End Sub

        Public ReadOnly Property GoalId As String
            Get
                Return _goalId
            End Get
        End Property

        Public ReadOnly Property RawUserRequest As String
            Get
                Return _rawUserRequest
            End Get
        End Property

        Public ReadOnly Property SourceClauses As IReadOnlyList(Of GoalSourceClause)
            Get
                Return _sourceClauses
            End Get
        End Property

        Public ReadOnly Property Criteria As IReadOnlyList(Of GoalCriterion)
            Get
                Return _criteria
            End Get
        End Property

        Public ReadOnly Property Constraints As IReadOnlyList(Of GoalConstraint)
            Get
                Return _constraints
            End Get
        End Property

        Public ReadOnly Property RequiredCapabilities As IReadOnlyCollection(Of String)
            Get
                Return _requiredCapabilities
            End Get
        End Property

        Public ReadOnly Property ContractHash As String
            Get
                Return _contractHash
            End Get
        End Property

        ''' <summary>
        ''' Canonical semantic graph identity. It ignores model-generated entity ids and list
        ''' order, but preserves raw authority, node multiplicity, typed fields and reference
        ''' topology. Runtime immutability continues to use ContractHash.
        ''' </summary>
        Public ReadOnly Property SemanticHash As String
            Get
                Return _semanticHash
            End Get
        End Property
    End Class

End Namespace
