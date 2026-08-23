Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Agent.Goals

    Public NotInheritable Class GoalSourceClause
        Private ReadOnly _id As String
        Private ReadOnly _text As String
        Private ReadOnly _explicit As Boolean
        Private ReadOnly _requiredCapability As String

        Friend Sub New(id As String,
                       text As String,
                       isExplicit As Boolean,
                       requiredCapability As String)
            _id = id
            _text = text
            _explicit = isExplicit
            _requiredCapability = requiredCapability
        End Sub

        Public ReadOnly Property Id As String
            Get
                Return _id
            End Get
        End Property

        Public ReadOnly Property Text As String
            Get
                Return _text
            End Get
        End Property

        Public ReadOnly Property IsExplicit As Boolean
            Get
                Return _explicit
            End Get
        End Property

        Public ReadOnly Property RequiredCapability As String
            Get
                Return _requiredCapability
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Mutable model/intake representation.  It is never accepted as the authoritative goal.
    ''' </summary>
    Public Class CandidateGoalSourceClause
        Public Property Id As String = ""
        Public Property Text As String = ""
        Public Property IsExplicit As Boolean = True
        Public Property RequiredCapability As String = ""
    End Class

    Public NotInheritable Class GoalCriterion
        Private ReadOnly _id As String
        Private ReadOnly _statement As String
        Private ReadOnly _kind As String
        Private ReadOnly _sourceClauseIds As ReadOnlyCollection(Of String)
        Private ReadOnly _required As Boolean
        Private ReadOnly _verificationCapability As String
        Private ReadOnly _capabilityId As String

        Friend Sub New(id As String,
                       statement As String,
                       kind As String,
                       sourceClauseIds As IEnumerable(Of String),
                       required As Boolean,
                       verificationCapability As String,
                       capabilityId As String)
            _id = id
            _statement = statement
            _kind = kind
            _sourceClauseIds = New List(Of String)(If(sourceClauseIds, New List(Of String)())).AsReadOnly()
            _required = required
            _verificationCapability = verificationCapability
            _capabilityId = capabilityId
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

        Public ReadOnly Property VerificationCapability As String
            Get
                Return _verificationCapability
            End Get
        End Property

        Public ReadOnly Property CapabilityId As String
            Get
                Return _capabilityId
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Mutable candidate emitted by interpretation.  GoalContractFreezer deep-copies it.
    ''' </summary>
    Public Class CandidateGoalCriterion
        Public Property Id As String = ""
        Public Property Statement As String = ""
        Public Property Kind As String = "semantic"
        Public Property SourceClauseIds As New List(Of String)()
        Public Property Required As Boolean = True
        Public Property VerificationCapability As String = ""
        Public Property CapabilityId As String = ""
    End Class

End Namespace
