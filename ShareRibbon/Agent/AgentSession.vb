Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Agent 会话 - 统一的数据模型
    ''' </summary>
    Public Class AgentSession
        Public Property Id As String = Guid.NewGuid().ToString()
        Public Property UserRequest As String
        Public Property AppType As String
        Public Property CurrentContent As String
        Public Property Skill As AgentSkill
        Public Property SelectedSkill As SkillFileDefinition
        Public Property Spec As AgentTaskSpec

        ' ReAct 循环记录
        Public Property Iterations As New List(Of ReActIteration)
        Public Property Status As AgentStatus = AgentStatus.Idle
        Public Property StartTime As DateTime = DateTime.Now
        Public Property CurrentIteration As Integer = 0
        Public Property MaxIterations As Integer = 15

        ' 执行计划
        Public Property Plan As ExecutionPlan

        ' 最终结果
        Public Property Result As AgentResult

        Public Sub New(userRequest As String, appType As String, currentContent As String)
            Me.UserRequest = userRequest
            Me.AppType = appType
            Me.CurrentContent = currentContent
        End Sub
    End Class

    ''' <summary>
    ''' ReAct 迭代记录
    ''' </summary>
    Public Class ReActIteration
        Public Property Index As Integer
        ''' <summary>
        ''' Stable reference exposed to the model and used by the completion gate.  A model
        ''' must cite these references instead of presenting an ungrounded success claim.
        ''' </summary>
        Public Property EvidenceId As String
        Public Property Thought As String
        Public Property Action As ToolCall
        Public Property AccessMode As String
        ''' <summary>
        ''' Evidence records whose data was actually bound into this action by the runtime.
        ''' This is data provenance, not a predicted workflow dependency.
        ''' </summary>
        Public Property DependsOnEvidenceIds As New List(Of String)()
        Public Property Observation As String
        ''' <summary>
        ''' Canonical, untruncated host observation used by deterministic outcome verification.
        ''' Large tool data remains outside this ledger and is carried by AgentToolDataflow.
        ''' </summary>
        Public Property OutcomeEvidence As JObject
        Public Property OutcomeArtifacts As JToken
        Public Property ContractEvidence As New List(Of OutcomeEvidenceRecord)()
        Public Property Explanation As ExecutionExplanation
        Public Property Timestamp As DateTime = DateTime.Now
    End Class

    ''' <summary>
    ''' 执行计划
    ''' </summary>
    Public Class ExecutionPlan
        Public Property Understanding As String
        Public Property Steps As New List(Of PlanStep)
        Public Property Summary As String
        Public Property Complexity As String = "medium"
        Public Property CapabilityGap As String = ""
        ''' <summary>
        ''' Outcome assertions compiled from the user goal during the initial planning turn.
        ''' The loop never executes these as steps; it freezes them on AgentTaskSpec and uses
        ''' them only to match host evidence at completion.
        ''' </summary>
        Public Property OutcomeContract As OutcomeContract
    End Class

    ''' <summary>
    ''' 计划步骤
    ''' </summary>
    Public Class PlanStep
        Public Property StepNumber As Integer
        Public Property Description As String
        ''' <summary>
        ''' Optional capability hint for explaining a high-level step in the UI. It never
        ''' authorizes tools, gates progress, or determines task completion.
        ''' </summary>
        Public Property ToolHint As String
        ''' <summary>
        ''' Legacy executable-plan payload retained for persisted sessions and older providers.
        ''' Adaptive ReAct never executes this value directly.
        ''' </summary>
        Public Property Code As String
        Public Property Language As String = "json"
        Public Property Status As StepStatus = StepStatus.Pending
        Public Property ErrorMessage As String
        Public Property LastExplanation As ExecutionExplanation
    End Class

    Public Class ExecutionExplanation
        Public Property StepIndex As Integer
        Public Property StepDescription As String
        Public Property ToolId As String
        Public Property ToolName As String
        Public Property ToolCategory As String
        Public Property RiskLevel As String
        Public Property ParametersJson As String
        Public Property StartedAt As DateTime
        Public Property FinishedAt As DateTime
        Public Property ElapsedMs As Long
        Public Property BeforeSummary As String
        Public Property AfterSummary As String
        Public Property ObservationJson As String
        Public Property DataSummaryJson As String
        Public Property Success As Boolean
        Public Property Message As String
        Public Property SkillName As String
        Public Property ScriptFileName As String
        Public Property McpToolName As String
        Public Property McpStatus As String
        Public Property FailureReason As String
        Public Property UndoPointName As String
        Public Property UndoHint As String
        Public Property CanUndo As Boolean
        Public Property AutoRepairSummary As String
        Public Property FixAttempts As Integer
        Public Property ExplanationText As String
    End Class

    ''' <summary>
    ''' Agent 执行结果
    ''' </summary>
    Public Class AgentResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property SessionId As String
        Public Property IterationsCompleted As Integer
        Public Property FinalOutput As String
        ''' <summary>Stable error code for callers that must distinguish terminal causes.</summary>
        Public Property ErrorCode As String = ""
        ''' <summary>Whether the current user task must stop.</summary>
        Public Property TaskFatal As Boolean = False
        ''' <summary>Whether the caller must discard/reset the current Agent session/runtime.</summary>
        Public Property SessionFatal As Boolean = False

        Public Shared Function SuccessResult(sessionId As String,
                                              Optional message As String = "",
                                              Optional finalOutput As String = "") As AgentResult
            Return New AgentResult With {
                .Success = True,
                .SessionId = sessionId,
                .Message = message,
                .FinalOutput = finalOutput,
                .ErrorCode = "",
                .TaskFatal = False,
                .SessionFatal = False
            }
        End Function

        Public Shared Function Failed(sessionId As String,
                                      message As String,
                                      Optional taskFatal As Boolean = False,
                                      Optional sessionFatal As Boolean = False,
                                      Optional errorCode As String = "") As AgentResult
            Return New AgentResult With {
                .Success = False,
                .SessionId = sessionId,
                .Message = message,
                .ErrorCode = If(errorCode, ""),
                .TaskFatal = taskFatal OrElse sessionFatal,
                .SessionFatal = sessionFatal
            }
        End Function
    End Class

    ''' <summary>
    ''' Agent 状态枚举
    ''' </summary>
    Public Enum AgentStatus
        Idle
        Thinking
        Planning
        WaitingApproval
        Executing
        Observing
        Reflecting
        Completed
        Failed
        Aborted
    End Enum

    ''' <summary>
    ''' 步骤状态枚举
    ''' </summary>
    Public Enum StepStatus
        Pending
        Running
        Completed
        Failed
        Skipped
    End Enum

    ''' <summary>
    ''' 任务规格（Spec驱动）
    ''' </summary>
    Public Class AgentTaskSpec
        Private _rawUserRequest As String = ""
        Private _goalCompilation As Goals.GoalCompilationResult
        Private _goalContract As Goals.GoalContract
        Private _goalInterpretationFallbackReason As String = ""

        Public Property Goal As String = ""
        Public Property TargetObject As String = ""
        Public Property Constraints As New List(Of String)()
        Public Property SuccessCriteria As New List(Of String)()
        Public Property RequiredTools As New List(Of String)()
        ''' <summary>
        ''' User-visible capability constraints whose use is itself part of the requested
        ''' outcome (for example an explicit request to calculate with Python). Unlike
        ''' RequiredTools, these are policy constraints, not a prescribed workflow.
        ''' </summary>
        Public Property RequiredCapabilities As New List(Of String)()
        ''' <summary>
        ''' Legacy persisted name for RequiredCapabilities. New code must not use this list
        ''' to prescribe plan steps or execution order.
        ''' </summary>
        Public Property MandatoryTools As New List(Of String)()
        ''' <summary>
        ''' Legacy persisted workflow metadata. Adaptive ReAct intentionally ignores this
        ''' sequence; data dependencies are expressed by tool inputs and observations.
        ''' </summary>
        Public Property MandatoryToolSequence As New List(Of String)()
        Public Property RiskLevel As String = "safe"
        Public Property Complexity As String = "medium"
        ''' <summary>
        ''' Controls whether the task may mutate the active Office document.  The planner and
        ''' execution gate use this semantic contract instead of inferring safety from wording.
        ''' </summary>
        Public Property MutationPolicy As String = "allow"
        Public Property ExpectedOutputs As New List(Of String)()
        Public Property ExpectedSlideCount As Integer = 0
        ''' <summary>
        ''' Legacy host-evidence acceptance projection. It remains operational during migration,
        ''' but is non-authoritative for user-goal semantics and MUST NOT be converted back into
        ''' Goals.GoalContract.
        ''' </summary>
        Public Property OutcomeContract As OutcomeContract

        ''' <summary>
        ''' Exact user language captured before interpretation. It can be captured once and
        ''' cannot be silently replaced by a later plan or repair turn.
        ''' </summary>
        Public ReadOnly Property RawUserRequest As String
            Get
                Return _rawUserRequest
            End Get
        End Property

        ''' <summary>
        ''' Frozen authoritative goal. Planner and ReAct may read it but cannot replace it.
        ''' </summary>
        Public ReadOnly Property GoalContract As Goals.GoalContract
            Get
                Return _goalContract
            End Get
        End Property

        ''' <summary>
        ''' Structured but still untrusted interpretation captured at intake. It is internal so
        ''' planners cannot mutate or replace the semantic authority before Validate -> Freeze.
        ''' </summary>
        Friend ReadOnly Property GoalCompilation As Goals.GoalCompilationResult
            Get
                Return _goalCompilation
            End Get
        End Property

        ''' <summary>
        ''' Observable provenance when a structured interpretation was rejected and the exact
        ''' raw request was used instead.  This status is not part of user semantics.
        ''' </summary>
        Public ReadOnly Property GoalInterpretationFallbackReason As String
            Get
                Return _goalInterpretationFallbackReason
            End Get
        End Property

        Friend Sub RecordGoalInterpretationFallback(reason As String)
            Dim normalized = If(reason, "").Trim()
            If normalized.Length = 0 Then Return
            If _goalInterpretationFallbackReason.Length = 0 Then
                _goalInterpretationFallbackReason = normalized
                Return
            End If
            If Not String.Equals(_goalInterpretationFallbackReason, normalized, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("Goal interpretation fallback provenance has already been recorded.")
            End If
        End Sub

        Friend Sub CaptureRawUserRequest(value As String)
            Dim candidate = If(value, "")
            If String.IsNullOrWhiteSpace(candidate) Then
                Throw New ArgumentException("Raw user request cannot be empty.", NameOf(value))
            End If
            If String.IsNullOrEmpty(_rawUserRequest) Then
                _rawUserRequest = candidate
                Return
            End If
            If Not String.Equals(_rawUserRequest, candidate, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("RawUserRequest has already been captured and cannot be replaced.")
            End If
        End Sub

        Friend Sub SetGoalCompilationOnce(value As Goals.GoalCompilationResult)
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If String.IsNullOrEmpty(_rawUserRequest) Then
                Throw New InvalidOperationException("RawUserRequest must be captured before GoalCompilation is attached.")
            End If
            If value.Candidate Is Nothing OrElse
               Not String.Equals(_rawUserRequest, value.Candidate.RawUserRequest, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("GoalCompilation does not represent the captured RawUserRequest.")
            End If

            If _goalCompilation IsNot Nothing Then
                Dim existingFingerprint = Goals.GoalCoverageValidator.ComputeCompilationFingerprint(_goalCompilation)
                Dim incomingFingerprint = Goals.GoalCoverageValidator.ComputeCompilationFingerprint(value)
                If Not String.Equals(existingFingerprint, incomingFingerprint, StringComparison.Ordinal) Then
                    Throw New InvalidOperationException("GoalCompilation has already been attached and cannot be replaced.")
                End If
                Return
            End If
            _goalCompilation = value
        End Sub

        Friend Sub SetGoalContractOnce(value As Goals.GoalContract)
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If String.IsNullOrEmpty(_rawUserRequest) Then
                Throw New InvalidOperationException("RawUserRequest must be captured before GoalContract is attached.")
            End If
            If _goalContract IsNot Nothing Then
                If Not String.Equals(_goalContract.ContractHash, value.ContractHash, StringComparison.Ordinal) Then
                    Throw New InvalidOperationException("GoalContract is frozen and cannot be replaced.")
                End If
                Return
            End If
            If Not String.Equals(_rawUserRequest, value.RawUserRequest, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("GoalContract does not represent the captured RawUserRequest.")
            End If
            _goalContract = value
        End Sub

        Public ReadOnly Property IsSimple As Boolean
            Get
                Return String.Equals(Complexity, "simple", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Structured, execution-independent definition of the observable end state.  This is
    ''' deliberately separate from PlanStep and RequiredCapabilities: outcome, strategy and
    ''' policy must not collapse into the same concept.
    ''' </summary>
    Public Class OutcomeContract
        Private _boundGoalContractHash As String = ""
        Private _bindingMode As String = ""
        Private _frozenOutcomeContractHash As String = ""

        Public Property SchemaVersion As String = "1.0"
        Public Property Requirements As New List(Of OutcomeRequirement)()
        ''' <summary>
        ''' Hash of the authoritative GoalContract that this verification projection was
        ''' validated against.  The model cannot provide this value; the harness binds it
        ''' when the initial outcome contract is frozen.
        ''' </summary>
        <Newtonsoft.Json.JsonIgnore>
        Public ReadOnly Property BoundGoalContractHash As String
            Get
                Return _boundGoalContractHash
            End Get
        End Property
        ''' <summary>
        ''' goal-v1 for a contract projected from an immutable GoalContract; legacy-v1 only
        ''' for persisted sessions that have no GoalContract.
        ''' </summary>
        <Newtonsoft.Json.JsonIgnore>
        Public ReadOnly Property BindingMode As String
            Get
                Return _bindingMode
            End Get
        End Property
        ''' <summary>
        ''' Integrity seal over the normalized verification contract.  Frozen=True alone is
        ''' not trusted because OutcomeRequirement remains mutable during the migration.
        ''' </summary>
        <Newtonsoft.Json.JsonIgnore>
        Public ReadOnly Property FrozenOutcomeContractHash As String
            Get
                Return _frozenOutcomeContractHash
            End Get
        End Property
        ''' <summary>
        ''' Concrete workbook identity captured when active/short Excel references are frozen.
        ''' Evidence is bound at observation time, so switching ActiveWorkbook cannot
        ''' reinterpret old proof as belonging to another file.
        ''' </summary>
        Public Property BoundWorkbook As String = ""
        ''' <summary>
        ''' Compute capabilities resolved against the trusted ToolRegistry while the contract
        ''' is frozen. Ignored during JSON binding so a model cannot self-authorize a non-compute
        ''' producer by emitting a similarly named field.
        ''' </summary>
        <Newtonsoft.Json.JsonIgnore>
        Public Property ValidatedComputeCapabilities As New List(Of String)()
        Public Property Frozen As Boolean = False

        Friend Sub BindToGoal(goalContractHash As String,
                              bindingMode As String)
            _boundGoalContractHash = If(goalContractHash, "")
            _bindingMode = If(bindingMode, "")
        End Sub

        Friend Sub SealIntegrity(fingerprint As String)
            _frozenOutcomeContractHash = If(fingerprint, "")
        End Sub
    End Class

    Public Class OutcomeRequirement
        Public Property Id As String
        Public Property AppType As String
        Public Property TargetRef As String
        Public Property EffectType As String
        Public Property PropertyName As String
        Public Property [Operator] As String = "equals"
        Public Property ExpectedValue As JToken
        Public Property DerivedFromCapability As String
        ''' <summary>
        ''' Stable IDs of required GoalContract criteria that this host-observable requirement
        ''' helps prove. For persisted sessions without a GoalContract only, criterion-N refers
        ''' to the legacy AgentTaskSpec.SuccessCriteria projection.
        ''' </summary>
        Public Property CriterionIds As New List(Of String)()
        Public Property Required As Boolean = True
        Public Property Description As String
    End Class

    ''' <summary>
    ''' Canonical proof produced from a successful tool result.  Expected describes the
    ''' requested postcondition; Actual and Satisfied come from the host observation.
    ''' </summary>
    Public Class OutcomeEvidenceRecord
        Public Property EvidenceId As String
        Public Property IterationEvidenceId As String
        Public Property TargetRef As String
        Public Property EffectType As String
        Public Property PropertyName As String
        Public Property Expected As JToken
        Public Property Actual As JToken
        Public Property Satisfied As Boolean
        ''' <summary>
        ''' This record is a later host observation that invalidates overlapping older
        ''' evidence. It may accompany positive proof, or be invalidation-only after a
        ''' partially applied/failed write.
        ''' </summary>
        Public Property InvalidatesPrior As Boolean
        ''' <summary>
        ''' True only when a structured host verifier linked the tool request to the observed
        ''' postcondition. This permits user-facing parameter aliases without treating an
        ''' unverified request as evidence.
        ''' </summary>
        Public Property RequestVerified As Boolean
        ''' <summary>
        ''' Only the high-level request fields explicitly linked to passed host verification.
        ''' Never contains the complete tool request merely because some other property passed.
        ''' </summary>
        Public Property VerifiedRequest As JToken
        Public Property SourceToolId As String
        Public Property DataHash As String
        ''' <summary>
        ''' Monotonic revision of the observed Office world after this action. Later
        ''' overlapping state evidence supersedes older assertions.
        ''' </summary>
        Public Property WorldRevision As Long
        ''' <summary>
        ''' Canonical aliases for the same host object (for example a stable chart anchor and
        ''' Excel's generated ChartObject ref).  They participate only in invalidation, never
        ''' in satisfying a contract target.
        ''' </summary>
        Public Property RelatedTargetRefs As New List(Of String)()
        Public Property DerivedFromEvidenceIds As New List(Of String)()
    End Class

End Namespace
