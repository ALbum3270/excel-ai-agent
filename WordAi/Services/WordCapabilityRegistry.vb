' WordAi\Services\WordCapabilityRegistry.vb
' Word capability metadata used by WordActionHarness planning.

Namespace Services

    Public Enum WordCapabilityRiskLevel
        Low
        Medium
        High
    End Enum

    Public Class WordCapabilityDescriptor
        Public Property Id As String = ""
        Public Property DisplayName As String = ""
        Public Property Kind As WordActionKind = WordActionKind.None
        Public Property Description As String = ""
        Public Property InputSchema As String = ""
        Public Property RiskLevel As WordCapabilityRiskLevel = WordCapabilityRiskLevel.Medium
        Public Property SupportsPreview As Boolean
        Public Property SupportsUndo As Boolean = True
        Public Property ObserveContract As String = ""
        Public Property RepairContract As String = ""
        Public Property ExplainContract As String = ""
        Public Property ExampleRequests As New List(Of String)()

        Public Function ToHumanReadableSummary() As String
            Dim parts As New List(Of String) From {
                $"{DisplayName} ({Id})",
                $"风险: {RiskLevel}",
                $"预览: {If(SupportsPreview, "支持", "不支持")}",
                $"撤销: {If(SupportsUndo, "进入 Word 撤销栈", "需执行器自行说明")}"
            }

            If Not String.IsNullOrWhiteSpace(ObserveContract) Then parts.Add("观察: " & ObserveContract)
            If Not String.IsNullOrWhiteSpace(RepairContract) Then parts.Add("修复: " & RepairContract)
            Return String.Join("；", parts)
        End Function
    End Class

    Public NotInheritable Class WordCapabilityRegistry

        Private Shared ReadOnly _capabilities As Lazy(Of List(Of WordCapabilityDescriptor)) =
            New Lazy(Of List(Of WordCapabilityDescriptor))(AddressOf BuildCapabilities)

        Private Sub New()
        End Sub

        Public Shared Function All() As IReadOnlyList(Of WordCapabilityDescriptor)
            Return _capabilities.Value
        End Function

        Public Shared Function Find(kind As WordActionKind) As WordCapabilityDescriptor
            Return _capabilities.Value.FirstOrDefault(Function(item) item.Kind = kind)
        End Function

        Public Shared Function Require(kind As WordActionKind) As WordCapabilityDescriptor
            Dim descriptor = Find(kind)
            If descriptor Is Nothing Then
                Return New WordCapabilityDescriptor With {
                    .Id = "word.unknown",
                    .DisplayName = "未知 Word 能力",
                    .Kind = kind,
                    .Description = "未登记的 Word capability",
                    .RiskLevel = WordCapabilityRiskLevel.High,
                    .SupportsPreview = False,
                    .SupportsUndo = False
                }
            End If
            Return descriptor
        End Function

        Private Shared Function BuildCapabilities() As List(Of WordCapabilityDescriptor)
            Return New List(Of WordCapabilityDescriptor) From {
                New WordCapabilityDescriptor With {
                    .Id = "word.proofread",
                    .DisplayName = "Word 校对",
                    .Kind = WordActionKind.Proofread,
                    .Description = "读取当前选区或全文，生成错别字、标点、语法、术语和风格一致性检查计划。",
                    .InputSchema = "ProofreadIntentPlan(scope, issueTypes, applyMode, showSidePanel)",
                    .RiskLevel = WordCapabilityRiskLevel.Low,
                    .SupportsPreview = True,
                    .SupportsUndo = True,
                    .ObserveContract = "校对结果应进入侧边栏或高置信修正流程，并保留用户确认入口。",
                    .RepairContract = "若无法读取文档或选区，回退到全文校对计划或提示当前没有可处理文档。",
                    .ExplainContract = "说明校对范围、问题类型、是否自动应用高置信修正。",
                    .ExampleRequests = New List(Of String) From {"校对全文", "检查选中段落的错别字", "审校这篇文档的标点和术语"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.direct-formatting",
                    .DisplayName = "Word 直接格式调整",
                    .Kind = WordActionKind.DirectFormatting,
                    .Description = "把明确的字号、字体、加粗、颜色、对齐等请求编译为 FormattingIntentPlan 并应用到 Word 文档。",
                    .InputSchema = "FormattingIntentPlan(scope, operations, confidence)",
                    .RiskLevel = WordCapabilityRiskLevel.Medium,
                    .SupportsPreview = False,
                    .SupportsUndo = True,
                    .ObserveContract = "执行后抽样读取目标 Range，验证字号、字体、对齐等操作已生效。",
                    .RepairContract = "观察失败时返回失败摘要，由上层回退到普通聊天或后续 repair loop。",
                    .ExplainContract = "说明作用范围、应用的格式操作、观察结果和撤销方式。",
                    .ExampleRequests = New List(Of String) From {"字体统一加大2号", "把选中文字改成宋体并加粗", "正文行距调整为1.5倍"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.numbering",
                    .DisplayName = "Word 自动编号重排",
                    .Kind = WordActionKind.Numbering,
                    .Description = "识别 Word 自动编号段落，并重排为连续递增的 1,2,3... 编号。",
                    .InputSchema = "NumberingRequest(scope, sequenceGoal)",
                    .RiskLevel = WordCapabilityRiskLevel.Medium,
                    .SupportsPreview = False,
                    .SupportsUndo = True,
                    .ObserveContract = "执行后读取前几个编号段落，展示 ListString 和文本预览。",
                    .RepairContract = "若没有自动编号段落，不转换普通文本编号，提示用户当前范围不可处理。",
                    .ExplainContract = "说明处理范围、检测数量、应用数量、观察预览和撤销方式。",
                    .ExampleRequests = New List(Of String) From {"把前面的序号改为12345", "将列表编号重排为连续递增", "整理全文自动编号"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.semantic-reformat",
                    .DisplayName = "Word 语义智能排版",
                    .Kind = WordActionKind.SemanticReformat,
                    .Description = "面向标题、层级、目录、结构化段落的智能排版计划，优先走预览确认。",
                    .InputSchema = "WordFormattingTaskPlan(scopeSummary, standardName, targetSummary, operations)",
                    .RiskLevel = WordCapabilityRiskLevel.High,
                    .SupportsPreview = True,
                    .SupportsUndo = True,
                    .ObserveContract = "应用后统计预期段落、已修改段落和修复段落，并说明未命中的原因。",
                    .RepairContract = "结构段落未命中时根据 observe 结果修正定位或保留为待确认项。",
                    .ExplainContract = "说明排版标准、预览数量、应用数量、修复数量和撤销方式。",
                    .ExampleRequests = New List(Of String) From {"整理标题层级", "按公文格式优化全文排版", "规范目录和编号结构"}
                }
            }
        End Function

    End Class

End Namespace
