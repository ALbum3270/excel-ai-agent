Imports System.Collections.Generic
Imports System.Linq

Namespace Agent

    ''' <summary>
    ''' Canonical outcome vocabulary for tool capabilities. This metadata describes what a
    ''' tool can prove; it never selects a tool or prescribes an execution sequence.
    ''' </summary>
    Public NotInheritable Class OutcomeEffectCatalog
        ' The generic object bridge may produce different semantic effects. Its host observer
        ' must emit one precise effectType; absent that, the multi-effect declaration fails
        ' closed instead of manufacturing a Cartesian product.
        Private Shared ReadOnly EffectsByTool As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase) From {
            {"ReadRange", {"read_coverage"}},
            {"PythonCompute", {"compute_artifact"}},
            {"WriteData", {"data_state"}},
            {"ApplyFormula", {"formula_state"}},
            {"CreateSheet", {"object_exists"}},
            {"CopySheet", {"object_exists"}},
            {"RenameSheet", {"property_state"}},
            {"DeleteSheet", {"object_absent"}},
            {"InsertRowCol", {"data_state"}},
            {"DeleteRowCol", {"data_state"}},
            {"MergeCells", {"property_state"}},
            {"HideRowCol", {"property_state"}},
            {"ProtectSheet", {"property_state"}},
            {"FormatRange", {"property_state"}},
            {"ConditionalFormat", {"property_state"}},
            {"AutoFit", {"property_state"}},
            {"SortData", {"order_state"}},
            {"FilterData", {"filter_state", "property_state"}},
            {"FindReplace", {"data_state"}},
            {"RemoveDuplicates", {"data_state"}},
            {"CleanData", {"data_state"}},
            {"TransformData", {"data_state"}},
            {"DataAnalysis", {"data_state"}},
            {"GenerateReport", {"artifact"}},
            {"CreateChart", {"artifact", "property_state"}},
            {"CreatePivotTable", {"artifact"}},
            {"OfficeObjectOperation", {"read_coverage", "property_state", "data_state", "formula_state", "order_state", "filter_state", "object_exists", "object_absent", "artifact", "unclassified_mutation"}}
        }

        Private Sub New()
        End Sub

        Public Shared Function GetEffects(tool As ToolDescriptor) As List(Of String)
            Dim result As New List(Of String)()
            If tool?.OutcomeEffects IsNot Nothing Then
                result.AddRange(tool.OutcomeEffects.Where(Function(value) Not String.IsNullOrWhiteSpace(value)))
            End If

            Dim configured As String() = Nothing
            If result.Count = 0 AndAlso tool IsNot Nothing AndAlso
               EffectsByTool.TryGetValue(If(tool.Id, ""), configured) Then
                result.AddRange(configured)
            End If

            If result.Count = 0 AndAlso tool IsNot Nothing Then
                Select Case If(tool.AccessMode, "write").Trim().ToLowerInvariant()
                    Case "read"
                        result.Add("read_coverage")
                    Case "compute"
                        result.Add("compute_artifact")
                    Case Else
                        ' Unknown mutating capabilities must not masquerade as a specific
                        ' outcome. A planner can only contract against declared effects.
                        result.Add("unclassified_mutation")
                End Select
            End If

            Return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        End Function
    End Class

End Namespace
