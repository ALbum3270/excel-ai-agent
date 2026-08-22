Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports Excel = Microsoft.Office.Interop.Excel
Imports ShareRibbon
Imports ShareRibbon.Agent
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    ''' <summary>
    ''' Read-only catalog for the bounded Excel object model exposed to the Agent.
    ''' Discovery and execution share this exact catalog; a member that is absent here
    ''' cannot be invoked by ExcelOperationExecutor.
    ''' </summary>
    Public NotInheritable Class ExcelApiCatalogProvider
        Private Shared ReadOnly CatalogCache As New Lazy(Of CatalogSnapshot)(
            AddressOf BuildCatalog,
            LazyThreadSafetyMode.ExecutionAndPublication)

        Private Shared ReadOnly BlockedMemberNames As New HashSet(Of String)(
            {
                "quit", "close", "saveas", "savecopyas", "run", "executeexcel4macro",
                "evaluate", "_evaluate", "followhyperlink", "sendmail", "sendfaxoverinternet",
                "printout", "printpreview", "publishobjects", "vbproject"
            },
            StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly BlockedMemberTokens As String() = {
            "save", "export", "print", "mail", "fax", "hyperlink", "publish",
            "execute", "run", "quit", "close", "vbproject"
        }

        Private Shared ReadOnly DestructiveTokens As String() = {"delete", "remove", "clear", "replace"}

        Private Sub New()
        End Sub

        Public Shared Function SearchAsToolResult(params As JObject) As ToolResult
            Const toolId As String = "DiscoverOfficeCapability"
            Try
                Dim request = If(params, New JObject()).ToObject(Of OfficeCapabilitySearchRequest)()
                Dim validation = OfficeOperationValidation.ValidateCapabilitySearch(request)
                If Not validation.IsValid Then
                    Return ToolResult.Failed(toolId,
                                             validation.ToErrorMessage(),
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="Excel 能力查询参数无效",
                                             recoverable:=True)
                End If

                Dim result = Search(request)
                Dim observation = New JObject From {
                    {"kind", "office_capability_search"},
                    {"summary", $"找到 {result.Members.Count} 个相关 Excel 对象成员"},
                    {"changed", False},
                    {"readOnly", True},
                    {"resultCount", result.Members.Count},
                    {"truncated", result.Truncated},
                    {"catalogFingerprint", result.CatalogFingerprint},
                    {"warnings", JArray.FromObject(result.Warnings)}
                }
                If result.Members.Count = 0 Then
                    Return ToolResult.Failed(toolId,
                                             "No matching Excel capability was found",
                                             data:=result,
                                             errorCode:=ExceptionClassifier.CodeCapabilityNotFound,
                                             userMessage:="未找到与目标匹配且允许执行的 Excel 对象能力",
                                             recoverable:=True,
                                             observation:=observation)
                End If
                Return ToolResult.Succeed(toolId,
                                          observation("summary").ToString(),
                                          data:=result,
                                          observation:=observation)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            End Try
        End Function

        Public Shared Function Search(request As OfficeCapabilitySearchRequest) As OfficeCapabilitySearchResult
            Dim validation = OfficeOperationValidation.ValidateCapabilitySearch(request)
            If Not validation.IsValid Then Throw New ArgumentException(validation.ToErrorMessage())

            Dim normalizedQuery = NormalizeSearchText(request.Query)
            Dim normalizedTargetType = NormalizeSearchText(request.TargetType)
            Dim scored As New List(Of ScoredCatalogEntry)()
            For Each entry In CatalogCache.Value.Entries
                If Not request.IncludeReadOnly AndAlso entry.IsReadOnly Then Continue For
                If Not String.IsNullOrWhiteSpace(normalizedTargetType) AndAlso
                   entry.NormalizedTypeName.IndexOf(normalizedTargetType, StringComparison.Ordinal) < 0 Then
                    Continue For
                End If
                Dim score = ScoreEntry(entry, normalizedQuery, request.Query)
                If score > 0 Then scored.Add(New ScoredCatalogEntry With {.Entry = entry, .Score = score})
            Next

            Dim ordered = scored.OrderByDescending(Function(item) item.Score).
                                 ThenBy(Function(item) item.Entry.Member.MemberId, StringComparer.Ordinal).
                                 ToList()
            Dim selected = ordered.Take(request.MaxResults).Select(Function(item) item.Entry.Member).ToList()
            Return New OfficeCapabilitySearchResult With {
                .AppType = "Excel",
                .Query = request.Query,
                .Members = selected,
                .CatalogFingerprint = CatalogCache.Value.Fingerprint,
                .Truncated = ordered.Count > selected.Count,
                .Warnings = New List(Of String)()
            }
        End Function

        Friend Shared Function TryGetMemberBinding(memberId As String,
                                                   ByRef member As OfficeCapabilityMember,
                                                   ByRef reflectedMember As MemberInfo) As Boolean
            member = Nothing
            reflectedMember = Nothing
            Dim entry = CatalogCache.Value.Entries.FirstOrDefault(
                Function(item) String.Equals(item.Member.MemberId, memberId, StringComparison.Ordinal))
            If entry Is Nothing Then Return False
            member = entry.Member
            reflectedMember = entry.ReflectedMember
            Return reflectedMember IsNot Nothing
        End Function

        Friend Shared Function FindMemberId(typeName As String,
                                            memberName As String,
                                            memberKind As String) As String
            Dim requestedType = NormalizeLogicalTypeName(typeName)
            Dim entry = CatalogCache.Value.Entries.FirstOrDefault(
                Function(item)
                    Return String.Equals(item.LogicalTypeName, requestedType, StringComparison.OrdinalIgnoreCase) AndAlso
                           String.Equals(item.Member.MemberName, memberName, StringComparison.OrdinalIgnoreCase) AndAlso
                           String.Equals(item.Member.MemberKind, memberKind, StringComparison.OrdinalIgnoreCase)
                End Function)
            Return If(entry?.Member?.MemberId, "")
        End Function

        Private Shared Function BuildCatalog() As CatalogSnapshot
            Dim entries As New Dictionary(Of String, CatalogEntry)(StringComparer.Ordinal)
            For Each apiType In GetCatalogTypes()
                AddProperties(apiType, entries)
                AddMethods(apiType, entries)
            Next
            Dim ordered = entries.Values.OrderBy(Function(item) item.Member.MemberId, StringComparer.Ordinal).ToList()
            Dim fingerprintSource = String.Join(vbLf, ordered.Select(Function(item) item.Member.MemberId))
            Return New CatalogSnapshot With {
                .Entries = ordered,
                .Fingerprint = ComputeHash(fingerprintSource)
            }
        End Function

        Private Shared Function GetCatalogTypes() As IEnumerable(Of Type)
            Dim interopAssembly = ResolveInteropAssembly()
            Return New Type() {
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel._Workbook", GetType(Excel.Workbook)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Sheets", GetType(Excel.Worksheets)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel._Worksheet", GetType(Excel.Worksheet)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Range", GetType(Excel.Range)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Font", GetType(Excel.Font)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Interior", GetType(Excel.Interior)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Borders", GetType(Excel.Borders)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Border", GetType(Excel.Border)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.ChartObjects", GetType(Excel.ChartObjects)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.ChartObject", GetType(Excel.ChartObject)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel._Chart", GetType(Excel.Chart)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.ChartTitle", GetType(Excel.ChartTitle)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Legend", GetType(Excel.Legend)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.SeriesCollection", GetType(Excel.SeriesCollection)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.Series", GetType(Excel.Series)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.ListObjects", GetType(Excel.ListObjects)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.ListObject", GetType(Excel.ListObject)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.PivotTables", GetType(Excel.PivotTables)),
                ResolveCatalogType(interopAssembly, "Microsoft.Office.Interop.Excel.PivotTable", GetType(Excel.PivotTable))
            }.Where(Function(item) item IsNot Nothing).
              GroupBy(Function(item) item.FullName, StringComparer.Ordinal).
              Select(Function(group) group.First()).ToList()
        End Function

        Private Shared Function ResolveCatalogType(interopAssembly As Assembly,
                                                   fullTypeName As String,
                                                   embeddedFallback As Type) As Type
            If interopAssembly IsNot Nothing Then
                Try
                    Dim resolved = interopAssembly.GetType(fullTypeName, throwOnError:=False, ignoreCase:=False)
                    If resolved IsNot Nothing Then Return resolved
                Catch
                End Try
            End If
            Return embeddedFallback
        End Function

        Private Shared Function ResolveInteropAssembly() As Assembly
            For Each loaded In AppDomain.CurrentDomain.GetAssemblies()
                Try
                    If String.Equals(loaded.GetName().Name, "Microsoft.Office.Interop.Excel", StringComparison.OrdinalIgnoreCase) AndAlso
                       loaded.GetTypes().Any(Function(item) String.Equals(item.FullName, "Microsoft.Office.Interop.Excel.Range", StringComparison.Ordinal)) Then
                        Return loaded
                    End If
                Catch
                End Try
            Next
            Try
                Return Assembly.Load("Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
            Catch
            End Try

            Dim windowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
            Dim candidates As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(windowsDir) Then
                candidates.Add(Path.Combine(windowsDir,
                                            "assembly",
                                            "GAC_MSIL",
                                            "Microsoft.Office.Interop.Excel",
                                            "15.0.0.0__71e9bce111e9429c",
                                            "Microsoft.Office.Interop.Excel.dll"))
                Dim gacRoot = Path.Combine(windowsDir, "assembly", "GAC_MSIL", "Microsoft.Office.Interop.Excel")
                If Directory.Exists(gacRoot) Then
                    Try
                        candidates.AddRange(Directory.GetFiles(gacRoot,
                                                               "Microsoft.Office.Interop.Excel.dll",
                                                               SearchOption.AllDirectories))
                    Catch
                    End Try
                End If
            End If
            For Each candidate In candidates.Distinct(StringComparer.OrdinalIgnoreCase)
                Try
                    If File.Exists(candidate) Then Return Assembly.LoadFrom(candidate)
                Catch ex As Exception
                    Debug.WriteLine($"[ExcelApiCatalog] Unable to load interop assembly {candidate}: {ex.Message}")
                End Try
            Next
            Return Nothing
        End Function

        Private Shared Sub AddProperties(apiType As Type, entries As Dictionary(Of String, CatalogEntry))
            For Each prop In apiType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                Try
                    If Not prop.CanRead AndAlso Not prop.CanWrite Then Continue For
                    Dim parameters = prop.GetIndexParameters().Select(AddressOf BuildParameter).ToList()
                    Dim isReadOnly = prop.CanRead AndAlso Not prop.CanWrite
                    AddEntry(entries,
                             BuildMember(apiType, prop.Name, "property", parameters, FriendlyTypeName(prop.PropertyType), isReadOnly),
                             prop,
                             isReadOnly)
                Catch ex As Exception
                    Debug.WriteLine($"[ExcelApiCatalog] Skip property {apiType.Name}.{prop.Name}: {ex.Message}")
                End Try
            Next
        End Sub

        Private Shared Sub AddMethods(apiType As Type, entries As Dictionary(Of String, CatalogEntry))
            For Each method In apiType.GetMethods(BindingFlags.Public Or BindingFlags.Instance)
                If method.IsSpecialName Then Continue For
                Try
                    Dim parameters = method.GetParameters().Select(AddressOf BuildParameter).ToList()
                    Dim isReadOnly = IsReadOnlyMethod(method.Name)
                    AddEntry(entries,
                             BuildMember(apiType, method.Name, "method", parameters, FriendlyTypeName(method.ReturnType), isReadOnly),
                             method,
                             isReadOnly)
                Catch ex As Exception
                    Debug.WriteLine($"[ExcelApiCatalog] Skip method {apiType.Name}.{method.Name}: {ex.Message}")
                End Try
            Next
        End Sub

        Private Shared Function BuildMember(apiType As Type,
                                            memberName As String,
                                            memberKind As String,
                                            parameters As List(Of OfficeCapabilityParameter),
                                            returnType As String,
                                            isReadOnly As Boolean) As OfficeCapabilityMember
            Dim logicalTypeName = NormalizeLogicalTypeName(apiType.Name)
            Dim parameterSignature = String.Join(",", parameters.Select(Function(item) item.ParameterType))
            Dim memberId = $"Excel.{logicalTypeName}.{memberKind}.{memberName}({parameterSignature})->{returnType}"
            Dim executable = Not IsBlockedMember(memberName) AndAlso Not HasExternalIoParameter(parameters)
            Return New OfficeCapabilityMember With {
                .MemberId = memberId,
                .DeclaringType = If(apiType.FullName, logicalTypeName).Replace("._", "."),
                .MemberName = memberName,
                .MemberKind = memberKind,
                .Parameters = parameters,
                .ReturnType = returnType,
                .RiskLevel = ClassifyRisk(memberName, isReadOnly),
                .Executable = executable,
                .UnsupportedReason = If(executable, "", "Lifecycle, macro, external execution, file/network input, and document output members are disabled"),
                .Aliases = BuildAliases(logicalTypeName, memberName)
            }
        End Function

        Private Shared Sub AddEntry(entries As Dictionary(Of String, CatalogEntry),
                                    member As OfficeCapabilityMember,
                                    reflectedMember As MemberInfo,
                                    isReadOnly As Boolean)
            If member Is Nothing OrElse entries.ContainsKey(member.MemberId) Then Return
            Dim logicalTypeName = NormalizeLogicalTypeName(reflectedMember.DeclaringType.Name)
            Dim terms As New List(Of String) From {
                logicalTypeName, member.DeclaringType, member.MemberName, member.MemberKind, member.ReturnType
            }
            terms.AddRange(member.Aliases)
            terms.AddRange(member.Parameters.Select(Function(item) item.Name & " " & item.ParameterType))
            entries(member.MemberId) = New CatalogEntry With {
                .Member = member,
                .ReflectedMember = reflectedMember,
                .IsReadOnly = isReadOnly,
                .LogicalTypeName = logicalTypeName,
                .NormalizedTypeName = NormalizeSearchText(logicalTypeName),
                .NormalizedSearchText = NormalizeSearchText(String.Join(" ", terms)),
                .NormalizedTerms = terms.Select(AddressOf NormalizeSearchText).
                                         Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                                         Distinct().ToList()
            }
        End Sub

        Private Shared Function BuildParameter(parameter As ParameterInfo) As OfficeCapabilityParameter
            Return New OfficeCapabilityParameter With {
                .Name = parameter.Name,
                .ParameterType = FriendlyTypeName(parameter.ParameterType),
                .Required = Not parameter.IsOptional,
                .DefaultValue = GetSafeDefaultValue(parameter),
                .Description = If(parameter.IsOptional, "optional", "required")
            }
        End Function

        Private Shared Function GetSafeDefaultValue(parameter As ParameterInfo) As Object
            If parameter Is Nothing OrElse Not parameter.IsOptional Then Return Nothing
            Try
                Dim value = parameter.DefaultValue
                If value Is Nothing OrElse value Is DBNull.Value OrElse value Is Missing.Value Then Return Nothing
                If value.GetType().IsEnum Then Return value.ToString()
                If TypeOf value Is String OrElse TypeOf value Is Boolean OrElse TypeOf value Is Decimal OrElse
                   TypeOf value Is Double OrElse TypeOf value Is Single OrElse TypeOf value Is Integer OrElse
                   TypeOf value Is Long OrElse TypeOf value Is Short OrElse TypeOf value Is Byte Then Return value
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function FriendlyTypeName(value As Type) As String
            If value Is Nothing Then Return "Object"
            If value.IsByRef Then value = value.GetElementType()
            If value Is GetType(Void) Then Return "Void"
            If value.IsArray Then Return FriendlyTypeName(value.GetElementType()) & "[]"
            Return If(String.IsNullOrWhiteSpace(value.FullName), value.Name, value.FullName)
        End Function

        Private Shared Function NormalizeLogicalTypeName(value As String) As String
            Dim normalized = If(value, "").Trim().TrimStart("_"c)
            Select Case normalized.ToLowerInvariant()
                Case "sheets", "worksheets"
                    Return "Worksheets"
                Case "workbook"
                    Return "Workbook"
                Case "worksheet"
                    Return "Worksheet"
                Case "chart"
                    Return "Chart"
                Case Else
                    Return normalized
            End Select
        End Function

        Private Shared Function BuildAliases(typeName As String, memberName As String) As List(Of String)
            Dim aliases As New List(Of String)()
            Select Case typeName.TrimStart("_"c).ToLowerInvariant()
                Case "workbook" : aliases.AddRange({"工作簿", "文件"})
                Case "worksheets" : aliases.AddRange({"工作表集合", "全部工作表"})
                Case "worksheet" : aliases.AddRange({"工作表", "sheet", "表"})
                Case "range" : aliases.AddRange({"单元格", "区域", "范围", "选区"})
                Case "font" : aliases.AddRange({"字体", "文字格式"})
                Case "interior" : aliases.AddRange({"填充", "背景", "底色"})
                Case "borders", "border" : aliases.AddRange({"边框", "框线"})
                Case "chartobjects", "chartobject", "chart" : aliases.AddRange({"图表", "图表对象"})
                Case "charttitle" : aliases.AddRange({"图表标题", "标题"})
                Case "legend" : aliases.AddRange({"图例", "图例位置"})
                Case "seriescollection", "series" : aliases.AddRange({"数据系列", "系列", "分类轴"})
                Case "listobjects", "listobject" : aliases.AddRange({"Excel表", "结构化表格"})
                Case "pivottables", "pivottable" : aliases.AddRange({"数据透视表", "透视表"})
            End Select

            Select Case memberName.ToLowerInvariant()
                Case "numberformat", "numberformatlocal" : aliases.AddRange({"数字格式", "千位分隔符", "小数位", "百分比", "日期格式", "货币格式"})
                Case "bold" : aliases.AddRange({"加粗", "粗体"})
                Case "italic" : aliases.AddRange({"斜体"})
                Case "color", "colorindex" : aliases.AddRange({"颜色", "字体颜色", "背景色", "填充色"})
                Case "horizontalalignment" : aliases.AddRange({"水平对齐", "左对齐", "居中", "右对齐"})
                Case "verticalalignment" : aliases.AddRange({"垂直对齐", "顶部对齐", "垂直居中", "底部对齐"})
                Case "wraptext" : aliases.AddRange({"自动换行", "换行"})
                Case "add" : aliases.AddRange({"新增", "添加", "创建", "插入"})
                Case "delete" : aliases.Add("删除")
                Case "name" : aliases.AddRange({"名称", "重命名", "标题"})
                Case "value", "value2" : aliases.AddRange({"值", "数据", "写入"})
                Case "formula", "formula2" : aliases.AddRange({"公式", "函数"})
                Case "autofit" : aliases.AddRange({"自动调整列宽", "自动调整行高"})
                Case "sort" : aliases.AddRange({"排序", "升序", "降序"})
                Case "autofilter" : aliases.AddRange({"筛选", "过滤"})
            End Select
            Return aliases
        End Function

        Private Shared Function ClassifyRisk(memberName As String, isReadOnly As Boolean) As String
            If IsBlockedMember(memberName) Then Return "risky"
            Dim normalized = If(memberName, "").ToLowerInvariant()
            If DestructiveTokens.Any(Function(token) normalized.Contains(token)) Then Return "risky"
            If isReadOnly Then Return "safe"
            Return "medium"
        End Function

        Private Shared Function IsReadOnlyMethod(memberName As String) As Boolean
            Dim normalized = If(memberName, "").ToLowerInvariant()
            Return normalized.StartsWith("get") OrElse normalized.StartsWith("find") OrElse
                   normalized.StartsWith("item") OrElse normalized.StartsWith("can") OrElse
                   normalized.StartsWith("is") OrElse normalized.StartsWith("count")
        End Function

        Private Shared Function IsBlockedMember(memberName As String) As Boolean
            If BlockedMemberNames.Contains(memberName) Then Return True
            Dim normalized = If(memberName, "").Trim().ToLowerInvariant()
            Return BlockedMemberTokens.Any(Function(token) normalized.Contains(token))
        End Function

        Private Shared Function HasExternalIoParameter(parameters As IEnumerable(Of OfficeCapabilityParameter)) As Boolean
            If parameters Is Nothing Then Return False
            Dim blockedNames As String() = {
                "filename", "filepath", "path", "url", "uri", "customdictionary", "dictionaryfile"
            }
            Return parameters.Any(
                Function(parameter)
                    Dim normalized = If(parameter?.Name, "").Replace("_", "").ToLowerInvariant()
                    Return blockedNames.Any(Function(blocked) normalized.Contains(blocked))
                End Function)
        End Function

        Private Shared Function ScoreEntry(entry As CatalogEntry,
                                           normalizedQuery As String,
                                           rawQuery As String) As Integer
            If entry Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedQuery) Then Return 0
            Dim score As Integer = 0
            Dim memberName = NormalizeSearchText(entry.Member.MemberName)
            If memberName = normalizedQuery Then score += 120
            If entry.NormalizedSearchText.IndexOf(normalizedQuery, StringComparison.Ordinal) >= 0 Then score += 60
            For Each term In entry.NormalizedTerms
                If term.Length < 2 Then Continue For
                If normalizedQuery.IndexOf(term, StringComparison.Ordinal) >= 0 Then
                    score += Math.Min(45, 12 + term.Length)
                ElseIf term.IndexOf(normalizedQuery, StringComparison.Ordinal) >= 0 Then
                    score += 15
                End If
            Next
            Dim lowerQuery = If(rawQuery, "").ToLowerInvariant()
            Dim lowerMember = If(entry.Member.MemberName, "").ToLowerInvariant()
            If ContainsAny(lowerQuery, {"create", "add", "insert", "创建", "新增", "添加", "插入"}) AndAlso
               ContainsAny(lowerMember, {"create", "add", "insert"}) Then score += 30
            If ContainsAny(lowerQuery, {"delete", "remove", "删除"}) AndAlso
               ContainsAny(lowerMember, {"delete", "remove"}) Then score += 30
            If entry.Member.Executable Then score += 2
            Return score
        End Function

        Private Shared Function ContainsAny(value As String, tokens As IEnumerable(Of String)) As Boolean
            Return tokens.Any(Function(token) If(value, "").IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

        Private Shared Function NormalizeSearchText(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim builder As New StringBuilder()
            For Each ch In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) Then builder.Append(ch)
            Next
            Return builder.ToString()
        End Function

        Private Shared Function ComputeHash(value As String) As String
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Class CatalogSnapshot
            Public Property Entries As List(Of CatalogEntry)
            Public Property Fingerprint As String
        End Class

        Private Class CatalogEntry
            Public Property Member As OfficeCapabilityMember
            Public Property ReflectedMember As MemberInfo
            Public Property IsReadOnly As Boolean
            Public Property LogicalTypeName As String
            Public Property NormalizedTypeName As String
            Public Property NormalizedSearchText As String
            Public Property NormalizedTerms As List(Of String)
        End Class

        Private Class ScoredCatalogEntry
            Public Property Entry As CatalogEntry
            Public Property Score As Integer
        End Class
    End Class

End Namespace
