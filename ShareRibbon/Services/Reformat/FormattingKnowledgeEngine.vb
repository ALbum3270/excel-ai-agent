' ShareRibbon\Services\Reformat\FormattingKnowledgeEngine.vb
' 排版知识引擎 - 管理内置和用户自定义的排版标准

' 注意：DocumentType 枚举在 DocumentAnalyzer.vb 中定义，此文件引用之
' 注意：DocumentAnalysisResult 也在 DocumentAnalyzer.vb 中定义

''' <summary>
''' 排版标准 - 定义一种文档类型的完整格式规范
''' </summary>
Public Class FormattingStandard
    ''' <summary>标准唯一标识</summary>
    Public Property Id As String = Guid.NewGuid().ToString()

    ''' <summary>标准名称（如 "GB/T 9704-2012"）</summary>
    Public Property Name As String = ""

    ''' <summary>标准描述</summary>
    Public Property Description As String = ""

    ''' <summary>适用的文档类型列表（DocumentType枚举值的名称）</summary>
    Public Property ApplicableDocumentTypes As List(Of String)

    ''' <summary>语义排版映射（核心格式数据）</summary>
    Public Property SemanticMapping As SemanticStyleMapping

    ''' <summary>是否为内置标准</summary>
    Public Property IsBuiltIn As Boolean = False

    ''' <summary>是否激活</summary>
    Public Property IsActive As Boolean = True

    Public Sub New()
        ApplicableDocumentTypes = New List(Of String)()
        SemanticMapping = New SemanticStyleMapping()
    End Sub

    Public Sub New(name As String, description As String)
        Me.Name = name
        Me.Description = description
        ApplicableDocumentTypes = New List(Of String)()
        SemanticMapping = New SemanticStyleMapping()
    End Sub

    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(Name), "(未命名)", Name)
    End Function
End Class

''' <summary>
''' 排版知识引擎 - 管理内置和用户自定义的排版标准
''' 提供内置标准检索。
''' </summary>
Public Class FormattingKnowledgeEngine
    Private ReadOnly _standards As New List(Of FormattingStandard)()

    Public Sub New()
        ' 从内置数据加载所有标准
        Dim builtInStandards = FormattingStandardData.GetAllBuiltInStandards()
        _standards.AddRange(builtInStandards)

    End Sub

    ''' <summary>
    ''' 获取适用于指定文档类型的标准
    ''' </summary>
    Public Function GetStandardForDocumentType(docType As DocumentType) As FormattingStandard
        Dim typeName = docType.ToString()
        Return _standards.FirstOrDefault(Function(s) s.ApplicableDocumentTypes.Contains(typeName) AndAlso s.IsActive)
    End Function

    ''' <summary>
    ''' 根据名称获取标准
    ''' </summary>
    Public Function GetStandardByName(name As String) As FormattingStandard
        If String.IsNullOrEmpty(name) Then Return Nothing
        Return _standards.FirstOrDefault(Function(s) s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
    End Function

    ''' <summary>
    ''' 获取所有已激活的标准
    ''' </summary>
    Public Function GetActiveStandards() As List(Of FormattingStandard)
        Return _standards.Where(Function(s) s.IsActive).ToList()
    End Function

    ''' <summary>
    ''' 获取适用于指定文档类型的标准（支持字符串类型名称）
    ''' </summary>
    Public Function GetStandardForDocumentType(docTypeName As String) As FormattingStandard
        If String.IsNullOrEmpty(docTypeName) Then Return Nothing
        Return _standards.FirstOrDefault(Function(s) s.ApplicableDocumentTypes.Contains(docTypeName) AndAlso s.IsActive)
    End Function
End Class
