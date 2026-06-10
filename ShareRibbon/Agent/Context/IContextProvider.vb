Namespace Agent.Context

    ''' <summary>
    ''' 上下文提供器接口 - 统一 Excel/Word/PowerPoint 的上下文采集
    ''' </summary>
    Public Interface IContextProvider

        ''' <summary>
        ''' 获取当前 Office 应用的上下文
        ''' </summary>
        ''' <returns>包含选区、结构等信息的上下文对象</returns>
        Function GetContext() As OfficeContext

    End Interface

End Namespace
