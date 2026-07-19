Imports System.Diagnostics

''' <summary>
''' VSTO 启动关键路径管理器：同步注册程序集解析和宿主状态栏。
''' WebView2 与 SQLite 继续由各宿主的 Lazy 按需初始化。
''' </summary>
Public Class PhaseStartupManager

    ''' <summary>
    ''' 全局单例 — 供 ThisAddIn 和 BaseOfficeRibbon 共享
    ''' </summary>
    Public Shared ReadOnly Instance As New PhaseStartupManager()

    ''' <summary>
    ''' Phase 0: 关键路径初始化（同步，主线程，必须在 VSTO Startup 事件中调用）
    ''' 仅注册事件处理器，不做任何 I/O 密集操作
    ''' </summary>
    Public Sub RunCriticalPhase(application As Object)
        Dim sw = Stopwatch.StartNew()

        ' 仅注册 AssemblyResolve 事件，不做预加载
        SqliteAssemblyResolver.EnsureRegistered()

        ' 初始化全局状态栏
        Try
            GlobalStatusStripAll.InitializeApplication(application)
        Catch ex As Exception
            Debug.WriteLine($"[Startup] Phase0 GlobalStatusStrip failed: {ex.Message}")
        End Try

        sw.Stop()
        Debug.WriteLine($"[Startup] Phase0-Critical 完成: {sw.ElapsedMilliseconds}ms")
    End Sub
End Class
