Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Python

    Public Class PythonComputeExecutionResult
        Public Property Success As Boolean
        Public Property Data As JToken
        Public Property ErrorCode As String = ""
        Public Property ErrorMessage As String = ""
        Public Property DebugDetail As String = ""
        Public Property ExitCode As Integer = -1
        Public Property ElapsedMs As Long
        Public Property TimedOut As Boolean
    End Class

    ''' <summary>
    ''' 基础版受控 Python 计算通道。
    ''' Python 只接收 JSON 数据并返回 JSON 结果，不持有或操作 Office COM 对象。
    ''' 这不是操作系统级强沙箱；工具本身必须继续经过 SafetyGate 的用户审批。
    ''' </summary>
    Public NotInheritable Class PythonComputeService
        Private Const DefaultTimeoutSeconds As Integer = 20
        Private Const MaxTimeoutSeconds As Integer = 60
        Private Const MaxCodeChars As Integer = 16000
        Private Const MaxInputChars As Integer = 1000000
        Private Const MaxOutputChars As Integer = 2000000

        Private Shared ReadOnly AllowedImportRoots As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "math", "statistics", "datetime", "decimal", "collections", "re", "functools", "itertools"
        }

        Private Shared ReadOnly BlockedTokens As String() = {
            "__", "open(", "eval(", "exec(", "compile(", "breakpoint(", "input(",
            "globals(", "locals(", "vars(", "getattr(", "setattr(", "delattr(",
            "os.", "sys.", "subprocess", "socket", "pathlib", "shutil", "tempfile",
            "ctypes", "win32", "xlwings", "openpyxl", "requests", "urllib", "http.client",
            "builtins", "importlib", "marshal", "pickle", "shelve"
        }

        Private Sub New()
        End Sub

        Public Shared Async Function ExecuteAsync(params As JObject) As Task(Of PythonComputeExecutionResult)
            Dim sw = Stopwatch.StartNew()
            Dim result As New PythonComputeExecutionResult()
            Dim workDir As String = ""

            Try
                If params Is Nothing Then
                    Return Fail(result, ExceptionClassifier.CodeArgument, "PythonCompute 缺少参数", sw)
                End If

                Dim code = If(params("code")?.ToString(), "").Trim()
                If String.IsNullOrWhiteSpace(code) Then
                    Return Fail(result, ExceptionClassifier.CodeArgument, "PythonCompute 缺少 code", sw)
                End If
                If code.Length > MaxCodeChars Then
                    Return Fail(result, ExceptionClassifier.CodeArgument, $"Python 代码超过 {MaxCodeChars} 字符限制", sw)
                End If

                Dim safetyError = ValidateCode(code)
                If Not String.IsNullOrWhiteSpace(safetyError) Then
                    Return Fail(result, ExceptionClassifier.CodeSafetyBlocked, safetyError, sw)
                End If

                Dim inputToken = If(params("input"), JValue.CreateNull())
                Dim inputJson = inputToken.ToString(Formatting.None)
                If inputJson.Length > MaxInputChars Then
                    Return Fail(result, ExceptionClassifier.CodeArgument, $"Python 输入超过 {MaxInputChars} 字符限制", sw)
                End If

                Dim timeoutSeconds = DefaultTimeoutSeconds
                If params("timeoutSeconds") IsNot Nothing Then
                    Integer.TryParse(params("timeoutSeconds").ToString(), timeoutSeconds)
                End If
                timeoutSeconds = Math.Max(1, Math.Min(MaxTimeoutSeconds, timeoutSeconds))

                Dim pythonPath = FindPython()
                If String.IsNullOrWhiteSpace(pythonPath) Then
                    Return Fail(result, ExceptionClassifier.CodeNotFound, "未找到 Python 3 解释器，请安装 Python 并加入 PATH，或设置 OFFICE_AI_PYTHON_PATH", sw)
                End If

                workDir = Path.Combine(Path.GetTempPath(), "office-ai-python", Guid.NewGuid().ToString("N"))
                Directory.CreateDirectory(workDir)
                Dim scriptPath = Path.Combine(workDir, "compute.py")
                File.WriteAllText(scriptPath, BuildWrapper(code), New UTF8Encoding(False))

                Using process As New Process()
                    process.StartInfo.FileName = pythonPath
                    process.StartInfo.Arguments = QuoteArgument(scriptPath)
                    process.StartInfo.WorkingDirectory = workDir
                    process.StartInfo.UseShellExecute = False
                    process.StartInfo.CreateNoWindow = True
                    process.StartInfo.RedirectStandardInput = True
                    process.StartInfo.RedirectStandardOutput = True
                    process.StartInfo.RedirectStandardError = True
                    process.StartInfo.StandardOutputEncoding = Encoding.UTF8
                    process.StartInfo.StandardErrorEncoding = Encoding.UTF8
                    process.StartInfo.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"
                    process.StartInfo.EnvironmentVariables("PYTHONDONTWRITEBYTECODE") = "1"
                    process.StartInfo.EnvironmentVariables("PYTHONNOUSERSITE") = "1"
                    process.StartInfo.EnvironmentVariables("PYTHONPATH") = ""

                    process.Start()
                    Dim stdoutTask = process.StandardOutput.ReadToEndAsync()
                    Dim stderrTask = process.StandardError.ReadToEndAsync()
                    Await process.StandardInput.WriteAsync(inputJson)
                    process.StandardInput.Close()

                    Dim exited = Await Task.Run(Function() process.WaitForExit(timeoutSeconds * 1000))
                    If Not exited Then
                        result.TimedOut = True
                        Try
                            process.Kill()
                            process.WaitForExit(3000)
                        Catch
                        End Try
                    End If

                    Dim stdout = Await stdoutTask
                    Dim stderr = Await stderrTask
                    result.ExitCode = If(result.TimedOut, -1, process.ExitCode)

                    If result.TimedOut Then
                        Return Fail(result, ExceptionClassifier.CodeTimeout, $"Python 计算超过 {timeoutSeconds} 秒，已终止", sw)
                    End If
                    If stdout.Length > MaxOutputChars OrElse stderr.Length > MaxOutputChars Then
                        Return Fail(result, ExceptionClassifier.CodeArgument, "Python 输出超过大小限制", sw)
                    End If
                    If process.ExitCode <> 0 Then
                        Dim detail = If(String.IsNullOrWhiteSpace(stderr), "Python 进程执行失败", stderr.Trim())
                        If detail.IndexOf("PythonCompute rejected:", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Return Fail(result,
                                        ExceptionClassifier.CodeSafetyBlocked,
                                        "Python 代码不符合受控计算策略",
                                        sw,
                                        detail)
                        End If
                        Return Fail(result, ExceptionClassifier.CodeUnknown, "Python 计算失败", sw, detail)
                    End If
                    If String.IsNullOrWhiteSpace(stdout) Then
                        Return Fail(result, ExceptionClassifier.CodeJson, "Python 没有返回 JSON 结果", sw)
                    End If

                    Try
                        result.Data = JToken.Parse(stdout.Trim())
                    Catch ex As JsonReaderException
                        Return Fail(result,
                                    ExceptionClassifier.CodeJson,
                                    "Python 输出不是合法 JSON；代码只能通过 result 返回结果，请不要 print 调试信息",
                                    sw,
                                    ex.Message)
                    End Try
                End Using

                sw.Stop()
                result.Success = True
                result.ElapsedMs = sw.ElapsedMilliseconds
                Return result
            Catch ex As Exception
                Dim classified = ExceptionClassifier.Classify(ex)
                Return Fail(result, classified.ErrorCode, classified.UserMessage, sw, classified.DebugDetail)
            Finally
                SafeDeleteDirectory(workDir)
            End Try
        End Function

        Private Shared Function ValidateCode(code As String) As String
            Dim normalized = code.ToLowerInvariant()
            For Each token In BlockedTokens
                If normalized.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return $"PythonCompute 禁止使用: {token}"
                End If
            Next

            Dim lines = code.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ChrW(10))
            For Each rawLine In lines
                Dim line = rawLine.Trim()
                If line.StartsWith("import ", StringComparison.OrdinalIgnoreCase) Then
                    Dim imports = line.Substring(7).Split(","c)
                    For Each item In imports
                        Dim root = item.Trim().Split("."c, " "c)(0)
                        If Not AllowedImportRoots.Contains(root) Then Return $"PythonCompute 不允许导入模块: {root}"
                    Next
                ElseIf line.StartsWith("from ", StringComparison.OrdinalIgnoreCase) Then
                    Dim parts = line.Split(New Char() {" "c, ChrW(9)}, StringSplitOptions.RemoveEmptyEntries)
                    If parts.Length < 4 OrElse Not String.Equals(parts(2), "import", StringComparison.OrdinalIgnoreCase) Then
                        Return "PythonCompute 的 from import 语句格式无效"
                    End If
                    Dim root = parts(1).Split("."c)(0)
                    If Not AllowedImportRoots.Contains(root) Then Return $"PythonCompute 不允许导入模块: {root}"
                End If
            Next

            If normalized.IndexOf("result", StringComparison.OrdinalIgnoreCase) < 0 Then
                Return "PythonCompute 代码必须给 result 变量赋值"
            End If
            Return ""
        End Function

        Private Shared Function BuildWrapper(code As String) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("import ast")
            sb.AppendLine("import builtins")
            sb.AppendLine("import json")
            sb.AppendLine("import sys")
            sb.AppendLine("source = " & JsonConvert.SerializeObject(code))
            sb.AppendLine("allowed_roots = {'math','statistics','datetime','decimal','collections','re','functools','itertools'}")
            sb.AppendLine("blocked_calls = {'open','eval','exec','compile','breakpoint','input','globals','locals','vars','getattr','setattr','delattr','__import__'}")
            sb.AppendLine("tree = ast.parse(source, filename='<PythonCompute>', mode='exec')")
            sb.AppendLine("for node in ast.walk(tree):")
            sb.AppendLine("    if isinstance(node, ast.Import):")
            sb.AppendLine("        for alias in node.names:")
            sb.AppendLine("            if alias.name.split('.')[0] not in allowed_roots: raise RuntimeError('PythonCompute rejected: import ' + alias.name)")
            sb.AppendLine("    elif isinstance(node, ast.ImportFrom):")
            sb.AppendLine("        root = (node.module or '').split('.')[0]")
            sb.AppendLine("        if node.level != 0 or root not in allowed_roots: raise RuntimeError('PythonCompute rejected: import ' + (node.module or 'relative'))")
            sb.AppendLine("    elif isinstance(node, ast.Attribute) and node.attr.startswith('_'):")
            sb.AppendLine("        raise RuntimeError('PythonCompute rejected: private attribute access')")
            sb.AppendLine("    elif isinstance(node, ast.Name) and node.id.startswith('__'):")
            sb.AppendLine("        raise RuntimeError('PythonCompute rejected: private name access')")
            sb.AppendLine("    elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in blocked_calls:")
            sb.AppendLine("        raise RuntimeError('PythonCompute rejected: call ' + node.func.id)")
            sb.AppendLine("def safe_import(name, globals=None, locals=None, fromlist=(), level=0):")
            sb.AppendLine("    root = name.split('.')[0]")
            sb.AppendLine("    if level != 0 or root not in allowed_roots: raise RuntimeError('PythonCompute rejected: import ' + name)")
            sb.AppendLine("    return builtins.__import__(name, globals, locals, fromlist, level)")
            sb.AppendLine("safe_builtins = {")
            sb.AppendLine("    'abs':abs,'all':all,'any':any,'bool':bool,'dict':dict,'enumerate':enumerate,")
            sb.AppendLine("    'filter':filter,'float':float,'int':int,'isinstance':isinstance,'iter':iter,'len':len,")
            sb.AppendLine("    'list':list,'map':map,'max':max,'min':min,'next':next,'range':range,'reversed':reversed,")
            sb.AppendLine("    'round':round,'set':set,'sorted':sorted,'str':str,'sum':sum,'tuple':tuple,'zip':zip,")
            sb.AppendLine("    'Exception':Exception,'ValueError':ValueError,'TypeError':TypeError,'__import__':safe_import")
            sb.AppendLine("}")
            sb.AppendLine("scope = {'__builtins__': safe_builtins, 'input_data': json.load(sys.stdin)}")
            sb.AppendLine("exec(compile(tree, '<PythonCompute>', 'exec'), scope, scope)")
            sb.AppendLine("if 'result' not in scope:")
            sb.AppendLine("    raise RuntimeError('PythonCompute code must assign result')")
            sb.AppendLine("sys.stdout.write(json.dumps(scope['result'], ensure_ascii=False, default=str, separators=(',', ':')))")
            Return sb.ToString()
        End Function

        Private Shared Function FindPython() As String
            Dim configured = Environment.GetEnvironmentVariable("OFFICE_AI_PYTHON_PATH")
            If Not String.IsNullOrWhiteSpace(configured) AndAlso File.Exists(configured) Then Return configured

            For Each candidate In New String() {"python.exe", "python3.exe", "python", "python3"}
                Try
                    Using process As New Process()
                        process.StartInfo.FileName = candidate
                        process.StartInfo.Arguments = "--version"
                        process.StartInfo.UseShellExecute = False
                        process.StartInfo.CreateNoWindow = True
                        process.StartInfo.RedirectStandardOutput = True
                        process.StartInfo.RedirectStandardError = True
                        process.Start()
                        If process.WaitForExit(3000) AndAlso process.ExitCode = 0 Then Return candidate
                    End Using
                Catch
                End Try
            Next
            Return ""
        End Function

        Private Shared Function QuoteArgument(value As String) As String
            Return """" & If(value, "") & """"
        End Function

        Private Shared Function Fail(result As PythonComputeExecutionResult,
                                     errorCode As String,
                                     message As String,
                                     sw As Stopwatch,
                                     Optional debugDetail As String = "") As PythonComputeExecutionResult
            sw.Stop()
            result.Success = False
            result.ErrorCode = If(errorCode, ExceptionClassifier.CodeUnknown)
            result.ErrorMessage = If(message, "Python 计算失败")
            result.DebugDetail = AppLogger.Redact(If(String.IsNullOrWhiteSpace(debugDetail), result.ErrorMessage, debugDetail))
            result.ElapsedMs = sw.ElapsedMilliseconds
            Return result
        End Function

        Private Shared Sub SafeDeleteDirectory(path As String)
            If String.IsNullOrWhiteSpace(path) OrElse Not Directory.Exists(path) Then Return
            Try
                Directory.Delete(path, True)
            Catch ex As Exception
                AppLogger.Warn("PythonCompute", $"临时目录清理失败: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub
    End Class

End Namespace
