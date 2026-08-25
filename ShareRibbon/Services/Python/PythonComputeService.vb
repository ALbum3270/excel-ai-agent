Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
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
        Private Shared ReadOnly PythonResolutionLock As New Object()
        Private Shared _cachedPythonPath As String = ""

        Private Shared ReadOnly AllowedImportRoots As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "math", "statistics", "datetime", "decimal", "collections", "re", "functools", "itertools", "json"
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

        Public Shared Async Function ExecuteAsync(
            params As JObject,
            Optional cancellationToken As CancellationToken = Nothing) As Task(Of PythonComputeExecutionResult)
            Dim sw = Stopwatch.StartNew()
            Dim result As New PythonComputeExecutionResult()
            Dim workDir As String = ""

            Try
                cancellationToken.ThrowIfCancellationRequested()
                If params Is Nothing Then
                    Return Fail(result, ExceptionClassifier.CodeArgument, "PythonCompute 缺少参数", sw)
                End If

                Dim code = NormalizeTransportEscapes(If(params("code")?.ToString(), "")).Trim()
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
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim stdoutTask = process.StandardOutput.ReadToEndAsync()
                    Dim stderrTask = process.StandardError.ReadToEndAsync()
                    ' .NET Framework ProcessStartInfo has no StandardInputEncoding property.
                    ' Writing through StreamWriter would use the Windows console code page and
                    ' corrupt Chinese JSON before Python decodes stdin as UTF-8.
                    Dim inputBytes = New UTF8Encoding(False).GetBytes(inputJson)
                    Dim stdinFailure As Exception = Nothing
                    Try
                        Await process.StandardInput.BaseStream.WriteAsync(inputBytes, 0, inputBytes.Length)
                        Await process.StandardInput.BaseStream.FlushAsync()
                    Catch ex As IOException
                        ' Invalid generated source can make Python exit during ast.parse before it
                        ' consumes stdin. Preserve that primary Python error below instead of
                        ' misclassifying the secondary broken pipe as a filesystem IO failure.
                        stdinFailure = ex
                    Catch ex As ObjectDisposedException
                        stdinFailure = ex
                    Finally
                        Try
                            process.StandardInput.Close()
                        Catch
                        End Try
                    End Try

                    Dim exitTask = Task.Run(Sub() process.WaitForExit())
                    Dim timeoutTask = Task.Delay(timeoutSeconds * 1000)
                    Dim cancellationSignal = New TaskCompletionSource(Of Boolean)(
                        TaskCreationOptions.RunContinuationsAsynchronously)
                    Dim completed As Task
                    Using cancellationRegistration = cancellationToken.Register(
                        Sub() cancellationSignal.TrySetResult(True))
                        completed = Await Task.WhenAny(exitTask, timeoutTask, cancellationSignal.Task)
                    End Using

                    If completed Is cancellationSignal.Task Then
                        Try
                            process.Kill()
                            process.WaitForExit(3000)
                        Catch
                        End Try
                        cancellationToken.ThrowIfCancellationRequested()
                    End If

                    If completed Is timeoutTask Then
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
                        If detail.IndexOf("SyntaxError", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           detail.IndexOf("IndentationError", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           detail.IndexOf("TabError", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Return Fail(result,
                                        ExceptionClassifier.CodeArgument,
                                        "Python 代码存在语法错误",
                                        sw,
                                        detail)
                        End If
                        Return Fail(result, ExceptionClassifier.CodeUnknown, "Python 计算失败", sw, detail)
                    End If
                    If stdinFailure IsNot Nothing Then
                        Dim classified = ExceptionClassifier.Classify(stdinFailure)
                        Return Fail(result, classified.ErrorCode, classified.UserMessage, sw, classified.DebugDetail)
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
            Catch ex As OperationCanceledException
                Throw
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
                    Dim importItems = line.Substring(7).Split(","c)
                    For Each item In importItems
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

        ''' <summary>
        ''' Some OpenAI-compatible providers double-escape a multiline tool argument, so the
        ''' parsed JSON string still contains the two characters "\n" instead of a newline.
        ''' Decode only that transport shape: real multiline source is left byte-for-byte
        ''' unchanged, while escaped quotes/backslashes remain Python source rather than being
        ''' interpreted a second time.
        ''' </summary>
        Private Shared Function NormalizeTransportEscapes(code As String) As String
            If String.IsNullOrEmpty(code) Then Return If(code, "")
            If code.IndexOf(ChrW(10)) >= 0 OrElse code.IndexOf(ChrW(13)) >= 0 Then Return code
            If code.IndexOf("\n", StringComparison.Ordinal) < 0 AndAlso
               code.IndexOf("\r", StringComparison.Ordinal) < 0 Then Return code

            Return code.Replace("\r\n", vbLf).
                        Replace("\n", vbLf).
                        Replace("\r", vbLf).
                        Replace("\t", vbTab)
        End Function

        Private Shared Function BuildWrapper(code As String) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("import ast")
            sb.AppendLine("import builtins")
            sb.AppendLine("import json")
            sb.AppendLine("import sys")
            sb.AppendLine("source = " & JsonConvert.SerializeObject(code))
            sb.AppendLine("allowed_roots = {'math','statistics','datetime','decimal','collections','re','functools','itertools','json'}")
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
            ' .NET Framework creates StandardInput through a StreamWriter which can emit a
            ' UTF-8 preamble before callers write bytes to BaseStream.  utf-8-sig accepts both
            ' BOM and non-BOM input while preserving non-ASCII workbook values such as Chinese.
            sb.AppendLine("input_bytes = sys.stdin.buffer.read()")
            sb.AppendLine("scope = {'__builtins__': safe_builtins, 'input_data': json.loads(input_bytes.decode('utf-8-sig'))}")
            sb.AppendLine("exec(compile(tree, '<PythonCompute>', 'exec'), scope, scope)")
            sb.AppendLine("if 'result' not in scope:")
            sb.AppendLine("    raise RuntimeError('PythonCompute code must assign result')")
            sb.AppendLine("sys.stdout.write(json.dumps(scope['result'], ensure_ascii=False, default=str, separators=(',', ':')))")
            Return sb.ToString()
        End Function

        Private Shared Function FindPython() As String
            SyncLock PythonResolutionLock
                If Not String.IsNullOrWhiteSpace(_cachedPythonPath) Then Return _cachedPythonPath

                Dim candidates As New List(Of String)()
                AddPythonCandidate(candidates, Environment.GetEnvironmentVariable("OFFICE_AI_PYTHON_PATH"))

                Dim condaPrefix = Environment.GetEnvironmentVariable("CONDA_PREFIX")
                If Not String.IsNullOrWhiteSpace(condaPrefix) Then
                    AddPythonCandidate(candidates, Path.Combine(condaPrefix, "python.exe"))
                End If

                Dim virtualEnv = Environment.GetEnvironmentVariable("VIRTUAL_ENV")
                If Not String.IsNullOrWhiteSpace(virtualEnv) Then
                    AddPythonCandidate(candidates, Path.Combine(virtualEnv, "Scripts", "python.exe"))
                    AddPythonCandidate(candidates, Path.Combine(virtualEnv, "bin", "python"))
                End If

                For Each commandName In New String() {"python.exe", "python3.exe", "python", "python3"}
                    AddPythonCandidate(candidates, commandName)
                Next

                For Each candidate In candidates
                    If CanRunPython(candidate) Then
                        _cachedPythonPath = candidate
                        AppLogger.Info("PythonCompute", $"Python runtime resolved source={DescribePythonSource(candidate)}")
                        Return _cachedPythonPath
                    End If
                Next
                Return ""
            End SyncLock
        End Function

        Private Shared Sub AddPythonCandidate(candidates As List(Of String), candidate As String)
            If candidates Is Nothing OrElse String.IsNullOrWhiteSpace(candidate) Then Return
            Dim normalized = candidate.Trim().Trim(""""c)
            If candidates.Any(Function(existing) String.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase)) Then Return
            candidates.Add(normalized)
        End Sub

        Private Shared Function CanRunPython(candidate As String) As Boolean
            If String.IsNullOrWhiteSpace(candidate) Then Return False
            If (candidate.Contains(Path.DirectorySeparatorChar) OrElse candidate.Contains(Path.AltDirectorySeparatorChar)) AndAlso
               Not File.Exists(candidate) Then Return False

            Try
                Using process As New Process()
                    process.StartInfo.FileName = candidate
                    process.StartInfo.Arguments = "--version"
                    process.StartInfo.UseShellExecute = False
                    process.StartInfo.CreateNoWindow = True
                    process.StartInfo.RedirectStandardOutput = True
                    process.StartInfo.RedirectStandardError = True
                    process.Start()
                    Return process.WaitForExit(1500) AndAlso process.ExitCode = 0
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Shared Function DescribePythonSource(candidate As String) As String
            Dim configured = Environment.GetEnvironmentVariable("OFFICE_AI_PYTHON_PATH")
            If Not String.IsNullOrWhiteSpace(configured) AndAlso
               String.Equals(configured.Trim().Trim(""""c), candidate, StringComparison.OrdinalIgnoreCase) Then Return "configured"

            Dim condaPrefix = Environment.GetEnvironmentVariable("CONDA_PREFIX")
            If Not String.IsNullOrWhiteSpace(condaPrefix) AndAlso
               candidate.StartsWith(condaPrefix, StringComparison.OrdinalIgnoreCase) Then Return "conda"

            Dim virtualEnv = Environment.GetEnvironmentVariable("VIRTUAL_ENV")
            If Not String.IsNullOrWhiteSpace(virtualEnv) AndAlso
               candidate.StartsWith(virtualEnv, StringComparison.OrdinalIgnoreCase) Then Return "virtualenv"
            Return "path"
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
