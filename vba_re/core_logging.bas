Attribute VB_Name = "core_logging"
' ==============================================================================================
' MODULE NAME       : core_logging
' LAYER             : core
' PURPOSE           : Provides a standardized, side-effect-free logging contract (formatting)
'                     plus lightweight performance timers (capability). Does not write logs.
' DEPENDS           : None
' NOTE              : - Query-only formatting helpers (no IO).
'                     - Timers use VBA.Timer (seconds since midnight) and handle midnight rollover.
'                     - Real sinks (file/sheet/immediate) and logging switches belong to Platform layer.
' STATUS            : Frozen
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Initial draft based on legacy implementation, iteratively refined during early refactor.

' v2.0.0
'   - Refactor: split project into layered architecture (Core / Platform / Business).
'   - Freeze: standardized log line formatting + context serialization + timer capability.

' v2.1.0
'   - Fix (Contract): StartTimer/ElapsedTime/ClearTimers contracts corrected to
'                     Core / Capability (stateful); Side Effects annotated accurately.
'   - Fix (Docs): LogInfo/LogWarn/LogError/LogDebug given independent comment blocks;
'                 LogDebug condition and empty-return behavior explicitly documented.
'   - Fix (Docs): FormatDateTimeWithMS annotated as pseudo-milliseconds (not hi-res clock).
'   - Fix (Reliability): BuildContextString On Error pattern unified to capture e before GoTo 0.
'   - Fix (Build): #Const DEBUG_MODE comment added; override via Project Properties documented.
'   - Removed: SafeWriteLog stub removed (was declared in TOC but never implemented).
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 01: LOG CONTRACTS & FORMAT (QUERY)
'   [F] BuildContextString      - Serializes context (primitive/dictionary/etc.) to string
'   [f] EscapeLogField          - Escapes separators to keep log lines parseable
'   [f] FormatDateTimeWithMS    - Formats datetime with pseudo-milliseconds slot
'   [F] FormatLogLine           - Formats standardized log line (no IO)
'   [F] LogInfo                 - Convenience: INFO line formatting
'   [F] LogWarn                 - Convenience: WARN line formatting
'   [F] LogError                - Convenience: ERROR line formatting
'   [F] LogDebug                - Convenience: DEBUG line formatting (#If DEBUG_MODE)
'
' SECTION 02: PERF TIMERS (CAPABILITY)
'   [S] StartTimer              - Starts/replaces a named timer
'   [F] ElapsedTime             - Returns elapsed seconds, optional reset
'   [F] FormatElapsedTime       - Formats seconds into readable string
'   [F] TimerExists             - Checks if a named timer exists
'   [F] GetAllTimers            - Returns all timer names
'   [S] ClearTimers             - Clears all timers
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function
' ==============================================================================================
Option Explicit
#Const DEBUG_MODE = False       ' Note: Override via Project Properties > Compile > Conditional Compilation Arguments: DEBUG_MODE=-1

' ============================================================
' SECTION 01: LOG CONTRACTS & FORMAT (QUERY)
' ============================================================

' ----------------------------------------------------------------------------------------------
' [F] BuildContextString
'
' 功能说明      : 将 context 序列化为可读字符串（Dictionary: k=v;...），用于日志行 Context 字段
' 参数          : context - 任意类型上下文（常见：Dictionary/String/Number/Boolean/Empty/Null）
' 返回          : String - 序列化后的上下文字符串（无则返回空字符串）
' Purpose       : Serializes context into a compact string for log output
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function BuildContextString(ByVal context As Variant) As String

    If IsEmpty(context) Or IsNull(context) Then
        BuildContextString = vbNullString
        Exit Function
    End If

    If IsObject(context) Then
        ' ---- Scripting.Dictionary (late-bound)
        If TypeName(context) = "Dictionary" Then
            If context.Count = 0 Then
                BuildContextString = vbNullString
                Exit Function
            End If

            Dim parts() As String
            ReDim parts(0 To context.Count - 1) As String

            Dim i As Long
            Dim k As Variant
            i = 0
            For Each k In context.Keys
                parts(i) = EscapeLogField(CStr(k)) & "=" & EscapeLogField(CStr(context.Item(k)))
                i = i + 1
            Next k

            BuildContextString = Join(parts, ";")
            Exit Function
        End If

        ' ---- Other objects: best-effort string conversion
        Dim e as Long
        On Error Resume Next
        BuildContextString = CStr(context)
        e = Err.Number
        On Error GoTo 0
        If e <> 0 Then
            BuildContextString = "[" & TypeName(context) & "]"
        End If
        
        Exit Function
    End If

    ' ---- Primitive types
    Dim ep as Long
    On Error Resume Next
    BuildContextString = EscapeLogField(CStr(context))
    ep = Err.Number
    On Error GoTo 0
    If ep <> 0 Then BuildContextString = vbNullString

End Function

' ----------------------------------------------------------------------------------------------
' [f] EscapeLogField
'
' 功能说明      : 转义日志字段中的分隔符，确保 FormatLogLine 输出满足“一条日志一行”的契约。
'               : 转义字符包括：\, |, ;, = 以及换行符 CR/LF。
' 参数          : value - 待转义的字段内容
' 返回          : String - 转义后的字段内容
' Purpose       : Escapes log fields to keep FormatLogLine output single-line and parseable
' ----------------------------------------------------------------------------------------------
Private Function EscapeLogField(ByVal value As String) As String
    Dim s As String
    s = value

    ' escape separators
    s = Replace$(s, "\", "\\")
    s = Replace$(s, "|", "\|")
    s = Replace$(s, ";", "\;")
    s = Replace$(s, "=", "\=")

    ' escape newlines to preserve single-line log contract
    s = Replace$(s, vbCr, "\r")
    s = Replace$(s, vbLf, "\n")

    EscapeLogField = s
End Function

' ----------------------------------------------------------------------------------------------
' [f] FormatDateTimeWithMS
'
' 功能说明      : 将日期时间格式化为包含毫秒的字符串（简洁时间格式：yyyy-mm-dd hh:nn:ss.000）
'               : 说明：毫秒位使用 VBA.Timer 偏移量计算，为伪毫秒（非系统高精度时钟）
' 参数          : dtValue - 要格式化的日期时间值
' 返回          : String - 包含毫秒的格式化日期时间字符串
' Purpose       : Formats datetime with pseudo-milliseconds slot (fixed format)
' ----------------------------------------------------------------------------------------------
Private Function FormatDateTimeWithMS(ByVal dtValue As Date) As String
    Dim ms As Long
    ms = CLng((dtValue - Int(dtValue)) * 86400000) Mod 1000
    If ms < 0 Then ms = ms + 1000
    
    FormatDateTimeWithMS = _
        Format$(dtValue, "yyyy-mm-dd hh:nn:ss") & "." & _
        Format$(ms, "000")
End Function

' ----------------------------------------------------------------------------------------------
' [F] FormatLogLine
'
' 功能说明      : 格式化标准日志行（不写入任何 sink，仅返回字符串）
' 参数          : level - 日志级别（DEBUG/INFO/WARN/ERROR/PERF）
'               : layer - 分层（CORE/PLAT/BIZ）
'               : moduleName - 模块名
'               : procName - 过程名
'               : message - 日志正文
'               : context - 可选上下文（Dictionary 或简单类型）
' 返回          : String - 标准化日志行
' Purpose       : Format standardized log line (no IO)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function FormatLogLine(ByVal level As String, _
                              ByVal layer As String, _
                              ByVal moduleName As String, _
                              ByVal procName As String, _
                              ByVal message As String, _
                              Optional ByVal context As Variant) As String

    Dim ts As String
    ts = FormatDateTimeWithMS(Now)

    Dim ctx As String
    ctx = BuildContextString(context)

    ' Keep line parseable: escape separators in free-text fields.
    level = Trim$(level)
    layer = Trim$(layer)
    moduleName = Trim$(moduleName)
    procName = Trim$(procName)

    FormatLogLine = ts _
        & " | " & level _
        & " | " & layer _
        & " | " & EscapeLogField(moduleName) _
        & " | " & EscapeLogField(procName) _
        & " | " & EscapeLogField(message) _
        & " | " & ctx

End Function

' ----------------------------------------------------------------------------------------------
' [F] LogInfo
'
' 功能说明      : 生成 INFO 级别的格式化日志字符串（不写入任何 sink）。
' 参数          : layer - 逻辑层级（如 "core"/"plat"/"biz" 或你项目约定的层）
'               : moduleName - 模块名称
'               : procName - 过程/函数名称
'               : message - 日志内容
'               : context - 可选，附加上下文（允许 Variant，用于传递更多信息）
' 返回          : String - 格式化后的日志行
' Purpose       : Generates a formatted INFO-level log line (formatting only; no sink)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function LogInfo(ByVal layer As String, _
                        ByVal moduleName As String, _
                        ByVal procName As String, _
                        ByVal message As String, _
                        Optional ByVal context As Variant) As String
    LogInfo = FormatLogLine("INFO", layer, moduleName, procName, message, context)
End Function

' ----------------------------------------------------------------------------------------------
' [F] LogWarn
'
' 功能说明      : 生成 WARN 级别的格式化日志字符串（不写入任何 sink）。
' 参数          : layer - 逻辑层级
'               : moduleName - 模块名称
'               : procName - 过程/函数名称
'               : message - 日志内容
'               : context - 可选，附加上下文（Variant）
' 返回          : String - 格式化后的日志行
' Purpose       : Generates a formatted WARN-level log line (formatting only; no sink)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function LogWarn(ByVal layer As String, _
                        ByVal moduleName As String, _
                        ByVal procName As String, _
                        ByVal message As String, _
                        Optional ByVal context As Variant) As String
    LogWarn = FormatLogLine("WARN", layer, moduleName, procName, message, context)
End Function

' ----------------------------------------------------------------------------------------------
' [F] LogError
'
' 功能说明      : 生成 ERROR 级别的格式化日志字符串（不写入任何 sink）。
' 参数          : layer - 逻辑层级
'               : moduleName - 模块名称
'               : procName - 过程/函数名称
'               : message - 日志内容
'               : context - 可选，附加上下文（Variant）
' 返回          : String - 格式化后的日志行
' Purpose       : Generates a formatted ERROR-level log line (formatting only; no sink)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function LogError(ByVal layer As String, _
                         ByVal moduleName As String, _
                         ByVal procName As String, _
                         ByVal message As String, _
                         Optional ByVal context As Variant) As String
    LogError = FormatLogLine("ERROR", layer, moduleName, procName, message, context)
End Function

' ----------------------------------------------------------------------------------------------
' [F] LogDebug
'
' 功能说明      : 生成 DEBUG 级别的格式化日志字符串（不写入任何 sink）。
'               : 仅在编译时定义了 DEBUG_MODE 时返回格式化日志行；否则始终返回空字符串
' 参数          : layer - 逻辑层级
'               : moduleName - 模块名称
'               : procName - 过程/函数名称
'               : message - 日志内容
'               : context - 可选，附加上下文（Variant）
' 返回          : String - DEBUG_MODE 下返回日志行；否则返回空字符串
' Purpose       : Generates a formatted DEBUG-level log line (formatting only; no sink)
'               : Returns empty string when DEBUG_MODE is not defined
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function LogDebug(ByVal layer As String, _
                         ByVal moduleName As String, _
                         ByVal procName As String, _
                         ByVal message As String, _
                         Optional ByVal context As Variant) As String
#If DEBUG_MODE Then
    LogDebug = FormatLogLine("DEBUG", layer, moduleName, procName, message, context)
#Else
    LogDebug = vbNullString
#End If
End Function

' ============================================================
' SECTION 02: PERF TIMERS (CAPABILITY)
' ============================================================

Private pTimers As Object   ' Scripting.Dictionary (late binding)

' ----------------------------------------------------------------------------------------------
' [S] StartTimer
'
' 功能说明      : 启动指定名称的计时器，记录当前时间用于后续性能测量
' 参数          : timerName - 计时器名称，用于唯一标识
' 返回          : 无 - Sub过程无返回值
' Purpose       : Starts a timer with the specified name, recording current time for subsequent performance measurement
' Contract      : Core / Capability (stateful)
' Side Effects  : Modifies internal timer state (pTimers dictionary)
' ----------------------------------------------------------------------------------------------
Public Sub StartTimer(ByVal timerName As String)
    
    If Len(timerName) = 0 Then Exit Sub
    
    If pTimers Is Nothing Then
        Set pTimers = CreateObject("Scripting.Dictionary")
    End If
    
    pTimers(timerName) = Timer
    
End Sub

' ----------------------------------------------------------------------------------------------
' [F] ElapsedTime
'
' 功能说明      : 获取指定计时器从启动到当前的时间差（秒），支持自动重置和跨午夜处理
' 参数          : timerName - 计时器名称
'               : reset - 可选，获取后是否重置（移除计时器），默认为True
'               : defaultValue - 可选，计时器不存在时返回的默认值，默认为0
' 返回          : Double - 经过的秒数，计时器不存在则返回默认值
' Purpose       : Gets the elapsed time in seconds for a specified timer, supports auto-reset and midnight crossing handling
' Contract      : Core / Capability (stateful)
' Side Effects  : If reset=True, removes timerName from pTimers; otherwise none
' ----------------------------------------------------------------------------------------------
Public Function ElapsedTime( _
        ByVal timerName As String, _
        Optional ByVal reset As Boolean = True, _
        Optional ByVal defaultValue As Double = 0) As Double
    
    If pTimers Is Nothing Then
        ElapsedTime = defaultValue
        Exit Function
    End If
    
    If Not pTimers.exists(timerName) Then
        ElapsedTime = defaultValue
        Exit Function
    End If
    
    Dim startTime As Double
    startTime = pTimers(timerName)
    
    Dim currentTime As Double
    currentTime = Timer
    
    Dim diff As Double
    diff = currentTime - startTime
    
    If diff < 0 Then diff = diff + 86400   ' 跨午夜
    
    If reset Then pTimers.Remove timerName
    
    ElapsedTime = diff
    
End Function

' ----------------------------------------------------------------------------------------------
' [F] FormatElapsedTime
'
' 功能说明      : 将秒数格式化为可读的时间字符串（保留三位小数并添加"s"单位）
' 参数          : seconds - 要格式化的秒数
' 返回          : String - 格式化后的时间字符串，如 "1.234s"
' Purpose       : Formats seconds to a readable time string (with 3 decimal places and "s" unit)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function FormatElapsedTime(ByVal seconds As Double) As String
    
    If seconds < 0 Then seconds = 0
    
    FormatElapsedTime = Format$(seconds, "0.000") & "s"
    
End Function

' ----------------------------------------------------------------------------------------------
' [F] TimerExists
'
' 功能说明      : 检查指定名称的计时器是否存在
' 参数          : timerName - 要检查的计时器名称
' 返回          : Boolean - 计时器是否存在，True表示存在
' Purpose       : Checks if a timer with the specified name exists
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function TimerExists(ByVal timerName As String) As Boolean
    
    If pTimers Is Nothing Then
        TimerExists = False
    Else
        TimerExists = pTimers.exists(timerName)
    End If
    
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetAllTimers
'
' 功能说明      : 获取所有已启动计时器的名称数组
'               : 说明：空时返回 `Array()`，调用方应先检查 `UBound(result) < LBound(result)`
' 参数          : None - 无参数
' 返回          : Variant - 包含所有计时器名称的数组，若无计时器则返回空数组
' Purpose       : Gets an array of all started timer names
' Contract      : Core / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function GetAllTimers() As Variant
    
    If pTimers Is Nothing Or pTimers.Count = 0 Then
        GetAllTimers = Array()
        Exit Function
    End If
    
    GetAllTimers = pTimers.Keys
    
End Function

' ----------------------------------------------------------------------------------------------
' [S] ClearTimers
'
' 功能说明      : 清除所有计时器，释放计时器字典资源
' 参数          : None - 无参数
' 返回          : 无 - Sub过程无返回值
' Purpose       : Clears all timers and releases the timer dictionary resource
' Contract      : Core / Capability (stateful)
' Side Effects  : Clears internal timer state (pTimers dictionary)
' ----------------------------------------------------------------------------------------------
Public Sub ClearTimers()
    
    If Not pTimers Is Nothing Then
        pTimers.RemoveAll
        Set pTimers = Nothing
    End If
    
End Sub
