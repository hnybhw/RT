Attribute VB_Name = "plat_runtime"
' ==============================================================================================
' MODULE NAME       : plat_runtime
' LAYER             : platform
' PURPOSE           : Platform runtime execution engine. Provides the TResult contract,
'                     Application state management, safe business entry execution with
'                     retry policy, unified error handling, Workbook IO helpers,
'                     logging convenience wrappers, and performance scaffolding.
' DEPENDS           : plat_context v2.1.2 (InitContext, GetLogger)
'                     plat_logger  v2.1.2 (WriteLog signature via CallByName)
'                     core_logging v2.1.0 (StartTimer, ElapsedTime, FormatElapsedTime, BuildContextString)
' NOTE              : - Business MUST call plat_runtime.ResultStore before returning from
'                       any entry point invoked by SafeExecute.
'                     - RunOptimized / RunOptimizedWithArg are compatibility entries for
'                       legacy procedures that do not yet implement the TResult contract.
'                     - This module is the sole owner of Application state save/restore;
'                       Business layer must not mutate Application settings directly.
'                     - Logging helpers route through plat_context.GetLogger() via late
'                       binding; logging failures are suppressed and never raised to callers.
' STATUS            : In Development (not yet reviewed for Stable Candidate)
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Runtime helpers existed as scattered procedures inside monolithic script; no contract, no lifecycle management.
'   - Init (Design): Application state mutation and error handling were inline per procedure; no unified failure semantics.

' v2.0.0
'   - Refactor (Architecture): Extracted into Platform layer under layered model (Core / Platform / Business); positioned as runtime boundary between Business logic and Excel Application state.
'   - Refactor (Contract): Introduced TResult as unified Business/Platform result contract; Business entry points must call ResultStore (no MsgBox, no Raise).
'   - Refactor (Lifecycle): Formalized SaveAppState / RestoreAppState / ApplyOptimizedSettings with deterministic cleanup via CleanExit pattern.
'   - Refactor (Safe Execute): Introduced SafeExecute with UNSET contract guard, retry loop, and RunId correlation; each attempt isolated in private helper to guarantee reliable On Error semantics (function-scoped).
'   - Refactor (Error Policy): Centralized HandleError / BuildErrorMessage; UI policy suppressed by default (log-only); MsgBox reserved for dev mode.
'   - Refactor (IO): Consolidated Workbook IO helpers (OpenWorkbook, GetOpenWorkbookByName, IsSharePointUrlAccessible, GetFilePathFromDialog) as Platform-owned IO.
'   - Refactor (Logging): Introduced LogDebug/Info/Warn/Error/Perf wrappers routing through WritePlatLog -> plat_context.GetLogger() via CallByName (late binding).
'   - Fix (Retry Reliability): SafeExecute retry loop delegates each attempt to private TryOnce helper; On Error is function-scoped per attempt, eliminating VBA loop-level error handler unreliability.
'   - Fix (Params): Removed RunOptimizedWithParams ParamArray forwarding to Application.Run; replaced with RunOptimizedWithArg single-Variant overload (VBA Application.Run limitation).
'   - Fix (StatusBar): RestoreAppState explicitly guards StatusBar restore type (False vs String) to avoid implicit coercion dependency.
'   - Fix (Interface): WritePlatLog calls plat_context.GetLogger() (Function form, v2.1.1 confirmed); aligns with frozen plat_context interface.

' v2.1.0
'   - Fix (Logging Contract): LogXxx / WritePlatLog context widened from String to Optional Variant; aligns with plat_logger.WriteLog and core_logging.BuildContextString signatures.
'   - Fix (Log Consistency): HandleError/TryOnce unified to two-segment convention (message=short label; context=structured key-value).
'   - Fix (IO Robustness): GetFilePathFromDialog uses late-bound Object for FileDialog to reduce Office reference coupling in platform module.
'   - Improve (Perf): MeasurePerf implemented as minimal viable start/end wrapper over core_logging.StartTimer / ElapsedTime / FormatElapsedTime; DEPENDS aligned.
'   - Align (Status): STATUS set to Stable Candidate pending review; not yet Frozen.
'   - Fix (StatusBar Save Guard): Added VarType guard to SaveAppState to match RestoreAppState; Save/Restore now form symmetric type guards (no implicit coercion).
'   - Fix (HandleError Contract): context widened to Variant; BuildErrorMessage uses BuildContextString for serialization
'   - Fix (TryOnce Observability): procName added to exception context; Err.Description trimmed

' v2.1.1
'   - Align (Dependencies): Updated DEPENDS to plat_context v2.1.2 / plat_logger v2.1.2.
'   - Fix (IO Leak): IsSharePointUrlAccessible adopts CleanExit pattern to guarantee workbook closure on both success and exception paths.
'   - Fix (TryOnce Context): Err.Description prefixed with err_desc= to conform to key=value convention throughout context strings.
'   - Fix (CreateRunId Seeding): Randomize moved to lazy one-time init via gRunIdSeeded flag; eliminates repeated same-tick reseeding and reduces collision risk.
'   - Fix (SafeExecute Clarity): Done: label now reads result from local var r instead of function name slot; eliminates recursive-call ambiguity.
'   - Fix (SafeExecute Declarations): sentinel and attemptOk moved to function top; aligns with module declaration convention (no loop-body Dim).

' v2.1.2
'   - Fix (IO Leak): IsSharePointUrlAccessible sets wb = Nothing after first Close; CleanExit guard now evaluates correctly on success path (eliminates double-close).
'   - Fix (Separator Consistency): BuildErrorMessage separators unified to ";" (no space) matching all other context strings in module; ctx= prefix applied consistently.
'   - Fix (TryOnce Declarations): Dim dbg moved to function top; aligns with module declaration convention (no EH-body Dim).
'   - Align (SafeExecute Note): InitContext failure propagation documented as intentional fatal design boundary.
'   - Align (GetFilePathFromDialog Note): No-guard policy documented; call site assumed interactive and UI-capable.
'   - Fix (BuildErrorMessage Key Consistency): err_desc= key added to errDescription field; aligns with TryOnce context convention.

' v2.2.0
'   - Add (Architecture): Introduced SECTION 08 Session Teardown, absorbing all worksheet lifecycle management from app_06_ws (v1 script module).
'   - Add (Contract): SessionTeardown / PurgeTempSheets / RefreshOutputSheetList adopt Boolean + ByRef errMsg contract; no MsgBox, no Raise.
'   - Add (Config-Driven): Sheet action rules migrated from hardcoded string-matching to Setup-table lookup via p_ResolveSheetAction; extensible for v3.
'   - Add (Safety): Permanent sheet guard enforced via p_IsPermanentSheet; logic centralized and not duplicated across callers.
'   - Align (Ownership): plat_runtime is now the sole owner of Workbook structural mutations (sheet delete / archive); Business layer must not delete sheets.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 00: MODULE STATE & CONSTANTS
'
' SECTION 01: RESULT CONTRACT
'   [T] TResult
'   [F] ResultOk                - Build success TResult
'   [F] ResultFail              - Build failure TResult
'   [S] ResultStore             - Store last TResult into module slot
'   [F] ResultFetch             - Fetch last stored TResult
'
' SECTION 02: APPLICATION STATE
'   [T] TAppState
'   [F] SaveAppState            - Snapshot Excel Application state
'   [S] RestoreAppState         - Restore Excel Application state (best-effort)
'   [s] ApplyOptimizedSettings  - Set optimized Excel runtime settings
'   [S] RunOptimized            - Execute VBA procedure by name under optimized settings
'   [S] RunOptimizedWithArg     - Execute VBA procedure by name with one argument under optimized settings
'
' SECTION 03: SAFE EXECUTE / RETRY POLICY
'   [F] SafeExecute             - Execute business entry with contract guard and optional retry
'   [f] TryOnce                 - Single attempt wrapper (isolated On Error scope)
'   [f] CreateRunId             - Generate correlation ID for log tracing
'   [f] WaitMs                  - Best-effort delay without WinAPI dependency
'
' SECTION 04: ERROR POLICY
'   [S] HandleError             - Unified platform error handling (log + optional UI)
'   [f] BuildErrorMessage       - Standardized error message format for logging
'
' SECTION 05: WORKBOOK IO
'   [F] OpenWorkbook            - Open workbook with unified error handling
'   [F] GetOpenWorkbookByName   - Get open workbook by name (best-effort)
'   [F] IsSharePointUrlAccessible - Minimal SharePoint URL access test
'   [F] GetFilePathFromDialog   - Standard single-file picker dialog
'
' SECTION 06: LOGGING HELPERS
'   [S] LogDebug                - Convenience wrapper: DEBUG level
'   [S] LogInfo                 - Convenience wrapper: INFO level
'   [S] LogWarn                 - Convenience wrapper: WARN level
'   [S] LogError                - Convenience wrapper: ERROR level
'   [S] LogPerf                 - Convenience wrapper: PERF level
'   [s] WritePlatLog            - Bridge to plat_logger via plat_context.GetLogger (late binding; context as Variant)
'
' SECTION 07: PERF INSTRUMENTATION
'   [S] MeasurePerf             - Perf start/end wrapper over core_logging.StartTimer / ElapsedTime / FormatElapsedTime

' SECTION 08: SESSION TEARDOWN
'   [F] SessionTeardown         - Archive and delete output sheets per config-driven action rules
'   [F] PurgeTempSheets         - Delete all sheets matching TEMP_ prefix (safe, skips permanent)
'   [F] RefreshOutputSheetList  - Refresh output sheet inventory in Setup anchor range
'   [f] p_ResolveSheetAction    - Resolve action rule for a sheet name from Setup config table
'   [f] p_ExecuteSheetAction    - Execute a single sheet action (archive-delete / delete / keep)
'   [f] p_IsPermanentSheet      - Return True if sheet must never be deleted
'   [f] p_EnsureArchiveSheet    - Return or create the Archive worksheet
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function, [T]=Type
'       Rule: Helper functions and private procedures inherit the Contract and
'             Side Effects of their parent public API unless explicitly stated otherwise.
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 00: MODULE STATE & CONSTANTS
' ==============================================================================================

Private Const PLAT_LAYER    As String = "PLAT"
Private Const THIS_MODULE   As String = "plat_runtime"

Private gLastResult As TResult
Private gRunIdSeeded As Boolean

' ==============================================================================================
' SECTION 01: RESULT CONTRACT
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [T] TResult
'
' 功能说明      : 业务层与平台层之间的统一结果契约类型
'               : 业务入口必须通过 ResultStore 返回此类型；禁止抛出异常或弹出 MsgBox
' Purpose       : Unified result contract between Business and Platform layers
' ----------------------------------------------------------------------------------------------
Public Type TResult
    Ok       As Boolean
    Code     As String      ' "OK" / "VALIDATION" / "IO" / "CONTRACT" / "UNSET" / "UNHANDLED"
    UserMsg  As String      ' 面向用户的消息（可用于 UI 展示）
    DebugMsg As String      ' 调试细节（可选，仅用于日志）
    RunId    As String      ' 关联追踪 ID（由 SafeExecute 生成并注入）
    Action   As String      ' 业务动作标签
    Value    As Variant     ' 可选业务数据载体（DTO / Variant / Object）
End Type

' ----------------------------------------------------------------------------------------------
' [F] ResultOk
'
' 功能说明      : 构造成功的 TResult
' 参数          : actionName - 业务动作标签
'               : RunId      - 关联追踪 ID
'               : Value      - 可选业务数据载体
' 返回          : TResult    - 成功结果（Ok = True，Code = "OK"）
' Purpose       : Factory for success TResult
' Contract      : Platform / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function ResultOk( _
    ByVal actionName As String, _
    ByVal RunId As String, _
    Optional ByVal Value As Variant _
) As TResult
    Dim r As TResult
    r.Ok = True
    r.Code = "OK"
    r.UserMsg = vbNullString
    r.DebugMsg = vbNullString
    r.RunId = RunId
    r.Action = actionName
    r.Value = Value
    ResultOk = r
End Function

' ----------------------------------------------------------------------------------------------
' [F] ResultFail
'
' 功能说明      : 构造失败的 TResult
' 参数          : actionName - 业务动作标签
'               : RunId      - 关联追踪 ID
'               : code       - 稳定错误码（如 "VALIDATION" / "IO" / "UNHANDLED"）
'               : userMsg    - 面向用户的消息
'               : debugMsg   - 可选技术细节
' 返回          : TResult    - 失败结果（Ok = False）
' Purpose       : Factory for failure TResult
' Contract      : Platform / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function ResultFail( _
    ByVal actionName As String, _
    ByVal RunId As String, _
    ByVal Code As String, _
    ByVal UserMsg As String, _
    Optional ByVal DebugMsg As String = vbNullString _
) As TResult
    Dim r As TResult
    r.Ok = False
    r.Code = Code
    r.UserMsg = UserMsg
    r.DebugMsg = DebugMsg
    r.RunId = RunId
    r.Action = actionName
    r.Value = Empty
    ResultFail = r
End Function

' ----------------------------------------------------------------------------------------------
' [S] ResultStore
'
' 功能说明      : 将业务结果写入模块级结果槽；业务入口必须在返回前调用此方法
' 参数          : r - 要存储的 TResult
' 返回          : 无
' Purpose       : Stores business result into module slot; mandatory for SafeExecute contract
' Contract      : Platform / State mutation
' Side Effects  : Modifies gLastResult
' ----------------------------------------------------------------------------------------------
Public Sub ResultStore(ByRef r As TResult)
    gLastResult = r
End Sub

' ----------------------------------------------------------------------------------------------
' [F] ResultFetch
'
' 功能说明      : 取回最近一次存储的 TResult
' 返回          : TResult - 上一次 ResultStore 写入的结果
' Purpose       : Fetches last stored TResult (called by SafeExecute after business entry returns)
' Contract      : Platform / Query-only
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function ResultFetch() As TResult
    ResultFetch = gLastResult
End Function

' ==============================================================================================
' SECTION 02: APPLICATION STATE
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [T] TAppState
'
' 功能说明      : Excel Application 状态快照类型，用于保存和恢复运行前状态
' Purpose       : Snapshot type for Excel Application settings
' ----------------------------------------------------------------------------------------------
Private Type TAppState
    ScreenUpdating As Boolean
    EnableEvents   As Boolean
    DisplayAlerts  As Boolean
    Calculation    As XlCalculation
    StatusBarIsAuto As Boolean      ' True = Application.StatusBar was False (auto-controlled)
    StatusBarText   As String       ' Saved text when StatusBar was not auto
End Type

' ----------------------------------------------------------------------------------------------
' [F] SaveAppState
'
' 功能说明      : 快照当前 Excel Application 状态，供后续 RestoreAppState 恢复
' 返回          : TAppState - 状态快照
' Purpose       : Snapshot current Excel Application state before optimized run
' Contract      : Platform / Query-only (reads Application state only)
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function SaveAppState() As TAppState
    Dim s As TAppState
    With Application
        s.ScreenUpdating = .ScreenUpdating
        s.EnableEvents = .EnableEvents
        s.DisplayAlerts = .DisplayAlerts
        s.Calculation = .Calculation

        ' VarType guard required: .StatusBar is False (Boolean) when auto-controlled, or a String.
        ' Implicit coercion on Boolean/String comparison is intentionally avoided.
        If VarType(.StatusBar) = vbBoolean And .StatusBar = False Then
            s.StatusBarIsAuto = True
            s.StatusBarText = vbNullString
        Else
            s.StatusBarIsAuto = False
            s.StatusBarText = CStr(.StatusBar)
        End If
    End With
    SaveAppState = s
End Function

' ----------------------------------------------------------------------------------------------
' [S] RestoreAppState
'
' 功能说明      : 恢复 Excel Application 状态至快照值，尽力恢复（单项失败不中断）
' 参数          : s - 由 SaveAppState 生成的状态快照
' 返回          : 无
' Purpose       : Restore Excel Application state from snapshot (best-effort per property)
' Contract      : Platform / State mutation
' Side Effects  : Modifies Application.ScreenUpdating, EnableEvents, DisplayAlerts,
'               : Calculation, StatusBar
' ----------------------------------------------------------------------------------------------
Public Sub RestoreAppState(ByRef s As TAppState)
    On Error Resume Next
    With Application
        .ScreenUpdating = s.ScreenUpdating
        .EnableEvents = s.EnableEvents
        .DisplayAlerts = s.DisplayAlerts
        .Calculation = s.Calculation

        ' Restore StatusBar using explicit type guard (no implicit coercion)
        If s.StatusBarIsAuto Then
            .StatusBar = False
        Else
            .StatusBar = s.StatusBarText
        End If
    End With
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------------------------
' [s] ApplyOptimizedSettings
'
' 功能说明      : 将 Excel 切换至高性能模式，关闭屏幕刷新、事件、弹窗，设置手动计算
' 返回          : 无
' Purpose       : Set Excel runtime to optimized mode for large-scale computation
' ----------------------------------------------------------------------------------------------
Private Sub ApplyOptimizedSettings()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "RT Running..."
    End With
End Sub

' ----------------------------------------------------------------------------------------------
' [S] RunOptimized
'
' 功能说明      : 以优化模式执行指定 VBA 过程（不使用 TResult 契约的兼容入口）
'               : 自动保存/恢复 AppState，调用前初始化 Context
' 参数          : procName - 要执行的过程名称（字符串形式）
' 返回          : 无
' Purpose       : Compatibility entry - run a named VBA procedure under optimized Excel settings
'               : Note - MeasurePerf can be added around Application.Run for legacy perf diagnostics.
' Contract      : Platform / Side-effecting
' Side Effects  : Modifies Application state (via ApplyOptimizedSettings / RestoreAppState)
'               : Calls plat_context.InitContext; errors logged via HandleError
' ----------------------------------------------------------------------------------------------
Public Sub RunOptimized(ByVal procName As String)
    Dim st As TAppState
    st = SaveAppState()

    On Error GoTo EH

    plat_context.InitContext False
    ApplyOptimizedSettings

    LogInfo THIS_MODULE, "RunOptimized", "start", "proc=" & procName
    Application.Run procName
    LogInfo THIS_MODULE, "RunOptimized", "end", "proc=" & procName

CleanExit:
    RestoreAppState st
    Exit Sub

EH:
    HandleError THIS_MODULE, "RunOptimized", Err.Number, Err.Description, "proc=" & procName
    Resume CleanExit
End Sub

' ----------------------------------------------------------------------------------------------
' [S] RunOptimizedWithArg
'
' 功能说明      : 以优化模式执行指定 VBA 过程，并传入一个 Variant 参数
'               : VBA Application.Run 不支持 ParamArray 展开；单参数足以覆盖 Legacy 过渡需求
' 参数          : procName - 要执行的过程名称
'               : arg      - 传入目标过程的单个参数
' 返回          : 无
' Purpose       : Compatibility entry - run a named VBA procedure with a single argument
'               : under optimized settings (single-arg form; Application.Run limitation)
'               : Inherits RunOptimized perf note; MeasurePerf wrapper applies equally.
' Contract      : Platform / Side-effecting
' Side Effects  : Inherits RunOptimized side effects
' ----------------------------------------------------------------------------------------------
Public Sub RunOptimizedWithArg(ByVal procName As String, ByVal arg As Variant)
    Dim st As TAppState
    st = SaveAppState()

    On Error GoTo EH

    plat_context.InitContext False
    ApplyOptimizedSettings

    LogInfo THIS_MODULE, "RunOptimizedWithArg", "start", "proc=" & procName
    Application.Run procName, arg
    LogInfo THIS_MODULE, "RunOptimizedWithArg", "end", "proc=" & procName

CleanExit:
    RestoreAppState st
    Exit Sub

EH:
    HandleError THIS_MODULE, "RunOptimizedWithArg", Err.Number, Err.Description, "proc=" & procName
    Resume CleanExit
End Sub

' ==============================================================================================
' SECTION 03: SAFE EXECUTE / RETRY POLICY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] SafeExecute
'
' 功能说明      : 以契约保护模式执行业务入口，支持可选重试，记录完整执行日志
'               : 业务过程签名约定：Sub BizEntry(ByVal actionName As String, ByVal RunId As String)
'               : 业务过程必须在返回前调用 plat_runtime.ResultStore(r As TResult)
'               : 每次尝试通过 TryOnce 隔离，保证 On Error 语义为函数级，避免循环内不可靠行为
'               : 只对异常重试，不对业务失败重试
' 参数          : actionName - 业务动作标签（用于日志与结果关联）
'               : procName   - 要执行的过程名称（Application.Run 形式）
'               : retries    - 重试次数（0 = 不重试，仅执行一次）
'               : delayMs    - 重试间隔毫秒数
' 返回          : TResult    - 最终执行结果
' Purpose       : Safe business entry executor with UNSET contract guard, retry loop,
'               : RunId correlation, and structured logging
' Note          : Retry applies to VBA runtime exceptions only (TryOnce returns False);
'               : business failures (Ok=False from ResultStore) are returned immediately
'               : without retry. To retry on specific result codes (e.g. transient IO),
'               : a retryOnCodes policy parameter would be needed (not yet implemented).
'               : InitContext failure is not guarded; propagates as unhandled VBA exception
'               : (platform init failure is treated as fatal by design).
' Contract      : Platform / Orchestration
' Side Effects  : Calls plat_context.InitContext; calls Application.Run procName;
'               : writes to log via LogXxx; modifies gLastResult via ResultStore/Fetch
' ----------------------------------------------------------------------------------------------
Public Function SafeExecute( _
    ByVal actionName As String, _
    ByVal procName As String, _
    Optional ByVal retries As Long = 0, _
    Optional ByVal delayMs As Long = 0 _
) As TResult

    Dim RunId As String
    RunId = CreateRunId()

    plat_context.InitContext False
    LogInfo THIS_MODULE, "SafeExecute", "start", _
            "action=" & actionName & ";run_id=" & RunId & ";proc=" & procName

    Dim i As Long
    Dim r As TResult
    Dim sentinel As TResult
    Dim attemptOk As Boolean

    For i = 0 To retries
        If i > 0 Then
            LogWarn THIS_MODULE, "SafeExecute", "retry", _
                    "action=" & actionName & ";attempt=" & CStr(i) & ";run_id=" & RunId
            If delayMs > 0 Then WaitMs delayMs
        End If

        ' Reset result slot to UNSET sentinel before each attempt (contract guard)
        sentinel = ResultFail(actionName, RunId, "UNSET", _
                              "未设置结果：业务入口必须调用 plat_runtime.ResultStore", _
                              "proc=" & procName)
        ResultStore sentinel

        ' Delegate to isolated helper; On Error is function-scoped inside TryOnce
        attemptOk = TryOnce(procName, actionName, RunId, i)
        
        If attemptOk Then
            r = ResultFetch()

            ' Contract guard: business must have replaced UNSET
            If r.Code = "UNSET" Then
                r = ResultFail(actionName, RunId, "CONTRACT", _
                            "业务未返回结果（契约违规）", "proc=" & procName)
            End If
            SafeExecute = r
            GoTo Done
        End If

        ' attemptOk = False means TryOnce caught a VBA exception; loop continues
    Next i

    ' All attempts raised unhandled exceptions
    r = ResultFail(actionName, RunId, "UNHANDLED", _
                   "运行失败，请查看日志", "proc=" & procName)
    SafeExecute = r

Done:
    LogInfo THIS_MODULE, "SafeExecute", "end", _
            "action=" & actionName & ";run_id=" & RunId & _
            ";ok=" & CStr(r.Ok) & ";code=" & r.Code
End Function

' ----------------------------------------------------------------------------------------------
' [f] TryOnce
'
' 功能说明      : 单次执行尝试的隔离包装；On Error 为函数级作用域，保证每次重试的错误处理可靠
'               : 成功时返回 True（业务可能仍返回失败 TResult，由 SafeExecute 判断）
'               : 捕获 VBA 运行时异常时记录日志并返回 False（触发重试）
' 参数          : procName   - 要执行的过程名称
'               : actionName - 业务动作标签
'               : RunId      - 关联追踪 ID
'               : attempt    - 当前尝试序号（用于日志）
' 返回          : Boolean    - True = 过程正常执行完毕；False = 抛出未处理异常
' Purpose       : Isolated single-attempt wrapper ensuring function-scoped On Error per retry
' ----------------------------------------------------------------------------------------------
Private Function TryOnce( _
    ByVal procName As String, _
    ByVal actionName As String, _
    ByVal RunId As String, _
    ByVal attempt As Long _
) As Boolean
    Dim dbg As String

    On Error GoTo EH

    Application.Run procName, actionName, RunId
    TryOnce = True
    Exit Function

EH:
    dbg = "proc=" & procName & ";attempt=" & CStr(attempt) & _
      ";err=" & CStr(Err.Number) & ";err_desc=" & Trim$(Err.Description)
    LogError THIS_MODULE, "TryOnce", "exception", _
             "action=" & actionName & ";run_id=" & RunId & ";" & dbg
    TryOnce = False
End Function

' ----------------------------------------------------------------------------------------------
' [f] CreateRunId
'
' 功能说明      : 生成用于日志关联追踪的运行 ID，格式：yyyymmdd-hhnnss-XXXXXX
'               : 精度为秒级；同一秒内高频调用存在极低碰撞概率（日志关联用途可接受）
' 返回          : String - 关联 ID 字符串
' Purpose       : Generate correlation ID for log tracing without external dependency
' ----------------------------------------------------------------------------------------------
Private Function CreateRunId() As String
    If Not gRunIdSeeded Then
        Randomize
        gRunIdSeeded = True
    End If
    CreateRunId = Format(Now, "yyyymmdd-hhnnss") & "-" & _
                  CStr(Int((999999 - 100000 + 1) * Rnd + 100000))
End Function

' ----------------------------------------------------------------------------------------------
' [f] WaitMs
'
' 功能说明      : 基于 VBA.Timer 的纯 VBA 阻塞等待，配合重试间隔使用
'               : 处理午夜跨越（Timer 在午夜归零）的边界情况，提前退出而非死循环
' 参数          : delayMs - 等待毫秒数（<= 0 时直接返回）
' 返回          : 无
' Purpose       : Best-effort delay for retry intervals without WinAPI dependency
' ----------------------------------------------------------------------------------------------
Private Sub WaitMs(ByVal delayMs As Long)
    If delayMs <= 0 Then Exit Sub

    Dim t0 As Double
    Dim tTarget As Double
    t0 = Timer
    tTarget = t0 + (CDbl(delayMs) / 1000#)

    Do While Timer < tTarget
        DoEvents
        ' Handle midnight Timer reset (Timer wraps back to near 0 at 00:00:00)
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ==============================================================================================
' SECTION 04: ERROR POLICY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] HandleError
'
' 功能说明      : 统一平台错误处理入口：构造标准错误消息，写入日志
'               : UI 策略：生产模式仅记录日志；开发模式可启用 MsgBox（注释切换）
' 参数          : sourceModule  - 出错模块名称
'               : sourceProc    - 出错过程名称
'               : errNumber     - VBA Err.Number
'               : errDescription - VBA Err.Description
'               : context       - 可选附加上下文
' 返回          : 无
' Purpose       : Unified platform error handler (log-only by default; UI policy configurable)
' Contract      : Platform / Side-effecting
' Side Effects  : Writes to log via LogError
' ----------------------------------------------------------------------------------------------
Public Sub HandleError( _
    ByVal sourceModule As String, _
    ByVal sourceProc As String, _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    Optional ByVal context As Variant _
)
    Dim ctx As String
    ctx = BuildErrorMessage(sourceModule, sourceProc, errNumber, errDescription, context)

    LogError sourceModule, sourceProc, "error", ctx

    ' UI policy:
    ' - Dev mode: uncomment MsgBox for modal error feedback
    ' - Prod mode: log-only; direct users to log sheet
    ' MsgBox msg, vbCritical, "Runtime Error"
End Sub

' ----------------------------------------------------------------------------------------------
' [f] BuildErrorMessage
'
' 功能说明      : 将错误信息格式化为标准日志字符串：err=N;err_desc=描述;at=module.proc[;ctx=...]
' 参数          : sourceModule   - 出错模块
'               : sourceProc     - 出错过程
'               : errNumber      - 错误号
'               : errDescription - 错误描述
'               : context        - 附加上下文
' 返回          : String - 格式化错误字符串
' Purpose       : Standardized error message format for logging and debugging
' ----------------------------------------------------------------------------------------------
Private Function BuildErrorMessage( _
    ByVal sourceModule As String, _
    ByVal sourceProc As String, _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    ByVal context As Variant _
) As String
    Dim s As String
    s = "err=" & CStr(errNumber) & ";err_desc=" & Trim$(errDescription) & _
    ";at=" & sourceModule & "." & sourceProc
    Dim ctxStr As String
    ctxStr = core_logging.BuildContextString(context)
    If Len(ctxStr) > 0 Then s = s & ";ctx=" & ctxStr
    BuildErrorMessage = s
End Function

' ==============================================================================================
' SECTION 05: WORKBOOK IO
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] OpenWorkbook
'
' 功能说明      : 打开工作簿，失败时记录日志并返回 Nothing（不向上抛出异常）
' 参数          : filePathOrUrl - 本地路径或 SharePoint URL
'               : openReadOnly  - 是否以只读方式打开，默认 True
' 返回          : Workbook - 打开的工作簿对象，失败时返回 Nothing
' Purpose       : Open workbook with unified error handling; returns Nothing on failure
'               : Note: readOnly:= render in lowercase due to Office type library identifier normalization in VBA IDE. Semantics are correct; no code defect.
' Contract      : Platform / IO
' Side Effects  : May open a Workbook into Application.Workbooks collection
' ----------------------------------------------------------------------------------------------
Public Function OpenWorkbook( _
    ByVal filePathOrUrl As String, _
    Optional ByVal openReadOnly As Boolean = True _
) As Workbook
    On Error GoTo EH
    Set OpenWorkbook = Application.Workbooks.Open(Filename:=filePathOrUrl, readOnly:=openReadOnly)
    Exit Function
EH:
    HandleError THIS_MODULE, "OpenWorkbook", Err.Number, Err.Description, "path=" & filePathOrUrl
    Set OpenWorkbook = Nothing
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetOpenWorkbookByName
'
' 功能说明      : 在已打开的工作簿集合中按名称查找，未找到时返回 Nothing
' 参数          : wbName - 工作簿名称（含扩展名）
' 返回          : Workbook - 找到的工作簿对象，未找到时返回 Nothing
' Purpose       : Get workbook object from Application.Workbooks by name (best-effort)
' Contract      : Platform / Query-only (no IO)
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function GetOpenWorkbookByName(ByVal wbName As String) As Workbook
    On Error Resume Next
    Set GetOpenWorkbookByName = Application.Workbooks(wbName)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsSharePointUrlAccessible
'
' 功能说明      : 通过尝试以只读方式打开再立即关闭，测试 SharePoint URL 是否可达
'               : 用于补充 Core 层的 URL 格式验证（格式合法不等于可访问）
' 参数          : url - 要测试的 SharePoint URL
' 返回          : Boolean - True = 可访问；False = 访问失败
' Purpose       : Minimal SharePoint URL accessibility test (open + close probe)
' Contract      : Platform / IO
' Side Effects  : Briefly opens and closes a remote Workbook
' ----------------------------------------------------------------------------------------------
Public Function IsSharePointUrlAccessible(ByVal url As String) As Boolean
    Dim wb As Workbook
    On Error GoTo EH
    Set wb = Application.Workbooks.Open(Filename:=url, readOnly:=True)
    wb.Close SaveChanges:=False
    Set wb = Nothing                  ' guard: ensures CleanExit wb Is Nothing check is correct
    IsSharePointUrlAccessible = True

CleanExit:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
        Set wb = Nothing
    End If
    Exit Function

EH:
    IsSharePointUrlAccessible = False
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetFilePathFromDialog
'
' 功能说明      : 打开标准文件选择对话框（单选），用户取消时返回 vbNullString
' 参数          : dialogTitle   - 对话框标题
'               : filterDesc    - 文件类型描述
'               : filterPattern - 文件类型过滤模式（如 "*.xlsx"）
' 返回          : String - 选定的文件路径，取消时返回 vbNullString
' Purpose       : Standard single-file picker dialog with cancellation support
' Note          : No On Error guard; FileDialog failure (e.g. ActiveX restricted policy)
'               : propagates as unhandled exception. Call site is assumed to be interactive
'               : and UI-capable. Add guard at call site if non-standard env is possible.
'               : .title renders lowercase due to Office type library identifier normalization
'               : in VBA IDE. Semantics are correct; no code defect.
' Contract      : Platform / IO (UI interaction)
' Side Effects  : Displays file picker dialog to user
' ----------------------------------------------------------------------------------------------
Public Function GetFilePathFromDialog( _
    Optional ByVal dialogTitle As String = "Select file", _
    Optional ByVal filterDesc As String = "All Files", _
    Optional ByVal filterPattern As String = "*.*" _
) As String
    Dim fd As Object                                    ' late-bind: reduces Office reference coupling
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = dialogTitle
        .Filters.Clear
        .Filters.Add filterDesc, filterPattern
        .AllowMultiSelect = False
        If .Show Then
            GetFilePathFromDialog = .SelectedItems(1)
        Else
            GetFilePathFromDialog = vbNullString
        End If
    End With
End Function

' ==============================================================================================
' SECTION 06: LOGGING HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] LogDebug / LogInfo / LogWarn / LogError / LogPerf
'
' 功能说明      : 各级别日志的便捷包装器；自动注入层级常量 PLAT_LAYER
'               : context 接受 Variant（String / Dictionary / 其他），与 plat_logger.WriteLog
'               : 及 core_logging.BuildContextString 签名对齐
' 参数          : moduleName - 调用方模块名称
'               : procName   - 调用方过程名称
'               : message    - 日志消息
'               : context    - 可选附加上下文（Variant：String / Dictionary / 其他）
' 返回          : 无
' Purpose       : Convenience level wrappers routing to WritePlatLog
' Contract      : Platform / Side-effecting
' Side Effects  : Inherits WritePlatLog side effects (writes to log sink; suppressed on failure)
' ----------------------------------------------------------------------------------------------
Public Sub LogDebug(ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WritePlatLog "DEBUG", moduleName, procName, message, context
End Sub

Public Sub LogInfo(ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WritePlatLog "INFO", moduleName, procName, message, context
End Sub

Public Sub LogWarn(ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WritePlatLog "WARN", moduleName, procName, message, context
End Sub

Public Sub LogError(ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WritePlatLog "ERROR", moduleName, procName, message, context
End Sub

Public Sub LogPerf(ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WritePlatLog "PERF", moduleName, procName, message, context
End Sub

' ----------------------------------------------------------------------------------------------
' [s] WritePlatLog
'
' 功能说明      : 日志路由核心：通过 plat_context.GetLogger() 获取 logger 实例，
'               : 以 CallByName 晚绑定调用 WriteLog（避免对 plat_logger 类型的硬引用）
'               : logger 为 Nothing 时回退至 Debug.Print，context 通过
'               : core_logging.BuildContextString 序列化（兜底 On Error Resume Next）
'               : 日志失败不向调用方抛出异常
' 参数          : level      - 日志级别字符串
'               : moduleName - 调用方模块名称
'               : procName   - 调用方过程名称
'               : message    - 日志消息
'               : context    - 附加上下文（Variant：String / Dictionary / 其他）
' 返回          : 无
' Purpose       : Bridge to plat_logger via plat_context.GetLogger (late binding via CallByName)
'               : Fallback to Debug.Print when logger unavailable; logging never raises
' Contract      : Platform / Side-effecting
' Side Effects  : May write to Immediate window (Debug.Print) or via plat_logger sink
' ----------------------------------------------------------------------------------------------
Private Sub WritePlatLog( _
    ByVal level As String, _
    ByVal moduleName As String, _
    ByVal procName As String, _
    ByVal message As String, _
    Optional ByVal context As Variant _
)
    On Error Resume Next

    Dim logger As Object
    Set logger = plat_context.GetLogger()

    If Not logger Is Nothing Then
        CallByName logger, "WriteLog", VbMethod, _
                   level, PLAT_LAYER, moduleName, procName, message, context
    Else
        Dim ctxStr As String
        ctxStr = core_logging.BuildContextString(context)
        Debug.Print "[" & level & "][" & PLAT_LAYER & "][" & moduleName & "." & procName & "] " & _
                    message & IIf(Len(ctxStr) > 0, " | " & ctxStr, vbNullString)
    End If

    On Error GoTo 0
End Sub

' ==============================================================================================
' SECTION 07: PERF INSTRUMENTATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] MeasurePerf
'
' 功能说明      : 性能计时包装器，支持启动（endAndLog = False）和结束并记录（endAndLog = True）两种模式
'               : 启动时调用 core_logging.StartTimer(label)
'               : 结束时调用 core_logging.ElapsedTime + FormatElapsedTime，以 PERF 级别写入日志
'               : ElapsedTime 返回 -1 时（计时器不存在）以 "perf=NA" 记录，不抛出异常
'               : 建议 label 携带 action/RunId/step 信息（如 "action:RunId:step"）以便日志聚合
' 参数          : label      - 计时标签（唯一标识一个计时区段，start/end 必须一致）
'               : procName   - 所在过程名称（用于日志定位）
'               : endAndLog  - False（默认）= 启动计时；True = 结束计时并写入 PERF 日志
' 返回          : 无
' Purpose       : Minimal viable perf start/end wrapper over core_logging timer API
' Contract      : Platform / Side-effecting
' Side Effects  : Calls core_logging.StartTimer (mutates core_logging timer state) on start;
'               : writes PERF log entry via LogPerf on end
' ----------------------------------------------------------------------------------------------
Public Sub MeasurePerf( _
    ByVal label As String, _
    ByVal procName As String, _
    Optional ByVal endAndLog As Boolean = False _
)
    If Not endAndLog Then
        core_logging.StartTimer label
        Exit Sub
    End If

    Dim sec As Double
    sec = core_logging.ElapsedTime(label, True, -1)

    If sec >= 0 Then
        LogPerf THIS_MODULE, procName, "perf=" & core_logging.FormatElapsedTime(sec), "label=" & label
    Else
        LogPerf THIS_MODULE, procName, "perf=NA", "label=" & label
    End If
End Sub

' ==============================================================================================
' SECTION 08: SESSION TEARDOWN
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] SessionTeardown
'
' 功能说明      : 会话结束后对输出工作表执行归档/删除/保留动作
'               : 动作规则由 Setup 表配置区域驱动，不允许硬编码
'               : 永久工作表（含 @ 标识或核心系统表）受安全锁保护，不参与处理
' 参数          : errMsg  - 输出：失败时的错误说明
' 返回          : Boolean - True=全部处理完成；False=至少一张表处理失败，errMsg 已填充
' Purpose       : Post-session worksheet lifecycle management driven by Setup config rules
' Contract      : Platform / Workbook mutation
' Side Effects  : May delete worksheets; may create Archive sheet; writes to Log sheet via Logger
' ----------------------------------------------------------------------------------------------
Public Function SessionTeardown(ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    SessionTeardown = False

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws          As Worksheet
    Dim wsName      As String
    Dim action      As String
    Dim failCount   As Long
    Dim totalCount  As Long
    Dim localErr    As String
    Dim loopCount   As Long

    Const MAX_SHEETS As Long = 200

    LogInfo PLAT_LAYER, THIS_MODULE, "SessionTeardown", "Session teardown started"

    For Each ws In wb.Worksheets
        loopCount = loopCount + 1
        If loopCount > MAX_SHEETS Then
            LogWarn PLAT_LAYER, THIS_MODULE, "SessionTeardown", _
                "sheet_count_exceeded_limit=" & MAX_SHEETS & ";teardown_truncated=true"
            Exit For
        End If

        wsName = ws.Name

        ' --- Permanent sheet guard: never touch these ---
        If p_IsPermanentSheet(wsName) Then
            LogInfo PLAT_LAYER, THIS_MODULE, "SessionTeardown", _
                "sheet=" & wsName & ";verdict=permanent;skipped=true"
            GoTo NextSheet
        End If

        ' --- Resolve action from config table ---
        action = p_ResolveSheetAction(wsName)

        ' --- Execute action ---
        totalCount = totalCount + 1
        localErr = vbNullString
        If Not p_ExecuteSheetAction(wsName, action, localErr) Then
            failCount = failCount + 1
            If Len(errMsg) = 0 Then errMsg = localErr
            LogError PLAT_LAYER, THIS_MODULE, "SessionTeardown", _
                "sheet=" & wsName & ";action=" & action & ";err=" & localErr
        End If

NextSheet:
    Next ws

    If failCount > 0 Then
        errMsg = THIS_MODULE & ".SessionTeardown: " & failCount & " of " & totalCount & _
                 " sheets failed; first_err=" & errMsg
        LogError PLAT_LAYER, THIS_MODULE, "SessionTeardown", errMsg
        Exit Function
    End If

    LogInfo PLAT_LAYER, THIS_MODULE, "SessionTeardown", _
        "teardown_complete;total_processed=" & totalCount
    SessionTeardown = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] PurgeTempSheets
'
' 功能说明      : 删除所有以 TEMP_ 前缀命名的临时工作表
'               : 永久工作表安全锁确保核心表不受影响
' 参数          : errMsg       - 输出：失败时的错误说明
' 返回          : Boolean - True=清理完成；False=至少一张表删除失败
' Purpose       : Safe removal of all TEMP_-prefixed transient worksheets
' Contract      : Platform / Workbook mutation
' Side Effects  : May delete worksheets matching TEMP_ prefix
' ----------------------------------------------------------------------------------------------
Public Function PurgeTempSheets(ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    PurgeTempSheets = False

    Const TEMP_PREFIX As String = "TEMP_"
    Const MAX_SHEETS  As Long = 200

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim toDelete()  As String
    Dim count       As Long
    Dim i           As Long
    Dim loopCount   As Long

    Set wb = ThisWorkbook
    count = 0
    ReDim toDelete(0 To MAX_SHEETS - 1)

    ' --- Collect candidates in first pass (avoid mutating collection mid-loop) ---
    For Each ws In wb.Worksheets
        loopCount = loopCount + 1
        If loopCount > MAX_SHEETS Then Exit For
        If InStr(1, ws.Name, TEMP_PREFIX, vbTextCompare) > 0 Then
            If Not p_IsPermanentSheet(ws.Name) Then
                toDelete(count) = ws.Name
                count = count + 1
            End If
        End If
    Next ws

    If count = 0 Then
        LogInfo PLAT_LAYER, THIS_MODULE, "PurgeTempSheets", "no_temp_sheets_found"
        PurgeTempSheets = True
        Exit Function
    End If

    ' --- Delete in second pass ---
    Dim failCount As Long
    failCount = 0

    For i = 0 To count - 1
        Dim localErr As String
        localErr = vbNullString
        If Not p_ExecuteSheetAction(toDelete(i), "Delete", localErr) Then
            failCount = failCount + 1
            If Len(errMsg) = 0 Then errMsg = localErr
        End If
    Next i

    If failCount > 0 Then
        errMsg = THIS_MODULE & ".PurgeTempSheets: " & failCount & " of " & count & _
                 " deletions failed; first_err=" & errMsg
        LogError PLAT_LAYER, THIS_MODULE, "PurgeTempSheets", errMsg
        Exit Function
    End If

    LogInfo PLAT_LAYER, THIS_MODULE, "PurgeTempSheets", _
        "purge_complete;deleted=" & count
    PurgeTempSheets = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] RefreshOutputSheetList
'
' 功能说明      : 扫描当前 Workbook 所有非永久工作表，将名称写入 Setup 表的锚点区域
'               : 用于在 Python 运行后刷新 UI 列表，供用户查看输出表清单
' 参数          : errMsg       - 输出：失败时的错误说明
' 返回          : Boolean - True=刷新成功；False=失败，errMsg 已填充
' Purpose       : Refresh output sheet inventory written to Setup anchor range
' Contract      : Platform / Workbook read + Setup sheet write
' Side Effects  : Writes sheet names to Setup worksheet anchor range
' ----------------------------------------------------------------------------------------------
Public Function RefreshOutputSheetList(ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    RefreshOutputSheetList = False

    Const MAX_ROWS    As Long = 200
    Const MAX_SHEETS  As Long = 200

    Dim anchorRange   As Range
    Dim ws            As Worksheet
    Dim rowIdx        As Long
    Dim loopCount     As Long

    ' --- Resolve anchor range from context ---
    On Error Resume Next
    Set anchorRange = plat_context.GetRange("rng_out_ws_list_anchor", errMsg)
    On Error GoTo 0

    If anchorRange Is Nothing Then
        errMsg = THIS_MODULE & ".RefreshOutputSheetList: anchor range not found [rng_out_ws_list_anchor]"
        LogError PLAT_LAYER, THIS_MODULE, "RefreshOutputSheetList", errMsg
        Exit Function
    End If

    ' --- Clear existing content ---
    On Error Resume Next
    anchorRange.Resize(MAX_ROWS, 1).ClearContents
    On Error GoTo 0

    rowIdx = 0

    For Each ws In ThisWorkbook.Worksheets
        loopCount = loopCount + 1
        If loopCount > MAX_SHEETS Then Exit For

        If Not p_IsPermanentSheet(ws.Name) Then
            If rowIdx >= MAX_ROWS Then
                LogWarn PLAT_LAYER, THIS_MODULE, "RefreshOutputSheetList", _
                    "max_rows_reached=" & MAX_ROWS & ";list_truncated=true"
                Exit For
            End If
            anchorRange.Offset(rowIdx, 0).Value = ws.Name
            rowIdx = rowIdx + 1
        End If
    Next ws

    LogInfo PLAT_LAYER, THIS_MODULE, "RefreshOutputSheetList", _
        "refresh_complete;listed=" & rowIdx
    RefreshOutputSheetList = True
End Function

' ----------------------------------------------------------------------------------------------
' [f] p_ResolveSheetAction
'
' 功能说明      : 从 Setup 表配置区域读取工作表名称对应的处理动作
'               : 若未找到匹配规则，默认返回 "Keep"（安全保守策略）
'               : 配置区域：Named Range "rng_sys_ws_action_rules"（两列：SheetPattern, Action）
' 参数          : sheetName - 工作表名称
' 返回          : String - "ArchiveDelete" / "Delete" / "Keep"（默认 "Keep"）
' Purpose       : Config-driven sheet action resolver; extensible for v3 output types
' ----------------------------------------------------------------------------------------------
Private Function p_ResolveSheetAction(ByVal sheetName As String) As String
    p_ResolveSheetAction = "Keep"   ' safe default

    Dim rulesRange  As Range
    Dim rules       As Variant
    Dim i           As Long
    Dim pattern     As String
    Dim action      As String
    Dim ignored     As String

    ' --- Load rules table from named range ---
    On Error Resume Next
    Set rulesRange = plat_context.GetRange("rng_sys_ws_action_rules", ignored)
    On Error GoTo 0

    If rulesRange Is Nothing Then
        ' No rules table configured: default Keep for all non-permanent sheets
        LogWarn PLAT_LAYER, THIS_MODULE, "p_ResolveSheetAction", _
            "rules_range_missing=rng_sys_ws_action_rules;default=Keep;sheet=" & sheetName
        Exit Function
    End If

    rules = rulesRange.Value

    If Not IsArray(rules) Then Exit Function

    ' --- Scan rules: first match wins ---
    For i = LBound(rules, 1) To UBound(rules, 1)
        pattern = Trim$(CStr(rules(i, 1)))
        action  = Trim$(CStr(rules(i, 2)))

        If Len(pattern) = 0 Then GoTo NextRule

        If InStr(1, sheetName, pattern, vbTextCompare) > 0 Then
            Select Case UCase$(action)
                Case "ARCHIVEDELETE", "ARCHIVE_DELETE", "APPEND_DELETE"
                    p_ResolveSheetAction = "ArchiveDelete"
                Case "DELETE"
                    p_ResolveSheetAction = "Delete"
                Case Else
                    p_ResolveSheetAction = "Keep"
            End Select
            Exit Function
        End If

NextRule:
    Next i
End Function

' ----------------------------------------------------------------------------------------------
' [f] p_ExecuteSheetAction
'
' 功能说明      : 对指定工作表执行单项动作（归档删除 / 直接删除 / 保留）
'               : 归档动作：将数据追加写入 Archive 表后删除源表
'               : 删除动作：关闭 DisplayAlerts 后直接删除
'               : 保留动作：仅记录日志，不做任何修改
' 参数          : sheetName - 工作表名称
'               : action    - "ArchiveDelete" / "Delete" / "Keep"
'               : errMsg    - 输出：失败时的错误说明
' 返回          : Boolean - True=执行成功；False=执行失败，errMsg 已填充
' ----------------------------------------------------------------------------------------------
Private Function p_ExecuteSheetAction(ByVal sheetName As String, _
                                      ByVal action     As String, _
                                      ByRef errMsg     As String) As Boolean
    errMsg = vbNullString
    p_ExecuteSheetAction = False

    Dim ws          As Worksheet
    Dim archiveWs   As Worksheet
    Dim archiveErr  As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        errMsg = THIS_MODULE & ".p_ExecuteSheetAction: sheet not found [" & sheetName & "]"
        LogWarn PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", errMsg
        Exit Function
    End If

    Select Case UCase$(action)

        Case "ARCHIVEDELETE"
            ' --- Step 1: Ensure Archive sheet exists ---
            archiveErr = vbNullString
            Set archiveWs = p_EnsureArchiveSheet(archiveErr)
            If archiveWs Is Nothing Then
                errMsg = THIS_MODULE & ".p_ExecuteSheetAction: cannot get archive sheet;" & archiveErr
                Exit Function
            End If

            ' --- Step 2: Append source data to Archive ---
            Dim srcRange    As Range
            Dim destCell    As Range
            Dim lastArchRow As Long

            On Error Resume Next
            Set srcRange = ws.UsedRange
            On Error GoTo 0

            If srcRange Is Nothing Then
                LogWarn PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", _
                    "sheet=" & sheetName & ";used_range_empty;archive_skipped"
            Else
                On Error Resume Next
                lastArchRow = archiveWs.Cells(archiveWs.Rows.Count, 1).End(xlUp).Row
                If lastArchRow < 1 Then lastArchRow = 1
                Set destCell = archiveWs.Cells(lastArchRow + 1, 1)
                srcRange.Copy destCell
                On Error GoTo 0
            End If

            LogInfo PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", _
                "sheet=" & sheetName & ";action=archive_copy_done"

            ' --- Step 3: Delete source sheet ---
            On Error Resume Next
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Dim delErr As String
            delErr = Err.Description
            On Error GoTo 0

            If Len(delErr) > 0 Then
                errMsg = THIS_MODULE & ".p_ExecuteSheetAction: delete_after_archive_failed;" & _
                         "sheet=" & sheetName & ";err=" & delErr
                Exit Function
            End If

            LogInfo PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", _
                "sheet=" & sheetName & ";action=archive_delete_done"

        Case "DELETE"
            On Error Resume Next
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Dim delErr2 As String
            delErr2 = Err.Description
            On Error GoTo 0

            If Len(delErr2) > 0 Then
                errMsg = THIS_MODULE & ".p_ExecuteSheetAction: direct_delete_failed;" & _
                         "sheet=" & sheetName & ";err=" & delErr2
                Exit Function
            End If

            LogInfo PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", _
                "sheet=" & sheetName & ";action=delete_done"

        Case Else
            ' Keep - no mutation, log only
            LogInfo PLAT_LAYER, THIS_MODULE, "p_ExecuteSheetAction", _
                "sheet=" & sheetName & ";action=keep"

    End Select

    p_ExecuteSheetAction = True
End Function

' ----------------------------------------------------------------------------------------------
' [f] p_IsPermanentSheet
'
' 功能说明      : 判定工作表是否为永久保护表（永远不应被删除）
'               : 判定条件：名称含 @ 标识，或为系统核心表（Setup / Main / Log / Archive）
' 参数          : sheetName - 工作表名称
' 返回          : Boolean - True=永久表，不可删除；False=非永久表
' Purpose       : Centralized permanent sheet guard; single source of truth for all callers
' ----------------------------------------------------------------------------------------------
Private Function p_IsPermanentSheet(ByVal sheetName As String) As Boolean
    If InStr(1, sheetName, "@", vbTextCompare) > 0 Then
        p_IsPermanentSheet = True
        Exit Function
    End If

    Select Case LCase$(Trim$(sheetName))
        Case "setup", "main", "log", "archive", "treaty", "sublob"
            p_IsPermanentSheet = True
        Case Else
            p_IsPermanentSheet = False
    End Select
End Function

' ----------------------------------------------------------------------------------------------
' [f] p_EnsureArchiveSheet
'
' 功能说明      : 返回 Archive 工作表对象；若不存在则自动创建并追加至末尾
' 参数          : errMsg - 输出：失败时的错误说明
' 返回          : Worksheet - Archive 工作表对象；失败时返回 Nothing
' Purpose       : Guarantee Archive sheet availability before archive-delete operations
' ----------------------------------------------------------------------------------------------
Private Function p_EnsureArchiveSheet(ByRef errMsg As String) As Worksheet
    errMsg = vbNullString

    Const ARCHIVE_NAME As String = "Archive"
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ARCHIVE_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set p_EnsureArchiveSheet = ws
        Exit Function
    End If

    ' --- Create Archive sheet ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    If Not ws Is Nothing Then ws.Name = ARCHIVE_NAME
    On Error GoTo 0

    If ws Is Nothing Then
        errMsg = THIS_MODULE & ".p_EnsureArchiveSheet: failed to create Archive sheet"
        LogError PLAT_LAYER, THIS_MODULE, "p_EnsureArchiveSheet", errMsg
        Exit Function
    End If

    LogInfo PLAT_LAYER, THIS_MODULE, "p_EnsureArchiveSheet", "archive_sheet_created"
    Set p_EnsureArchiveSheet = ws
End Function