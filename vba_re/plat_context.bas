Attribute VB_Name = "plat_context"
' ==============================================================================================
' MODULE NAME       : plat_context
' LAYER             : platform
' PURPOSE           : Provides centralized Workbook runtime context,
'                     cached worksheet accessors, and shared object lifecycle control.
' DEPENDS           : core_utils, core_logging, Excel Object Model
' NOTE              : - Manages runtime state and shared resources.
'                     - All worksheet access must go through context getters.
'                     - No business logic allowed.
' STATUS            : Frozen
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Introduced centralized workbook context module to replace scattered worksheet and named-range access.
'   - Init (Config Access): Implemented configuration loading from named ranges with default fallback semantics.
'   - Init (Range Binding): Added basic named range registration and lookup helpers to reduce direct Range(...) usage in business logic.
'   - Init (Logging Hook): Provided initial logger bootstrap entry point (pre-layered design, script-style integration).

' v2.0.0
'   - Refactor (Architecture): Adopted layered model (Core / Platform / Business); positioned plat_context as Platform lifecycle boundary.
'   - Refactor (Lifecycle): Formalized InitContext / ResetContext contract.
'   - Refactor (Encapsulation): Restricted configuration registry scope (GetConfigDictionary kept Private); enforced access via GetConfigValue only.
'   - Fix (Sheet Binding): Added explicit worksheet binding validation with fail-fast error reporting for missing tabs.
'   - Align (Layer Discipline): Explicitly qualified Core utility calls to preserve dependency direction (Core → Platform).
'   - Improve (Naming Clarity): Separated configuration keys (CFG_*) from named range identifiers (NR_*).

' v2.1.0
'   - Fix (Init State Machine): Introduced Init-In-Progress guard to eliminate half-initialized window and getter re-entry risks.
'   - Fix (Logger Re-entry): Removed implicit InitContext call from InitProjectLogger; now asserts prerequisites (Fail Fast).
'   - Fix (Config Failure Semantics): GetConfigDictionary now Fail-Fast by default; optional silent fallback only for non-critical reads.
'   - Fix (Config Access Timing): GetConfigValue now enforces Context Ready before load.
'   - Fix (Range Read): Fixed ReadRangeFromSheet missing Exit Function bug.
'   - Clarify (Registry Entry): Declared setup-driven registry as primary entry; hardcoded registry remains as fallback (Deprecated).

' v2.1.1
'   - Fix (Fail Msg Capture): Captures Err.Description before ResetContext clears Err.
'   - Fix (Init Logger Config Path): Added internal GetConfigValueDirect to allow logger init during init phase.
'   - Fix (Numeric Convert): Avoided Banker's rounding by enforcing strict integer semantics for vbLong/vbInteger reads.
'   - Fix (Range Cache): Replaced Dictionary.Add with overwrite assignment to avoid nested-call key collisions.
'   - Fix (Deprecated Mapping): Restored legacy NR worksheet pairing inside deprecated RegisterRanges.
'   - Fix (EH Msg Capture): Captured Err.Description at the entry of GetConfigDictionary error handler to prevent message loss after cleanup operations.
'   - Fix (Registry Shape Guard): Added validation for rng_sys_range_registry return shape (must be 2D with required columns) to avoid out-of-range failures on single-row/empty registries.
'   - Fix (Compile): Removed duplicate wsLog/WsLog property definition (VBA identifiers are case-insensitive); standardized on WsLog.
'   - Refine (Encapsulation): Made InitProjectLogger Private to enforce single lifecycle control via InitContext and prevent external factory misuse.
'   - Improve (Registry Validation): Added explicit row-level validation for missing SheetName/RangeName and switched registry read to Value2 for stable typing.

' v2.1.2
'   - Refine (Factory Semantics): Removed redundant gLogger assignment inside InitProjectLogger; state ownership consolidated to InitContext caller.
'   - Clarify (Contract): Documented GetRangeValue caller-owned context responsibility; EnsureContextReady intentionally omitted as function depends only on caller-supplied ws.
'   - Refine (Clarity): Replaced direct function return-slot read in GetRangeOrRaise with explicit local variable to eliminate return-slot-as-intermediate ambiguity.
'   - Fix (Defensive Consistency): Added null guard to Config Property to match Logger Property pattern; both raise on Nothing after successful Init.
'   - Refine (Encapsulation): Demoted RegisterRanges from Public to Private; legacy fallback retained internally but hidden from external callers to prevent misuse of deprecated API.
'   - Note (Encoding): Bilingual comments (Chinese/English) retained for development readability; full migration to English pending before production release to resolve GBK/UTF-8 mismatch.
'   - Fix (Config Default): Changed CFG_STR_LOG_LEVEL default from "MODULE" to "INFO" to eliminate undocumented cross-module sentinel coupling with plat_logger.LevelRank.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 00: MODULE STATE
'
' SECTION 01: INIT / RESET (LIFECYCLE)
'   [S] InitContext                             - Initialize platform context, bind worksheets and load config
'   [S] ResetContext                            - Reset context, release all resources
'   [f] BindSheetOrRaise                        - Bind worksheet or raise error
'   [s] EnsureContextReady                      - Ensure context is ready
'
' SECTION 02: WORKSHEET GETTERS
'   [P] ContextInited                           - Return context initialization status
'   [P] WsSetup/WsMain/WsLog/WsTreaty/WsSubLoB  - Return Setup/Main/Log/Treaty/SubLoB worksheet object
'   [P] WsGN/WsEL/WsRE/WsKPI                    - Return GN/EL/RE/KPI worksheet object
'
' SECTION 03: LOGGER ACCESS
'   [P] Logger                                  - Return logger object
'   [F] GetLogger                               - Get logger object (function form)
'   [f] InitProjectLogger                       - Initialize project logger
'
' SECTION 04: CONFIG REGISTRY
'   [f] GetConfigDictionary                     - Get configuration dictionary
'   [F] GetConfigValue                          - Get configuration value
'   [f] GetConfigValueDirect                    - Direct config value access (internal)
'   [P] Config                                  - Return configuration dictionary object
'
' SECTION 05: RANGE REGISTRY (CACHED OBJECT ACCESS)
'   [f] ReadRangeFromSheet                      - Read range value from worksheet
'   [F] GetRange                                - Get range object (cached)
'   [F] GetRangeOrRaise                         - Get range object or raise error
'   [F] GetRangeValue                           - Get range value
'   [S] ClearRangeCache                         - Clear range cache
'
' SECTION 06: RANGE REGISTRY WARMUP (PRIMARY + FALLBACK)
'   [S] RegisterRangesFromSetup                 - Register ranges from Setup worksheet
'   [s] RegisterRanges                          - Register ranges (legacy method)
'
' SECTION 07: INTERNAL CONFIG REGISTRATION HELPERS
'   [s] RegisterConfig                          - Register configuration item
'   [f] ReadValueFromSheet                      - Read value from worksheet
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function, [T]=Type
'       Rule: Helper functions and private procedures inherit the Contract and
'             Side Effects of their parent public API unless explicitly stated otherwise.
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 00: MODULE STATE
' ==============================================================================================

Private gWsSetup As Worksheet
Private gWsMain As Worksheet
Private gWsLog As Worksheet
Private gWsTreaty As Worksheet
Private gWsSubLoB As Worksheet
Private gWsGN As Worksheet
Private gWsEL As Worksheet
Private gWsRE As Worksheet
Private gWsKPI As Worksheet

Private gConfig As Object          ' Scripting.Dictionary
Private gLogger As Object          ' plat_logger (object)
Private gRangeCache As Object      ' Scripting.Dictionary (String -> Range)

Private gContextInited As Boolean
Private gContextInitInProgress As Boolean

Private Const CFG_IS_SYS_MESSAGE_ENABLED  As String = "is_sys_message_enabled"
Private Const CFG_IS_SYS_ERROR_ENABLED    As String = "is_sys_error_enabled"
Private Const CFG_IS_LOG_WRITE_ENABLED    As String = "is_log_write_enabled"
Private Const CFG_STR_LOG_LEVEL           As String = "str_log_level"
Private Const CFG_STR_LOG_SHEET_NAME      As String = "str_log_sheet_name"
Private Const CFG_STR_LOG_ANCHOR_CELL     As String = "str_log_anchor_cell"
Private Const CFG_STR_10K_FILE_PATH       As String = "str_10k_file_path"
Private Const CFG_STR_10K_SEGMENT         As String = "str_10k_segment"
Private Const CFG_STR_10K_LOSS_LIST       As String = "str_10k_loss_list"
Private Const CFG_INT_10K_START_ROW       As String = "int_10k_start_row"
Private Const CFG_LNG_SYS_SIM_YEARS       As String = "lng_sys_sim_years"
Private Const CFG_LNG_SYS_CHUNK_THRESHOLD As String = "lng_sys_chunk_threshold"

Private Const NR_RNG_SYS_RANGE_REGISTRY   As String = "rng_sys_range_registry"
Private Const NR_RNG_SYS_MAIN_CONFIG      As String = "rng_sys_main_config"
Private Const NR_RNG_SYS_SEGMENT          As String = "rng_sys_segment"
Private Const NR_RNG_SYS_SUBLOB_TO_LOB    As String = "rng_sys_sublob_to_lob"
Private Const NR_RNG_OUT_WS_LIST_ANCHOR   As String = "rng_out_ws_list_anchor"
Private Const NR_RNG_TREATY_REFERENCE     As String = "rng_treaty_reference"
Private Const NR_RNG_MAP_SUBLOB_TO_TREATY As String = "rng_map_sublob_to_treaty"

' ==============================================================================================
' SECTION 01: INIT / RESET (LIFECYCLE)
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] InitContext
'
' 功能说明      : 初始化平台上下文，绑定工作表并加载配置
' 参数          : forceReload - 可选，是否强制重新加载，默认为False
' 返回          : 无
' Purpose       : Initializes the platform context, binds worksheets and loads configuration
' Contract      : Platform / Lifecycle
' Side Effects  : Modifies module state (gWs*, gConfig, gLogger, gContextInited)
' ----------------------------------------------------------------------------------------------
Public Sub InitContext(Optional ByVal forceReload As Boolean = False)
    If gContextInited And Not forceReload Then Exit Sub

    If gContextInitInProgress Then
        Err.Raise vbObjectError + 1200, "plat_context.InitContext", _
                  "InitContext re-entry is not allowed (initialization already in progress)."
    End If
    On Error GoTo Fail
    gContextInitInProgress = True

    If forceReload Then ResetContext

    ' ---- Bind worksheets (must not use getters here) ----
    Set gWsSetup = BindSheetOrRaise("Setup@SYS")
    Set gWsMain = BindSheetOrRaise("Main@SYS")
    Set gWsLog = BindSheetOrRaise("Log@SYS")
    Set gWsTreaty = BindSheetOrRaise("Treaty@REF")
    Set gWsSubLoB = BindSheetOrRaise("SubLoB@REF")
    Set gWsGN = BindSheetOrRaise("GN@OUT")
    Set gWsEL = BindSheetOrRaise("EL@OUT")
    Set gWsRE = BindSheetOrRaise("RE@OUT")
    Set gWsKPI = BindSheetOrRaise("KPI@ANL")

    ' ---- Config (Fail Fast in Init) ----
    Set gConfig = GetConfigDictionary(True, False)
    If gConfig Is Nothing Then
        Err.Raise vbObjectError + 1204, "plat_context.InitContext", _
                  "Config initialization returned Nothing (unexpected)."
    End If

    ' ---- Logger (must be allowed during init; reads config via direct accessor) ----
    Set gLogger = InitProjectLogger()
    If gLogger Is Nothing Then
        Err.Raise vbObjectError + 1205, "plat_context.InitContext", _
                  "Logger initialization returned Nothing (unexpected)."
    End If

    gContextInited = True
    gContextInitInProgress = False
    Exit Sub

Fail:
    Dim failMsg As String
    failMsg = Err.Description ' capture BEFORE ResetContext clears Err via On Error Resume Next

    ResetContext
    gContextInitInProgress = False

    Err.Raise vbObjectError + 1201, "plat_context.InitContext", _
              "InitContext failed. Original error: " & failMsg
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ResetContext
'
' 功能说明      : 重置上下文，释放所有资源
' 参数          : 无
' 返回          : 无
' Purpose       : Resets context, releases all resources
' Contract      : Platform / Lifecycle
' Side Effects  : Clears all module state (gWs*, gConfig, gLogger, gRangeCache)
' ----------------------------------------------------------------------------------------------
Public Sub ResetContext()
    On Error Resume Next

    Set gWsSetup = Nothing
    Set gWsMain = Nothing
    Set gWsLog = Nothing
    Set gWsTreaty = Nothing
    Set gWsSubLoB = Nothing
    Set gWsGN = Nothing
    Set gWsEL = Nothing
    Set gWsRE = Nothing
    Set gWsKPI = Nothing

    Set gConfig = Nothing
    Set gLogger = Nothing
    Set gRangeCache = Nothing

    gContextInited = False
    gContextInitInProgress = False

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------------------------
' [f] BindSheetOrRaise
'
' 功能说明      : 绑定工作表，如果找不到则引发错误
' 参数          : sheetName - 工作表名称
' 返回          : Worksheet - 绑定后的工作表对象
' Purpose       : Binds a worksheet by name or raises error if not found
' ----------------------------------------------------------------------------------------------
Private Function BindSheetOrRaise(ByVal sheetName As String) As Worksheet
    On Error GoTo EH
    Set BindSheetOrRaise = ThisWorkbook.Worksheets(sheetName)
    Exit Function
EH:
    Err.Raise vbObjectError + 1202, "plat_context.BindSheetOrRaise", _
              "Worksheet not found: " & sheetName
End Function

' ----------------------------------------------------------------------------------------------
' [s] EnsureContextReady
'
' 功能说明      : 确保上下文已准备就绪，如果未初始化则自动初始化
' 参数          : 无
' 返回          : 无
' Purpose       : Ensures context is ready, automatically initializes if not already initialized
' ----------------------------------------------------------------------------------------------
Private Sub EnsureContextReady()
    If gContextInitInProgress Then
        Err.Raise vbObjectError + 1203, "plat_context.EnsureContextReady", _
                  "Context is initializing; accessor re-entry is not allowed."
    End If

    If Not gContextInited Then
        InitContext False
    End If
End Sub

' ==============================================================================================
' SECTION 02: WORKSHEET GETTERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [P] ContextInited
'
' 功能说明      : 返回上下文初始化状态
' 参数          : 无
' 返回          : Boolean - 上下文是否已初始化
' Purpose       : Returns context initialization status
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Property Get ContextInited() As Boolean
    ContextInited = gContextInited
End Property

' ----------------------------------------------------------------------------------------------
' [P] WsSetup/WsMain/WsLog/WsTreaty/WsSubLoB
'
' 功能说明      : 返回Setup/Main/Log/Treaty/SubLoB等系统类和引用类工作表对象
' 参数          : 无
' 返回          : Worksheet - Setup/Main/Log/Treaty/SubLoB工作表对象
' Purpose       : Returns Setup/Main/Log/Treaty/SubLoB worksheet object
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Property Get WsSetup() As Worksheet
    EnsureContextReady
    Set WsSetup = gWsSetup
End Property

Public Property Get WsMain() As Worksheet
    EnsureContextReady
    Set WsMain = gWsMain
End Property

Public Property Get WsLog() As Worksheet
    EnsureContextReady
    Set WsLog = gWsLog
End Property

Public Property Get WsTreaty() As Worksheet
    EnsureContextReady
    Set WsTreaty = gWsTreaty
End Property

Public Property Get WsSubLoB() As Worksheet
    EnsureContextReady
    Set WsSubLoB = gWsSubLoB
End Property

' ----------------------------------------------------------------------------------------------
' [P] WsGN/WsEL/WsRE/WsKPI
'
' 功能说明      : 返回GN/EL/RE/KPI等输出类工作表对象
' 参数          : 无
' 返回          : Worksheet - GN/EL/RE/KPI工作表对象
' Purpose       : Returns GN/EL/RE/KPI worksheet object
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Property Get WsGN() As Worksheet
    EnsureContextReady
    Set WsGN = gWsGN
End Property

Public Property Get WsEL() As Worksheet
    EnsureContextReady
    Set WsEL = gWsEL
End Property

Public Property Get WsRE() As Worksheet
    EnsureContextReady
    Set WsRE = gWsRE
End Property

Public Property Get WsKPI() As Worksheet
    EnsureContextReady
    Set WsKPI = gWsKPI
End Property

' ==============================================================================================
' SECTION 03: LOGGER ACCESS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [P] Logger
'
' 功能说明      : 返回日志记录器对象
' 参数          : 无
' 返回          : Object - 日志记录器对象
' Purpose       : Returns logger object
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Property Get Logger() As Object
    EnsureContextReady
    If gLogger Is Nothing Then
        Err.Raise vbObjectError + 1300, "plat_context.Logger", _
                  "Logger not initialized (unexpected)."
    End If
    Set Logger = gLogger
End Property

' ----------------------------------------------------------------------------------------------
' [F] GetLogger
'
' 功能说明      : 获取日志记录器对象（函数形式）
' 参数          : 无
' 返回          : Object - 日志记录器对象
' Purpose       : Get logger object (function form)
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function GetLogger() As Object
    Set GetLogger = Logger
End Function

' ----------------------------------------------------------------------------------------------
' [f] InitProjectLogger
'
' 功能说明      : 初始化项目日志记录器
' 参数          : 无
' 返回          : Object - 初始化后的日志记录器对象
' Purpose       : Initializes project logger
' Contract      : Platform / Internal
' Side Effects  : None (returns new logger instance; state ownership consolidated to InitContext)
' ----------------------------------------------------------------------------------------------
Private Function InitProjectLogger() As Object
    ' DO NOT call InitContext here (no hidden lifecycle decisions).
    ' Prerequisites must be satisfied by InitContext ordering.

    If gWsLog Is Nothing Then
        Err.Raise vbObjectError + 1310, "plat_context.InitProjectLogger", _
                  "Logger init requires worksheets already bound (gWsLog is Nothing)."
    End If

    If gConfig Is Nothing Then
        Err.Raise vbObjectError + 1311, "plat_context.InitProjectLogger", _
                  "Logger init requires config already loaded (gConfig is Nothing)."
    End If

    ' ---- config (DIRECT) ----
    ' NOTE: Must NOT call public GetConfigValue here because InitContext is still in progress.
    Dim enableSheet As Boolean
    Dim enableImmediate As Boolean  ' always enabled: Immediate window is dev-only; no runtime side effect
    Dim minLevel As String
    Dim logSheetName As String
    Dim anchorCell As String

    enableSheet = core_utils.ToSafeBoolean(GetConfigValueDirect(CFG_IS_LOG_WRITE_ENABLED, True), True)
    enableImmediate = True
    minLevel = CStr(GetConfigValueDirect(CFG_STR_LOG_LEVEL, "INFO"))
    logSheetName = CStr(GetConfigValueDirect(CFG_STR_LOG_SHEET_NAME, "Log@SYS"))
    anchorCell = CStr(GetConfigValueDirect(CFG_STR_LOG_ANCHOR_CELL, "B3"))

    ' ---- resolve ws ----
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(logSheetName)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = gWsLog

    ' ---- build logger ----
    Dim lgr As plat_logger
    Set lgr = New plat_logger
    Call lgr.Init(enableSheet, enableImmediate, minLevel, ws, anchorCell)

    Set InitProjectLogger = lgr
End Function

' ==============================================================================================
' SECTION 04: CONFIG REGISTRY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] GetConfigDictionary
'
' 功能说明      : 获取配置字典，支持强制重新加载和静默回退
' 参数          : forceReload - 可选，是否强制重新加载，默认为False
'               : allowSilentFallback - 可选，是否允许静默回退，默认为False
' 返回          : Object - 配置字典对象
' Purpose       : Gets configuration dictionary with force reload and silent fallback options
' Contract      : Platform / Internal
' Side Effects  : May modify gConfig state
' ----------------------------------------------------------------------------------------------
Private Function GetConfigDictionary( _
    Optional ByVal forceReload As Boolean = False, _
    Optional ByVal allowSilentFallback As Boolean = False _
) As Object

    If (gConfig Is Nothing) Or forceReload Then
        On Error GoTo EH

        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")

        RegisterConfig d, CFG_IS_SYS_MESSAGE_ENABLED, gWsMain, True
        RegisterConfig d, CFG_IS_SYS_ERROR_ENABLED, gWsMain, True
        RegisterConfig d, CFG_IS_LOG_WRITE_ENABLED, gWsMain, True

        RegisterConfig d, CFG_STR_LOG_LEVEL, gWsLog, "INFO"
        RegisterConfig d, CFG_STR_LOG_SHEET_NAME, gWsLog, "Log@SYS"
        RegisterConfig d, CFG_STR_LOG_ANCHOR_CELL, gWsLog, "B3"

        RegisterConfig d, CFG_STR_10K_FILE_PATH, gWsMain, "C:\Users\Richard\Desktop\VBA\Data_10k_Sims_2025_v0.5_MM_Large_RDS.xlsb"
        RegisterConfig d, CFG_STR_10K_SEGMENT, gWsMain, "RDS"
        RegisterConfig d, CFG_STR_10K_LOSS_LIST, gWsMain, "RDS_Gross"
        RegisterConfig d, CFG_INT_10K_START_ROW, gWsMain, 4

        RegisterConfig d, CFG_LNG_SYS_SIM_YEARS, gWsMain, 10000
        RegisterConfig d, CFG_LNG_SYS_CHUNK_THRESHOLD, gWsMain, 6000000

        Set gConfig = d
    End If

    Set GetConfigDictionary = gConfig
    Exit Function

EH:
    Dim ehMsg As String
    ehMsg = Err.Description
    
    Set gConfig = Nothing
    If allowSilentFallback Then
        Set GetConfigDictionary = Nothing
        Exit Function
    End If

    Err.Raise vbObjectError + 1400, "plat_context.GetConfigDictionary", _
              "Config registry build failed. Original error: " & ehMsg
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetConfigValue
'
' 功能说明      : 获取配置值，如果键不存在则返回默认值
'               : Public - external callers must only access config when Context is READY
' 参数          : key - 配置键
'               : defaultValue - 可选，默认值
' 返回          : Variant - 配置值或默认值
' Purpose       : Gets configuration value, returns default if key not found
' Contract      : Platform / Query-only
' Side Effects  : May modify gConfig state (silent fallback on None)
' ----------------------------------------------------------------------------------------------
Public Function GetConfigValue(ByVal key As String, Optional ByVal defaultValue As Variant) As Variant
    EnsureContextReady

    If gConfig Is Nothing Then
        Set gConfig = GetConfigDictionary(False, True) ' allow silent fallback at runtime
    End If

    If gConfig Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If

    If gConfig.Exists(key) Then
        GetConfigValue = gConfig(key)
    Else
        GetConfigValue = defaultValue
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetConfigValueDirect
'
' 功能说明      : 直接获取配置值（内部使用），不检查上下文状态
'               : Internal - direct accessor that does not call EnsureContextReady
'               : Used during InitContext (InitProjectLogger) when init is in progress but gConfig is already built
' 参数          : key - 配置键
'               : defaultValue - 可选，默认值
' 返回          : Variant - 配置值或默认值
' Purpose       : Direct config value access (internal), does not check context readiness
' Contract      : Platform / Internal
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Private Function GetConfigValueDirect(ByVal key As String, Optional ByVal defaultValue As Variant) As Variant
    If gConfig Is Nothing Then
        GetConfigValueDirect = defaultValue
        Exit Function
    End If

    If gConfig.Exists(key) Then
        GetConfigValueDirect = gConfig(key)
    Else
        GetConfigValueDirect = defaultValue
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [P] Config
'
' 功能说明      : 返回配置字典对象
' 参数          : 无
' 返回          : Object - 配置字典对象
' Purpose       : Returns configuration dictionary object
' Contract      : Platform / Query-only
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Property Get Config() As Object
    EnsureContextReady
    If gConfig Is Nothing Then
        Err.Raise vbObjectError + 1401, "plat_context.Config", _
                  "Config not initialized (unexpected)."
    End If
    Set Config = gConfig
End Property

' ==============================================================================================
' SECTION 05: RANGE REGISTRY (CACHED OBJECT ACCESS)
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] ReadRangeFromSheet
'
' 功能说明      : 从工作表读取范围值，如果读取失败则返回Empty
' 参数          : ws - 工作表对象
'               : nameOrAddress - 范围名称或地址
' 返回          : Variant - 范围值或Empty
' Purpose       : Reads range value from worksheet, returns Empty on failure
' Contract      : Platform / Internal
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Private Function ReadRangeFromSheet(ByVal ws As Worksheet, _
                                   ByVal nameOrAddress As String) As Variant
    On Error GoTo EH

    Dim rng As Range
    Set rng = ws.Range(nameOrAddress)

    ReadRangeFromSheet = rng.Value
    Exit Function

EH:
    ReadRangeFromSheet = Empty
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetRange
'
' 功能说明      : 获取范围对象（带缓存），支持强制重新加载
' 参数          : ws - 工作表对象
'               : nameOrAddress - 范围名称或地址
'               : forceReload - 可选，是否强制重新加载，默认为False
' 返回          : Range - 范围对象，如果找不到则返回Nothing
' Purpose       : Gets range object (cached), supports force reload
' Contract      : Platform / Query-only
' Side Effects  : May modify gRangeCache state
' ----------------------------------------------------------------------------------------------
Public Function GetRange(ByVal ws As Worksheet, ByVal nameOrAddress As String, _
                         Optional ByVal forceReload As Boolean = False) As Range
    EnsureContextReady

    If gRangeCache Is Nothing Then Set gRangeCache = CreateObject("Scripting.Dictionary")

    Dim k As String
    k = ws.Name & "!" & nameOrAddress

    If forceReload Then
        If gRangeCache.Exists(k) Then gRangeCache.Remove k
    End If

    If gRangeCache.Exists(k) Then
        Set GetRange = gRangeCache(k)
        Exit Function
    End If

    Dim rng As Range
    On Error GoTo EH
    Set rng = ws.Range(nameOrAddress)
    On Error GoTo 0

    ' overwrite assignment avoids "key exists" errors under nested call stacks
    Set gRangeCache(k) = rng
    Set GetRange = rng
    Exit Function

EH:
    Set GetRange = Nothing
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetRangeOrRaise
'
' 功能说明      : 获取范围对象或引发错误，如果找不到则抛出异常
' 参数          : ws - 工作表对象
'               : nameOrAddress - 范围名称或地址
'               : forceReload - 可选，是否强制重新加载，默认为False
' 返回          : Range - 范围对象
' Purpose       : Gets range object or raises error if not found
' Contract      : Platform / Query-only
' Side Effects  : May modify gRangeCache state
' ----------------------------------------------------------------------------------------------
Public Function GetRangeOrRaise(ByVal ws As Worksheet, ByVal nameOrAddress As String, _
                               Optional ByVal forceReload As Boolean = False) As Range
    Dim rng As Range
    Set rng = GetRange(ws, nameOrAddress, forceReload)
    If rng Is Nothing Then
        Err.Raise vbObjectError + 1500, "plat_context.GetRangeOrRaise", _
                  "Range not found: " & ws.Name & "!" & nameOrAddress
    End If
    Set GetRangeOrRaise = rng
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetRangeValue
'
' 功能说明      : 获取范围值，如果读取失败则返回默认值
' 参数          : ws - 工作表对象
'               : nameOrAddress - 范围名称或地址
'               : defaultValue - 可选，默认值
' 返回          : Variant - 范围值或默认值
' Purpose       : Gets range value, returns default on failure
' Contract      : Platform / Query-only
'               : Caller is responsible for context readiness; this function depends only on
'               : the caller-supplied ws and does not access module state directly
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function GetRangeValue(ByVal ws As Worksheet, ByVal nameOrAddress As String, _
                             Optional ByVal defaultValue As Variant) As Variant
    Dim v As Variant
    v = ReadRangeFromSheet(ws, nameOrAddress)
    If IsEmpty(v) Then
        GetRangeValue = defaultValue
    Else
        GetRangeValue = v
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [S] ClearRangeCache
'
' 功能说明      : 清除范围缓存
' 参数          : 无
' 返回          : 无
' Purpose       : Clears range cache
' Contract      : Platform / State mutation
' Side Effects  : Clears gRangeCache state
' ----------------------------------------------------------------------------------------------
Public Sub ClearRangeCache()
    Set gRangeCache = Nothing
End Sub

' ==============================================================================================
' SECTION 06: RANGE REGISTRY WARMUP (PRIMARY + FALLBACK)
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RegisterRangesFromSetup
'
' 功能说明      : 从Setup工作表注册范围，基于范围注册表
' 参数          : forceReload - 可选，是否强制重新加载，默认为False
' 返回          : 无
' Purpose       : Registers ranges from Setup worksheet based on range registry
' Contract      : Platform / State mutation
' Side Effects  : Populates gRangeCache state
' ----------------------------------------------------------------------------------------------
Public Sub RegisterRangesFromSetup(Optional ByVal forceReload As Boolean = False)
    EnsureContextReady

    Dim rngReg As Range
    Set rngReg = GetRangeOrRaise(gWsSetup, NR_RNG_SYS_RANGE_REGISTRY, forceReload)

    Dim data As Variant
    data = rngReg.Value2

    ' ---- Shape validation ----
    If Not IsArray(data) Then
        Err.Raise vbObjectError + 1603, "plat_context.RegisterRangesFromSetup", _
                  "Range registry must be a 2D table (at least header + 1 data row)."
    End If

    Dim rUB As Long, cUB As Long
    rUB = 0: cUB = 0

    On Error Resume Next
    rUB = UBound(data, 1)
    cUB = UBound(data, 2)
    On Error GoTo 0

    If rUB = 0 Or cUB = 0 Then
        Err.Raise vbObjectError + 1604, "plat_context.RegisterRangesFromSetup", _
                  "Range registry must return a 2D array."
    End If

    If rUB < 2 Then
        Err.Raise vbObjectError + 1605, "plat_context.RegisterRangesFromSetup", _
                  "Range registry contains no data rows (expected: header + at least 1 row)."
    End If

    If cUB < 3 Then
        Err.Raise vbObjectError + 1606, "plat_context.RegisterRangesFromSetup", _
                  "Range registry must have at least 3 columns: SheetName, RangeName, Required."
    End If

    ' ---- Process rows ----
    Dim i As Long
    Dim sheetName As String
    Dim rangeName As String
    Dim required As Boolean
    Dim ws As Worksheet
    Dim rng As Range

    For i = 2 To rUB
        sheetName = Trim$(CStr(data(i, 1)))
        rangeName = Trim$(CStr(data(i, 2)))

        If sheetName = vbNullString And rangeName = vbNullString Then GoTo ContinueNext
        
        If sheetName = vbNullString Or rangeName = vbNullString Then
            Err.Raise vbObjectError + 1607, "plat_context.RegisterRangesFromSetup", _
                      "Invalid registry row: SheetName and RangeName must both be provided (row " & i & ")."
        End If

        required = core_utils.ToSafeBoolean(data(i, 3), False)

        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            If required Then
                Err.Raise vbObjectError + 1601, "plat_context.RegisterRangesFromSetup", _
                          "Required worksheet missing in registry: " & sheetName
            End If
            GoTo ContinueNext
        End If

        Set rng = Nothing
        Set rng = GetRange(ws, rangeName, forceReload)

        If rng Is Nothing And required Then
            Err.Raise vbObjectError + 1602, "plat_context.RegisterRangesFromSetup", _
                      "Required range missing in registry: " & sheetName & "!" & rangeName
        End If

ContinueNext:
    Next i
End Sub

' ----------------------------------------------------------------------------------------------
' [s] RegisterRanges
'
' 功能说明      : 注册范围（传统方法），基于硬编码的映射
'               : DEPRECATED - keep legacy worksheet pairing for compatibility
' 参数          : forceReload - 可选，是否强制重新加载，默认为False
' 返回          : 无
' Purpose       : Registers ranges (legacy method) based on hardcoded mapping
' Contract      : Platform / State mutation
' Side Effects  : Populates gRangeCache state
' ----------------------------------------------------------------------------------------------
Private Sub RegisterRanges(Optional ByVal forceReload As Boolean = False)
    EnsureContextReady

    ' Legacy minimal set (fallback)
    Call GetRange(gWsMain, NR_RNG_SYS_MAIN_CONFIG, forceReload)
    Call GetRange(gWsMain, NR_RNG_SYS_SEGMENT, forceReload)

    ' NOTE: legacy pairing (per original mapping)
    Call GetRange(gWsSetup, NR_RNG_SYS_SUBLOB_TO_LOB, forceReload)
    Call GetRange(gWsTreaty, NR_RNG_TREATY_REFERENCE, forceReload)
    Call GetRange(gWsSubLoB, NR_RNG_MAP_SUBLOB_TO_TREATY, forceReload)

    Call GetRange(gWsGN, NR_RNG_OUT_WS_LIST_ANCHOR, forceReload)
End Sub

' ==============================================================================================
' SECTION 07: INTERNAL CONFIG REGISTRATION HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [s] RegisterConfig
'
' 功能说明      : 注册配置项到字典中
' 参数          : d - 字典对象
'               : key - 配置键
'               : ws - 工作表对象
'               : defaultValue - 默认值
' 返回          : 无
' Purpose       : Registers configuration item into dictionary
' Contract      : Platform / Internal
' Side Effects  : Modifies dictionary state
' ----------------------------------------------------------------------------------------------
Private Sub RegisterConfig(ByVal d As Object, ByVal key As String, ByVal ws As Worksheet, ByVal defaultValue As Variant)
    Dim v As Variant
    v = ReadValueFromSheet(ws, key, defaultValue)
    d(key) = v
End Sub

' ----------------------------------------------------------------------------------------------
' [f] ReadValueFromSheet
'
' 功能说明      : 从工作表读取值，根据默认值类型进行安全转换
' 参数          : ws - 工作表对象
'               : key - 配置键
'               : defaultValue - 默认值
' 返回          : Variant - 转换后的值或默认值
' Purpose       : Reads value from worksheet with safe conversion based on default value type
' Contract      : Platform / Internal
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Private Function ReadValueFromSheet(ByVal ws As Worksheet, ByVal key As String, ByVal defaultValue As Variant) As Variant
    ' Silent Config Fallback (Design Intent):
    ' - Any read / conversion issue returns defaultValue.
    ' - However, numeric conversion must avoid silent Banker's rounding for integer configs.
    On Error GoTo EH

    If ws Is Nothing Then
        ReadValueFromSheet = defaultValue
        Exit Function
    End If

    Dim v As Variant
    v = ws.Range(key).Value2

    Select Case VarType(defaultValue)
        Case vbBoolean
            ReadValueFromSheet = core_utils.ToSafeBoolean(v, CBool(defaultValue))

        Case vbInteger, vbLong
            ' Strict integer semantics: reject non-integer values (including 2.5) to avoid Banker's rounding.
            If IsNumeric(v) Then
                Dim dv As Double
                dv = CDbl(v)
                If dv = Fix(dv) Then
                    ReadValueFromSheet = CLng(dv)
                Else
                    ReadValueFromSheet = defaultValue
                End If
            Else
                ReadValueFromSheet = defaultValue
            End If

        Case vbSingle, vbDouble, vbCurrency
            If IsNumeric(v) Then
                ReadValueFromSheet = CDbl(v)
            Else
                ReadValueFromSheet = defaultValue
            End If

        Case vbString
            If IsEmpty(v) Or IsNull(v) Or (Trim$(CStr(v)) = vbNullString) Then
                ReadValueFromSheet = defaultValue
            Else
                ReadValueFromSheet = Trim$(CStr(v))
            End If

        Case Else
            ReadValueFromSheet = v
    End Select

    Exit Function

EH:
    ReadValueFromSheet = defaultValue
End Function
