VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plat_logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' MODULE NAME       : plat_logger
' LAYER             : platform
' PURPOSE           : Platform logging service providing structured log writing with severity filtering,
'                     worksheet output, and convenience wrappers for Business layer usage.
' DEPENDS           : core_logging, Excel Object Model
' NOTE              : - Primary contract WriteLog must maintain signature compatibility with CallByName
'                       from plat_runtime (level, layer, module, proc, message, context).
'                     - Severity levels: DEBUG/INFO/PERF/WARN/ERROR with configurable minimum level.
'                     - Log lines are formatted by core_logging.FormatLogLine (frozen contract).
' STATUS            : Frozen
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Introduced logging service with severity filtering and worksheet output.
'   - Init (Contract): Established WriteLog signature for CallByName compatibility.
'   - Init (Levels): Implemented DEBUG/INFO/PERF/WARN/ERROR severity levels with ranking.

' v2.0.0
'   - Refactor (Architecture): Positioned as Platform layer service under layered model.
'   - Refactor (Encapsulation): Separated formatting logic to core_logging (frozen contract).
'   - Refactor (Interface): Added convenience wrappers (Info, Warn, Err, LogDebug, Perf).

' v2.1.0
'   - Fix (Level Ranking): Added PERF level with rank 25 (between INFO and WARN).
'   - Fix (Sheet Write): Added safety checks for pWs and pAnchor before writing.
'   - Fix (Error Handling): Added On Error Resume Next in WriteLog to prevent logging failures.

' v2.1.1
'   - Fix (Level Handling): Unknown levels default to INFO rank (20) for safe fallback.
'   - Fix (Init Validation): Added UCase$ and Trim$ for minLevel parameter normalization.
'   - Fix (Anchor Safety): Added check for WsLog Nothing before setting pAnchor.
'   - Refine (API Naming): Renamed convenience wrappers (Info/Warn/Err/LogDebug/Perf → InfoLog/WarnLog/ErrLog/DebugLog/PerfLog)
'       to avoid shadowing VBA built-ins and improve call-site clarity.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 01: INITIALIZATION
'   [S] Init                                    - Initialize logger with configuration (enable sheet/immediate, min level, worksheet, anchor)
'
' SECTION 02: PRIMARY CONTRACT
'   [S] WriteLog                                - Primary logging contract - writes structured log entry with severity filtering
'
' SECTION 03: CONVENIENCE WRAPPERS
'   [S] InfoLog/WarnLog/ErrLog/DebugLog/PerfLog - Convenience wrapper for Info/Warn/Err/Debug/Perf level logging
'
' SECTION 04: INTERNAL HELPERS
'   [f] ShouldWrite                             - Determines if a log level should be written based on configured minimum level
'   [f] LevelRank                               - Converts log level string to numeric rank for comparison
'   [s] WriteLineToSheet                        - Writes log line to worksheet at anchor position with row offset
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

Private pEnableSheet As Boolean
Private pEnableImmediate As Boolean
Private pMinLevel As String          ' "DEBUG"/"INFO"/"WARN"/"ERROR"/"PERF"
Private pWs As Worksheet
Private pAnchor As Range
Private pNextRowOffset As Long
Private pSheetDisabledDueToOverflow As Boolean

' ==============================================================================================
' SECTION 01: INITIALIZATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Init
'
' 功能说明      : 初始化日志记录器，设置配置参数
' 参数          : enableSheet - 是否启用工作表输出
'               : enableImmediate - 是否启用即时窗口输出
'               : minLevel - 最小日志级别（DEBUG/INFO/PERF/WARN/ERROR）
'               : WsLog - 日志工作表对象
'               : anchorCellAddress - 锚点单元格地址
' 返回          : 无
' Purpose       : Initializes logger with configuration parameters
' Contract      : Platform / Lifecycle
' Side Effects  : Modifies private module state (pEnableSheet, pEnableImmediate, pMinLevel,
'               : pWs, pAnchor, pNextRowOffset, pSheetDisabledDueToOverflow)
' ----------------------------------------------------------------------------------------------
Public Sub Init( _
    ByVal enableSheet As Boolean, _
    ByVal enableImmediate As Boolean, _
    ByVal minLevel As String, _
    ByVal WsLog As Worksheet, _
    ByVal anchorCellAddress As String _
)
    pEnableSheet = enableSheet
    pEnableImmediate = enableImmediate
    pMinLevel = UCase$(Trim$(minLevel))
    pSheetDisabledDueToOverflow = False

    Set pWs = WsLog
    If Not WsLog Is Nothing Then
        On Error Resume Next
        Set pAnchor = WsLog.Range(anchorCellAddress).Cells(1, 1)
        On Error GoTo 0
    Else
        Set pAnchor = Nothing
    End If

    pNextRowOffset = 0
End Sub

' ==============================================================================================
' SECTION 02: PRIMARY CONTRACT
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] WriteLog
'
' 功能说明      : 主要日志契约，写入结构化日志条目，支持级别过滤
' 参数          : level - 日志级别 (DEBUG/INFO/PERF/WARN/ERROR)
'               : layer - 层级名称 (Core/Platform/Business)
'               : moduleName - 模块名称
'               : procName - 过程名称
'               : message - 日志消息
'               : context - 可选的上下文信息
' 返回          : 无
' Purpose       : Primary logging contract; delegates formatting to core_logging.FormatLogLine
'               : (frozen contract); called by plat_runtime via CallByName
' Contract      : Platform / Primary API (signature must remain compatible with CallByName)
' Side Effects  : May write to Immediate window (Debug.Print) if pEnableImmediate is True
'               : May write to worksheet (WriteLineToSheet) if pEnableSheet is True
'               : Sink errors are suppressed (On Error Resume Next)
' ----------------------------------------------------------------------------------------------
Public Sub WriteLog( _
    ByVal level As String, _
    ByVal layer As String, _
    ByVal moduleName As String, _
    ByVal procName As String, _
    ByVal message As String, _
    Optional ByVal context As Variant _
)
    level = UCase$(Trim$(level))
    If Not ShouldWrite(level) Then Exit Sub

    Dim line As String
    line = core_logging.FormatLogLine(level, layer, moduleName, procName, message, context)

    On Error Resume Next    ' sink operations only; logging must never raise to caller
    If pEnableImmediate Then Debug.Print line
    If pEnableSheet Then WriteLineToSheet line
    On Error GoTo 0
End Sub

' ==============================================================================================
' SECTION 03: CONVENIENCE WRAPPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] InfoLog/WarnLog/ErrLog/DebugLog/PerfLog
'
' 功能说明      : Info/Warn/Err/Debug/Perf级别日志记录的便捷包装器
' 参数          : layer - 层级名称
'               : moduleName - 模块名称
'               : procName - 过程名称
'               : message - 日志消息
'               : context - 可选的上下文信息
' 返回          : 无
' Purpose       : Convenience wrapper for Info/Warn/Err/Debug/Perf level logging
' Contract      : Platform / Public API
' Side Effects  : Calls WriteLog with Info/Warn/Err/Debug/Perf level; inherits WriteLog side effects
' ----------------------------------------------------------------------------------------------
Public Sub InfoLog(ByVal layer As String, ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WriteLog "INFO", layer, moduleName, procName, message, context
End Sub

Public Sub WarnLog(ByVal layer As String, ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WriteLog "WARN", layer, moduleName, procName, message, context
End Sub

Public Sub ErrLog(ByVal layer As String, ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WriteLog "ERROR", layer, moduleName, procName, message, context
End Sub

Public Sub DebugLog(ByVal layer As String, ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WriteLog "DEBUG", layer, moduleName, procName, message, context
End Sub

Public Sub PerfLog(ByVal layer As String, ByVal moduleName As String, ByVal procName As String, ByVal message As String, Optional ByVal context As Variant)
    WriteLog "PERF", layer, moduleName, procName, message, context
End Sub

' ==============================================================================================
' SECTION 04: INTERNAL HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] ShouldWrite
'
' 功能说明      : 根据配置的最小级别确定是否应写入指定级别的日志
' 参数          : level - 日志级别
' 返回          : Boolean - 是否应写入该级别日志
' Purpose       : Determines if a log level should be written based on configured minimum level
' Contract      : Platform / Internal
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Private Function ShouldWrite(ByVal level As String) As Boolean
    Dim lv As Long, minv As Long
    lv = LevelRank(level)
    minv = LevelRank(pMinLevel)
    ShouldWrite = (lv >= minv)
End Function

' ----------------------------------------------------------------------------------------------
' [f] LevelRank
'
' 功能说明      : 将日志级别字符串转换为数值等级以便比较
' 参数          : level - 日志级别字符串
' 返回          : Long - 数值等级（DEBUG=10, INFO=20, PERF=25, WARN=30, ERROR=40）
' Purpose       : Converts log level string to numeric rank for comparison
' Contract      : Platform / Internal
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Private Function LevelRank(ByVal level As String) As Long
    Select Case UCase$(Trim$(level))
        Case "DEBUG": LevelRank = 10
        Case "INFO":  LevelRank = 20
        Case "PERF":  LevelRank = 25
        Case "WARN":  LevelRank = 30
        Case "ERROR": LevelRank = 40
        Case Else:    LevelRank = 20 ' unknown treated as INFO
    End Select
End Function

' ----------------------------------------------------------------------------------------------
' [s] WriteLineToSheet
'
' 功能说明      : 将日志行写入工作表的锚点位置，自动扫描已有内容确定追加行，并防止超出工作表行限制
' 参数          : line - 要写入的日志行
' 返回          : 无
' Purpose       : Writes log line to worksheet at anchor position with append-mode row detection
'               : and overflow fuse to prevent exceeding worksheet row limit
' Contract      : Platform / Internal
' Side Effects  : Modifies worksheet cell values; increments pNextRowOffset; may set pSheetDisabledDueToOverflow
' ----------------------------------------------------------------------------------------------
Private Sub WriteLineToSheet(ByVal line As String)
    If pWs Is Nothing Then Exit Sub
    If pAnchor Is Nothing Then Exit Sub
    If Len(line) = 0 Then Exit Sub

    ' ---- Overflow fuse ----
    ' Once tripped, sheet sink is permanently disabled for this session.
    If pSheetDisabledDueToOverflow Then Exit Sub

    ' ---- Append mode: scan anchor column to find next write row ----
    ' Handles cases where the sheet already contains log entries from a previous run.
    If pNextRowOffset = 0 Then
        Dim lastRow As Long
        lastRow = pWs.Cells(pWs.Rows.Count, pAnchor.Column).End(xlUp).Row

        If lastRow < pAnchor.Row Then
            ' Anchor row itself is empty: start writing at anchor
            pNextRowOffset = 0
        ElseIf lastRow = pAnchor.Row Then
            ' Anchor row has content: check if it is empty
            If IsEmpty(pAnchor.Value2) And Len(pAnchor.Formula) = 0 Then
                pNextRowOffset = 0
            Else
                pNextRowOffset = 1
            End If
        Else
            ' Rows below anchor have content: append after last used row
            pNextRowOffset = lastRow - pAnchor.Row + 1
        End If
    End If

    ' ---- Row overflow check ----
    Dim targetRow As Long
    targetRow = pAnchor.Row + pNextRowOffset

    If targetRow > pWs.Rows.Count Then
        pSheetDisabledDueToOverflow = True
        Exit Sub
    End If

    pAnchor.Offset(pNextRowOffset, 0).Value2 = line
    pNextRowOffset = pNextRowOffset + 1
End Sub
