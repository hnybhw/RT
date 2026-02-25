Attribute VB_Name = "app_01_basic"
' ==============================================================================================
' MODULE NAME     : app_01_basic
' PURPOSE         : 系统核心基础模块，提供全局常量、工作表属性、配置管理、错误处理、服务定位、通用工具函数等核心能力
'                 : Core foundation module providing global constants, worksheet properties, configuration management,
'                 : error handling, service location, and common utility functions
' DEPENDS         : C10KProcessor (处理器类，用于实例化与管理 / Processor class for instantiation and management)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 全局常量声明 / Global Constants
'   [C] 工作表名称常量         - 定义系统固定工作表名称，统一工作表访问标识 / Worksheet name constants
'   [C] 通用全局常量           - 定义Excel边界、列标识等通用常量 / General constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'
' SECTION 2: 模块级私有变量 / Module-Level Variables
'   [V] 配置与实例管理变量     - 全局配置字典、活动处理器/工作簿集合，仅模块内可访问 / Configuration and instance management variables
'
' SECTION 3: 工作表快捷属性 / Worksheet Properties
'   [F] 工作表只读属性         - 提供系统固定工作表的直接访问属性 / Read-only worksheet access properties
'   [F] 系统配置开关属性       - 提供错误/消息显示的配置开关 / System configuration flags
'
' SECTION 4: 全局配置管理 / Global Configuration Management
'   [F] GetConfig              - 懒加载创建全局配置字典 / Lazy-loaded global configuration dictionary
'   [S] RefreshConfig          - 重新从工作表加载配置 / Refresh configuration from worksheet
'
' SECTION 5: C10KProcessor类相关服务管理 / C10KProcessor Service Management
'   [F] GetProcessor           - 获取/创建C10KProcessor实例，支持实例缓存 / Get or create C10KProcessor instance
'   [S] ReleaseProcessor       - 显式释放指定的C10KProcessor实例 / Explicitly release processor instance
'   [S] TrackWorkbook          - 跟踪外部打开的工作簿 / Track external workbooks for cleanup
'   [S] ServiceLocator_Cleanup - 批量清理所有活动实例和跟踪的工作簿 / Clean up all active instances and tracked workbooks
'
' SECTION 6: 系统环境与提示工具 / System Environment Utilities
'   [F] MessageStart           - 启动任务计时，关闭屏幕刷新 / Start task timing, disable screen updating
'   [S] MessageEnd             - 结束任务计时，恢复Excel环境 / End task timing, restore Excel environment
'   [S] ResetGlobalVariables   - 重置全局系统状态 / Reset global system state
'   [S] FormatSheetStandard    - 工作表标准化格式化 / Standard worksheet formatting
'
' SECTION 7: 全局错误处理 / Global Error Handling
'   [S] HandleError            - 集中式错误处理 / Centralized error handling
'
' SECTION 8: 通用核心工具函数 / Common Utility Functions
'   [F] IsWorkSheetExist       - 检查工作表是否存在 / Check if worksheet exists
'
' SECTION 9: 全局日志系统 / Global Logging System
'   [S] WriteLog               - 全局标准化日志写入 / Standardized logging
'   [S] ClearLog               - 清空日志区域 / Clear log area
'
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 1: 全局常量声明 / Global Constants
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 模块名称常量 - 用于日志记录，避免硬编码 / Module name constant for logging
' ----------------------------------------------------------------------------------------------
Public Const MODULE_NAME As String = "app_01_basic"

' ----------------------------------------------------------------------------------------------
' [C] 工作表名称常量 / Worksheet name constants
' 说明：定义系统固定工作表的名称常量，统一工作表访问入口，避免硬编码
'       Constants for system worksheet names, providing unified access without hardcoding
' ----------------------------------------------------------------------------------------------
Public Const SHT_NAME_LOG     As String = "1@Log"
Public Const SHT_NAME_SETUP   As String = "2@Setup"
Public Const SHT_NAME_MAIN    As String = "3@Main"
Public Const SHT_NAME_TREATY  As String = "4@Treaty"
Public Const SHT_NAME_SUBLOB  As String = "5@SubLoB"
Public Const SHT_NAME_GN      As String = "6@GN"
Public Const SHT_NAME_EL      As String = "7@EL"
Public Const SHT_NAME_RE      As String = "8@RE"
Public Const SHT_NAME_KPI     As String = "9@KPI"

' ----------------------------------------------------------------------------------------------
' [C] 通用全局常量 / General constants
' 说明：定义Excel系统边界、通用标识等常量，统一系统通用参数
'       Constants for Excel boundaries and general identifiers
' ----------------------------------------------------------------------------------------------
Public Const surfix   As String = "@Asm"      ' 假设工作表后缀 / Assumption sheet suffix
Public Const rowMax   As Long = 1048576       ' Excel最大行数 / Excel maximum rows
Public Const colNew   As String = "AA"        ' 溢出列标识 / Overflow column identifier

' ==============================================================================================
' SECTION 2: 模块级私有变量 / Module-Level Variables
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [V] 配置与实例管理变量 / Configuration and instance management variables
' 说明：模块级私有变量，仅模块内可访问，用于管理全局配置和类实例/工作簿资源
'       Module-level private variables for managing global configuration and resources
' ----------------------------------------------------------------------------------------------
Private pConfig As Object                    ' 全局配置字典 / Global configuration dictionary
Private pActiveProcessors As Collection      ' 活动C10KProcessor实例集合 / Active C10KProcessor instances
Private pActiveWorkbooks As Collection       ' 跟踪的外部工作簿集合 / Tracked external workbooks

' ==============================================================================================
' SECTION 3: 工作表快捷属性 / Worksheet Properties
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] Log - 获取Log工作表 / Get Log worksheet
' 返回值：Worksheet - Log工作表对象 / Setup worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get Log() As Worksheet
    Set Log = ThisWorkbook.Sheets(SHT_NAME_LOG)
End Property

' ----------------------------------------------------------------------------------------------
' [F] Setup - 获取Setup工作表 / Get Setup worksheet
' 返回值：Worksheet - Setup工作表对象 / Setup worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get Setup() As Worksheet
    Set Setup = ThisWorkbook.Sheets(SHT_NAME_SETUP)
End Property

' ----------------------------------------------------------------------------------------------
' [F] Main - 获取Main工作表 / Get Main worksheet
' 返回值：Worksheet - Main工作表对象 / Main worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get Main() As Worksheet
    Set Main = ThisWorkbook.Sheets(SHT_NAME_MAIN)
End Property

' ----------------------------------------------------------------------------------------------
' [F] Treaty - 获取Treaty工作表 / Get Treaty worksheet
' 返回值：Worksheet - Treaty工作表对象 / Treaty worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get Treaty() As Worksheet
    Set Treaty = ThisWorkbook.Sheets(SHT_NAME_TREATY)
End Property

' ----------------------------------------------------------------------------------------------
' [F] SubLoB - 获取SubLoB工作表 / Get SubLoB worksheet
' 返回值：Worksheet - SubLoB工作表对象 / SubLoB worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get subLoB() As Worksheet
    Set subLoB = ThisWorkbook.Sheets(SHT_NAME_SUBLOB)
End Property

' ----------------------------------------------------------------------------------------------
' [F] GN - 获取GN工作表 / Get GN worksheet
' 返回值：Worksheet - GN工作表对象 / GN worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get GN() As Worksheet
    Set GN = ThisWorkbook.Sheets(SHT_NAME_GN)
End Property

' ----------------------------------------------------------------------------------------------
' [F] EL - 获取EL工作表 / Get EL worksheet
' 返回值：Worksheet - EL工作表对象 / EL worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get EL() As Worksheet
    Set EL = ThisWorkbook.Sheets(SHT_NAME_EL)
End Property

' ----------------------------------------------------------------------------------------------
' [F] RE - 获取RE工作表 / Get RE worksheet
' 返回值：Worksheet - RE工作表对象 / RE worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get RE() As Worksheet
    Set RE = ThisWorkbook.Sheets(SHT_NAME_RE)
End Property

' ----------------------------------------------------------------------------------------------
' [F] KPI - 获取KPI工作表 / Get KPI worksheet
' 返回值：Worksheet - KPI工作表对象 / KPI worksheet object
' ----------------------------------------------------------------------------------------------
Public Property Get KPI() As Worksheet
    Set KPI = ThisWorkbook.Sheets(SHT_NAME_KPI)
End Property

' ----------------------------------------------------------------------------------------------
' [F] ShowError - 获取错误显示开关 / Get error display flag
' 说明：从Main工作表读取配置，控制是否弹出错误提示框
'       Read from Main worksheet to control error message display
' 返回值：Boolean - True=显示错误提示 / Show error messages
' ----------------------------------------------------------------------------------------------
Public Property Get ShowError() As Boolean
    ShowError = Main.Range("p_show_error").value
End Property

' ----------------------------------------------------------------------------------------------
' [F] ShowMessage - 获取消息显示开关 / Get message display flag
' 说明：从Main工作表读取配置，控制是否弹出操作完成提示框
'       Read from Main worksheet to control completion message display
' 返回值：Boolean - True=显示消息提示 / Show completion messages
' ----------------------------------------------------------------------------------------------
Public Property Get ShowMessage() As Boolean
    ShowMessage = Main.Range("p_show_message").value
End Property

' ==============================================================================================
' SECTION 4: 全局配置管理 / Global Configuration Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetConfig
' 说明：懒加载创建并返回全局配置字典，从Main工作表读取配置项，无值/无效时设置默认值
'       Lazy-loaded global configuration dictionary with defaults from Main worksheet
' 返回值：Object - 晚绑定的Scripting.Dictionary配置字典 / Late-bound Scripting.Dictionary
' ----------------------------------------------------------------------------------------------
Public Function GetConfig() As Object
    ' 懒加载：字典为空时才创建 / Lazy initialization
    If pConfig Is Nothing Then
        Set pConfig = CreateObject("Scripting.Dictionary")
        
        ' 读取模拟最大年份 / Read simulation year maximum
        On Error Resume Next
        pConfig("yearMax") = CLng(val(Main.Range("p_sim_yrs").value))
        On Error GoTo 0
        
        ' 无效值设默认值 / Set default if invalid
        If Not pConfig.Exists("yearMax") Or pConfig("yearMax") <= 0 Then
            pConfig("yearMax") = 10000
        End If
        
        ' 设置基础默认配置项 / Set default configuration
        pConfig("materialityThreshold") = 0.5    ' 损失重要性阈值 / Loss materiality threshold
        pConfig("chunkSize") = 50000             ' 数据分块处理大小 / Chunk size for large data
        pConfig("pythonTimeout") = 300           ' Python调用超时时间（秒）/ Python timeout in seconds
        
        ' 读取分块处理阈值 / Read chunk processing threshold
        On Error Resume Next
        pConfig("chunkThreshold") = CLng(val(Main.Range("p_chunk_threshold").value))
        On Error GoTo 0
        
        If Not pConfig.Exists("chunkThreshold") Or pConfig("chunkThreshold") <= 0 Then
            pConfig("chunkThreshold") = 2000000  ' 默认200万行 / Default 2 million rows
        End If
        
        Call WriteLog(MODULE_NAME, "GetConfig", "全局配置初始化完成 / Global configuration initialized", "配置管理")
    End If
    
    Set GetConfig = pConfig
End Function

' ----------------------------------------------------------------------------------------------
' [S] RefreshConfig
' 说明：重新从Main工作表加载核心配置项，更新全局配置字典
'       Reload core configuration from Main worksheet and update global configuration
' ----------------------------------------------------------------------------------------------
Public Sub RefreshConfig()
    If Not pConfig Is Nothing Then
        On Error Resume Next
        pConfig("yearMax") = CLng(val(Main.Range("p_sim_yrs").value))
        On Error GoTo 0
        
        If pConfig("yearMax") <= 0 Then pConfig("yearMax") = 10000
        
        Call WriteLog(MODULE_NAME, "RefreshConfig", "全局配置已刷新 / Global configuration refreshed", "配置管理")
    End If
End Sub

' ==============================================================================================
' SECTION 5: C10KProcessor类相关服务管理 / C10KProcessor Service Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetProcessor
' 说明：获取/创建C10KProcessor实例，支持空ID创建新实例、指定ID获取缓存实例
'       Get or create C10KProcessor instance with optional caching
' 参数：instanceId - String（可选），实例唯一标识，空值则创建新实例 / Instance ID, empty for new instance
' 返回值：C10KProcessor - 处理器实例对象 / Processor instance
' ----------------------------------------------------------------------------------------------
Public Function GetProcessor(Optional ByVal instanceId As String = "") As C10KProcessor
    ' 实例集合为空时初始化 / Initialize collection if needed
    If pActiveProcessors Is Nothing Then
        Set pActiveProcessors = New Collection
    End If
    
    ' 空ID：创建新实例并注入全局配置 / Create new instance with global configuration
    If instanceId = "" Then
        Dim proc As C10KProcessor
        Set proc = New C10KProcessor
        
        Dim config As Object
        Set config = GetConfig()
        
        ' 注入全局配置 / Inject global configuration
        On Error Resume Next
        proc.config("yearMax") = CLng(config("yearMax"))
        proc.config("materialityThreshold") = CDbl(config("materialityThreshold"))
        proc.config("chunkSize") = CLng(config("chunkSize"))
        On Error GoTo 0
        
        ' 生成唯一实例ID，加入集合缓存 / Generate unique ID and cache
        instanceId = "PROC_" & Timer & "_" & pActiveProcessors.count
        pActiveProcessors.Add proc, instanceId
        
        Call WriteLog(MODULE_NAME, "GetProcessor", "创建新处理器实例 / New processor instance created: " & instanceId, "实例管理")
        Set GetProcessor = proc
    Else
        ' 指定ID：从集合中获取缓存实例 / Get cached instance by ID
        On Error Resume Next
        Set GetProcessor = pActiveProcessors(instanceId)
        On Error GoTo 0
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [S] ReleaseProcessor
' 说明：显式释放指定ID的C10KProcessor实例，调用实例清理方法并从集合中移除
'       Explicitly release processor instance by ID
' 参数：instanceId - String，待释放实例的唯一标识 / Instance ID to release
' ----------------------------------------------------------------------------------------------
Public Sub ReleaseProcessor(ByVal instanceId As String)
    On Error Resume Next
    Dim proc As C10KProcessor
    Set proc = pActiveProcessors(instanceId)
    
    If Not proc Is Nothing Then
        proc.Cleanup
        pActiveProcessors.Remove instanceId
        Call WriteLog(MODULE_NAME, "ReleaseProcessor", "释放处理器实例 / Released processor instance: " & instanceId, "实例管理")
    End If
End Sub

' ----------------------------------------------------------------------------------------------
' [S] TrackWorkbook
' 说明：跟踪外部打开的工作簿，加入跟踪集合，用于后续自动清理
'       Track external workbooks for automatic cleanup
' 参数：wb - Workbook，待跟踪的外部工作簿对象 / Workbook to track
' ----------------------------------------------------------------------------------------------
Public Sub TrackWorkbook(ByRef wb As Workbook)
    If pActiveWorkbooks Is Nothing Then
        Set pActiveWorkbooks = New Collection
    End If
    
    ' 仅跟踪非本工作簿的外部工作簿 / Only track external workbooks
    If wb.path <> ThisWorkbook.path Then
        pActiveWorkbooks.Add wb, wb.Name
        Call WriteLog(MODULE_NAME, "TrackWorkbook", "跟踪外部工作簿 / Tracking external workbook: " & wb.Name, "实例管理")
    End If
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ServiceLocator_Cleanup
' 说明：批量清理所有资源，释放所有活动C10KProcessor实例，关闭所有跟踪的外部工作簿
'       Clean up all resources: release processors and close tracked workbooks
' ----------------------------------------------------------------------------------------------
Public Sub ServiceLocator_Cleanup()
    On Error Resume Next
    
    ' 清理所有活动处理器实例 / Clean up all processor instances
    If Not pActiveProcessors Is Nothing Then
        Dim proc As C10KProcessor
        For Each proc In pActiveProcessors
            proc.Cleanup
        Next proc
        Set pActiveProcessors = Nothing
        Call WriteLog(MODULE_NAME, "ServiceLocator_Cleanup", "已清理所有处理器实例 / All processor instances cleaned up", "实例管理")
    End If
    
    ' 关闭所有跟踪的外部工作簿 / Close all tracked workbooks
    If Not pActiveWorkbooks Is Nothing Then
        Dim wb As Workbook
        For Each wb In pActiveWorkbooks
            wb.Close SaveChanges:=False
        Next wb
        Set pActiveWorkbooks = Nothing
        Call WriteLog(MODULE_NAME, "ServiceLocator_Cleanup", "已关闭所有跟踪工作簿 / All tracked workbooks closed", "实例管理")
    End If
End Sub

' ==============================================================================================
' SECTION 6: 系统环境与提示工具 / System Environment Utilities
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] MessageStart
' 说明：启动任务计时，关闭Excel屏幕刷新提升运行速度，设置状态栏初始提示
'       Start task timing, disable screen updating, set status bar message
' 参数：context - String（可选），任务描述，显示在状态栏 / Task description for status bar
' 返回值：Double - 任务启动的计时器时间 / Start time for duration calculation
' ----------------------------------------------------------------------------------------------
Public Function MessageStart(Optional ByVal context As String = "") As Double
    If context <> "" Then
        Application.StatusBar = ">>> 开始执行 / Starting: " & context
        Call WriteLog(MODULE_NAME, "MessageStart", "任务开始 / Task started: " & context, "计时")
    End If
    
    MessageStart = Timer
    Application.ScreenUpdating = False
End Function

' ----------------------------------------------------------------------------------------------
' [S] MessageEnd
' 说明：结束任务计时，恢复Excel系统环境，计算并显示任务耗时
'       End task timing, restore Excel environment, calculate and display duration
' 参数：startTime - Double（可选），任务启动时间 / Start time from MessageStart
'       context - String（可选），任务描述或自定义消息 / Task description or custom message
' ----------------------------------------------------------------------------------------------
Public Sub MessageEnd(Optional ByVal startTime As Double = 0, Optional ByVal context As String = "")
    Dim endTime As Double: endTime = Timer
    Dim totalSeconds As Double
    
    ' 计算任务耗时 / Calculate duration
    If startTime > 0 Then
        totalSeconds = endTime - startTime
        ' 处理跨午夜的时间差 / Handle midnight crossover
        If totalSeconds < 0 Then totalSeconds = totalSeconds + 86400
    End If
    
    ' 格式化耗时信息 / Format duration message
    Dim timeMessage As String
    If totalSeconds < 60 Then
        timeMessage = Format(totalSeconds, "0.00") & " 秒 / seconds"
    Else
        Dim minPart As Long: minPart = Int(totalSeconds / 60)
        Dim secPart As Double: secPart = totalSeconds - (minPart * 60)
        timeMessage = minPart & " 分 / min " & Format(secPart, "0.00") & " 秒 / sec"
    End If
    
    ' 恢复Excel系统环境 / Restore Excel environment
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' 激活系统主工作簿和Main工作表 / Activate main workbook and worksheet
    On Error Resume Next
    ThisWorkbook.Activate
    Main.Activate
    On Error GoTo 0
    
    Call WriteLog(MODULE_NAME, "MessageEnd", "任务完成 / Task completed: " & context & "，耗时 / Duration: " & timeMessage, "计时")
    
    ' 按配置弹出提示框 / Show message if configured
    If ShowMessage Then
        If context = "" Then
            MsgBox "流程执行完成！" & vbCrLf & "总耗时: " & timeMessage, vbInformation, "再保险引擎"
        Else
            MsgBox context & vbCrLf & "总耗时 / Total time: " & timeMessage, vbInformation, "再保险引擎"
        End If
    End If
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ResetGlobalVariables
' 说明：重置全局系统状态，清理服务定位器资源，恢复Excel状态栏和屏幕刷新
'       Reset global system state, clean up resources, restore Excel environment
' ----------------------------------------------------------------------------------------------
Public Sub ResetGlobalVariables()
    ServiceLocator_Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call WriteLog(MODULE_NAME, "ResetGlobalVariables", "全局变量已重置 / Global variables reset", "系统工具")
End Sub

' ----------------------------------------------------------------------------------------------
' [S] FormatSheetStandard
' 说明：工作表标准化格式化，统一表头样式、自动筛选、冻结窗格、列宽自适应、数字格式
'       Standard worksheet formatting: header style, autofilter, freeze panes, auto-fit, number formats
' 参数：ws - Worksheet，待格式化的工作表对象 / Worksheet to format
'       commaColIdx - Long（可选），需设置千分位的列索引 / Column index for thousand separator
'       pctColIdx - Long（可选），需设置百分比的列索引 / Column index for percentage format
' ----------------------------------------------------------------------------------------------
Public Sub FormatSheetStandard(ByRef ws As Worksheet, Optional ByVal commaColIdx As Long = 0, Optional ByVal pctColIdx As Long = 0)
    On Error Resume Next
    
    With ws
        .Activate
        
        ' 表头样式：加粗、灰色背景、居中对齐、下边框 / Header style
        With .rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        ' 开启自动筛选 / Enable autofilter
        .AutoFilterMode = False
        .Range("A1").CurrentRegion.AutoFilter
        
        ' 设置数字格式 / Set number formats
        If commaColIdx > 0 Then
            .Columns(commaColIdx).NumberFormat = "#,##0"
        End If
        If pctColIdx > 0 Then
            .Columns(pctColIdx).NumberFormat = "0.00%"
        End If
        
        ' 列宽自适应，冻结首行 / Auto-fit columns, freeze first row
        .Columns.AutoFit
        With ActiveWindow
            .FreezePanes = False
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
    End With
    
    On Error GoTo 0
    Call WriteLog(MODULE_NAME, "FormatSheetStandard", "工作表已格式化 / Worksheet formatted: " & ws.Name, "系统工具")
End Sub

' ==============================================================================================
' SECTION 7: 全局错误处理 / Global Error Handling
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] HandleError
' 说明：集中式全局错误处理，恢复Excel环境、输出错误信息、写入日志文件、可选弹出提示框
'       Centralized error handling: restore Excel environment, log error, optionally show message box
' 参数：source - String，错误来源（模块.过程名） / Error source (Module.Procedure)
'       message - String，错误描述 / Error description
'       errNumber - Long（可选），错误代码 / Error number
' ----------------------------------------------------------------------------------------------
Public Sub HandleError(ByVal source As String, ByVal message As String, Optional ByVal errNumber As Long = 0)
    ' 恢复Excel基础环境 / Restore Excel environment
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
    ' 拼接完整错误信息 / Build full error message
    Dim fullMessage As String
    If errNumber <> 0 Then
        fullMessage = "错误代码 / Error #" & errNumber & ": 发生在 / Occurred in [" & source & "]" & vbCrLf & _
                      vbCrLf & "错误描述 / Description: " & message
    Else
        fullMessage = "发生在 / Occurred in [" & source & "]" & vbCrLf & vbCrLf & "错误描述 / Description: " & message
    End If
    
    ' 写入日志 / Write to log
    Call WriteLog(MODULE_NAME, "HandleError", fullMessage, "错误信息")
    
    ' 按配置弹出错误提示框 / Show message if configured
    If ShowError Then
        MsgBox fullMessage, vbCritical, "再保险引擎 - 运行错误 / Reinsurance Engine - Runtime Error"
    End If
    
    ' 写入错误信息到日志文件 / Write to error log file
    On Error Resume Next
    Dim logFile As String
    logFile = ThisWorkbook.path & "\error_log.txt"
    Open logFile For Append As #1
    Print #1, Now & " - " & fullMessage & vbCrLf & "-----------------------------------------"
    Close #1
    On Error GoTo 0
End Sub

' ==============================================================================================
' SECTION 8: 通用核心工具函数 / Common Utility Functions
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IsWorkSheetExist
' 说明：检查指定工作簿中是否存在目标名称的工作表
'       Check if a worksheet exists in the specified workbook
' 参数：wb - Workbook，待检查的工作簿对象 / Workbook to check
'       sheetName - String，目标工作表名称 / Worksheet name
' 返回值：Boolean - True=工作表存在 / Worksheet exists
' ----------------------------------------------------------------------------------------------
Public Function IsWorkSheetExist(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    IsWorkSheetExist = Not ws Is Nothing
    On Error GoTo 0
End Function

' ==============================================================================================
' SECTION 9: 全局日志系统 / Global Logging System
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] WriteLog
' 说明：全局标准化日志写入方法，所有模块统一调用
'       Standardized logging function for all modules
' 参数：moduleName - String，模块名 / Module name
'       procName   - String，过程/函数名 / Procedure name
'       logContent - String，日志具体内容 / Log content
'       logTitle   - String（可选），第二列标题行内容，用于分类标注 / Optional title for classification
' ----------------------------------------------------------------------------------------------
Public Sub WriteLog(ByVal moduleName As String, ByVal procName As String, _
                    ByVal logContent As String, Optional ByVal logTitle As String = "")
    On Error Resume Next
    
    Dim logSwitch As Boolean      ' 日志开关（是否写入Excel）/ Log switch
    Dim logStartCell As Range     ' 日志起始单元格（Setup!E2）/ Log start cell
    Dim lastLogRow As Long        ' 日志最后一行行号 / Last log row
    Dim logFullText As String     ' 格式化后的完整日志内容 / Formatted log text
    
    ' 1. 读取日志开关 / Read log switch from Main worksheet
    logSwitch = Main.Range("p_show_log").value
    
    ' 2. 定义日志起始单元格 / Define log start cell
    Set logStartCell = Log.Range("B3")
    
    ' 3. 构造标准化日志内容 / Build standardized log text
    logFullText = "[" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "] [" & moduleName & "." & procName & "] " & logContent
    
    ' 4. 必输：Debug窗口输出 / Always output to Debug window
    Debug.Print logFullText
    
    ' 5. 可选：写入Excel Setup表 / Optionally write to Excel Setup sheet
    If logSwitch Then
        lastLogRow = Log.Cells(Log.rows.count, logStartCell.Column).End(xlUp).row
        If lastLogRow < logStartCell.row Then lastLogRow = logStartCell.row
        
        Log.Cells(lastLogRow, logStartCell.Column).value = logFullText
        Log.Cells(lastLogRow, logStartCell.Column + 1).value = logTitle
    End If
    
    Set logStartCell = Nothing
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ClearLog
' 说明：清空Setup工作表日志区域（E2:F列）的所有内容
'       Clear all log content in Setup worksheet (columns E2:F)
' ----------------------------------------------------------------------------------------------
Public Sub ClearLog()
    On Error Resume Next
    Log.Range("B3:C" & Log.rows.count).ClearContents
    MsgBox "Log表日志区域（B3:C列）已清空！/ Log area (B3:C) cleared!", vbInformation, "日志清空完成 / Log Clear Complete"
    Call WriteLog(MODULE_NAME, "ClearLog", "日志区域已清空 / Log area cleared", "日志系统")
End Sub

