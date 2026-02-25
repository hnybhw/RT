Attribute VB_Name = "app_07_engine"
' ==============================================================================================
' MODULE NAME     : app_07_engine
' PURPOSE         : Python引擎调用、批处理与运行环境校验核心模块，提供单分段执行、全分段批量处理、进度显示、运行统计等功能
'                 : Core module for Python engine invocation, batch processing and environment validation:
'                 : single segment execution, full segment batch processing, progress display, execution statistics
' DEPENDS         : app_01_basic (Main/Setup/GetFilePath/HandleError/MessageStart/MessageEnd/IsWorkSheetExist/WriteLog)
'                 : app_06_ws (CommitProcess/RefreshOutputSheetList/SmartCleanUp/g_COMMIT_SILENT)
'                 : xlwings (Python与Excel交互核心库 / Python-Excel interaction library)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'   [C] 批处理相关常量         - 定义列宽、进度条长度、最大分段数等 / Batch processing constants
'
' SECTION 2: 单分段执行 / Single Segment Execution
'   [S] SpecifyPyPath         - 打开文件对话框选择Python脚本路径，更新Main工作表配置
'                              / Open file dialog to select Python script path, update Main worksheet
'   [S] Run_Python_Engine     - 通过xlwings触发Python计算，支持静默模式（用于批量处理）
'                              / Trigger Python calculation via xlwings, support silent mode
'   [f] IsXlwingsReady        - 校验xlwings运行环境是否可用，报错时提供完整配置步骤
'                              / Check if xlwings environment is ready, provide setup instructions on error
'   [f] ValidatePythonPath    - 验证Python脚本路径是否存在 / Validate Python script path
'   [f] BuildPythonCommand    - 构建安全的Python命令字符串，处理路径转义 / Build safe Python command with path escaping
'
' SECTION 3: 批处理与统计 / Batch Processing & Statistics
'   [S] Run_All_Segments      - 批量处理所有分段，自动静默模式，生成执行报告并记录到Setup表
'                              / Batch process all segments, auto silent mode, generate execution report to Setup sheet
'   [f] GetSegmentList        - 获取分段列表命名区域 / Get segment list named range
'   [f] ConfirmBatchStart     - 显示批量启动确认对话框 / Show batch start confirmation dialog
'   [f] InitializeLogArea     - 初始化Setup表日志区域 / Initialize log area in Setup sheet
'   [f] BuildReportHeader     - 构建执行报告表头 / Build execution report header
'   [f] ProcessSingleSegment  - 处理单个分段（静默模式）/ Process single segment in silent mode
'   [f] LogSegmentDuration    - 记录分段执行耗时到Setup表 / Log segment duration to Setup sheet
'   [f] FinalizeBatch         - 完成批量处理，显示汇总报告 / Finalize batch processing, show summary report
'   [f] PadString             - 字符串填充，用于对齐显示 / String padding for aligned display
'   [f] GetProgressBar        - 生成进度条字符串（■/□）用于状态栏显示 / Generate progress bar for status bar
'
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 1: 模块常量声明 / Module Constants
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 模块名称常量 - 用于日志记录，避免硬编码 / Module name constant for logging
' ----------------------------------------------------------------------------------------------
Public Const MODULE_NAME As String = "app_07_engine"

' ----------------------------------------------------------------------------------------------
' [C] 批处理相关常量 / Batch processing constants
' ----------------------------------------------------------------------------------------------
Private Const DEFAULT_COL_WIDTH      As Integer = 22     ' 默认列宽 / Default column width
Private Const PROGRESS_BAR_LENGTH    As Integer = 20     ' 进度条长度 / Progress bar length
Private Const MAX_SEGMENTS           As Integer = 100    ' 最大分段数 / Maximum segments

' ==============================================================================================
' SECTION 2: 单分段执行 / Single Segment Execution
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] SpecifyPyPath
' 说明：打开文件对话框供用户选择Python脚本，选择后更新Main工作表的路径配置单元格
'       Open file dialog for Python script selection, update Main worksheet configuration
' ----------------------------------------------------------------------------------------------
Public Sub SpecifyPyPath()
    On Error GoTo ErrorHandler
    
    Dim selectedPath As String
    selectedPath = GetFilePath("选择Python再保险脚本 / Select Python Reinsurance Script", _
                               "Python脚本 / Python Scripts, *.py")
    
    If selectedPath <> "" Then
        Main.Range("ref_Py_ScriptPath").value = selectedPath
        Call WriteLog(MODULE_NAME, "SpecifyPyPath", "Python脚本路径已更新 / Python script path updated: " & selectedPath, "路径配置")
        
        ' 非静默模式才显示弹窗 / Show message only in non-silent mode
        If Not app_06_ws.g_COMMIT_SILENT Then
            MsgBox "Python脚本路径已更新！" & vbCrLf & "Python script path updated!", _
                   vbInformation, "路径配置成功 / Path Configuration Successful"
        End If
    Else
        Call WriteLog(MODULE_NAME, "SpecifyPyPath", "用户取消选择 / User cancelled selection", "路径配置")
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME & ".SpecifyPyPath", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [S] Run_Python_Engine
' 说明：通过xlwings触发Python计算，支持静默模式，用于批量处理时关闭消息弹窗
'       Trigger Python calculation via xlwings, support silent mode for batch processing
' 参数：isSilent - Boolean（可选），是否静默模式（不显示MessageEnd弹窗），默认False
'                / Silent mode (no MessageEnd popup), default False
' ----------------------------------------------------------------------------------------------
Public Sub Run_Python_Engine(Optional ByVal isSilent As Boolean = False)
    On Error GoTo ErrHandler
    
    ' 开启xlwings环境校验，报错含完整配置步骤 / Check xlwings environment
    If Not IsXlwingsReady() Then Exit Sub
    
    ' 非静默模式启动计时 / Start timing in non-silent mode
    Dim startTime As Double
    If Not isSilent Then
        startTime = MessageStart("Python引擎 / Python Engine")
    End If
    
    Call WriteLog(MODULE_NAME, "Run_Python_Engine", "开始执行Python引擎 / Starting Python engine", "Python调用")
    
    ' 从Main工作表读取参数 / Read parameters from Main worksheet
    Dim pyScriptPath As String
    Dim targetSegment As String
    
    pyScriptPath = Main.Range("ref_Py_ScriptPath").value
    targetSegment = Main.Range("ref_10K_Segment").value
    
    Call WriteLog(MODULE_NAME, "Run_Python_Engine", "脚本路径 / Script: " & pyScriptPath & _
                  ", 分段 / Segment: " & targetSegment, "Python调用")
    
    ' 验证Python脚本路径 / Validate Python script path
    If Not ValidatePythonPath(pyScriptPath) Then Exit Sub
    
    ' 构建并执行Python命令 / Build and execute Python command
    Dim pyCmd As String
    pyCmd = BuildPythonCommand(pyScriptPath)
    
    ' 更新状态栏 / Update status bar
    Application.StatusBar = ">>> 正在执行Python / Running Python for segment: " & targetSegment
    Call WriteLog(MODULE_NAME, "Run_Python_Engine", "执行Python命令 / Executing Python command", "Python调用")
    
    ' 执行Python / Execute Python
    Application.Run "RunPython", pyCmd
    
    Call WriteLog(MODULE_NAME, "Run_Python_Engine", "Python执行完成 / Python execution completed", "Python调用")

    ' 刷新UI列表（新建的工作表）/ Refresh UI list for newly created sheets
    Call app_06_ws.RefreshOutputSheetList
    
    ' 非静默模式结束计时 / End timing in non-silent mode
    If Not isSilent Then
        MessageEnd startTime, "Python引擎 / Python Engine"
    Else
        Application.StatusBar = False
    End If
    
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Call WriteLog(MODULE_NAME, "Run_Python_Engine", "Python执行失败 / Python execution failed: " & Err.Description, "错误")
    HandleError MODULE_NAME & ".Run_Python_Engine", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [f] IsXlwingsReady
' 说明：校验xlwings运行环境是否可用，通过尝试执行最小Python命令验证
'       Check if xlwings environment is ready by executing minimal Python command
' 返回值：Boolean - True=环境可用，False=环境不可用 / True if ready, False otherwise
' 备注：报错时提供完整的xlwings配置步骤说明 / Provide complete setup instructions on error
' ----------------------------------------------------------------------------------------------
Private Function IsXlwingsReady() As Boolean
    On Error GoTo ErrorHandler
    
    ' 尝试执行最小Python命令 / Attempt minimal Python command
    Application.Run "RunPython", "print('xlwings test')"
    
    ' 成功则返回True / Return True if successful
    IsXlwingsReady = True
    Exit Function
    
ErrorHandler:
    ' 环境不可用 / Environment not ready
    IsXlwingsReady = False
    
    ' 完整的报错提示，仅非静默模式弹窗 / Complete error message, only in non-silent mode
    If Not app_06_ws.g_COMMIT_SILENT Then
        Dim errMsg As String
        errMsg = "xlwings环境校验失败！请按以下步骤配置：" & vbCrLf & vbCrLf
        errMsg = errMsg & "【1】安装Python + xlwings" & vbCrLf
        errMsg = errMsg & "1. 安装Python并勾选'Add Python to PATH'" & vbCrLf
        errMsg = errMsg & "2. 打开命令提示符执行：pip install xlwings" & vbCrLf & vbCrLf
        errMsg = errMsg & "【2】添加Excel VBA引用" & vbCrLf
        errMsg = errMsg & "1. 按 Alt+F11 打开VBA编辑器" & vbCrLf
        errMsg = errMsg & "2. 点击菜单 工具 → 引用" & vbCrLf
        errMsg = errMsg & "3. 勾选或浏览添加 xlwings.bas 文件" & vbCrLf & vbCrLf
        errMsg = errMsg & "【3】配置Excel宏安全" & vbCrLf
        errMsg = errMsg & "1. 文件 → 选项 → 信任中心 → 信任中心设置" & vbCrLf
        errMsg = errMsg & "2. 宏设置：启用所有宏" & vbCrLf
        errMsg = errMsg & "3. 勾选'信任对VBA工程对象模型的访问'" & vbCrLf
        errMsg = errMsg & "4. 外部内容：启用所有数据连接" & vbCrLf
        errMsg = errMsg & "5. 重启Excel使设置生效"
        
        MsgBox errMsg, vbCritical, "xlwings配置错误 / xlwings Configuration Error"
    End If
    
    ' 记录错误日志 / Log error
    Call WriteLog(MODULE_NAME, "IsXlwingsReady", "xlwings环境校验失败 / xlwings check failed: " & Err.Description, "错误")
End Function

' ----------------------------------------------------------------------------------------------
' [f] ValidatePythonPath
' 说明：验证Python脚本路径是否存在，为空或不存在时显示错误
'       Validate Python script path exists
' 参数：pyScriptPath - String，Python脚本路径 / Python script path
' 返回值：Boolean - True=路径有效，False=路径无效 / True if valid, False otherwise
' ----------------------------------------------------------------------------------------------
Private Function ValidatePythonPath(ByVal pyScriptPath As String) As Boolean
    ' 空路径检查 / Empty path check
    If pyScriptPath = "" Then
        If Not app_06_ws.g_COMMIT_SILENT Then
            MsgBox "Python脚本路径为空，请先指定路径！" & vbCrLf & _
                   "Python script path is empty, please specify first!", _
                   vbCritical, "路径错误 / Path Error"
        End If
        Call WriteLog(MODULE_NAME, "ValidatePythonPath", "Python脚本路径为空 / Empty path", "错误")
        ValidatePythonPath = False
        Exit Function
    End If
    
    ' 文件存在性检查 / File existence check
    If Dir(pyScriptPath) = "" Then
        If Not app_06_ws.g_COMMIT_SILENT Then
            MsgBox "Python脚本不存在：" & pyScriptPath & vbCrLf & _
                   "Python script not found: " & pyScriptPath, _
                   vbCritical, "文件错误 / File Error"
        End If
        Call WriteLog(MODULE_NAME, "ValidatePythonPath", "Python脚本不存在 / File not found: " & pyScriptPath, "错误")
        ValidatePythonPath = False
        Exit Function
    End If
    
    ValidatePythonPath = True
End Function

' ----------------------------------------------------------------------------------------------
' [f] BuildPythonCommand
' 说明：构建安全的Python命令字符串，处理路径中的特殊字符（如单引号、反斜杠）
'       Build safe Python command string, handle special characters in path
' 参数：pyScriptPath - String，Python脚本路径 / Python script path
' 返回值：String - 格式化的Python命令字符串 / Formatted Python command string
' ----------------------------------------------------------------------------------------------
Private Function BuildPythonCommand(ByVal pyScriptPath As String) As String
    Dim pyModuleName As String
    pyModuleName = GetFileNameWithoutExt(pyScriptPath)
    
    ' 转义路径中的特殊字符 / Escape special characters in path
    Dim safePath As String
    safePath = Replace(pyScriptPath, "'", "\'")   ' 转义单引号 / Escape single quotes
    safePath = Replace(safePath, "/", "\")        ' 统一反斜杠 / Unify backslashes
    
    BuildPythonCommand = "import sys, os; " & _
                        "path = r'" & safePath & "'; " & _
                        "sys.path.append(os.path.dirname(path)); " & _
                        "engine = __import__('" & pyModuleName & "'); " & _
                        "engine.main()"
    
    Call WriteLog(MODULE_NAME, "BuildPythonCommand", "Python命令构建完成 / Python command built: " & pyModuleName, "Python调用")
End Function

' ==============================================================================================
' SECTION 3: 批处理与统计 / Batch Processing & Statistics
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Run_All_Segments
' 说明：批量处理所有分段，自动启用静默模式，生成执行报告并记录到Setup表
'       Batch process all segments, auto silent mode, generate execution report to Setup sheet
' ----------------------------------------------------------------------------------------------
Public Sub Run_All_Segments()
    ' 声明变量 / Declare variables
    Dim originalSilentState As Boolean
    Dim segCell As Range, segList As Range
    Dim totalSegs As Integer, currentIdx As Integer
    Dim tStart As Double
    Dim logMessage As String
    Dim confirmMsg As VbMsgBoxResult
    
    On Error GoTo ErrHandler
    
    Call WriteLog(MODULE_NAME, "Run_All_Segments", "开始批量执行所有分段 / Starting batch execution of all segments", "批量执行")
    
    ' 保存原始静默状态，强制启用静默模式 / Save original silent state, force silent mode
    originalSilentState = app_06_ws.g_COMMIT_SILENT
    app_06_ws.g_COMMIT_SILENT = True
    
    ' 1. 获取分段列表 / Get segment list
    Set segList = GetSegmentList()
    If segList Is Nothing Then GoTo ExitSub
    
    ' 2. 用户确认 / User confirmation
    If Not ConfirmBatchStart() Then
        Call WriteLog(MODULE_NAME, "Run_All_Segments", "用户取消批量执行 / User cancelled batch execution", "批量执行")
        GoTo ExitSub
    End If
    
    ' 3. 初始化 / Initialize
    totalSegs = Application.WorksheetFunction.CountA(segList)
    currentIdx = 0
    tStart = Timer
    
    Call WriteLog(MODULE_NAME, "Run_All_Segments", "待处理分段数量 / Segments to process: " & totalSegs, "批量执行")
    
    ' 4. 清空并初始化日志区域 / Clear and initialize log area
    InitializeLogArea
    
    ' 5. 构建报告表头 / Build report header
    logMessage = BuildReportHeader(DEFAULT_COL_WIDTH)
    
    ' 6. 批量处理主循环 / Batch processing main loop
    Application.ScreenUpdating = False
    
    For Each segCell In segList
        If segCell.value <> "" Then
            currentIdx = currentIdx + 1
            ProcessSingleSegment segCell.value, currentIdx, totalSegs, logMessage
        End If
    Next segCell
    
    ' 7. 完成批量处理 / Finalize batch
    FinalizeBatch tStart, totalSegs, currentIdx, logMessage, DEFAULT_COL_WIDTH
    
ExitSub:
    ' 恢复原始静默状态 / Restore original silent state
    app_06_ws.g_COMMIT_SILENT = originalSilentState
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call WriteLog(MODULE_NAME, "Run_All_Segments", "批量执行出错 / Batch execution error: " & Err.Description, "错误")
    HandleError MODULE_NAME & ".Run_All_Segments", Err.Description
    GoTo ExitSub
End Sub

' ----------------------------------------------------------------------------------------------
' [f] GetSegmentList
' 说明：获取分段列表命名区域 rng_segment_list
'       Get segment list named range
' 返回值：Range - 分段列表区域，未找到则返回Nothing / Segment list range, Nothing if not found
' ----------------------------------------------------------------------------------------------
Private Function GetSegmentList() As Range
    On Error Resume Next
    Set GetSegmentList = Main.Range("rng_segment_list")
    On Error GoTo 0
    
    If GetSegmentList Is Nothing Then
        Call WriteLog(MODULE_NAME, "GetSegmentList", "命名区域 'rng_segment_list' 未找到 / Named range not found", "错误")
        If Not app_06_ws.g_COMMIT_SILENT Then
            MsgBox "错误：找不到命名区域 'rng_segment_list'，请检查名称管理器。" & vbCrLf & _
                   "Error: Named range 'rng_segment_list' not found. Please check Name Manager.", _
                   vbCritical, "配置错误 / Configuration Error"
        End If
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] ConfirmBatchStart
' 说明：显示批量启动确认对话框
'       Show batch start confirmation dialog
' 返回值：Boolean - True=用户确认，False=用户取消 / True if confirmed, False if cancelled
' ----------------------------------------------------------------------------------------------
Private Function ConfirmBatchStart() As Boolean
    Dim confirmMsg As VbMsgBoxResult
    confirmMsg = MsgBox("是否启动全自动批量处理（所有分段）？" & vbCrLf & _
                       "Start fully automated batch processing (All Segments)?", _
                       vbQuestion + vbYesNo, "批量启动 / Batch Start")
    ConfirmBatchStart = (confirmMsg = vbYes)
End Function

' ----------------------------------------------------------------------------------------------
' [f] InitializeLogArea
' 说明：初始化Setup表日志区域（B2:C列），清空旧数据并设置表头
'       Initialize log area in Setup sheet (B2:C), clear old data and set headers
' ----------------------------------------------------------------------------------------------
Private Sub InitializeLogArea()
    Setup.Range("B2:C100").ClearContents
    Setup.Range("B2").value = "分段名称 / Segment Name"
    Setup.Range("C2").value = "耗时 / Duration"
    Call WriteLog(MODULE_NAME, "InitializeLogArea", "日志区域初始化完成 / Log area initialized", "批量执行")
End Sub

' ----------------------------------------------------------------------------------------------
' [f] BuildReportHeader
' 说明：构建执行报告表头字符串
'       Build execution report header string
' 参数：colWidth - Integer，列宽 / Column width
' 返回值：String - 格式化的报告表头 / Formatted report header
' ----------------------------------------------------------------------------------------------
Private Function BuildReportHeader(ByVal colWidth As Integer) As String
    Dim headerWidth As Integer
    headerWidth = colWidth + 15
    
    BuildReportHeader = "批量执行报告 / BATCH EXECUTION REPORT" & vbCrLf & _
                        String(headerWidth, "=") & vbCrLf & _
                        PadString("分段名称 / Segment Name", colWidth, True) & " | 耗时 / Duration" & vbCrLf & _
                        String(headerWidth, "-") & vbCrLf
End Function

' ----------------------------------------------------------------------------------------------
' [f] ProcessSingleSegment
' 说明：处理单个分段（静默模式），更新状态栏和日志
'       Process single segment in silent mode, update status bar and log
' 参数：segName - String，分段名称 / Segment name
'       currentIdx - Integer，当前索引 / Current index
'       totalSegs - Integer，总分段数 / Total segments
'       logMessage - String，日志消息（ByRef）/ Log message
' ----------------------------------------------------------------------------------------------
Private Sub ProcessSingleSegment(ByVal segName As String, ByVal currentIdx As Integer, _
                                  ByVal totalSegs As Integer, ByRef logMessage As String)
    Dim tSeg As Double
    tSeg = Timer
    
    ' 更新状态栏 / Update status bar
    Application.StatusBar = GetProgressBar(currentIdx, totalSegs) & " 处理中 / Processing: " & segName
    
    Call WriteLog(MODULE_NAME, "ProcessSingleSegment", "开始处理分段 / Starting segment: " & segName & _
                  " (" & currentIdx & "/" & totalSegs & ")", "批量执行")
    
    ' 设置当前分段 / Set current segment
    Main.Range("ref_10K_Segment").value = segName
    
    ' 执行Python引擎（静默模式）/ Run Python engine (silent mode)
    Call Run_Python_Engine(isSilent:=True)
    
    ' 执行数据追加与清理 / Execute data append and cleanup
    Call app_06_ws.CommitProcess
    
    ' 计算耗时 / Calculate duration
    Dim timeStr As String
    timeStr = Format(Timer - tSeg, "0.00") & " s"
    
    ' 更新日志 / Update log
    logMessage = logMessage & PadString(segName, DEFAULT_COL_WIDTH, False) & " | " & timeStr & vbCrLf
    
    ' 记录到立即窗口和日志 / Log to immediate window and log file
    Debug.Print PadString(segName, DEFAULT_COL_WIDTH, False) & " | " & timeStr
    Call WriteLog(MODULE_NAME, "ProcessSingleSegment", PadString(segName, DEFAULT_COL_WIDTH, False) & " | " & timeStr, "批量执行")
    
    ' 记录到Setup表 / Log to Setup sheet
    LogSegmentDuration segName, timeStr, currentIdx
End Sub

' ----------------------------------------------------------------------------------------------
' [f] LogSegmentDuration
' 说明：记录分段执行耗时到Setup表
'       Log segment duration to Setup sheet
' 参数：segName - String，分段名称 / Segment name
'       timeStr - String，耗时字符串 / Duration string
'       idx - Integer，索引 / Index
' ----------------------------------------------------------------------------------------------
Private Sub LogSegmentDuration(ByVal segName As String, ByVal timeStr As String, ByVal idx As Integer)
    Setup.Range("B" & 2 + idx).value = segName
    Setup.Range("C" & 2 + idx).value = timeStr
End Sub

' ----------------------------------------------------------------------------------------------
' [f] FinalizeBatch
' 说明：完成批量处理，显示汇总报告
'       Finalize batch processing, show summary report
' 参数：startTime - Double，开始时间 / Start time
'       totalSegs - Integer，总分段数 / Total segments
'       processedCount - Integer，已处理分段数 / Processed segments
'       logMessage - String，日志消息 / Log message
'       colWidth - Integer，列宽 / Column width
' ----------------------------------------------------------------------------------------------
Private Sub FinalizeBatch(ByVal startTime As Double, ByVal totalSegs As Integer, _
                          ByVal processedCount As Integer, ByVal logMessage As String, _
                          ByVal colWidth As Integer)
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Dim totalTimeStr As String
    totalTimeStr = Format((Timer - startTime) / 60, "0.00") & " 分钟 / mins"
    
    ' 记录总耗时到Setup表 / Log total time to Setup sheet
    Setup.Range("B" & 3 + processedCount).value = "总耗时 / TOTAL TIME"
    Setup.Range("C" & 3 + processedCount).value = totalTimeStr
    
    ' 完成报告 / Complete report
    Dim headerWidth As Integer
    headerWidth = colWidth + 15
    
    logMessage = logMessage & String(headerWidth, "=") & vbCrLf & _
                 PadString("总耗时 / TOTAL TIME:", colWidth, False) & " | " & totalTimeStr
    
    Application.StatusBar = ">>>>>>>>>> 批量处理完成！总耗时 / Batch completed! Total time: " & totalTimeStr
    
    Call WriteLog(MODULE_NAME, "FinalizeBatch", "批量处理完成，总耗时 / Batch completed, total time: " & totalTimeStr, "批量执行")
    
    ' 显示报告对话框 / Show report dialog
    MsgBox logMessage, vbInformation, "批量执行完成 / Batch Execution Complete"
End Sub

' ----------------------------------------------------------------------------------------------
' [f] PadString
' 说明：字符串填充，用于对齐显示（简化版，保持原有视觉效果）
'       String padding for aligned display (simplified, keeps original visual effect)
' 参数：text - String，输入字符串 / Input string
'       targetWidth - Integer，目标宽度 / Target width
'       isHeader - Boolean，是否为表头（不截断）/ Is header (no truncation)
' 返回值：String - 填充后的字符串 / Padded string
' ----------------------------------------------------------------------------------------------
Private Function PadString(ByVal text As String, ByVal targetWidth As Integer, _
                           Optional ByVal isHeader As Boolean = False) As String
    Dim textLength As Integer
    textLength = Len(text)
    
    ' 超长则截断 / Truncate if too long
    If textLength > targetWidth Then
        If isHeader Then
            PadString = text
        Else
            PadString = Left(text, targetWidth - 3) & "..."
        End If
        Exit Function
    End If
    
    ' 简单空格填充 / Simple space padding
    Dim spacesToAdd As Integer
    spacesToAdd = targetWidth - textLength
    
    PadString = text & Space(spacesToAdd)
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetProgressBar
' 说明：生成进度条字符串（■/□）用于状态栏显示
'       Generate progress bar string for status bar display
' 参数：current - Integer，当前进度 / Current progress
'       total - Integer，总进度 / Total
' 返回值：String - 进度条字符串 / Progress bar string
' ----------------------------------------------------------------------------------------------
Private Function GetProgressBar(ByVal current As Integer, ByVal total As Integer) As String
    Dim pct As Double
    Dim filledLen As Integer
    
    pct = current / total
    filledLen = Int(pct * PROGRESS_BAR_LENGTH)
    
    GetProgressBar = "[" & String(filledLen, "■") & _
                    String(PROGRESS_BAR_LENGTH - filledLen, "□") & "] " & _
                    Format(pct, "0%")
End Function

