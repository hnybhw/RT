VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C10KProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : C10KProcessor
' PURPOSE         : 封装万年模拟数据的所有处理操作，包含数据加载、格式转换、聚合统计、资源管理等核心能力
'                 : Encapsulates all 10K simulation data processing operations: data loading,
'                 : format conversion, aggregation statistics, resource management
' DEPENDS         : app_01_basic (HandleError, IsWorkSheetExist, WriteLog 等基础工具函数)
'                 : app_02_wb (ValidateFilePath 文件路径校验函数)
'                 : CParserFactory (解析器工厂类，创建分段专属解析器)
'                 : IDataParser (数据解析接口，定义解析器标准方法)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 常量与私有成员变量声明 / Constants & Private Member Variables
'   [C] 列索引常量               - 定义输出数组的固定列索引 / Output array column indices
'   [V] 私有成员变量             - 类的核心数据载体、状态标记、对象引用 / Core data, state flags, object references
'
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
'   [S] Class_Initialize        - 类初始化，创建字典/工厂对象，设置默认配置 / Class initialization
'   [S] Class_Terminate         - 类销毁，自动调用清理方法释放所有资源 / Class termination
'
' SECTION 3: 公共属性暴露 / Public Properties
'   [F] IsLoaded                - 只读，返回数据是否成功加载的状态 / Data loaded status
'   [F] SourcePath              - 只读，返回加载的源文件路径 / Source file path
'   [F] RowCount                - 只读，返回加载矩阵的行数 / Matrix row count
'   [F] ColumnCount             - 只读，返回加载矩阵的列数 / Matrix column count
'   [F] Result                  - 只读，返回最近一次处理的结果数组 / Last result array
'   [F] Config (Get)            - 只读，根据键获取配置字典中的值 / Get config value by key
'   [F] Config (Let)            - 可写，根据键设置配置字典中的值 / Set config value by key
'
' SECTION 4: 数据加载与资源清理 / Data Loading & Resource Cleanup
'   [F] LoadFromFile            - 从指定文件加载指定分段的万年数据，核心数据入口 / Load data from file and segment
'   [F] LoadFromMatrix          - 直接加载数组数据，用于测试或内部数据传递 / Load from array for testing
'   [F] GetResult               - 返回最近一次处理的结果数组（方法形式，保持兼容性）/ Get result array (method for compatibility)
'   [S] Cleanup                 - 公共清理方法，释放工作簿/解析器/数组等所有资源 / Public cleanup method
'
' SECTION 5: 核心数据处理 / Core Data Processing
'   [F] TransformToLongList     - 宽表转长表核心逻辑，支持大数据量分块处理判断 / Wide-to-long conversion
'   [F] AggregateBySubLoB       - 按SubLoB维度聚合数据，支持LoB配置匹配的容错/严格模式 / SubLoB aggregation
'   [f] TransformToLongListChunked - 分块处理占位方法，适配超大数据量 / Chunked processing placeholder
'
' SECTION 6: 私有辅助工具方法 / Private Helper Methods
'   [f] ValidateMatrix          - 验证加载的矩阵数据是否为有效的二维数组 / Validate matrix data
'   [f] GetArrayDimensions      - 获取任意数组的维度数量 / Get array dimensions
'   [f] ExtractUniqueSubLoBs    - 从矩阵表头行提取唯一的SubLoB名称 / Extract unique SubLoB names
'   [f] LoadLoBParameters       - 从Setup工作表加载LoB配置 / Load LoB parameters from Setup
'
' SECTION 7: 调试与信息输出 / Debugging & Information
'   [S] DumpInfo                - 输出类实例的核心状态信息到立即窗口和日志 / Dump instance info
'
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 1: 常量与私有成员变量声明 / Constants & Private Member Variables
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 列索引常量 - 定义输出数组的固定列索引，统一列结构规范 / Output array column indices
' ----------------------------------------------------------------------------------------------
Private Const COL_LOSS_SEQ As Long = 1      ' 损失序号 / Loss Sequence
Private Const COL_COLUMN_ID As Long = 2     ' 列ID / Column ID
Private Const COL_YEAR As Long = 3          ' 年份 / Year
Private Const COL_SUBLOB As Long = 4        ' 次级业务线 / Sub LoB
Private Const COL_PERIL As Long = 5         ' 风险类型 / Peril
Private Const COL_EVENT As Long = 6         ' 事件类型 / Event
Private Const COL_LOSS_AMT As Long = 7      ' 损失金额 / Loss Amount
Private Const COL_TID As Long = 8           ' TID（预留）/ TID (reserved)
Private Const COL_GROUP As Long = 9         ' 分组 / Group

' ----------------------------------------------------------------------------------------------
' [V] 私有成员变量 - 类的核心数据载体、状态标记、外部对象引用 / Private member variables
' ----------------------------------------------------------------------------------------------
Private m_sourceWb As Workbook              ' 源工作簿对象引用 / Source workbook reference
Private m_filePath As String                ' 源文件的完整路径 / Source file path
Private m_matrix As Variant                 ' 加载的矩阵数据（二维数组）/ Loaded matrix data
Private m_result As Variant                 ' 处理后的结果数组 / Result array
Private m_config As Object                  ' 配置字典（晚绑定）/ Configuration dictionary
Private m_isInitialized As Boolean          ' 初始化状态标记 / Initialization flag
Private m_parserFactory As CParserFactory   ' 解析器工厂对象 / Parser factory
Private m_currentParser As IDataParser      ' 当前解析器对象 / Current parser

' ==============================================================================================
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Class_Initialize
' 说明：类实例创建时自动执行，完成对象初始化、默认配置设置、状态初始化
'       Auto-executed on instance creation: object initialization, default config, state init
' ----------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ' 晚绑定创建配置字典，避免引用依赖 / Late-bound dictionary to avoid reference dependency
    Set m_config = CreateObject("Scripting.Dictionary")
    
    ' 创建解析器工厂实例 / Create parser factory
    Set m_parserFactory = New CParserFactory
    
    ' 初始化默认配置项 / Initialize default configuration
    m_config("MaterialityThreshold") = 0.5      ' 损失重要性阈值 / Loss materiality threshold
    m_config("DefaultYearMax") = 10000          ' 默认最大统计年份 / Default max years
    m_config("MaxRows") = 1048576               ' Excel最大行数 / Excel max rows
    m_config("ChunkSize") = 2000000             ' 分块处理大小 / Chunk size
    m_config("ChunkThreshold") = 6000000        ' 分块处理阈值 / Chunk threshold
    m_config("ShowProgress") = True             ' 显示进度条 / Show progress bar
    
    ' 初始化类核心状态 / Initialize core state
    m_isInitialized = False
    
    Call WriteLog("C10KProcessor", "Class_Initialize", "类实例初始化完成 / Class instance initialized", "生命周期")
End Sub

' ----------------------------------------------------------------------------------------------
' [S] Class_Terminate
' 说明：类实例销毁时自动执行，调用清理方法释放所有占用的资源，避免内存泄漏
'       Auto-executed on instance termination: cleanup all resources to prevent memory leaks
' ----------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Cleanup
    Call WriteLog("C10KProcessor", "Class_Terminate", "类实例已销毁 / Class instance terminated", "生命周期")
End Sub

' ==============================================================================================
' SECTION 3: 公共属性暴露 / Public Properties
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IsLoaded (Get)
' 说明：只读属性，返回数据是否成功加载的状态（初始化完成且矩阵数据非空）
'       Read-only: data loaded status (initialized and matrix non-empty)
' 返回值：Boolean - True=数据已加载 / Data loaded, False=数据未加载 / No data
' ----------------------------------------------------------------------------------------------
Public Property Get IsLoaded() As Boolean
    IsLoaded = m_isInitialized And Not IsEmpty(m_matrix)
End Property

' ----------------------------------------------------------------------------------------------
' [F] SourcePath (Get)
' 说明：只读属性，返回最近一次加载的源文件完整路径
'       Read-only: source file path of last loaded data
' 返回值：String - 源文件路径 / Source file path
' ----------------------------------------------------------------------------------------------
Public Property Get SourcePath() As String
    SourcePath = m_filePath
End Property

' ----------------------------------------------------------------------------------------------
' [F] RowCount (Get)
' 说明：只读属性，返回加载矩阵数据的总行数
'       Read-only: total rows in loaded matrix
' 返回值：Long - 矩阵行数 / Matrix rows
' ----------------------------------------------------------------------------------------------
Public Property Get RowCount() As Long
    If IsLoaded Then
        RowCount = UBound(m_matrix, 1)
    Else
        RowCount = 0
    End If
End Property

' ----------------------------------------------------------------------------------------------
' [F] ColumnCount (Get)
' 说明：只读属性，返回加载矩阵数据的总列数
'       Read-only: total columns in loaded matrix
' 返回值：Long - 矩阵列数 / Matrix columns
' ----------------------------------------------------------------------------------------------
Public Property Get ColumnCount() As Long
    If IsLoaded Then
        ColumnCount = UBound(m_matrix, 2)
    Else
        ColumnCount = 0
    End If
End Property

' ----------------------------------------------------------------------------------------------
' [F] Result (Get)
' 说明：只读属性，返回最近一次数据处理的结果数组
'       Read-only: result array from last processing
' 返回值：Variant - 处理后的二维数组 / Result 2D array
' ----------------------------------------------------------------------------------------------
Public Property Get result() As Variant
    result = m_result
End Property

' ----------------------------------------------------------------------------------------------
' [F] Config (Get)
' 说明：只读属性，根据指定键获取配置字典中的值
'       Read-only: get configuration value by key
' 参数：key - String，配置项键名 / Configuration key
' 返回值：Variant - 配置项值 / Configuration value
' ----------------------------------------------------------------------------------------------
Public Property Get config(key As String) As Variant
    If m_config.Exists(key) Then
        config = m_config(key)
    Else
        config = Empty
    End If
End Property

' ----------------------------------------------------------------------------------------------
' [F] Config (Let)
' 说明：可写属性，根据指定键设置配置字典中的值
'       Writeable: set configuration value by key
' 参数：key - String，配置项键名 / Configuration key
'       value - Variant，配置项值 / Configuration value
' ----------------------------------------------------------------------------------------------
Public Property Let config(key As String, value As Variant)
    m_config(key) = value
End Property

' ==============================================================================================
' SECTION 4: 数据加载与资源清理 / Data Loading & Resource Cleanup
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] LoadFromFile
' 说明：从指定文件加载指定分段的万年模拟数据，类的核心数据入口方法
'       Core data entry method: load 10K data from file and segment
' 参数：filePath - String，源文件完整路径 / Source file path
'       segment - String，数据分段名称 / Segment name
' 返回值：Boolean - True=加载成功 / Success, False=加载失败 / Failure
' ----------------------------------------------------------------------------------------------
Public Function LoadFromFile(ByVal filePath As String, ByVal segment As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 错误码常量定义 / Error code constants
    Const ERR_INVALID_FILE_PATH As Long = 1001          ' 无效文件路径 / Invalid file path
    Const ERR_FAILED_OPEN_WB As Long = 1002             ' 工作簿打开失败 / Failed to open workbook
    Const ERR_SEGMENT_SHEET_NOT_FOUND As Long = 1003    ' 分段工作表不存在 / Segment sheet not found
    Const ERR_INVALID_MATRIX As Long = 1004             ' 矩阵数据无效 / Invalid matrix data
    Const ERR_NO_PARSER_FOR_SEGMENT As Long = 1005      ' 无对应分段解析器 / No parser for segment
    
    ' 声明局部变量 / Declare local variables
    Dim ws As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    
    Call WriteLog("C10KProcessor", "LoadFromFile", "开始加载数据 / Starting data load: 文件/File=" & filePath & ", 分段/Segment=" & segment, "数据加载")
    
    ' 1. 输入验证（守卫逻辑）/ Input validation
    If Not app_02_wb.ValidateFilePath(filePath) Then
        Err.Raise ERR_INVALID_FILE_PATH, "C10KProcessor.LoadFromFile", "无效文件路径 / Invalid file path: " & filePath
    End If
    
    ' 2. 清理旧数据 / Cleanup old data
    Cleanup
    
    ' 3. 保存原始Excel环境设置 / Save original Excel environment settings
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    
    ' 4. 关闭弹窗和事件，提升工作簿打开的稳定性 / Disable alerts and events for stability
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' 只读打开工作簿，避免文件锁定 / Open workbook read-only
    m_filePath = filePath
    Set m_sourceWb = Workbooks.Open(fileName:=filePath, readOnly:=True, UpdateLinks:=False)
    
    ' 校验工作簿是否成功打开 / Validate workbook opened
    If m_sourceWb Is Nothing Then
        Err.Raise ERR_FAILED_OPEN_WB, "C10KProcessor.LoadFromFile", "工作簿打开失败 / Failed to open workbook: " & filePath
    End If
    
    ' 5. 校验目标分段工作表是否存在 / Validate segment worksheet exists
    On Error Resume Next
    Set ws = m_sourceWb.Worksheets(segment)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Err.Raise ERR_SEGMENT_SHEET_NOT_FOUND, "C10KProcessor.LoadFromFile", "分段工作表不存在 / Segment sheet not found: " & segment
    End If
    Set ws = Nothing
    
    ' 6. 获取对应分段的解析器 / Get parser for segment
    Set m_currentParser = m_parserFactory.GetParser(segment)
    If m_currentParser Is Nothing Then
        Err.Raise ERR_NO_PARSER_FOR_SEGMENT, "C10KProcessor.LoadFromFile", "无对应分段解析器 / No parser for segment: " & segment
    End If
    
    ' 7. 加载数据到二维数组 / Load data to 2D array
    m_matrix = m_sourceWb.Worksheets(segment).Range("A1").CurrentRegion.value
    
    ' 8. 验证加载的矩阵数据 / Validate matrix data
    If Not ValidateMatrix() Then
        Err.Raise ERR_INVALID_MATRIX, "C10KProcessor.LoadFromFile", "分段矩阵数据无效 / Invalid matrix data for segment: " & segment
    End If
    
    ' 9. 恢复Excel环境设置，标记初始化完成 / Restore Excel environment, mark initialized
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    
    m_isInitialized = True
    
    Call WriteLog("C10KProcessor", "LoadFromFile", "数据加载成功 / Data loaded successfully: " & _
                  RowCount & "行 x " & ColumnCount & "列", "数据加载")
    
    LoadFromFile = True
    Exit Function

ErrorHandler:
    ' 优先恢复Excel环境设置 / Restore Excel environment first
    On Error Resume Next
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    On Error GoTo 0
    
    ' 增强错误信息 / Enhance error message
    Dim errMsg As String
    errMsg = Err.Description & " [文件/File: " & filePath & ", 分段/Segment: " & segment & "]"
    
    Call WriteLog("C10KProcessor", "LoadFromFile", "数据加载失败 / Data load failed: " & errMsg, "错误")
    HandleError "C10KProcessor.LoadFromFile", errMsg
    Cleanup
    LoadFromFile = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] LoadFromMatrix
' 说明：直接加载二维数组数据，用于测试场景或内部数据传递，跳过文件IO步骤
'       Direct load from 2D array for testing or internal data passing
' 参数：matrix - Variant，待加载的二维数组 / 2D array to load
' 返回值：Boolean - True=加载成功 / Success, False=加载失败 / Failure
' ----------------------------------------------------------------------------------------------
Public Function LoadFromMatrix(ByRef matrix As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 守卫逻辑，拦截空数组 / Guard clause for empty array
    If IsEmpty(matrix) Then
        Call WriteLog("C10KProcessor", "LoadFromMatrix", "输入数组为空 / Input array is empty", "警告")
        LoadFromMatrix = False
        Exit Function
    End If
    
    ' 清理旧数据，加载新数组 / Cleanup and load new array
    Cleanup
    m_matrix = matrix
    m_isInitialized = True
    
    Call WriteLog("C10KProcessor", "LoadFromMatrix", "数组加载成功 / Array loaded successfully", "数据加载")
    LoadFromMatrix = True
    Exit Function

ErrorHandler:
    Call WriteLog("C10KProcessor", "LoadFromMatrix", "数组加载失败 / Array load failed: " & Err.Description, "错误")
    HandleError "C10KProcessor.LoadFromMatrix", Err.Description
    LoadFromMatrix = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetResult
' 说明：返回最近一次数据处理的结果数组（方法形式，保持与app_04_tenk的兼容性）
'       Return result array from last processing (method for compatibility with app_04_tenk)
' 返回值：Variant - 处理后的二维数组，无结果时返回Empty / Result 2D array, Empty if no result
' ----------------------------------------------------------------------------------------------
Public Function GetResult() As Variant
    If IsEmpty(m_result) Then
        GetResult = Empty
        Call WriteLog("C10KProcessor", "GetResult", "结果数组为空 / Result array is empty", "警告")
    Else
        GetResult = m_result
        Call WriteLog("C10KProcessor", "GetResult", "返回结果数组 / Returning result array: " & _
                      UBound(m_result, 1) & "行 x " & UBound(m_result, 2) & "列", "结果输出")
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [S] Cleanup
' 说明：公共清理方法，释放所有占用的对象资源、清空数组、重置状态
'       Public cleanup method: release all resources, clear arrays, reset state
' ----------------------------------------------------------------------------------------------
Public Sub Cleanup()
    On Error Resume Next
    
    Call WriteLog("C10KProcessor", "Cleanup", "开始清理资源 / Starting resource cleanup", "资源管理")
    
    ' 清理源工作簿引用 - 显式关闭文件 / Release source workbook - explicitly close file
    If Not m_sourceWb Is Nothing Then
        ' 检查是否为有效文件路径，避免关闭本工作簿 / Check if valid path, avoid closing this workbook
        If m_sourceWb.path <> "" And m_sourceWb.FullName <> ThisWorkbook.FullName Then
            m_sourceWb.Close SaveChanges:=False  ' 显式关闭文件 / Explicitly close file
            Call WriteLog("C10KProcessor", "Cleanup", "源工作簿已关闭 / Source workbook closed: " & m_sourceWb.Name, "资源管理")
        End If
        Set m_sourceWb = Nothing
    End If
    
    ' 清理解析器对象引用 / Release parser reference
    Set m_currentParser = Nothing
    
    ' 清空矩阵和结果数组 / Clear arrays
    Erase m_matrix
    Erase m_result
    
    ' 重置文件路径和初始化状态 / Reset path and state
    m_filePath = ""
    m_isInitialized = False
    
    Call WriteLog("C10KProcessor", "Cleanup", "资源清理完成 / Resource cleanup completed", "资源管理")
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ClearRawData
' 说明：手动释放原始矩阵数据，为后续操作（如Python调用）腾出内存空间
'       执行此操作后，将无法再次使用 TransformToLongList 或 AggregateBySubLoB，
'       除非重新调用 LoadFromFile 加载数据
'       Manually release raw matrix data to free memory for subsequent operations (e.g., Python calls)
'       After this, TransformToLongList or AggregateBySubLoB cannot be used again
'       unless reloading data via LoadFromFile
' ----------------------------------------------------------------------------------------------
Public Sub ClearRawData()
    On Error Resume Next
    
    Call WriteLog("C10KProcessor", "ClearRawData", "开始释放原始矩阵数据 / Starting to release raw matrix data", "内存管理")
    
    ' 记录释放前的状态 / Log before release
    Dim rowsBefore As Long, colsBefore As Long
    If Not IsEmpty(m_matrix) And IsArray(m_matrix) Then
        On Error GoTo 0
        rowsBefore = UBound(m_matrix, 1)
        colsBefore = UBound(m_matrix, 2)
    End If
    
    ' 释放原始矩阵内存 / Release raw matrix memory
    Erase m_matrix
    
    ' 重置初始化状态（但保留结果数组）/ Reset initialization state (but keep result)
    m_isInitialized = False
    
    If rowsBefore > 0 Then
        Call WriteLog("C10KProcessor", "ClearRawData", "原始矩阵已释放 / Raw matrix released: " & _
                      rowsBefore & "行 x " & colsBefore & "列，内存已释放", "内存管理")
    Else
        Call WriteLog("C10KProcessor", "ClearRawData", "原始矩阵为空或已释放 / Raw matrix already empty or released", "内存管理")
    End If
End Sub

' ==============================================================================================
' SECTION 5: 核心数据处理 / Core Data Processing
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] TransformToLongList
' 说明：宽表转长表核心业务逻辑，支持大数据量分块处理判断，带进度显示和阈值过滤
'       针对稀疏数据优化：快速跳过空值，减少无效检查
'       Core wide-to-long conversion logic with chunk handling and progress display
'       Optimized for sparse data: fast empty value skip, reduce invalid checks
' 参数：startRow - Long，数据起始行 / Start row
'       segment - String，数据分段名称 / Segment name
' 返回值：Variant - 转换后的长表二维数组 / Converted long-format 2D array
' ----------------------------------------------------------------------------------------------
Public Function TransformToLongList(ByVal startRow As Long, ByVal segment As String) As Variant
    On Error GoTo ErrorHandler
    
    ' 常量定义 / Constants
    Const COL_COUNT As Long = 9                  ' 输出列数 / Output columns
    Const PROGRESS_STEP As Integer = 10          ' 进度更新步长 / Progress update step
    Const ERROR_NO_DATA As Long = 2001           ' 无数据错误码 / No data error
    Const ERROR_NO_PARSER As Long = 2002         ' 无解析器错误码 / No parser error
    
    ' 局部变量声明 / Local variables
    Dim rMax As Long, cMax As Long, r As Long, c As Long
    Dim count As Long, i As Long, j As Long
    Dim tempArray() As Variant
    Dim estimatedSize As Long, totalRowsToProcess As Long
    Dim lastProgress As Integer, progress As Integer
    Dim threshold As Long, materialityThreshold As Double
    Dim headers As Variant, parsedRow As Variant
    Dim config As Object, showProgress As Boolean
    Dim originalScreenUpdating As Boolean
    Dim actualDataCount As Long  ' 实际有效数据行数统计 / Actual valid data count
    
    Call WriteLog("C10KProcessor", "TransformToLongList", "开始宽表转长表处理: 分段=" & segment, "数据处理")
    
    ' 守卫检查 / Guard checks
    If Not IsLoaded Then
        Err.Raise ERROR_NO_DATA, "C10KProcessor.TransformToLongList", "分段无数据加载: " & segment
    End If
    
    If m_currentParser Is Nothing Then
        Err.Raise ERROR_NO_PARSER, "C10KProcessor.TransformToLongList", "分段无解析器可用: " & segment
    End If
    
    ' 记录解析器类型 / Log parser type
    Call WriteLog("C10KProcessor", "TransformToLongList", "使用解析器: " & TypeName(m_currentParser), "调试")
    
    ' 初始化基础参数 / Initialize basic parameters
    rMax = UBound(m_matrix, 1)
    cMax = UBound(m_matrix, 2)
    
    Call WriteLog("C10KProcessor", "TransformToLongList", "矩阵维度: " & rMax & "x" & cMax, "调试")
    Call WriteLog("C10KProcessor", "TransformToLongList", "起始行: " & startRow, "调试")
    
    ' 校验起始行合法性 / Validate start row
    If startRow < 1 Or startRow > rMax Then
        Err.Raise 2003, "C10KProcessor.TransformToLongList", "无效起始行: " & startRow & _
                  " (有效范围: 1-" & rMax & ")"
    End If
    
    ' 获取配置项 / Get configuration
    Set config = app_01_basic.GetConfig()
    threshold = m_config("ChunkThreshold")
    materialityThreshold = m_config("MaterialityThreshold")
    showProgress = m_config("ShowProgress")
    
    Call WriteLog("C10KProcessor", "TransformToLongList", "阈值: " & materialityThreshold, "调试")
    
    ' 计算待处理总行数 / Calculate total rows to process
    totalRowsToProcess = rMax - startRow + 1
    
    ' 分块处理判断 / Check if chunked processing needed
    estimatedSize = totalRowsToProcess * (cMax - 1)
    Call WriteLog("C10KProcessor", "TransformToLongList", "预估数据量: " & estimatedSize & " 行", "调试")
    
    If estimatedSize > threshold Then
        Call WriteLog("C10KProcessor", "TransformToLongList", "数据量超过阈值，启用分块处理", "数据处理")
        TransformToLongList = TransformToLongListChunked(startRow, segment, estimatedSize)
        GoTo Cleanup
    End If
    
    ' Excel环境优化 / Excel environment optimization
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 预分配临时数组 / Pre-allocate temporary array
    ReDim tempArray(1 To estimatedSize + 1, 1 To COL_COUNT)
    count = 1 ' 表头行 / Header row
    actualDataCount = 0
    
    ' 写入表头 / Write headers
    headers = m_currentParser.GetHeaders()
    If Not IsArray(headers) Then
        Err.Raise 2004, "C10KProcessor.TransformToLongList", "解析器返回表头非数组"
    End If
    
    Call WriteLog("C10KProcessor", "TransformToLongList", "表头列数: " & (UBound(headers) + 1), "调试")
    
    If UBound(headers) <> COL_COUNT - 1 Then
        Err.Raise 2004, "C10KProcessor.TransformToLongList", "表头列数不正确: " & (UBound(headers) + 1) & " (预期: " & COL_COUNT & ")"
    End If
    
    For c = 0 To COL_COUNT - 1
        tempArray(1, c + 1) = headers(c)
    Next c
    
    ' 主循环：宽表转长表 / Main loop: wide to long conversion
    For r = startRow To rMax
        ' 进度显示 / Progress display
        If showProgress Then
            progress = Int((r - startRow) / totalRowsToProcess * 100)
            If progress > lastProgress And progress Mod PROGRESS_STEP = 0 Then
                Application.StatusBar = "处理中 / Processing: " & progress & "% (" & r & "/" & rMax & ")"
                lastProgress = progress
                DoEvents
            End If
        End If
        
        ' 遍历列（从第2列开始，第1列为年份）/ Iterate columns (column 1 is year)
        For c = 2 To cMax
            ' 优化1：快速跳过空值和非数值 / Fast skip for empty and non-numeric
            If IsEmpty(m_matrix(r, c)) Then GoTo NextColumn
            If Not IsNumeric(m_matrix(r, c)) Then GoTo NextColumn
            
            ' 阈值过滤 / Materiality filter
            Dim val As Double
            val = CDbl(m_matrix(r, c))
            If val > materialityThreshold Then
                count = count + 1
                actualDataCount = actualDataCount + 1
                
                ' 基础字段赋值 / Basic fields
                tempArray(count, COL_LOSS_SEQ) = count - 1
                tempArray(count, COL_COLUMN_ID) = c
                tempArray(count, COL_YEAR) = m_matrix(r, 1)
                tempArray(count, COL_LOSS_AMT) = val
                tempArray(count, COL_TID) = ""
                
                ' 解析器处理动态字段 / Parser handles dynamic fields
                parsedRow = m_currentParser.ParseRow(m_matrix, r, c, startRow, segment)
                
                ' 优化2：简化解析器结果处理 / Simplified parser result handling
                If IsArray(parsedRow) Then
                    Dim parsedUBound As Long
                    On Error Resume Next
                    parsedUBound = UBound(parsedRow)
                    If Err.Number = 0 And parsedUBound >= 9 Then
                        tempArray(count, COL_SUBLOB) = parsedRow(4)
                        tempArray(count, COL_PERIL) = parsedRow(5)
                        tempArray(count, COL_EVENT) = parsedRow(6)
                        tempArray(count, COL_GROUP) = parsedRow(9)
                    Else
                        ' 解析器返回数组格式错误，填充默认值 / Invalid parser array, fill with defaults
                        tempArray(count, COL_SUBLOB) = "Error"
                        tempArray(count, COL_PERIL) = "Error"
                        tempArray(count, COL_EVENT) = "Error"
                        tempArray(count, COL_GROUP) = "Error"
                        Call WriteLog("C10KProcessor", "TransformToLongList", _
                                    "解析器返回数组格式错误 / Invalid parser array (行/row=" & r & ", 列/col=" & c & ")", "警告")
                    End If
                    On Error GoTo 0
                Else
                    ' 解析器返回非数组，填充空值 / Parser returned non-array, fill with empty
                    tempArray(count, COL_SUBLOB) = ""
                    tempArray(count, COL_PERIL) = ""
                    tempArray(count, COL_EVENT) = ""
                    tempArray(count, COL_GROUP) = ""
                End If
            End If
NextColumn:
        Next c
    Next r
    
    ' 优化3：输出实际有效数据行数 / Log actual valid data count
    Call WriteLog("C10KProcessor", "TransformToLongList", _
                  "转换完成 / Conversion completed: " & actualDataCount & " 行有效数据 / valid data rows", "数据处理")
    
    ' 调整数组到实际大小 / Resize to actual size
    If count > 1 Then
        ReDim m_result(1 To count, 1 To COL_COUNT)
        For i = 1 To count
            For j = 1 To COL_COUNT
                m_result(i, j) = tempArray(i, j)
            Next j
        Next i
        TransformToLongList = m_result
    Else
        Call WriteLog("C10KProcessor", "TransformToLongList", "无有效数据 / No valid data after conversion", "警告")
        TransformToLongList = Empty
    End If
    
Cleanup:
    ' 恢复Excel环境 / Restore Excel environment
    Application.ScreenUpdating = originalScreenUpdating
    Application.StatusBar = False
    Erase tempArray
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = Err.Description
    If errMsg = "" Then
        ' 根据错误号给出更具体的提示 / More specific messages based on error number
        Select Case Err.Number
            Case 9
                errMsg = "下标越界 (可能数组索引超出范围) / Subscript out of range"
            Case 13
                errMsg = "类型不匹配 (可能数据类型错误) / Type mismatch"
            Case 91
                errMsg = "对象变量未设置 / Object variable not set"
            Case Else
                errMsg = "未知错误 (错误号: " & Err.Number & ") / Unknown error"
        End Select
    End If
    
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    
    Call WriteLog("C10KProcessor", "TransformToLongList", "转换失败: " & errMsg, "错误")
    HandleError "C10KProcessor.TransformToLongList", errMsg
    TransformToLongList = Empty
End Function

' ----------------------------------------------------------------------------------------------
' [F] AggregateBySubLoB
' 说明：按SubLoB维度聚合万年数据，用于GN工作表统计，支持LoB配置匹配的容错/严格模式
'       SubLoB aggregation for GN worksheet, with tolerance/strict mode for LoB matching
' 参数：startRow - Long，数据起始行 / Start row
'       segment - String，数据分段名称 / Segment name
'       yearMaxParam - Long（可选），最大统计年份 / Max years (optional)
'       isStrictLoBMatching - Boolean（可选），LoB匹配严格模式 / Strict LoB matching mode
' 返回值：Variant - 聚合后的二维数组 / Aggregated 2D array
' ----------------------------------------------------------------------------------------------
Public Function AggregateBySubLoB(ByVal startRow As Long, ByVal segment As String, _
                                   Optional ByVal yearMaxParam As Long = 0, _
                                   Optional ByVal isStrictLoBMatching As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    
    ' 常量定义 / Constants
    Const ERR_NO_DATA_LOADED As Long = 3001          ' 无数据加载错误码 / No data loaded
    Const ERR_INVALID_YEARMAX As Long = 3002         ' 无效年份最大值 / Invalid yearMax
    Const ERR_MATRIX_NOT_ARRAY As Long = 3003        ' 矩阵非数组 / Matrix not array
    Const ERR_INVALID_STARTROW As Long = 3004        ' 无效起始行 / Invalid start row
    Const ERR_SUBLOB_NO_LOB_MAPPING As Long = 3005   ' SubLoB无对应LoB配置 / No LoB mapping
    Const HEADER_COLS As Long = 10                   ' 固定表头列数 / Fixed header columns
    
    ' 局部变量声明 / Local variables
    Dim yearMax As Long, rMax As Long, cMax As Long
    Dim numSub As Long, i As Long, r As Long, c As Long
    Dim uniqueSubs As Variant, resArray() As Variant
    Dim dictLoB As Object, dictSubRow As Object
    Dim sName As String, currentSub As String
    Dim limitRow As Long, targetRow As Long, yrIdx As Long
    Dim originalScreenUpdating As Boolean
    
    Call WriteLog("C10KProcessor", "AggregateBySubLoB", "开始聚合处理 / Starting aggregation: 分段/Segment=" & segment, "数据处理")
    
    ' 守卫检查 / Guard checks
    If Not IsLoaded Then
        Err.Raise ERR_NO_DATA_LOADED, "C10KProcessor.AggregateBySubLoB", _
                  "分段无数据加载 / No data loaded for segment: " & segment
    End If
    
    ' 处理年份参数 / Handle year parameter
    yearMax = IIf(yearMaxParam > 0, yearMaxParam, m_config("DefaultYearMax"))
    If yearMax <= 0 Then
        Err.Raise ERR_INVALID_YEARMAX, "C10KProcessor.AggregateBySubLoB", _
                  "无效最大年份值 / Invalid yearMax: " & yearMax
    End If
    
    ' 获取矩阵维度 / Get matrix dimensions
    If Not IsArray(m_matrix) Then
        Err.Raise ERR_MATRIX_NOT_ARRAY, "C10KProcessor.AggregateBySubLoB", "矩阵数据非有效数组 / Matrix is not an array"
    End If
    
    rMax = UBound(m_matrix, 1)
    cMax = UBound(m_matrix, 2)
    
    ' 校验起始行 / Validate start row
    If startRow < 1 Or startRow > rMax Then
        Err.Raise ERR_INVALID_STARTROW, "C10KProcessor.AggregateBySubLoB", _
                  "无效起始行 / Invalid start row: " & startRow
    End If
    
    ' Excel环境优化 / Excel environment optimization
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 提取唯一SubLoB / Extract unique SubLoBs
    uniqueSubs = ExtractUniqueSubLoBs()
    If IsEmpty(uniqueSubs) Then
        Call WriteLog("C10KProcessor", "AggregateBySubLoB", "无唯一SubLoB / No unique SubLoBs found", "警告")
        AggregateBySubLoB = Empty
        GoTo Cleanup
    End If
    
    numSub = UBound(uniqueSubs) + 1
    
    ' 加载LoB配置参数 / Load LoB parameters
    Set dictLoB = LoadLoBParameters()
    
    ' 提前检查LoB配置（守卫逻辑）/ Early check for LoB config
    If dictLoB Is Nothing And isStrictLoBMatching Then
        Err.Raise ERR_SUBLOB_NO_LOB_MAPPING, "C10KProcessor.AggregateBySubLoB", _
                  "LoB配置字典为空且开启严格匹配模式 / LoB config empty in strict mode [分段/Segment: " & segment & "]"
    End If
    
    ' 预分配结果数组 + 构建SubLoB行映射 / Pre-allocate result array + build row mapping
    Set dictSubRow = CreateObject("Scripting.Dictionary")
    dictSubRow.CompareMode = vbTextCompare
    
    ReDim resArray(1 To numSub, 1 To HEADER_COLS + yearMax)
    
    ' 填充SubLoB基础信息 / Populate SubLoB basic info
    For i = 0 To numSub - 1
        sName = Trim(uniqueSubs(i))
        If sName <> "" Then
            dictSubRow(sName) = i + 1
            
            resArray(i + 1, 1) = segment
            resArray(i + 1, 9) = sName
            
            ' LoB配置匹配 / LoB config matching
            If Not dictLoB Is Nothing And dictLoB.Exists(sName) Then
                On Error Resume Next
                resArray(i + 1, 7) = dictLoB(sName)(0)  ' Class
                resArray(i + 1, 8) = dictLoB(sName)(1)  ' Type
                On Error GoTo ErrorHandler
            Else
                If isStrictLoBMatching Then
                    Err.Raise ERR_SUBLOB_NO_LOB_MAPPING, "C10KProcessor.AggregateBySubLoB", _
                              "SubLoB无对应LoB配置 / No LoB mapping for: " & sName & " [分段/Segment: " & segment & "]"
                Else
                    resArray(i + 1, 7) = "N/A"
                    resArray(i + 1, 8) = "N/A"
                    Call WriteLog("C10KProcessor", "AggregateBySubLoB", "警告 / Warning: SubLoB无对应LoB配置 / No LoB mapping for: " & sName, "配置警告")
                End If
            End If
        End If
    Next i
    
    ' 核心聚合逻辑 / Core aggregation logic
    limitRow = IIf(rMax < (startRow + yearMax - 1), rMax, startRow + yearMax - 1)
    
    For c = 2 To cMax
        currentSub = Trim(m_matrix(1, c))
        If dictSubRow.Exists(currentSub) Then
            targetRow = dictSubRow(currentSub)
            
            For r = startRow To limitRow
                If IsNumeric(m_matrix(r, c)) Then
                    yrIdx = r - startRow + 1
                    resArray(targetRow, yrIdx + HEADER_COLS) = _
                        resArray(targetRow, yrIdx + HEADER_COLS) + CDbl(m_matrix(r, c))
                    resArray(targetRow, HEADER_COLS) = _
                        resArray(targetRow, HEADER_COLS) + CDbl(m_matrix(r, c))
                End If
            Next r
        End If
    Next c
    
    ' 计算均值 / Calculate mean
    For i = 1 To numSub
        If yearMax > 0 Then
            resArray(i, HEADER_COLS) = resArray(i, HEADER_COLS) / yearMax
        Else
            resArray(i, HEADER_COLS) = 0
        End If
    Next i
    
    m_result = resArray
    
    Call WriteLog("C10KProcessor", "AggregateBySubLoB", "聚合完成 / Aggregation completed: " & numSub & " 个唯一SubLoB", "数据处理")
    AggregateBySubLoB = resArray
    
Cleanup:
    Application.ScreenUpdating = originalScreenUpdating
    Set dictLoB = Nothing
    Set dictSubRow = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    Dim errMsg As String
    errMsg = Err.Description & " [分段/Segment: " & segment & ", 最大年份/YearMax: " & yearMax & _
             ", LoB严格匹配/Strict: " & isStrictLoBMatching & "]"
    
    Call WriteLog("C10KProcessor", "AggregateBySubLoB", "聚合失败 / Aggregation failed: " & errMsg, "错误")
    HandleError "C10KProcessor.AggregateBySubLoB", errMsg
    
    Set dictLoB = Nothing
    Set dictSubRow = Nothing
    AggregateBySubLoB = Empty
End Function

' ----------------------------------------------------------------------------------------------
' [f] TransformToLongListChunked
' 说明：宽表转长表分块处理，适配超大数据量的处理场景，将数据分批处理并合并结果
'       针对稀疏数据优化：快速跳过空值，减少无效检查
'       Chunked processing for very large datasets, process in batches and merge results
'       Optimized for sparse data: fast empty value skip, reduce invalid checks
' 参数：startRow - Long，数据起始行 / Start row
'       segment - String，数据分段名称 / Segment name
'       estimatedSize - Long，估算数据大小 / Estimated size
' 返回值：Variant - 处理后的数组 / Result array
' ----------------------------------------------------------------------------------------------
Private Function TransformToLongListChunked(ByVal startRow As Long, ByVal segment As String, _
                                             ByVal estimatedSize As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' 常量定义 / Constants
    Const COL_COUNT As Long = 9
    Const PROGRESS_STEP As Integer = 10
    
    ' 局部变量声明 / Local variables
    Dim rMax As Long, cMax As Long, r As Long, c As Long
    Dim chunkSize As Long
    Dim chunkStart As Long, chunkEnd As Long
    Dim chunkCount As Long, totalChunks As Long
    Dim i As Long, j As Long
    Dim allResults As Collection
    Dim headers As Variant, parsedRow As Variant
    Dim materialityThreshold As Double
    Dim progress As Integer, lastProgress As Integer
    Dim totalRowsToProcess As Long
    Dim actualDataCount As Long  ' 实际有效数据行数统计 / Actual valid data count
    
    ' 获取配置 / Get configuration
    chunkSize = m_config("ChunkSize")
    materialityThreshold = m_config("MaterialityThreshold")
    rMax = UBound(m_matrix, 1)
    cMax = UBound(m_matrix, 2)
    totalRowsToProcess = rMax - startRow + 1
    totalChunks = Int((totalRowsToProcess - 1) / chunkSize) + 1
    actualDataCount = 0
    
    Call WriteLog("C10KProcessor", "TransformToLongListChunked", _
                  "开始分块处理 / Starting chunked processing: 总行数/Total rows=" & totalRowsToProcess & _
                  ", 块大小/Chunk size=" & chunkSize & ", 总块数/Total chunks=" & totalChunks, "数据处理")
    
    ' 获取表头 / Get headers
    headers = m_currentParser.GetHeaders()
    If Not IsArray(headers) Or UBound(headers) <> COL_COUNT - 1 Then
        Err.Raise 2004, "C10KProcessor.TransformToLongListChunked", "解析器返回表头格式无效 / Invalid header format from parser"
    End If
    
    ' 初始化结果集合 / Initialize results collection
    Set allResults = New Collection
    
    ' 添加表头行 / Add header row
    allResults.Add headers
    
    ' Excel环境优化 / Excel environment optimization
    Application.ScreenUpdating = False
    lastProgress = -1
    
    ' 按块循环处理 / Process chunk by chunk
    For chunkStart = startRow To rMax Step chunkSize
        chunkEnd = Application.Min(chunkStart + chunkSize - 1, rMax)
        chunkCount = chunkCount + 1
        
        ' 进度显示 / Progress display
        progress = Int((chunkStart - startRow) / totalRowsToProcess * 100)
        If progress > lastProgress And progress Mod PROGRESS_STEP = 0 Then
            Application.StatusBar = "分块处理中 / Chunked processing [" & segment & "]: " & progress & _
                                   "% (块/Chunk " & chunkCount & "/" & totalChunks & ")"
            lastProgress = progress
            DoEvents
        End If
        
        ' 处理当前块 / Process current chunk
        For r = chunkStart To chunkEnd
            For c = 2 To cMax
                ' 优化：快速跳过空值和非数值 / Fast skip for empty and non-numeric
                If IsEmpty(m_matrix(r, c)) Then GoTo NextColumn
                If Not IsNumeric(m_matrix(r, c)) Then GoTo NextColumn
                
                ' 阈值过滤 / Materiality filter
                Dim val As Double
                val = CDbl(m_matrix(r, c))
                If val > materialityThreshold Then
                    actualDataCount = actualDataCount + 1
                    
                    ' 基础字段赋值 / Basic fields
                    Dim rowData(1 To COL_COUNT) As Variant
                    rowData(COL_LOSS_SEQ) = allResults.count  ' 使用当前计数作为序号 / Use current count as sequence
                    rowData(COL_COLUMN_ID) = c
                    rowData(COL_YEAR) = m_matrix(r, 1)
                    rowData(COL_LOSS_AMT) = val
                    rowData(COL_TID) = ""
                    
                    ' 解析器处理动态字段 / Parser handles dynamic fields
                    parsedRow = m_currentParser.ParseRow(m_matrix, r, c, startRow, segment)
                    
                    ' 优化：简化解析器结果处理 / Simplified parser result handling
                    If IsArray(parsedRow) Then
                        Dim parsedUBound As Long
                        On Error Resume Next
                        parsedUBound = UBound(parsedRow)
                        If Err.Number = 0 And parsedUBound >= 9 Then
                            rowData(COL_SUBLOB) = parsedRow(4)
                            rowData(COL_PERIL) = parsedRow(5)
                            rowData(COL_EVENT) = parsedRow(6)
                            rowData(COL_GROUP) = parsedRow(9)
                        Else
                            ' 解析器返回数组格式错误，填充默认值 / Invalid parser array, fill with defaults
                            rowData(COL_SUBLOB) = "Error"
                            rowData(COL_PERIL) = "Error"
                            rowData(COL_EVENT) = "Error"
                            rowData(COL_GROUP) = "Error"
                            Call WriteLog("C10KProcessor", "TransformToLongListChunked", _
                                        "解析器返回数组格式错误 / Invalid parser array (行/row=" & r & ", 列/col=" & c & ")", "警告")
                        End If
                        On Error GoTo 0
                    Else
                        ' 解析器返回非数组，填充空值 / Parser returned non-array, fill with empty
                        rowData(COL_SUBLOB) = ""
                        rowData(COL_PERIL) = ""
                        rowData(COL_EVENT) = ""
                        rowData(COL_GROUP) = ""
                    End If
                    
                    ' 将行数据添加到集合 / Add row data to collection
                    Dim rowArray As Variant
                    rowArray = rowData
                    allResults.Add rowArray
                End If
NextColumn:
            Next c
        Next r
    Next chunkStart
    
    ' 优化：输出实际有效数据行数 / Log actual valid data count
    Call WriteLog("C10KProcessor", "TransformToLongListChunked", _
                  "分块处理完成 / Chunked processing completed: " & actualDataCount & _
                  " 行有效数据 / valid data rows (总行/Total: " & (allResults.count - 1) & ")", "数据处理")
    
    ' 将集合转换为二维数组 / Convert collection to 2D array
    If allResults.count > 1 Then
        ReDim m_result(1 To allResults.count, 1 To COL_COUNT)
        
        ' 第一行是表头 / First row is header
        For j = 1 To COL_COUNT
            m_result(1, j) = allResults(1)(j - 1)
        Next j
        
        ' 数据行 / Data rows
        For i = 2 To allResults.count
            Dim rowValues As Variant
            rowValues = allResults(i)
            For j = 1 To COL_COUNT
                m_result(i, j) = rowValues(j)
            Next j
        Next i
        
        TransformToLongListChunked = m_result
    Else
        Call WriteLog("C10KProcessor", "TransformToLongListChunked", "无有效数据 / No valid data", "警告")
        TransformToLongListChunked = Empty
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call WriteLog("C10KProcessor", "TransformToLongListChunked", "分块处理失败 / Chunked processing failed: " & Err.Description, "错误")
    HandleError "C10KProcessor.TransformToLongListChunked", Err.Description
    TransformToLongListChunked = Empty
End Function

' ==============================================================================================
' SECTION 6: 私有辅助工具方法 / Private Helper Methods
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] ValidateMatrix
' 说明：验证加载的矩阵数据是否为有效的二维数组
'       Validate loaded matrix is a valid 2D array
' 返回值：Boolean - True=有效二维数组 / Valid 2D array, False=无效数据 / Invalid
' ----------------------------------------------------------------------------------------------
Private Function ValidateMatrix() As Boolean
    On Error Resume Next
    
    If IsEmpty(m_matrix) Then
        ValidateMatrix = False
        Exit Function
    End If
    
    If Not IsArray(m_matrix) Then
        ValidateMatrix = False
        Exit Function
    End If
    
    ValidateMatrix = (GetArrayDimensions(m_matrix) = 2)
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetArrayDimensions
' 说明：获取任意数组的维度数量，用于矩阵数据的有效性校验
'       Get number of dimensions of an array
' 参数：arr - Variant，待检测的数组 / Array to check
' 返回值：Integer - 数组的维度数量 / Number of dimensions
' ----------------------------------------------------------------------------------------------
Private Function GetArrayDimensions(ByRef arr As Variant) As Integer
    On Error Resume Next
    Dim dimCount As Integer
    dimCount = 0
    
    Do While True
        dimCount = dimCount + 1
        Dim ub As Long
        ub = UBound(arr, dimCount)
        If Err.Number <> 0 Then
            GetArrayDimensions = dimCount - 1
            Exit Function
        End If
    Loop
End Function

' ----------------------------------------------------------------------------------------------
' [f] ExtractUniqueSubLoBs
' 说明：从矩阵表头行（第一行）提取唯一的SubLoB名称
'       Extract unique SubLoB names from matrix header row
' 返回值：Variant - 唯一SubLoB名称的一维数组 / 1D array of unique SubLoB names
' ----------------------------------------------------------------------------------------------
Private Function ExtractUniqueSubLoBs() As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim c As Long
    For c = 2 To UBound(m_matrix, 2)
        If Not IsEmpty(m_matrix(1, c)) Then
            dict(CStr(m_matrix(1, c))) = Empty
        End If
    Next c
    
    ExtractUniqueSubLoBs = dict.Keys
End Function

' ----------------------------------------------------------------------------------------------
' [f] LoadLoBParameters
' 说明：从Setup工作表的LoB配置区域加载数据，构建SubLoB到LoB(分类/类型)的映射字典
'       Load LoB configuration from Setup sheet, build SubLoB to LoB mapping
' 返回值：Object - 映射字典 / Mapping dictionary
' ----------------------------------------------------------------------------------------------
Private Function LoadLoBParameters() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Dim paraData As Variant
    paraData = app_01_basic.Setup.Range("rng_LoB").value
    
    If IsArray(paraData) Then
        Dim i As Long
        For i = 1 To UBound(paraData, 1)
            dict(CStr(paraData(i, 1))) = Array(paraData(i, 2), paraData(i, 3))
        Next i
    End If
    On Error GoTo 0
    
    Set LoadLoBParameters = dict
End Function

' ==============================================================================================
' SECTION 7: 调试与信息输出 / Debugging & Information
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] DumpInfo
' 说明：输出类实例的核心状态信息到立即窗口和日志
'       Dump core instance info to immediate window and log
' ----------------------------------------------------------------------------------------------
Public Sub DumpInfo()
    Call WriteLog("C10KProcessor", "DumpInfo", "=== C10KProcessor 实例信息 / Instance Info ===", "调试信息")
    Call WriteLog("C10KProcessor", "DumpInfo", "数据加载状态 / Loaded: " & IsLoaded, "调试信息")
    Call WriteLog("C10KProcessor", "DumpInfo", "源文件路径 / Source: " & m_filePath, "调试信息")
    Call WriteLog("C10KProcessor", "DumpInfo", "矩阵数据维度 / Matrix: " & RowCount & " 行 x " & ColumnCount & " 列", "调试信息")
    Call WriteLog("C10KProcessor", "DumpInfo", "配置项 / Config:", "调试信息")
    
    Dim key As Variant
    For Each key In m_config.Keys
        Call WriteLog("C10KProcessor", "DumpInfo", "  " & key & ": " & m_config(key), "调试信息")
    Next key
    
    Call WriteLog("C10KProcessor", "DumpInfo", "===============================", "调试信息")
End Sub

