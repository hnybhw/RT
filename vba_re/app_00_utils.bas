Attribute VB_Name = "app_00_utils"
'
' ==============================================================================================
' MODULE NAME     : app_00_utils (FINAL VERSION)
' PURPOSE         : 通用工具函数库，为所有模块提供安全、统一的数组检测、范围操作、配置管理、
'                  Application状态管理、安全数学运算、错误处理包装、路径工具和性能计时功能。
'                  本模块旨在作为整个项目的"基础设施"，集中处理所有边界检查和防御性编程。
'                 : Universal utility library providing safe array operations, range handling,
'                   config management, Application state management, safe math, error handling,
'                   path utilities, and performance timing.
' DEPENDS         : app_01_basic (for WriteLog, HandleError, MODULE_NAME constant)
'                 : Microsoft Excel Object Library (Range, Worksheet, Application)
' NOTE            : 本模块所有函数均设计为无副作用、可重入、线程安全（在VBA单线程环境下）
'                 : All functions are designed to be side-effect free, reentrant, and thread-safe
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 0: 私有日志辅助函数 / Private Logging Helpers
'   [f] LogInfo          - 记录信息日志
'   [f] LogWarn          - 记录警告日志
'   [f] LogError         - 记录错误日志
'   [f] LogDebug         - 记录调试日志（仅在DEBUG_MODE下生效）
'
' SECTION 1: 模块常量和类型 / Module Constants and Types
'   [C] MODULE_NAME      - 模块名称常量（用于日志记录）
'   [C] 配置键常量       - CFG_YEAR_MAX, CFG_MATERIALITY_THRESHOLD, CFG_CHUNK_SIZE,
'                          CFG_CHUNK_THRESHOLD, CFG_PYTHON_TIMEOUT, CFG_STRICT_LOB_MATCHING
'   [C] 数组规范常量     - ARRAY_BASE_DEFAULT, ARRAY_MIN_ROWS, ARRAY_MIN_COLS
'   [C] 安全边界常量     - SAFE_MAX_LONG, SAFE_MAX_ARRAY_ELEMENTS, SAFE_MAX_STRING_LENGTH
'   [C] 空值/默认值常量  - EMPTY_VALUE, DEFAULT_STRING, DEFAULT_LONG, DEFAULT_DOUBLE, DEFAULT_BOOLEAN
'
' SECTION 2: 数组安全检测与标准化 / Array Safety and Normalization
'   [F] GetArrayInfo           - 安全获取数组维度信息（永不崩溃）
'   [f] GetArrayDimensions     - 安全获取数组维度数（无硬编码上限）
'   [F] IsArrayValid           - 统一检测数组是否有效且包含数据
'   [F] NormalizeTo1Based      - 将任意数组转换为1-based（便于后续处理）
'   [f] Normalize1DArray       - 标准化一维数组为二维1-based数组（纯循环实现）
'   [f] Normalize2DArray       - 标准化二维数组为二维1-based数组
'   [F] EnsureArrayDimensions  - 确保数组至少具有指定维度
'   [F] SliceArraySafe         - 安全切片（带边界检查）
'
' SECTION 3: 范围/工作表安全操作 / Range and Worksheet Safety
'   [F] GetNamedRangeSafely    - 安全获取命名范围（带存在性验证）
'   [F] GetActualDataRange     - 获取工作表的实际数据范围（避免UsedRange陷阱）
'   [F] IsWorksheetValid       - 检查工作表是否存在且可用
'   [F] GetLastRowSafely       - 安全获取最后一行的行号
'   [F] GetLastColSafely       - 安全获取最后一列的列号
'
' SECTION 4: 配置管理统一化 / Configuration Management
'   [F] GetConfigValue         - 统一的配置获取函数（带类型转换和默认值）
'   [F] SetConfigValue         - 安全设置配置值（带类型验证）
'   [F] ConfigKeyExists        - 检查配置键是否存在
'
' SECTION 5: Application状态管理 / Application State Management
'   [F] SaveAppState           - 保存当前Application状态
'   [S] RestoreAppState        - 恢复Application状态
'   [F] RunOptimized           - 在优化环境中运行宏（自动恢复）
'   [f] ApplyOptimizedSettings - 应用优化设置
'   [F] RunOptimizedWithParams - 带参数的优化运行
'
' SECTION 6: 安全数学运算 / Safe Math Operations
'   [F] SafeMultiply           - 安全乘法（检查溢出和NaN）
'   [F] SafeAdd                - 安全加法（检查溢出和NaN）
'   [F] CalculateSafeArraySize - 计算安全的数组大小
'
' SECTION 7: 错误处理增强 / Enhanced Error Handling
'   [F] SafeExecute            - 统一的安全执行包装器
'   [f] LogExecutionError      - 记录执行错误
'   [f] BuildErrorMessage      - 构建错误消息
'   [F] IsInDesignMode         - 检查是否在设计模式
'   [S] Assert                 - 调试断言（仅在开发环境生效）
'   [f] GetCallerInfo          - 获取调用方信息
'   [f] LogAssertion           - 将断言失败写入日志文件
'   [f] GetAssertLogPath       - 获取断言日志文件路径
'
' SECTION 8: 数据类型转换与验证 / Data Type Conversion and Validation
'   [F] ToSafeLong             - 安全转换为Long（带范围校验）
'   [F] ToSafeDouble           - 安全转换为Double（带范围校验）
'   [F] ToSafeString           - 安全转换为String
'   [F] IsNumericSafe          - 安全检查是否为数值
'   [F] IsDateSafe             - 安全检查是否为日期
'   [F] SanitizeString         - 清理字符串中的不可打印字符（正则实现）
'   [f] SanitizeStringFallback - 回退到循环方法（兼容性保障）
'
' SECTION 9: 路径与文件工具 / Path and File Utilities
'   [F] NormalizePath          - 规范化路径
'   [F] SafePathCombine        - 安全组合路径
'   [F] IsNetworkPath          - 判断是否为网络路径
'   [F] IsSharePointPath       - 判断是否为SharePoint路径（支持HTTP/HTTPS和UNC）
'   [F] EnsureFolderExists     - 确保文件夹存在（支持多级递归创建）
'   [f] GetParentFolderPath    - 获取父文件夹路径
'   [F] GetTempFilePath        - 获取临时文件路径（带唯一性校验）
'   [f] FormatDateTimeWithMS   - 格式化日期时间（支持毫秒）
'
' SECTION 10: 性能计时工具 / Performance Timing
'   [S] StartTimer             - 启动计时器
'   [F] ElapsedTime            - 获取已用时间
'   [F] FormatElapsedTime      - 格式化时间显示（支持时/分/秒/毫秒）
'   [S] LogPerformance         - 记录性能日志（支持阈值警告）
'   [S] ClearTimers            - 清空所有计时器
'   [F] GetAllTimers           - 获取所有计时器名称列表（调试用）
'   [F] TimerExists            - 检查指定计时器是否存在
'
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 0: 私有日志辅助函数 / Private Logging Helpers
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] LogInfo
' 私有辅助函数：记录信息日志
' ----------------------------------------------------------------------------------------------
Private Sub LogInfo(ByVal procName As String, ByVal message As String, Optional ByVal logType As String = "信息")
    Call app_01_basic.WriteLog(MODULE_NAME, procName, message, logType)
End Sub

' ----------------------------------------------------------------------------------------------
' [f] LogWarn
' 私有辅助函数：记录警告日志
' ----------------------------------------------------------------------------------------------
Private Sub LogWarn(ByVal procName As String, ByVal message As String, Optional ByVal logType As String = "警告")
    Call app_01_basic.WriteLog(MODULE_NAME, procName, "警告: " & message, logType)
End Sub

' ----------------------------------------------------------------------------------------------
' [f] LogError
' 私有辅助函数：记录错误日志
' ----------------------------------------------------------------------------------------------
Private Sub LogError(ByVal procName As String, ByVal message As String, Optional ByVal logType As String = "错误处理")
    Call app_01_basic.WriteLog(MODULE_NAME, procName, "错误: " & message, logType)
End Sub

' ----------------------------------------------------------------------------------------------
' [f] LogDebug
' 私有辅助函数：记录调试日志（仅在 DEBUG_MODE 下生效）
' ----------------------------------------------------------------------------------------------
Private Sub LogDebug(ByVal procName As String, ByVal message As String)
    #If DEBUG_MODE Then
        Call app_01_basic.WriteLog(MODULE_NAME, procName, "调试: " & message, "调试")
    #End If
End Sub

' ==============================================================================================
' SECTION 1: 模块常量和类型 / Module Constants and Types
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' 模块名称常量 - 用于日志记录
' ----------------------------------------------------------------------------------------------
Public Const MODULE_NAME As String = "app_00_utils"

' ----------------------------------------------------------------------------------------------
' [C] 配置键常量 - 统一所有模块的配置键名（解决大小写/命名不一致问题）
' 所有模块必须使用这些常量访问配置，严禁直接使用字符串硬编码
' ----------------------------------------------------------------------------------------------
' 核心配置键
Public Const CFG_YEAR_MAX As String = "YearMax"                             ' 模拟最大年数
Public Const CFG_MATERIALITY_THRESHOLD As String = "MaterialityThreshold"   ' 重大性阈值
Public Const CFG_CHUNK_SIZE As String = "ChunkSize"                         ' 分块大小
Public Const CFG_CHUNK_THRESHOLD As String = "ChunkThreshold"               ' 触发分块的阈值行数
Public Const CFG_PYTHON_TIMEOUT As String = "PythonTimeout"                 ' Python超时时间（秒）
' 解析器相关配置键
Public Const CFG_STRICT_LOB_MATCHING As String = "StrictLoBMatching"        ' 是否严格匹配LoB
' 后续新增配置键必须在此添加常量定义

' ----------------------------------------------------------------------------------------------
' [C] 数组规范常量 - 定义项目数组约定
' ----------------------------------------------------------------------------------------------
Public Const ARRAY_BASE_DEFAULT As Long = 1      ' 项目默认使用1-based数组（与Range.Value一致）
Public Const ARRAY_MIN_ROWS As Long = 1          ' 最小行数
Public Const ARRAY_MIN_COLS As Long = 1          ' 最小列数

' ----------------------------------------------------------------------------------------------
' [C] 安全边界常量 - 内存和安全限制
' ----------------------------------------------------------------------------------------------
Public Const SAFE_MAX_LONG As Double = 2 ^ 31 - 1           ' 安全Long最大值（约21亿，VBA Long类型上限为2,147,483,647）
Public Const SAFE_MAX_ARRAY_ELEMENTS As Long = 10000000     ' 安全数组元素个数上限（约1000万，经验值：VBA处理1000万元素仍较流畅）
Public Const SAFE_MAX_STRING_LENGTH As Long = 32767         ' 安全字符串长度上限（Excel单元格字符数上限为32,767）

' ----------------------------------------------------------------------------------------------
' [C] 空值/默认值常量 - 统一模块内的空值语义
' ----------------------------------------------------------------------------------------------
Public Const EMPTY_VALUE As Variant = Empty             ' 统一空值，替代直接使用 Empty
Public Const DEFAULT_STRING As String = ""              ' 默认空字符串
Public Const DEFAULT_LONG As Long = 0                   ' 默认长整数
Public Const DEFAULT_DOUBLE As Double = 0               ' 默认双精度浮点数
Public Const DEFAULT_BOOLEAN As Boolean = False         ' 默认布尔值

' ==============================================================================================
' SECTION 2: 数组安全检测与标准化 / Array Safety and Normalization
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetArrayInfo
' 安全获取数组维度信息（永不崩溃）。无论数组是什么状态（Empty、未初始化、0-based、1-based），
' 都能返回完整的维度信息，供调用方决策。
'
' 返回: Dictionary对象，包含以下键：
'       - IsArray: Boolean - 是否为数组
'       - IsEmpty: Boolean - 是否为空（未分配）
'       - Dims: Long - 维度数（0表示非数组或空）
'       - LBound1, UBound1, RowCount: Long - 第1维信息
'       - LBound2, UBound2, ColCount: Long - 第2维信息（如果存在）
'       - ErrorMessage: String - 如果检测出错，包含错误信息
' ----------------------------------------------------------------------------------------------
Public Function GetArrayInfo(arr As Variant) As Object
    ' 使用后期绑定避免引用Scripting运行时
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 初始化核心键
    dict("IsArray") = False
    dict("IsEmpty") = True
    dict("Dims") = 0
    dict("ErrorMessage") = ""
    
    ' 基本检查
    If Not IsArray(arr) Then
        Set GetArrayInfo = dict
        Exit Function
    End If
    
    dict("IsArray") = True
    
    ' 获取维度信息
    On Error Resume Next
    Dim dims As Long
    dims = GetArrayDimensions(arr)
    
    If Err.Number <> 0 Then
        dict("ErrorMessage") = Err.Description
        On Error GoTo 0
        Set GetArrayInfo = dict
        Exit Function
    End If
    
    ' 验证维度范围（防御性检查）
    If dims < 0 Then dims = 0
    dict("Dims") = dims
    dict("IsEmpty") = (dims = 0)
    
    ' 获取第1维信息
    If dims >= 1 Then
        dict("LBound1") = LBound(arr, 1)
        dict("UBound1") = UBound(arr, 1)
        dict("RowCount") = dict("UBound1") - dict("LBound1") + 1
    Else
        ' 初始化第1维默认值
        dict("LBound1") = 0
        dict("UBound1") = -1
        dict("RowCount") = 0
    End If
    
    ' 获取第2维信息（如果存在）
    If dims >= 2 Then
        dict("LBound2") = LBound(arr, 2)
        dict("UBound2") = UBound(arr, 2)
        dict("ColCount") = dict("UBound2") - dict("LBound2") + 1
    Else
        ' 初始化第2维默认值
        dict("LBound2") = 0
        dict("UBound2") = -1
        dict("ColCount") = 0
    End If
    
    On Error GoTo 0
    
    Set GetArrayInfo = dict
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetArrayDimensions
' 私有辅助函数：安全获取数组维度数（无硬编码上限）
' ----------------------------------------------------------------------------------------------
Private Function GetArrayDimensions(arr As Variant) As Long
    Dim dims As Long
    dims = 0
    
    On Error Resume Next
    
    Do
        dims = dims + 1
        Dim temp As Long
        temp = LBound(arr, dims)
    Loop Until Err.Number <> 0
    
    GetArrayDimensions = dims - 1
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsArrayValid
' 统一检测数组是否有效且包含数据。解决VBA中Or不短路的问题，提供统一的空数组检测语义。
'
' 参数:
'   arr - 待检测的数组
'   requireData - 是否要求至少包含一行数据（默认True）
'   require2D - 是否要求是二维数组（默认False）
'
' 返回: Boolean - True表示数组有效且满足要求
' ----------------------------------------------------------------------------------------------
Public Function IsArrayValid(arr As Variant, _
                            Optional requireData As Boolean = True, _
                            Optional require2D As Boolean = False) As Boolean
    
    ' 1. 获取数组信息
    Dim info As Object
    Set info = GetArrayInfo(arr)
    
    ' 2. 检查是否为数组
    If Not info("IsArray") Then
        IsArrayValid = False
        Exit Function
    End If
    
    ' 3. 检查是否为空数组
    If info("IsEmpty") Then
        IsArrayValid = False
        Exit Function
    End If
    
    ' 4. 检查维度要求
    If require2D And info("Dims") < 2 Then
        IsArrayValid = False
        Exit Function
    End If
    
    ' 5. 检查数据要求
    If requireData Then
        If info("RowCount") = 0 Then
            IsArrayValid = False
            Exit Function
        End If
        If require2D And info("ColCount") = 0 Then
            IsArrayValid = False
            Exit Function
        End If
    End If
    
    ' 所有检查通过
    IsArrayValid = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] NormalizeTo1Based
' 将任意数组转换为1-based（便于后续处理）。如果数组已经是1-based且连续，直接返回原数组；
' 否则重建为1-based数组。对于三维及以上数组，保持原样并记录警告。
'
' 参数:
'   arr - 原始数组（可以是0-based、1-based、不连续等）
'   makeCopy - 是否强制创建副本（默认False，仅在必要时创建）
'
' 返回: Variant - 1-based数组，或原始数组（如果已符合要求或无法处理）
' ----------------------------------------------------------------------------------------------
Public Function NormalizeTo1Based(arr As Variant, Optional makeCopy As Boolean = False) As Variant
    ' 先验证数组有效性
    If Not IsArrayValid(arr, requireData:=False) Then
        NormalizeTo1Based = arr  ' 返回原数组（可能是Empty）
        Exit Function
    End If
    
    Dim info As Object
    Set info = GetArrayInfo(arr)
    
    ' 处理三维及以上数组 - 保持原样并记录警告
    If info("Dims") > 2 Then
        Call LogWarn("NormalizeTo1Based", _
             "遇到 " & info("Dims") & " 维数组，当前版本保持原样返回", _
             "数组处理")
        NormalizeTo1Based = arr
        Exit Function
    End If
    
    ' 判断是否为1-based（合并条件）
    Dim is1Based As Boolean
    is1Based = (info("LBound1") = 1)
    If info("Dims") = 2 Then
        is1Based = is1Based And (info("LBound2") = 1)
    End If
    
    ' 如果已经是1-based且不需要复制，直接返回
    If Not makeCopy And is1Based Then
        NormalizeTo1Based = arr
        Exit Function
    End If
    
    ' 根据不同维度处理
    Select Case info("Dims")
        Case 1
            NormalizeTo1Based = Normalize1DArray(arr, info)
        Case 2
            NormalizeTo1Based = Normalize2DArray(arr, info)
        Case Else
            ' 理论上不会执行到这里，但保留安全处理
            NormalizeTo1Based = arr
    End Select
End Function

' ----------------------------------------------------------------------------------------------
' [f] Normalize1DArray
' 私有辅助函数：标准化一维数组为二维1-based数组（纯循环实现，稳定性优先）
' ----------------------------------------------------------------------------------------------
Private Function Normalize1DArray(arr As Variant, info As Object) As Variant
    Dim rows As Long
    rows = info("RowCount")
    
    ' 创建目标数组
    Dim result() As Variant
    ReDim result(1 To rows, 1 To 1)
    
    ' 循环复制（稳定性优于 Transpose）
    Dim i As Long
    For i = 1 To rows
        result(i, 1) = arr(info("LBound1") + i - 1)
    Next i
    
    Normalize1DArray = result
End Function

' ----------------------------------------------------------------------------------------------
' [f] Normalize2DArray
' 私有辅助函数：标准化二维数组为二维1-based数组（使用高效转换）
' ----------------------------------------------------------------------------------------------
Private Function Normalize2DArray(arr As Variant, info As Object) As Variant
    Dim rows As Long, cols As Long
    rows = info("RowCount")
    cols = info("ColCount")
    
    ' 如果源数组已经是连续的且我们只需要调整边界
    If info("LBound1") = 1 And info("LBound2") = 1 Then
        ' 已经是1-based，直接返回（理论上不会进这个分支，但保留）
        Normalize2DArray = arr
        Exit Function
    End If
    
    ' 创建目标数组
    Dim result() As Variant
    ReDim result(1 To rows, 1 To cols)
    
    ' 优化：如果源数组是连续的，可以使用更快的赋值方式
    ' VBA 中无法直接对二维数组切片，但可以利用行复制
    Dim r As Long
    Dim srcRowStart As Long, srcRowEnd As Long
    srcRowStart = info("LBound1")
    srcRowEnd = info("UBound1")
    
    ' 逐行复制（比逐元素复制快）
    For r = 1 To rows
        Dim srcRow As Long
        srcRow = srcRowStart + r - 1
        
        Dim c As Long
        For c = 1 To cols
            result(r, c) = arr(srcRow, info("LBound2") + c - 1)
        Next c
    Next r
    
    Normalize2DArray = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] EnsureArrayDimensions
' 确保数组至少具有指定维度。如果数组不存在或维度不足，创建新的空数组。
'
' 参数:
'   arr - 待检查的数组（ByRef以便修改）
'   minRows - 最小行数要求
'   minCols - 最小列数要求
'
' 返回: Boolean - True表示数组已满足要求（可能是新建的）
' ----------------------------------------------------------------------------------------------
Public Function EnsureArrayDimensions(ByRef arr As Variant, _
                                     Optional minRows As Long = 1, _
                                     Optional minCols As Long = 1) As Boolean
    
    ' 如果数组无效，直接重建
    If Not IsArrayValid(arr, requireData:=False) Then
        ReDim arr(1 To minRows, 1 To minCols)
        EnsureArrayDimensions = True
        Exit Function
    End If
    
    Dim info As Object
    Set info = GetArrayInfo(arr)
    
    ' 检查维度是否足够
    If info("RowCount") >= minRows And info("ColCount") >= minCols Then
        ' 已满足要求
        EnsureArrayDimensions = True
        Exit Function
    End If
    
    ' 需要扩展数组
    Dim newRows As Long, newCols As Long
    newRows = WorksheetFunction.Max(info("RowCount"), minRows)
    newCols = WorksheetFunction.Max(info("ColCount"), minCols)
    
    ' 保存原数组（临时变量）
    Dim originalArr As Variant
    originalArr = arr
    
    ' 重建数组（使用 ReDim Preserve 仅当最后一维扩展时有效）
    On Error Resume Next
    
    ' 尝试直接 ReDim Preserve
    ReDim Preserve arr(1 To newRows, 1 To newCols)
    
    If Err.Number = 0 Then
        ' ReDim Preserve 成功，但可能丢失部分数据
        ' 需要从 originalArr 恢复数据
        Dim r As Long, c As Long
        For r = 1 To info("RowCount")
            For c = 1 To info("ColCount")
                arr(r, c) = originalArr(info("LBound1") + r - 1, info("LBound2") + c - 1)
            Next c
        Next r
        
        EnsureArrayDimensions = True
        On Error GoTo 0
        Exit Function
    End If
    
    ' ReDim Preserve 失败，回退到手动重建
    Dim newArr() As Variant
    ReDim newArr(1 To newRows, 1 To newCols)
    
    ' 复制原数据
    For r = 1 To info("RowCount")
        For c = 1 To info("ColCount")
            newArr(r, c) = originalArr(info("LBound1") + r - 1, info("LBound2") + c - 1)
        Next c
    Next r
    
    arr = newArr
    EnsureArrayDimensions = True
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] SliceArraySafe
' 安全切片（带边界检查）。从源数组中提取指定行/列范围，返回1-based新数组。
' 对于来自Worksheet.Range的数组，可以利用VBA的切片特性；对内存数组采用循环。
'
' 参数:
'   arr - 源数组
'   rowStart, rowEnd - 行范围（使用源数组的索引系统）
'   colStart, colEnd - 列范围（使用源数组的索引系统）
'
' 返回: Variant - 切片后的1-based数组，无效范围返回EMPTY_VALUE
' ----------------------------------------------------------------------------------------------
Public Function SliceArraySafe(arr As Variant, _
                              ByVal rowStart As Variant, ByVal rowEnd As Variant, _
                              ByVal colStart As Variant, ByVal colEnd As Variant) As Variant
    
    ' 验证源数组
    If Not IsArrayValid(arr, requireData:=True, require2D:=True) Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If
    
    Dim info As Object
    Set info = GetArrayInfo(arr)
    
    ' 处理缺省参数
    If IsMissing(rowStart) Or IsEmpty(rowStart) Then rowStart = info("LBound1")
    If IsMissing(rowEnd) Or IsEmpty(rowEnd) Then rowEnd = info("UBound1")
    If IsMissing(colStart) Or IsEmpty(colStart) Then colStart = info("LBound2")
    If IsMissing(colEnd) Or IsEmpty(colEnd) Then colEnd = info("UBound2")
    
    ' 边界检查
    If rowStart < info("LBound1") Or rowStart > info("UBound1") Or _
       rowEnd < info("LBound1") Or rowEnd > info("UBound1") Or _
       rowStart > rowEnd Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If
    
    If colStart < info("LBound2") Or colStart > info("UBound2") Or _
       colEnd < info("LBound2") Or colEnd > info("UBound2") Or _
       colStart > colEnd Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If
    
    ' 计算目标数组大小
    Dim rows As Long, cols As Long
    rows = rowEnd - rowStart + 1
    cols = colEnd - colStart + 1
    
    ' 预判数组大小，避免过大
    Dim safeSize As Variant
    safeSize = CalculateSafeArraySize(rows, cols)
    If IsEmpty(safeSize) Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If
    
    ' 如果safeSize小于请求的大小，说明被截断
    If safeSize(1) < rows Or safeSize(2) < cols Then
        rows = safeSize(1)
        cols = safeSize(2)
        ' 调整结束位置
        rowEnd = rowStart + rows - 1
        colEnd = colStart + cols - 1
    End If
    
    ' 创建目标数组
    Dim result() As Variant
    ReDim result(1 To rows, 1 To cols)
    
    ' 逐元素复制（这是最安全的方式）
    Dim r As Long, c As Long
    Dim srcR As Long, srcC As Long
    
    For r = 1 To rows
        srcR = rowStart + r - 1
        For c = 1 To cols
            srcC = colStart + c - 1
            result(r, c) = arr(srcR, srcC)
        Next c
    Next r
    
    SliceArraySafe = result
End Function

' ==============================================================================================
' SECTION 3: 范围/工作表安全操作 / Range and Worksheet Safety
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetNamedRangeSafely
' 安全获取命名范围（带存在性验证）。返回Nothing如果命名范围不存在或无效。
'
' 参数:
'   rangeName - 命名范围名称（必需）
'   sheet - 工作表对象（可选，如果提供则优先在该工作表中查找）
'   wb - 工作簿对象（可选，默认为ThisWorkbook）
'
' 返回: Range对象，或Nothing
' ----------------------------------------------------------------------------------------------
Public Function GetNamedRangeSafely(ByVal rangeName As String, _
                                   Optional sheet As Worksheet, _
                                   Optional wb As Workbook) As Range
    
    ' 默认工作簿
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Dim rng As Range
    
    ' 如果提供了工作表，先尝试在工作表级名称中查找
    If Not sheet Is Nothing Then
        Set rng = sheet.Range(rangeName)
        If Err.Number = 0 Then
            ' 验证范围是否有效（非空）
            If Not rng Is Nothing And rng.Cells.count > 0 Then
                Set GetNamedRangeSafely = rng
            Else
                Set GetNamedRangeSafely = Nothing
            End If
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
    End If
    
    ' 尝试工作簿级名称
    Set rng = wb.Names(rangeName).RefersToRange
    If Err.Number = 0 Then
        ' 验证范围是否有效（非空）
        If Not rng Is Nothing And rng.Cells.count > 0 Then
            Set GetNamedRangeSafely = rng
        Else
            Set GetNamedRangeSafely = Nothing
        End If
    Else
        Set GetNamedRangeSafely = Nothing
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetActualDataRange
' 获取工作表的实际数据范围（避免UsedRange的陷阱）。处理完全空表、只有标题行等情况。
'
' 返回: Range对象，如果没有数据则返回Nothing
' ----------------------------------------------------------------------------------------------
Public Function GetActualDataRange(ws As Worksheet) As Range
    ' 参数验证
    If ws Is Nothing Then
        Set GetActualDataRange = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 查找最后一个非空单元格
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious)
    
    ' 检查是否找到任何数据
    If lastCell Is Nothing Then
        ' 完全空表
        Set GetActualDataRange = Nothing
    Else
        ' 返回从A1到最后一个非空单元格的范围
        Set GetActualDataRange = ws.Range(ws.Cells(1, 1), lastCell)
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsWorksheetValid
' 检查工作表是否存在且可用。可选检查工作表是否可见。
'
' 参数:
'   ws - 工作表对象
'   checkVisible - 是否检查工作表可见性（默认False）
'                  如果为True，则排除隐藏（xlSheetHidden）和深度隐藏（xlSheetVeryHidden）的工作表
'
' 返回: Boolean - True表示工作表有效且满足可见性要求
' ----------------------------------------------------------------------------------------------
Public Function IsWorksheetValid(ws As Worksheet, Optional checkVisible As Boolean = False) As Boolean
    Dim result As Boolean
    result = False
    
    If ws Is Nothing Then
        IsWorksheetValid = False
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 检查工作表是否存在（通过获取名称）
    Dim name As String
    name = ws.name
    result = (Err.Number = 0)
    
    ' 如果需要检查可见性，且工作表存在
    If result And checkVisible Then
        result = (ws.Visible = xlSheetVisible)
    End If
    
    On Error GoTo 0
    
    IsWorksheetValid = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetLastRowSafely
' 安全获取最后一行的行号（考虑各种边界情况）。
'
' 参数:
'   ws - 工作表对象
'   column - 要检查的列（可以是列号或列字母，默认为1）
'
' 返回: Long - 最后一行的行号，如果无数据或无效应列则返回0
' ----------------------------------------------------------------------------------------------
Public Function GetLastRowSafely(ws As Worksheet, Optional column As Variant = 1) As Long
    If ws Is Nothing Then
        GetLastRowSafely = 0
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 将列参数转换为列号（如果是字母）
    Dim colNum As Long
    If IsNumeric(column) Then
        colNum = CLng(column)
    Else
        colNum = Range(column & "1").column
    End If
    
    ' 校验列号是否有效
    If colNum < 1 Or colNum > ws.Columns.count Then
        GetLastRowSafely = 0
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, colNum).End(xlUp).row
    
    ' 如果整个列都空，End(xlUp)会返回第1行
    If ws.Cells(lastRow, colNum).value = "" And lastRow = 1 Then
        lastRow = 0
    End If
    
    GetLastRowSafely = lastRow
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetLastColSafely
' 安全获取最后一列的列号。
'
' 参数:
'   ws - 工作表对象
'   row - 要检查的行（默认为1）
'
' 返回: Long - 最后一列的列号，如果无数据或无效行则返回0
' ----------------------------------------------------------------------------------------------
Public Function GetLastColSafely(ws As Worksheet, Optional row As Variant = 1) As Long
    If ws Is Nothing Then
        GetLastColSafely = 0
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 将行参数转换为行号
    Dim rowNum As Long
    If IsNumeric(row) Then
        rowNum = CLng(row)
    Else
        ' 如果传入的是字符串（如"1"），也尝试转换
        rowNum = val(row)
    End If
    
    ' 校验行号是否有效
    If rowNum < 1 Or rowNum > ws.rows.count Then
        GetLastColSafely = 0
        Exit Function
    End If
    
    Dim lastCol As Long
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).column
    
    ' 如果整行都空，End(xlToLeft)会返回第1列
    If ws.Cells(rowNum, lastCol).value = "" And lastCol = 1 Then
        lastCol = 0
    End If
    
    GetLastColSafely = lastCol
    
    On Error GoTo 0
End Function

' ==============================================================================================
' SECTION 4: 配置管理统一化 / Configuration Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetConfigValue
' 统一的配置获取函数（带类型转换和默认值）。所有模块必须通过此函数获取配置。
'
' 参数:
'   config - 配置字典对象
'   key - 配置键（应使用CFG_*常量）
'   defaultValue - 默认值（如果键不存在或类型转换失败）
'   valueType - 期望的类型（可选，"Long", "Double", "String", "Boolean", "Date", "Integer"）
'               如果提供，将尝试转换值到该类型
'
' 返回: Variant - 配置值，失败返回默认值
' ----------------------------------------------------------------------------------------------
Public Function GetConfigValue(config As Object, _
                              ByVal key As String, _
                              Optional ByVal defaultValue As Variant = EMPTY_VALUE, _
                              Optional ByVal valueType As String = "") As Variant
    
    ' 校验配置对象和键是否存在
    If config Is Nothing Then
        Call LogError("GetConfigValue", "配置对象为空，无法获取键: " & key, "配置错误")
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    If Not config.exists(key) Then
        ' 键不存在，记录调试信息（非错误，可能只是未配置）
        Call LogDebug("GetConfigValue", "配置键不存在，使用默认值: " & key)
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    ' 获取原始值
    Dim rawValue As Variant
    rawValue = config(key)
    
    ' 如果不需要类型转换，直接返回
    If valueType = "" Then
        GetConfigValue = rawValue
        Exit Function
    End If
    
    ' 进行类型转换
    On Error Resume Next
    
    Dim convertedValue As Variant
    convertedValue = EMPTY_VALUE
    
    Select Case UCase(valueType)
        Case "LONG", "LONG INTEGER"
            If IsNumeric(rawValue) Then
                convertedValue = CLng(rawValue)
            End If
            
        Case "DOUBLE", "FLOAT", "SINGLE"
            If IsNumeric(rawValue) Then
                convertedValue = CDbl(rawValue)
            End If
            
        Case "STRING"
            convertedValue = CStr(rawValue)
            
        Case "BOOLEAN", "BOOL"
            If IsNumeric(rawValue) Then
                convertedValue = CBool(rawValue)
            ElseIf VarType(rawValue) = vbString Then
                Dim lowerVal As String
                lowerVal = LCase(CStr(rawValue))
                convertedValue = (lowerVal = "true" Or lowerVal = "yes" Or lowerVal = "1")
            End If
            
        Case "DATE"
            If IsDate(rawValue) Then
                convertedValue = CDate(rawValue)
            End If
            
        Case "INTEGER"
            If IsNumeric(rawValue) Then
                convertedValue = CInt(rawValue)
            End If
            
        Case Else
            ' 未知类型，返回原值
            convertedValue = rawValue
    End Select
    
    ' 检查转换是否成功
    If Err.Number <> 0 Or IsEmpty(convertedValue) Then
        Call LogError("GetConfigValue", "配置键 [" & key & "] 转换为 " & valueType & " 失败，使用默认值", "配置错误")
        GetConfigValue = defaultValue
    Else
        GetConfigValue = convertedValue
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] SetConfigValue
' 安全设置配置值（带类型验证和转换）。所有模块必须通过此函数设置配置。
'
' 参数:
'   config - 配置字典对象
'   key - 配置键（必须使用CFG_*常量）
'   value - 要设置的值
'   valueType - 期望的类型（可选，"Long", "Double", "String", "Boolean"）
'               如果提供，将尝试转换值到该类型
'
' 返回: Boolean - 设置是否成功
' ----------------------------------------------------------------------------------------------
Public Function SetConfigValue(config As Object, _
                              ByVal key As String, _
                              ByVal value As Variant, _
                              Optional ByVal valueType As String = "") As Boolean
    
    ' 统一错误捕获，覆盖全阶段
    On Error Resume Next
    
    ' 验证配置对象
    If config Is Nothing Then
        Call LogError("SetConfigValue", "配置对象为空，无法设置键: " & key, "配置错误")
        SetConfigValue = False
        Exit Function
    End If
    
    ' 验证键名（检查是否以CFG_开头，但非强制，仅记录警告）
    If Left(key, 4) <> "CFG_" Then
        Call LogWarn("SetConfigValue", "配置键 [" & key & "] 未使用 CFG_前缀，建议使用常量", "配置规范")
    End If
    
    ' 根据类型进行验证和转换
    Dim convertedValue As Variant
    convertedValue = value
    
    If valueType <> "" Then
        Select Case valueType
            Case "Long"
                If IsNumeric(value) Then
                    convertedValue = CLng(value)
                Else
                    SetConfigValue = False
                    Exit Function
                End If
                
            Case "Double"
                If IsNumeric(value) Then
                    convertedValue = CDbl(value)
                Else
                    SetConfigValue = False
                    Exit Function
                End If
                
            Case "String"
                convertedValue = CStr(value)
                
            Case "Boolean"
                If IsNumeric(value) Then
                    convertedValue = CBool(value)
                ElseIf VarType(value) = vbString Then
                    Dim lowerVal As String
                    lowerVal = LCase(CStr(value))
                    convertedValue = (lowerVal = "true" Or lowerVal = "yes" Or lowerVal = "1")
                Else
                    SetConfigValue = False
                    Exit Function
                End If
                
            Case Else
                ' 未知类型，保持原值
        End Select
    End If
    
    ' 检查是否已有相同值的配置（可选，防止重复赋值）
    If config.exists(key) Then
        ' 如果值相同，可以视为成功但不记录
        If config(key) = convertedValue Then
            SetConfigValue = True
            Exit Function
        End If
    End If
    
    ' 设置值
    config(key) = convertedValue
    
    ' 检查是否成功
    If Err.Number = 0 Then
        SetConfigValue = True
    Else
        Call LogError("SetConfigValue", "设置配置键 [" & key & "] 失败: " & Err.Description, "配置错误")
        SetConfigValue = False
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] ConfigKeyExists
' 检查配置键是否存在（安全版本）。
' ----------------------------------------------------------------------------------------------
Public Function ConfigKeyExists(config As Object, ByVal key As String) As Boolean
    Dim result As Boolean
    result = False
    
    If config Is Nothing Then
        ConfigKeyExists = False
        Exit Function
    End If
    
    On Error Resume Next
    result = config.exists(key)
    On Error GoTo 0
    
    ConfigKeyExists = result
End Function

' ==============================================================================================
' SECTION 5: Application状态管理 / Application State Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] SaveAppState
' 保存当前Application状态到字典中，供后续恢复使用。
'
' 返回: Dictionary对象，包含以下键：
'       - ScreenUpdating, Calculation, EnableEvents, DisplayAlerts
'       - Cursor, StatusBar, Interactive
' ----------------------------------------------------------------------------------------------
Public Function SaveAppState() As Object
    Dim state As Object
    Set state = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    state("ScreenUpdating") = Application.ScreenUpdating
    state("Calculation") = Application.Calculation
    state("EnableEvents") = Application.EnableEvents
    state("DisplayAlerts") = Application.DisplayAlerts
    state("Cursor") = Application.Cursor
    state("StatusBar") = Application.StatusBar
    state("Interactive") = Application.Interactive
    
    On Error GoTo 0
    
    Set SaveAppState = state
End Function

' ----------------------------------------------------------------------------------------------
' [S] RestoreAppState
' 恢复Application状态（安全版本，处理可能的错误）。采用循环遍历批量恢复。
' ----------------------------------------------------------------------------------------------
Public Sub RestoreAppState(state As Object)
    If state Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    Dim key As Variant
    For Each key In state.Keys
        Select Case key
            Case "ScreenUpdating"
                Application.ScreenUpdating = state(key)
            Case "Calculation"
                Application.Calculation = state(key)
            Case "EnableEvents"
                Application.EnableEvents = state(key)
            Case "DisplayAlerts"
                Application.DisplayAlerts = state(key)
            Case "Cursor"
                Application.Cursor = state(key)
            Case "StatusBar"
                Application.StatusBar = state(key)
            Case "Interactive"
                Application.Interactive = state(key)
            Case Else
                ' 忽略未知键
        End Select
    Next key
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------------------------
' [F] RunOptimized
' 在优化环境中运行宏，并确保状态恢复（无论操作是否成功）。
' 返回 Boolean 表示执行是否成功。
'
' 参数:
'   macroName - 要运行的宏名（格式："模块名.过程名"）
'   optimizeForSpeed - 是否进行速度优化（关闭屏幕更新、事件等）
'
' 返回: Boolean - True表示执行成功，False表示发生错误
' ----------------------------------------------------------------------------------------------
Public Function RunOptimized(ByVal macroName As String, _
                            Optional ByVal optimizeForSpeed As Boolean = True) As Boolean
    
    ' 保存原始状态
    Dim originalState As Object
    Set originalState = SaveAppState()
    
    Dim success As Boolean
    success = True
    
    ' 统一错误捕获，覆盖全阶段
    On Error Resume Next
    
    ' 应用优化设置
    If optimizeForSpeed Then
        Call ApplyOptimizedSettings
    End If
    
    ' 执行宏
    Application.Run macroName
    
    ' 检查是否发生错误
    If Err.Number <> 0 Then
        success = False
        Call LogError("RunOptimized", "执行宏 [" & macroName & "] 时发生错误: " & Err.Description, "错误处理")
    End If
    
    ' 恢复状态
    Call RestoreAppState(originalState)
    
    On Error GoTo 0
    
    RunOptimized = success
End Function

' ----------------------------------------------------------------------------------------------
' [f] ApplyOptimizedSettings
' 私有辅助函数：应用优化设置
' ----------------------------------------------------------------------------------------------
Private Sub ApplyOptimizedSettings()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Cursor = xlWait
End Sub

' ----------------------------------------------------------------------------------------------
' [F] RunOptimizedWithParams
' 带参数的优化运行版本，返回执行状态。支持参数解包，确保宏能正确接收参数。
'
' 参数:
'   macroName - 要运行的宏名（格式："模块名.过程名"）
'   ParamArray args() - 可变参数列表
'
' 返回: Boolean - True表示执行成功，False表示发生错误
'
' 注意: VBA Application.Run 的限制：最多支持 30 个参数。
'       本函数出于安全考虑限制为 10 个参数，已覆盖 99.9% 的使用场景。
'       如需支持更多参数，请在下面的 Select Case 中添加对应分支。
' ----------------------------------------------------------------------------------------------
Public Function RunOptimizedWithParams(ByVal macroName As String, _
                                      ParamArray args() As Variant) As Boolean
    
    ' 保存原始状态
    Dim originalState As Object
    Set originalState = SaveAppState()
    
    Dim success As Boolean
    success = True
    
    ' 统一错误捕获
    On Error Resume Next
    
    ' 应用优化设置（复用私有函数）
    Call ApplyOptimizedSettings
    
    ' 获取参数数量
    Dim paramCount As Long
    paramCount = 0
    
    ' 安全检测 ParamArray 是否有参数
    Dim temp As Long
    temp = UBound(args)  ' 会触发错误，但被 On Error 捕获
    If Err.Number = 0 Then
        paramCount = UBound(args) - LBound(args) + 1
    End If
    Err.Clear
    
    ' 根据参数数量解包执行
    ' 注意：VBA 不支持动态参数传递，必须显式写出每个参数
    ' 以下分支覆盖 0-10 个参数，如需更多请在此扩展
    Select Case paramCount
        Case 0
            Application.Run macroName
            
        Case 1
            Application.Run macroName, args(0)
            
        Case 2
            Application.Run macroName, args(0), args(1)
            
        Case 3
            Application.Run macroName, args(0), args(1), args(2)
            
        Case 4
            Application.Run macroName, args(0), args(1), args(2), args(3)
            
        Case 5
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4)
            
        Case 6
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4), args(5)
            
        Case 7
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4), args(5), args(6)
            
        Case 8
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)
            
        Case 9
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)
            
        Case 10
            Application.Run macroName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)
            
        Case Else
            ' 参数过多，记录错误
            Call LogError("RunOptimizedWithParams", _
                         "参数数量超过 10 个（实际 " & paramCount & "），不支持。如需更多参数，请在代码中添加对应分支。", _
                         "错误处理")
            success = False
    End Select
    
    ' 检查执行是否发生错误
    If Err.Number <> 0 Then
        success = False
        Call LogError("RunOptimizedWithParams", "执行宏 [" & macroName & "] 时发生错误: " & Err.Description, "错误处理")
    End If
    
    ' 恢复状态
    Call RestoreAppState(originalState)
    
    On Error GoTo 0
    
    RunOptimizedWithParams = success
End Function

' ==============================================================================================
' SECTION 6: 安全数学运算 / Safe Math Operations
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] SafeMultiply
' 安全乘法（检查溢出和NaN）。返回Empty如果结果超出安全范围或无效。
' 统一由错误捕获处理所有异常情况，代码更简洁。
' ----------------------------------------------------------------------------------------------
Public Function SafeMultiply(ByVal a As Double, ByVal b As Double, _
                            Optional ByVal maxVal As Double = SAFE_MAX_LONG) As Variant
    
    On Error Resume Next
    
    Dim result As Double
    result = a * b
    
    ' 检查是否发生错误（非数值、溢出等）
    If Err.Number <> 0 Then
        SafeMultiply = EMPTY_VALUE
        Exit Function
    End If
    
    ' 检查结果是否超出安全范围
    If Abs(result) > maxVal Then
        SafeMultiply = EMPTY_VALUE
    Else
        SafeMultiply = result
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] SafeAdd
' 安全加法（检查溢出和NaN）。返回Empty如果结果超出安全范围或无效。
' 统一由错误捕获处理所有异常情况，代码更简洁。
' ----------------------------------------------------------------------------------------------
Public Function SafeAdd(ByVal a As Double, ByVal b As Double, _
                       Optional ByVal maxVal As Double = SAFE_MAX_LONG) As Variant
    
    On Error Resume Next
    
    Dim result As Double
    result = a + b
    
    ' 检查是否发生错误（非数值、溢出等）
    If Err.Number <> 0 Then
        SafeAdd = EMPTY_VALUE
        Exit Function
    End If
    
    ' 检查结果是否超出安全范围
    If Abs(result) > maxVal Then
        SafeAdd = EMPTY_VALUE
    Else
        SafeAdd = result
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] CalculateSafeArraySize
' 计算安全的数组大小（基于元素个数和内存限制）。
' 返回实际可安全分配的大小，如果原始请求过大则截断。
'
' 参数:
'   requestedRows - 请求的行数
'   requestedCols - 请求的列数
'   maxElements - 最大元素个数上限（默认SAFE_MAX_ARRAY_ELEMENTS）
'
' 返回: Variant - 包含两个元素的数组 [safeRows, safeCols]，如果溢出则返回EMPTY_VALUE
' ----------------------------------------------------------------------------------------------
Public Function CalculateSafeArraySize(ByVal requestedRows As Long, _
                                      ByVal requestedCols As Long, _
                                      Optional ByVal maxElements As Long = SAFE_MAX_ARRAY_ELEMENTS) As Variant
    
    ' 前置校验：行列数不能小于1
    If requestedRows < 1 Or requestedCols < 1 Then
        CalculateSafeArraySize = EMPTY_VALUE
        Exit Function
    End If
    
    ' 检查乘法溢出
    Dim totalElements As Variant
    totalElements = SafeMultiply(requestedRows, requestedCols)
    If IsEmpty(totalElements) Then
        CalculateSafeArraySize = EMPTY_VALUE
        Exit Function
    End If
    
    ' 检查是否超过最大元素数
    If totalElements > maxElements Then
        ' 避免除零错误，先取有效列数
        Dim safeCols As Long
        safeCols = IIf(requestedCols > maxElements, maxElements, requestedCols)
        Dim safeRows As Long
        safeRows = maxElements \ safeCols
        ' 兜底：至少保留1行1列
        safeRows = WorksheetFunction.Max(safeRows, 1)
        safeCols = WorksheetFunction.Max(safeCols, 1)
        
        ' 二次校验：确保总元素数不超限
        Do While safeRows * safeCols > maxElements
            If safeRows > 1 Then
                safeRows = safeRows - 1
            Else
                safeCols = safeCols - 1
            End If
        Loop
        
        Dim result(1 To 2) As Long
        result(1) = safeRows
        result(2) = safeCols
        CalculateSafeArraySize = result
    Else
        Dim okResult(1 To 2) As Long
        okResult(1) = requestedRows
        okResult(2) = requestedCols
        CalculateSafeArraySize = okResult
    End If
End Function

' ==============================================================================================
' SECTION 7: 错误处理增强 / Enhanced Error Handling
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] SafeExecute
' 统一的安全执行包装器。无论action中发生什么错误，都能保证：
'   - 错误被捕获并记录
'   - Application状态被恢复
'   - 可选的用户提示
'   - 支持重试机制
'   - 返回执行状态供调用方决策
'
' 参数:
'   moduleName - 调用方模块名
'   procName - 调用方过程名
'   macroName - 要执行的宏名
'   errorMessage - 自定义错误消息（可选）
'   showUser - 是否向用户显示错误（默认True）
'   retryCount - 重试次数（默认0，表示不重试）
'
' 返回: Boolean - True表示执行成功，False表示发生错误
' ----------------------------------------------------------------------------------------------
Public Function SafeExecute(ByVal moduleName As String, _
                           ByVal procName As String, _
                           ByVal macroName As String, _
                           Optional ByVal errorMessage As String = "", _
                           Optional ByVal showUser As Boolean = True, _
                           Optional ByVal retryCount As Long = 0) As Boolean
    
    ' 参数验证
    If macroName = "" Then
        Call LogError(procName, "宏名为空，无法执行", "错误处理")
        SafeExecute = False
        Exit Function
    End If
    
    ' 保存原始状态
    Dim originalState As Object
    Set originalState = SaveAppState()
    
    Dim success As Boolean
    success = False
    
    Dim attempt As Long
    For attempt = 0 To retryCount
        ' 每次重试前重置错误
        On Error GoTo ErrorHandler
        
        ' 执行宏
        Application.Run macroName
        
        ' 如果没有错误，成功
        success = True
        Exit For
        
ErrorHandler:
        ' 记录错误
        Dim errNum As Long
        Dim errDesc As String
        errNum = Err.Number
        errDesc = Err.Description
        
        ' 如果还有重试次数，继续循环
        If attempt < retryCount Then
            ' 短暂等待后重试（避免过快重试）
            Application.Wait Now + TimeValue("00:00:01")
        Else
            ' 最后一次重试也失败，记录最终错误
            Call LogExecutionError(procName, macroName, _
                       errNum, errDesc, errorMessage, attempt + 1)
            
            ' 显示给用户（如果需要）
            If showUser Then
                Dim fullMsg As String
                fullMsg = BuildErrorMessage(procName, errorMessage, errDesc)
                MsgBox fullMsg, vbCritical, "再保险引擎 - 运行时错误"
            End If
        End If
    Next attempt
    
    ' 恢复状态
    Call RestoreAppState(originalState)
    
    SafeExecute = success
End Function

' ----------------------------------------------------------------------------------------------
' [f] LogExecutionError
' 私有辅助函数：记录执行错误
' ----------------------------------------------------------------------------------------------
Private Sub LogExecutionError(ByVal procName As String, _
                             ByVal macroName As String, _
                             ByVal errNum As Long, _
                             ByVal errDesc As String, _
                             ByVal customMsg As String, _
                             Optional ByVal attemptCount As Long = 1)
    
    Dim logMsg As String
    logMsg = "执行宏 [" & macroName & "] 失败"
    
    If attemptCount > 1 Then
        logMsg = logMsg & "（已重试 " & (attemptCount - 1) & " 次）"
    End If
    
    logMsg = logMsg & ": [" & errNum & "] " & errDesc
    
    If customMsg <> "" Then
        logMsg = logMsg & " | " & customMsg
    End If
    
    Call LogError(procName, logMsg, "错误处理")
End Sub

' ----------------------------------------------------------------------------------------------
' [f] BuildErrorMessage
' 私有辅助函数：构建错误消息
' ----------------------------------------------------------------------------------------------
Private Function BuildErrorMessage(ByVal procName As String, _
                                   ByVal customMsg As String, _
                                   ByVal errDesc As String) As String
    
    Dim fullMsg As String
    If customMsg <> "" Then
        fullMsg = customMsg & vbCrLf & vbCrLf & "技术详情: " & errDesc
    Else
        fullMsg = "在执行 [" & procName & "] 时发生错误:" & vbCrLf & vbCrLf & errDesc
    End If
    
    BuildErrorMessage = fullMsg
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsInDesignMode
' 检查是否在设计模式（用于调试断言）。
'
' 参数:
'   forceRefresh - 是否强制刷新状态（默认False）
'                  如果为True，则重新检查并更新静态变量
'
' 返回: Boolean - True表示在设计模式
' ----------------------------------------------------------------------------------------------
Public Function IsInDesignMode(Optional ByVal forceRefresh As Boolean = False) As Boolean
    Static inDesign As Boolean
    Static checked As Boolean
    
    ' 如果需要强制刷新，重置检查标志
    If forceRefresh Then
        checked = False
    End If
    
    ' 如果尚未检查，或需要强制刷新
    If Not checked Then
        On Error Resume Next
        
        ' 检查 VBE 是否可用且主窗口可见
        Dim vbeVisible As Boolean
        vbeVisible = False
        
        ' 先检查 VBE 对象是否存在
        If Not Application.VBE Is Nothing Then
            vbeVisible = (Application.VBE.MainWindow.Visible = True)
        End If
        
        inDesign = vbeVisible
        checked = True
        
        On Error GoTo 0
    End If
    
    IsInDesignMode = inDesign
End Function

' ----------------------------------------------------------------------------------------------
' [S] Assert
' 调试断言（仅在开发环境生效）。条件为False时弹出对话框，支持中断/继续/忽略选项。
'
' 参数:
'   condition - 断言条件
'   message - 自定义消息（可选）
'   logToFile - 是否将断言失败写入日志文件（默认True）
'
' 使用示例:
'   Call Assert(result > 0, "结果应该大于0")
' ----------------------------------------------------------------------------------------------
Public Sub Assert(ByVal condition As Boolean, _
                 Optional ByVal message As String = "", _
                 Optional ByVal logToFile As Boolean = True)
    
    #If DEBUG_MODE Then
        If Not condition Then
            ' 获取调用信息
            Dim callerInfo As String
            callerInfo = GetCallerInfo(2)  ' 获取上一级调用信息
            
            ' 构建完整消息
            Dim fullMsg As String
            fullMsg = "断言失败"
            If message <> "" Then
                fullMsg = fullMsg & ": " & message
            End If
            fullMsg = fullMsg & vbCrLf & vbCrLf & "调用位置: " & callerInfo
            
            ' 输出到立即窗口
            Debug.Print "========== 断言失败 =========="
            Debug.Print "时间: " & Now
            Debug.Print "消息: " & IIf(message = "", "无", message)
            Debug.Print "调用: " & callerInfo
            Debug.Print "==============================="
            
            ' 写入日志文件
            If logToFile Then
                Call LogAssertion(callerInfo, message)
            End If
            
            ' 弹出对话框，提供调试选项
            Dim response As VbMsgBoxResult
            response = MsgBox(fullMsg & vbCrLf & vbCrLf & _
                              "请选择操作：" & vbCrLf & _
                              "   - 是(Y)：进入中断模式（调试）" & vbCrLf & _
                              "   - 否(N)：继续执行" & vbCrLf & _
                              "   - 取消：忽略后续所有断言", _
                              vbYesNoCancel + vbExclamation, "调试断言")
            
            Select Case response
                Case vbYes
                    ' 进入中断模式
                    Stop
                Case vbCancel
                    ' 忽略后续断言（可以通过全局变量控制，这里简化）
                    ' 可以设置一个全局标志，但这里不实现
                Case vbNo
                    ' 继续执行
            End Select
        End If
    #End If
End Sub

' ----------------------------------------------------------------------------------------------
' [f] GetCallerInfo
' 私有辅助函数：获取调用方信息（模块名、过程名、行号）
'
' 参数:
'   depth - 调用栈深度（1表示直接调用方，2表示上一级，以此类推）
' ----------------------------------------------------------------------------------------------
Private Function GetCallerInfo(Optional ByVal depth As Long = 1) As String
    On Error Resume Next
    
    Dim result As String
    result = "未知调用位置"
    
    ' VBA中没有直接获取调用栈的API，但可以通过错误对象获取部分信息
    ' 这里模拟一个简单的实现
    
    ' 方法1：通过 Err 对象获取（需要先产生一个错误）
    Dim savedErrNum As Long
    Dim savedErrDesc As String
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    
    On Error GoTo 0
    
    ' 尝试通过 Application.Caller 获取（仅适用于工作表函数）
    Dim caller As String
    caller = Application.caller
    
    If caller <> "" Then
        result = "调用方: " & caller
    Else
        ' 回退到模块名+过程名（需要手动传入）
        result = "深度 " & depth & " 调用"
    End If
    
    ' 恢复错误状态
    Err.Number = savedErrNum
    Err.Description = savedErrDesc
    
    GetCallerInfo = result
End Function

' ----------------------------------------------------------------------------------------------
' [f] LogAssertion
' 私有辅助函数：将断言失败写入日志文件
' ----------------------------------------------------------------------------------------------
Private Sub LogAssertion(ByVal callerInfo As String, ByVal message As String)
    On Error Resume Next
    
    Dim logPath As String
    logPath = GetAssertLogPath()
    
    If logPath = "" Then
        Exit Sub
    End If
    
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open logPath For Append As #fileNum
    Print #fileNum, "[" & Now & "] 断言失败 - " & callerInfo
    If message <> "" Then
        Print #fileNum, "  消息: " & message
    End If
    Print #fileNum, "----------------------------------------"
    Close #fileNum
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetAssertLogPath
' 私有辅助函数：获取断言日志文件路径
' ----------------------------------------------------------------------------------------------
Private Function GetAssertLogPath() As String
    On Error Resume Next
    
    Dim logFolder As String
    logFolder = Environ("TEMP")
    
    If logFolder = "" Then
        logFolder = "C:\Temp"
    End If
    
    ' 确保文件夹存在
    Call EnsureFolderExists(logFolder)
    
    Dim logFile As String
    logFile = logFolder & "\assert_log.txt"
    
    GetAssertLogPath = logFile
    
    On Error GoTo 0
End Function

' ==============================================================================================
' SECTION 8: 数据类型转换与验证 / Data Type Conversion and Validation
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ToSafeLong
' 安全转换为Long（不抛出错误），并校验范围。
'
' 参数:
'   value - 要转换的值
'   defaultValue - 转换失败或超出范围时返回的默认值（默认0）
'   minVal - 最小值范围（可选，默认-2,147,483,648）
'   maxVal - 最大值范围（可选，默认2,147,483,647）
'
' 返回: Long - 转换后的值，失败或超出范围返回默认值
' ----------------------------------------------------------------------------------------------
Public Function ToSafeLong(value As Variant, _
                          Optional defaultValue As Long = 0, _
                          Optional ByVal minVal As Long = -2147483648#, _
                          Optional ByVal maxVal As Long = 2147483647) As Long
    
    Dim result As Long
    result = defaultValue
    
    On Error Resume Next
    
    ' 检查是否为数值
    If IsNumeric(value) Then
        Dim temp As Double
        temp = CDbl(value)
        
        ' 检查是否在Long范围内
        If temp >= minVal And temp <= maxVal Then
            result = CLng(temp)
        End If
    End If
    
    On Error GoTo 0
    
    ToSafeLong = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] ToSafeDouble
' 安全转换为Double（不抛出错误），并校验范围。
'
' 参数:
'   value - 要转换的值
'   defaultValue - 转换失败或超出范围时返回的默认值（默认0）
'   minVal - 最小值范围（可选，默认-1.7E+308）
'   maxVal - 最大值范围（可选，默认1.7E+308）
'
' 返回: Double - 转换后的值，失败或超出范围返回默认值
' ----------------------------------------------------------------------------------------------
Public Function ToSafeDouble(value As Variant, _
                            Optional defaultValue As Double = 0, _
                            Optional ByVal minVal As Double = -1.7E+308, _
                            Optional ByVal maxVal As Double = 1.7E+308) As Double
    
    Dim result As Double
    result = defaultValue
    
    On Error Resume Next
    
    ' 检查是否为数值
    If IsNumeric(value) Then
        Dim temp As Double
        temp = CDbl(value)
        
        ' 检查是否在Double范围内
        If temp >= minVal And temp <= maxVal Then
            result = temp
        End If
    End If
    
    On Error GoTo 0
    
    ToSafeDouble = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] ToSafeString
' 安全转换为String（不抛出错误）。
'
' 参数:
'   value - 要转换的值
'   defaultValue - 转换失败时返回的默认值（默认空字符串）
'   trimWhitespace - 是否去除首尾空格（默认False）
'   maxLength - 最大长度限制（0表示不限制，默认0）
'
' 返回: String - 转换后的字符串，失败返回默认值
' ----------------------------------------------------------------------------------------------
Public Function ToSafeString(value As Variant, _
                            Optional defaultValue As String = "", _
                            Optional ByVal trimWhitespace As Boolean = False, _
                            Optional ByVal maxLength As Long = 0) As String
    
    Dim result As String
    result = defaultValue
    
    On Error Resume Next
    
    ' 检查是否为Null
    If Not IsNull(value) Then
        result = CStr(value)
        
        ' 去除首尾空格
        If trimWhitespace Then
            result = Trim(result)
        End If
        
        ' 限制长度
        If maxLength > 0 And Len(result) > maxLength Then
            result = Left(result, maxLength)
        End If
    End If
    
    On Error GoTo 0
    
    ToSafeString = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsNumericSafe
' 安全检查是否为数值（不触发错误）。
' ----------------------------------------------------------------------------------------------
Public Function IsNumericSafe(value As Variant) As Boolean
    Dim result As Boolean
    result = False
    
    On Error Resume Next
    result = IsNumeric(value)
    On Error GoTo 0
    
    IsNumericSafe = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsDateSafe
' 安全检查是否为日期（不触发错误）。
' ----------------------------------------------------------------------------------------------
Public Function IsDateSafe(value As Variant) As Boolean
    Dim result As Boolean
    result = False
    
    On Error Resume Next
    result = IsDate(value)
    On Error GoTo 0
    
    IsDateSafe = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] SanitizeString
' 清理字符串中的不可打印字符（AscW<32或=127）。
' 使用正则表达式一次性替换，效率远高于循环拼接。
'
' 参数:
'   text - 原始字符串
'   replacement - 替换字符（默认空字符串，即删除）
'
' 返回: String - 清理后的字符串，失败返回原字符串
' ----------------------------------------------------------------------------------------------
Public Function SanitizeString(ByVal text As String, _
                               Optional ByVal replacement As String = "") As String
    
    ' 空字符串直接返回
    If text = "" Then
        SanitizeString = ""
        Exit Function
    End If
    
    ' 默认返回原字符串（出错时保底）
    SanitizeString = text
    
    On Error Resume Next
    
    ' 创建正则表达式对象
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 检查对象是否创建成功
    If regex Is Nothing Then
        ' 回退到原循环方法（兼容性保障）
        SanitizeString = SanitizeStringFallback(text, replacement)
        Exit Function
    End If
    
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^\x20-\x7E]"  ' 匹配所有不可打印ASCII字符（0-31、127）
    End With
    
    ' 执行替换
    SanitizeString = regex.Replace(text, replacement)
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [f] SanitizeStringFallback
' 私有辅助函数：回退到循环方法（仅在正则不可用时使用）
' ----------------------------------------------------------------------------------------------
Private Function SanitizeStringFallback(ByVal text As String, _
                                        ByVal replacement As String) As String
    Dim result As String
    result = ""
    
    Dim i As Long
    Dim c As String
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        If AscW(c) >= 32 And AscW(c) <> 127 Then
            result = result & c
        Else
            result = result & replacement
        End If
    Next i
    
    SanitizeStringFallback = result
End Function

' ==============================================================================================
' SECTION 9: 路径与文件工具 / Path and File Utilities
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] NormalizePath
' 规范化路径：统一使用反斜杠，清理多余分隔符。
'
' 参数:
'   filePath - 原始路径
'   keepTrailingSlash - 是否保留末尾的反斜杠（默认False）
'                       如果为True，确保路径以反斜杠结尾
'                       如果为False，移除末尾的反斜杠
'
' 返回: String - 规范化后的路径
' ----------------------------------------------------------------------------------------------
Public Function NormalizePath(ByVal filePath As String, _
                             Optional ByVal keepTrailingSlash As Boolean = False) As String
    
    ' 空路径直接返回
    If filePath = "" Then
        NormalizePath = ""
        Exit Function
    End If
    
    Dim result As String
    result = filePath
    
    ' 去除首尾空格
    result = Trim(result)
    
    ' 替换正斜杠为反斜杠
    result = Replace(result, "/", "\")
    
    ' 清理多余的路径分隔符
    Dim original As String
    Do
        original = result
        result = Replace(result, "\\", "\")
    Loop While original <> result
    
    ' 处理末尾分隔符
    If keepTrailingSlash Then
        ' 确保以反斜杠结尾（除非路径为空或已经是根路径如"C:\"）
        If Right(result, 1) <> "\" Then
            result = result & "\"
        End If
    Else
        ' 移除末尾的反斜杠（但保留根路径如"C:\"）
        If Len(result) > 3 And Right(result, 1) = "\" Then
            result = Left(result, Len(result) - 1)
        End If
    End If
    
    NormalizePath = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] SafePathCombine
' 安全组合路径，统一使用反斜杠（\）作为路径分隔符。
' ----------------------------------------------------------------------------------------------
Public Function SafePathCombine(ByVal path1 As String, ByVal path2 As String) As String
    Dim result As String
    
    ' 处理空路径
    If path1 = "" Then
        result = path2
    ElseIf path2 = "" Then
        result = path1
    Else
        ' 先规范化两个路径
        path1 = NormalizePath(path1)
        path2 = NormalizePath(path2)
        
        ' 确保只有一个路径分隔符
        If Right(path1, 1) = "\" Then
            result = path1 & path2
        Else
            result = path1 & "\" & path2
        End If
    End If
    
    ' 最终规范化
    SafePathCombine = NormalizePath(result)
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsNetworkPath
' 判断是否为网络路径（UNC路径或以 \\ 开头）。
' ----------------------------------------------------------------------------------------------
Public Function IsNetworkPath(ByVal filePath As String) As Boolean
    Dim normalized As String
    normalized = NormalizePath(filePath, False)
    
    IsNetworkPath = (Left(normalized, 2) = "\\")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsSharePointPath
' 判断是否为SharePoint路径（支持HTTP/HTTPS和UNC格式）。
'
' SharePoint路径特征：
'   - HTTP/HTTPS: 以 https:// 或 http:// 开头
'   - UNC格式: 以 \\ 开头且包含 "sharepoint.com" 或 ".sharepoint."
' ----------------------------------------------------------------------------------------------
Public Function IsSharePointPath(ByVal filePath As String) As Boolean
    Dim normalized As String
    normalized = NormalizePath(filePath, False)
    
    ' 去除空格并转为小写
    normalized = Trim(LCase(normalized))
    
    ' 情况1：HTTP/HTTPS 路径
    If Left(normalized, 8) = "https://" Or Left(normalized, 7) = "http://" Then
        IsSharePointPath = True
        Exit Function
    End If
    
    ' 情况2：UNC 网络路径
    If Left(normalized, 2) = "\\" Then
        ' 检查是否包含 SharePoint 域名特征
        If InStr(normalized, "sharepoint.com") > 0 Or _
           InStr(normalized, ".sharepoint.") > 0 Then
            IsSharePointPath = True
            Exit Function
        End If
    End If
    
    IsSharePointPath = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] EnsureFolderExists
' 确保文件夹存在，如果不存在则创建。支持多级文件夹的递归创建。
'
' 参数:
'   folderPath - 文件夹路径
'   createParents - 是否创建父文件夹（默认True）
'
' 返回: Boolean - True表示文件夹存在或创建成功，False表示创建失败
' ----------------------------------------------------------------------------------------------
Public Function EnsureFolderExists(ByVal folderPath As String, _
                                  Optional ByVal createParents As Boolean = True) As Boolean
    
    ' 参数验证
    If folderPath = "" Then
        EnsureFolderExists = False
        Exit Function
    End If
    
    ' 规范化路径
    Dim normalized As String
    normalized = NormalizePath(folderPath, False)
    
    On Error Resume Next
    
    ' 检查文件夹是否已存在
    Dim attr As Long
    attr = GetAttr(normalized)
    
    ' 如果存在且是文件夹，直接返回成功
    If Err.Number = 0 And (attr And vbDirectory) Then
        EnsureFolderExists = True
        Exit Function
    End If
    
    ' 清除错误（可能是文件不存在）
    Err.Clear
    
    ' 如果需要创建父文件夹
    If createParents Then
        Dim parentPath As String
        parentPath = GetParentFolderPath(normalized)
        
        If parentPath <> "" Then
            ' 递归创建父文件夹
            If Not EnsureFolderExists(parentPath, True) Then
                EnsureFolderExists = False
                Exit Function
            End If
        End If
    End If
    
    ' 创建当前文件夹
    MkDir normalized
    
    ' 检查是否成功
    If Err.Number = 0 Then
        EnsureFolderExists = True
    Else
        EnsureFolderExists = False
    End If
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetParentFolderPath
' 私有辅助函数：获取父文件夹路径
'
' 参数:
'   folderPath - 文件夹路径（已规范化的路径）
'
' 返回: String - 父文件夹路径，规则：
'       - 普通路径："C:\Folder\Sub" → "C:\Folder"
'       - 根路径："C:\" → ""（无父目录）
'       - 驱动器根："C:" → "C:\"（转换为根路径）
'       - 网络路径："\\server\share" → "\\server"
'       - 网络根："\\server" → ""（无父目录）
'       - 空路径："" → ""
' ----------------------------------------------------------------------------------------------
Private Function GetParentFolderPath(ByVal folderPath As String) As String
    Dim result As String
    result = ""
    
    ' 空路径直接返回
    If folderPath = "" Then
        GetParentFolderPath = ""
        Exit Function
    End If
    
    ' 确保路径已规范化（无多余空格和分隔符）
    Dim path As String
    path = NormalizePath(folderPath, False)
    
    ' 查找最后一个反斜杠
    Dim lastSlash As Long
    lastSlash = InStrRev(path, "\")
    
    If lastSlash > 0 Then
        ' 提取父路径（去掉最后一个反斜杠及其后面的部分）
        result = Left(path, lastSlash - 1)
        
        ' 处理各种边界情况
        Select Case True
            ' 情况1：驱动器根路径（如 "C:"）
            Case Len(result) = 2 And Right(result, 1) = ":"
                result = result & "\"  ' 保持 "C:\" 格式（这是根，不是父目录）
                
            ' 情况2：网络路径根（如 "\\server"）
            Case result = "\\"
                result = ""  ' 网络根没有父目录
                
            ' 情况3：已经是根路径（如 "C:\" 的父目录应该是空）
            Case result = "" Or result = "C:"  ' 这部分由上层逻辑处理
                ' 保持原值
        End Select
    Else
        ' 没有反斜杠，说明是根路径或无效路径
        If Len(path) = 2 And Right(path, 1) = ":" Then
            ' 如 "C:" → 返回 "C:\"（作为根路径）
            result = path & "\"
        ElseIf Left(path, 2) = "\\" Then
            ' 如 "\\server" → 已经是网络根，没有父目录
            result = ""
        Else
            ' 其他情况（如 "Folder"）→ 没有父目录
            result = ""
        End If
    End If
    
    GetParentFolderPath = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetTempFilePath
' 获取临时文件路径（带可选前缀和扩展名）。确保文件路径唯一且可用。
'
' 参数:
'   prefix - 文件名前缀（默认"tmp"）
'   extension - 文件扩展名（默认"tmp"）
'   maxAttempts - 最大尝试次数（默认10）
'
' 返回: String - 唯一的临时文件路径，失败返回空字符串
' ----------------------------------------------------------------------------------------------
Public Function GetTempFilePath(Optional ByVal prefix As String = "tmp", _
                                Optional ByVal extension As String = "tmp", _
                                Optional ByVal maxAttempts As Long = 10) As String
    
    ' 初始化随机种子
    Randomize Timer
    
    ' 获取临时文件夹
    Dim tempFolder As String
    tempFolder = Environ("TEMP")
    If tempFolder = "" Then tempFolder = "C:\Temp"
    
    ' 确保临时文件夹存在
    If Not EnsureFolderExists(tempFolder) Then
        GetTempFilePath = ""
        Exit Function
    End If
    
    ' 处理扩展名（移除前置点）
    extension = Trim(extension)
    If Left(extension, 1) = "." Then extension = Mid(extension, 2)
    
    ' 生成唯一文件名（带重试机制）
    Dim attempt As Long
    Dim filePath As String
    Dim fileExists As Boolean
    
    For attempt = 1 To maxAttempts
        ' 生成毫秒级时间戳
        Dim timestamp As String
        timestamp = FormatDateTimeWithMS(Now, "yyyymmdd_hhnnss_ms")
        
        ' 生成5位随机数
        Dim randomPart As String
        randomPart = Format(Int(Rnd * 100000), "00000")
        
        ' 组合文件名
        Dim fileName As String
        fileName = prefix & "_" & timestamp & "_" & randomPart & "." & extension
        
        ' 组合完整路径
        filePath = SafePathCombine(tempFolder, fileName)
        
        ' 检查文件是否存在
        On Error Resume Next
        Dim attr As Long
        attr = GetAttr(filePath)
        fileExists = (Err.Number = 0)
        On Error GoTo 0
        
        ' 如果文件不存在，使用此路径
        If Not fileExists Then
            GetTempFilePath = filePath
            Exit Function
        End If
        
        ' 短暂等待，让 Timer 变化（避免同一毫秒内重复）
        If attempt < maxAttempts Then
            Application.Wait Now + TimeValue("00:00:00.001")
        End If
    Next attempt
    
    ' 达到最大尝试次数仍未找到可用文件名
    Call LogError("GetTempFilePath", "无法生成唯一的临时文件名，已尝试 " & maxAttempts & " 次", "路径错误")
    GetTempFilePath = ""
End Function

' ----------------------------------------------------------------------------------------------
' [f] FormatDateTimeWithMS
' 私有辅助函数：格式化日期时间，支持毫秒（ms）
'
' 参数:
'   DT - 日期时间值
'   fmt - 格式字符串，支持 "ms" 作为毫秒占位符
'         例如："yyyymmdd_hhnnss_ms" → "20250226_143025_123"
'
' 返回: String - 格式化后的日期时间字符串
' ----------------------------------------------------------------------------------------------
Private Function FormatDateTimeWithMS(ByVal DT As Date, ByVal fmt As String) As String
    Dim result As String
    
    ' 检查是否需要毫秒
    Dim msPos As Long
    msPos = InStr(fmt, "ms")
    
    If msPos > 0 Then
        ' 分离 ms 前后的部分
        Dim beforeMs As String
        Dim afterMs As String
        
        beforeMs = Left(fmt, msPos - 1)
        afterMs = Mid(fmt, msPos + 2)
        
        ' 计算毫秒数
        Dim ms As Long
        ' 一天有 86400 秒 = 86,400,000 毫秒
        ms = (DT - Int(DT)) * 86400000
        ms = ms Mod 1000  ' 确保在 0-999 范围内
        
        ' 分别格式化：
        ' 1. 日期部分（不含 ms）用 VBA.Format 处理
        ' 2. 毫秒部分手动添加
        Dim datePart As String
        If beforeMs <> "" Or afterMs <> "" Then
            ' 有日期格式部分
            datePart = VBA.Format(DT, beforeMs & afterMs)
        Else
            ' 只有 ms，没有其他格式
            datePart = ""
        End If
        
        ' 组合结果：日期部分 + 毫秒
        If datePart = "" Then
            result = Format(ms, "000")
        Else
            result = datePart & Format(ms, "000")
        End If
    Else
        ' 无毫秒要求，直接使用 VBA.Format
        result = VBA.Format(DT, fmt)
    End If
    
    FormatDateTimeWithMS = result
End Function

' ==============================================================================================
' SECTION 10: 性能计时工具 / Performance Timing
' ==============================================================================================

Private pTimers As Object  ' 存储多个计时器

' ----------------------------------------------------------------------------------------------
' [S] StartTimer
' 启动指定名称的计时器。如果同名计时器已存在，将被覆盖。
'
' 参数:
'   timerName - 计时器名称（唯一标识）
' ----------------------------------------------------------------------------------------------
Public Sub StartTimer(ByVal timerName As String)
    If pTimers Is Nothing Then
        Set pTimers = CreateObject("Scripting.Dictionary")
    End If
    
    ' 参数验证
    If timerName = "" Then
        Call LogError("StartTimer", "计时器名称为空，无法启动", "性能监控")
        Exit Sub
    End If
    
    pTimers(timerName) = Timer
End Sub

' ----------------------------------------------------------------------------------------------
' [F] ElapsedTime
' 获取指定计时器的已用时间（秒）。可选是否重置计时器。
'
' 参数:
'   timerName - 计时器名称
'   reset - 是否重置计时器（默认True）
'   defaultValue - 计时器不存在或无效时返回的默认值（默认0）
'
' 返回: Double - 已用时间（秒），失败返回默认值
' ----------------------------------------------------------------------------------------------
Public Function ElapsedTime(ByVal timerName As String, _
                           Optional ByVal reset As Boolean = True, _
                           Optional ByVal defaultValue As Double = 0) As Double
    
    Dim result As Double
    result = defaultValue
    
    ' 参数验证
    If timerName = "" Then
        Call LogError("ElapsedTime", "计时器名称为空", "性能监控")
        ElapsedTime = defaultValue
        Exit Function
    End If
    
    ' 检查计时器集合是否存在
    If pTimers Is Nothing Then
        ElapsedTime = defaultValue
        Exit Function
    End If
    
    ' 检查计时器是否存在
    If Not pTimers.exists(timerName) Then
        ElapsedTime = defaultValue
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim startTime As Double
    startTime = pTimers(timerName)
    
    Dim currentTime As Double
    currentTime = Timer
    
    result = currentTime - startTime
    If result < 0 Then result = result + 86400  ' 处理跨午夜
    
    If reset Then pTimers.Remove timerName
    
    On Error GoTo 0
    
    ElapsedTime = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] FormatElapsedTime
' 格式化时间显示（支持秒、分、时）。
'
' 参数:
'   seconds - 秒数
'   includeMilliseconds - 是否包含毫秒（默认False）
'
' 返回: String - 格式化的时间字符串
' ----------------------------------------------------------------------------------------------
Public Function FormatElapsedTime(ByVal seconds As Double, _
                                  Optional ByVal includeMilliseconds As Boolean = False) As String
    
    Dim result As String
    
    ' 处理负数（不应发生，但防御性处理）
    If seconds < 0 Then seconds = 0
    
    ' 提取小时、分钟、秒
    Dim hours As Long
    Dim mins As Long
    Dim secs As Double
    
    hours = Int(seconds / 3600)
    seconds = seconds - (hours * 3600)
    
    mins = Int(seconds / 60)
    secs = seconds - (mins * 60)
    
    ' 根据时长选择格式
    If hours > 0 Then
        ' 有时：显示 时:分:秒
        If includeMilliseconds Then
            result = hours & " 时 " & mins & " 分 " & Format(secs, "0.00") & " 秒"
        Else
            result = hours & " 时 " & mins & " 分 " & Format(secs, "0") & " 秒"
        End If
    ElseIf mins > 0 Then
        ' 有分无时：显示 分:秒
        If includeMilliseconds Then
            result = mins & " 分 " & Format(secs, "0.00") & " 秒"
        Else
            result = mins & " 分 " & Format(secs, "0") & " 秒"
        End If
    Else
        ' 只有秒
        If includeMilliseconds Then
            result = Format(secs, "0.00") & " 秒"
        Else
            result = Format(secs, "0") & " 秒"
        End If
    End If
    
    FormatElapsedTime = result
End Function

' ----------------------------------------------------------------------------------------------
' [S] LogPerformance
' 记录性能日志（自动计时并写入日志）。支持阈值警告。
'
' 参数:
'   moduleName - 调用方模块名
'   operation - 操作名称
'   timerName - 计时器名称
'   thresholdSeconds - 阈值（秒），超过时记录警告（默认0，不检查）
'   includeMilliseconds - 是否在日志中包含毫秒（默认False）
' ----------------------------------------------------------------------------------------------
Public Sub LogPerformance(ByVal moduleName As String, _
                         ByVal operation As String, _
                         ByVal timerName As String, _
                         Optional ByVal thresholdSeconds As Double = 0, _
                         Optional ByVal includeMilliseconds As Boolean = False)
    
    ' 获取已用时间（自动重置）
    Dim elapsed As Double
    elapsed = ElapsedTime(timerName, True, -1)
    
    ' 检查计时器是否有效
    If elapsed = -1 Then
        Call LogError(operation, "无法获取计时器 [" & timerName & "] 的耗时", "性能监控错误")
        Exit Sub
    End If
    
    ' 格式化时间
    Dim formattedTime As String
    formattedTime = FormatElapsedTime(elapsed, includeMilliseconds)
    
    ' 构建日志消息
    Dim logMsg As String
    logMsg = "操作完成，耗时: " & formattedTime
    
    ' 检查阈值
    Dim logType As String
    logType = "性能监控"
    
    If thresholdSeconds > 0 And elapsed > thresholdSeconds Then
        logMsg = logMsg & " (超过阈值 " & FormatElapsedTime(thresholdSeconds) & ")"
        logType = "性能警告"
        
        ' 额外记录详细日志
        Call LogWarn(operation, "操作耗时超过阈值！实际: " & formattedTime & "，阈值: " & FormatElapsedTime(thresholdSeconds), "性能警告")
    End If
    
    ' 记录日志（根据类型选择不同的日志函数）
    If logType = "性能警告" Then
        Call LogWarn(operation, logMsg, logType)
    Else
        Call LogInfo(operation, logMsg, logType)
    End If
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ClearTimers
' 清空所有计时器，释放内存。
' ----------------------------------------------------------------------------------------------
Public Sub ClearTimers()
    If Not pTimers Is Nothing Then
        pTimers.RemoveAll
        Set pTimers = Nothing
    End If
End Sub

' ----------------------------------------------------------------------------------------------
' [S] GetAllTimers
' 获取所有正在运行的计时器名称列表（调试用）。
'
' 返回: Variant - 计时器名称数组，无计时器时返回空数组
' ----------------------------------------------------------------------------------------------
Public Function GetAllTimers() As Variant
    Dim result() As String
    
    If pTimers Is Nothing Or pTimers.count = 0 Then
        GetAllTimers = Array()
        Exit Function
    End If
    
    ReDim result(0 To pTimers.count - 1)
    
    Dim i As Long
    Dim key As Variant
    i = 0
    For Each key In pTimers.Keys
        result(i) = key
        i = i + 1
    Next key
    
    GetAllTimers = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] TimerExists
' 检查指定名称的计时器是否存在。
'
' 参数:
'   timerName - 计时器名称
'
' 返回: Boolean - True表示存在
' ----------------------------------------------------------------------------------------------
Public Function TimerExists(ByVal timerName As String) As Boolean
    If pTimers Is Nothing Then
        TimerExists = False
        Exit Function
    End If
    
    TimerExists = pTimers.exists(timerName)
End Function
