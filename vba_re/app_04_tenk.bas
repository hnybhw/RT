Attribute VB_Name = "app_04_tenk"
' ==============================================================================================
' MODULE NAME     : app_04_tenk
' PURPOSE         : 万年模拟数据处理调度核心模块，提供文件路径选择、宽表转长表、SubLoB聚合的可视化操作入口，直接绑定表单按钮执行
'                 : Core scheduling module for 10K simulation data processing: file path selection,
'                 : wide-to-long table conversion, SubLoB aggregation, UI button integration
' DEPENDS         : app_01_basic (GetConfig、HandleError、MessageStart/MessageEnd、FormatSheetStandard、WriteLog 等)
'                 : app_01_basic (GetConfig, HandleError, MessageStart/MessageEnd, FormatSheetStandard, WriteLog, etc.)
'                 : app_02_wb (GetFilePath、ValidateFilePath 等)
'                 : app_02_wb (GetFilePath, ValidateFilePath, etc.)
'                 : app_03_list (WriteToSheet、AppendToSheet 等)
'                 : app_03_list (WriteToSheet, AppendToSheet, etc.)
'                 : C10KProcessor (万年数据处理核心类 / Core 10K data processing class)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'
' SECTION 2: 万年数据文件路径配置 / 10K Data File Path Configuration
'   [S] Specify10KPath          - 弹出文件选择对话框，选择万年模拟数据文件并更新Main工作表路径配置
'                              / Show file dialog to select 10K file and update Main worksheet
'
' SECTION 3: 万年数据核心处理调度 / Core 10K Data Processing
'   [S] TenK_Convert_To_Sheet   - 万年数据宽表转长表调度入口，加载数据并转换为损失清单写入指定工作表
'                              / Wide-to-long table conversion entry point
'   [S] TenK_Aggregate_To_GN    - 万年数据按SubLoB聚合调度入口，加载数据并聚合结果追加到GN工作表
'                              / SubLoB aggregation entry point
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
Public Const MODULE_NAME As String = "app_04_tenk"

' ==============================================================================================
' SECTION 2: 万年数据文件路径配置 / 10K Data File Path Configuration
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Specify10KPath
' 说明：弹出文件选择对话框，供用户选择万年模拟数据文件，选择后更新Main工作表的路径配置单元格
'       Show file dialog for user to select 10K simulation file and update Main worksheet
' 备注：直接绑定表单按钮，无参数无返回值 / UI button binding, no parameters, no return value
' ----------------------------------------------------------------------------------------------
Public Sub Specify10KPath()
    On Error GoTo ErrorHandler
    
    ' 弹出文件选择对话框，限定Excel相关格式 / File dialog with Excel formats
    Dim selectedPath As String
    selectedPath = GetFilePath("选择万年模拟数据文件 / Select 10K Simulation File", _
                               "Excel Files, *.csv;*.xlsb;*.xlsx")
    
    ' 选择有效路径则更新配置并提示 / Update if path selected
    If selectedPath <> "" Then
        Main.Range("ref_10K_FilePath").value = selectedPath
        Call WriteLog(MODULE_NAME, "Specify10KPath", "万年数据文件路径已更新 / 10K file path updated: " & selectedPath, "路径配置")
        
        If ShowMessage Then
            MsgBox "万年数据文件路径已更新: " & selectedPath & vbCrLf & _
                   "10K file path updated: " & selectedPath, vbInformation, "路径配置成功 / Path Configuration Successful"
        End If
    Else
        Call WriteLog(MODULE_NAME, "Specify10KPath", "用户取消选择 / User cancelled selection", "路径配置")
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME & ".Specify10KPath", Err.Description
End Sub

' ==============================================================================================
' SECTION 3: 万年数据核心处理调度 / Core 10K Data Processing
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] TenK_Convert_To_Sheet
' 说明：万年数据宽表转长表核心调度入口，从Main工作表读取配置参数，调用C10KProcessor完成数据加载、转换，
'       结果写入指定工作表并标准化格式化
'       Core entry point for wide-to-long table conversion: read parameters from Main worksheet,
'       call C10KProcessor for data loading and transformation, write results to target worksheet
' 备注：直接绑定表单按钮，无参数无返回值 / UI button binding, no parameters, no return value
' ----------------------------------------------------------------------------------------------
Public Sub TenK_Convert_To_Sheet()
    On Error GoTo ErrHandler
    
    ' 启动计时并设置状态栏 / Start timing and set status bar
    Dim startTime As Double
    startTime = MessageStart("万年数据转长表处理 / 10K Wide-to-Long Conversion")
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "开始万年数据宽表转长表处理 / Starting 10K wide-to-long conversion", "数据处理")
    
    ' 初始化配置并获取核心参数，做默认值兜底 / Get configuration with defaults
    Dim config As Object
    Set config = GetConfig()
    
    Dim matThreshold As Double
    matThreshold = 0.5
    Dim yrMax As Long
    yrMax = 10000
    
    ' 安全读取全局配置，避免配置字典为空导致错误 / Safely read configuration
    On Error Resume Next
    If Not config Is Nothing Then
        matThreshold = CDbl(config("materialityThreshold"))
        yrMax = CLng(config("yearMax"))
    End If
    On Error GoTo 0
    
    ' 配置值有效性校验，确保参数合法 / Validate configuration values
    If matThreshold <= 0 Then matThreshold = 0.5
    If yrMax <= 0 Then yrMax = 10000
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "配置参数 / Config: 阈值/Threshold=" & matThreshold & _
                  ", 最大年份/YearMax=" & yrMax, "数据处理")
    
    ' 从Main工作表读取业务配置参数 / Read parameters from Main worksheet
    Dim filePath As String, segment As String, outputSheetName As String, startRow As Long
    filePath = Main.Range("ref_10K_FilePath").value
    segment = Main.Range("ref_10K_Segment").value
    outputSheetName = Main.Range("ref_10K_LossList").value
    startRow = val(Main.Range("ref_10K_StartRow").value)
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "业务参数 / Params: 文件/File=" & filePath & _
                  ", 分段/Segment=" & segment & ", 输出/Output=" & outputSheetName & ", 起始行/StartRow=" & startRow, "数据处理")
    
    ' 入参有效性校验，拦截无效配置 / Validate input parameters
    If Not ValidateFilePath(filePath) Then
        Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "无效文件路径 / Invalid file path: " & filePath, "错误")
        MsgBox "无效的文件路径: " & filePath & vbCrLf & _
               "Invalid file path: " & filePath, vbCritical, "参数错误 / Parameter Error"
        Exit Sub
    End If
    
    If outputSheetName = "" Then
        Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "输出工作表名称为空 / Output sheet name is empty", "错误")
        MsgBox "输出工作表名称未配置，请填写Main工作表对应单元格" & vbCrLf & _
               "Output sheet name not configured, please fill in Main worksheet", vbCritical, "参数错误 / Parameter Error"
        Exit Sub
    End If
    
    If startRow < 1 Then
        Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "无效起始行 / Invalid start row: " & startRow, "错误")
        MsgBox "数据起始行配置无效，必须大于0" & vbCrLf & _
               "Start row must be greater than 0", vbCritical, "参数错误 / Parameter Error"
        Exit Sub
    End If
    
    ' 创建万年数据处理器实例并配置参数 / Create processor instance and configure
    Dim processor As C10KProcessor
    Set processor = New C10KProcessor
    
    processor.config("materialityThreshold") = matThreshold
    processor.config("yearMax") = yrMax
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "处理器实例创建完成 / Processor instance created", "数据处理")
    
    ' 加载指定文件和分段的万年数据 / Load data from file and segment
    If Not processor.LoadFromFile(filePath, segment) Then
        Err.Raise 998, , "万年数据加载失败，分段不存在或数据无效: " & segment & _
                        " / Failed to load 10K data, segment not found or invalid: " & segment
    End If
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "数据加载成功 / Data loaded successfully: " & _
                  processor.RowCount & "行 x " & processor.ColumnCount & "列", "数据处理")
    
    ' 执行宽表转长表处理，获取转换结果 / Execute transformation
    Dim result As Variant
    result = processor.TransformToLongList(startRow, segment)
    
    ' 结果有效性校验，无数据则主动抛错 / Validate result
    If IsEmpty(result) Then
        Err.Raise 997, , "万年数据转换后无有效数据，请检查源数据和重要性阈值" & _
                        " / No valid data after transformation, check source data and materiality threshold"
    End If
    
    ' 获取结果数组的行数 / Get result row count
    Dim resultRows As Long
    resultRows = UBound(result, 1)
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "数据转换成功 / Transformation successful: " & _
                  (resultRows - 1) & "行结果 / result rows", "数据处理")
    
    ' 新增：提前释放原始矩阵内存，为后续操作腾出空间
    Call processor.ClearRawData
    
    ' 获取结果数组并写入指定工作表 / Get result array and write to worksheet
    Dim convertResult As Variant
    convertResult = processor.GetResult()
    
    If WriteToSheet(convertResult, outputSheetName) Then
        Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "结果写入成功 / Results written to: " & outputSheetName, "工作表IO")
    Else
        Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "结果写入失败 / Failed to write results to: " & outputSheetName, "错误")
    End If
    
    ' 对输出工作表进行标准化格式化，第7列设置千分位 / Format worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(outputSheetName)
    FormatSheetStandard ws, 7, 0
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "工作表格式化完成 / Worksheet formatted: " & outputSheetName, "工作表IO")
    
    ' 释放处理器实例资源 / Release processor
    Set processor = Nothing
    
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "万年数据宽表转长表处理完成 / 10K wide-to-long conversion completed", "数据处理")
    
    ' 新增：在MessageEnd前显示实际转换行数
    Dim completionMessage As String
    completionMessage = "数据处理完成！" & vbCrLf & _
                       "转换结果: " & (resultRows - 1) & " 行" & vbCrLf & _
                       "输出工作表: " & outputSheetName & vbCrLf & _
                       "Data processing completed!" & vbCrLf & _
                       "Results: " & (resultRows - 1) & " rows" & vbCrLf & _
                       "Output sheet: " & outputSheetName
    
    ' 结束计时，恢复Excel环境并弹出处理完成提示（带行数信息）
    MessageEnd startTime, completionMessage
    Exit Sub

ErrHandler:
    ' 集中式错误处理 / Centralized error handling
    Call WriteLog(MODULE_NAME, "TenK_Convert_To_Sheet", "处理失败 / Processing failed: " & Err.Description, "错误")
    HandleError MODULE_NAME & ".TenK_Convert_To_Sheet", Err.Description
    
    ' 异常时确保处理器实例资源释放 / Ensure processor cleanup on error
    If Not processor Is Nothing Then Set processor = Nothing
End Sub

' ----------------------------------------------------------------------------------------------
' [S] TenK_Aggregate_To_GN
' 说明：万年数据按SubLoB聚合核心调度入口，从Main工作表读取配置参数，调用C10KProcessor完成数据加载、聚合，
'       结果追加到GN标准工作表
'       Core entry point for SubLoB aggregation: read parameters from Main worksheet,
'       call C10KProcessor for data loading and aggregation, append results to GN worksheet
' 备注：直接绑定表单按钮，无参数无返回值 / UI button binding, no parameters, no return value
' ----------------------------------------------------------------------------------------------
Public Sub TenK_Aggregate_To_GN()
    On Error GoTo ErrHandler
    
    ' 启动计时并设置状态栏 / Start timing and set status bar
    Dim startTime As Double
    startTime = MessageStart("万年数据按SubLoB聚合处理 / 10K SubLoB Aggregation")
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "开始万年数据按SubLoB聚合处理 / Starting 10K SubLoB aggregation", "数据处理")
    
    ' 初始化配置并获取最大年份参数，做默认值兜底 / Get configuration with defaults
    Dim config As Object
    Set config = GetConfig()
    
    Dim yrMax As Long
    yrMax = 10000
    
    ' 安全读取全局配置，避免配置字典为空导致错误 / Safely read configuration
    On Error Resume Next
    If Not config Is Nothing Then
        yrMax = CLng(config("yearMax"))
    End If
    On Error GoTo 0
    
    ' 配置值有效性校验，确保参数合法 / Validate configuration values
    If yrMax <= 0 Then yrMax = 10000
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "配置参数 / Config: 最大年份/YearMax=" & yrMax, "数据处理")
    
    ' 从Main工作表读取业务配置参数 / Read parameters from Main worksheet
    Dim filePath As String, segment As String, startRow As Long
    filePath = Main.Range("ref_10K_FilePath").value
    segment = Main.Range("ref_10K_Segment").value
    startRow = val(Main.Range("ref_10K_StartRow").value)
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "业务参数 / Params: 文件/File=" & filePath & _
                  ", 分段/Segment=" & segment & ", 起始行/StartRow=" & startRow, "数据处理")
    
    ' 入参有效性校验，拦截无效配置 / Validate input parameters
    If Not ValidateFilePath(filePath) Then
        Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "无效文件路径 / Invalid file path: " & filePath, "错误")
        MsgBox "无效的文件路径: " & filePath & vbCrLf & _
               "Invalid file path: " & filePath, vbCritical, "参数错误 / Parameter Error"
        Exit Sub
    End If
    
    If startRow < 1 Then
        Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "无效起始行 / Invalid start row: " & startRow, "错误")
        MsgBox "数据起始行配置无效，必须大于0" & vbCrLf & _
               "Start row must be greater than 0", vbCritical, "参数错误 / Parameter Error"
        Exit Sub
    End If
    
    ' 创建万年数据处理器实例并配置最大年份参数 / Create processor instance and configure
    Dim processor As C10KProcessor
    Set processor = New C10KProcessor
    
    processor.config("yearMax") = yrMax
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "处理器实例创建完成 / Processor instance created", "数据处理")
    
    ' 加载指定文件和分段的万年数据 / Load data from file and segment
    If Not processor.LoadFromFile(filePath, segment) Then
        Err.Raise 998, , "万年数据加载失败，分段不存在或数据无效: " & segment & _
                        " / Failed to load 10K data, segment not found or invalid: " & segment
    End If
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "数据加载成功 / Data loaded successfully: " & _
                  processor.RowCount & "行 x " & processor.ColumnCount & "列", "数据处理")
    
    ' 执行按SubLoB聚合处理，获取聚合结果 / Execute aggregation
    Dim result As Variant
    result = processor.AggregateBySubLoB(startRow, segment, yrMax, False) ' 默认容错模式 / Default tolerance mode
    
    ' 结果有效性校验，无数据则主动抛错 / Validate result
    If IsEmpty(result) Then
        Err.Raise 997, , "万年数据聚合后无有效数据，请检查源数据和分段配置" & _
                        " / No valid data after aggregation, check source data and segment configuration"
    End If
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "数据聚合成功 / Aggregation successful: " & _
                  UBound(result, 1) & "行结果 / result rows", "数据处理")
    
    ' 获取结果数组并追加到GN标准工作表 / Get result array and append to GN worksheet
    Dim aggResult As Variant
    aggResult = processor.GetResult()
    
    If AppendToSheet(aggResult, SHT_NAME_GN) Then
        Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "结果追加成功 / Results appended to: " & SHT_NAME_GN, "工作表IO")
    Else
        Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "结果追加失败 / Failed to append results to: " & SHT_NAME_GN, "错误")
    End If
    
    ' 释放处理器实例资源 / Release processor
    Set processor = Nothing
    
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "万年数据按SubLoB聚合处理完成 / 10K SubLoB aggregation completed", "数据处理")
    
    ' 结束计时，恢复Excel环境并弹出处理完成提示 / End timing and show completion message
    MessageEnd startTime, "万年数据按SubLoB聚合处理 / 10K SubLoB Aggregation"
    Exit Sub

ErrHandler:
    ' 集中式错误处理 / Centralized error handling
    Call WriteLog(MODULE_NAME, "TenK_Aggregate_To_GN", "处理失败 / Processing failed: " & Err.Description, "错误")
    HandleError MODULE_NAME & ".TenK_Aggregate_To_GN", Err.Description
    
    ' 异常时确保处理器实例资源释放 / Ensure processor cleanup on error
    If Not processor Is Nothing Then Set processor = Nothing
End Sub

