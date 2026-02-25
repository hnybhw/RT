Attribute VB_Name = "app_03_list"
' ==============================================================================================
' MODULE NAME     : app_03_list
' PURPOSE         : 高速数组IO核心模块，提供数组内存操作、工作表读写、超大数组（超Excel行限制）拆分/合并读写能力，适配类模块数据交互
'                 : Core module for high-speed array I/O: memory operations, worksheet read/write,
'                 : oversized array splitting/merging for Excel row limit, class module integration
' DEPENDS         : app_01_basic (HandleError, IsWorkSheetExist, WriteLog, rowMax/colNew 等基础工具/常量)
'                 : app_01_basic (HandleError, IsWorkSheetExist, WriteLog, rowMax/colNew constants, etc.)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'
' SECTION 2: 基础数组内存操作工具 / Basic Array Memory Operations
'   [F] ArrayFromSheetRange     - 从工作表区域安全读取数据到二维数组，含空值校验
'                              / Safely read worksheet range to 2D array
'   [F] GetLastNonEmptyRow      - 获取二维数组中指定列的最后一个非空行号
'                              / Get last non-empty row in specified column
'   [F] ArrayDimensions         - 安全获取任意数组的维度数量
'                              / Safely get array dimensions
'   [F] IsArrayEmpty            - 校验数组是否为空/未初始化/无效
'                              / Check if array is empty/uninitialized/invalid
'   [F] SliceArray              - 高效截取二维数组的指定行/列区域，生成新数组
'                              / Efficiently slice 2D array to new array
'   [F] AppendArrayVertical     - 垂直拼接两个二维数组（行追加），需列数一致
'                              / Vertically concatenate two 2D arrays
'
' SECTION 3: 常规数组工作表IO / Standard Array Worksheet I/O
'   [F] WriteToSheet            - 通用数组写入工作表（覆盖模式），自动创建工作表
'                              / Write array to worksheet (overwrite mode)
'   [F] AppendToSheet           - 通用数组追加到工作表（追加模式），无表则自动创建
'                              / Append array to worksheet (append mode)
'
' SECTION 4: 超大数组工作表IO / Oversized Array Worksheet I/O
'   [F] WriteArrayToSheet_Overspill - 带溢出处理的数组写入，超行限制时拆分到指定列
'                                   / Write array with overspill handling
'   [s] SplitArrayToSheet       - 私有子过程，核心拆分逻辑
'                              / Private sub for array splitting logic
'   [F] ReadArrayFromSheet_Overspill - 带溢出还原的数组读取
'                                   / Read array with overspill reconstruction
'   [f] MergeArrayFromSheet     - 私有函数，核心合并逻辑
'                              / Private function for array merging logic
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
Public Const MODULE_NAME As String = "app_03_list"

' ==============================================================================================
' SECTION 2: 基础数组内存操作工具 / Basic Array Memory Operations
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ArrayFromSheetRange
' 说明：从指定工作表的指定起始单元格安全读取数据到二维数组，含空区域/无效工作表校验
'       Safely read data from worksheet range to 2D array with validation
' 参数：ws - Worksheet，目标工作表对象 / Target worksheet
'       startCell - String（可选），起始单元格，默认A1 / Start cell, default "A1"
' 返回值：Variant - 读取的二维数组，空区域/无效则返回空数组 / 2D array or empty array if invalid
' ----------------------------------------------------------------------------------------------
Public Function ArrayFromSheetRange(ByVal ws As Worksheet, Optional ByVal startCell As String = "A1") As Variant
    On Error GoTo ErrorHandler
    
    ' 无效工作表直接返回空数组 / Return empty for invalid worksheet
    If ws Is Nothing Then
        Call WriteLog(MODULE_NAME, "ArrayFromSheetRange", "无效工作表对象 / Invalid worksheet object", "警告")
        ArrayFromSheetRange = Array()
        Exit Function
    End If
    
    Dim targetRange As Range
    Set targetRange = ws.Range(startCell).CurrentRegion
    
    ' 校验区域是否为空，空则返回空数组 / Check if range is empty
    If targetRange.rows.count < 1 Or IsEmpty(targetRange.Cells(1, 1)) Then
        Call WriteLog(MODULE_NAME, "ArrayFromSheetRange", "工作表区域为空 / Empty worksheet range: " & ws.Name & "!" & startCell, "警告")
        ArrayFromSheetRange = Array()
        Exit Function
    End If
    
    ' 读取区域数据到数组 / Read range data to array
    ArrayFromSheetRange = targetRange.value
    
    Call WriteLog(MODULE_NAME, "ArrayFromSheetRange", "成功读取数组 / Array read successful: " & ws.Name & "!" & startCell & _
                  " [" & UBound(ArrayFromSheetRange, 1) & "行 x " & UBound(ArrayFromSheetRange, 2) & "列]", "数组操作")
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".ArrayFromSheetRange", Err.Description
    ArrayFromSheetRange = Array()
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetLastNonEmptyRow
' 说明：获取二维数组中指定列的最后一个非空行号，忽略空值/纯空格
'       Get last non-empty row in specified column, ignoring empty/blank values
' 参数：dataArray - Variant，待检测的二维数组 / 2D array to check
'       checkColumn - Long（可选），检测列索引，默认1 / Column index to check, default 1
' 返回值：Long - 最后一个非空行号，数组无效则返回0 / Last non-empty row number, 0 if invalid
' ----------------------------------------------------------------------------------------------
Public Function GetLastNonEmptyRow(ByRef dataArray As Variant, Optional ByVal checkColumn As Long = 1) As Long
    On Error Resume Next
    
    ' 数组无效直接返回0 / Return 0 for invalid array
    If IsArrayEmpty(dataArray) Then
        GetLastNonEmptyRow = 0
        Exit Function
    End If
    
    ' 从后往前遍历，找到第一个非空行 / Scan from bottom to find first non-empty
    Dim r As Long
    For r = UBound(dataArray, 1) To LBound(dataArray, 1) Step -1
        If Not IsEmpty(dataArray(r, checkColumn)) Then
            If Len(Trim(dataArray(r, checkColumn) & "")) > 0 Then
                GetLastNonEmptyRow = r
                Exit Function
            End If
        End If
    Next r
    
    GetLastNonEmptyRow = 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] ArrayDimensions
' 说明：安全获取任意数组的维度数量，非数组返回0
'       Safely get number of dimensions of an array
' 参数：arr - Variant，待检测的数组/变量 / Array to check
' 返回值：Integer - 数组维度数量，非数组/无效则返回0 / Number of dimensions, 0 if not array
' ----------------------------------------------------------------------------------------------
Public Function ArrayDimensions(ByRef arr As Variant) As Integer
    On Error Resume Next
    Dim dimCount As Integer
    dimCount = 0
    
    ' 循环获取维度，触发错误则表示维度结束 / Loop until error indicates end of dimensions
    Do While True
        dimCount = dimCount + 1
        Dim ub As Long
        ub = UBound(arr, dimCount)
        If Err.Number <> 0 Then
            ArrayDimensions = dimCount - 1
            Exit Function
        End If
    Loop
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsArrayEmpty
' 说明：校验数组是否为空/未初始化/维度无效（如行上界<下界）
'       Check if array is empty/uninitialized/invalid
' 参数：arr - Variant，待检测的数组/变量 / Array to check
' 返回值：Boolean - True=数组为空/无效，False=数组有效 / True if empty/invalid, False if valid
' ----------------------------------------------------------------------------------------------
Public Function IsArrayEmpty(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayEmpty = IsEmpty(arr) Or Not IsArray(arr) Or (UBound(arr, 1) < LBound(arr, 1))
End Function

' ----------------------------------------------------------------------------------------------
' [F] SliceArray
' 说明：高效截取二维数组的指定行/列区域，生成新的二维数组，支持行/列范围自定义
'       Efficiently slice 2D array to new array with specified row/column ranges
' 参数：source - Variant，源二维数组 / Source 2D array
'       rowStart - Long，起始行号 / Start row
'       rowEnd - Long，结束行号 / End row
'       colStart - Long（可选），起始列号，默认1 / Start column, default 1
'       colEnd - Long（可选），结束列号，默认源数组最后一列 / End column, default last column
' 返回值：Variant - 截取后的新二维数组，源数组无效则返回空数组 / Sliced 2D array, empty if invalid
' ----------------------------------------------------------------------------------------------
Public Function SliceArray(ByRef source As Variant, ByVal rowStart As Long, ByVal rowEnd As Long, _
                           Optional ByVal colStart As Long = 1, Optional ByVal colEnd As Long = 0) As Variant
    On Error GoTo ErrorHandler
    
    ' 源数组无效直接返回空数组 / Return empty for invalid source
    If IsArrayEmpty(source) Then
        Call WriteLog(MODULE_NAME, "SliceArray", "源数组无效 / Invalid source array", "警告")
        SliceArray = Array()
        Exit Function
    End If
    
    ' 未指定结束列则取源数组最后一列 / Default colEnd to last column
    If colEnd = 0 Then colEnd = UBound(source, 2)
    
    ' 计算截取的行列数 / Calculate slice dimensions
    Dim RowCount As Long: RowCount = rowEnd - rowStart + 1
    Dim colCount As Long: colCount = colEnd - colStart + 1
    
    ' 预分配新数组内存 / Pre-allocate result array
    Dim result As Variant
    ReDim result(1 To RowCount, 1 To colCount)
    
    ' 逐行逐列拷贝数据 / Copy data row by row, column by column
    Dim r As Long, c As Long, srcR As Long, srcC As Long
    For r = 1 To RowCount
        srcR = rowStart + r - 1
        For c = 1 To colCount
            srcC = colStart + c - 1
            result(r, c) = source(srcR, srcC)
        Next c
    Next r
    
    Call WriteLog(MODULE_NAME, "SliceArray", "数组截取成功 / Array sliced successfully: " & _
                  "原数组 / Original [" & UBound(source, 1) & "x" & UBound(source, 2) & "] -> " & _
                  "新数组 / New [" & RowCount & "x" & colCount & "]", "数组操作")
    
    SliceArray = result
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".SliceArray", Err.Description
    SliceArray = Array()
End Function

' ----------------------------------------------------------------------------------------------
' [F] AppendArrayVertical
' 说明：垂直拼接两个二维数组（行追加），要求两个数组列数一致，否则拼接失败并返回原基础数组
'       Vertically concatenate two 2D arrays, requires same column count
' 参数：baseArray - Variant，基础数组（被追加的数组） / Base array
'       appendArray - Variant，追加数组（需拼接的数组） / Array to append
' 返回值：Variant - 拼接后的新二维数组，列数不一致/数组无效则返回原基础数组 / Concatenated array or base array if invalid
' ----------------------------------------------------------------------------------------------
Public Function AppendArrayVertical(ByRef baseArray As Variant, ByRef appendArray As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' 基础数组为空则直接返回追加数组 / Return append array if base is empty
    If IsArrayEmpty(baseArray) Then
        AppendArrayVertical = appendArray
        Exit Function
    End If
    
    ' 追加数组为空则直接返回基础数组 / Return base array if append is empty
    If IsArrayEmpty(appendArray) Then
        AppendArrayVertical = baseArray
        Exit Function
    End If
    
    ' 获取两个数组的行列数 / Get dimensions
    Dim baseRows As Long: baseRows = UBound(baseArray, 1)
    Dim baseCols As Long: baseCols = UBound(baseArray, 2)
    Dim appendRows As Long: appendRows = UBound(appendArray, 1)
    Dim appendCols As Long: appendCols = UBound(appendArray, 2)
    
    ' 列数不一致则触发错误并返回基础数组 / Column mismatch -> return base array
    If baseCols <> appendCols Then
        Call WriteLog(MODULE_NAME, "AppendArrayVertical", "数组列数不匹配 / Column mismatch: " & _
                      "基础数组 / Base [" & baseCols & "列], 追加数组 / Append [" & appendCols & "列]", "错误")
        AppendArrayVertical = baseArray
        Exit Function
    End If
    
    ' 预分配拼接后数组的内存 / Pre-allocate result array
    Dim result As Variant
    ReDim result(1 To baseRows + appendRows, 1 To baseCols)
    
    ' 拷贝基础数组数据 / Copy base array
    Dim r As Long, c As Long
    For r = 1 To baseRows
        For c = 1 To baseCols
            result(r, c) = baseArray(r, c)
        Next c
    Next r
    
    ' 拷贝追加数组数据 / Copy append array
    For r = 1 To appendRows
        For c = 1 To baseCols
            result(r + baseRows, c) = appendArray(r, c)
        Next c
    Next r
    
    Call WriteLog(MODULE_NAME, "AppendArrayVertical", "数组垂直拼接成功 / Arrays concatenated: " & _
                  "基础 / Base [" & baseRows & "行], 追加 / Append [" & appendRows & "行] -> 总计 / Total [" & (baseRows + appendRows) & "行]", "数组操作")
    
    AppendArrayVertical = result
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".AppendArrayVertical", Err.Description
    AppendArrayVertical = baseArray
End Function

' ==============================================================================================
' SECTION 3: 常规数组工作表IO / Standard Array Worksheet I/O
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] WriteToSheet
' 说明：通用数组写入工作表（覆盖模式），目标表不存在则自动创建，全程关闭屏幕刷新提升速度
'       Write array to worksheet (overwrite mode), auto-create worksheet if needed
' 参数：dataArr - Variant，要写入的二维数组 / 2D array to write
'       targetSheetName - String，目标工作表名 / Target worksheet name
'       startCell - String（可选），起始单元格，默认A1 / Start cell, default "A1"
' 返回值：Boolean - True=写入成功，False=写入失败/数组为空 / True if successful, False otherwise
' ----------------------------------------------------------------------------------------------
Public Function WriteToSheet(ByVal dataArr As Variant, ByVal targetSheetName As String, Optional ByVal startCell As String = "A1") As Boolean
    On Error GoTo ErrorHandler
    
    ' 守卫检查：数组为空直接返回失败 / Return False if array is empty
    If IsEmpty(dataArr) Then
        Call WriteLog(MODULE_NAME, "WriteToSheet", "数组为空，写入取消 / Empty array, write cancelled: " & targetSheetName, "警告")
        WriteToSheet = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim rows As Long, cols As Long
    Dim originalScreenUpdating As Boolean
    
    ' Excel环境优化：关闭屏幕刷新提升速度 / Disable screen updating for performance
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 获取/创建工作表 / Get or create worksheet
    If Not IsWorkSheetExist(ThisWorkbook, targetSheetName) Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = targetSheetName
        Call WriteLog(MODULE_NAME, "WriteToSheet", "创建新工作表 / Created new worksheet: " & targetSheetName, "工作表操作")
    Else
        Set ws = ThisWorkbook.Sheets(targetSheetName)
    End If
    
    ' 清空旧数据 + 一次性写入新数据（性能最优）/ Clear existing data and write new data
    ws.Cells.ClearContents
    rows = UBound(dataArr, 1)
    cols = UBound(dataArr, 2)
    ws.Range(startCell).Resize(rows, cols).value = dataArr
    
    ' 恢复环境 / Restore environment
    Application.ScreenUpdating = originalScreenUpdating
    
    Call WriteLog(MODULE_NAME, "WriteToSheet", "数组写入成功 / Array written successfully: " & targetSheetName & _
                  " [" & rows & "行 x " & cols & "列]", "工作表IO")
    
    WriteToSheet = True
    Exit Function

ErrorHandler:
    ' 异常时恢复环境 / Restore environment on error
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    HandleError MODULE_NAME & ".WriteToSheet", Err.Description
    WriteToSheet = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] AppendToSheet
' 说明：通用数组追加到工作表（追加模式），目标表不存在则调用WriteToSheet创建
'       Append array to worksheet, auto-create if not exists
' 参数：dataArr - Variant，要追加的二维数组 / 2D array to append
'       targetSheetName - String，目标工作表名 / Target worksheet name
' 返回值：Boolean - True=追加成功，False=追加失败/数组为空 / True if successful, False otherwise
' ----------------------------------------------------------------------------------------------
Public Function AppendToSheet(ByVal dataArr As Variant, ByVal targetSheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 守卫检查：数组为空直接返回失败 / Return False if array is empty
    If IsEmpty(dataArr) Then
        Call WriteLog(MODULE_NAME, "AppendToSheet", "数组为空，追加取消 / Empty array, append cancelled: " & targetSheetName, "警告")
        AppendToSheet = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim lastRow As Long, rows As Long, cols As Long
    Dim originalScreenUpdating As Boolean
    
    ' Excel环境优化：关闭屏幕刷新提升速度 / Disable screen updating for performance
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 无表则调用WriteToSheet创建，有表则追加 / Create if not exists, otherwise append
    If Not IsWorkSheetExist(ThisWorkbook, targetSheetName) Then
        AppendToSheet = WriteToSheet(dataArr, targetSheetName)
    Else
        Set ws = ThisWorkbook.Sheets(targetSheetName)
        
        ' 计算最后一行（防空表）/ Find last row, handle empty sheet
        lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
        If lastRow < 1 Then lastRow = 1
        
        ' 一次性追加数据 / Append data in one operation
        rows = UBound(dataArr, 1)
        cols = UBound(dataArr, 2)
        ws.Cells(lastRow + 1, 1).Resize(rows, cols).value = dataArr
        
        Call WriteLog(MODULE_NAME, "AppendToSheet", "数组追加成功 / Array appended successfully: " & targetSheetName & _
                      " [" & rows & "行 x " & cols & "列] 至行 / to row " & (lastRow + 1), "工作表IO")
        
        AppendToSheet = True
    End If
    
    ' 恢复环境 / Restore environment
    Application.ScreenUpdating = originalScreenUpdating
    Exit Function

ErrorHandler:
    ' 异常时恢复环境 / Restore environment on error
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    HandleError MODULE_NAME & ".AppendToSheet", Err.Description
    AppendToSheet = False
End Function

' ==============================================================================================
' SECTION 4: 超大数组工作表IO / Oversized Array Worksheet I/O
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] WriteArrayToSheet_Overspill
' 说明：带溢出处理的数组写入，适配超Excel最大行限制的超大数组，超行时自动拆分到colNew指定列（默认AA）
'       Write array with overspill handling for arrays exceeding Excel row limit
' 参数：targetSheet - Worksheet，目标工作表对象 / Target worksheet
'       dataArray - Variant，待写入的二维数组 / 2D array to write
'       startCell - String（可选），主区域起始单元格，默认A1 / Main area start cell, default "A1"
' 返回值：Boolean - True=写入成功，False=写入失败/无效工作表 / True if successful, False otherwise
' ----------------------------------------------------------------------------------------------
Public Function WriteArrayToSheet_Overspill(ByVal targetSheet As Worksheet, ByRef dataArray As Variant, _
                                   Optional ByVal startCell As String = "A1") As Boolean
    On Error GoTo ErrorHandler
    
    ' 无效工作表直接返回失败 / Return False for invalid worksheet
    If targetSheet Is Nothing Then
        Call WriteLog(MODULE_NAME, "WriteArrayToSheet_Overspill", "无效工作表对象 / Invalid worksheet object", "错误")
        WriteArrayToSheet_Overspill = False
        Exit Function
    End If
    
    ' 数组为空则清空工作表并返回成功 / Clear worksheet if array is empty
    If IsArrayEmpty(dataArray) Then
        targetSheet.Cells.ClearContents
        Call WriteLog(MODULE_NAME, "WriteArrayToSheet_Overspill", "数组为空，清空工作表 / Empty array, cleared worksheet: " & targetSheet.Name, "工作表IO")
        WriteArrayToSheet_Overspill = True
        Exit Function
    End If
    
    ' 清空工作表原有内容 / Clear existing content
    targetSheet.Cells.ClearContents
    
    Dim RowCount As Long, colCount As Long
    RowCount = UBound(dataArray, 1)
    colCount = UBound(dataArray, 2)
    
    ' 超行限制则调用拆分逻辑，否则直接常规写入 / Split if exceeds row limit, otherwise normal write
    If RowCount > rowMax Then
        Call SplitArrayToSheet(targetSheet, dataArray, startCell, RowCount, colCount)
        Call WriteLog(MODULE_NAME, "WriteArrayToSheet_Overspill", "超大数组拆分写入 / Oversized array split written: " & _
                      targetSheet.Name & " [" & RowCount & "行 > " & rowMax & "行限制]", "工作表IO")
    Else
        targetSheet.Range(startCell).Resize(RowCount, colCount).value = dataArray
        Call WriteLog(MODULE_NAME, "WriteArrayToSheet_Overspill", "数组常规写入 / Normal array write: " & _
                      targetSheet.Name & " [" & RowCount & "行 x " & colCount & "列]", "工作表IO")
    End If
    
    WriteArrayToSheet_Overspill = True
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".WriteArrayToSheet_Overspill", Err.Description
    WriteArrayToSheet_Overspill = False
End Function

' ----------------------------------------------------------------------------------------------
' [s] SplitArrayToSheet
' 说明：私有核心拆分逻辑，将超大数组拆分为主区域（A列开始，最大行）+溢出区域（colNew列开始）写入工作表
'       Private core splitting logic: main area (A column) + overflow area (colNew column)
' 参数：ws - Worksheet，目标工作表对象 / Target worksheet
'       dataArray - Variant，待拆分的超大二维数组 / Oversized 2D array to split
'       startCell - String，主区域起始单元格 / Main area start cell
'       totalRows - Long，数组总行数 / Total rows in array
'       colCount - Long，数组总列数 / Total columns in array
' ----------------------------------------------------------------------------------------------
Private Sub SplitArrayToSheet(ByVal ws As Worksheet, ByRef dataArray As Variant, _
                                     ByVal startCell As String, ByVal totalRows As Long, ByVal colCount As Long)
    
    ' 主区域写入最大行数据 / Write main area up to row limit
    ws.Range(startCell).Resize(rowMax, colCount).value = dataArray
    
    ' 计算溢出数据行数并截取溢出数组 / Calculate overflow rows and slice array
    Dim overflowCount As Long
    overflowCount = totalRows - rowMax
    Dim overflowArray As Variant
    overflowArray = SliceArray(dataArray, rowMax + 1, totalRows, 1, colCount)
    
    ' 溢出数据写入到colNew指定列（默认AA）/ Write overflow to colNew column
    ws.Range(colNew & "1").Resize(overflowCount, colCount).value = overflowArray
    
    Call WriteLog(MODULE_NAME, "SplitArrayToSheet", "数组拆分完成 / Array split complete: " & _
                  "主区域 / Main [" & rowMax & "行], 溢出区域 / Overflow [" & overflowCount & "行] 至列 / to column " & colNew, "工作表IO")
End Sub

' ----------------------------------------------------------------------------------------------
' [F] ReadArrayFromSheet_Overspill
' 说明：带溢出还原的数组读取，将工作表中主区域（A列开始）+溢出区域（colNew列开始）的数据合并为完整数组
'       Read array with overspill reconstruction from worksheet
' 参数：sourceSheet - Worksheet，源工作表对象 / Source worksheet
'       spillColName - String（可选），溢出区域起始列名，默认AA / Overflow column name, default "AA"
' 返回值：Variant - 合并后的完整二维数组，无效工作表则返回空数组 / Reconstructed 2D array, empty if invalid
' ----------------------------------------------------------------------------------------------
Public Function ReadArrayFromSheet_Overspill(ByVal sourceSheet As Worksheet, Optional ByVal spillColName As String = "AA") As Variant
    On Error GoTo ErrorHandler
    
    ' 无效工作表直接返回空数组 / Return empty for invalid worksheet
    If sourceSheet Is Nothing Then
        Call WriteLog(MODULE_NAME, "ReadArrayFromSheet_Overspill", "无效工作表对象 / Invalid worksheet object", "错误")
        ReadArrayFromSheet_Overspill = Array()
        Exit Function
    End If
    
    ' 获取主区域数据 / Get main area data
    Dim mainRange As Range
    Set mainRange = sourceSheet.Range("A1").CurrentRegion
    
    ' 检查溢出区域是否有数据 / Check if overflow area has data
    Dim lastSpillRow As Long
    lastSpillRow = sourceSheet.Cells(sourceSheet.rows.count, spillColName).End(xlUp).row
    
    ' 无溢出数据则直接返回主区域数组，有则调用合并逻辑 / Return main array if no overflow, otherwise merge
    If lastSpillRow < 1 Or IsEmpty(sourceSheet.Range(spillColName & "1").value) Then
        ReadArrayFromSheet_Overspill = mainRange.value
        Call WriteLog(MODULE_NAME, "ReadArrayFromSheet_Overspill", "读取常规数组 / Normal array read: " & sourceSheet.Name, "工作表IO")
    Else
        ReadArrayFromSheet_Overspill = MergeArrayFromSheet(sourceSheet, mainRange, spillColName, lastSpillRow)
        Call WriteLog(MODULE_NAME, "ReadArrayFromSheet_Overspill", "读取拆分数组 / Split array read: " & sourceSheet.Name, "工作表IO")
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".ReadArrayFromSheet_Overspill", Err.Description
    ReadArrayFromSheet_Overspill = Array()
End Function

' ----------------------------------------------------------------------------------------------
' [f] MergeArrayFromSheet
' 说明：私有核心合并逻辑，将工作表主区域和溢出区域的数组数据垂直拼接为完整的超大数组
'       Private core merging logic: vertically concatenate main and overflow arrays
' 参数：ws - Worksheet，源工作表对象 / Source worksheet
'       mainRange - Range，主区域范围 / Main area range
'       spillColName - String，溢出区域起始列名 / Overflow column name
'       lastSpillRow - Long，溢出区域最后一行行号 / Last row of overflow area
' 返回值：Variant - 合并后的完整二维数组 / Reconstructed 2D array
' ----------------------------------------------------------------------------------------------
Private Function MergeArrayFromSheet(ByVal ws As Worksheet, ByRef mainRange As Range, _
                                      ByVal spillColName As String, ByVal lastSpillRow As Long) As Variant
    
    ' 读取主区域和溢出区域数据到数组 / Read main and overflow areas to arrays
    Dim vMain As Variant
    vMain = mainRange.value
    
    Dim mainRows As Long, mainCols As Long
    mainRows = UBound(vMain, 1)
    mainCols = UBound(vMain, 2)
    
    Dim spillRange As Range
    Set spillRange = ws.Range(spillColName & "1").Resize(lastSpillRow, mainCols)
    Dim vSpill As Variant
    vSpill = spillRange.value
    
    ' 计算合并后总行数并预分配内存 / Calculate total rows and pre-allocate
    Dim totalRows As Long
    totalRows = mainRows + UBound(vSpill, 1)
    
    Dim result As Variant
    ReDim result(1 To totalRows, 1 To mainCols)
    
    ' 拷贝主区域数据 / Copy main area data
    Dim r As Long, c As Long
    For r = 1 To mainRows
        For c = 1 To mainCols
            result(r, c) = vMain(r, c)
        Next c
    Next r
    
    ' 拷贝溢出区域数据 / Copy overflow area data
    For r = 1 To UBound(vSpill, 1)
        For c = 1 To mainCols
            result(r + mainRows, c) = vSpill(r, c)
        Next c
    Next r
    
    Call WriteLog(MODULE_NAME, "MergeArrayFromSheet", "数组合并完成 / Array merge complete: " & _
                  "主区域 / Main [" & mainRows & "行], 溢出区域 / Overflow [" & UBound(vSpill, 1) & "行] -> 总计 / Total [" & totalRows & "行]", "数组操作")
    
    MergeArrayFromSheet = result
End Function

