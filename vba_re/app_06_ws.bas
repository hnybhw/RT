Attribute VB_Name = "app_06_ws"
' ==============================================================================================
' MODULE NAME     : app_06_ws
' PURPOSE         : 工作表编排、结果归档与UI维护核心模块，支持弹窗/静默双模式，动作类型常量化便于统一维护
'                 : Core module for worksheet orchestration, result archival and UI maintenance,
'                 : supporting popup/silent dual modes, action type constants for unified maintenance
' DEPENDS         : app_01_basic (Main/EL/RE工作表属性、IsWorkSheetExist、HandleError、WriteLog 等)
'                 : app_01_basic (Main/EL/RE worksheet properties, IsWorkSheetExist, HandleError, WriteLog, etc.)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'   [C] 工作表处理动作类型常量 - 定义工作表归档/删除/保留等动作类型 / Worksheet action type constants
'   [C] 其他常量               - 定义循环限制、列表大小等常量 / Other constants
'
' SECTION 2: 模块级变量 / Module-Level Variables
'   [V] 全局静默模式开关       - 控制弹窗/静默模式 / Global silent mode switch
'
' SECTION 3: 模块初始化 / Module Initialization
'   [s] InitializeModule       - 模块初始化，给全局变量赋默认值 / Module initialization
'
' SECTION 4: 结果归档与工作表清理 / Result Archival & Worksheet Cleanup
'   [S] CommitProcess          - 核心执行入口，模块常量控制弹窗/静默，适配表单控件+批量循环
'                              / Core execution entry, popup/silent controlled by module constant
'   [S] RefreshOutputSheetList - 扫描非永久工作表，更新Main表UI列表并赋予默认动作
'                              / Scan non-permanent sheets, update Main UI list with default actions
'   [S] SmartCleanUp           - 按常量定义的动作类型执行归档/删除逻辑，永久工作表安全锁保护
'                              / Execute archive/delete logic with permanent sheet protection
'   [s] AppendPythonResultToEL - 源工作表数据归档至EL主表，高速内存级数据传输
'                              / Archive source sheet data to EL master with memory transfer
'
' SECTION 5: 主数据清理 / Master Data Clearing
'   [S] ws_Clear_EL            - 清空EL表历史数据，保留表头与格式 / Clear EL data, preserve headers
'   [S] ws_Clear_RE            - 清空RE表历史数据，保留表头与格式 / Clear RE data, preserve headers
'
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
'       1. 双模式控制：修改g_COMMIT_SILENT为True=静默模式（批量循环），False=弹窗模式（表单手动调用）
'       2. 动作类型：所有业务动作均声明为常量，后续修改仅需调整常量值，无需修改业务代码
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 1: 模块常量声明 / Module Constants
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 模块名称常量 - 用于日志记录，避免硬编码 / Module name constant for logging
' ----------------------------------------------------------------------------------------------
Public Const MODULE_NAME As String = "app_06_ws"

' ----------------------------------------------------------------------------------------------
' [C] 工作表处理动作类型常量 / Worksheet action type constants
' 说明：所有业务动作均声明为常量，统一维护，避免硬编码字符串
'       All business actions declared as constants for unified maintenance
' ----------------------------------------------------------------------------------------------
Public Const ACTION_APPEND_DELETE   As String = "AppendAndDelete"  ' 归档后删除 / Archive then delete
Public Const ACTION_DELETE          As String = "Delete"           ' 直接删除 / Direct delete
Public Const ACTION_KEEP            As String = "Keep"             ' 保留工作表 / Keep worksheet

' ----------------------------------------------------------------------------------------------
' [C] 其他常量 / Other constants
' ----------------------------------------------------------------------------------------------
Private Const CONST_PERMANENT_SHEET  As String = "@"               ' 永久工作表标识 / Permanent sheet marker
Private Const CONST_PIVOT_SHEET      As String = "Pivot"           ' 透视表标识 / Pivot table marker
Private Const CONST_MAX_LOOP_COUNT   As Long = 100                 ' 最大循环次数 / Max loop count
Private Const CONST_LIST_ROWS        As Long = 65                  ' 列表最大行数 / Max list rows
Private Const CONST_LIST_COLS        As Long = 2                   ' 列表列数 / List columns

' ==============================================================================================
' SECTION 2: 模块级变量 / Module-Level Variables
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [V] 全局静默模式开关 - 跨模块调用，控制弹窗/静默
'     Global silent mode switch - cross-module, controls popup/silent mode
' ----------------------------------------------------------------------------------------------
Public g_COMMIT_SILENT As Boolean

' ==============================================================================================
' SECTION 3: 模块初始化 / Module Initialization
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [s] InitializeModule
' 说明：模块初始化，给全局变量赋默认值，确保模块首次使用时状态正确
'       Module initialization, set default values for global variables
' ----------------------------------------------------------------------------------------------
Private Sub InitializeModule()
    g_COMMIT_SILENT = False ' 默认非静默模式（手动操作有弹窗）/ Default non-silent mode (popup for manual)
    Call WriteLog(MODULE_NAME, "InitializeModule", "全局静默变量初始化完成 / Global silent variable initialized", "模块初始化")
End Sub

' ==============================================================================================
' SECTION 4: 结果归档与工作表清理 / Result Archival & Worksheet Cleanup
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] CommitProcess
' 说明：工作表归档/清理核心流程，根据ws_created_start_cell区域的列表执行操作
'       Core process for worksheet archive/cleanup, execute based on list in ws_created_start_cell
' ----------------------------------------------------------------------------------------------
Public Sub CommitProcess()
    On Error GoTo ErrHandler
    
    Dim startCell As Range
    Dim rOffset As Long
    Dim wsName As String
    Dim actionType As String
    Dim processCount As Long
    
    ' 确保模块已初始化 / Ensure module is initialized
    If IsEmpty(g_COMMIT_SILENT) Then
        Call InitializeModule
    End If
    
    Call WriteLog(MODULE_NAME, "CommitProcess", "工作表归档/清理流程开始执行 / Starting worksheet archive/cleanup", "工作表处理")
    
    ' 1. 获取起始单元格（Main表ws_created_start_cell）/ Get start cell
    On Error Resume Next
    Set startCell = Main.Range("ws_created_start_cell")
    On Error GoTo 0
    
    If startCell Is Nothing Then
        Call WriteLog(MODULE_NAME, "CommitProcess", "命名区域ws_created_start_cell不存在 / Named range not found", "错误")
        If Not g_COMMIT_SILENT Then
            MsgBox "未找到输出工作表列表起始单元格！" & vbCrLf & _
                   "Named range 'ws_created_start_cell' not found!", vbCritical, "配置错误 / Config Error"
        End If
        Exit Sub
    End If
    
    ' 2. 遍历列表中的每一行，处理工作表 / Iterate through each row in the list
    rOffset = 0
    processCount = 0
    
    Do While startCell.Offset(rOffset, 0).value <> "" And rOffset < CONST_LIST_ROWS
        wsName = startCell.Offset(rOffset, 0).value
        actionType = startCell.Offset(rOffset, 1).value
        
        Call WriteLog(MODULE_NAME, "CommitProcess", "处理工作表 / Processing sheet: " & wsName & ", 动作 / Action: " & actionType, "工作表处理")
        
        ' 检查工作表是否存在 / Check if worksheet exists
        If IsWorkSheetExist(ThisWorkbook, wsName) Then
            Call ExecuteWorksheetAction(wsName, actionType)
            processCount = processCount + 1
        Else
            Call WriteLog(MODULE_NAME, "CommitProcess", "工作表不存在，跳过 / Sheet not found, skipping: " & wsName, "警告")
        End If
        
        rOffset = rOffset + 1
    Loop
    
    ' 3. 流程完成 / Process complete
    If processCount = 0 Then
        Call WriteLog(MODULE_NAME, "CommitProcess", "无待处理工作表 / No worksheets to process", "工作表处理")
        If Not g_COMMIT_SILENT Then
            MsgBox "无待处理的工作表！" & vbCrLf & "No worksheets to process!", vbInformation, "执行完成 / Complete"
        End If
    Else
        Call WriteLog(MODULE_NAME, "CommitProcess", "工作表归档/清理流程执行完成，共处理 " & processCount & " 个工作表 / Process completed, processed " & processCount & " worksheets", "工作表处理")
        
        If Not g_COMMIT_SILENT Then
            MsgBox "工作表处理完成！共处理 " & processCount & " 个工作表。" & vbCrLf & _
                   "Worksheet processing completed! " & processCount & " worksheets processed.", _
                   vbInformation, "执行完成 / Complete"
        End If
    End If
    
    ' 4. 刷新列表（可选，可根据需要决定是否自动刷新）/ Refresh list (optional)
    Call RefreshOutputSheetList
    
    Exit Sub

ErrHandler:
    Call WriteLog(MODULE_NAME, "CommitProcess", "流程执行出错 / Process error: " & Err.Description, "错误")
    If Not g_COMMIT_SILENT Then
        MsgBox "工作表处理出错：" & Err.Description & vbCrLf & _
               "Worksheet processing error: " & Err.Description, vbCritical, "执行失败 / Failed"
    End If
    HandleError MODULE_NAME & ".CommitProcess", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [S] RefreshOutputSheetList
' 说明：刷新Main表中输出工作表列表，从ws_created_start_cell开始向下填充65行，右边一列显示默认动作
'       Refresh output worksheet list in Main sheet, from ws_created_start_cell down 65 rows,
'       right column shows default actions (Keep/Delete/AppendAndDelete)
' ----------------------------------------------------------------------------------------------
Public Sub RefreshOutputSheetList()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim startCell As Range
    Dim rowIdx As Long
    Dim defaultAction As String
    
    ' 确保模块已初始化 / Ensure module is initialized
    If IsEmpty(g_COMMIT_SILENT) Then
        Call InitializeModule
    End If
    
    Call WriteLog(MODULE_NAME, "RefreshOutputSheetList", "开始刷新输出工作表列表 / Starting to refresh output sheet list", "列表更新")
    
    ' 获取起始单元格（Main表ws_created_start_cell）/ Get start cell
    On Error Resume Next
    Set startCell = Main.Range("ws_created_start_cell")
    On Error GoTo 0
    
    If startCell Is Nothing Then
        Call WriteLog(MODULE_NAME, "RefreshOutputSheetList", "命名区域ws_created_start_cell不存在 / Named range not found", "错误")
        If Not g_COMMIT_SILENT Then
            MsgBox "未找到输出工作表列表起始单元格！" & vbCrLf & _
                   "Named range 'ws_created_start_cell' not found!", vbCritical, "配置错误 / Config Error"
        End If
        Exit Sub
    End If
    
    ' 清空原有列表（清除65行2列的区域）/ Clear existing list (65 rows, 2 columns)
    startCell.Resize(CONST_LIST_ROWS, CONST_LIST_COLS).ClearContents
    rowIdx = 0
    
    ' 遍历工作表，填充列表 / Iterate worksheets, populate list
    For Each ws In ThisWorkbook.Worksheets
        ' 跳过永久工作表（含@）/ Skip permanent sheets (contain @)
        If InStr(1, ws.Name, CONST_PERMANENT_SHEET, vbTextCompare) = 0 Then
            
            ' 根据工作表名称确定默认动作 / Determine default action based on name
            If InStr(1, ws.Name, "EL_", vbTextCompare) > 0 Or _
               InStr(1, ws.Name, "RE_", vbTextCompare) > 0 Then
                defaultAction = ACTION_APPEND_DELETE
            ElseIf InStr(1, ws.Name, "TEMP_", vbTextCompare) > 0 Then
                defaultAction = ACTION_DELETE
            Else
                defaultAction = ACTION_KEEP
            End If
            
            ' 写入工作表名称和默认动作 / Write sheet name and default action
            startCell.Offset(rowIdx, 0).value = ws.Name
            startCell.Offset(rowIdx, 1).value = defaultAction
            
            rowIdx = rowIdx + 1
            
            ' 最多填充65行 / Max 65 rows
            If rowIdx >= CONST_LIST_ROWS Then
                Call WriteLog(MODULE_NAME, "RefreshOutputSheetList", "工作表列表超过最大行数 / Exceeded max rows: " & CONST_LIST_ROWS, "警告")
                Exit For
            End If
        End If
    Next ws
    
    Call WriteLog(MODULE_NAME, "RefreshOutputSheetList", "输出工作表列表已刷新，共 " & rowIdx & " 个工作表 / List refreshed", "列表更新")
    Exit Sub

ErrHandler:
    Call WriteLog(MODULE_NAME, "RefreshOutputSheetList", "刷新工作表列表出错 / Error: " & Err.Description, "错误")
    HandleError MODULE_NAME & ".RefreshOutputSheetList", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [S] SmartCleanUp
' 说明：智能清理临时工作表，保留永久表和核心表，按常量定义的动作类型执行
'       Smart cleanup of temporary worksheets, preserve permanent and core sheets
' ----------------------------------------------------------------------------------------------
Public Sub SmartCleanUp()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim loopCount As Long
    Dim deleteCount As Long
    
    ' 确保模块已初始化 / Ensure module is initialized
    If IsEmpty(g_COMMIT_SILENT) Then
        Call InitializeModule
    End If
    
    deleteCount = 0
    loopCount = 0
    
    Call WriteLog(MODULE_NAME, "SmartCleanUp", "智能清理临时工作表流程开始 / Starting smart cleanup of temporary sheets", "工作表清理")
    
    For Each ws In ThisWorkbook.Worksheets
        loopCount = loopCount + 1
        If loopCount > CONST_MAX_LOOP_COUNT Then
            Call WriteLog(MODULE_NAME, "SmartCleanUp", "清理循环超过最大次数 / Exceeded max loop count: " & CONST_MAX_LOOP_COUNT, "警告")
            Exit For
        End If
        
        ' 跳过永久表、核心配置表 / Skip permanent and core sheets
        If InStr(1, ws.Name, CONST_PERMANENT_SHEET, vbTextCompare) > 0 Or _
           ws.Name = "Main" Or ws.Name = "Setup" Or ws.Name = "Archive" Then
            Call WriteLog(MODULE_NAME, "SmartCleanUp", "核心工作表跳过 / Core sheet skipped: " & ws.Name, "清理过滤")
            GoTo NextWS
        End If
        
        ' 清理临时表（含TEMP_前缀）/ Cleanup temporary sheets
        If InStr(1, ws.Name, "TEMP_", vbTextCompare) > 0 Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            deleteCount = deleteCount + 1
            Call WriteLog(MODULE_NAME, "SmartCleanUp", "临时工作表已清理 / Temporary sheet deleted: " & ws.Name, "清理操作")
        End If
        
NextWS:
    Next ws
    
    Call WriteLog(MODULE_NAME, "SmartCleanUp", "智能清理完成，共删除 " & deleteCount & " 个临时工作表 / Cleanup completed, deleted " & deleteCount & " temporary sheets", "清理完成")
    
    If Not g_COMMIT_SILENT Then
        MsgBox "智能清理完成！共删除 " & deleteCount & " 个临时工作表。" & vbCrLf & _
               "Smart cleanup completed! " & deleteCount & " temporary sheets deleted.", _
               vbInformation, "清理完成 / Cleanup Complete"
    End If
    
    Exit Sub

ErrHandler:
    Call WriteLog(MODULE_NAME, "SmartCleanUp", "智能清理出错 / Error: " & Err.Description, "错误")
    If Not g_COMMIT_SILENT Then
        MsgBox "智能清理临时工作表出错：" & Err.Description & vbCrLf & _
               "Smart cleanup error: " & Err.Description, vbCritical, "清理失败 / Cleanup Failed"
    End If
    HandleError MODULE_NAME & ".SmartCleanUp", Err.Description
End Sub

' ==============================================================================================
' SECTION 5: 私有辅助方法 / Private Helper Methods
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] GetWorksheetAction
' 说明：根据工作表名称规则判定处理动作（AppendAndDelete/Delete/Keep）
'       Determine action based on worksheet name rules
' 参数：wsName - String，工作表名称 / Worksheet name
' 返回值：String - 处理动作类型 / Action type
' ----------------------------------------------------------------------------------------------
Private Function GetWorksheetAction(ByVal wsName As String) As String
    On Error GoTo ErrHandler
    
    Dim result As String
    result = ACTION_KEEP ' 默认保留 / Default keep
    
    ' 示例规则：可根据实际业务调整 / Customize rules as needed
    If InStr(1, wsName, "EL_", vbTextCompare) > 0 Or _
       InStr(1, wsName, "RE_", vbTextCompare) > 0 Then
        result = ACTION_APPEND_DELETE
    ElseIf InStr(1, wsName, "TEMP_", vbTextCompare) > 0 Then
        result = ACTION_DELETE
    End If
    
    GetWorksheetAction = result
    Exit Function

ErrHandler:
    GetWorksheetAction = ACTION_KEEP ' 异常时默认保留 / Default keep on error
    Call WriteLog(MODULE_NAME, "GetWorksheetAction", "动作判定出错 / Action determination error: " & Err.Description, "错误")
End Function

' ----------------------------------------------------------------------------------------------
' [f] ExecuteWorksheetAction
' 说明：根据动作类型执行归档删除/直接删除/保留操作
'       Execute archive/delete/keep based on action type
' 参数：wsName - String，工作表名称 / Worksheet name
'       actionType - String，处理动作类型 / Action type
' ----------------------------------------------------------------------------------------------
Private Sub ExecuteWorksheetAction(ByVal wsName As String, ByVal actionType As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim archiveWs As Worksheet
    
    Select Case actionType
        Case ACTION_APPEND_DELETE
            ' 归档逻辑：复制数据到归档表，再删除原表 / Archive: copy to Archive sheet, then delete
            If Not IsWorkSheetExist(ThisWorkbook, "Archive") Then
                ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)).Name = "Archive"
                Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "归档表Archive不存在，已自动创建 / Archive sheet created", "归档操作")
            End If
            
            Set archiveWs = ThisWorkbook.Worksheets("Archive")
            Set ws = ThisWorkbook.Worksheets(wsName)
            
            ' 复制数据到归档表末尾 / Copy data to end of Archive sheet
            ws.UsedRange.Copy archiveWs.Cells(archiveWs.rows.count, 1).End(xlUp).Offset(1, 0)
            Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "工作表数据已归档 / Data archived: " & wsName, "归档操作")
            
            ' 删除原表 / Delete original sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "工作表已删除（归档后）/ Sheet deleted (after archive): " & wsName, "删除操作")
            
        Case ACTION_DELETE
            ' 直接删除 / Direct delete
            Set ws = ThisWorkbook.Worksheets(wsName)
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "工作表已直接删除 / Sheet directly deleted: " & wsName, "删除操作")
            
        Case ACTION_KEEP
            ' 保留工作表，仅记录日志 / Keep sheet, log only
            Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "工作表已保留 / Sheet kept: " & wsName, "保留操作")
    End Select
    
    Exit Sub

ErrHandler:
    Call WriteLog(MODULE_NAME, "ExecuteWorksheetAction", "执行动作 '" & actionType & "' 出错 / Error: " & Err.Description, "错误")
    If Not g_COMMIT_SILENT Then
        MsgBox "处理工作表 '" & wsName & "' 时出错：" & Err.Description & vbCrLf & _
               "Error processing worksheet '" & wsName & "': " & Err.Description, _
               vbExclamation, "操作警告 / Operation Warning"
    End If
End Sub

' ==============================================================================================
' SECTION 6: 主数据清理 / Master Data Clearing
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] ws_Clear_EL
' 说明：清空EL表历史数据，保留表头与格式
'       Clear EL sheet historical data, preserve headers and formatting
' ----------------------------------------------------------------------------------------------
Public Sub ws_Clear_EL()
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    startTime = MessageStart("清空EL表 / Clear EL Sheet")
    
    Call WriteLog(MODULE_NAME, "ws_Clear_EL", "开始清空EL表数据 / Starting to clear EL sheet data", "数据清理")
    
    ' 获取最后一行 / Get last row
    Dim lastR As Long
    lastR = EL.Cells(EL.rows.count, "A").End(xlUp).row
    
    ' 有数据则清空（保留表头）/ Clear if has data (keep header)
    If lastR > 1 Then
        EL.rows("2:" & lastR).ClearContents
        Call WriteLog(MODULE_NAME, "ws_Clear_EL", "EL表数据已清空 / EL sheet cleared: " & (lastR - 1) & "行数据 / rows", "数据清理")
    Else
        Call WriteLog(MODULE_NAME, "ws_Clear_EL", "EL表无数据可清空 / No data to clear in EL sheet", "数据清理")
    End If
    
    MessageEnd startTime, "清空EL表 / Clear EL Sheet"
    
    If ShowMessage Then
        MsgBox "EL表数据已清空！" & vbCrLf & "EL sheet data cleared!", vbInformation, "数据清理完成 / Data Clear Complete"
    End If
    
    Exit Sub

ErrHandler:
    HandleError MODULE_NAME & ".ws_Clear_EL", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [S] ws_Clear_RE
' 说明：清空RE表历史数据，保留表头与格式
'       Clear RE sheet historical data, preserve headers and formatting
' ----------------------------------------------------------------------------------------------
Public Sub ws_Clear_RE()
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    startTime = MessageStart("清空RE表 / Clear RE Sheet")
    
    Call WriteLog(MODULE_NAME, "ws_Clear_RE", "开始清空RE表数据 / Starting to clear RE sheet data", "数据清理")
    
    ' 获取最后一行 / Get last row
    Dim lastR As Long
    lastR = RE.Cells(RE.rows.count, "A").End(xlUp).row
    
    ' 有数据则清空（保留表头）/ Clear if has data (keep header)
    If lastR > 1 Then
        RE.rows("2:" & lastR).ClearContents
        Call WriteLog(MODULE_NAME, "ws_Clear_RE", "RE表数据已清空 / RE sheet cleared: " & (lastR - 1) & "行数据 / rows", "数据清理")
    Else
        Call WriteLog(MODULE_NAME, "ws_Clear_RE", "RE表无数据可清空 / No data to clear in RE sheet", "数据清理")
    End If
    
    MessageEnd startTime, "清空RE表 / Clear RE Sheet"
    
    If ShowMessage Then
        MsgBox "RE表数据已清空！" & vbCrLf & "RE sheet data cleared!", vbInformation, "数据清理完成 / Data Clear Complete"
    End If
    
    Exit Sub

ErrHandler:
    HandleError MODULE_NAME & ".ws_Clear_RE", Err.Description
End Sub

