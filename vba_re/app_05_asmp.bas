Attribute VB_Name = "app_05_asmp"
' ==============================================================================================
' MODULE NAME     : app_05_asmp
' PURPOSE         : 再保险假设编排与精算映射核心模块，实现Treaty与SubLoB的TID关联映射，批量生成分保组假设数据，绑定表单更新按钮执行
'                 : Core module for reinsurance assumption orchestration and actuarial mapping:
'                 : TID-based Treaty-SubLoB mapping, batch generation of reinsurance group data, UI button integration
' DEPENDS         : app_01_basic (Treaty/SubLoB工作表属性、HandleError、MessageStart/MessageEnd、IsWorkSheetExist、FormatSheetStandard、WriteLog 等)
'                 : app_01_basic (Treaty/SubLoB worksheet properties, HandleError, MessageStart/MessageEnd,
'                 : IsWorkSheetExist, FormatSheetStandard, WriteLog, etc.)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'   [C] 输出结果表列索引常量   - 定义输出数组的固定列索引 / Output array column index constants
'   [C] 源数据表列索引常量     - 定义Treaty表和SubLoB表的列索引 / Source table column index constants
'
' SECTION 2: 分保假设编排 / Reinsurance Assumption Orchestration
'   [S] Assumption_UpdateAll    - 批量更新所有定义的分保组工作表，为Update Assumptions按钮提供执行入口
'                              / Batch update all treaty group worksheets, UI button entry point
'   [s] Assumption_Update       - 核心映射引擎，通过字典实现TID高速查找，关联Treaty与SubLoB生成分保假设数据
'                              / Core mapping engine with dictionary-based TID lookup
'
' SECTION 3: 精算逻辑辅助方法 / Actuarial Logic Helpers
'   [f] GetLimitForm            - 根据业务分段和子业务线，判定分保限额适用形式（Risk/Event/Not Applicable/Unknown）
'                              / Determine limit application form based on segment and subLoB
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
Public Const MODULE_NAME As String = "app_05_asmp"

' ----------------------------------------------------------------------------------------------
' [C] 输出结果表 - 列索引常量 / Output array column index constants
' 说明：定义输出数组的固定列索引，统一列结构规范，避免魔法数字
'       Define fixed column indices for output array, avoid magic numbers
' ----------------------------------------------------------------------------------------------
Private Const COL_TREATY_GROUP    As Long = 1      ' 分保组 / Treaty Group
Private Const COL_TID             As Long = 2      ' 分保协议ID / Treaty ID
Private Const COL_TREATY_NAME     As Long = 3      ' 分保协议名称 / Treaty Name
Private Const COL_SEGMENT         As Long = 4      ' 业务分段 / Segment
Private Const COL_MAJOR_LOB       As Long = 5      ' 主业务线 / Major LoB
Private Const COL_SUB_LOB         As Long = 6      ' 子业务线 / Sub LoB
Private Const COL_PERIL           As Long = 7      ' 风险类型 / Peril
Private Const COL_CCY             As Long = 8      ' 币种 / Currency
Private Const COL_LIMIT_RISK      As Long = 9      ' 风险限额 / Limit_Risk
Private Const COL_LIMIT_EVENT     As Long = 10     ' 事故限额 / Limit_Event
Private Const COL_RETENTION       As Long = 11     ' 自留额 / Retention
Private Const COL_SHARE           As Long = 12     ' 分保比例 / Share %
Private Const COL_INURING         As Long = 13     ' 生效顺序 / Inuring Order
Private Const COL_LIMIT_FORM      As Long = 14     ' 限额适用形式 / Limit Application Form
Private Const COL_UNIQUE_KEY      As Long = 15     ' 唯一键 / Unique Key
Private Const COL_COUNT_ASSUMP    As Long = 15     ' 输出结果表总列数 / Total output columns

' ----------------------------------------------------------------------------------------------
' [C] SubLoB表 - 列索引常量 / SubLoB table column index constants
' 说明：定义SubLoB工作表中各列的索引，便于维护和修改
'       Define column indices in SubLoB worksheet for maintainability
' ----------------------------------------------------------------------------------------------
Private Const SUBLOB_COL_SEGMENT  As Long = 2      ' 业务分段列 / Segment column
Private Const SUBLOB_COL_MAJOR    As Long = 3      ' 主业务线列 / Major LoB column
Private Const SUBLOB_COL_SUB      As Long = 4      ' 子业务线列 / Sub LoB column
Private Const SUBLOB_COL_PERIL    As Long = 5      ' 风险类型列 / Peril column
Private Const SUBLOB_TID_START    As Long = 6      ' TID起始列（从第6列开始）/ TID start column

' ----------------------------------------------------------------------------------------------
' [C] Treaty表 - 列索引常量 / Treaty table column index constants
' 说明：定义Treaty工作表中各列的索引，便于维护和修改
'       Define column indices in Treaty worksheet for maintainability
' ----------------------------------------------------------------------------------------------
Private Const TREATY_COL_TID      As Long = 1      ' 分保协议ID列 / TID column
Private Const TREATY_COL_NAME     As Long = 2      ' 分保协议名称列 / Treaty name column
Private Const TREATY_COL_GROUP    As Long = 4      ' 分保组列 / Treaty group column
Private Const TREATY_COL_INURING  As Long = 5      ' 生效顺序列 / Inuring order column
Private Const TREATY_COL_CCY      As Long = 6      ' 币种列 / Currency column
Private Const TREATY_COL_LIMIT_RISK As Long = 7   ' 风险限额列 / Limit_Risk column
Private Const TREATY_COL_LIMIT_EVENT As Long = 8  ' 事故限额列 / Limit_Event column
Private Const TREATY_COL_RETENTION As Long = 9    ' 自留额列 / Retention column
Private Const TREATY_COL_SHARE    As Long = 10     ' 分保比例列 / Share % column

' ----------------------------------------------------------------------------------------------
' [C] 分保组输出表名称后缀 / Reinsurance group output sheet suffix
' ----------------------------------------------------------------------------------------------
Private Const SUFFIX_ASMP         As String = "_Asmp"

' ==============================================================================================
' SECTION 2: 分保假设编排 / Reinsurance Assumption Orchestration
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Assumption_UpdateAll
' 说明：批量更新所有定义的分保组工作表（QS, LoB, MM, Cat, Cyber, Clash），为表单按钮提供执行入口
'       Batch update all treaty group worksheets, entry point for "Update Assumptions" button
' ----------------------------------------------------------------------------------------------
Public Sub Assumption_UpdateAll()
    On Error GoTo ErrHandler
    
    Dim groupNames As Variant
    groupNames = Array("QS", "LoB", "MM", "Cat", "Cyber", "Clash")
    Dim i As Long
    
    Call WriteLog(MODULE_NAME, "Assumption_UpdateAll", "开始批量更新分保组假设 / Starting batch assumption update", "分保编排")
    
    ' 启动全局计时器，关闭屏幕刷新 / Start global timer, disable screen updating
    MessageStart
    
    For i = LBound(groupNames) To UBound(groupNames)
        Application.StatusBar = ">>> 处理分保组 / Processing Group: " & groupNames(i) & _
                                " (" & (i + 1) & "/" & (UBound(groupNames) - LBound(groupNames) + 1) & ")"
        Call Assumption_Update(CStr(groupNames(i)))
    Next i
    
    Application.StatusBar = False
    MessageEnd
    
    Call WriteLog(MODULE_NAME, "Assumption_UpdateAll", "批量更新分保组假设完成 / Batch assumption update completed", "分保编排")
    Exit Sub
    
ErrHandler:
    Application.StatusBar = False
    HandleError MODULE_NAME & ".Assumption_UpdateAll", Err.Description
End Sub

' ----------------------------------------------------------------------------------------------
' [s] Assumption_Update
' 说明：核心映射引擎，通过字典实现TID高速查找，关联Treaty与SubLoB生成分保假设数据
'       Core mapping engine with dictionary-based TID lookup, generate reinsurance assumption data
' 参数：groupName - String，分保组名称（如"QS", "LoB", "Cat"等）/ Treaty group name
' ----------------------------------------------------------------------------------------------
Private Sub Assumption_Update(ByVal groupName As String)
    On Error GoTo ErrHandler
    
    Call WriteLog(MODULE_NAME, "Assumption_Update", "开始处理分保组 / Starting group: " & groupName, "分保编排")
    
    ' 加载源数据 / Load source data
    Dim vTreaty As Variant
    vTreaty = Treaty.Range("rng_Treaty").value
    
    Dim vSubLoB As Variant
    vSubLoB = subLoB.Range("rng_SubLoB").value
    
    ' 验证数据加载是否成功 / Validate data loading
    If IsEmpty(vTreaty) Or Not IsArray(vTreaty) Then
        Call WriteLog(MODULE_NAME, "Assumption_Update", "Treaty数据为空或无效 / Treaty data empty or invalid", "错误")
        Exit Sub
    End If
    
    If IsEmpty(vSubLoB) Or Not IsArray(vSubLoB) Then
        Call WriteLog(MODULE_NAME, "Assumption_Update", "SubLoB数据为空或无效 / SubLoB data empty or invalid", "错误")
        Exit Sub
    End If
    
    ' 创建TID映射字典 / Create TID mapping dictionary
    Dim dictTID As Object
    Set dictTID = CreateObject("Scripting.Dictionary")
    dictTID.CompareMode = vbTextCompare ' 不区分大小写匹配 / Case-insensitive matching
    
    ' 1. 从SubLoB表头映射TID到列索引 / Map TID headers to column indices
    Dim k As Long
    For k = SUBLOB_TID_START To UBound(vSubLoB, 2)
        If Not IsEmpty(vSubLoB(1, k)) Then
            dictTID(CStr(vSubLoB(1, k))) = k
        End If
    Next k
    
    Call WriteLog(MODULE_NAME, "Assumption_Update", "TID映射字典创建完成，共 " & dictTID.count & " 个映射 / TID mapping dictionary created", "分保编排")
    
    ' 2. 初始化输出数组（预分配足够空间）/ Initialize output array
    Dim outData() As Variant
    ReDim outData(1 To (UBound(vTreaty, 1) * 100), 1 To COL_COUNT_ASSUMP)
    Dim outRow As Long
    outRow = 0
    
    Dim i As Long, j As Long
    
    ' 3. 核心映射逻辑 / Core mapping logic
    For i = 1 To UBound(vTreaty, 1)
        ' 按分保组过滤 / Filter by group name
        If vTreaty(i, TREATY_COL_GROUP) = groupName Then
            
            Dim tid As String
            tid = CStr(vTreaty(i, TREATY_COL_TID))
            
            ' 检查TID是否存在映射 / Check if TID exists in mapping
            If dictTID.Exists(tid) Then
                Dim targetCol As Long
                targetCol = dictTID(tid)
                
                ' 扫描SubLoB表的行 / Scan SubLoB rows
                For j = 2 To UBound(vSubLoB, 1)
                    ' 映射标记为1表示该分段受此分保协议覆盖 / Flag '1' indicates coverage
                    If vSubLoB(j, targetCol) = 1 Then
                        outRow = outRow + 1
                        
                        ' 基本标识 / Basic Identifiers
                        outData(outRow, COL_TREATY_GROUP) = groupName              ' 分保组 / Treaty Group
                        outData(outRow, COL_TID) = tid                              ' 分保协议ID / TID
                        outData(outRow, COL_TREATY_NAME) = vTreaty(i, TREATY_COL_NAME)  ' 分保协议名称 / Treaty Name
                        
                        ' 分段元数据（来自SubLoB表）/ Segment Metadata (from SubLoB)
                        outData(outRow, COL_SEGMENT) = vSubLoB(j, SUBLOB_COL_SEGMENT)   ' 业务分段 / Segment
                        outData(outRow, COL_MAJOR_LOB) = vSubLoB(j, SUBLOB_COL_MAJOR)  ' 主业务线 / Major LoB
                        outData(outRow, COL_SUB_LOB) = vSubLoB(j, SUBLOB_COL_SUB)      ' 子业务线 / Sub LoB
                        outData(outRow, COL_PERIL) = vSubLoB(j, SUBLOB_COL_PERIL)      ' 风险类型 / Peril
                        
                        ' 财务条款（来自Treaty表）/ Financial Terms (from Treaty)
                        outData(outRow, COL_CCY) = vTreaty(i, TREATY_COL_CCY)          ' 币种 / Currency
                        outData(outRow, COL_LIMIT_RISK) = vTreaty(i, TREATY_COL_LIMIT_RISK)  ' 风险限额 / Limit_Risk
                        outData(outRow, COL_LIMIT_EVENT) = vTreaty(i, TREATY_COL_LIMIT_EVENT) ' 事故限额 / Limit_Event
                        outData(outRow, COL_RETENTION) = vTreaty(i, TREATY_COL_RETENTION)     ' 自留额 / Retention
                        outData(outRow, COL_SHARE) = vTreaty(i, TREATY_COL_SHARE)            ' 分保比例 / Share %
                        outData(outRow, COL_INURING) = vTreaty(i, TREATY_COL_INURING)         ' 生效顺序 / Inuring Order
                        
                        ' 4. 精算逻辑：确定限额适用形式 / Actuarial Logic: Determine Limit Application Form
                        outData(outRow, COL_LIMIT_FORM) = GetLimitForm( _
                            CStr(vSubLoB(j, SUBLOB_COL_SEGMENT)), _
                            CStr(vSubLoB(j, SUBLOB_COL_SUB)))
                        
                        ' 5. 唯一键：用于Python数据处理 / Unique Key for Python processing
                        outData(outRow, COL_UNIQUE_KEY) = groupName & "|" & tid & "|" & _
                                                          vSubLoB(j, SUBLOB_COL_SEGMENT) & "|" & _
                                                          vSubLoB(j, SUBLOB_COL_PERIL)
                    End If
                Next j
            End If
        End If
    Next i
    
    Call WriteLog(MODULE_NAME, "Assumption_Update", "分保组 " & groupName & " 映射完成，共 " & outRow & " 行数据 / Mapping completed", "分保编排")
    
    ' 6. 输出到工作表 / Output to sheet
    Dim targetSheetName As String
    targetSheetName = groupName & SUFFIX_ASMP
    
    If IsWorkSheetExist(ThisWorkbook, targetSheetName) Then
        Dim wsOut As Worksheet
        Set wsOut = ThisWorkbook.Sheets(targetSheetName)
        
        ' 清除原有数据（保留表头）/ Clear existing data (keep headers)
        wsOut.Range("A2").CurrentRegion.Offset(1, 0).ClearContents
        
        If outRow > 0 Then
            ' 调整到实际行数，避免尾部空行 / Resize to actual rows
            wsOut.Range("A2").Resize(outRow, COL_COUNT_ASSUMP).value = outData
            Call WriteLog(MODULE_NAME, "Assumption_Update", "数据写入成功 / Data written: " & targetSheetName & " [" & outRow & "行]", "工作表IO")
        Else
            Call WriteLog(MODULE_NAME, "Assumption_Update", "无数据写入 / No data to write: " & targetSheetName, "警告")
        End If
        
        ' 标准化工作表格式，格式化数值列 / Standard worksheet formatting
        FormatSheetStandard wsOut, COL_SHARE, COL_LIMIT_FORM
        Call WriteLog(MODULE_NAME, "Assumption_Update", "工作表格式化完成 / Worksheet formatted: " & targetSheetName, "工作表IO")
    Else
        Call WriteLog(MODULE_NAME, "Assumption_Update", "目标工作表不存在 / Target sheet not found: " & targetSheetName, "错误")
    End If
    
    ' 释放对象 / Release objects
    Set dictTID = Nothing
    Set wsOut = Nothing
    
    Call WriteLog(MODULE_NAME, "Assumption_Update", "分保组处理完成 / Group processing completed: " & groupName, "分保编排")
    Exit Sub
    
ErrHandler:
    HandleError MODULE_NAME & ".Assumption_Update", "分保组 / Group [" & groupName & "] 处理失败: " & Err.Description
    Set dictTID = Nothing
    Set wsOut = Nothing
End Sub

' ==============================================================================================
' SECTION 3: 精算逻辑辅助方法 / Actuarial Logic Helpers
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] GetLimitForm
' 说明：根据业务分段和子业务线，判定分保限额适用形式（Risk/Event/Not Applicable/Unknown）
'       Determine limit application form based on segment and subLoB
' 参数：segment - String，业务分段名称 / Segment name
'       subLoB - String，子业务线名称 / Sub LoB name
' 返回值：String - 限额适用形式 / Limit application form
' ----------------------------------------------------------------------------------------------
Private Function GetLimitForm(ByVal segment As String, ByVal subLoB As String) As String
    ' 清除首尾空格，避免脏数据导致匹配失效 / Trim to avoid matching issues
    segment = Trim(segment)
    subLoB = Trim(subLoB)
    
    Dim result As String
    result = "Unknown" ' 默认值 / Default value
    
    Select Case True
        '  attritional损失通常不适用分保限额 / Attritional losses - no reinsurance limits
        Case segment Like "*_Att"
            result = "Not Applicable"
            
        ' 自然灾害分段按事故限额适用 / Natural Catastrophe - per Event
        Case segment = "CAT_Large", segment = "CAT_NM", segment Like "CAT_*"
            result = "Event"
            
        ' 人为重大损失按风险限额适用 / Man-Made Large - per Risk
        Case segment = "MM_Large", segment Like "NonCat_Large*"
            result = "Risk"
            
        ' 特殊场景：RDS（现实灾害情景）/ Special Case: RDS
        Case segment = "MM_RDS"
            ' 根据子业务线名称判定 / Determine based on subLoB name
            If InStr(1, subLoB, "Event", vbTextCompare) > 0 Then
                result = "Event"
            Else
                result = "Risk"
            End If
            
        Case Else
            result = "Unknown"
    End Select
    
    ' 记录调试信息 / Log for debugging (optional)
    If result = "Unknown" Then
        Call WriteLog(MODULE_NAME, "GetLimitForm", "未知限额形式 / Unknown limit form: 分段/Segment=" & segment & ", 子业务线/SubLoB=" & subLoB, "警告")
    End If
    
    GetLimitForm = result
End Function

