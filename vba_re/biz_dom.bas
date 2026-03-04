Attribute VB_Name = "biz_dom"
' ==============================================================================================
' MODULE NAME       : biz_dom
' LAYER             : Business
' PURPOSE           : Business Domain layer. Owns all pure array transformation logic for the
'                     reinsurance analysis platform. Three responsibility areas:
'                     (1) 10K wide-to-long conversion for Python LOSS schema consumption,
'                     (2) SubLoB aggregation for GN sheet append,
'                     (3) Treaty x SubLoB assumption mapping (TID cross-reference engine).
'                     No reinsurance calculation logic. No worksheet IO.
' DEPENDS           : core_utils    v2.1.1 (EMPTY_VALUE, IsEmptyValue)
'                     plat_runtime  v2.1.2 (LogInfo, LogWarn, LogError)
' NOTE              : - All reinsurance calculation logic (XL, QS, inuring, event aggregation,
'                       net recovery) is owned exclusively by the Python engine
'                       (RI_Engine_Core_v2.py). This module must never replicate it.
'                     - biz_io owns all worksheet reads and writes. This module receives
'                       pre-loaded Variant arrays and returns Variant arrays only.
'                     - All public functions are query-only: no IO, no worksheet access,
'                       no Application state mutation.
'                     - TransformToLongList output column order aligns 1-to-1 with Python
'                       SCHEMA["LOSS"] to eliminate VBA-to-Python re-mapping:
'                         seq(0)=LossSeq, colume(1)=Colume, year(2)=Year,
'                         sub_lob(3)=SubLoB, peril(4)=Peril, event(5)=Event,
'                         amount(6)=Amount, tid(7)=TID, group(8)=Group
'                     - TransformToLongList uses segment-aware field parsing (p_ParseRowFields)
'                       to derive Peril / Event / Group per the three parser rules originally
'                       implemented as class modules CRDSParser / CCATParser / CMMParser:
'
'                         CAT_*            Peril = matrix(2,c)  [row 2 of header band]
'                                          Event = matrix(3,c)  [row 3 of header band]
'                                          Group = "CAT"        [fixed]
'
'                         RDS              Peril = matrix(2,c)  [row 2 of header band]
'                                          Event = matrix(3,c)  [row 3 of header band]
'                                          Group = "Aviation"   if SubLoB contains "Aviation"
'                                                  "Marine"     if SubLoB contains "Marine",
'                                                               "PVT", or "Offshore"
'                                                  empty        otherwise
'
'                         NonCat_Large*    Peril = "MM_Large"   [derived from segment]
'                                          Event = empty
'                                          Group = empty
'
'                         NonCat_Attrit.   Peril = "MM_Att"     [derived from segment]
'                                          Event = empty
'                                          Group = empty
'
'                       Matrix header band layout for CAT_* and RDS:
'                         row 1 = SubLoB labels   (all segment types)
'                         row 2 = Peril labels    (CAT_* and RDS only)
'                         row 3 = Event labels    (CAT_* and RDS only)
'                         row 4+ = loss data      => startRow >= 4 for these segments
'                       MM matrix header band:
'                         row 1 = SubLoB labels   (single header row)
'                         row 2+ = loss data      => startRow >= 2
'
'                       The class-module parser hierarchy (IDataParser / CParserFactory /
'                       CCATParser / CMMParser / CRDSParser) is superseded by the inline
'                       dispatch in p_ParseRowFields. No class modules needed in v2.
'
'                     - BuildAssumptionArray receives two pre-loaded arrays (Treaty, SubLoB)
'                       and returns the fully mapped assumption array. IO and entry stay out.
'                     - ResolveLimitForm is the canonical limit form implementation;
'                       supersedes app_05_asmp.GetLimitForm (v1 private function).
' STATUS            : Draft
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Domain transformation logic was distributed across
'                             C10KProcessor (class module), app_04_tenk, and app_05_asmp;
'                             no unified domain boundary existed.
'   - Init (Design): Wide-to-long conversion owned by C10KProcessor with segment-aware
'                    parsing delegated to class hierarchy (IDataParser / CParserFactory /
'                    CCATParser / CMMParser / CRDSParser); assumption mapping inline in
'                    app_05_asmp.Assumption_Update.
'   - Init (Scope): Domain logic entangled with IO and orchestration; no query-only contract;
'                   no EMPTY_VALUE failure sentinel.

' v2.0.0
'   - Init (Architecture): Introduced biz_dom as the Business Domain transformation module
'                          under the three-layer model (Core / Platform / Business).
'   - Init (Boundary): Wide-to-Long conversion, SubLoB aggregation, and assumption mapping
'                      are all pure array transforms; all belong here. IO and entry stay out.
'   - Init (Contract): Boolean + ByRef errMsg for all public functions; EMPTY_VALUE for
'                      array-returning failures; no exceptions raised, no silent fallbacks.
'   - Init (Schema Alignment): TransformToLongList output column order aligned 1-to-1 with
'                              Python SCHEMA["LOSS"] to eliminate VBA-to-Python re-mapping.
'   - Init (Parser Consolidation): Segment-aware Peril/Event/Group derivation rules from
'                                  CRDSParser / CCATParser / CMMParser inlined into private
'                                  function p_ParseRowFields; class-module hierarchy superseded.
'   - Init (Multi-row Header): CAT_* and RDS matrices carry Peril (row 2) and Event (row 3)
'                              in the header band; startRow contract is >= 4 for these
'                              segments. MM matrices are single-header; startRow >= 2.
'   - Init (Assumption Mapping): Extracted Treaty x SubLoB TID cross-reference engine from
'                                app_05_asmp into BuildAssumptionArray; IO and entry decoupled.
'   - Init (LimitForm Authority): ResolveLimitForm declared the canonical implementation;
'                                 supersedes app_05_asmp.GetLimitForm (v1 private function).
'   - Init (Query-only): All public functions are side-effect free except plat_runtime
'                        logging calls (non-business side effects).
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 00: MODULE CONSTANTS
'
' SECTION 01: WIDE-TO-LONG TRANSFORMATION
'   [F] TransformToLongList     - Convert 10K wide matrix to Python LOSS schema long list
'
' SECTION 02: SUBLOB AGGREGATION
'   [F] AggregateToGN           - Aggregate 10K wide matrix by SubLoB across simulation years
'
' SECTION 03: INPUT VALIDATION
'   [F] ValidateTenKMatrix      - Validate 10K source matrix for structural pre-conditions
'
' SECTION 04: ASSUMPTION MAPPING
'   [F] BuildAssumptionArray    - Cross-reference Treaty x SubLoB to produce assumption array
'   [F] ResolveLimitForm        - Resolve limit application form from segment / sub-LoB
'
' SECTION 05: PRIVATE HELPERS
'   [f] p_ParseRowFields        - Derive Peril / Event / Group per segment parser rules
'
' ==============================================================================================
' NOTE: [C]=Constant, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
'       Rule: Private helpers inherit the Contract and Side Effects of their calling public
'             function unless explicitly stated otherwise.
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 00: MODULE CONSTANTS
' ==============================================================================================

Private Const BIZ_LAYER     As String = "BIZ"
Private Const THIS_MODULE   As String = "biz_dom"

' -----------------------------------------------------------------------
' 10K LONG-LIST schema constants
' -----------------------------------------------------------------------

' Output column count for TransformToLongList.
' Aligns with Python SCHEMA["LOSS"]: 9 fields (0-indexed Python = 1-indexed VBA).
Private Const LONG_COL_COUNT        As Long = 9

' Output column indices for TransformToLongList (1-based VBA).
' MUST match Python SCHEMA["LOSS"] offset + 1 exactly.
Private Const LC_SEQ                As Long = 1    ' seq    (0)
Private Const LC_COLUME             As Long = 2    ' colume (1)
Private Const LC_YEAR               As Long = 3    ' year   (2)
Private Const LC_SUB_LOB            As Long = 4    ' sub_lob(3)
Private Const LC_PERIL              As Long = 5    ' peril  (4)
Private Const LC_EVENT              As Long = 6    ' event  (5)
Private Const LC_AMOUNT             As Long = 7    ' amount (6)
Private Const LC_TID                As Long = 8    ' tid    (7) - empty; Python fills via Asmp
Private Const LC_GROUP              As Long = 9    ' group  (8)

' Header strings for TransformToLongList row 1 (match Python SCHEMA["LOSS"] keys).
Private Const LH_SEQ                As String = "LossSeq"
Private Const LH_COLUME             As String = "Colume"
Private Const LH_YEAR               As String = "Year"
Private Const LH_SUB_LOB            As String = "SubLoB"
Private Const LH_PERIL              As String = "Peril"
Private Const LH_EVENT              As String = "Event"
Private Const LH_AMOUNT             As String = "Amount"
Private Const LH_TID                As String = "TID"
Private Const LH_GROUP              As String = "Group"

' Default materiality threshold: losses at or below this value are suppressed.
Private Const DEFAULT_MATERIALITY   As Double = 0.5

' -----------------------------------------------------------------------
' GN AGGREGATION schema constants
' -----------------------------------------------------------------------

' Fixed header column count in AggregateToGN output before year columns begin.
' Layout: col 1 = SubLoB, col 2 = Peril, col 3 = Mean, col 4..N = Year 1..yearMax.
Private Const GN_HEADER_COLS        As Long = 3

' -----------------------------------------------------------------------
' PARSER DISPATCH constants  (used by p_ParseRowFields)
' -----------------------------------------------------------------------

' Wide-matrix header band row positions.
' CAT_* and RDS carry Peril in row 2 and Event in row 3 of the column header band.
' MM matrices have a single header row; Peril is derived from the segment name string.
Private Const HDR_ROW_SUBLOB        As Long = 1    ' SubLoB label row  (all segment types)
Private Const HDR_ROW_PERIL         As Long = 2    ' Peril label row   (CAT_* and RDS only)
Private Const HDR_ROW_EVENT         As Long = 3    ' Event label row   (CAT_* and RDS only)

' Fixed Group tag for all CAT_* segments.
Private Const PARSER_GROUP_CAT      As String = "CAT"

' Fixed Peril string constants for MM segments (derived from segment name, not matrix rows).
Private Const PARSER_PERIL_MM_LARGE As String = "MM_Large"
Private Const PARSER_PERIL_MM_ATT   As String = "MM_Att"

' RDS Group keyword constants.
' Evaluation order: Aviation checked first; Marine keywords checked second.
Private Const PARSER_GROUP_AVIATION As String = "Aviation"
Private Const PARSER_GROUP_MARINE   As String = "Marine"
' RDS Marine sub-keywords (any match → Group = "Marine").
Private Const PARSER_KW_PVT         As String = "PVT"
Private Const PARSER_KW_OFFSHORE    As String = "Offshore"

' -----------------------------------------------------------------------
' ASSUMPTION MAPPING schema constants
' -----------------------------------------------------------------------

' Output column count for BuildAssumptionArray (15-column schema from app_05_asmp).
Private Const ASMP_COL_COUNT        As Long = 15

' Output column indices for BuildAssumptionArray (1-based VBA).
' Align with Python SCHEMA["ASMP"] offset + 1.
Private Const AC_GROUP              As Long = 1    ' Treaty Group
Private Const AC_TID                As Long = 2    ' Treaty ID
Private Const AC_TNAME              As Long = 3    ' Treaty Name
Private Const AC_SEGMENT            As Long = 4    ' Segment
Private Const AC_MAJOR_LOB          As Long = 5    ' Major LoB
Private Const AC_SUB_LOB            As Long = 6    ' Sub LoB
Private Const AC_PERIL              As Long = 7    ' Peril
Private Const AC_CCY                As Long = 8    ' Currency
Private Const AC_LIMIT_RISK         As Long = 9    ' Limit_Risk
Private Const AC_LIMIT_EVENT        As Long = 10   ' Limit_Event
Private Const AC_RETENTION          As Long = 11   ' Retention
Private Const AC_SHARE              As Long = 12   ' Share %
Private Const AC_INURING            As Long = 13   ' Inuring Order
Private Const AC_LIMIT_FORM         As Long = 14   ' Limit Application Form
Private Const AC_UNIQUE_KEY         As Long = 15   ' Unique Key: Group|TID|Segment|Peril

' Treaty array (rng_Treaty) source column indices.
Private Const TC_TID                As Long = 1
Private Const TC_NAME               As Long = 2
Private Const TC_GROUP              As Long = 4
Private Const TC_INURING            As Long = 5
Private Const TC_CCY                As Long = 6
Private Const TC_LIMIT_RISK         As Long = 7
Private Const TC_LIMIT_EVENT        As Long = 8
Private Const TC_RETENTION          As Long = 9
Private Const TC_SHARE              As Long = 10

' SubLoB array (rng_SubLoB) source column indices.
Private Const SC_SEGMENT            As Long = 2
Private Const SC_MAJOR              As Long = 3
Private Const SC_SUB                As Long = 4
Private Const SC_PERIL              As Long = 5
Private Const SC_TID_START          As Long = 6    ' TID mapping flag columns start here

' Limit form canonical string constants.
Private Const LF_RISK               As String = "Risk"
Private Const LF_EVENT              As String = "Event"
Private Const LF_NA                 As String = "Not Applicable"
Private Const LF_UNKNOWN            As String = "Unknown"

' ==============================================================================================
' SECTION 01: WIDE-TO-LONG TRANSFORMATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] TransformToLongList
'
' 功能说明      : 将 10K 宽表矩阵转换为符合 Python SCHEMA["LOSS"] 字段顺序的 Long Form 数组
'               : 使用段感知字段解析（p_ParseRowFields）派生每行的 Peril / Event / Group
'               : 三类矩阵的表头结构不同（见 NOTE 中的 Parser 规则）：
'                   CAT_* / RDS  : 三行表头带（SubLoB / Peril / Event）=> startRow >= 4
'                   MM           : 单行表头 => startRow >= 2
'               : 损失金额 <= materialityThreshold 的格被过滤，不写入输出
'               : TID 列留空，由 Python 引擎通过 Assumption 映射填入
'               : 输出数组由 biz_io.WriteArrayToSheet 写入目标工作表后供 Python 消费
' 参数          : wideMatrix            - 二维 Variant 数组（来自 biz_io；含表头行）
'               : startRow              - 数据起始行（1-based；由调用方从 Named Range 读取）
'               : segment               - 分段标签，驱动 Peril/Event/Group 派生逻辑
'               : outLongList           - 输出：Long Form 二维 Variant 数组（含表头行）
'               : errMsg                - 输出：失败时的错误说明
'               : materialityThreshold  - 可选：重要性阈值（默认 0.5）
' 返回          : Boolean - True=转换成功，outLongList 可用；False=失败，errMsg 已填充
' Purpose       : Transform 10K wide matrix to Python-consumable LOSS schema long list
' Contract      : Business / Domain (query-only; no IO, no Application state mutation)
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function TransformToLongList(ByRef wideMatrix As Variant, _
                                     ByVal startRow As Long, _
                                     ByVal Segment As String, _
                                     ByRef outLongList As Variant, _
                                     ByRef errMsg As String, _
                                     Optional ByVal materialityThreshold As Double = DEFAULT_MATERIALITY) As Boolean
    errMsg = vbNullString
    TransformToLongList = False
    outLongList = core_utils.EMPTY_VALUE

    ' --- Pre-condition validation ---
    Dim matErr As String
    If Not ValidateTenKMatrix(wideMatrix, startRow, matErr) Then
        errMsg = THIS_MODULE & ".TransformToLongList: " & matErr
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "TransformToLongList", errMsg
        Exit Function
    End If

    Dim rMax As Long: rMax = UBound(wideMatrix, 1)
    Dim cMax As Long: cMax = UBound(wideMatrix, 2)

    ' Worst-case pre-allocation: all data cells pass materiality filter (+1 for header).
    Dim estimatedRows As Long
    estimatedRows = (rMax - startRow + 1) * (cMax - 1) + 1

    Dim temp() As Variant
    ReDim temp(1 To estimatedRows, 1 To LONG_COL_COUNT)

    ' --- Write header row ---
    temp(1, LC_SEQ) = LH_SEQ
    temp(1, LC_COLUME) = LH_COLUME
    temp(1, LC_YEAR) = LH_YEAR
    temp(1, LC_SUB_LOB) = LH_SUB_LOB
    temp(1, LC_PERIL) = LH_PERIL
    temp(1, LC_EVENT) = LH_EVENT
    temp(1, LC_AMOUNT) = LH_AMOUNT
    temp(1, LC_TID) = LH_TID
    temp(1, LC_GROUP) = LH_GROUP

    ' --- Main conversion loop: wide → long ---
    ' wideMatrix col 1      = simulation year value (all segment types)
    ' wideMatrix(1, c)      = SubLoB label (all segment types)
    ' wideMatrix(2, c)      = Peril label  (CAT_* and RDS only)
    ' wideMatrix(3, c)      = Event label  (CAT_* and RDS only)
    ' wideMatrix(r, c) r>=startRow = loss amount
    Dim count As Long: count = 1    ' row 1 is header
    Dim r As Long
    Dim c As Long
    Dim lossAmt As Double
    Dim parsedPeril As String
    Dim parsedEvent As String
    Dim parsedGroup As String

    For r = startRow To rMax
        For c = 2 To cMax
            If IsNumeric(wideMatrix(r, c)) Then
                lossAmt = CDbl(wideMatrix(r, c))
                If lossAmt > materialityThreshold Then

                    ' Derive segment-specific fields from column header band
                    p_ParseRowFields wideMatrix, c, Segment, _
                                     parsedPeril, parsedEvent, parsedGroup

                    count = count + 1
                    temp(count, LC_SEQ) = count - 1
                    temp(count, LC_COLUME) = c
                    temp(count, LC_YEAR) = wideMatrix(r, 1)
                    temp(count, LC_SUB_LOB) = wideMatrix(HDR_ROW_SUBLOB, c)
                    temp(count, LC_PERIL) = parsedPeril
                    temp(count, LC_EVENT) = parsedEvent
                    temp(count, LC_AMOUNT) = lossAmt
                    temp(count, LC_TID) = vbNullString       ' Python fills via Assumption
                    temp(count, LC_GROUP) = parsedGroup

                End If
            End If
        Next c
    Next r

    ' --- Trim to actual row count ---
    If count < 2 Then
        errMsg = THIS_MODULE & ".TransformToLongList: no rows passed materiality filter" & _
                 " [segment=" & Segment & " threshold=" & materialityThreshold & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "TransformToLongList", errMsg
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To LONG_COL_COUNT)
    Dim i As Long
    Dim j As Long
    For i = 1 To count
        For j = 1 To LONG_COL_COUNT
            result(i, j) = temp(i, j)
        Next j
    Next i
    Erase temp

    outLongList = result
    TransformToLongList = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "TransformToLongList", _
        "Wide-to-Long complete; segment=[" & Segment & "]" & _
        " outputRows=" & count - 1 & _
        " sourceRows=" & rMax - startRow + 1 & _
        " sourceCols=" & cMax - 1 & _
        " threshold=" & materialityThreshold
End Function

' ==============================================================================================
' SECTION 02: SUBLOB AGGREGATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] AggregateToGN
'
' 功能说明      : 将 10K 宽表矩阵按 SubLoB 维度跨模拟年累加，生成供追加到 GN 工作表的汇总数组
'               : 矩阵约定与 TransformToLongList 相同：col 1 = Year，row 1 = SubLoB 标题
'               : startRow 需反映实际数据起始行（CAT_*/RDS 为 4；MM 为 2）
'               : 输出格式（含表头行）：SubLoB(1) | Peril(2) | Mean(3) | Year1..YearN
'               : Mean = 所有模拟年该 SubLoB 总损失 / yearMax（防除零内置）
'               : 输出由 biz_io.AppendArrayToSheet 追加到 GN 表，不覆盖既有数据
' 参数          : wideMatrix  - 二维 Variant 数组（来自 biz_io；含表头行）
'               : startRow    - 数据起始行（1-based）
'               : segment     - 分段标签（用于日志标识）
'               : yearMax     - 模拟年数上限（均值计算分母；通常 10000）
'               : outSummary  - 输出：二维 Variant 数组（含表头行）
'               : errMsg      - 输出：失败时的错误说明
' 返回          : Boolean - True=聚合成功，outSummary 可用；False=失败，errMsg 已填充
' Purpose       : Aggregate 10K wide matrix by SubLoB for GN sheet append
' Contract      : Business / Domain (query-only; no IO, no Application state mutation)
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function AggregateToGN(ByRef wideMatrix As Variant, _
                                ByVal startRow As Long, _
                                ByVal Segment As String, _
                                ByVal yearMax As Long, _
                                ByRef outSummary As Variant, _
                                ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    AggregateToGN = False
    outSummary = core_utils.EMPTY_VALUE

    Dim matErr As String
    If Not ValidateTenKMatrix(wideMatrix, startRow, matErr) Then
        errMsg = THIS_MODULE & ".AggregateToGN: " & matErr
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "AggregateToGN", errMsg
        Exit Function
    End If

    If yearMax <= 0 Then
        errMsg = THIS_MODULE & ".AggregateToGN: yearMax must be > 0; got " & yearMax
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "AggregateToGN", errMsg
        Exit Function
    End If

    Dim rMax As Long: rMax = UBound(wideMatrix, 1)
    Dim cMax As Long: cMax = UBound(wideMatrix, 2)

    ' --- Step 1: Extract unique SubLoB names from header row 1 ---
    Dim dictSubRow As Object
    Set dictSubRow = CreateObject("Scripting.Dictionary")
    dictSubRow.CompareMode = vbTextCompare

    Dim c As Long
    Dim subLobStr As String
    For c = 2 To cMax
        subLobStr = Trim$(CStr(wideMatrix(HDR_ROW_SUBLOB, c)))
        If Len(subLobStr) > 0 And Not dictSubRow.Exists(subLobStr) Then
            dictSubRow(subLobStr) = dictSubRow.count + 1
        End If
    Next c

    Dim numSub As Long: numSub = dictSubRow.count
    If numSub = 0 Then
        errMsg = THIS_MODULE & ".AggregateToGN: no SubLoB columns found in matrix header" & _
                 " [segment=" & Segment & "]"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "AggregateToGN", errMsg
        Set dictSubRow = Nothing
        Exit Function
    End If

    ' --- Step 2: Allocate output (1 header row + numSub data rows) ---
    Dim totalCols As Long: totalCols = GN_HEADER_COLS + yearMax
    Dim result() As Variant
    ReDim result(1 To numSub + 1, 1 To totalCols)

    result(1, 1) = "SubLoB"
    result(1, 2) = "Peril"
    result(1, 3) = "Mean"
    Dim yr As Long
    For yr = 1 To yearMax
        result(1, GN_HEADER_COLS + yr) = yr
    Next yr

    Dim subKey As Variant
    Dim subRowIdx As Long
    For Each subKey In dictSubRow.Keys
        subRowIdx = CLng(dictSubRow(subKey)) + 1    ' +1: row 1 is header
        result(subRowIdx, 1) = subKey
        result(subRowIdx, 2) = vbNullString
        result(subRowIdx, 3) = 0#
        For yr = 1 To yearMax
            result(subRowIdx, GN_HEADER_COLS + yr) = 0#
        Next yr
    Next subKey

    ' --- Step 3: Accumulate losses up to yearMax boundary ---
    Dim limitRow As Long
    limitRow = startRow + yearMax - 1
    If limitRow > rMax Then limitRow = rMax

    Dim r As Long
    Dim yrIdx As Long
    Dim lossAmt As Double
    Dim targetRow As Long

    For c = 2 To cMax
        subLobStr = Trim$(CStr(wideMatrix(HDR_ROW_SUBLOB, c)))
        If dictSubRow.Exists(subLobStr) Then
            targetRow = CLng(dictSubRow(subLobStr)) + 1
            For r = startRow To limitRow
                If IsNumeric(wideMatrix(r, c)) Then
                    lossAmt = CDbl(wideMatrix(r, c))
                    yrIdx = r - startRow + 1
                    If yrIdx >= 1 And yrIdx <= yearMax Then
                        result(targetRow, GN_HEADER_COLS + yrIdx) = _
                            CDbl(result(targetRow, GN_HEADER_COLS + yrIdx)) + lossAmt
                        result(targetRow, 3) = CDbl(result(targetRow, 3)) + lossAmt
                    End If
                End If
            Next r
        End If
    Next c

    ' --- Step 4: Compute Mean = cumulative total / yearMax ---
    Dim i As Long
    For i = 2 To numSub + 1
        result(i, 3) = CDbl(result(i, 3)) / CDbl(yearMax)
    Next i

    Set dictSubRow = Nothing
    outSummary = result
    AggregateToGN = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "AggregateToGN", _
        "GN aggregation complete; segment=[" & Segment & "]" & _
        " subLoBCount=" & numSub & _
        " yearMax=" & yearMax & _
        " sourceRows=" & limitRow - startRow + 1
End Function

' ==============================================================================================
' SECTION 03: INPUT VALIDATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ValidateTenKMatrix
'
' 功能说明      : 验证 10K 源矩阵满足 TransformToLongList 和 AggregateToGN 的结构前置条件
'               : 检查：非空二维数组、至少 2 行（表头 + 数据）、至少 2 列（Year + SubLoB）
'               : startRow 合法性（>= 1，<= UBound 行）
'               : 注意：startRow 本身由调用方从 Named Range 读取，它已反映表头行数
'               : 本函数不感知分段类型，不检查 startRow >= 4 这一业务约定
' 参数          : wideMatrix - 待验证的二维 Variant 数组
'               : startRow   - 数据起始行号（1-based）
'               : errMsg     - 输出：首个发现的验证错误说明
' 返回          : Boolean - True=验证通过；False=失败，errMsg 已填充
' Purpose       : Structural pre-condition guard for 10K wide matrix
' Contract      : Business / Domain (query-only)
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function ValidateTenKMatrix(ByRef wideMatrix As Variant, _
                                    ByVal startRow As Long, _
                                    ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ValidateTenKMatrix = False

    If core_utils.IsEmptyValue(wideMatrix) Then
        errMsg = "wideMatrix is EMPTY_VALUE"
        Exit Function
    End If

    If Not IsArray(wideMatrix) Then
        errMsg = "wideMatrix is not an array"
        Exit Function
    End If

    Dim nRows As Long
    Dim nCols As Long
    On Error Resume Next
    nRows = UBound(wideMatrix, 1) - LBound(wideMatrix, 1) + 1
    nCols = UBound(wideMatrix, 2) - LBound(wideMatrix, 2) + 1
    Dim e As Long: e = Err.Number
    On Error GoTo 0

    If e <> 0 Then
        errMsg = "wideMatrix is not a valid 2D array"
        Exit Function
    End If

    If nCols < 2 Then
        errMsg = "wideMatrix must have >= 2 cols (Year + 1 SubLoB); got " & nCols
        Exit Function
    End If

    If nRows < 2 Then
        errMsg = "wideMatrix must have >= 2 rows (header + 1 data); got " & nRows
        Exit Function
    End If

    If startRow < 1 Then
        errMsg = "startRow must be >= 1; got " & startRow
        Exit Function
    End If

    Dim rMax As Long: rMax = UBound(wideMatrix, 1)
    If startRow > rMax Then
        errMsg = "startRow (" & startRow & ") exceeds matrix row bound (" & rMax & ")"
        Exit Function
    End If

    ValidateTenKMatrix = True
End Function

' ==============================================================================================
' SECTION 04: ASSUMPTION MAPPING
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] BuildAssumptionArray
'
' 功能说明      : 将 Treaty 数组与 SubLoB 映射矩阵交叉关联，生成指定分保组的完整假设数组
'               : 对每个属于 groupName 的 TID，在 SubLoB 矩阵中定位对应列，
'               : 遍历所有标记为 1 的数据行，将 Treaty 财务字段与 SubLoB 元数据合并为一行
'               : 输出数组含表头行；列顺序与 Python SCHEMA["ASMP"] 严格对齐
'               : LimitForm 由 ResolveLimitForm 派发；UniqueKey = Group|TID|Segment|Peril
'               : IO（读 Treaty/SubLoB 表，写 GroupName_Asmp 表）属于 biz_io
'               : 分保组列表遍历与批量执行属于 biz_entry
' 参数          : treatyArray  - Treaty 表的二维 Variant 数组（无表头行；数据从 row 1 起）
'                                来自 biz_io.ReadNamedRangeArray("rng_Treaty")
'               : subLoBArray  - SubLoB 表的二维 Variant 数组（含表头行 row 1 = TID 名称）
'                                来自 biz_io.ReadNamedRangeArray("rng_SubLoB")
'               : groupName    - 分保组名称（如 "QS" / "Cat" / "MM"）
'               : outArray     - 输出：完整假设二维 Variant 数组（含表头行）
'               : errMsg       - 输出：失败时的错误说明
' 返回          : Boolean - True=映射成功，outArray 可用；False=失败，errMsg 已填充
' Purpose       : Treaty x SubLoB TID cross-reference engine; produces typed assumption array
' Contract      : Business / Domain (query-only; no IO, no Application state mutation)
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function BuildAssumptionArray(ByRef treatyArray As Variant, _
                                      ByRef subLoBArray As Variant, _
                                      ByVal GroupName As String, _
                                      ByRef outArray As Variant, _
                                      ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    BuildAssumptionArray = False
    outArray = core_utils.EMPTY_VALUE

    ' --- Input guards ---
    If core_utils.IsEmptyValue(treatyArray) Or Not IsArray(treatyArray) Then
        errMsg = THIS_MODULE & ".BuildAssumptionArray: treatyArray is not a valid array"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", errMsg
        Exit Function
    End If

    If core_utils.IsEmptyValue(subLoBArray) Or Not IsArray(subLoBArray) Then
        errMsg = THIS_MODULE & ".BuildAssumptionArray: subLoBArray is not a valid array"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", errMsg
        Exit Function
    End If

    If Len(Trim$(GroupName)) = 0 Then
        errMsg = THIS_MODULE & ".BuildAssumptionArray: groupName is empty"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", errMsg
        Exit Function
    End If

    ' --- Step 1: Build TID → SubLoB column index dictionary ---
    ' SubLoB header row (row 1): TID names occupy columns SC_TID_START onward.
    Dim dictTID As Object
    Set dictTID = CreateObject("Scripting.Dictionary")
    dictTID.CompareMode = vbTextCompare

    Dim k As Long
    For k = SC_TID_START To UBound(subLoBArray, 2)
        If Not IsEmpty(subLoBArray(1, k)) Then
            dictTID(Trim$(CStr(subLoBArray(1, k)))) = k
        End If
    Next k

    If dictTID.count = 0 Then
        errMsg = THIS_MODULE & ".BuildAssumptionArray: no TID headers found in subLoBArray" & _
                 " (expected from column " & SC_TID_START & " onward)"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", errMsg
        Set dictTID = Nothing
        Exit Function
    End If

    ' --- Step 2: Worst-case pre-allocation ---
    Dim treatyRows As Long: treatyRows = UBound(treatyArray, 1)
    Dim subLoBRows As Long: subLoBRows = UBound(subLoBArray, 1)

    Dim temp() As Variant
    ReDim temp(1 To treatyRows * subLoBRows + 1, 1 To ASMP_COL_COUNT)

    ' --- Step 3: Write header row ---
    temp(1, AC_GROUP) = "Treaty_Group"
    temp(1, AC_TID) = "TID"
    temp(1, AC_TNAME) = "Treaty_Name"
    temp(1, AC_SEGMENT) = "Segment"
    temp(1, AC_MAJOR_LOB) = "Major_LoB"
    temp(1, AC_SUB_LOB) = "Sub_LoB"
    temp(1, AC_PERIL) = "Peril"
    temp(1, AC_CCY) = "Ccy"
    temp(1, AC_LIMIT_RISK) = "Limit_Risk"
    temp(1, AC_LIMIT_EVENT) = "Limit_Event"
    temp(1, AC_RETENTION) = "Retention"
    temp(1, AC_SHARE) = "Share"
    temp(1, AC_INURING) = "Inuring"
    temp(1, AC_LIMIT_FORM) = "Limit_Form"
    temp(1, AC_UNIQUE_KEY) = "Unique_Key"

    ' --- Step 4: Cross-reference loop ---
    Dim outRow As Long: outRow = 1      ' row 1 is header
    Dim i As Long
    Dim j As Long
    Dim tid As String
    Dim targetCol As Long
    Dim segStr As String
    Dim subStr As String

    For i = 1 To treatyRows
        If Trim$(CStr(treatyArray(i, TC_GROUP))) = GroupName Then
            tid = Trim$(CStr(treatyArray(i, TC_TID)))

            If dictTID.Exists(tid) Then
                targetCol = CLng(dictTID(tid))

                ' SubLoB data rows start at row 2 (row 1 = TID header)
                For j = 2 To subLoBRows
                    If subLoBArray(j, targetCol) = 1 Then
                        outRow = outRow + 1

                        segStr = Trim$(CStr(subLoBArray(j, SC_SEGMENT)))
                        subStr = Trim$(CStr(subLoBArray(j, SC_SUB)))

                        temp(outRow, AC_GROUP) = GroupName
                        temp(outRow, AC_TID) = tid
                        temp(outRow, AC_TNAME) = treatyArray(i, TC_NAME)
                        temp(outRow, AC_SEGMENT) = segStr
                        temp(outRow, AC_MAJOR_LOB) = subLoBArray(j, SC_MAJOR)
                        temp(outRow, AC_SUB_LOB) = subStr
                        temp(outRow, AC_PERIL) = subLoBArray(j, SC_PERIL)
                        temp(outRow, AC_CCY) = treatyArray(i, TC_CCY)
                        temp(outRow, AC_LIMIT_RISK) = treatyArray(i, TC_LIMIT_RISK)
                        temp(outRow, AC_LIMIT_EVENT) = treatyArray(i, TC_LIMIT_EVENT)
                        temp(outRow, AC_RETENTION) = treatyArray(i, TC_RETENTION)
                        temp(outRow, AC_SHARE) = treatyArray(i, TC_SHARE)
                        temp(outRow, AC_INURING) = treatyArray(i, TC_INURING)
                        temp(outRow, AC_LIMIT_FORM) = ResolveLimitForm(segStr, subStr)
                        temp(outRow, AC_UNIQUE_KEY) = GroupName & "|" & tid & "|" & _
                                                       segStr & "|" & _
                                                       CStr(subLoBArray(j, SC_PERIL))
                    End If
                Next j
            End If
        End If
    Next i

    If outRow < 2 Then
        errMsg = THIS_MODULE & ".BuildAssumptionArray: no rows produced for group=[" & _
                 GroupName & "]; check Treaty group column and SubLoB mapping flags"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", errMsg
        Set dictTID = Nothing
        Exit Function
    End If

    ' --- Step 5: Trim to actual size ---
    Dim result() As Variant
    ReDim result(1 To outRow, 1 To ASMP_COL_COUNT)
    Dim p As Long
    Dim q As Long
    For p = 1 To outRow
        For q = 1 To ASMP_COL_COUNT
            result(p, q) = temp(p, q)
        Next q
    Next p
    Erase temp

    Set dictTID = Nothing
    outArray = result
    BuildAssumptionArray = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "BuildAssumptionArray", _
        "Assumption mapping complete; group=[" & GroupName & "]" & _
        " outputRows=" & outRow - 1 & _
        " treatyRows=" & treatyRows & _
        " subLoBRows=" & subLoBRows - 1
End Function

' ----------------------------------------------------------------------------------------------
' [F] ResolveLimitForm
'
' 功能说明      : 根据业务分段和子业务线判定再保险限额适用形式
'               : 返回规范字符串常量（Risk / Event / Not Applicable / Unknown）
'               : 本函数为 Domain 层权威实现，取代 app_05_asmp.GetLimitForm（v1 私有函数）
'               : 永不失败：未匹配分段返回 "Unknown" 而非报错
' 参数          : segment - 损失分段名称（如 CAT_Large / MM_Large / *_Att / MM_RDS）
'               : subLoB  - 子业务线名称（仅用于 MM_RDS 分支中的 Event/Risk 判断）
' 返回          : String - 限额适用形式（Risk / Event / Not Applicable / Unknown）
' Purpose       : Canonical limit form resolver; single source of truth for form dispatch
' Contract      : Business / Domain (query-only; pure function)
' Side Effects  : None (query-only)
' ----------------------------------------------------------------------------------------------
Public Function ResolveLimitForm(ByVal Segment As String, _
                                  ByVal subLoB As String) As String
    Segment = Trim$(Segment)
    subLoB = Trim$(subLoB)

    Select Case True
        Case Segment Like "*_Att"
            ResolveLimitForm = LF_NA

        Case Segment = "CAT_Large", Segment = "CAT_NM"
            ResolveLimitForm = LF_EVENT

        Case Segment = "MM_Large"
            ResolveLimitForm = LF_RISK

        Case Segment = "MM_RDS"
            If InStr(1, subLoB, "Event", vbTextCompare) > 0 Then
                ResolveLimitForm = LF_EVENT
            Else
                ResolveLimitForm = LF_RISK
            End If

        Case Else
            ResolveLimitForm = LF_UNKNOWN
    End Select
End Function

' ==============================================================================================
' SECTION 05: PRIVATE HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] p_ParseRowFields
'
' 功能说明      : 根据分段类型从宽矩阵列标题带派生单列的 Peril / Event / Group 值
'               : 封装 CRDSParser / CCATParser / CMMParser 三类解析器的完整派发规则：
'
'               : CAT_*（CCATParser）
'                   Peril = wideMatrix(2, c)     矩阵第 2 行
'                   Event = wideMatrix(3, c)     矩阵第 3 行
'                   Group = "CAT"               固定值
'
'               : RDS（CRDSParser）
'                   Peril = wideMatrix(2, c)     矩阵第 2 行
'                   Event = wideMatrix(3, c)     矩阵第 3 行
'                   Group 由 SubLoB（row 1）关键词决定：
'                     SubLoB 含 "Aviation"              → "Aviation"
'                     SubLoB 含 "Marine"/"PVT"/"Offshore" → "Marine"
'                     否则                                → 空字符串
'
'               : NonCat_Large*（CMMParser）
'                   Peril = "MM_Large"          由 segment 名称派生
'                   Event = 空字符串
'                   Group = 空字符串
'
'               : NonCat_Attritional（CMMParser）
'                   Peril = "MM_Att"            由 segment 名称派生
'                   Event = 空字符串
'                   Group = 空字符串
'
'               : 未匹配分段（Fallback）
'                   Peril = 空字符串
'                   Event = 空字符串
'                   Group = 空字符串
'
' 参数          : wideMatrix  - 宽矩阵（TransformToLongList 的调用方已验证其维度）
'               : c           - 当前列索引（2-based，列 1 = Year 列）
'               : segment     - 分段名称，驱动分支派发
'               : outPeril    - 输出：Peril 字符串
'               : outEvent    - 输出：Event 字符串
'               : outGroup    - 输出：Group 字符串
' Purpose       : Inline parser dispatch; replaces class-module parser hierarchy
' Contract      : Inherited from TransformToLongList (query-only; no IO; no side effects)
' ----------------------------------------------------------------------------------------------
Private Sub p_ParseRowFields(ByRef wideMatrix As Variant, _
                              ByVal c As Long, _
                              ByVal Segment As String, _
                              ByRef outPeril As String, _
                              ByRef outEvent As String, _
                              ByRef outGroup As String)
    ' Initialise outputs to empty — fallback state for unrecognised segment
    outPeril = vbNullString
    outEvent = vbNullString
    outGroup = vbNullString

    Dim subLobVal As String

    Select Case True

        ' ---------------------------------------------------------------
        ' CAT_* branch  (mirrors CCATParser)
        ' Peril and Event are read from fixed header band rows 2 and 3.
        ' Group is always "CAT".
        ' ---------------------------------------------------------------
        Case Segment Like "CAT_*"
            outPeril = Trim$(CStr(wideMatrix(HDR_ROW_PERIL, c)))
            outEvent = Trim$(CStr(wideMatrix(HDR_ROW_EVENT, c)))
            outGroup = PARSER_GROUP_CAT

        ' ---------------------------------------------------------------
        ' RDS branch  (mirrors CRDSParser)
        ' Peril and Event from header band rows 2 and 3.
        ' Group derived from SubLoB keyword matching.
        ' Evaluation order: Aviation first, then Marine keywords.
        ' ---------------------------------------------------------------
        Case Segment = "RDS"
            outPeril = Trim$(CStr(wideMatrix(HDR_ROW_PERIL, c)))
            outEvent = Trim$(CStr(wideMatrix(HDR_ROW_EVENT, c)))
            subLobVal = Trim$(CStr(wideMatrix(HDR_ROW_SUBLOB, c)))

            If InStr(1, subLobVal, PARSER_GROUP_AVIATION, vbTextCompare) > 0 Then
                outGroup = PARSER_GROUP_AVIATION
            ElseIf InStr(1, subLobVal, PARSER_GROUP_MARINE, vbTextCompare) > 0 Or _
                   InStr(1, subLobVal, PARSER_KW_PVT, vbTextCompare) > 0 Or _
                   InStr(1, subLobVal, PARSER_KW_OFFSHORE, vbTextCompare) > 0 Then
                outGroup = PARSER_GROUP_MARINE
            End If

        ' ---------------------------------------------------------------
        ' NonCat_Large* branch  (mirrors CMMParser NonCat_Large* arm)
        ' Peril derived from segment name; Event and Group are empty.
        ' ---------------------------------------------------------------
        Case Segment Like "NonCat_Large*"
            outPeril = PARSER_PERIL_MM_LARGE

        ' ---------------------------------------------------------------
        ' NonCat_Attritional branch  (mirrors CMMParser Attritional arm)
        ' Peril derived from segment name; Event and Group are empty.
        ' ---------------------------------------------------------------
        Case Segment = "NonCat_Attritional"
            outPeril = PARSER_PERIL_MM_ATT

        ' ---------------------------------------------------------------
        ' Fallback: unrecognised segment — all fields remain empty.
        ' Caller (TransformToLongList) logs at Info level; no error raised.
        ' ---------------------------------------------------------------
        Case Else
            ' outPeril / outEvent / outGroup already vbNullString

    End Select
End Sub


