Attribute VB_Name = "biz_entry"
' ==============================================================================================
' MODULE NAME       : biz_entry
' LAYER             : Business
' PURPOSE           : Business Entry layer. Owns all orchestration logic for the reinsurance
'                     analysis platform. Three responsibility areas:
'                     (1) Single-segment loss transform pipeline (10K → LOSS sheet),
'                     (2) Single-group assumption mapping pipeline (Treaty × SubLoB → Asmp sheet),
'                     (3) SubLoB aggregation pipeline (10K → GN sheet append).
'                     Batch orchestration (RunAllSegments / RunAllAssumptions) loops over the
'                     above single-unit entry points.
'                     No calculation logic. No worksheet IO. No Application state mutation.
' DEPENDS           : biz_io       v2.0.0 (ReadNamedRangeValue, ReadNamedRangeArray,
'                                          ReadSheetArea, WriteArrayToSheet, AppendArrayToSheet,
'                                          ClearOutputSheet, RequireOutputSheet)
'                     biz_dom      v2.0.0 (TransformToLongList, AggregateToGN,
'                                          ValidateTenKMatrix, BuildAssumptionArray)
'                     plat_runtime v2.1.2 (SafeExecute, ResultStore, ResultOk, ResultFail,
'                                          LogInfo, LogWarn, LogError, MeasurePerf)
'                     plat_context v2.1.2 (GetConfigValue)
'       : plat_runtime v2.2.0 (SessionTeardown, RefreshOutputSheetList)
' NOTE              : - This module is the sole orchestration boundary of the Business layer.
'                       All entry points are designed to be invoked via plat_runtime.SafeExecute.
'                     - Entry points carry the signature:
'                         Sub BizEntry(ByVal actionName As String, ByVal RunId As String)
'                       and MUST call plat_runtime.ResultStore before returning.
'                     - No direct worksheet access is permitted; all IO is delegated to biz_io.
'                     - No calculation logic is permitted; all transforms are delegated to biz_dom.
'                     - No Application state mutation (ScreenUpdating, Calculation, etc.) is
'                       permitted; that responsibility belongs exclusively to plat_runtime.
'                     - Python engine invocation (xlwings RunPython) is NOT owned here.
'                       RunPythonEngine is a VBA-side orchestration stub that prepares data and
'                       signals readiness; the actual dispatch is Platform-owned IO. The stub
'                       is retained as the architectural coordination point for the Python
'                       engine boundary. Full Python dispatch migration is tracked as
'                       TODO(v2.1): migrate xlwings dispatch to plat_runtime IO layer.
'                     - CommitProcess (worksheet archive/delete) is Platform-level Workbook
'                       management, not Business orchestration. biz_entry does not call it.
'                       Migration is tracked as TODO(v2.1): CommitProcess → plat_runtime.
'                     - Batch entry points (RunAllSegments, RunAllAssumptions) loop over
'                       single-unit entries; they do not duplicate orchestration logic.
'                     - All public entry point Subs carry (actionName, RunId) as parameters
'                       for TResult contract compliance via SafeExecute.
'                     - Named range constants used by this module:
'                         CFG_TENK_FILE_PATH   = "ref_10K_FilePath"
'                         CFG_TENK_SEGMENT     = "ref_10K_Segment"
'                         CFG_TENK_LOSS_SHEET  = "ref_10K_LossList"
'                         CFG_TENK_START_ROW   = "ref_10K_StartRow"
'                         CFG_TENK_MAT_THRESH  = "p_materiality_threshold"
'                         CFG_SEGMENT_LIST     = "rng_segment_list"
'                         CFG_GROUP_LIST       = "rng_group_list"
'                     - Assumption group names are read from rng_group_list at runtime;
'                       no hardcoded group list in this module.
' STATUS            : Draft
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Orchestration logic distributed across app_04_tenk,
'                             app_05_asmp, and app_07_engine; no unified entry boundary.
'   - Init (Design): Entry points directly coupled to IO, domain logic, and Application
'                    state mutation; no TResult contract; no SafeExecute wrapper.
'   - Init (Scope): Python engine invocation (Run_Python_Engine / Run_All_Segments) and
'                   session teardown (CommitProcess) interleaved inline; silent mode
'                   controlled via global variable (g_COMMIT_SILENT) across modules.

' v2.0.0
'   - Init (Architecture): Introduced biz_entry as the Business orchestration boundary module
'                          under the three-layer model (Core / Platform / Business).
'   - Init (Boundary): Established strict separation: IO → biz_io, calculation → biz_dom,
'                      orchestration → biz_entry; no layer may cross into another's domain.
'   - Init (Contract): All public entry points implement TResult contract via ResultStore;
'                      no MsgBox, no Err.Raise, no silent fallback in entry layer.
'   - Init (Batch Orchestration): RunAllSegments and RunAllAssumptions loop over single-unit
'                                 entries; batch logic does not replicate orchestration steps.
'   - Init (Python Stub): RunPythonEngine retained as VBA-side coordination stub;
'                         xlwings dispatch is Platform-owned IO (not yet migrated).
'   - Init (CommitProcess Exclusion): CommitProcess excluded from biz_entry scope;
'                                     Platform-level Workbook management; migration pending.
'   - Init (Failure Semantics): Boolean + ByRef errMsg from biz_io / biz_dom; failures map
'                               to TResult codes (VALIDATION / IO / UNHANDLED); no exceptions.
'   - Refactor (Python Dispatch): RunPythonEngine upgraded from stub to full dispatch entry.
'   - Refactor (Batch Teardown): RunAllSegments calls plat_runtime.SessionTeardown post-batch.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 00: MODULE CONSTANTS
'
' SECTION 01: LOSS TRANSFORM ENTRY
'   [S] RunLossTransform        - Single segment: 10K wide matrix → LOSS sheet (TResult contract)
'
' SECTION 02: ASSUMPTION MAPPING ENTRY
'   [S] RunAssumptionMapping    - Single group: Treaty × SubLoB → Asmp sheet (TResult contract)
'
' SECTION 03: SUBLOB AGGREGATION ENTRY
'   [S] RunSubLoBAggregate      - Single segment: 10K wide matrix → GN sheet append (TResult contract)
'
' SECTION 04: PYTHON ENGINE ENTRY
'   [S] RunPythonEngine         - VBA-side orchestration stub for Python engine dispatch
'
' SECTION 05: BATCH ORCHESTRATION
'   [S] RunAllSegments          - Batch: loop segment list → RunLossTransform per segment
'   [S] RunAllAssumptions       - Batch: loop group list  → RunAssumptionMapping per group
'
' SECTION 06: PRIVATE HELPERS
'   [f] p_ReadSegmentConfig     - Read and validate all named range config for a segment run
'   [f] p_ReadGroupConfig       - Read and validate all named range config for a group run
'   [f] p_ReadColumnList        - Read a named-range column vector into a 1D String array
'
' ==============================================================================================
' NOTE: [C]=Constant, [S]=Public Sub, [s]=Private Sub, [F]=Public Function, [f]=Private Function
'       Rule: Private helpers inherit the Contract and Side Effects of their calling public
'             entry point unless explicitly stated otherwise.
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 00: MODULE CONSTANTS
' ==============================================================================================

Private Const BIZ_LAYER     As String = "BIZ"
Private Const THIS_MODULE   As String = "biz_entry"

' Named range keys for segment-level configuration (read via biz_io.ReadNamedRangeValue)
Private Const NR_TENK_FILE_PATH  As String = "ref_10K_FilePath"
Private Const NR_TENK_SEGMENT    As String = "ref_10K_Segment"
Private Const NR_TENK_LOSS_SHEET As String = "ref_10K_LossList"
Private Const NR_TENK_START_ROW  As String = "ref_10K_StartRow"
Private Const NR_MAT_THRESHOLD   As String = "p_materiality_threshold"

' Named range keys for assumption-level configuration
Private Const NR_ASMP_GROUP      As String = "ref_Asmp_Group"

' Named range keys for batch orchestration
Private Const NR_SEGMENT_LIST    As String = "rng_segment_list"
Private Const NR_GROUP_LIST      As String = "rng_group_list"

' Named range keys for input data arrays
Private Const NR_TREATY_ARRAY    As String = "rng_Treaty"
Private Const NR_SUBLOB_ARRAY    As String = "rng_SubLoB"

' GN sheet name (output for SubLoB aggregation)
Private Const SH_GN              As String = "GN"

' Assumption output sheet suffix (GroupName & suffix = sheet name)
Private Const ASMP_SHEET_SUFFIX  As String = "_Asmp"

' Default materiality threshold when named range is absent or invalid
Private Const DEFAULT_MAT_THRESH As Double = 0.5

Private Const NR_PY_SCRIPT_PATH  As String = "ref_Py_ScriptPath"

' ==============================================================================================
' SECTION 01: LOSS TRANSFORM ENTRY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RunLossTransform
'
' 功能说明      : 单分段 10K 宽表转长表完整流程编排入口
'               : 流程：读取配置 → 读取 10K 矩阵 → 验证矩阵 → 宽转长 → 清空目标表 → 写入结果
'               : 所有 IO 通过 biz_io；所有转换通过 biz_dom；无直接 worksheet 访问
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入，传递给 TResult）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入，传递给 TResult）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Orchestrate single-segment 10K wide-to-long pipeline under TResult contract
' Contract      : Business / Entry (orchestration only; no IO, no calculation, no App mutation)
' Side Effects  : Calls biz_io (read + write); calls plat_runtime.ResultStore
' ----------------------------------------------------------------------------------------------
Public Sub RunLossTransform(ByVal actionName As String, ByVal RunId As String)

    Dim r       As tResult
    Dim errMsg  As String

    ' --- Step 1: Read and validate segment configuration ---
    Dim cfg     As TSegmentConfig
    If Not p_ReadSegmentConfig(cfg, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "Config read failed; action=" & actionName & " run_id=" & RunId & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "配置读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
        "Config loaded; segment=" & cfg.Segment & " startRow=" & cfg.startRow & _
        " threshold=" & cfg.MatThreshold & " outSheet=" & cfg.LossSheetName & _
        " action=" & actionName & " run_id=" & RunId

    ' --- Step 2: Assert output sheet exists (Fail Fast) ---
    If Not biz_io.RequireOutputSheet(cfg.LossSheetName, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "Output sheet missing; sheet=" & cfg.LossSheetName & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "输出工作表不存在：" & cfg.LossSheetName, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 3: Read 10K wide matrix from named range ---
    Dim wideMatrix  As Variant
    If Not biz_io.ReadSheetArea(cfg.TenKFilePath, cfg.Segment, wideMatrix, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "10K matrix read failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "10K 矩阵读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 4: Validate matrix structure before transform ---
    If Not biz_dom.ValidateTenKMatrix(wideMatrix, cfg.Segment, cfg.startRow, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "Matrix validation failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "10K 矩阵结构验证失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 5: Transform wide → long ---
    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunLossTransform", "transform_start", _
        "segment=" & cfg.Segment

    Dim longArray   As Variant
    If Not biz_dom.TransformToLongList(wideMatrix, cfg.Segment, cfg.startRow, _
                                        cfg.MatThreshold, longArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "TransformToLongList failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
                "宽转长失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunLossTransform", "transform_end", _
        "segment=" & cfg.Segment

    ' --- Step 6: Clear output sheet (preserve header row 1) ---
    If Not biz_io.ClearOutputSheet(cfg.LossSheetName, errMsg) Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "ClearOutputSheet failed (non-fatal); sheet=" & cfg.LossSheetName & " err=" & errMsg
        ' Non-fatal: continue; write will overwrite existing data
    End If

    ' --- Step 7: Write long array to output sheet ---
    If Not biz_io.WriteArrayToSheet(cfg.LossSheetName, longArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
            "WriteArrayToSheet failed; sheet=" & cfg.LossSheetName & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "写入 LOSS 工作表失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Done ---
    Dim rowCount As Long
    rowCount = UBound(longArray, 1) - LBound(longArray, 1)   ' exclude header row
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunLossTransform", _
        "Pipeline complete; segment=" & cfg.Segment & _
        " dataRows=" & rowCount & _
        " outSheet=" & cfg.LossSheetName & _
        " action=" & actionName & " run_id=" & RunId

    r = plat_runtime.ResultOk(actionName, RunId, CLng(rowCount))
    plat_runtime.ResultStore r

End Sub

' ==============================================================================================
' SECTION 02: ASSUMPTION MAPPING ENTRY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RunAssumptionMapping
'
' 功能说明      : 单分保组 Treaty × SubLoB 假设映射完整流程编排入口
'               : 流程：读取组名 → 读取 Treaty/SubLoB 数组 → 映射 → 清空 Asmp 表 → 写入
'               : 所有 IO 通过 biz_io；所有映射逻辑通过 biz_dom；无直接 worksheet 访问
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Orchestrate single-group Treaty × SubLoB assumption mapping pipeline
' Contract      : Business / Entry (orchestration only; no IO, no calculation, no App mutation)
' Side Effects  : Calls biz_io (read + write); calls plat_runtime.ResultStore
' ----------------------------------------------------------------------------------------------
Public Sub RunAssumptionMapping(ByVal actionName As String, ByVal RunId As String)

    Dim r       As tResult
    Dim errMsg  As String

    ' --- Step 1: Read group configuration ---
    Dim cfg     As TGroupConfig
    If Not p_ReadGroupConfig(cfg, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "Group config read failed; action=" & actionName & " run_id=" & RunId & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "分保组配置读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
        "Group config loaded; groupName=" & cfg.GroupName & _
        " outSheet=" & cfg.AsmpSheetName & _
        " action=" & actionName & " run_id=" & RunId

    ' --- Step 2: Assert output Asmp sheet exists (Fail Fast) ---
    If Not biz_io.RequireOutputSheet(cfg.AsmpSheetName, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "Asmp sheet missing; sheet=" & cfg.AsmpSheetName & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "假设输出工作表不存在：" & cfg.AsmpSheetName, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 3: Read Treaty array ---
    Dim treatyArray As Variant
    If Not biz_io.ReadNamedRangeArray(NR_TREATY_ARRAY, treatyArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "Treaty array read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "Treaty 数据读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 4: Read SubLoB array ---
    Dim subLoBArray As Variant
    If Not biz_io.ReadNamedRangeArray(NR_SUBLOB_ARRAY, subLoBArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "SubLoB array read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "SubLoB 数据读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 5: Build assumption array via domain engine ---
    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", "mapping_start", _
        "group=" & cfg.GroupName

    Dim asmpArray   As Variant
    If Not biz_dom.BuildAssumptionArray(treatyArray, subLoBArray, cfg.GroupName, _
                                         asmpArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "BuildAssumptionArray failed; group=" & cfg.GroupName & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
                "假设映射失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", "mapping_end", _
        "group=" & cfg.GroupName

    ' --- Step 6: Clear Asmp output sheet (preserve header row 1) ---
    If Not biz_io.ClearOutputSheet(cfg.AsmpSheetName, errMsg) Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "ClearOutputSheet failed (non-fatal); sheet=" & cfg.AsmpSheetName & " err=" & errMsg
        ' Non-fatal: continue; overwrite will cover stale data
    End If

    ' --- Step 7: Write assumption array to Asmp sheet ---
    If Not biz_io.WriteArrayToSheet(cfg.AsmpSheetName, asmpArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
            "WriteArrayToSheet failed; sheet=" & cfg.AsmpSheetName & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "写入假设工作表失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Done ---
    Dim rowCount As Long
    rowCount = UBound(asmpArray, 1) - LBound(asmpArray, 1)   ' exclude header row
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAssumptionMapping", _
        "Pipeline complete; group=" & cfg.GroupName & _
        " dataRows=" & rowCount & _
        " outSheet=" & cfg.AsmpSheetName & _
        " action=" & actionName & " run_id=" & RunId

    r = plat_runtime.ResultOk(actionName, RunId, CLng(rowCount))
    plat_runtime.ResultStore r

End Sub

' ==============================================================================================
' SECTION 03: SUBLOB AGGREGATION ENTRY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RunSubLoBAggregate
'
' 功能说明      : 单分段 10K 宽表按 SubLoB 聚合并追加至 GN 表完整流程编排入口
'               : 流程：读取配置 → 读取 10K 矩阵 → 验证 → 聚合 → 追加写入 GN 表
'               : 追加模式（AppendArrayToSheet）：不清空 GN 表，多分段数据累积追加
'               : 所有 IO 通过 biz_io；所有聚合逻辑通过 biz_dom；无直接 worksheet 访问
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Orchestrate single-segment SubLoB aggregation pipeline (append to GN)
' Contract      : Business / Entry (orchestration only; no IO, no calculation, no App mutation)
' Side Effects  : Calls biz_io (read + append); calls plat_runtime.ResultStore
' ----------------------------------------------------------------------------------------------
Public Sub RunSubLoBAggregate(ByVal actionName As String, ByVal RunId As String)

    Dim r       As tResult
    Dim errMsg  As String

    ' --- Step 1: Read and validate segment configuration ---
    Dim cfg     As TSegmentConfig
    If Not p_ReadSegmentConfig(cfg, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "Config read failed; action=" & actionName & " run_id=" & RunId & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "配置读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
        "Config loaded; segment=" & cfg.Segment & _
        " action=" & actionName & " run_id=" & RunId

    ' --- Step 2: Assert GN output sheet exists (Fail Fast) ---
    If Not biz_io.RequireOutputSheet(SH_GN, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "GN sheet missing; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "GN 输出工作表不存在", errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 3: Read 10K wide matrix ---
    Dim wideMatrix  As Variant
    If Not biz_io.ReadSheetArea(cfg.TenKFilePath, cfg.Segment, wideMatrix, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "10K matrix read failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "10K 矩阵读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 4: Validate matrix structure ---
    If Not biz_dom.ValidateTenKMatrix(wideMatrix, cfg.Segment, cfg.startRow, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "Matrix validation failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "10K 矩阵结构验证失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Step 5: Aggregate wide matrix by SubLoB ---
    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", "aggregate_start", _
        "segment=" & cfg.Segment

    Dim gnArray     As Variant
    If Not biz_dom.AggregateToGN(wideMatrix, cfg.Segment, cfg.startRow, gnArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "AggregateToGN failed; segment=" & cfg.Segment & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
                "SubLoB 聚合失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    plat_runtime.MeasurePerf BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", "aggregate_end", _
        "segment=" & cfg.Segment

    ' --- Step 6: Append aggregated rows to GN sheet ---
    ' Append (not overwrite): multiple segments accumulate in GN across a batch run.
    ' Caller is responsible for clearing GN before a full batch if a fresh run is needed.
    If Not biz_io.AppendArrayToSheet(SH_GN, gnArray, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
            "AppendArrayToSheet failed; sheet=" & SH_GN & " err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "IO", _
                "追加写入 GN 工作表失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' --- Done ---
    Dim rowCount As Long
    rowCount = UBound(gnArray, 1) - LBound(gnArray, 1) + 1
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunSubLoBAggregate", _
        "Pipeline complete; segment=" & cfg.Segment & _
        " appendedRows=" & rowCount & _
        " outSheet=" & SH_GN & _
        " action=" & actionName & " run_id=" & RunId

    r = plat_runtime.ResultOk(actionName, RunId, CLng(rowCount))
    plat_runtime.ResultStore r

End Sub

' ==============================================================================================
' SECTION 04: PYTHON ENGINE ENTRY
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RunPythonEngine
'
' 功能说明      : VBA 侧 Python 引擎完整调度入口
'               : 流程：读取分段配置 → 验证脚本路径 → 验证 LOSS 表就绪（可选警告）→
'               :        通过 plat_runtime.RunOptimized 调用 Python 引擎 →
'               :        刷新输出表列表 → 返回 TResult
'               : Application 状态（ScreenUpdating / Calculation）由 plat_runtime 管理
'               : Session 归档/删除由 plat_runtime.SessionTeardown 管理（批量模式由
'               :   RunAllSegments 统一在批次结束后调用；单次模式由调用方决定是否调用）
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Full Python engine dispatch entry; delegates xlwings call to plat_runtime
' Contract      : Business / Entry (orchestration; no App mutation; no direct xlwings call)
' Side Effects  : Calls plat_runtime.RunOptimized (triggers xlwings RunPython via app_07_engine);
'               : calls plat_runtime.RefreshOutputSheetList;
'               : calls plat_runtime.ResultStore
' Note          : TODO(v2.1): migrate xlwings dispatch fully into plat_runtime IO layer;
'               : at that point, the RunOptimized call below becomes a plat_runtime IO call
'               : and the dependency on app_07_engine.Run_Python_Engine is eliminated.
' ----------------------------------------------------------------------------------------------
Public Sub RunPythonEngine(ByVal actionName As String, ByVal RunId As String)

    Dim r       As tResult
    Dim errMsg  As String

    ' -----------------------------------------------------------------------
    ' Step 1: Read and validate segment name
    ' -----------------------------------------------------------------------
    Dim segValue As Variant
    If Not biz_io.ReadNamedRangeValue(NR_TENK_SEGMENT, segValue, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
            "Segment config read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "当前分段配置读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    Dim Segment As String
    Segment = Trim$(CStr(segValue))
    If Len(Segment) = 0 Then
        errMsg = THIS_MODULE & ".RunPythonEngine: ref_10K_Segment is empty"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "分段名称为空，无法调用 Python 引擎", errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' Step 2: Read and validate Python script path
    ' -----------------------------------------------------------------------
    Dim pyPathValue As Variant
    If Not biz_io.ReadNamedRangeValue(NR_PY_SCRIPT_PATH, pyPathValue, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
            "Python script path read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "Python 脚本路径读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    Dim pyPath As String
    pyPath = Trim$(CStr(pyPathValue))
    If Len(pyPath) = 0 Then
        errMsg = THIS_MODULE & ".RunPythonEngine: ref_Py_ScriptPath is empty"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "Python 脚本路径为空，请先通过 SpecifyPyPath 设置路径", errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    If Dir(pyPath) = vbNullString Then
        errMsg = THIS_MODULE & ".RunPythonEngine: script file not found [" & pyPath & "]"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "Python 脚本文件不存在：" & pyPath, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' Step 3: Advisory readiness check — LOSS sheet existence (non-fatal)
    ' -----------------------------------------------------------------------
    Dim lossSheetValue As Variant
    If biz_io.ReadNamedRangeValue(NR_TENK_LOSS_SHEET, lossSheetValue, errMsg) Then
        Dim lossSheet As String
        lossSheet = Trim$(CStr(lossSheetValue))
        If Len(lossSheet) > 0 Then
            If biz_io.SheetIsEmpty(lossSheet) Then
                plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
                    "LOSS sheet appears empty before Python dispatch; segment=" & Segment & _
                    " sheet=" & lossSheet & "; verify RunLossTransform was called"
            End If
        End If
    End If

    ' -----------------------------------------------------------------------
    ' Step 4: Dispatch Python engine via plat_runtime.RunOptimized
    ' Application state (ScreenUpdating / Calculation) is managed by RunOptimized.
    ' xlwings RunPython is invoked inside app_07_engine.Run_Python_Engine.
    ' TODO(v2.1): replace this call with a plat_runtime IO layer call once
    '             xlwings dispatch is fully migrated out of app_07_engine.
    ' -----------------------------------------------------------------------
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
        "Dispatching Python engine; segment=" & Segment & _
        " script=" & pyPath & _
        " action=" & actionName & " run_id=" & RunId

    On Error GoTo EH_Dispatch
    plat_runtime.RunOptimized "app_07_engine.Run_Python_Engine"
    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' Step 5: Refresh output sheet list in Setup anchor range
    ' -----------------------------------------------------------------------
    Dim refreshErr As String
    If Not plat_runtime.RefreshOutputSheetList(refreshErr) Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
            "RefreshOutputSheetList failed (non-fatal); err=" & refreshErr
    End If

    ' -----------------------------------------------------------------------
    ' Done
    ' -----------------------------------------------------------------------
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunPythonEngine", _
        "Python engine dispatch complete; segment=" & Segment & _
        " action=" & actionName & " run_id=" & RunId

    r = plat_runtime.ResultOk(actionName, RunId, Segment)
    plat_runtime.ResultStore r
    Exit Sub

EH_Dispatch:
    errMsg = THIS_MODULE & ".RunPythonEngine: xlwings dispatch raised exception;" & _
             " segment=" & Segment & " err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunPythonEngine", errMsg
    r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
            "Python 引擎调用异常：" & Err.Description, errMsg)
    plat_runtime.ResultStore r

End Sub


' ==============================================================================================
' SECTION 05: BATCH ORCHESTRATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] RunAllSegments
'
' 功能说明      : 批量执行完整 Python 引擎流程：遍历 rng_segment_list →
'               :   逐段执行 RunLossTransform → RunPythonEngine → 批次结束后 SessionTeardown
'               : 每个分段以独立 SafeExecute 调用执行，失败时记录警告并继续（容错批量模式）
'               : Application 状态由 plat_runtime 管理；不依赖全局静默变量
'               : Session 归档/删除在所有分段执行完毕后统一由 plat_runtime.SessionTeardown 执行
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Batch orchestration: segment list → LossTransform + PythonEngine per segment,
'               : followed by unified SessionTeardown
' Contract      : Business / Entry (orchestration only; no IO, no calculation, no App mutation)
' Side Effects  : Writes ref_10K_Segment named range per iteration (config mutation);
'               : calls plat_runtime.SafeExecute per segment (two calls per segment);
'               : calls plat_runtime.SessionTeardown after batch loop;
'               : calls plat_runtime.ResultStore
' ----------------------------------------------------------------------------------------------
Public Sub RunAllSegments(ByVal actionName As String, ByVal RunId As String)

    Dim r       As tResult
    Dim errMsg  As String

    ' -----------------------------------------------------------------------
    ' Step 1: Read segment list
    ' -----------------------------------------------------------------------
    Dim segmentList() As String
    If Not p_ReadColumnList(NR_SEGMENT_LIST, segmentList, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
            "Segment list read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "分段列表读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    Dim totalCount As Long
    totalCount = UBound(segmentList) - LBound(segmentList) + 1

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
        "Batch start; totalSegments=" & totalCount & _
        " action=" & actionName & " run_id=" & RunId

    ' -----------------------------------------------------------------------
    ' Step 2: Loop — LossTransform + PythonEngine per segment
    ' -----------------------------------------------------------------------
    Dim failCount   As Long
    Dim failLog     As String
    Dim i           As Long
    failCount = 0
    failLog = vbNullString

    For i = LBound(segmentList) To UBound(segmentList)

        Dim segName As String
        segName = Trim$(segmentList(i))
        If Len(segName) = 0 Then GoTo NextSegment

        ' --- 2a: Inject segment name into config named range ---
        Dim WriteErr As String
        If Not biz_io.WriteNamedRangeValue(NR_TENK_SEGMENT, segName, WriteErr) Then
            plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
                "Config write failed; segment=" & segName & " err=" & WriteErr
            failCount = failCount + 1
            failLog = failLog & segName & "(config_write_fail);"
            GoTo NextSegment
        End If

        ' --- 2b: RunLossTransform for this segment ---
        Dim tResult As tResult
        tResult = plat_runtime.SafeExecute( _
                      actionName & "[" & segName & "].LossTransform", _
                      THIS_MODULE & ".RunLossTransform")

        If Not tResult.Ok Then
            plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
                "LossTransform failed (continuing); segment=" & segName & _
                " code=" & tResult.Code & " msg=" & tResult.UserMsg
            failCount = failCount + 1
            failLog = failLog & segName & ".LossTransform(" & tResult.Code & ");"
            GoTo NextSegment
        End If

        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
            "LossTransform complete; segment=" & segName & _
            " rows=" & CStr(tResult.value)

        ' --- 2c: RunPythonEngine for this segment ---
        Dim pyResult As tResult
        pyResult = plat_runtime.SafeExecute( _
                       actionName & "[" & segName & "].PythonEngine", _
                       THIS_MODULE & ".RunPythonEngine")

        If Not pyResult.Ok Then
            plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
                "PythonEngine failed (continuing); segment=" & segName & _
                " code=" & pyResult.Code & " msg=" & pyResult.UserMsg
            failCount = failCount + 1
            failLog = failLog & segName & ".PythonEngine(" & pyResult.Code & ");"
            GoTo NextSegment
        End If

        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
            "PythonEngine complete; segment=" & segName

NextSegment:
    Next i

    ' -----------------------------------------------------------------------
    ' Step 3: Session teardown — archive / delete output sheets
    ' Executed once after all segments, regardless of individual failures.
    ' plat_runtime.SessionTeardown owns all workbook structural mutations.
    ' -----------------------------------------------------------------------
    Dim teardownErr As String
    If Not plat_runtime.SessionTeardown(teardownErr) Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
            "SessionTeardown completed with errors (non-fatal for batch result);" & _
            " err=" & teardownErr
    Else
        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
            "SessionTeardown complete"
    End If

    ' -----------------------------------------------------------------------
    ' Step 4: Build batch summary result
    ' -----------------------------------------------------------------------
    Dim successCount As Long
    successCount = totalCount - failCount

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllSegments", _
        "Batch complete; total=" & totalCount & _
        " success=" & successCount & " failed=" & failCount & _
        " action=" & actionName & " run_id=" & RunId

    If failCount = 0 Then
        r = plat_runtime.ResultOk(actionName, RunId, CLng(successCount))
    Else
        ' Partial success: batch ran to completion but some segments failed.
        ' Caller can inspect DebugMsg for the failure log.
        r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
                "批量执行部分失败：成功 " & successCount & _
                " / 共 " & totalCount & " 个分段", failLog)
        r.Ok = (failCount = 0)   ' already False; explicit for readability
    End If

    plat_runtime.ResultStore r

End Sub

' ----------------------------------------------------------------------------------------------
' [S] RunAllAssumptions
'
' 功能说明      : 批量执行：遍历 rng_group_list → 逐组调用 RunAssumptionMapping
'               : 每次迭代前将当前组名写入 ref_Asmp_Group，随后通过 SafeExecute 触发入口
'               : 任一组失败时记录警告并继续，不中断批量流程（容错批量模式）
'               : 批量结束后返回汇总 TResult；失败组数记录在 DebugMsg
'               : 必须通过 plat_runtime.SafeExecute 调用；返回前调用 ResultStore
' 参数          : actionName - 业务动作标签（由 SafeExecute 注入）
'               : RunId      - 关联追踪 ID（由 SafeExecute 注入）
' 返回          : 无（通过 plat_runtime.ResultStore 返回 TResult）
' Purpose       : Batch orchestration over group list for AssumptionMapping pipeline
' Contract      : Business / Entry (orchestration only; no IO, no calculation, no App mutation)
' Side Effects  : Writes ref_Asmp_Group named range per iteration (config mutation);
'               : calls plat_runtime.ResultStore; calls plat_runtime.SafeExecute per group
' ----------------------------------------------------------------------------------------------
Public Sub RunAllAssumptions(ByVal actionName As String, ByVal RunId As String)

    Dim r           As tResult
    Dim errMsg      As String

    ' --- Step 1: Read group list ---
    Dim groupList() As String
    If Not p_ReadColumnList(NR_GROUP_LIST, groupList, errMsg) Then
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
            "Group list read failed; err=" & errMsg
        r = plat_runtime.ResultFail(actionName, RunId, "VALIDATION", _
                "分保组列表读取失败：" & errMsg, errMsg)
        plat_runtime.ResultStore r
        Exit Sub
    End If

    Dim totalCount  As Long
    totalCount = UBound(groupList) - LBound(groupList) + 1

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
        "Batch start; totalGroups=" & totalCount & _
        " action=" & actionName & " run_id=" & RunId

    ' --- Step 2: Loop over groups ---
    Dim failCount   As Long
    Dim failLog     As String
    Dim i           As Long
    failCount = 0
    failLog = vbNullString

    For i = LBound(groupList) To UBound(groupList)
        Dim GroupName   As String
        GroupName = groupList(i)
        If Len(GroupName) = 0 Then GoTo NextGroup

        ' Inject current group name into config cell so RunAssumptionMapping reads it
        Dim WriteErr    As String
        If Not biz_io.WriteNamedRangeValue(NR_ASMP_GROUP, GroupName, WriteErr) Then
            plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
                "Failed to set group config; group=" & GroupName & " err=" & WriteErr
            failCount = failCount + 1
            failLog = failLog & GroupName & "(config_write_fail);"
            GoTo NextGroup
        End If

        ' Execute single-group pipeline under SafeExecute
        Dim groupResult As tResult
        groupResult = plat_runtime.SafeExecute( _
                          actionName & "[" & GroupName & "]", _
                          THIS_MODULE & ".RunAssumptionMapping")

        If Not groupResult.Ok Then
            plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
                "Group failed (continuing batch); group=" & GroupName & _
                " code=" & groupResult.Code & " msg=" & groupResult.UserMsg
            failCount = failCount + 1
            failLog = failLog & GroupName & "(" & groupResult.Code & ");"
        Else
            plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
                "Group complete; group=" & GroupName & _
                " dataRows=" & CStr(groupResult.value)
        End If

NextGroup:
    Next i

    ' --- Step 3: Build batch summary result ---
    Dim successCount As Long
    successCount = totalCount - failCount

    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RunAllAssumptions", _
        "Batch complete; total=" & totalCount & _
        " success=" & successCount & " failed=" & failCount & _
        " action=" & actionName & " run_id=" & RunId

    If failCount = 0 Then
        r = plat_runtime.ResultOk(actionName, RunId, CLng(successCount))
    Else
        r = plat_runtime.ResultFail(actionName, RunId, "UNHANDLED", _
                "批量执行部分失败：成功 " & successCount & " / 共 " & totalCount, _
                "failedGroups=" & failLog)
    End If

    plat_runtime.ResultStore r

End Sub

' ==============================================================================================
' SECTION 06: PRIVATE HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [T] TSegmentConfig
'
' 功能说明      : 分段运行所需配置的内部聚合类型
'               : 由 p_ReadSegmentConfig 填充，由 RunLossTransform / RunSubLoBAggregate 消费
' Purpose       : Internal config bundle for segment-level pipeline runs
' ----------------------------------------------------------------------------------------------
Private Type TSegmentConfig
    TenKFilePath    As String
    Segment         As String
    LossSheetName   As String
    startRow        As Long
    MatThreshold    As Double
End Type

' ----------------------------------------------------------------------------------------------
' [T] TGroupConfig
'
' 功能说明      : 分保组运行所需配置的内部聚合类型
'               : 由 p_ReadGroupConfig 填充，由 RunAssumptionMapping 消费
' Purpose       : Internal config bundle for group-level assumption mapping runs
' ----------------------------------------------------------------------------------------------
Private Type TGroupConfig
    GroupName       As String
    AsmpSheetName   As String
End Type

' ----------------------------------------------------------------------------------------------
' [f] p_ReadSegmentConfig
'
' 功能说明      : 读取并验证单分段运行所需的所有命名区域配置值
'               : 任一必填项缺失或无效均返回 False + errMsg（Fail Fast）
'               : MatThreshold 为可选配置；缺失时使用 DEFAULT_MAT_THRESH 并记录 LogWarn
' 参数          : outCfg  - 输出：填充后的 TSegmentConfig（成功时有效）
'               : errMsg  - 输出：失败时的错误说明
' 返回          : Boolean - True=配置完整有效；False=缺失或无效，errMsg 已填充
' ----------------------------------------------------------------------------------------------
Private Function p_ReadSegmentConfig(ByRef outCfg As TSegmentConfig, _
                                      ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    p_ReadSegmentConfig = False

    Dim v   As Variant
    Dim e   As String

    ' File path
    If Not biz_io.ReadNamedRangeValue(NR_TENK_FILE_PATH, v, e) Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_FILE_PATH & " read failed; " & e
        Exit Function
    End If
    outCfg.TenKFilePath = Trim$(CStr(v))
    If Len(outCfg.TenKFilePath) = 0 Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_FILE_PATH & " is empty"
        Exit Function
    End If

    ' Segment
    If Not biz_io.ReadNamedRangeValue(NR_TENK_SEGMENT, v, e) Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_SEGMENT & " read failed; " & e
        Exit Function
    End If
    outCfg.Segment = Trim$(CStr(v))
    If Len(outCfg.Segment) = 0 Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_SEGMENT & " is empty"
        Exit Function
    End If

    ' Loss sheet name
    If Not biz_io.ReadNamedRangeValue(NR_TENK_LOSS_SHEET, v, e) Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_LOSS_SHEET & " read failed; " & e
        Exit Function
    End If
    outCfg.LossSheetName = Trim$(CStr(v))
    If Len(outCfg.LossSheetName) = 0 Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_LOSS_SHEET & " is empty"
        Exit Function
    End If

    ' Start row
    If Not biz_io.ReadNamedRangeValue(NR_TENK_START_ROW, v, e) Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_START_ROW & " read failed; " & e
        Exit Function
    End If
    If Not IsNumeric(v) Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_START_ROW & " is not numeric; value=[" & CStr(v) & "]"
        Exit Function
    End If
    outCfg.startRow = CLng(v)
    If outCfg.startRow < 1 Then
        errMsg = "p_ReadSegmentConfig: " & NR_TENK_START_ROW & " must be >= 1; got " & outCfg.startRow
        Exit Function
    End If

    ' Materiality threshold (optional; default on missing/invalid)
    If Not biz_io.ReadNamedRangeValue(NR_MAT_THRESHOLD, v, e) Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "p_ReadSegmentConfig", _
            NR_MAT_THRESHOLD & " not found; using default=" & DEFAULT_MAT_THRESH
        outCfg.MatThreshold = DEFAULT_MAT_THRESH
    ElseIf Not IsNumeric(v) Or CDbl(v) <= 0 Then
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "p_ReadSegmentConfig", _
            NR_MAT_THRESHOLD & " value invalid [" & CStr(v) & "]; using default=" & DEFAULT_MAT_THRESH
        outCfg.MatThreshold = DEFAULT_MAT_THRESH
    Else
        outCfg.MatThreshold = CDbl(v)
    End If

    p_ReadSegmentConfig = True

End Function

' ----------------------------------------------------------------------------------------------
' [f] p_ReadGroupConfig
'
' 功能说明      : 读取并验证单分保组运行所需的所有命名区域配置值
'               : AsmpSheetName 由 GroupName + ASMP_SHEET_SUFFIX 派生，不读取独立命名区域
'               : 任一必填项缺失或无效均返回 False + errMsg（Fail Fast）
' 参数          : outCfg  - 输出：填充后的 TGroupConfig（成功时有效）
'               : errMsg  - 输出：失败时的错误说明
' 返回          : Boolean - True=配置完整有效；False=缺失或无效，errMsg 已填充
' ----------------------------------------------------------------------------------------------
Private Function p_ReadGroupConfig(ByRef outCfg As TGroupConfig, _
                                    ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    p_ReadGroupConfig = False

    Dim v   As Variant
    Dim e   As String

    ' Group name
    If Not biz_io.ReadNamedRangeValue(NR_ASMP_GROUP, v, e) Then
        errMsg = "p_ReadGroupConfig: " & NR_ASMP_GROUP & " read failed; " & e
        Exit Function
    End If
    outCfg.GroupName = Trim$(CStr(v))
    If Len(outCfg.GroupName) = 0 Then
        errMsg = "p_ReadGroupConfig: " & NR_ASMP_GROUP & " is empty"
        Exit Function
    End If

    ' Derive Asmp sheet name from group name + suffix
    outCfg.AsmpSheetName = outCfg.GroupName & ASMP_SHEET_SUFFIX

    p_ReadGroupConfig = True

End Function

' ----------------------------------------------------------------------------------------------
' [f] p_ReadColumnList
'
' 功能说明      : 从命名区域读取单列向量，返回去空后的 String 数组
'               : 命名区域为多行单列时按行顺序展开；空格行自动跳过
'               : 若有效值为零则返回 False + errMsg
' 参数          : rangeName   - 命名区域名称（单列多行）
'               : outList()   - 输出：String 数组（1-based，仅包含非空值）
'               : errMsg      - 输出：失败时的错误说明
' 返回          : Boolean - True=至少读取到一个有效值；False=失败，errMsg 已填充
' ----------------------------------------------------------------------------------------------
Private Function p_ReadColumnList(ByVal rangeName As String, _
                                   ByRef outList() As String, _
                                   ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    p_ReadColumnList = False

    Dim rawArray    As Variant
    Dim e           As String

    If Not biz_io.ReadNamedRangeArray(rangeName, rawArray, e) Then
        errMsg = "p_ReadColumnList: " & rangeName & " read failed; " & e
        Exit Function
    End If

    ' Count valid (non-empty) entries
    Dim validCount  As Long
    Dim i           As Long
    validCount = 0

    For i = LBound(rawArray, 1) To UBound(rawArray, 1)
        Dim cellVal As String
        cellVal = Trim$(CStr(rawArray(i, 1)))
        If Len(cellVal) > 0 Then validCount = validCount + 1
    Next i

    If validCount = 0 Then
        errMsg = "p_ReadColumnList: " & rangeName & " contains no non-empty values"
        Exit Function
    End If

    ' Build output string array (1-based)
    ReDim outList(1 To validCount)
    Dim idx As Long
    idx = 0
    For i = LBound(rawArray, 1) To UBound(rawArray, 1)
        cellVal = Trim$(CStr(rawArray(i, 1)))
        If Len(cellVal) > 0 Then
            idx = idx + 1
            outList(idx) = cellVal
        End If
    Next i

    p_ReadColumnList = True

End Function


