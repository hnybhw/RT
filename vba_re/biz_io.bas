Attribute VB_Name = "biz_io"
' ==============================================================================================
' MODULE NAME       : biz_io
' LAYER             : Business
' PURPOSE           : Business IO layer. Owns all worksheet read/write operations for the
'                     reinsurance analysis platform. Provides typed read contracts (named range
'                     and array reads) and typed write contracts (array flush, header write,
'                     clear/reset) for all known output sheets. No calculation logic allowed.
' DEPENDS           : plat_context  v2.1.2 (GetWorksheet, GetLogger)
'                     plat_runtime  v2.1.2 (LogInfo, LogWarn, LogError)
'                     core_files    v2.1.0 (GetLastRowSafely, GetWorksheetByName)
'                     core_utils    v2.1.1 (EMPTY_VALUE, IsValidArray)
'                     Excel Object Model (Worksheet, Range)
' NOTE              : - This module is the sole owner of Excel worksheet IO within the
'                       Business layer. Domain and Entry modules must not access sheets directly.
'                     - All public functions return Boolean (success flag) with ByRef errMsg.
'                       Callers must check return value before consuming out-parameters.
'                     - Write operations are array-bulk-only (single Range.Value assignment).
'                       Cell-by-cell writes are prohibited for performance reasons.
'                     - Read operations return Variant arrays; EMPTY_VALUE signals failure.
'                     - Sheet names are resolved via plat_context; biz_io does not hold
'                       hardcoded Worksheet object references.
'                     - Clear operations preserve row 1 (header row) by contract.
'                     - This module does not manage Application state (ScreenUpdating etc.);
'                       that responsibility belongs to plat_runtime.
' STATUS            : Draft
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Legacy Baseline): Worksheet read/write operations were inline within script-style
'                             orchestration modules (app_04_tenk, app_05_asmp, app_06_ws);
'                             no dedicated IO boundary existed.
'   - Init (Design): Array writes mixed cell-by-cell and bulk patterns without unified policy;
'                    no contract for failure semantics or header row preservation.
'   - Init (Scope): Sheet access coupled directly to business logic; no separation between
'                   IO, domain calculation, and orchestration.

' v2.0.0
'   - Init (Architecture): Introduced biz_io as the dedicated Business IO boundary module
'                          under the three-layer model (Core / Platform / Business).
'   - Init (Contract): Established Boolean + ByRef errMsg as the unified public API contract
'                      for all read and write operations; EMPTY_VALUE for read failures.
'   - Init (Boundary): Defined strict separation between IO (biz_io), calculation (biz_domain),
'                      and orchestration (biz_entry); no business logic permitted here.
'   - Init (Write Policy): Mandated bulk array write (Range.Value = array) as the sole write
'                          mechanism; cell-by-cell iteration prohibited.
'   - Init (Read Policy): Named range reads and last-row-anchored area reads established
'                         as the two supported read patterns.
'   - Init (Sheet Identity): Sheet resolution delegated to plat_context; no direct
'                            Worksheet object storage in module state.
'   - Init (Clear Contract): ClearOutputSheet and ClearDataRange preserve header row (row 1)
'                            by contract; callers must not re-clear headers.
'   - Init (Failure Semantics): All failures return False + errMsg; no exceptions raised,
'                               no silent fallbacks, no default value substitution.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 00: MODULE CONSTANTS
'
' SECTION 01: NAMED RANGE READ
'   [F] ReadNamedRangeValue     - Read single-cell named range value (scalar)
'   [F] ReadNamedRangeArray     - Read multi-cell named range into Variant array
'
' SECTION 02: SHEET AREA READ
'   [F] ReadSheetArea           - Read contiguous data area from sheet (header row + data)
'   [F] ReadColumnVector        - Read single column from sheet as 1D Variant array
'
' SECTION 03: SHEET WRITE
'   [F] WriteArrayToSheet       - Bulk-write 2D Variant array to sheet at anchor cell (overwrite)
'   [F] AppendArrayToSheet      - Append 2D Variant array below last used row on sheet
'   [F] WriteHeaderRow          - Write 1D header array to row 1 of target sheet
'
' SECTION 04: SHEET CLEAR / RESET
'   [F] ClearOutputSheet        - Clear data rows (row 2 onward), preserve header row 1
'   [F] ClearDataRange          - Clear a specific named range contents (no format clear)
'
' SECTION 05: SHEET EXISTENCE & VALIDATION
'   [F] RequireOutputSheet      - Assert output sheet exists; fail-fast with errMsg if absent
'   [F] SheetIsEmpty            - Query whether a sheet has no data below row 1
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
Private Const THIS_MODULE   As String = "biz_io"

' Maximum rows allowed in a single bulk-write operation.
' Protects against accidental oversized writes; caller must chunk if needed.
Private Const MAX_WRITE_ROWS As Long = 1100000

' Maximum columns allowed in a single bulk-write operation.
Private Const MAX_WRITE_COLS As Long = 2000

' ==============================================================================================
' SECTION 01: NAMED RANGE READ
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ReadNamedRangeValue
'
' 功能说明      : 读取指定工作簿内单格命名区域的标量值，不支持多格区域
' 参数          : rangeName  - 命名区域名称（字符串）
'               : outValue   - 输出：读取到的值（成功时有效）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=读取成功，outValue 可用；False=失败，errMsg 已填充
' Purpose       : Scalar named range read; single source of truth for config / parameter reads
' Contract      : Business / IO Read
' Side Effects  : None (read-only worksheet access)
' ----------------------------------------------------------------------------------------------
Public Function ReadNamedRangeValue(ByVal rangeName As String, _
                                    ByRef outValue As Variant, _
                                    ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ReadNamedRangeValue = False
    outValue = core_utils.EMPTY_VALUE

    If Len(Trim$(rangeName)) = 0 Then
        errMsg = THIS_MODULE & ".ReadNamedRangeValue: rangeName is empty"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadNamedRangeValue", errMsg
        Exit Function
    End If

    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        errMsg = THIS_MODULE & ".ReadNamedRangeValue: named range not found [" & rangeName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadNamedRangeValue", errMsg
        Exit Function
    End If

    If rng.Cells.CountLarge > 1 Then
        errMsg = THIS_MODULE & ".ReadNamedRangeValue: named range [" & rangeName & "] spans " & _
                 rng.Cells.CountLarge & " cells; use ReadNamedRangeArray for multi-cell reads"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadNamedRangeValue", errMsg
        Exit Function
    End If

    outValue = rng.Cells(1, 1).value
    ReadNamedRangeValue = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadNamedRangeValue", _
        "Read [" & rangeName & "] = " & CStr(outValue)
End Function

' ----------------------------------------------------------------------------------------------
' [F] ReadNamedRangeArray
'
' 功能说明      : 读取多格命名区域为二维 Variant 数组（单行/单列区域同样返回二维数组）
' 参数          : rangeName  - 命名区域名称（字符串）
'               : outArray   - 输出：读取到的二维数组（成功时有效）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=读取成功，outArray 可用；False=失败，errMsg 已填充
' Purpose       : Multi-cell named range read; normalizes single-cell result to 2D array
' Contract      : Business / IO Read
' Side Effects  : None (read-only worksheet access)
' ----------------------------------------------------------------------------------------------
Public Function ReadNamedRangeArray(ByVal rangeName As String, _
                                    ByRef outArray As Variant, _
                                    ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ReadNamedRangeArray = False
    outArray = core_utils.EMPTY_VALUE

    If Len(Trim$(rangeName)) = 0 Then
        errMsg = THIS_MODULE & ".ReadNamedRangeArray: rangeName is empty"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadNamedRangeArray", errMsg
        Exit Function
    End If

    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        errMsg = THIS_MODULE & ".ReadNamedRangeArray: named range not found [" & rangeName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadNamedRangeArray", errMsg
        Exit Function
    End If

    Dim raw As Variant
    raw = rng.value

    ' Normalize scalar to 2D array so callers always receive a consistent type
    If Not IsArray(raw) Then
        Dim normalized(1 To 1, 1 To 1) As Variant
        normalized(1, 1) = raw
        outArray = normalized
    Else
        outArray = raw
    End If

    ReadNamedRangeArray = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadNamedRangeArray", _
        "Read [" & rangeName & "] rows=" & (UBound(outArray, 1) - LBound(outArray, 1) + 1) & _
        " cols=" & (UBound(outArray, 2) - LBound(outArray, 2) + 1)
End Function

' ==============================================================================================
' SECTION 02: SHEET AREA READ
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ReadSheetArea
'
' 功能说明      : 从指定工作表读取连续数据区域（含表头行），返回二维 Variant 数组
'               : 数据范围由锚点列（anchorCol）的最后非空行决定
' 参数          : sheetName  - 工作表名称
'               : startCell  - 数据区域左上角单元格地址（如 "A1"）
'               : anchorCol  - 用于确定最后数据行的列号（1-based）
'               : outArray   - 输出：读取到的二维数组（行数包含表头）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=读取成功，outArray 可用；False=失败，errMsg 已填充
' Purpose       : Primary sheet read contract; anchor-column last-row detection guards against
'                 sparse trailing rows contaminating the read area
' Contract      : Business / IO Read
' Side Effects  : None (read-only worksheet access)
' ----------------------------------------------------------------------------------------------
Public Function ReadSheetArea(ByVal sheetName As String, _
                              ByVal startCell As String, _
                              ByVal anchorCol As Long, _
                              ByRef outArray As Variant, _
                              ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ReadSheetArea = False
    outArray = core_utils.EMPTY_VALUE

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    If anchorCol < 1 Then
        errMsg = THIS_MODULE & ".ReadSheetArea: anchorCol must be >= 1 [sheet=" & sheetName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadSheetArea", errMsg
        Exit Function
    End If

    Dim anchorErrMsg As String
    Dim lastRow As Long
    lastRow = core_files.GetLastRowSafely(ws, anchorCol, anchorErrMsg)
    If lastRow = 0 Then
        errMsg = THIS_MODULE & ".ReadSheetArea: GetLastRowSafely failed [sheet=" & sheetName & _
                 "] reason=" & anchorErrMsg
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadSheetArea", errMsg
        Exit Function
    End If

    Dim anchorRange As Range
    On Error Resume Next
    Set anchorRange = ws.Range(startCell)
    On Error GoTo 0

    If anchorRange Is Nothing Then
        errMsg = THIS_MODULE & ".ReadSheetArea: invalid startCell address [" & startCell & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadSheetArea", errMsg
        Exit Function
    End If

    Dim startRow As Long
    Dim startColNum As Long
    startRow = anchorRange.row
    startColNum = anchorRange.Column

    If lastRow < startRow Then
        errMsg = THIS_MODULE & ".ReadSheetArea: lastRow (" & lastRow & ") < startRow (" & _
                 startRow & ") [sheet=" & sheetName & "] - sheet appears empty"
        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadSheetArea", errMsg
        Exit Function
    End If

    ' Determine column extent from row 1 of the sheet area (header-driven width)
    Dim lastCol As Long
    lastCol = ws.Cells(startRow, ws.Columns.count).End(xlToLeft).Column
    If lastCol < startColNum Then lastCol = startColNum

    Dim readRange As Range
    Set readRange = ws.Range(ws.Cells(startRow, startColNum), ws.Cells(lastRow, lastCol))
    outArray = readRange.value

    ' Normalize scalar edge case (1x1 range) to 2D array
    If Not IsArray(outArray) Then
        Dim norm(1 To 1, 1 To 1) As Variant
        norm(1, 1) = outArray
        outArray = norm
    End If

    ReadSheetArea = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadSheetArea", _
        "Read sheet=[" & sheetName & "] rows=" & (lastRow - startRow + 1) & _
        " cols=" & (lastCol - startColNum + 1)
End Function

' ----------------------------------------------------------------------------------------------
' [F] ReadColumnVector
'
' 功能说明      : 从指定工作表读取单列数据，返回一维 Variant 数组（跳过表头行1）
' 参数          : sheetName  - 工作表名称
'               : colIndex   - 列号（1-based）
'               : outVector  - 输出：读取到的一维数组（不含表头，下标 1-based）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=读取成功，outVector 可用；False=失败，errMsg 已填充
' Purpose       : Single-column read for list extraction (e.g. segment list, TID list)
' Contract      : Business / IO Read
' Side Effects  : None (read-only worksheet access)
' ----------------------------------------------------------------------------------------------
Public Function ReadColumnVector(ByVal sheetName As String, _
                                 ByVal colIndex As Long, _
                                 ByRef outVector As Variant, _
                                 ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ReadColumnVector = False
    outVector = core_utils.EMPTY_VALUE

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    If colIndex < 1 Then
        errMsg = THIS_MODULE & ".ReadColumnVector: colIndex must be >= 1 [sheet=" & sheetName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadColumnVector", errMsg
        Exit Function
    End If

    Dim anchorErrMsg As String
    Dim lastRow As Long
    lastRow = core_files.GetLastRowSafely(ws, colIndex, anchorErrMsg)
    If lastRow = 0 Then
        errMsg = THIS_MODULE & ".ReadColumnVector: GetLastRowSafely failed [sheet=" & sheetName & _
                 "] reason=" & anchorErrMsg
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ReadColumnVector", errMsg
        Exit Function
    End If

    ' Data starts at row 2 (row 1 is header by contract)
    Dim dataStartRow As Long
    dataStartRow = 2
    If lastRow < dataStartRow Then
        errMsg = THIS_MODULE & ".ReadColumnVector: no data rows found below header [sheet=" & _
                 sheetName & " col=" & colIndex & "]"
        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadColumnVector", errMsg
        Exit Function
    End If

    Dim rowCount As Long
    rowCount = lastRow - dataStartRow + 1
    Dim raw As Variant
    raw = ws.Range(ws.Cells(dataStartRow, colIndex), ws.Cells(lastRow, colIndex)).value

    ' Flatten 2D single-column array to 1D
    Dim vec() As Variant
    ReDim vec(1 To rowCount)
    Dim i As Long
    For i = 1 To rowCount
        vec(i) = raw(i, 1)
    Next i

    outVector = vec
    ReadColumnVector = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ReadColumnVector", _
        "Read column vector sheet=[" & sheetName & "] col=" & colIndex & " count=" & rowCount
End Function

' ==============================================================================================
' SECTION 03: SHEET WRITE
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] WriteArrayToSheet
'
' 功能说明      : 将二维 Variant 数组以覆盖模式批量写入目标工作表的锚点单元格起始区域
'               : 不清除目标区域以外的现有数据；调用方负责预先清空（如需）
' 参数          : sheetName  - 目标工作表名称
'               : anchorCell - 写入起始单元格地址（如 "A1"）
'               : dataArray  - 二维 Variant 数组（必须已初始化，不接受 EMPTY_VALUE）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=写入成功；False=失败，errMsg 已填充
' Purpose       : Primary bulk write contract; single Range.Value assignment for performance
' Contract      : Business / IO Write
' Side Effects  : Writes to worksheet cells within the computed target range
' ----------------------------------------------------------------------------------------------
Public Function WriteArrayToSheet(ByVal sheetName As String, _
                                  ByVal anchorCell As String, _
                                  ByRef dataArray As Variant, _
                                  ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    WriteArrayToSheet = False

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    If Not p_ValidateWriteArray(dataArray, errMsg) Then
        errMsg = THIS_MODULE & ".WriteArrayToSheet [sheet=" & sheetName & "]: " & errMsg
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "WriteArrayToSheet", errMsg
        Exit Function
    End If

    Dim anchor As Range
    On Error Resume Next
    Set anchor = ws.Range(anchorCell)
    On Error GoTo 0
    If anchor Is Nothing Then
        errMsg = THIS_MODULE & ".WriteArrayToSheet: invalid anchorCell [" & anchorCell & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "WriteArrayToSheet", errMsg
        Exit Function
    End If

    Dim nRows As Long
    Dim nCols As Long
    nRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    nCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1

    Dim targetRange As Range
    Set targetRange = ws.Range(anchor, anchor.Offset(nRows - 1, nCols - 1))

    On Error GoTo WriteErr
    targetRange.value = dataArray
    On Error GoTo 0

    WriteArrayToSheet = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "WriteArrayToSheet", _
        "Wrote sheet=[" & sheetName & "] anchor=" & anchorCell & _
        " rows=" & nRows & " cols=" & nCols
    Exit Function

WriteErr:
    errMsg = THIS_MODULE & ".WriteArrayToSheet: write failed [sheet=" & sheetName & _
             " anchor=" & anchorCell & "] err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "WriteArrayToSheet", errMsg
End Function

' ----------------------------------------------------------------------------------------------
' [F] AppendArrayToSheet
'
' 功能说明      : 将二维 Variant 数组追加写入目标工作表的最后已用行下方
'               : 锚点列（anchorCol）用于确定当前最后已用行；表头行（row 1）受保护不被追加覆盖
' 参数          : sheetName  - 目标工作表名称
'               : anchorCol  - 确定最后已用行的列号（1-based）
'               : dataArray  - 二维 Variant 数组（不含表头行）
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=追加成功；False=失败，errMsg 已填充
' Purpose       : Append contract for incremental writes (e.g. segment-by-segment accumulation)
' Contract      : Business / IO Write
' Side Effects  : Writes to worksheet rows below the current last used row
' ----------------------------------------------------------------------------------------------
Public Function AppendArrayToSheet(ByVal sheetName As String, _
                                   ByVal anchorCol As Long, _
                                   ByRef dataArray As Variant, _
                                   ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    AppendArrayToSheet = False

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    If Not p_ValidateWriteArray(dataArray, errMsg) Then
        errMsg = THIS_MODULE & ".AppendArrayToSheet [sheet=" & sheetName & "]: " & errMsg
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "AppendArrayToSheet", errMsg
        Exit Function
    End If

    If anchorCol < 1 Then
        errMsg = THIS_MODULE & ".AppendArrayToSheet: anchorCol must be >= 1 [sheet=" & sheetName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "AppendArrayToSheet", errMsg
        Exit Function
    End If

    Dim anchorErrMsg As String
    Dim lastRow As Long
    lastRow = core_files.GetLastRowSafely(ws, anchorCol, anchorErrMsg)
    ' lastRow = 0 means sheet is empty (only header or fully blank); start at row 2
    Dim appendRow As Long
    If lastRow = 0 Then
        appendRow = 2
    Else
        appendRow = lastRow + 1
    End If

    ' Guard: never overwrite row 1 (header row contract)
    If appendRow < 2 Then appendRow = 2

    Dim nRows As Long
    Dim nCols As Long
    nRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    nCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1

    Dim anchor As Range
    Set anchor = ws.Cells(appendRow, 1)
    Dim targetRange As Range
    Set targetRange = ws.Range(anchor, anchor.Offset(nRows - 1, nCols - 1))

    On Error GoTo AppendErr
    targetRange.value = dataArray
    On Error GoTo 0

    AppendArrayToSheet = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "AppendArrayToSheet", _
        "Appended sheet=[" & sheetName & "] startRow=" & appendRow & _
        " rows=" & nRows & " cols=" & nCols
    Exit Function

AppendErr:
    errMsg = THIS_MODULE & ".AppendArrayToSheet: write failed [sheet=" & sheetName & _
             " appendRow=" & appendRow & "] err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "AppendArrayToSheet", errMsg
End Function

' ----------------------------------------------------------------------------------------------
' [F] WriteHeaderRow
'
' 功能说明      : 将一维表头数组写入目标工作表的第一行（A1 起始）
'               : 仅写入 row 1；不影响任何数据行
' 参数          : sheetName   - 目标工作表名称
'               : headerArray - 一维 Variant 数组，包含各列表头名称
'               : errMsg      - 输出：失败时的错误说明
' 返回          : Boolean - True=写入成功；False=失败，errMsg 已填充
' Purpose       : Dedicated header write; separates structural setup from data writes
' Contract      : Business / IO Write
' Side Effects  : Writes to row 1 of the target worksheet (columns 1 to header count)
' ----------------------------------------------------------------------------------------------
Public Function WriteHeaderRow(ByVal sheetName As String, _
                                ByRef headerArray As Variant, _
                                ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    WriteHeaderRow = False

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    If Not IsArray(headerArray) Then
        errMsg = THIS_MODULE & ".WriteHeaderRow: headerArray is not an array [sheet=" & sheetName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "WriteHeaderRow", errMsg
        Exit Function
    End If

    Dim nCols As Long
    nCols = UBound(headerArray) - LBound(headerArray) + 1
    If nCols < 1 Then
        errMsg = THIS_MODULE & ".WriteHeaderRow: headerArray is empty [sheet=" & sheetName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "WriteHeaderRow", errMsg
        Exit Function
    End If

    ' Reshape 1D to 2D row array for Range.Value assignment
    Dim h2D(1 To 1, 1 To 1) As Variant
    ReDim h2D(1 To 1, 1 To nCols)
    Dim i As Long
    For i = 1 To nCols
        h2D(1, i) = headerArray(LBound(headerArray) + i - 1)
    Next i

    Dim targetRange As Range
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))

    On Error GoTo HeaderErr
    targetRange.value = h2D
    On Error GoTo 0

    WriteHeaderRow = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "WriteHeaderRow", _
        "Wrote header sheet=[" & sheetName & "] cols=" & nCols
    Exit Function

HeaderErr:
    errMsg = THIS_MODULE & ".WriteHeaderRow: write failed [sheet=" & sheetName & _
             "] err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "WriteHeaderRow", errMsg
End Function

' ==============================================================================================
' SECTION 04: SHEET CLEAR / RESET
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ClearOutputSheet
'
' 功能说明      : 清空目标工作表的数据行（row 2 起），保留第一行表头不清除
'               : 仅清除内容（ClearContents），不清除格式
' 参数          : sheetName  - 目标工作表名称
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=清除成功；False=失败，errMsg 已填充
' Purpose       : Standard pre-write reset; header-row preservation is a hard contract
' Contract      : Business / IO Write (destructive within data rows)
' Side Effects  : Clears contents of rows 2 through last used row on the target sheet
' ----------------------------------------------------------------------------------------------
Public Function ClearOutputSheet(ByVal sheetName As String, _
                                 ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ClearOutputSheet = False

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    Dim lastRow As Long
    ' Use column 1 as anchor; if empty, nothing to clear
    Dim anchorErrMsg As String
    lastRow = core_files.GetLastRowSafely(ws, 1, anchorErrMsg)
    If lastRow <= 1 Then
        ' Sheet is empty or header-only: nothing to clear, succeed silently
        ClearOutputSheet = True
        plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ClearOutputSheet", _
            "Sheet=[" & sheetName & "] already empty (lastRow=" & lastRow & "), no action"
        Exit Function
    End If

    On Error GoTo ClearErr
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Columns.count)).ClearContents
    On Error GoTo 0

    ClearOutputSheet = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ClearOutputSheet", _
        "Cleared sheet=[" & sheetName & "] rows 2 to " & lastRow
    Exit Function

ClearErr:
    errMsg = THIS_MODULE & ".ClearOutputSheet: clear failed [sheet=" & sheetName & _
             "] err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "ClearOutputSheet", errMsg
End Function

' ----------------------------------------------------------------------------------------------
' [F] ClearDataRange
'
' 功能说明      : 清空指定命名区域的内容（ClearContents），不清除格式
'               : 适用于参数区域或局部数据区域的重置
' 参数          : rangeName  - 命名区域名称
'               : errMsg     - 输出：失败时的错误说明
' 返回          : Boolean - True=清除成功；False=失败，errMsg 已填充
' Purpose       : Targeted named range clear for parameter zones and partial resets
' Contract      : Business / IO Write (destructive within named range)
' Side Effects  : Clears contents of the specified named range cells
' ----------------------------------------------------------------------------------------------
Public Function ClearDataRange(ByVal rangeName As String, _
                                ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    ClearDataRange = False

    If Len(Trim$(rangeName)) = 0 Then
        errMsg = THIS_MODULE & ".ClearDataRange: rangeName is empty"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ClearDataRange", errMsg
        Exit Function
    End If

    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        errMsg = THIS_MODULE & ".ClearDataRange: named range not found [" & rangeName & "]"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "ClearDataRange", errMsg
        Exit Function
    End If

    On Error GoTo ClearRngErr
    rng.ClearContents
    On Error GoTo 0

    ClearDataRange = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "ClearDataRange", _
        "Cleared named range [" & rangeName & "]"
    Exit Function

ClearRngErr:
    errMsg = THIS_MODULE & ".ClearDataRange: clear failed [range=" & rangeName & _
             "] err=" & Err.Description
    plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "ClearDataRange", errMsg
End Function

' ==============================================================================================
' SECTION 05: SHEET EXISTENCE & VALIDATION
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] RequireOutputSheet
'
' 功能说明      : 断言指定工作表存在；不存在则立即填充 errMsg 并返回 False（Fail Fast 语义）
'               : 用于业务流程入口的前置检查，确保 IO 目标在执行前可用
' 参数          : sheetName  - 工作表名称
'               : errMsg     - 输出：失败时的错误说明（含明确的操作建议）
' 返回          : Boolean - True=工作表存在；False=不存在，errMsg 已填充
' Purpose       : Fail-fast sheet existence assertion for Business entry pre-condition checks
' Contract      : Business / IO Validation (read-only query)
' Side Effects  : None (read-only)
' ----------------------------------------------------------------------------------------------
Public Function RequireOutputSheet(ByVal sheetName As String, _
                                   ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    RequireOutputSheet = False

    If Len(Trim$(sheetName)) = 0 Then
        errMsg = THIS_MODULE & ".RequireOutputSheet: sheetName is empty"
        plat_runtime.LogWarn BIZ_LAYER, THIS_MODULE, "RequireOutputSheet", errMsg
        Exit Function
    End If

    Dim ws As Worksheet
    Dim ignored As String
    If Not p_ResolveSheet(sheetName, ws, ignored) Then
        errMsg = THIS_MODULE & ".RequireOutputSheet: required output sheet not found [" & _
                 sheetName & "] - verify workbook structure before running"
        plat_runtime.LogError BIZ_LAYER, THIS_MODULE, "RequireOutputSheet", errMsg
        Exit Function
    End If

    RequireOutputSheet = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "RequireOutputSheet", _
        "Output sheet confirmed [" & sheetName & "]"
End Function

' ----------------------------------------------------------------------------------------------
' [F] SheetIsEmpty
'
' 功能说明      : 查询工作表在表头行（row 1）以下是否无数据
' 参数          : sheetName  - 工作表名称
'               : outIsEmpty - 输出：True=无数据行；False=有数据行
'               : errMsg     - 输出：失败时的错误说明（工作表不存在等）
' 返回          : Boolean - True=查询成功，outIsEmpty 可用；False=查询失败，errMsg 已填充
' Purpose       : Pre-write / pre-read guard to detect whether a sheet has been populated
' Contract      : Business / IO Validation (read-only query)
' Side Effects  : None (read-only)
' ----------------------------------------------------------------------------------------------
Public Function SheetIsEmpty(ByVal sheetName As String, _
                              ByRef outIsEmpty As Boolean, _
                              ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    SheetIsEmpty = False
    outIsEmpty = False

    Dim ws As Worksheet
    If Not p_ResolveSheet(sheetName, ws, errMsg) Then Exit Function

    Dim anchorErrMsg As String
    Dim lastRow As Long
    lastRow = core_files.GetLastRowSafely(ws, 1, anchorErrMsg)

    ' lastRow = 0 (failure) or <= 1 (header only) both mean effectively empty for data purposes
    outIsEmpty = (lastRow <= 1)
    SheetIsEmpty = True
    plat_runtime.LogInfo BIZ_LAYER, THIS_MODULE, "SheetIsEmpty", _
        "sheet=[" & sheetName & "] lastRow=" & lastRow & " isEmpty=" & outIsEmpty
End Function

' ==============================================================================================
' PRIVATE HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] p_ResolveSheet
'
' 功能说明      : 通过 plat_context 解析工作表对象；不存在则填充 errMsg 并返回 False
' 参数          : sheetName - 工作表名称
'               : outWs     - 输出：解析到的 Worksheet 对象
'               : errMsg    - 输出：失败时的错误说明
' 返回          : Boolean - True=解析成功；False=失败
' ----------------------------------------------------------------------------------------------
Private Function p_ResolveSheet(ByVal sheetName As String, _
                                 ByRef outWs As Worksheet, _
                                 ByRef errMsg As String) As Boolean
    p_ResolveSheet = False
    Set outWs = Nothing

    If Len(Trim$(sheetName)) = 0 Then
        errMsg = THIS_MODULE & ".p_ResolveSheet: sheetName is empty"
        Exit Function
    End If

    Dim resolveErrMsg As String
    Set outWs = core_files.GetWorksheetByName(ThisWorkbook, sheetName, resolveErrMsg)

    If outWs Is Nothing Then
        errMsg = THIS_MODULE & ".p_ResolveSheet: worksheet not found [" & sheetName & _
                 "] reason=" & resolveErrMsg
        Exit Function
    End If

    p_ResolveSheet = True
End Function

' ----------------------------------------------------------------------------------------------
' [f] p_ValidateWriteArray
'
' 功能说明      : 验证写入数组是否为有效的二维 Variant 数组，并检查行列数是否在安全阈值内
' 参数          : arr    - 待验证的数组
'               : errMsg - 输出：失败时的错误说明
' 返回          : Boolean - True=数组有效；False=无效，errMsg 已填充
' ----------------------------------------------------------------------------------------------
Private Function p_ValidateWriteArray(ByRef arr As Variant, _
                                       ByRef errMsg As String) As Boolean
    p_ValidateWriteArray = False

    If Not IsArray(arr) Then
        errMsg = "dataArray is not an array"
        Exit Function
    End If

    If core_utils.IsEmptyValue(arr) Then
        errMsg = "dataArray is EMPTY_VALUE"
        Exit Function
    End If

    Dim dims As Integer
    On Error Resume Next
    dims = 0
    Dim ub1 As Long: ub1 = UBound(arr, 1): If Err.Number = 0 Then dims = 1
    Dim ub2 As Long: ub2 = UBound(arr, 2): If Err.Number = 0 Then dims = 2
    Dim e As Long: e = Err.Number
    On Error GoTo 0

    If dims < 2 Then
        errMsg = "dataArray must be a 2D array (dims=" & dims & ")"
        Exit Function
    End If

    Dim nRows As Long
    Dim nCols As Long
    nRows = ub1 - LBound(arr, 1) + 1
    nCols = ub2 - LBound(arr, 2) + 1

    If nRows > MAX_WRITE_ROWS Then
        errMsg = "dataArray row count (" & nRows & ") exceeds MAX_WRITE_ROWS (" & MAX_WRITE_ROWS & ")"
        Exit Function
    End If

    If nCols > MAX_WRITE_COLS Then
        errMsg = "dataArray column count (" & nCols & ") exceeds MAX_WRITE_COLS (" & MAX_WRITE_COLS & ")"
        Exit Function
    End If

    p_ValidateWriteArray = True
End Function




