Attribute VB_Name = "core_utils"
' ==============================================================================================
' MODULE NAME       : core_utils
' LAYER             : core
' PURPOSE           : Core utility toolbox for safe array handling, type conversion/validation,
'                     and safe math helpers used across the project (query-only, no IO).
' DEPENDS           : None
' NOTE              : - Query-only by default: no IO, no Application state changes.
'                     - Performance-first for Excel-centric workloads (mostly 1D/2D Variant arrays).
'                     - Functions are designed to be re-entrant and side-effect free.
' STATUS            : Frozen
' ==============================================================================================
' VERSION HISTORY   :
' v2.0.0
'   - Refactor: split project into layered architecture (Core / Platform / Business).
'   - Freeze: stabilized array inspect/transform, conversion/validate, safe math utilities.
'
' v1.0.0
'   - Initial draft based on legacy implementation, iteratively refined during early refactor.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 01: ARRAY INSPECT VALIDATE
'   [T] ArrayInfo               - Array metadata (dims/bounds/counts)
'   [F] GetArrayInfo            - Fast array inspection (0/1/2D; allocated/bounds)
'   [F] IsArrayValid            - Validates array allocation/dimension/data requirements
'   [F] EnsureArrayDimensions   - Ensures array is 2D and meets minimum size (copy/expand)
'   [F] CalculateSafeArraySize  - Clamps requested size to safe maximum elements
'
' SECTION 02: ARRAY TRANSFORM
'   [F] NormalizeTo1Based       - Normalizes array to 1-based 2D (1D upcasts to Nx1)
'   [f] Normalize1DArray        - Helper: 1D -> 2D (Nx1), 1-based
'   [f] Normalize2DArray        - Helper: 2D -> 2D, 1-based
'   [F] SliceArraySafe          - Slice a 2D array safely with bounds validation, returning a 1-based copy
'   [f] CoerceLongOrDefault     - Safely coerce optional Variant inputs to Long without raising type errors
'   [F] SliceArraySafeFull      - Convenience wrapper for full-range safe slicing (returns a 1-based copy)
'   [F] AppendArrayVertical     - Vertical concat (supports 1D/2D mix, logical cols match)
'
' SECTION 03: CONVERT VALIDATE
'   [F] ToSafeLong              - Safe Long conversion with range check
'   [F] ToSafeDouble            - Safe Double conversion with range check
'   [F] ToSafeString            - Safe String conversion with trim/length options
'   [F] IsNumericSafe           - Safe IsNumeric check
'   [F] IsDateSafe              - Safe IsDate check
'   [F] SanitizeString          - Removes control chars using RegExp (fallback supported)
'   [f] SanitizeStringFallback  - Loop-based sanitizer fallback
'   [f] ParseBoolean            - Helper: parses common boolean representations
'
' SECTION 04: SAFE MATH
'   [F] SafeMultiply            - Safe multiply (finite check + abs clamp)
'   [F] SafeAdd                 - Safe add (finite check + abs clamp)
'   [f] IsFiniteDouble          - Helper: finite (not NaN/Inf) check
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function, [T]=Type
' ==============================================================================================
Option Explicit
Public Const EMPTY_VALUE As Variant = Empty

' ==============================================================================================
' SECTION 01: ARRAY INSPECT VALIDATE
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [T] ArrayInfo
'
' 功能说明      : 数组信息结构体，用于存储数组的维度、上下界及元素数量
' 参数          : None - 这是一个类型定义，无参数
' 返回          : 类型 - 用户自定义类型，包含数组状态和边界信息
' ----------------------------------------------------------------------------------------------
Public Type ArrayInfo
    IsArray As Boolean
    IsAllocated As Boolean      ' True => dim1 bounds readable
    Dims As Long                ' 0/1/2 only (intentional for speed)

    LBound1 As Long
    UBound1 As Long
    Count1  As Long

    LBound2 As Long
    UBound2 As Long
    Count2  As Long
End Type

' ----------------------------------------------------------------------------------------------
' [F] GetArrayInfo
'
' 功能说明      : 获取数组的维度、上下界及元素数量等信息
' 参数          : arr - 要检查的数组变量
'               : info - 用于存储数组信息的结构体
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 是否成功获取数组信息，True表示数组已分配
' Purpose       : Retrieves dimension, bounds and count information of an array
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetArrayInfo(ByRef arr As Variant, _
                             ByRef info As ArrayInfo, _
                             Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString

    ' defaults
    info.IsArray = False
    info.IsAllocated = False
    info.Dims = 0
    info.LBound1 = 0: info.UBound1 = -1: info.Count1 = 0
    info.LBound2 = 0: info.UBound2 = -1: info.Count2 = 0

    If Not IsArray(arr) Then
        errMsg = "Not an array."
        Exit Function
    End If
    info.IsArray = True

    On Error Resume Next
    Err.Clear

    ' ---- Dim 1 (required)
    info.LBound1 = LBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        errMsg = "Array not allocated (dim1 bounds unreadable)."
        GoTo Finalize
    End If

    info.UBound1 = UBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        errMsg = "Array dim1 UBound unreadable."
        GoTo Finalize
    End If

    info.Dims = 1
    info.Count1 = info.UBound1 - info.LBound1 + 1
    If info.Count1 < 0 Then info.Count1 = 0

    ' ---- Dim 2 (best-effort; stop here intentionally)
    info.LBound2 = LBound(arr, 2)
    If Err.Number = 0 Then
        info.UBound2 = UBound(arr, 2)
        If Err.Number = 0 Then
            info.Dims = 2
            info.Count2 = info.UBound2 - info.LBound2 + 1
            If info.Count2 < 0 Then info.Count2 = 0
        Else
            ' rare: LBound2 ok but UBound2 fails => keep as 1D; clear and keep defaults
            Err.Clear
            info.LBound2 = 0: info.UBound2 = -1: info.Count2 = 0
        End If
    Else
        Err.Clear
        info.LBound2 = 0: info.UBound2 = -1: info.Count2 = 0
    End If

Finalize:
    On Error GoTo 0

    info.IsAllocated = (info.Dims > 0)
    GetArrayInfo = info.IsAllocated
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsArrayValid
'
' 功能说明      : 验证数组是否有效，可检查是否为二维数组以及是否包含数据
' 参数          : arr - 要验证的数组变量
'               : requireData - 可选，是否要求数组包含数据，默认为True
'               : require2D - 可选，是否要求为二维数组，默认为False
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 数组是否有效，True表示满足所有验证条件
' Purpose       : Validates if an array meets specified criteria including dimensions and data presence
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsArrayValid(ByRef arr As Variant, _
                             Optional ByVal requireData As Boolean = True, _
                             Optional ByVal require2D As Boolean = False, _
                             Optional ByRef errMsg As String) As Boolean
    Dim info As ArrayInfo
    If Not GetArrayInfo(arr, info, errMsg) Then Exit Function

    If require2D And info.Dims < 2 Then
        errMsg = "2D array required."
        Exit Function
    End If

    If requireData Then
        If info.Count1 <= 0 Then
            errMsg = "Array has no elements."
            Exit Function
        End If
        If require2D And info.Count2 <= 0 Then
            errMsg = "2D array has zero columns."
            Exit Function
        End If
    End If

    IsArrayValid = True
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetArrayDimensions
'
' 功能说明      : 获取数组的维度数量
' 参数          : arr - 要检查的数组变量
' 返回          : Long - 数组的维度数，0表示非数组或未分配数组
' ----------------------------------------------------------------------------------------------
Private Function GetArrayDimensions(ByRef arr As Variant) As Long
    Dim info As ArrayInfo
    If GetArrayInfo(arr, info) Then
        GetArrayDimensions = info.Dims
    Else
        GetArrayDimensions = 0
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] EnsureArrayDimensions
'
' 功能说明      : 确保数组具有指定的最小行数和列数，必要时重新创建或扩展数组
' 参数          : arr - 要检查和调整的数组变量
'               : minRows - 可选，要求的最小行数，默认为1
'               : minCols - 可选，要求的最小列数，默认为1
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 是否成功确保数组维度，True表示数组已满足或调整为所需尺寸
' Purpose       : Ensures an array has at least the specified minimum rows and columns, recreating or resizing if necessary
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function EnsureArrayDimensions(ByRef arr As Variant, _
                                      Optional ByVal minRows As Long = 1, _
                                      Optional ByVal minCols As Long = 1, _
                                      Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    If minRows < 1 Then minRows = 1
    If minCols < 1 Then minCols = 1

    Dim info As ArrayInfo
    If Not GetArrayInfo(arr, info, errMsg) Then
        ReDim arr(1 To minRows, 1 To minCols)
        EnsureArrayDimensions = True
        Exit Function
    End If

    ' Fast exit: already good 2D capacity
    If info.Dims = 2 Then
        If info.Count1 >= minRows And info.Count2 >= minCols Then
            EnsureArrayDimensions = True
            Exit Function
        End If
    End If

    Dim targetRows As Long, targetCols As Long
    targetRows = info.Count1
    If targetRows < minRows Then targetRows = minRows

    If info.Dims = 1 Then
        targetCols = minCols

        Dim newFrom1D() As Variant
        ReDim newFrom1D(1 To targetRows, 1 To targetCols)

        If info.Count1 > 0 Then
            Dim r As Long, src1 As Long
            For r = 1 To info.Count1
                src1 = info.LBound1 + (r - 1)
                newFrom1D(r, 1) = arr(src1)
            Next r
        End If

        arr = newFrom1D
        EnsureArrayDimensions = True
        Exit Function
    End If

    ' 2D resize path
    targetCols = info.Count2
    If targetCols < minCols Then targetCols = minCols

    Dim new2D() As Variant
    ReDim new2D(1 To targetRows, 1 To targetCols)

    If info.Count1 > 0 And info.Count2 > 0 Then
        Dim rr As Long, cc As Long
        Dim srcR As Long, srcC As Long
        For rr = 1 To info.Count1
            srcR = info.LBound1 + (rr - 1)
            For cc = 1 To info.Count2
                srcC = info.LBound2 + (cc - 1)
                new2D(rr, cc) = arr(srcR, srcC)
            Next cc
        Next rr
    End If

    arr = new2D
    EnsureArrayDimensions = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] CalculateSafeArraySize
'
' 功能说明      : 计算安全的数组尺寸，确保数组元素总数不超过指定的最大元素数
' 参数          : requestedRows - 请求的行数
'               : requestedCols - 请求的列数
'               : maxElements - 可选，允许的最大元素总数，默认为1000000
' 返回          : Long() - 包含两个元素的一维数组，result(1)为安全行数，result(2)为安全列数
' Purpose       : Calculates safe array dimensions ensuring total elements don't exceed specified maximum
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function CalculateSafeArraySize(ByVal requestedRows As Long, _
                                       ByVal requestedCols As Long, _
                                       Optional ByVal maxElements As Long = 0) As Long()
    Dim result() As Long
    ReDim result(1 To 2)

    If requestedRows < 1 Then requestedRows = 1
    If requestedCols < 1 Then requestedCols = 1
    If maxElements < 1 Then maxElements = 1000000

    Dim safeCols As Long
    safeCols = requestedCols
    If safeCols > maxElements Then safeCols = maxElements
    If safeCols < 1 Then safeCols = 1

    Dim maxRows As Long
    maxRows = maxElements \ safeCols
    If maxRows < 1 Then maxRows = 1

    Dim safeRows As Long
    safeRows = requestedRows
    If safeRows > maxRows Then safeRows = maxRows

    result(1) = safeRows
    result(2) = safeCols
    CalculateSafeArraySize = result
End Function

' ==============================================================================================
' SECTION 02: ARRAYS TRANSFORM
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] NormalizeTo1Based
'
' 功能说明      : 将数组标准化为 1-based 结构，便于后续统一处理。
'               : - 1D 数组：始终升级为二维 1-based (rows x 1)
'               : - 2D 数组：如需则重建为二维 1-based
'               : - 其他维度：保持原样返回
' 参数          : arr - 原始数组（可为 0-based / 1-based / 1D / 2D）
'               : makeCopy - 是否强制重建副本（对 2D 且已 1-based 时有意义）
' 返回          : Variant - 标准化后的数组（1D -> 2D Nx1；2D -> 2D 1-based）
' Purpose       : Normalize arrays into a consistent 1-based shape (1D upcasts to Nx1 2D).
' Contract      : Core / Query-only.
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function NormalizeTo1Based(ByVal arr As Variant, Optional ByVal makeCopy As Boolean = False) As Variant

    ' 非数组或未分配：原样返回
    If Not IsArrayValid(arr, requireData:=False) Then
        NormalizeTo1Based = arr
        Exit Function
    End If

    Dim info As ArrayInfo
    Dim ok As Boolean
    ok = GetArrayInfo(arr, info)
    If Not ok Then
        NormalizeTo1Based = arr
        Exit Function
    End If

    Select Case info.Dims
        Case 1
            ' 契约：1D 一律升级为 2D Nx1（Excel 写入友好）
            NormalizeTo1Based = Normalize1DArray(arr, info)
            Exit Function

        Case 2
            ' 2D：只有在“已是 1-based 且不强制拷贝”时走 fast path
            If (Not makeCopy) And (info.LBound1 = 1) And (info.LBound2 = 1) Then
                NormalizeTo1Based = arr
            Else
                NormalizeTo1Based = Normalize2DArray(arr, info)
            End If
            Exit Function

        Case Else
            ' >=3D：保持原样（不做额外开销）
            NormalizeTo1Based = arr
            Exit Function
    End Select

End Function

' ----------------------------------------------------------------------------------------------
' [f] Normalize1DArray
'
' 功能说明      : 将一维数组转换为二维单列数组，索引从1开始
' 参数          : arr - 原始一维数组
'               : info - 包含原始数组边界信息的结构体
' 返回          : Variant - 二维单列数组(rows x 1)，行数与原数组元素数相同
' ----------------------------------------------------------------------------------------------
Private Function Normalize1DArray(ByRef arr As Variant, ByRef info As ArrayInfo) As Variant
    Dim rows As Long
    rows = info.Count1
    If rows <= 0 Then
        Normalize1DArray = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To rows, 1 To 1)

    Dim i As Long, srcIdx As Long
    For i = 1 To rows
        srcIdx = info.LBound1 + (i - 1)
        result(i, 1) = arr(srcIdx)
    Next i

    Normalize1DArray = result
End Function

' ----------------------------------------------------------------------------------------------
' [f] Normalize2DArray (Private/Helper)
'
' 功能说明      : 将二维数组的索引规范化为从1开始，若已是1-based则直接返回原数组
' 参数          : arr - 原始二维数组
'               : info - 包含原始数组边界信息的结构体
' 返回          : Variant - 索引从1开始的二维数组
' ----------------------------------------------------------------------------------------------
Private Function Normalize2DArray(ByRef arr As Variant, ByRef info As ArrayInfo) As Variant
    Dim rows As Long, cols As Long
    rows = info.Count1
    cols = info.Count2
    If rows <= 0 Or cols <= 0 Then
        Normalize2DArray = Empty
        Exit Function
    End If

    If info.LBound1 = 1 And info.LBound2 = 1 Then
        Normalize2DArray = arr
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To rows, 1 To cols)

    Dim r As Long, c As Long
    Dim srcR As Long, srcC As Long
    For r = 1 To rows
        srcR = info.LBound1 + (r - 1)
        For c = 1 To cols
            srcC = info.LBound2 + (c - 1)
            result(r, c) = arr(srcR, srcC)
        Next c
    Next r

    Normalize2DArray = result
End Function

' ----------------------------------------------------------------------------------------------
' [F] SliceArraySafe
'
' 功能说明      : 安全切片（带边界检查），从源二维数组中提取指定行/列范围，返回 1-based 新数组。
'               : rowStart/rowEnd/colStart/colEnd 允许省略（Missing/Empty/Null），省略时使用源数组边界。
'               : 若输入参数非法（如非数值字符串）或范围越界，则返回 EMPTY_VALUE。
' 参数          : arr - 源数组（必须为已分配二维数组）
'               : rowStart - 起始行（可选）
'               : rowEnd - 结束行（可选）
'               : colStart - 起始列（可选）
'               : colEnd - 结束列（可选）
' 返回          : Variant - 切片后的 1-based 数组；失败返回 EMPTY_VALUE
' Purpose       : Slice a 2D array safely with bounds validation, returning a 1-based copy.
' Contract      : Core / Query-only.
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SliceArraySafe(ByVal arr As Variant, _
                              Optional ByVal rowStart As Variant, _
                              Optional ByVal rowEnd As Variant, _
                              Optional ByVal colStart As Variant, _
                              Optional ByVal colEnd As Variant) As Variant

    ' 源数组必须是已分配 2D 且有数据
    If Not IsArrayValid(arr, requireData:=True, require2D:=True) Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    Dim info As ArrayInfo
    If Not GetArrayInfo(arr, info) Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    ' ---- 参数解析（不抛异常）
    Dim ok As Boolean
    Dim rs As Long, re As Long, cs As Long, ce As Long

    rs = CoerceLongOrDefault(rowStart, info.LBound1, ok)
    If Not ok Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    re = CoerceLongOrDefault(rowEnd, info.UBound1, ok)
    If Not ok Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    cs = CoerceLongOrDefault(colStart, info.LBound2, ok)
    If Not ok Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    ce = CoerceLongOrDefault(colEnd, info.UBound2, ok)
    If Not ok Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    ' ---- 边界检查
    If rs < info.LBound1 Or rs > info.UBound1 Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If re < info.LBound1 Or re > info.UBound1 Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If rs > re Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If cs < info.LBound2 Or cs > info.UBound2 Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If ce < info.LBound2 Or ce > info.UBound2 Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If cs > ce Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    ' ---- 目标尺寸
    Dim rowCount As Long, colCount As Long
    rowCount = re - rs + 1
    colCount = ce - cs + 1

    ' ---- 安全尺寸（避免巨大数组导致崩溃）
    Dim safeSize As Variant
    safeSize = CalculateSafeArraySize(rowCount, colCount)
    If IsEmpty(safeSize) Then
        SliceArraySafe = EMPTY_VALUE
        Exit Function
    End If

    If safeSize(1) < rowCount Then rowCount = safeSize(1)
    If safeSize(2) < colCount Then colCount = safeSize(2)

    ' 调整实际结束位置（静默截断）
    re = rs + rowCount - 1
    ce = cs + colCount - 1

    ' ---- 复制
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To colCount)

    Dim r As Long, c As Long
    Dim srcR As Long, srcC As Long

    For r = 1 To rowCount
        srcR = rs + r - 1
        For c = 1 To colCount
            srcC = cs + c - 1
            result(r, c) = arr(srcR, srcC)
        Next c
    Next r

    SliceArraySafe = result

End Function

' ----------------------------------------------------------------------------------------------
' [f] CoerceLongOrDefault
'
' 功能说明      : 将 Variant 参数安全解析为 Long；若为 Missing/Empty/Null 则使用默认值；
'               : 若非数值则返回 ok=False（不抛异常）
' 参数          : v - 输入 Variant（可能 Missing/Empty/Null/String/Number）
'               : defVal - 缺省时使用的默认值
'               : ok - ByRef 输出，True=成功解析/采用默认值，False=非法输入
' 返回          : Long - 解析结果或默认值
' ----------------------------------------------------------------------------------------------
Private Function CoerceLongOrDefault(ByVal v As Variant, ByVal defVal As Long, ByRef ok As Boolean) As Long

    ok = True

    If IsMissing(v) Or IsEmpty(v) Or IsNull(v) Then
        CoerceLongOrDefault = defVal
        Exit Function
    End If

    If IsNumeric(v) Then
        ' CLng 在此处不会因类型不匹配崩溃；若是极端字符串数值如 "10" 也可正常处理
        CoerceLongOrDefault = CLng(v)
        Exit Function
    End If

    ok = False

End Function

' ----------------------------------------------------------------------------------------------
' [F] SliceArraySafeFull
'
' 功能说明      : 返回源数组的完整切片副本（等价于 SliceArraySafe(arr) 省略所有范围参数）。
'               : - 仅接受已分配二维数组
'               : - 返回 1-based 新数组（完整复制）
'               : - 若源数组无效则返回 EMPTY_VALUE
'               : - 若源数组过大，仍会遵循 SliceArraySafe 的安全截断策略（silent clamp）
' 参数          : arr - 源数组（必须为已分配二维数组）
' 返回          : Variant - 完整副本（1-based 2D）；失败返回 EMPTY_VALUE
' Purpose       : Convenience wrapper for full-range safe slicing (returns a 1-based copy).
' Contract      : Core / Query-only.
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SliceArraySafeFull(ByVal arr As Variant) As Variant
    SliceArraySafeFull = SliceArraySafe(arr:=arr)
End Function

' ----------------------------------------------------------------------------------------------
' [F] AppendArrayVertical
'
' 功能说明      : 垂直拼接两个数组（一维数组被视为单列），返回新的二维数组，自动进行索引规范化和大小检查
' 参数          : baseArray - 基础数组（位于上方）
'               : appendArray - 要追加的数组（位于下方）
'               : errMsg - 可选，返回错误信息
' 返回          : Variant - 垂直拼接后的二维数组，列数取两者列数（一维为1列）
' Purpose       : Vertically concatenates two arrays (1D treated as single column), returns new 2D array with auto-normalization and size checking
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function AppendArrayVertical(ByRef baseArray As Variant, _
                                   ByRef appendArray As Variant, _
                                   Optional ByRef errMsg As String) As Variant
    errMsg = vbNullString

    Dim bInfo As ArrayInfo, aInfo As ArrayInfo
    Dim bOk As Boolean, aOk As Boolean

    bOk = GetArrayInfo(baseArray, bInfo, errMsg)
    aOk = GetArrayInfo(appendArray, aInfo, errMsg)

    ' If either invalid => return normalized other (pass-through if other invalid too)
    If Not bOk Then
        AppendArrayVertical = NormalizeTo1Based(appendArray, False, errMsg)
        Exit Function
    End If
    If Not aOk Then
        AppendArrayVertical = NormalizeTo1Based(baseArray, False, errMsg)
        Exit Function
    End If

    ' Determine logical columns (1D => 1 col; 2D => Count2)
    Dim bCols As Long, aCols As Long
    If bInfo.Dims = 1 Then
        bCols = 1
    ElseIf bInfo.Dims = 2 Then
        bCols = bInfo.Count2
    Else
        bCols = 0
    End If

    If aInfo.Dims = 1 Then
        aCols = 1
    ElseIf aInfo.Dims = 2 Then
        aCols = aInfo.Count2
    Else
        aCols = 0
    End If

    If bInfo.Count1 <= 0 Or aInfo.Count1 <= 0 Or bCols <= 0 Or aCols <= 0 Then
        errMsg = "Allocated array with data required."
        AppendArrayVertical = Empty
        Exit Function
    End If

    If bCols <> aCols Then
        errMsg = "Column count mismatch (Base:" & bCols & " vs App:" & aCols & ")."
        AppendArrayVertical = Empty
        Exit Function
    End If

    Dim outRows As Long, outCols As Long
    outRows = bInfo.Count1 + aInfo.Count1
    outCols = bCols

    Dim safeSize() As Long
    safeSize = CalculateSafeArraySize(outRows, outCols)
    If safeSize(2) < outCols Then
        errMsg = "Unsafe output size."
        AppendArrayVertical = Empty
        Exit Function
    End If
    If safeSize(1) < outRows Then outRows = safeSize(1)

    Dim result() As Variant
    ReDim result(1 To outRows, 1 To outCols)

    Dim r As Long, c As Long
    Dim srcR As Long, srcC As Long

    ' Copy base (as much as fits)
    Dim baseRowsToCopy As Long
    baseRowsToCopy = bInfo.Count1
    If baseRowsToCopy > outRows Then baseRowsToCopy = outRows

    If bInfo.Dims = 1 Then
        For r = 1 To baseRowsToCopy
            srcR = bInfo.LBound1 + (r - 1)
            result(r, 1) = baseArray(srcR)
        Next r
    Else
        For r = 1 To baseRowsToCopy
            srcR = bInfo.LBound1 + (r - 1)
            For c = 1 To outCols
                srcC = bInfo.LBound2 + (c - 1)
                result(r, c) = baseArray(srcR, srcC)
            Next c
        Next r
    End If

    ' Copy append after base
    Dim remaining As Long
    remaining = outRows - baseRowsToCopy
    If remaining > 0 Then
        Dim appRowsToCopy As Long
        appRowsToCopy = aInfo.Count1
        If appRowsToCopy > remaining Then appRowsToCopy = remaining

        Dim outStart As Long
        outStart = baseRowsToCopy + 1

        If aInfo.Dims = 1 Then
            For r = 1 To appRowsToCopy
                srcR = aInfo.LBound1 + (r - 1)
                result(outStart + (r - 1), 1) = appendArray(srcR)
            Next r
        Else
            For r = 1 To appRowsToCopy
                srcR = aInfo.LBound1 + (r - 1)
                For c = 1 To outCols
                    srcC = aInfo.LBound2 + (c - 1)
                    result(outStart + (r - 1), c) = appendArray(srcR, srcC)
                Next c
            Next r
        End If
    End If

    AppendArrayVertical = result
End Function

' ==============================================================================================
' SECTION 03: CONVERT VALIDATE
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ToSafeLong
'
' 功能说明      : 将各种类型的输入值安全地转换为Long整数，支持范围检查和默认值
' 参数          : value - 要转换的输入值
'               : defaultValue - 可选，转换失败时返回的默认值，默认为0
'               : minVal - 可选，允许的最小值，默认为Long类型最小值(-2147483648)
'               : maxVal - 可选，允许的最大值，默认为Long类型最大值(2147483647)
'               : errMsg - 可选，返回错误信息
' 返回          : Long - 转换后的Long整数值，失败则返回defaultValue
' Purpose       : Safely converts various input types to Long integer with range checking and default value
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function ToSafeLong(ByVal value As Variant, _
                           Optional ByVal defaultValue As Long = 0, _
                           Optional ByVal minVal As Long = -2147483647 - 1, _
                           Optional ByVal maxVal As Long = 2147483647, _
                           Optional ByRef errMsg As String) As Long
    errMsg = vbNullString
    ToSafeLong = defaultValue

    If IsEmpty(value) Or IsNull(value) Then
        errMsg = "Empty/Null."
        Exit Function
    End If

    Dim vt As VbVarType
    vt = VarType(value)

    On Error GoTo Fail

    Dim d As Double
    Select Case vt
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            d = CDbl(value)

        Case vbBoolean
            If CBool(value) Then d = 1# Else d = 0#

        Case vbDate
            d = CDbl(CDate(value))

        Case vbString
            If Not IsNumeric(value) Then
                errMsg = "Not numeric string."
                Exit Function
            End If
            d = CDbl(value)

        Case Else
            If Not IsNumeric(value) Then
                errMsg = "Not numeric."
                Exit Function
            End If
            d = CDbl(value)
    End Select

    If d < CDbl(minVal) Or d > CDbl(maxVal) Then
        errMsg = "Out of range."
        Exit Function
    End If

    Dim l As Long
    l = CLng(d) ' may overflow on banker's rounding edge
    If l < minVal Or l > maxVal Then
        errMsg = "Out of range after rounding."
        Exit Function
    End If

    ToSafeLong = l
    Exit Function

Fail:
    errMsg = Err.Description
    ToSafeLong = defaultValue
End Function

' ----------------------------------------------------------------------------------------------
' [F] ToSafeDouble
'
' 功能说明      : 将各种类型的输入值安全地转换为Double浮点数，支持范围检查和默认值
' 参数          : value - 要转换的输入值
'               : defaultValue - 可选，转换失败时返回的默认值，默认为0
'               : minVal - 可选，允许的最小值，默认为Double类型最小值(-1.7E+308)
'               : maxVal - 可选，允许的最大值，默认为Double类型最大值(1.7E+308)
'               : errMsg - 可选，返回错误信息
' 返回          : Double - 转换后的Double浮点数值，失败则返回defaultValue
' Purpose       : Safely converts various input types to Double with range checking and default value
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function ToSafeDouble(ByVal value As Variant, _
                             Optional ByVal defaultValue As Double = 0#, _
                             Optional ByVal minVal As Double = -1.7E+308, _
                             Optional ByVal maxVal As Double = 1.7E+308, _
                             Optional ByRef errMsg As String) As Double
    errMsg = vbNullString
    ToSafeDouble = defaultValue

    If IsEmpty(value) Or IsNull(value) Then
        errMsg = "Empty/Null."
        Exit Function
    End If

    Dim vt As VbVarType
    vt = VarType(value)

    On Error GoTo Fail

    Dim d As Double
    Select Case vt
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            d = CDbl(value)

        Case vbBoolean
            If CBool(value) Then d = 1# Else d = 0#

        Case vbDate
            d = CDbl(CDate(value))

        Case vbString
            If Not IsNumeric(value) Then
                errMsg = "Not numeric string."
                Exit Function
            End If
            d = CDbl(value)

        Case Else
            If Not IsNumeric(value) Then
                errMsg = "Not numeric."
                Exit Function
            End If
            d = CDbl(value)
    End Select

    If d < minVal Or d > maxVal Then
        errMsg = "Out of range."
        Exit Function
    End If

    ToSafeDouble = d
    Exit Function

Fail:
    errMsg = Err.Description
    ToSafeDouble = defaultValue
End Function

' ----------------------------------------------------------------------------------------------
' [F] ToSafeString
'
' 功能说明      : 将各种类型的输入值安全地转换为字符串，支持修剪空格和长度限制
' 参数          : value - 要转换的输入值
'               : defaultValue - 可选，转换失败时返回的默认值，默认为空字符串
'               : trimWhitespace - 可选，是否修剪前后空格，默认为False
'               : maxLength - 可选，最大字符串长度，0表示不限制，默认为0
'               : errMsg - 可选，返回错误信息
' 返回          : String - 转换后的字符串，失败则返回defaultValue
' Purpose       : Safely converts various input types to String with trimming and length limiting options
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function ToSafeString(ByVal value As Variant, _
                             Optional ByVal defaultValue As String = vbNullString, _
                             Optional ByVal trimWhitespace As Boolean = False, _
                             Optional ByVal maxLength As Long = 0, _
                             Optional ByRef errMsg As String) As String
    errMsg = vbNullString
    ToSafeString = defaultValue

    If IsNull(value) Or IsEmpty(value) Then
        errMsg = "Null/Empty."
        Exit Function
    End If

    On Error GoTo Fail

    Dim s As String
    s = CStr(value)

    If trimWhitespace Then s = Trim$(s)
    If maxLength > 0 Then
        If Len(s) > maxLength Then s = Left$(s, maxLength)
    End If

    ToSafeString = s
    Exit Function

Fail:
    errMsg = Err.Description
    ToSafeString = defaultValue
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsNumericSafe
'
' 功能说明      : 安全地检查输入值是否为数值类型，自动处理Empty和Null值
' 参数          : value - 要检查的输入值
' 返回          : Boolean - 是否为数值，True表示是数值，False表示非数值或Empty/Null
' Purpose       : Safely checks if input value is numeric, automatically handles Empty and Null
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsNumericSafe(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    On Error Resume Next
    IsNumericSafe = IsNumeric(value)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsDateSafe
'
' 功能说明      : 安全地检查输入值是否为有效的日期，自动处理Empty和Null值
' 参数          : value - 要检查的输入值
' 返回          : Boolean - 是否为有效日期，True表示是日期，False表示非日期或Empty/Null
' Purpose       : Safely checks if input value is a valid date, automatically handles Empty and Null
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsDateSafe(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    On Error Resume Next
    IsDateSafe = IsDate(value)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] SanitizeString
'
' 功能说明      : 清理字符串中的无效字符，使用默认的清理函数
' 参数          : text - 要清理的原始文本
'               : replacement - 可选，用于替换无效字符的字符串，默认为空字符串
' 返回          : String - 清理后的安全字符串
' Purpose       : Sanitizes a string by removing invalid characters, using the default sanitization function
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SanitizeString(ByVal text As String, _
                               Optional ByVal replacement As String = vbNullString) As String
    SanitizeString = SanitizeStringFallback(text, replacement)
End Function

' ----------------------------------------------------------------------------------------------
' [f] SanitizeStringFallback
'
' 功能说明      : 清理字符串中的控制字符（ASCII码<32和127），将其替换为指定的替换字符串
' 参数          : text - 要清理的原始文本
'               : replacement - 用于替换无效字符的字符串
' 返回          : String - 清理后的安全字符串，所有控制字符被替换
' ----------------------------------------------------------------------------------------------
Private Function SanitizeStringFallback(ByVal text As String, _
                                        ByVal replacement As String) As String
    Dim n As Long
    n = Len(text)
    If n = 0 Then
        SanitizeStringFallback = vbNullString
        Exit Function
    End If

    Dim repLen As Long
    repLen = Len(replacement)

    Dim i As Long, code As Long
    Dim ch As String

    If repLen <= 1 Then
        Dim buf As String
        buf = String$(n, vbNullChar)

        For i = 1 To n
            ch = Mid$(text, i, 1)
            code = AscW(ch)

            If code >= 32 And code <> 127 Then
                Mid$(buf, i, 1) = ch
            Else
                If repLen = 1 Then
                    Mid$(buf, i, 1) = replacement
                End If
            End If
        Next i

        If repLen = 0 Then
            SanitizeStringFallback = Replace$(buf, vbNullChar, vbNullString)
        Else
            SanitizeStringFallback = buf
        End If

        Exit Function
    End If

    Dim parts() As String
    ReDim parts(1 To n)

    For i = 1 To n
        ch = Mid$(text, i, 1)
        code = AscW(ch)

        If code >= 32 And code <> 127 Then
            parts(i) = ch
        Else
            parts(i) = replacement
        End If
    Next i

    SanitizeStringFallback = Join(parts, vbNullString)
End Function

' ----------------------------------------------------------------------------------------------
' [f] ParseBoolean
'
' 功能说明      : 将各种类型的输入值解析为布尔值，支持数字、字符串和布尔类型
' 参数          : raw - 要解析的原始输入值
'               : outVal - 输出参数，返回解析后的布尔值
' 返回          : Boolean - 解析是否成功，True表示成功解析为布尔值
' ----------------------------------------------------------------------------------------------
Private Function ParseBoolean(ByVal raw As Variant, _
                              ByRef outVal As Boolean) As Boolean
    Dim vt As VbVarType
    vt = VarType(raw)

    Select Case vt
        Case vbBoolean
            outVal = CBool(raw)
            ParseBoolean = True

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            outVal = (CDbl(raw) <> 0#)
            ParseBoolean = True

        Case vbString
            Dim s As String
            s = Trim$(CStr(raw))
            If Len(s) = 0 Then Exit Function

            s = LCase$(s)
            If s = "true" Or s = "yes" Or s = "y" Or s = "1" Then
                outVal = True
                ParseBoolean = True
            ElseIf s = "false" Or s = "no" Or s = "n" Or s = "0" Then
                outVal = False
                ParseBoolean = True
            End If

        Case Else
            If IsNumeric(raw) Then
                outVal = (CDbl(raw) <> 0#)
                ParseBoolean = True
            End If
    End Select
End Function

' ==============================================================================================
' SECTION 04: SAFE MATH
' ==============================================================================================

Private Const DEFAULT_MAX_ABS As Double = 1.79769313486231E+308 ' Double.Max (approx)

' ----------------------------------------------------------------------------------------------
' [F] SafeMultiply

' 功能说明      : 安全地执行两个双精度数的乘法，检查结果是否为有限数且不超过最大绝对值
' 参数          : a - 第一个乘数
'               : b - 第二个乘数
'               : maxAbs - 可选，允许的最大绝对值，默认为DEFAULT_MAX_ABS
' 返回          : Variant - 乘法结果，若结果无效或超出范围则返回Empty
' Purpose       : Safely multiplies two double values, checking if result is finite and within maximum absolute value
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SafeMultiply(ByVal a As Double, ByVal b As Double, _
                             Optional ByVal maxAbs As Double = DEFAULT_MAX_ABS) As Variant
    Dim r As Double
    r = a * b

    If Not IsFiniteDouble(r) Then
        SafeMultiply = Empty
        Exit Function
    End If

    If Abs(r) > maxAbs Then
        SafeMultiply = Empty
    Else
        SafeMultiply = r
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] SafeAdd

' 功能说明      : 安全地执行两个双精度数的加法，检查结果是否为有限数且不超过最大绝对值
' 参数          : a - 第一个加数
'               : b - 第二个加数
'               : maxAbs - 可选，允许的最大绝对值，默认为DEFAULT_MAX_ABS
' 返回          : Variant - 加法结果，若结果无效或超出范围则返回Empty
' Purpose       : Safely adds two double values, checking if result is finite and within maximum absolute value
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SafeAdd(ByVal a As Double, ByVal b As Double, _
                        Optional ByVal maxAbs As Double = DEFAULT_MAX_ABS) As Variant
    Dim r As Double
    r = a + b

    If Not IsFiniteDouble(r) Then
        SafeAdd = Empty
        Exit Function
    End If

    If Abs(r) > maxAbs Then
        SafeAdd = Empty
    Else
        SafeAdd = r
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] IsFiniteDouble

' 功能说明      : 判断双精度数是否为有限数（非无穷大且非NaN）
' 参数          : x - 要检查的双精度数
' 返回          : Boolean - 是否为有限数，True表示是有限数
' ----------------------------------------------------------------------------------------------
Private Function IsFiniteDouble(ByVal x As Double) As Boolean
    ' Only finite values satisfy (x - x = 0).
    ' Inf - Inf = NaN, NaN - NaN = NaN, and NaN <> NaN.
    Dim diff As Double
    diff = x - x
    IsFiniteDouble = (diff = diff) ' False only for NaN
End Function
