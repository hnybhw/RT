Attribute VB_Name = "plat_support"
' ==============================================================================================
' MODULE NAME       : plat_support
' LAYER             : Platform
' PURPOSE           : Provides standalone diagnostic and support utilities for development,
'                     inspection, and data-validation source generation.
'                     Not part of the core runtime lifecycle.
' DEPENDS           : Excel Object Model (ThisWorkbook.Names, Worksheets)
' NOTE              : - All functions in this module are standalone and do not require
'                       plat_context to be initialized.
'                     - Functions marked [UDF] are callable from Excel worksheet formulas.
'                     - No business logic allowed.
' STATUS            : Draft
' ==============================================================================================
' VERSION HISTORY   :
' v1.0.0
'   - Init (Module Baseline): Established plat_support as Platform-layer home for diagnostic
'     and development-support utilities, separate from runtime lifecycle modules.
'   - Add (Named Range Extraction): Introduced ExtractNamedRangesToSheet (Sub) and
'     GetNamedRangeList (UDF) for workbook Name enumeration with system-name filtering.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 01: NAMED RANGE DIAGNOSTICS
'   [S] ExtractNamedRangesToSheet  - Extract user-defined Named Ranges to Log sheet
'   [F] GetNamedRangeList          - [UDF] Return Named Range name array for INDEX() reference
'
' SECTION 02: INTERNAL HELPERS
'   [f] IsSystemName               - Detect Excel auto-generated system names
'   [f] GetLocalName               - Strip SheetName! prefix from a Name string
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function, [T]=Type
'       [UDF] = Callable from Excel worksheet formula (must be Public Function)
'       Rule: Helper functions and private procedures inherit the Contract and
'             Side Effects of their parent public API unless explicitly stated otherwise.
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 00: MODULE CONSTANTS
' ==============================================================================================

Private Const DIAG_OUTPUT_SHEET As String = "Log@SYS"
Private Const DIAG_COL_NAME     As Long = 1
Private Const DIAG_COL_REF      As Long = 2

' ==============================================================================================
' SECTION 01: NAMED RANGE DIAGNOSTICS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] ExtractNamedRangesToSheet
'
' 功能说明      : 提取当前工作簿中所有用户定义的命名区域（过滤 Excel 自动生成的系统名称），
'               : 将名称和引用位置写入诊断工作表，供数据有效性列表等场景使用
' 参数          : 无
' 返回          : 无
' Purpose       : Enumerates user-defined Names in ThisWorkbook, filters system-generated
'               : names, and writes Name / RefersTo pairs to the diagnostic output sheet
' Contract      : Platform / Diagnostic
'               : Output sheet must exist; aborts with message if not found
'               : Individual Name objects that fail to resolve are skipped with a warning entry
'               : System names (Print_Area, Print_Titles, _FilterDatabase) are excluded
' Side Effects  : Clears content (not format) of output sheet before writing
'               : Writes Name and RefersTo columns starting at row 2 (row 1 = headers)
' ----------------------------------------------------------------------------------------------
Public Sub ExtractNamedRangesToSheet()

    Dim wks       As Worksheet
    Dim nm        As Name
    Dim outputRow As Long
    Dim nameCount As Long
    Dim refText   As String

    ' ---- Bind output sheet (Fail Fast) ----
    On Error Resume Next
    Set wks = ThisWorkbook.Sheets(DIAG_OUTPUT_SHEET)
    On Error GoTo 0

    If wks Is Nothing Then
        MsgBox "诊断工作表 [" & DIAG_OUTPUT_SHEET & "] 不存在，操作已中止。", _
               vbCritical, "ExtractNamedRangesToSheet"
        Exit Sub
    End If

    ' ---- Clear content only (preserve format and logger anchor) ----
    wks.UsedRange.ClearContents

    ' ---- Write headers ----
    wks.Cells(1, DIAG_COL_NAME) = "名称"
    wks.Cells(1, DIAG_COL_REF) = "引用位置"

    outputRow = 2
    nameCount = 0

    ' ---- Enumerate all Names, filter system names ----
    For Each nm In ThisWorkbook.Names

        ' Skip Excel auto-generated system names
        If IsSystemName(nm.Name) Then GoTo NextName

        ' Isolate per-item read failure (corrupt Name objects, broken external links)
        refText = vbNullString
        On Error Resume Next
        refText = nm.RefersTo
        On Error GoTo 0

        wks.Cells(outputRow, DIAG_COL_NAME) = nm.Name

        If Len(refText) = 0 Then
            ' Mark unresolvable entries explicitly rather than silently skipping
            wks.Cells(outputRow, DIAG_COL_REF) = "[无法读取引用位置]"
        Else
            ' Prefix with space to prevent Excel from interpreting "=" as a formula
            wks.Cells(outputRow, DIAG_COL_REF) = " " & refText
        End If

        outputRow = outputRow + 1
        nameCount = nameCount + 1

NextName:
    Next nm

    MsgBox "命名区域提取完成，共提取 " & nameCount & " 个。", _
           vbInformation, "ExtractNamedRangesToSheet"

End Sub

' ----------------------------------------------------------------------------------------------
' [F] GetNamedRangeList
'
' 功能说明      : 返回工作簿中所有用户定义命名区域的名称数组，供 Excel 公式通过
'               : INDEX() 逐行引用，可用于在 Setup 表生成数据有效性的动态来源
' 参数          : 无
' 返回          : Variant - 一维数组，每个元素为一个命名区域的名称字符串
'               : 若无用户定义名称，返回包含单个空字符串的数组（避免公式错误）
' Purpose       : UDF entry point for worksheet formula access to workbook Name list
'               : Primary use case: data validation source in Setup sheet via INDEX()
' Contract      : Platform / Query-only
'               : Does not require plat_context to be initialized
'               : System names filtered by the same rule as ExtractNamedRangesToSheet
'               : Caller uses INDEX(GetNamedRangeList(), ROW(A1)) to pull items row by row
' Side Effects  : None (Query-only)
' ----------------------------------------------------------------------------------------------
Public Function GetNamedRangeList() As Variant

    Dim nm          As Name
    Dim nameCount   As Long
    Dim result()    As String
    Dim i           As Long

    ' ---- First pass: count valid names to size the array ----
    nameCount = 0
    For Each nm In ThisWorkbook.Names
        If Not IsSystemName(nm.Name) Then
            nameCount = nameCount + 1
        End If
    Next nm

    ' ---- Guard: return single-element array to prevent INDEX() formula error ----
    If nameCount = 0 Then
        GetNamedRangeList = Array("")
        Exit Function
    End If

    ' ---- Second pass: populate result array ----
    ReDim result(1 To nameCount)
    i = 1
    For Each nm In ThisWorkbook.Names
        If Not IsSystemName(nm.Name) Then
            result(i) = nm.Name
            i = i + 1
        End If
    Next nm

    GetNamedRangeList = result

End Function

' ==============================================================================================
' SECTION 02: INTERNAL HELPERS
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [f] IsSystemName
'
' 功能说明      : 判断给定名称是否为 Excel 自动生成的系统名称
'               : 当前过滤规则：Print_Area / Print_Titles / _FilterDatabase
' 参数          : fullName - Name 对象的完整名称字符串（可含 SheetName! 前缀）
' 返回          : Boolean - True 表示应跳过该名称
' Purpose       : Centralizes system-name detection logic to keep enumeration loops clean
' Contract      : Platform / Internal / Query-only
' Side Effects  : None
' ----------------------------------------------------------------------------------------------
Private Function IsSystemName(ByVal fullName As String) As Boolean

    Dim localName As String
    localName = GetLocalName(fullName)

    Select Case True
        Case localName Like "_FilterDatabase": IsSystemName = True
        Case localName Like "Print_Area": IsSystemName = True
        Case localName Like "Print_Titles": IsSystemName = True
        Case Else: IsSystemName = False
    End Select

End Function

' ----------------------------------------------------------------------------------------------
' [f] GetLocalName
'
' 功能说明      : 从名称字符串中去除 SheetName! 前缀，返回本地名称部分
'               : 例如："Setup!rng_config" → "rng_config"
'               : 例如："rng_sys_segment"  → "rng_sys_segment"（无前缀时原样返回）
' 参数          : fullName - Name 对象的完整名称字符串
' 返回          : String - 不含 SheetName! 前缀的本地名称
' Purpose       : Normalizes Name strings for pattern matching regardless of scope
' Contract      : Platform / Internal / Query-only
' Side Effects  : None
' ----------------------------------------------------------------------------------------------
Private Function GetLocalName(ByVal fullName As String) As String

    Dim bangPos As Long
    bangPos = InStr(fullName, "!")

    If bangPos > 0 Then
        GetLocalName = Mid$(fullName, bangPos + 1)
    Else
        GetLocalName = fullName
    End If

    MsgBox "命名区域提取完成，共提取 " & (outputRow - 2) & " 个。", vbInformation, "完成"
End Function
