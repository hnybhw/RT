Attribute VB_Name = "core_files"
' ==============================================================================================
' MODULE NAME       : core_files
' LAYER             : core
' PURPOSE           : Query-only Excel object queries (worksheet/range/name) plus pure path/URL
'                     parsing and validation helpers (string-only; no IO).
' DEPENDS           : Excel Object Model (Workbook/Worksheet/Range/Name)
' NOTE              : - No IO: does not open workbooks, does not touch filesystem, no dialogs.
'                     - No Application state changes (no Find-state pollution, no ScreenUpdating).
'                     - Path/URL logic is purely string-based; “accessibility checks” belong to Platform.
' STATUS            : Frozen
' ==============================================================================================
' VERSION HISTORY   :
' v2.0.0
'   - Refactor: split project into layered architecture (Core / Platform / Business).
'   - Freeze: query-only Excel helpers + pure path/URL parsing/shape-validation.
'   - Merge: previous path/url sections consolidated into SECTION 02.
'
' v1.0.0
'   - Initial draft based on legacy implementation, iteratively refined during early refactor.
' ==============================================================================================
' TABLE OF CONTENTS :
'
' SECTION 01: EXCEL WORKSHEET & RANGE & NAME
'   [F] IsWorksheetValid        - Validates worksheet object, optional visible check
'   [F] IsWorksheetExist        - Checks if worksheet exists in workbook (optional outWs)
'   [F] GetLastRowSafely        - Safely gets last row for a target column (filter-mode aware)
'   [F] GetLastColSafely        - Safely gets last col for a target row
'   [F] TryGetNamedRange        - Tries sheet-level then workbook-level named range
'   [F] GetActualDataRangeCore  - Gets CurrentRegion around an anchor address
'   [f] ResolveRowIndex         - Parses/validates row index
'   [f] ResolveColumnIndex      - Parses/validates column index (number or letters)
'   [f] ColumnLettersToNumber   - Converts column letters (A..XFD) to number
'
' SECTION 02: PATH & URL (PURE)
'   [F] NormalizePath           - Normalizes file path or URL (slashes, duplicates, trailing)
'   [F] SafePathCombine         - Combines base + relative part safely (absolute passthrough)
'   [F] IsNetworkPath           - UNC detection (\\)
'   [F] IsSharePointPath        - SharePoint shape detection (URL/UNC)
'   [F] IsSharePointURLValid    - SharePoint URL shape validation (no open/IO)
'   [F] GetNameFromURL          - Extracts last segment (supports URL decode + strip query)
'   [F] GetFileNameWithoutExt   - Extracts name without extension
'   [F] DecodeURL               - Lightweight %XX decoder (+ -> space)
'   [f] IsAbsolutePath          - Absolute path detection (drive/UNC/URL)
'   [f] GetParentFolderPath     - Parent folder for path/URL
'   [f] IsHttpUrl               - Checks http/https
'   [f] NormalizeUrlSlashes     - Normalizes URL slashes
'   [f] NormalizeRelativeUrlPiece - Normalizes relative URL path part
'   [f] CollapseBackslashesPreserveUNC - Collapses backslashes preserving UNC prefix
'   [f] StripUrlQueryFragment   - Removes ?query and #fragment
'   [f] GetUrlHost              - Extracts host from URL
'   [f] ContainsInvalidPathChars - Windows invalid char / control char check
'   [f] HexPairToLong           - Hex pair to byte (0..255)
'   [f] HexCharToVal            - Hex char to nibble (0..15)
' ==============================================================================================
' NOTE: [C]=Constant, [V]=Variable, [P]=Property, [S]=Public Sub, [s]=Private Sub,
'       [F]=Public Function, [f]=Private Function
' ==============================================================================================
Option Explicit

' ==============================================================================================
' SECTION 01: EXCEL WORKSHEET & RANGE & NAME
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IsWorksheetValid
'
' 功能说明      : 验证工作表对象是否有效，可选项检查工作表是否可见
' 参数          : ws - 要验证的工作表对象
'               : checkVisible - 可选，是否检查工作表可见性，默认为False
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 工作表是否有效，True表示有效（且如果要求可见，则可见）
' Purpose       : Validates if a worksheet object is valid, optionally checking if the worksheet is visible
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsWorksheetValid(ByVal ws As Worksheet, _
                                 Optional ByVal checkVisible As Boolean = False, _
                                 Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    If ws Is Nothing Then
        errMsg = "Worksheet is Nothing."
        Exit Function
    End If

    On Error GoTo Fail

    Dim tmp As String
    tmp = ws.Name

    If checkVisible Then
        IsWorksheetValid = (ws.Visible = xlSheetVisible)
        If Not IsWorksheetValid Then errMsg = "Worksheet not visible."
    Else
        IsWorksheetValid = True
    End If

    Exit Function

Fail:
    errMsg = Err.Description
    IsWorksheetValid = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsWorksheetExist
'
' 功能说明      : 检查工作簿中是否存在指定名称的工作表（仅 Worksheets 集合，不包含 Chart sheets），并可选返回工作表对象
' 参数          : wb - 要检查的工作簿对象
'               : sheetName - 要查找的工作表名称（会自动 Trim）
'               : outWs - 可选，输出参数，返回找到的工作表对象（未找到则为 Nothing）
'               : errMsg - 可选，返回错误信息（成功时为空字符串）
' 返回          : Boolean - 工作表是否存在，True表示存在
' Purpose       : Checks if a worksheet exists in a workbook (Worksheets only), optionally returns the worksheet object.
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsWorksheetExist(ByVal wb As Workbook, _
                                 ByVal sheetName As String, _
                                 Optional ByRef outWs As Worksheet, _
                                 Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    Set outWs = Nothing
    IsWorksheetExist = False

    If wb Is Nothing Then
        errMsg = "Workbook is Nothing."
        Exit Function
    End If

    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then
        errMsg = "Sheet name is empty."
        Exit Function
    End If

    ' Query-only: avoid raising runtime errors for missing sheets
    On Error Resume Next
    Set outWs = wb.Worksheets(sheetName)
    On Error GoTo 0

    IsWorksheetExist = Not (outWs Is Nothing)
    If Not IsWorksheetExist Then
        errMsg = "Worksheet not found: " & sheetName
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetLastRowSafely
'
' 功能说明      : 安全地获取工作表中指定列的最后非空行号，处理筛选模式和空表情况
' 参数          : ws - 目标工作表对象
'               : targetCol - 可选，目标列（可以是列号或列字母），默认为第1列
'               : errMsg - 可选，返回错误信息
' 返回          : Long - 最后非空行号，0表示空表，-1表示工作表处于筛选模式
' Purpose       : Safely gets the last non-empty row number for a specified column in a worksheet, handles filter mode and empty sheet cases
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetLastRowSafely(ByVal ws As Worksheet, _
                                 Optional ByVal targetCol As Variant = 1, _
                                 Optional ByRef errMsg As String) As Long
    errMsg = vbNullString
    GetLastRowSafely = 0

    If ws Is Nothing Then
        errMsg = "Worksheet is Nothing."
        Exit Function
    End If

    Dim isFiltered As Boolean
    On Error Resume Next
    isFiltered = ws.FilterMode
    On Error GoTo 0

    If isFiltered Then
        GetLastRowSafely = -1
        errMsg = "Worksheet is in FilterMode."
        Exit Function
    End If

    Dim colNum As Long
    colNum = ResolveColumnIndex(ws, targetCol)
    If colNum <= 0 Then
        errMsg = "Invalid target column."
        Exit Function
    End If

    On Error GoTo Fail

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, colNum).End(xlUp).Row

    If lastRow = 1 Then
        Dim c As Range
        Set c = ws.Cells(1, colNum)
        If IsEmpty(c.Value2) And Len(c.Formula) = 0 Then lastRow = 0
    End If

    GetLastRowSafely = lastRow
    Exit Function

Fail:
    errMsg = Err.Description
    GetLastRowSafely = 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetLastColSafely
'
' 功能说明      : 安全地获取工作表中指定行的最后非空列号，处理空表情况
' 参数          : ws - 目标工作表对象
'               : targetRow - 可选，目标行号，默认为第1行
'               : errMsg - 可选，返回错误信息
' 返回          : Long - 最后非空列号，0表示该行全为空
' Purpose       : Safely gets the last non-empty column number for a specified row in a worksheet, handles empty row cases
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetLastColSafely(ByVal ws As Worksheet, _
                                 Optional ByVal targetRow As Variant = 1, _
                                 Optional ByRef errMsg As String) As Long
    errMsg = vbNullString
    GetLastColSafely = 0

    If ws Is Nothing Then
        errMsg = "Worksheet is Nothing."
        Exit Function
    End If

    Dim rowNum As Long
    rowNum = ResolveRowIndex(ws, targetRow)
    If rowNum <= 0 Then
        errMsg = "Invalid target row."
        Exit Function
    End If

    On Error GoTo Fail

    Dim lastCol As Long
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column

    If lastCol = 1 Then
        Dim c As Range
        Set c = ws.Cells(rowNum, 1)
        If IsEmpty(c.Value2) And Len(c.Formula) = 0 Then lastCol = 0
    End If

    GetLastColSafely = lastCol
    Exit Function

Fail:
    errMsg = Err.Description
    GetLastColSafely = 0
End Function

' ----------------------------------------------------------------------------------------------
' [F] TryGetNamedRange
'
' 功能说明      : 尝试从工作簿或工作表中获取指定名称的命名区域，支持工作表级和工作簿级名称
' 参数          : wbOrWs - 工作簿或工作表对象
'               : rangeName - 要查找的命名区域名称
'               : outRng - 输出参数，返回找到的区域对象
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 是否成功找到并获取命名区域，True表示成功
' Purpose       : Attempts to get a named range from a workbook or worksheet, supports both sheet-level and workbook-level names
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function TryGetNamedRange(ByVal wbOrWs As Object, _
                                 ByVal rangeName As String, _
                                 ByRef outRng As Range, _
                                 Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString
    Set outRng = Nothing
    TryGetNamedRange = False

    rangeName = Trim$(rangeName)
    If Len(rangeName) = 0 Then
        errMsg = "Name is empty."
        Exit Function
    End If

    If wbOrWs Is Nothing Then
        errMsg = "wbOrWs is Nothing."
        Exit Function
    End If

    On Error GoTo Fail

    Dim nm As Name

    If TypeOf wbOrWs Is Worksheet Then
        Dim ws As Worksheet
        Set ws = wbOrWs

        ' Sheet-level name
        On Error Resume Next
        Set nm = ws.Names(rangeName)
        On Error GoTo 0

        If Not nm Is Nothing Then
            On Error Resume Next
            Set outRng = nm.RefersToRange
            On Error GoTo 0
            If Not outRng Is Nothing Then
                TryGetNamedRange = True
                Exit Function
            End If
        End If

        ' Workbook-level name
        Dim wb As Workbook
        Set wb = ws.parent

        Set nm = Nothing
        On Error Resume Next
        Set nm = wb.Names(rangeName)
        On Error GoTo 0

        If Not nm Is Nothing Then
            On Error Resume Next
            Set outRng = nm.RefersToRange
            On Error GoTo 0
            If Not outRng Is Nothing Then
                TryGetNamedRange = True
                Exit Function
            End If
        End If

        errMsg = "Named range not found or not a range."
        Exit Function
    End If

    If TypeOf wbOrWs Is Workbook Then
        Dim wb2 As Workbook
        Set wb2 = wbOrWs

        On Error Resume Next
        Set nm = wb2.Names(rangeName)
        On Error GoTo 0

        If nm Is Nothing Then
            errMsg = "Named range not found."
            Exit Function
        End If

        On Error Resume Next
        Set outRng = nm.RefersToRange
        On Error GoTo 0

        If outRng Is Nothing Then
            errMsg = "Name does not refer to a range."
            Exit Function
        End If

        TryGetNamedRange = True
        Exit Function
    End If

    errMsg = "wbOrWs must be Worksheet or Workbook."
    Exit Function

Fail:
    errMsg = Err.Description
    Set outRng = Nothing
    TryGetNamedRange = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetActualDataRangeCore
'
' 功能说明      : 获取工作表中以指定地址为中心的实际数据区域（CurrentRegion），自动处理空单元格情况
' 参数          : ws - 目标工作表对象
'               : startAddress - 起始单元格地址（即使传入区域也只会使用第一个单元格）
'               : errMsg - 可选，返回错误信息
' 返回          : Range - 实际数据区域，若无数据或出错则返回Nothing
' Purpose       : Gets the actual data region (CurrentRegion) around a specified address in a worksheet, handles empty cell cases
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetActualDataRangeCore(ByVal ws As Worksheet, _
                                       ByVal startAddress As String, _
                                       Optional ByRef errMsg As String) As Range
    errMsg = vbNullString
    Set GetActualDataRangeCore = Nothing

    If ws Is Nothing Then
        errMsg = "Worksheet is Nothing."
        Exit Function
    End If

    startAddress = Trim$(startAddress)
    If Len(startAddress) = 0 Then
        errMsg = "startAddress is empty."
        Exit Function
    End If

    On Error GoTo Fail

    Dim anchor As Range
    ' Force single cell even if caller passes "A1:B2"
    Set anchor = ws.Range(startAddress).Cells(1, 1)

    Dim region As Range
    Set region = anchor.CurrentRegion
    If region Is Nothing Then Exit Function

    ' Valid if NOT (single-cell AND empty)
    If region.Cells.Count = 1 Then
        If IsEmpty(region.Value2) And Len(region.Formula) = 0 Then
            errMsg = "No data found around " & startAddress
            Exit Function
        End If
    End If

    Set GetActualDataRangeCore = region
    Exit Function

Fail:
    errMsg = Err.Description
    Set GetActualDataRangeCore = Nothing
End Function

' ----------------------------------------------------------------------------------------------
' [f] ResolveRowIndex
'
' 功能说明      : 将行标识解析为有效的行号，验证是否在工作表行范围内
' 参数          : ws - 目标工作表对象
'               : targetRow - 要解析的行标识（仅支持数值类型）
' 返回          : Long - 有效的行号，若无效则返回0
' ----------------------------------------------------------------------------------------------
Private Function ResolveRowIndex(ByVal ws As Worksheet, ByVal targetRow As Variant) As Long
    ResolveRowIndex = 0
    If IsNumeric(targetRow) Then
        ResolveRowIndex = CLng(targetRow)
        If ResolveRowIndex < 1 Or ResolveRowIndex > ws.rows.Count Then ResolveRowIndex = 0
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] ResolveColumnIndex
'
' 功能说明      : 将列标识（列号或列字母）解析为有效的列号，验证是否在工作表列范围内
' 参数          : ws - 目标工作表对象
'               : targetCol - 要解析的列标识（数值列号或字符串列字母）
' 返回          : Long - 有效的列号，若无效则返回0
' ----------------------------------------------------------------------------------------------
Private Function ResolveColumnIndex(ByVal ws As Worksheet, ByVal targetCol As Variant) As Long
    ResolveColumnIndex = 0

    If IsNumeric(targetCol) Then
        ResolveColumnIndex = CLng(targetCol)
    ElseIf VarType(targetCol) = vbString Then
        ResolveColumnIndex = ColumnLettersToNumber(CStr(targetCol))
    Else
        ResolveColumnIndex = 0
    End If

    If ResolveColumnIndex < 1 Or ResolveColumnIndex > ws.Columns.Count Then
        ResolveColumnIndex = 0
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] ColumnLettersToNumber
'
' 功能说明      : 将Excel列字母（如"A", "AB", "XFD"）转换为对应的列号
' 参数          : letters - 要转换的列字母字符串
' 返回          : Long - 对应的列号，若字母无效则返回0
' ----------------------------------------------------------------------------------------------
Private Function ColumnLettersToNumber(ByVal letters As String) As Long
    letters = UCase$(Trim$(letters))
    If Len(letters) = 0 Then Exit Function

    Dim i As Long, c As Integer, n As Long
    For i = 1 To Len(letters)
        c = AscW(Mid$(letters, i, 1))
        If c < 65 Or c > 90 Then
            ColumnLettersToNumber = 0
            Exit Function
        End If
        n = n * 26 + (c - 64)
    Next i

    ColumnLettersToNumber = n
End Function

' ==============================================================================================
' SECTION 02: PATH & URL (PURE)
' ==============================================================================================

Private Const SCHEME_HTTP As String = "http://"
Private Const SCHEME_HTTPS As String = "https://"

' ----------------------------------------------------------------------------------------------
' [F] NormalizePath
'
' 功能说明      : 规范化文件路径或URL，统一路径分隔符，可选保留尾部斜杠
' 参数          : filePath - 要规范化的原始路径字符串
'               : keepTrailingSlash - 可选，是否保留尾部斜杠，默认为False
' 返回          : String - 规范化后的路径（文件路径使用反斜杠，URL保留正斜杠）
' Purpose       : Normalizes file path or URL by standardizing separators, optionally preserving trailing slash
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function NormalizePath(ByVal filePath As String, _
                              Optional ByVal keepTrailingSlash As Boolean = False) As String
    Dim s As String
    s = Trim$(filePath)

    If Len(s) = 0 Then
        NormalizePath = vbNullString
        Exit Function
    End If

    Dim isUrl As Boolean
    isUrl = IsHttpUrl(s)

    If isUrl Then
        NormalizePath = NormalizeUrlSlashes(s)
        Exit Function
    End If

    ' File-system path
    s = Replace$(s, "/", "\")
    s = CollapseBackslashesPreserveUNC(s)

    If keepTrailingSlash Then
        If Right$(s, 1) <> "\" Then s = s & "\"
    Else
        ' Remove trailing "\" except drive root "C:\"
        If Len(s) > 3 Then
            If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
        End If
    End If

    NormalizePath = s
End Function

' ----------------------------------------------------------------------------------------------
' [F] SafePathCombine
'
' 功能说明      : 安全地组合两个路径部分，自动处理分隔符和规范化，支持文件路径和URL
' 参数          : path1 - 基础路径部分
'               : path2 - 要追加的路径部分
' 返回          : String - 组合后的完整规范化路径
' Purpose       : Safely combines two path parts with automatic separator handling and normalization, supports both file paths and URLs
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function SafePathCombine(ByVal path1 As String, ByVal path2 As String) As String
    path1 = Trim$(path1)
    path2 = Trim$(path2)

    If Len(path1) = 0 Then
        SafePathCombine = NormalizePath(path2)
        Exit Function
    End If

    If Len(path2) = 0 Then
        SafePathCombine = NormalizePath(path1)
        Exit Function
    End If

    If IsAbsolutePath(path2) Then
        SafePathCombine = NormalizePath(path2)
        Exit Function
    End If

    ' If path1 is URL, combine with "/" semantics
    If IsHttpUrl(path1) Then
        Dim u1 As String, u2 As String
        u1 = NormalizeUrlSlashes(path1)
        u2 = NormalizeRelativeUrlPiece(path2)

        If Right$(u1, 1) = "/" Then
            SafePathCombine = u1 & u2
        Else
            SafePathCombine = u1 & "/" & u2
        End If
        Exit Function
    End If

    ' File path combine
    path1 = NormalizePath(path1, False)
    path2 = NormalizePath(path2, False)

    If Right$(path1, 1) = "\" Then
        SafePathCombine = NormalizePath(path1 & path2, False)
    Else
        SafePathCombine = NormalizePath(path1 & "\" & path2, False)
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] IsAbsolutePath
'
' 功能说明      : 判断给定路径是否为绝对路径（包括URL、UNC路径、驱动器根路径和系统根路径）
' 参数          : path - 要检查的路径字符串
' 返回          : Boolean - 是否为绝对路径，True表示是绝对路径
' ----------------------------------------------------------------------------------------------
Private Function IsAbsolutePath(ByVal path As String) As Boolean
    Dim s As String
    s = Trim$(path)
    If Len(s) = 0 Then Exit Function

    If IsHttpUrl(s) Then
        IsAbsolutePath = True
        Exit Function
    End If

    s = Replace$(s, "/", "\") ' do not call NormalizePath here (avoid extra work)

    ' UNC
    If Left$(s, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If

    ' Drive letter: "C:" or "C:\"
    If Len(s) >= 2 Then
        If Mid$(s, 2, 1) = ":" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If

    ' Rooted path like "\Windows\..."
    If Left$(s, 1) = "\" Then
        IsAbsolutePath = True
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsNetworkPath
'
' 功能说明      : 判断给定路径是否为网络路径（UNC路径，以双反斜杠开头）
' 参数          : filePath - 要检查的路径字符串
' 返回          : Boolean - 是否为网络路径，True表示是UNC网络路径
' Purpose       : Determines if the given path is a network path (UNC path starting with double backslash)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsNetworkPath(ByVal filePath As String) As Boolean
    Dim s As String
    s = Trim$(filePath)
    If Len(s) = 0 Then Exit Function

    s = Replace$(s, "/", "\")
    IsNetworkPath = (Left$(s, 2) = "\\")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsSharePointPath
'
' 功能说明      : 判断给定路径是否为SharePoint路径（包括HTTP URL和UNC路径，特征包含"sharepoint"关键字）
' 参数          : filePath - 要检查的路径字符串
' 返回          : Boolean - 是否为SharePoint路径，True表示是SharePoint相关路径
' Purpose       : Determines if the given path is a SharePoint path (including HTTP URLs and UNC paths, characterized by "sharepoint" keyword)
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsSharePointPath(ByVal filePath As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(filePath))
    If Len(s) = 0 Then Exit Function

    If IsHttpUrl(s) Then
        ' Keep it simple and fast: SharePoint URLs almost always include "sharepoint"
        IsSharePointPath = (InStr(1, s, "sharepoint", vbTextCompare) > 0)
        Exit Function
    End If

    s = Replace$(s, "/", "\")
    If Left$(s, 2) = "\\" Then
        IsSharePointPath = (InStr(1, s, "sharepoint.com", vbTextCompare) > 0) Or _
                           (InStr(1, s, ".sharepoint.", vbTextCompare) > 0)
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetParentFolderPath
'
' 功能说明      : 获取给定路径的父文件夹路径，支持文件路径和URL，自动处理不同路径类型的边界情况
' 参数          : path - 要获取父文件夹的路径字符串
' 返回          : String - 父文件夹路径，若无父文件夹则返回空字符串
' ----------------------------------------------------------------------------------------------
Private Function GetParentFolderPath(ByVal path As String) As String
    Dim s As String
    s = Trim$(path)
    If Len(s) = 0 Then Exit Function

    If IsHttpUrl(s) Then
        s = NormalizeUrlSlashes(s)

        ' Remove query/fragment for parent determination (simple + fast)
        s = StripUrlQueryFragment(s)

        Dim p As Long
        p = InStrRev(s, "/")
        If p <= 0 Then Exit Function

        ' Keep "https://host" without trailing "/"
        If p <= InStr(s, "://") + 2 Then Exit Function
        GetParentFolderPath = Left$(s, p - 1)
        Exit Function
    End If

    s = NormalizePath(s, False)

    ' Drive root "C:\" has no parent
    If Len(s) = 3 And Mid$(s, 2, 1) = ":" And Right$(s, 1) = "\" Then Exit Function

    Dim lastSlash As Long
    lastSlash = InStrRev(s, "\")
    If lastSlash <= 0 Then
        ' "C:" -> treat as root "C:\"
        If Len(s) = 2 And Mid$(s, 2, 1) = ":" Then
            GetParentFolderPath = s & "\"
        End If
        Exit Function
    End If

    Dim parent As String
    parent = Left$(s, lastSlash - 1)

    ' "C:" => "C:\"
    If Len(parent) = 2 And Mid$(parent, 2, 1) = ":" Then parent = parent & "\"

    ' UNC: "\\server" has no parent; "\\server\share" parent is "\\server"
    If Left$(s, 2) = "\\" Then
        If parent = "\" Then parent = vbNullString
        ' If parent becomes "\\": means just UNC prefix; treat as ""
        If parent = "\\" Then parent = vbNullString
    End If

    GetParentFolderPath = parent
End Function

' ----------------------------------------------------------------------------------------------
' [F] ValidateFilePath
'
' 功能说明      : 验证文件路径或URL的格式是否有效，检查空路径、非法字符和必要的路径结构
' 参数          : filePath - 要验证的路径字符串
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - 路径格式是否有效，True表示路径格式正确
' Purpose       : Validates if a file path or URL has a valid format, checking for empty path, invalid characters, and required path structure
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function ValidateFilePath(ByVal filePath As String, _
                                 Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString

    Dim s As String
    s = Trim$(filePath)
    If Len(s) = 0 Then
        errMsg = "Empty path."
        Exit Function
    End If

    If IsHttpUrl(s) Then
        ValidateFilePath = IsSharePointURLValid(s, errMsg)
        Exit Function
    End If

    s = NormalizePath(s, False)

    ' Basic invalid characters for Windows file paths (excluding ":" allowed at drive position)
    ' < > " | ? *  and control chars
    If ContainsInvalidPathChars(s) Then
        errMsg = "Invalid characters in path."
        Exit Function
    End If

    ' Drive path like "C:\..."
    If Len(s) >= 2 And Mid$(s, 2, 1) = ":" Then
        ValidateFilePath = True
        Exit Function
    End If

    ' UNC path "\\server\share\..."
    If Left$(s, 2) = "\\" Then
        ' require at least "\\x\y"
        Dim p As Long
        p = InStr(3, s, "\")
        If p > 3 Then
            ValidateFilePath = True
        Else
            errMsg = "UNC path too short."
        End If
        Exit Function
    End If

    ' Relative path allowed? In Core, treat as valid format (caller decides)
    ValidateFilePath = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsSharePointURLValid
'
' 功能说明      : 验证SharePoint URL的格式是否有效，检查协议、主机名和SharePoint特征
' 参数          : sharePointUrl - 要验证的SharePoint URL字符串
'               : errMsg - 可选，返回错误信息
' 返回          : Boolean - SharePoint URL是否有效，True表示格式正确且包含SharePoint特征
' Purpose       : Validates if a SharePoint URL has a valid format, checking protocol, hostname, and SharePoint characteristics
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function IsSharePointURLValid(ByVal sharePointUrl As String, _
                                     Optional ByRef errMsg As String) As Boolean
    errMsg = vbNullString

    Dim s As String
    s = Trim$(sharePointUrl)
    If Len(s) = 0 Then
        errMsg = "Empty URL."
        Exit Function
    End If

    If Not IsHttpUrl(s) Then
        errMsg = "Not http/https."
        Exit Function
    End If

    s = NormalizeUrlSlashes(s)
    s = StripUrlQueryFragment(s)

    ' Extract host between "://" and next "/"
    Dim host As String
    host = GetUrlHost(s)
    If Len(host) = 0 Then
        errMsg = "Missing host."
        Exit Function
    End If

    If InStr(1, host, "sharepoint", vbTextCompare) = 0 Then
        errMsg = "Host does not look like SharePoint."
        Exit Function
    End If

    IsSharePointURLValid = True
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetNameFromURL
'
' 功能说明      : 从文件路径或URL中提取文件名或最后一部分名称，支持URL解码和参数清理
' 参数          : filePath - 要提取名称的路径字符串
' 返回          : String - 提取的文件名或URL最后一部分，对URL自动解码并去除查询参数
' Purpose       : Extracts the file name or last part from a file path or URL, supports URL decoding and query parameter cleanup
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetNameFromURL(ByVal filePath As String) As String
    Dim s As String
    s = Trim$(filePath)
    If Len(s) = 0 Then Exit Function

    Dim isUrl As Boolean
    isUrl = IsHttpUrl(s)

    Dim p1 As Long, p2 As Long, lastPos As Long
    p1 = InStrRev(s, "/")
    p2 = InStrRev(s, "\")
    If p1 > p2 Then lastPos = p1 Else lastPos = p2

    Dim namePart As String
    If lastPos > 0 Then
        namePart = Mid$(s, lastPos + 1)
    Else
        namePart = s
    End If

    ' Strip query/fragment if URL
    If isUrl Then
        namePart = StripUrlQueryFragment(namePart)
        namePart = DecodeURL(namePart)
    End If

    GetNameFromURL = namePart
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetFileNameWithoutExt
'
' 功能说明      : 从文件路径中提取不带扩展名的文件名
' 参数          : filePath - 要处理的文件路径字符串
' 返回          : String - 不带扩展名的文件名，若无扩展名则返回完整文件名
' Purpose       : Extracts the file name without extension from a file path
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function GetFileNameWithoutExt(ByVal filePath As String) As String
    Dim fn As String
    fn = GetNameFromURL(filePath)
    If Len(fn) = 0 Then Exit Function

    Dim dotPos As Long
    dotPos = InStrRev(fn, ".")
    If dotPos > 1 Then
        GetFileNameWithoutExt = Left$(fn, dotPos - 1)
    Else
        GetFileNameWithoutExt = fn
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] DecodeURL
'
' 功能说明      : 解码URL编码的字符串，将%XX形式的编码和加号(+)转换为原始字符
'               : 只保证常见的ASCII %20、%2F等，复杂UTF-8编码不保证完全正确（但也不会崩溃），适用于解码文件名等简单URL组件
' 参数          : inputUrl - 要解码的URL编码字符串
' 返回          : String - 解码后的原始字符串
' Purpose       : Decodes a URL-encoded string, converting %XX encoded sequences and plus signs to original characters
' Contract      : Core / Query-only
' Side Effects  : None (Query-only).
' ----------------------------------------------------------------------------------------------
Public Function DecodeURL(ByVal inputUrl As String) As String
    Dim s As String
    s = inputUrl
    If Len(s) = 0 Then
        DecodeURL = vbNullString
        Exit Function
    End If

    Dim n As Long
    n = Len(s)

    Dim i As Long
    Dim ch As String
    Dim out As String

    ' Pre-allocate output buffer roughly (string concatenation is ok for typical file name lengths)
    out = vbNullString

    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)

        If ch = "+" Then
            out = out & " "
            i = i + 1

        ElseIf ch = "%" And i + 2 <= n Then
            Dim h1 As String, h2 As String
            h1 = Mid$(s, i + 1, 1)
            h2 = Mid$(s, i + 2, 1)

            Dim v As Long
            v = HexPairToLong(h1, h2)
            If v >= 0 Then
                out = out & Chr$(v)
                i = i + 3
            Else
                out = out & ch
                i = i + 1
            End If

        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    DecodeURL = out
End Function

' ----------------------------------------------------------------------------------------------
' [f] IsHttpUrl
' 功能说明      : 判断字符串是否为 HTTP 或 HTTPS URL（内部自动 Trim 与小写化处理）
' 参数          : s - 待判断的字符串
' 返回          : Boolean - True 表示为 http:// 或 https:// 开头的 URL
' ----------------------------------------------------------------------------------------------
Private Function IsHttpUrl(ByVal s As String) As Boolean
    ' 防御性处理：自动去除首尾空格并统一为小写
    s = LCase$(Trim$(s))

    If Len(s) = 0 Then Exit Function

    IsHttpUrl = (Left$(s, 7) = "http://") _
             Or (Left$(s, 8) = "https://")

End Function

' ----------------------------------------------------------------------------------------------
' [f] NormalizeUrlSlashes
' 功能说明      : 规范化URL中的斜杠，将反斜杠转换为正斜杠，并折叠连续的斜杠
' 参数          : url - 要规范化的URL字符串
' 返回          : String - 斜杠规范化后的URL
' ----------------------------------------------------------------------------------------------
Private Function NormalizeUrlSlashes(ByVal url As String) As String
    Dim s As String
    s = Trim$(url)

    Dim protoPos As Long
    protoPos = InStr(1, s, "://", vbTextCompare)
    If protoPos = 0 Then
        NormalizeUrlSlashes = s
        Exit Function
    End If

    Dim head As String, tail As String
    head = Left$(s, protoPos + 2)     ' include "://"
    tail = Mid$(s, protoPos + 3)

    ' Convert "\" to "/"
    tail = Replace$(tail, "\", "/")

    ' Collapse multiple "/" in tail
    Do While InStr(tail, "//") > 0
        tail = Replace$(tail, "//", "/")
    Loop

    NormalizeUrlSlashes = head & tail
End Function

' ----------------------------------------------------------------------------------------------
' [f] NormalizeRelativeUrlPiece
' 功能说明      : 规范化URL的相对路径部分，将反斜杠转换为正斜杠，并去除开头的斜杠
' 参数          : piece - 要规范化的相对路径字符串
' 返回          : String - 规范化后的相对路径，无开头的斜杠
' ----------------------------------------------------------------------------------------------
Private Function NormalizeRelativeUrlPiece(ByVal piece As String) As String
    Dim s As String
    s = Trim$(piece)
    If Len(s) = 0 Then
        NormalizeRelativeUrlPiece = vbNullString
        Exit Function
    End If
    s = Replace$(s, "\", "/")
    Do While Left$(s, 1) = "/"
        s = Mid$(s, 2)
    Loop
    NormalizeRelativeUrlPiece = s
End Function

' ----------------------------------------------------------------------------------------------
' [f] CollapseBackslashesPreserveUNC
' 功能说明      : 折叠路径中的连续反斜杠为单个反斜杠，同时保留UNC路径的双反斜杠前缀
' 参数          : path - 要处理的路径字符串
' 返回          : String - 反斜杠折叠后的路径，UNC路径前缀保持双反斜杠
' ----------------------------------------------------------------------------------------------
Private Function CollapseBackslashesPreserveUNC(ByVal path As String) As String
    Dim s As String
    s = path

    Dim isUNC As Boolean
    isUNC = (Left$(s, 2) = "\\")

    If isUNC Then
        s = Mid$(s, 3)
    End If

    Do While InStr(s, "\\") > 0
        s = Replace$(s, "\\", "\")
    Loop

    If isUNC Then
        s = "\\" & s
    End If

    CollapseBackslashesPreserveUNC = s
End Function

' ----------------------------------------------------------------------------------------------
' [f] StripUrlQueryFragment
' 功能说明      : 去除URL中的查询参数（?后）和片段标识（#后），只保留基础路径部分
' 参数          : urlOrPart - 要处理的URL字符串
' 返回          : String - 去除查询参数和片段标识后的URL基础部分
' ----------------------------------------------------------------------------------------------
Private Function StripUrlQueryFragment(ByVal urlOrPart As String) As String
    Dim s As String
    s = urlOrPart

    Dim q As Long, f As Long, cut As Long
    q = InStr(1, s, "?", vbBinaryCompare)
    f = InStr(1, s, "#", vbBinaryCompare)

    If q = 0 And f = 0 Then
        StripUrlQueryFragment = s
        Exit Function
    End If

    If q = 0 Then
        cut = f
    ElseIf f = 0 Then
        cut = q
    ElseIf q < f Then
        cut = q
    Else
        cut = f
    End If

    StripUrlQueryFragment = Left$(s, cut - 1)
End Function

' ----------------------------------------------------------------------------------------------
' [f] GetUrlHost
' 功能说明      : 从URL中提取主机名（域名或IP地址部分）
' 参数          : url - 要提取主机名的URL字符串
' 返回          : String - URL中的主机名，若无协议或格式错误则返回空字符串
' ----------------------------------------------------------------------------------------------
Private Function GetUrlHost(ByVal url As String) As String
    Dim s As String
    s = url

    Dim protoPos As Long
    protoPos = InStr(1, s, "://", vbTextCompare)
    If protoPos = 0 Then Exit Function

    Dim startPos As Long
    startPos = protoPos + 3

    Dim slashPos As Long
    slashPos = InStr(startPos, s, "/", vbBinaryCompare)

    If slashPos = 0 Then
        GetUrlHost = Mid$(s, startPos)
    ElseIf slashPos > startPos Then
        GetUrlHost = Mid$(s, startPos, slashPos - startPos)
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [f] ContainsInvalidPathChars
' 功能说明      : 检查字符串是否包含Windows文件路径中的非法字符（控制字符及 " * < > ? |）
' 参数          : s - 要检查的字符串
' 返回          : Boolean - 是否包含非法字符，True表示包含
' ----------------------------------------------------------------------------------------------
Private Function ContainsInvalidPathChars(ByVal s As String) As Boolean
    ' Windows invalids: < > " | ? * and control chars (0-31)
    Dim i As Long, ch As Integer
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 32 Then
            ContainsInvalidPathChars = True
            Exit Function
        End If
        Select Case ch
            Case 34, 42, 60, 62, 63, 124 ' " * < > ? |
                ContainsInvalidPathChars = True
                Exit Function
        End Select
    Next i
End Function

' ----------------------------------------------------------------------------------------------
' [f] HexPairToLong
' 功能说明      : 将两个十六进制字符转换为对应的字节值（0-255）
' 参数          : h1 - 第一个十六进制字符（高位）
'               : h2 - 第二个十六进制字符（低位）
' 返回          : Long - 转换后的字节值，若字符无效则返回-1
' ----------------------------------------------------------------------------------------------
Private Function HexPairToLong(ByVal h1 As String, ByVal h2 As String) As Long
    Dim v1 As Long, v2 As Long
    v1 = HexCharToVal(h1)
    If v1 < 0 Then
        HexPairToLong = -1
        Exit Function
    End If
    v2 = HexCharToVal(h2)
    If v2 < 0 Then
        HexPairToLong = -1
        Exit Function
    End If
    HexPairToLong = (v1 * 16) + v2
End Function

' ----------------------------------------------------------------------------------------------
' [f] HexCharToVal
' 功能说明      : 将单个十六进制字符（0-9, A-F, a-f）转换为对应的数值（0-15）
' 参数          : h - 要转换的十六进制字符
' 返回          : Long - 转换后的数值，若字符无效则返回-1
' ----------------------------------------------------------------------------------------------
Private Function HexCharToVal(ByVal h As String) As Long
    Dim c As Integer
    c = AscW(h)

    ' 0-9
    If c >= 48 And c <= 57 Then
        HexCharToVal = c - 48
        Exit Function
    End If
    ' A-F
    If c >= 65 And c <= 70 Then
        HexCharToVal = c - 55
        Exit Function
    End If
    ' a-f
    If c >= 97 And c <= 102 Then
        HexCharToVal = c - 87
        Exit Function
    End If

    HexCharToVal = -1
End Function
