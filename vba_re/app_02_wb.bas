Attribute VB_Name = "app_02_wb"
' ==============================================================================================
' MODULE NAME     : app_02_wb
' PURPOSE         : 工作簿高级操作核心模块，提供工作簿打开/校验、文件路径处理、SharePoint链接验证、文件选择对话框等跨工作簿操作能力
'                 : Core module for advanced workbook operations: open/validate workbooks, file path handling,
'                 : SharePoint URL validation, file dialog selection, and cross-workbook capabilities
' DEPENDS         : app_01_basic (HandleError, TrackWorkbook, WriteLog, GetNameFromURL等基础工具函数/过程)
'                 : app_01_basic (HandleError, TrackWorkbook, WriteLog, GetNameFromURL, etc.)
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 模块常量声明 / Module Constants
'   [C] 模块名称常量           - 定义当前模块名称，用于日志记录 / Module name constant for logging
'
' SECTION 2: 工作簿核心操作 / Core Workbook Operations
'   [F] OpenWorkbook           - 打开工作簿，支持只读模式，优先获取已打开工作簿避免重复打开
'                              / Open workbook with read-only mode, reuse if already open
'   [F] GetOpenWorkbookByName  - 根据工作簿名称检查并返回已打开的工作簿对象
'                              / Get open workbook by name
'
' SECTION 3: 路径与文件校验 / Path & File Validation
'   [F] ValidateFilePath       - 验证文件路径有效性，兼容本地路径和SharePoint网络路径
'                              / Validate file path (local or SharePoint)
'   [F] IsSharePointURLValid   - 验证SharePoint URL的有效性，通过尝试打开方式校验
'                              / Validate SharePoint URL by attempting to open
'
' SECTION 4: 文件路径选择与解析 / File Path Selection & Parsing
'   [F] GetFilePath            - 弹出文件选择对话框，获取用户选择的文件完整路径
'                              / Show file dialog and return selected path
'   [F] GetNameFromURL         - 从本地路径/网络URL中提取文件名，支持URL解码
'                              / Extract filename from path/URL with URL decoding
'   [F] GetFileNameWithoutExt  - 从文件路径中提取无扩展名的纯文件名
'                              / Extract filename without extension
'
' SECTION 5: URL工具函数 / URL Utilities
'   [F] DecodeURL              - URL解码，还原URL中编码的特殊字符（如%20=空格）
'                              / Decode URL encoded characters
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
Public Const MODULE_NAME As String = "app_02_wb"

' ==============================================================================================
' SECTION 2: 工作簿核心操作 / Core Workbook Operations
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] OpenWorkbook
' 说明：打开指定路径的工作簿，支持只读模式，优先检测已打开工作簿避免重复打开，打开后自动跟踪工作簿
'       Open workbook at specified path, check if already open, track for cleanup
' 参数：filePath - String，工作簿完整路径（本地/SharePoint） / Full workbook path (local/SharePoint)
'       readOnly - Boolean（可选），是否只读打开，默认True / Open as read-only, default True
' 返回值：Workbook - 打开的工作簿对象，失败则返回Nothing / Opened workbook object, Nothing if failed
' ----------------------------------------------------------------------------------------------
Public Function OpenWorkbook(ByVal filePath As String, Optional ByVal readOnly As Boolean = True) As Workbook
    On Error GoTo ErrorHandler
    
    ' 空路径直接返回空 / Return Nothing for empty path
    If Len(Trim(filePath)) = 0 Then
        Call WriteLog(MODULE_NAME, "OpenWorkbook", "空文件路径 / Empty file path", "警告")
        Set OpenWorkbook = Nothing
        Exit Function
    End If
    
    Dim fileName As String
    ' 从路径中提取纯文件名 / Extract filename from path
    fileName = GetNameFromURL(filePath)
    
    ' 先检查是否已打开，避免重复打开 / Check if already open
    Dim wb As Workbook
    Set wb = GetOpenWorkbookByName(fileName)
    
    ' 未打开则执行新打开操作 / Open if not already open
    If wb Is Nothing Then
        ' 关闭Excel提示和事件，提升打开稳定性 / Disable alerts and events for stability
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        
        ' 打开工作簿，关闭链接更新 / Open workbook with links update disabled
        Set wb = Workbooks.Open(filePath, UpdateLinks:=False, readOnly:=readOnly)
        
        ' 恢复Excel提示和事件 / Restore alerts and events
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        
        ' 跟踪外部工作簿，便于后续自动清理 / Track for automatic cleanup
        TrackWorkbook wb
        
        Call WriteLog(MODULE_NAME, "OpenWorkbook", "打开新工作簿 / Opened new workbook: " & fileName, "工作簿操作")
    Else
        Call WriteLog(MODULE_NAME, "OpenWorkbook", "复用已打开工作簿 / Reusing open workbook: " & fileName, "工作簿操作")
    End If
    
    Set OpenWorkbook = wb
    Exit Function

ErrorHandler:
    ' 异常时恢复Excel环境 / Restore Excel environment on error
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    HandleError MODULE_NAME & ".OpenWorkbook", Err.Description
    Set OpenWorkbook = Nothing
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetOpenWorkbookByName
' 说明：根据工作簿名称检查是否已在Excel中打开，返回对应工作簿对象
'       Check if workbook is already open by name
' 参数：fileName - String，工作簿名称（含扩展名） / Workbook name (with extension)
' 返回值：Workbook - 已打开的工作簿对象，未打开则返回Nothing / Open workbook object, Nothing if not open
' ----------------------------------------------------------------------------------------------
Public Function GetOpenWorkbookByName(ByVal fileName As String) As Workbook
    Dim wb As Workbook
    
    ' 屏蔽错误，避免工作簿未打开时触发运行时错误 / Suppress error if workbook not open
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    
    Set GetOpenWorkbookByName = wb
    
    If Not wb Is Nothing Then
        Call WriteLog(MODULE_NAME, "GetOpenWorkbookByName", "找到已打开工作簿 / Found open workbook: " & fileName, "工作簿操作")
    End If
End Function

' ==============================================================================================
' SECTION 3: 路径与文件校验 / Path & File Validation
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] ValidateFilePath
' 说明：验证文件路径的有效性，自动区分本地路径和SharePoint网络路径，分别采用不同校验方式
'       Validate file path, automatically handle local vs SharePoint paths
' 参数：filePath - String，待验证的文件路径/SharePoint URL / File path or SharePoint URL to validate
' 返回值：Boolean - True=路径有效，False=路径无效/不存在 / True if valid, False otherwise
' ----------------------------------------------------------------------------------------------
Public Function ValidateFilePath(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 空路径直接判定为无效 / Empty path is invalid
    If filePath = "" Then
        ValidateFilePath = False
        Exit Function
    End If
    
    ' 包含https:则判定为SharePoint路径，否则为本地路径 / SharePoint or local path
    If InStr(1, filePath, "https:", vbTextCompare) > 0 Then
        ValidateFilePath = IsSharePointURLValid(filePath)
        Call WriteLog(MODULE_NAME, "ValidateFilePath", "SharePoint路径校验 / SharePoint path validation: " & filePath & " => " & ValidateFilePath, "路径校验")
    Else
        ValidateFilePath = (Dir(filePath) <> "")
        Call WriteLog(MODULE_NAME, "ValidateFilePath", "本地路径校验 / Local path validation: " & filePath & " => " & ValidateFilePath, "路径校验")
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".ValidateFilePath", Err.Description
    ValidateFilePath = False
End Function

' ----------------------------------------------------------------------------------------------
' [F] IsSharePointURLValid
' 说明：验证SharePoint URL的有效性，通过尝试只读打开的方式进行实际校验
'       Validate SharePoint URL by attempting read-only open
' 参数：sharePointURL - String，待验证的SharePoint工作簿URL / SharePoint workbook URL to validate
' 返回值：Boolean - True=URL有效可访问，False=URL无效/无法访问 / True if accessible, False otherwise
' ----------------------------------------------------------------------------------------------
Public Function IsSharePointURLValid(ByVal sharePointURL As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 非http/https开头直接判定为无效 / Not HTTP/HTTPS is invalid
    If Not (sharePointURL Like "https://*" Or sharePointURL Like "http://*") Then
        IsSharePointURLValid = False
        Exit Function
    End If
    
    ' 关闭Excel环境提示/事件/刷新，提升校验速度 / Disable UI elements for performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wbTmp As Workbook
    ' 尝试只读打开SharePoint工作簿，验证可访问性 / Attempt to open as read-only
    Set wbTmp = Workbooks.Open(sharePointURL, UpdateLinks:=False, readOnly:=True)
    
    ' 打开成功则判定为有效，立即关闭临时工作簿 / If open succeeds, URL is valid
    If Not wbTmp Is Nothing Then
        IsSharePointURLValid = True
        wbTmp.Close SaveChanges:=False
        Call WriteLog(MODULE_NAME, "IsSharePointURLValid", "SharePoint URL有效 / Valid SharePoint URL: " & sharePointURL, "路径校验")
    End If

CleanExit:
    ' 无论是否成功，均恢复Excel原始环境 / Restore Excel environment
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    IsSharePointURLValid = False
    Resume CleanExit
End Function

' ==============================================================================================
' SECTION 4: 文件路径选择与解析 / File Path Selection & Parsing
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetFilePath
' 说明：弹出Excel文件选择对话框，支持自定义标题和文件筛选，获取用户选择的单个文件完整路径
'       Show Excel file dialog with custom title and filter, return selected path
' 参数：title - String，对话框标题 / Dialog title
'       filterDescription - String，文件筛选规则（格式：描述,扩展名，如Excel文件,.xlsx;.xlsm）
'                          / File filter (format: description,extension e.g. "Excel Files,.xlsx;.xlsm")
' 返回值：String - 用户选择的文件完整路径，取消选择则返回空字符串 / Selected file path, empty if cancelled
' ----------------------------------------------------------------------------------------------
Public Function GetFilePath(ByVal title As String, ByVal filterDescription As String) As String
    On Error GoTo ErrorHandler
    
    ' 晚绑定创建文件对话框对象 / Late-bound file dialog object
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(3) ' 3=msoFileDialogFilePicker
    
    With fileDialog
        .title = title
        .AllowMultiSelect = False ' 禁止多选，仅支持单个文件选择 / Single file selection only
        
        ' 清空默认筛选规则，添加自定义筛选 / Clear default filters, add custom filter
        .Filters.Clear
        
        Dim filterParts() As String
        filterParts = Split(filterDescription, ",")
        If UBound(filterParts) >= 1 Then
            .Filters.Add Trim(filterParts(0)), Trim(filterParts(1))
        End If
        
        ' 用户选择文件则返回路径，否则返回空 / Return selected path or empty if cancelled
        If .Show = -1 Then
            GetFilePath = .SelectedItems(1)
            Call WriteLog(MODULE_NAME, "GetFilePath", "用户选择文件 / User selected: " & GetFilePath, "文件选择")
        Else
            GetFilePath = ""
            Call WriteLog(MODULE_NAME, "GetFilePath", "用户取消选择 / User cancelled selection", "文件选择")
        End If
    End With
    
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME & ".GetFilePath", Err.Description
    GetFilePath = ""
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetNameFromURL
' 说明：从本地文件路径或网络URL中提取纯文件名（含扩展名），网络URL会先进行URL解码
'       Extract filename (with extension) from local path or URL, decode URL if needed
' 参数：filePath - String，本地路径/网络URL / Local path or URL
' 返回值：String - 提取的文件名（含扩展名），空路径则返回空 / Filename with extension, empty if path empty
' ----------------------------------------------------------------------------------------------
Public Function GetNameFromURL(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    ' 空路径直接返回空 / Return empty for empty path
    If Len(Trim(filePath)) = 0 Then
        GetNameFromURL = ""
        Exit Function
    End If
    
    ' 取最后一个/或\的位置，截取后续字符串为文件名 / Find last slash or backslash
    Dim lastPos As Long
    lastPos = Application.WorksheetFunction.Max(InStrRev(filePath, "/"), InStrRev(filePath, "\"))
    
    Dim result As String
    If lastPos > 0 Then
        result = Mid(filePath, lastPos + 1)
    Else
        result = filePath
    End If
    
    ' 包含http则判定为URL，进行URL解码后返回 / Decode if URL
    If InStr(1, filePath, "http", vbTextCompare) > 0 Then
        result = DecodeURL(result)
    End If
    
    GetNameFromURL = result
    Exit Function

ErrorHandler:
    GetNameFromURL = ""
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetFileNameWithoutExt
' 说明：从文件路径中提取无扩展名的纯文件名，先提取含扩展名文件名再截取扩展名前部分
'       Extract filename without extension from path
' 参数：filePath - String，本地路径/网络URL / Local path or URL
' 返回值：String - 无扩展名的纯文件名，空路径则返回空 / Filename without extension, empty if path empty
' ----------------------------------------------------------------------------------------------
Public Function GetFileNameWithoutExt(ByVal filePath As String) As String
    ' 先提取含扩展名的文件名 / Get filename with extension first
    Dim fileName As String
    fileName = GetNameFromURL(filePath)
    
    ' 取最后一个.的位置，截取前面部分为无扩展名文件名 / Find last dot and truncate
    Dim lastDot As Long
    lastDot = InStrRev(fileName, ".")
    
    Dim result As String
    If lastDot > 1 Then
        result = Left(fileName, lastDot - 1)
    Else
        ' 无扩展名则直接返回原文件名 / Return as-is if no extension
        result = fileName
    End If
    
    GetFileNameWithoutExt = result
End Function

' ==============================================================================================
' SECTION 5: URL工具函数 / URL Utilities
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] DecodeURL
' 说明：URL解码核心函数，还原URL中被编码的常见特殊字符，适配SharePoint路径解析
'       Decode URL encoded characters for SharePoint path parsing
' 参数：inputURL - String，待解码的URL字符串 / URL string to decode
' 返回值：String - 解码后的原始字符串 / Decoded string
' ----------------------------------------------------------------------------------------------
Public Function DecodeURL(ByVal inputURL As String) As String
    Dim res As String
    res = inputURL
    
    ' 替换常见的URL编码字符，按常用程度排序 / Replace common URL encoded characters
    res = Replace(res, "%20", " ")   ' 空格 / Space
    res = Replace(res, "%5B", "[")   ' 左方括号 / Left bracket
    res = Replace(res, "%5D", "]")   ' 右方括号 / Right bracket
    res = Replace(res, "%2F", "/")   ' 斜杠 / Forward slash
    res = Replace(res, "%2E", ".")   ' 点 / Dot
    res = Replace(res, "%2C", ",")   ' 逗号 / Comma
    res = Replace(res, "%2B", "+")   ' 加号 / Plus
    res = Replace(res, "%23", "#")   ' 井号 / Hash
    res = Replace(res, "%40", "@")   ' At符号 / At sign
    
    DecodeURL = res
End Function
