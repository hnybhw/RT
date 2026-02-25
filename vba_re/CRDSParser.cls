VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRDSParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : CRDSParser
' PURPOSE         : 解析RDS格式的万年模拟数据，实现IDataParser接口
'                 : Parse RDS format 10K data, implements IDataParser interface
' DEPENDS         : IDataParser
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 常量与私有成员变量 / Constants & Private Members
'   [C] 列索引常量           - 输出数组列索引 / Output array column indices
'   [V] m_searchTerms        - 搜索关键词列表 / Search terms list
'
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
'   [S] Class_Initialize     - 初始化搜索关键词 / Initialize search terms
'
' SECTION 3: 接口实现 / Interface Implementation
'   [F] IDataParser_CanParse    - 判断是否能处理RDS分段 / Check if can parse RDS
'   [F] IDataParser_GetHeaders  - 返回RDS格式的表头 / Return RDS headers
'   [F] IDataParser_ParseRow    - 解析RDS格式的单行数据 / Parse single RDS row
'
' ==============================================================================================

Option Explicit
Implements IDataParser

' ==============================================================================================
' SECTION 1: 常量与私有成员变量 / Constants & Private Members
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 列索引常量 - 输出数组列索引 / Output array column indices
' ----------------------------------------------------------------------------------------------
Private Const COL_SUBLOB As Long = 4
Private Const COL_PERIL As Long = 5
Private Const COL_EVENT As Long = 6
Private Const COL_GROUP As Long = 9

' ----------------------------------------------------------------------------------------------
' [V] m_searchTerms - 用于分组判断的搜索关键词列表 / Search terms for group determination
' ----------------------------------------------------------------------------------------------
Private m_searchTerms As Variant

' ==============================================================================================
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Class_Initialize
' 说明：初始化搜索关键词数组
'       Initialize search terms array
' ----------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    m_searchTerms = Array("Marine", "PVT", "Offshore")
    Call WriteLog("CRDSParser", "Class_Initialize", "RDS解析器初始化完成 / RDS parser initialized", "解析器")
End Sub

' ==============================================================================================
' SECTION 3: 接口实现 / Interface Implementation
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_CanParse
' 说明：判断是否能处理RDS分段
'       Determine if can parse RDS segment
' 参数：segment - String，分段名称 / Segment name
' 返回值：Boolean - True=可以处理 / Can handle, False=不能处理 / Cannot handle
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_CanParse(ByVal segment As String) As Boolean
    IDataParser_CanParse = (segment = "RDS")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_GetHeaders
' 说明：返回RDS格式的表头数组
'       Return RDS format headers array
' 返回值：Variant - 表头数组 / Headers array
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_GetHeaders() As Variant
    IDataParser_GetHeaders = Array("LossSeq", "Column", "Year", "SubLoB", "Peril", "Event", "Subject_Loss", "TID", "Group")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_ParseRow
' 说明：解析RDS格式的单行数据，处理Peril、Event和Group字段
'       Parse single RDS row, handle Peril, Event and Group fields
' 参数：matrix - Variant，源数据矩阵 / Source data matrix
'       r - Long，当前行号 / Current row
'       c - Long，当前列号 / Current column
'       startRow - Long，数据起始行 / Data start row
'       segment - String，分段名称 / Segment name
' 返回值：Variant - 包含9列数据的数组 / Array with 9 columns of data
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_ParseRow(ByRef matrix As Variant, ByVal r As Long, ByVal c As Long, _
                                      ByVal startRow As Long, ByVal segment As String) As Variant
    ' 确保返回9列数据
    Dim result(1 To 9) As Variant
    
    ' 初始化所有列
    Dim i As Long
    For i = 1 To 9
        result(i) = ""
    Next i
    
    ' 安全检查
    On Error Resume Next
    
    ' 确保矩阵有足够的行和列
    If UBound(matrix, 1) >= 1 And UBound(matrix, 2) >= c Then
        result(4) = matrix(1, c)
    End If
    
    If UBound(matrix, 1) >= 2 And UBound(matrix, 2) >= c Then
        result(5) = matrix(2, c)
    End If
    
    If UBound(matrix, 1) >= 3 And UBound(matrix, 2) >= c Then
        result(6) = matrix(3, c)
    End If
    
    ' Group逻辑
    If Not IsEmpty(result(4)) Then
        Dim sl As String
        sl = CStr(result(4))
        
        If InStr(1, sl, "Aviation", vbTextCompare) > 0 Then
            result(9) = "Aviation"
        ElseIf segment = "RDS" Then
            Dim t As Variant
            For Each t In m_searchTerms
                If InStr(1, sl, CStr(t), vbTextCompare) > 0 Then
                    result(9) = "Marine"
                    Exit For
                End If
            Next t
        End If
    End If
    
    On Error GoTo 0
    
    IDataParser_ParseRow = result
End Function

