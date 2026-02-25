VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMMParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : CMMParser
' PURPOSE         : 解析MM_Large和MM_Attritional格式的万年模拟数据，实现IDataParser接口
'                 : Parse MM_Large and MM_Attritional format 10K data, implements IDataParser
' DEPENDS         : IDataParser
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 常量定义 / Constants
'   [C] 列索引常量           - 输出数组列索引 / Output array column indices
'
' SECTION 2: 接口实现 / Interface Implementation
'   [F] IDataParser_CanParse    - 判断是否能处理MM分段 / Check if can parse MM segments
'   [F] IDataParser_GetHeaders  - 返回MM格式的表头 / Return MM headers
'   [F] IDataParser_ParseRow    - 解析MM格式的单行数据 / Parse single MM row
'
' ==============================================================================================

Option Explicit
Implements IDataParser

' ==============================================================================================
' SECTION 1: 常量定义 / Constants
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [C] 列索引常量 - 输出数组列索引 / Output array column indices
' ----------------------------------------------------------------------------------------------
Private Const COL_SUBLOB As Long = 4
Private Const COL_PERIL As Long = 5

' ==============================================================================================
' SECTION 2: 接口实现 / Interface Implementation
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_CanParse
' 说明：判断是否能处理MM分段（所有以NonCat_Large开头的分段以及NonCat_Attritional）
'       Determine if can parse MM segments (all segments starting with NonCat_Large and NonCat_Attritional)
' 参数：segment - String，分段名称 / Segment name
' 返回值：Boolean - True=可以处理 / Can handle, False=不能处理 / Cannot handle
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_CanParse(ByVal segment As String) As Boolean
    IDataParser_CanParse = (segment Like "NonCat_Large*" Or segment = "NonCat_Attritional")
    
    If IDataParser_CanParse Then
        Call WriteLog("CMMParser", "IDataParser_CanParse", "可以处理分段 / Can handle segment: " & segment, "解析器")
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_GetHeaders
' 说明：返回MM格式的表头数组
'       Return MM format headers array
' 返回值：Variant - 表头数组 / Headers array
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_GetHeaders() As Variant
    IDataParser_GetHeaders = Array("LossSeq", "Column", "Year", "SubLoB", "Peril", "Event", "Subject_Loss", "TID", "Group")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_ParseRow
' 说明：解析MM格式的单行数据，根据分段名称设置Peril字段
'       Parse single MM row, set Peril field based on segment name
' 参数：matrix - Variant，源数据矩阵 / Source data matrix
'       r - Long，当前行号 / Current row
'       c - Long，当前列号 / Current column
'       startRow - Long，数据起始行 / Data start row
'       segment - String，分段名称 / Segment name
' 返回值：Variant - 包含9列数据的数组 / Array with 9 columns of data
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_ParseRow(ByRef matrix As Variant, ByVal r As Long, ByVal c As Long, _
                                      ByVal startRow As Long, ByVal segment As String) As Variant
    Dim result(1 To 9) As Variant
    
    ' SubLoB (从第1行获取) / SubLoB from row 1
    result(COL_SUBLOB) = matrix(1, c)
    
    ' Peril逻辑 - 根据分段名称设置 / Peril logic based on segment name
    If segment Like "NonCat_Large*" Then
        result(COL_PERIL) = "MM_Large"
    ElseIf segment = "NonCat_Attritional" Then
        result(COL_PERIL) = "MM_Att"
    End If
    
    IDataParser_ParseRow = result
End Function

