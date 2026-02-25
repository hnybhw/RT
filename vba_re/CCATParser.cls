VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCATParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : CCATParser
' PURPOSE         : 解析CAT格式的万年模拟数据，实现IDataParser接口
'                 : Parse CAT format 10K data, implements IDataParser interface
' DEPENDS         : IDataParser
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 常量定义 / Constants
'   [C] 列索引常量           - 输出数组列索引 / Output array column indices
'
' SECTION 2: 接口实现 / Interface Implementation
'   [F] IDataParser_CanParse    - 判断是否能处理CAT分段 / Check if can parse CAT segments
'   [F] IDataParser_GetHeaders  - 返回CAT格式的表头 / Return CAT headers
'   [F] IDataParser_ParseRow    - 解析CAT格式的单行数据 / Parse single CAT row
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
Private Const COL_EVENT As Long = 6
Private Const COL_GROUP As Long = 9

' ==============================================================================================
' SECTION 2: 接口实现 / Interface Implementation
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_CanParse
' 说明：判断是否能处理CAT分段（所有以CAT_开头的分段）
'       Determine if can parse CAT segments (all segments starting with CAT_)
' 参数：segment - String，分段名称 / Segment name
' 返回值：Boolean - True=可以处理 / Can handle, False=不能处理 / Cannot handle
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_CanParse(ByVal segment As String) As Boolean
    IDataParser_CanParse = (segment Like "CAT_*")
    
    If IDataParser_CanParse Then
        Call WriteLog("CCATParser", "IDataParser_CanParse", "可以处理分段 / Can handle segment: " & segment, "解析器")
    End If
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_GetHeaders
' 说明：返回CAT格式的表头数组
'       Return CAT format headers array
' 返回值：Variant - 表头数组 / Headers array
' ----------------------------------------------------------------------------------------------
Private Function IDataParser_GetHeaders() As Variant
    IDataParser_GetHeaders = Array("LossSeq", "Column", "Year", "SubLoB", "Peril", "Event", "Subject_Loss", "TID", "Group")
End Function

' ----------------------------------------------------------------------------------------------
' [F] IDataParser_ParseRow
' 说明：解析CAT格式的单行数据，设置Peril、Event和Group字段
'       Parse single CAT row, set Peril, Event and Group fields
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
    
    ' Peril (从第2行获取) / Peril from row 2
    result(COL_PERIL) = matrix(2, c)
    
    ' Event (从第3行获取) / Event from row 3
    result(COL_EVENT) = matrix(3, c)
    
    ' Group - CAT类型统一设为"CAT" / Group set to "CAT" for all CAT segments
    result(COL_GROUP) = "CAT"
    
    IDataParser_ParseRow = result
End Function

