VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : IDataParser
' PURPOSE         : 定义数据解析器的标准接口，所有具体解析器必须实现此接口
'                 : Define standard interface for data parsers, all concrete parsers must implement
' DEPENDS         : 无 / None
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 接口方法定义 / Interface Method Definitions
'   [F] CanParse    - 判断解析器是否能处理指定的分段 / Check if parser can handle segment
'   [F] ParseRow    - 解析单行数据，返回格式化的结果数组 / Parse single row, return formatted array
'   [F] GetHeaders  - 返回输出表的表头数组 / Return output table headers
'
' ==============================================================================================
' NOTE: 此为接口类，所有方法均为抽象方法，不包含实现代码
'       This is an interface class, all methods are abstract with no implementation
' ==============================================================================================

Option Explicit

' ----------------------------------------------------------------------------------------------
' [F] CanParse
' 说明：判断解析器是否能处理指定的分段
'       Determine if parser can handle the specified segment
' 参数：segment - String，分段名称 / Segment name
' 返回值：Boolean - True=可以处理 / Can handle, False=不能处理 / Cannot handle
' ----------------------------------------------------------------------------------------------
Public Function CanParse(ByVal segment As String) As Boolean
End Function

' ----------------------------------------------------------------------------------------------
' [F] ParseRow
' 说明：解析单行数据，返回格式化的结果数组
'       Parse single row of data, return formatted result array
' 参数：matrix - Variant，源数据矩阵 / Source data matrix
'       r - Long，当前行号 / Current row
'       c - Long，当前列号 / Current column
'       startRow - Long，数据起始行 / Data start row
'       segment - String，分段名称 / Segment name
' 返回值：Variant - 包含9列数据的数组 / Array with 9 columns of data
' ----------------------------------------------------------------------------------------------
Public Function ParseRow(ByRef matrix As Variant, ByVal r As Long, ByVal c As Long, _
                         ByVal startRow As Long, ByVal segment As String) As Variant
End Function

' ----------------------------------------------------------------------------------------------
' [F] GetHeaders
' 说明：返回输出表的表头数组
'       Return output table headers array
' 返回值：Variant - 表头数组 / Headers array
' ----------------------------------------------------------------------------------------------
Public Function GetHeaders() As Variant
End Function
