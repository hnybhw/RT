VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParserFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================================
' CLASS NAME      : CParserFactory
' PURPOSE         : 根据分段名称返回合适的解析器实例，实现解析器的集中管理和注册
'                 : Return appropriate parser instance based on segment name,
'                 : centralize parser management and registration
' DEPENDS         : IDataParser, CRDSParser, CMMParser, CCATParser
' ==============================================================================================
' TABLE OF CONTENTS:
'
' SECTION 1: 私有成员变量 / Private Member Variables
'   [V] m_parsers  - 存储所有已注册解析器的集合 / Collection of registered parsers
'
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
'   [S] Class_Initialize - 初始化时注册所有支持的解析器 / Register all supported parsers
'
' SECTION 3: 公共方法 / Public Methods
'   [F] GetParser  - 根据分段名称获取对应的解析器 / Get parser for segment
'
' ==============================================================================================
' NOTE: 新增解析器时，只需在 Class_Initialize 中注册，无需修改其他代码
'       To add new parser, just register in Class_Initialize, no other code changes needed
' ==============================================================================================

Option Explicit

' ==============================================================================================
' SECTION 1: 私有成员变量 / Private Member Variables
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [V] m_parsers - 存储所有已注册解析器的集合 / Collection of registered parsers
' ----------------------------------------------------------------------------------------------
Private m_parsers As Collection

' ==============================================================================================
' SECTION 2: 类生命周期管理 / Class Lifecycle Management
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [S] Class_Initialize
' 说明：类初始化时创建集合并注册所有支持的解析器
'       Create collection and register all supported parsers on initialization
' ----------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_parsers = New Collection
    
    ' 注册所有支持的解析器 / Register all supported parsers
    m_parsers.Add New CRDSParser
    m_parsers.Add New CMMParser
    m_parsers.Add New CCATParser
    
    ' 写入日志 / Write to log
    Call WriteLog("CParserFactory", "Class_Initialize", "解析器工厂初始化完成，已注册 " & m_parsers.count & " 个解析器 / Factory initialized, " & m_parsers.count & " parsers registered", "工厂模式")
End Sub

' ==============================================================================================
' SECTION 3: 公共方法 / Public Methods
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' [F] GetParser
' 说明：根据分段名称遍历已注册解析器，返回第一个能处理该分段的解析器
'       Iterate registered parsers, return first parser that can handle the segment
' 参数：segment - String，分段名称 / Segment name
' 返回值：IDataParser - 解析器对象，无匹配则返回 Nothing / Parser object, Nothing if no match
' ----------------------------------------------------------------------------------------------
Public Function GetParser(ByVal segment As String) As IDataParser
    Dim parser As IDataParser
    
    ' 遍历所有已注册解析器 / Iterate all registered parsers
    For Each parser In m_parsers
        If parser.CanParse(segment) Then
            Set GetParser = parser
            Call WriteLog("CParserFactory", "GetParser", "找到匹配解析器 / Found matching parser: " & TypeName(parser) & " 对于分段 / for segment: " & segment, "工厂模式")
            Exit Function
        End If
    Next parser
    
    ' 无匹配解析器 / No matching parser found
    Call WriteLog("CParserFactory", "GetParser", "未找到匹配解析器 / No matching parser found for segment: " & segment, "警告")
    Set GetParser = Nothing
End Function
