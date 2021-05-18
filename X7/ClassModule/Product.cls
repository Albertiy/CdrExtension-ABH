VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit    'VB6强制显式声明变量

Private mvarSkuName As String ' 名称
Private mvarTemplate As String ' 类型
Private mvarColor As String ' 颜色

Property Let skuName(v As String)
    mvarSkuName = v
End Property
Property Get skuName() As String
    skuName = mvarSkuName
End Property

Property Let Template(v As String)
    mvarTemplate = v
End Property
Property Get Template() As String
    Template = mvarTemplate
End Property

Property Let color(v As String)
    mvarColor = v
End Property
Property Get color() As String
    color = mvarColor
End Property

Private Sub Class_Initialize() 'vb6没有类似构造函数的参数

End Sub
'初始化函数
Public Function init(iSkuName As String, iTemplate As String, iColor As String)
    skuName = iSkuName
    Template = iTemplate
    color = iColor
End Function

Public Function toString() As String
    toString = "{ skuName= '" & skuName & "', template= '" & Template & "', color= '" & color & "' }"
End Function
