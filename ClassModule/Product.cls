VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit    'VB6ǿ����ʽ��������

Private mvarSkuName As String ' ����
Private mvarTemplate As String ' ����
Private mvarColor As String ' ��ɫ

Property Let skuName(v As String)
    mvarSkuName = v
End Property
Property Get skuName() As String
    skuName = mvarSkuName
End Property

Property Let template(v As String)
    mvarTemplate = v
End Property
Property Get template() As String
    template = mvarTemplate
End Property

Property Let color(v As String)
    mvarColor = v
End Property
Property Get color() As String
    color = mvarColor
End Property

Private Sub Class_Initialize() 'vb6û�����ƹ��캯���Ĳ���

End Sub
'��ʼ������
Public Function init(iSkuName As String, iTemplate As String, iColor As String)
    skuName = iSkuName
    template = iTemplate
    color = iColor
End Function

Public Function toString() As String
    toString = "{ skuName= '" & skuName & "', template= '" & template & "', color= '" & color & "' }"
End Function