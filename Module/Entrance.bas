Attribute VB_Name = "Entrance"
Public ColorList As New Collection
Public ProductList As New Collection '����ȫ�ֱ����б�
Public ResourcePath As String

'�������д�ں������ӹ����У��ʵ���д��ʼ������
Public Function InitProductList()
    'If ProductList.Count <> 0 Then Exit Function
    If ProductList.Count <> 0 Then Set ProductList = New Collection
    
    '�б��е�Ԫ������Ƕ����������ã�������ͬһ���󣬸���һ���ȫ��
    '��Ҫ���� new
    Dim a As Product
    Set a = New Product
    a.init "���ķ�", "type-1", "pink"
    'MsgBox a.toString
    ProductList.Add a, "���ķ�"
   
    Set a = New Product
    a.init "������", "type-1", "blue"
    ProductList.Add a, "������"
    
    Set a = New Product
    a.init "���Ʒ�", "type-3", "pink"
    ProductList.Add a, "���Ʒ�"
    
    Set a = New Product
    a.init "������", "type-3", "blue"
    ProductList.Add a, "������"
    
    Set a = New Product
    a.init "������+�������", "type-4", "blue"
    ProductList.Add a, "������+�������"
    
    Set a = New Product
    a.init "����+ԲȦ��", "type-5", "pink"
    ProductList.Add a, "����+ԲȦ��"
    
    Set a = New Product
    a.init "��ʨ", "type-6", "classicred"
    ProductList.Add a, "��ʨ"
    
'    MsgBox ("ProductList Count = " & ProductList.Count)
'    Dim item As Product
'    For Each item In ProductList
'     MsgBox "Product List Add " & item.toString
'    Next item

End Function
'��ʼ����ɫ�б�
Public Function InitColorList()
    If ColorList.Count <> 0 Then Set ColorList = New Collection
    Dim Pink As color
    Dim Blue As color
    Dim Yellow As color
    Dim Classicred As color
    Dim Black As color
    
    Set Pink = Application.CreateCMYKColor(0, 62, 8, 0)
    Set Blue = Application.CreateCMYKColor(62, 0, 8, 0)
    Set Yellow = Application.CreateCMYKColor(0, 8, 62, 0)
    Set Classicred = Application.CreateCMYKColor(10, 100, 100, 0)
    Set Black = Application.CreateCMYKColor(0, 0, 0, 100)
    
    ColorList.Add Pink, "pink"
    ColorList.Add Blue, "blue"
    ColorList.Add Yellow, "yellow"
    ColorList.Add Classicred, "classicred"
    ColorList.Add Black, "black"
    
'    MsgBox ("ColorList Count = " & ColorList.Count)
'    Dim item As color
'    For counter = 1 To ColorList.Count
'     MsgBox "Color List Add " & ColorList.item(counter).toString
'    Next counter
    
End Function


Sub ShowMainWindow()
'    MsgBox "HelloWorld"
    UserForm1.Show
    
End Sub
