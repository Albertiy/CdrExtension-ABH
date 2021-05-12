Attribute VB_Name = "Entrance"
Public ColorList As New Collection
Public ProductList As New Collection '定义全局变量列表
Public ResourcePath As String

'代码必须写在函数或子过程中，故单独写初始化函数
Public Function InitProductList()
    'If ProductList.Count <> 0 Then Exit Function
    If ProductList.Count <> 0 Then Set ProductList = New Collection
    
    '列表中的元素如果是对象，则是引用，多次添加同一对象，改其一则改全部
    '需要重新 new
    Dim a As Product
    Set a = New Product
    a.init "爱心粉", "type-1", "pink"
    'MsgBox a.toString
    ProductList.Add a, "爱心粉"
   
    Set a = New Product
    a.init "爱心蓝", "type-1", "blue"
    ProductList.Add a, "爱心蓝"
    
    Set a = New Product
    a.init "条纹粉", "type-3", "pink"
    ProductList.Add a, "条纹粉"
    
    Set a = New Product
    a.init "条纹蓝", "type-3", "blue"
    ProductList.Add a, "条纹蓝"
    
    Set a = New Product
    a.init "独角兽+五角星蓝", "type-4", "blue"
    ProductList.Add a, "独角兽+五角星蓝"
    
    Set a = New Product
    a.init "黛西+圆圈粉", "type-5", "pink"
    ProductList.Add a, "黛西+圆圈粉"
    
    Set a = New Product
    a.init "醒狮", "type-6", "classicred"
    ProductList.Add a, "醒狮"
    
'    MsgBox ("ProductList Count = " & ProductList.Count)
'    Dim item As Product
'    For Each item In ProductList
'     MsgBox "Product List Add " & item.toString
'    Next item

End Function
'初始化颜色列表
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
