VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "生日帽"
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OutlineButton_Click()
    GenOutLine
End Sub

Private Sub TypeBox_Change()
    Dim typeStr As String
    typeStr = TypeBox.Text
    Dim p As Product
    Set p = GetItemFromProductList(typeStr)
    imagePath = ResourcePath & p.Template & "-" & p.color & ".bmp"
    'MsgBox imagePath
    'TypeImage.Picture = Nothing
    TypeImage.Picture = LoadPicture(imagePath)
End Sub


'窗口初始化，在激活之前
Private Sub UserForm_Initialize()
    InitProductList
    InitColorList
    'ResourcePath = "D:\MyWorkspace\cdr\生日帽-单图片\"
    'MsgBox Application.Path  ' 到 Draw 为止
    ResourcePath = Application.Path & "GMS\ABH-Resource\"
    'MsgBox ResourcePath
    TypeImage.Picture = LoadPicture(ResourcePath & "type.bmp")
    'MsgBox ("load type: " & ProductList.Count)
    For Index = 1 To ProductList.Count
        TypeBox.AddItem ProductList(Index).skuName
    Next Index
    
    TypeBox.ListIndex = 0 ' First is 0
    
End Sub

Private Sub GenerateButton_Click()
    Dim nameStr As String
    Dim ageStr As String
    Dim typeStr As String
    
    nameStr = Trim(NameTextBox.Text)
    ageStr = Trim(AgeTextBox.Text)
    typeStr = TypeBox.Text
    'MsgBox NameTextBox.Text & AgeTextBox.Text & TypeBox.Value
    
    If nameStr = "" Then
        MsgBox "姓名是必填项"
        Exit Sub
    End If
    
    If typeStr = "" Then
        MsgBox "请选择一个产品"
        Exit Sub
    End If
    
    Dim ptype As String
    Dim pcolor As String
    ptype = ""
    Dim pItem As Product
    For Each pItem In ProductList
        If pItem.skuName = typeStr Then
            ptype = pItem.Template
            pcolor = pItem.color 'Is String
            'MsgBox ColorList(pcolor).toString
        End If
    Next pItem
    
    If ptype = "" Then
        MsgBox "无效的产品名"
        Exit Sub
    End If
    
    GenFile nameStr, ageStr, ptype, pcolor
End Sub

Public Function GetItemFromProductList(nameStr As String) As Product
    For Each pItem In ProductList
        If pItem.skuName = nameStr Then
            Set GetItemFromProductList = pItem
            Exit Function
        End If
    Next pItem
End Function

Public Function GenFile(nameStr As String, ageStr As String, typeStr As String, colorStr As String)
    'If not input, nameStr = ''
    
    Dim Doc As Document
    Dim p As Page
    Dim l As layer
    Dim s As shape
    
    templatePath = ResourcePath & typeStr & ".cdt"
    'MsgBox (templatePath)
    Set Doc = Application.CreateDocumentFromTemplate(templatePath, True)
    'MsgBox (doc.Name)
    Set p = Doc.Pages(1) ' 0 is MainPage
    'MsgBox (p.Name)
    Set l = p.Layers("图层 1")
    'Exit Function
    'MsgBox (l.Name)
    
    'MsgBox Doc.Unit
    Doc.Unit = cdrMillimeter ' set unit to mm
    
    Dim group1 As shape
    Set group1 = l.Shapes("group1")
    
    Dim nameTextShape As shape
    Dim ageTextShape As shape
    
    Dim logoBack1 As shape
    Set logoBack1 = l.Shapes("logoback1")
    
    Set nameTextShape = l.Shapes("name1")
    oldText = nameTextShape.Text.Story.Text '获取字符串
    'MsgBox (oldText)
    nameTextShape.Text.Replace oldText, nameStr, True, ReplaceAll:=True '替换文字
    
    Set ageTextShape = l.Shapes("age1")
    oldText = ageTextShape.Text.Story.Text 'Get Old Age Str
    'MsgBox (oldText)
    ageTextShape.Text.Replace oldText, ageStr, True, ReplaceAll:=True
    
    If typeStr = "type-5" And ageStr <> "" Then
        'MsgBox "type-5" & ageStr
        If Mid(ageStr, 1, 1) = "1" Then
            ageTextShape.Move -8, 0
        End If
        If Mid(ageStr, 1, 1) = "7" Then
            ageTextShape.Move -4, 0
        End If
    End If
    
    '设置字体 仓耳舒圆体 W05 TsangerShuYuanT-W05 和颜色;
    '字体单位是 pt，1pt=0.353毫米 40mm= 113.314
    
    On Error GoTo FontSetError '其实不会报错
    nameTextShape.Text.Story.Font = "仓耳舒圆体 W05"
    ageTextShape.Text.Story.Font = "仓耳舒圆体 W05"
    On Error GoTo 0
    
    nameTextShape.Text.Story.Fill.UniformColor = ColorList(colorStr)
    
    If typeStr = "type-4" Or typeStr = "type-5" Then
        ageTextShape.Text.Story.Fill.UniformColor = ColorList("yellow")
    Else
        ageTextShape.Text.Story.Fill.UniformColor = ColorList("black")
    End If
    
    '生成边框
    Dim nameEffect As Effect
    Dim nameSrFromContour As ShapeRange
    Dim nameUpperBack1 As shape
    Dim nameBack1 As shape
    'nameTextShape.ConvertToCurves
    'nameTextShape.ConvertToCurves
    Set nameEffect = nameTextShape.CreateContour(cdrContourOutside, 4, 1, OutlineColor:=ColorList("black"), FillColor:=ColorList("black"), CornerType:=cdrContourCornerRound)
    ' Not Suggested Sub : Shape.Separate; 因为没有返回值
    ' Suggested Function: Effect.Separate; 返回ShapeRange
    Set nameSrFromContour = nameEffect.Separate
    Set nameUpperBack1 = nameSrFromContour(1)
    nameUpperBack1.Name = "nameupperback1"
    Set nameBack1 = nameUpperBack1.Duplicate
    nameBack1.Name = "nameback1"
    nameBack1.OrderBackOf logoBack1


    Dim ageEffect As Effect
    Dim ageSrFromContour As ShapeRange
    Dim ageUpperBack1 As shape
    Dim ageBack1 As shape
    
    If typeStr = "type-4" Or typeStr = "type-5" Then
        If ageStr <> "" Then
            Set ageEffect = ageTextShape.CreateContour(cdrContourOutside, 4, 1, OutlineColor:=ColorList("black"), FillColor:=ColorList("black"), CornerType:=cdrContourCornerRound)
            Set ageSrFromContour = ageEffect.Separate   ' 从原形状上拆分轮廓图
            Set ageUpperBack1 = ageSrFromContour(1)
            ageUpperBack1.Name = "ageupperback1"
            Set ageBack1 = ageUpperBack1.Duplicate
            ageBack1.Name = "ageback1"
            ageBack1.OrderBackOf logoBack1
        End If
    End If
    
    Dim backSr As New ShapeRange
    Dim backBoundary As shape
    
    backSr.Add nameBack1
    If typeStr = "type-4" Or typeStr = "type-5" Then
        If ageBack1 Is Nothing Then
        Else
            backSr.Add ageBack1
        End If
    End If
    If typeStr <> "type-3" Then
        backSr.Add logoBack1
    End If
    
    'X4中没有CreateBoundary，但可以像调用插件一样调用此功能
    'Set backBoundary = backSr.CreateBoundary(0, 0, True, False)
    Set backBoundary = backSr.CustomCommand("Boundary", "CreateBoundary") ', 0, 0, True, False
    backBoundary.Name = "backBoundary"
    backBoundary.Fill.UniformColor = ColorList("black")
    
    backBoundary.OrderBackOf group1
        
    ' Create Outline
'    Dim allBoundary As Shape
'    Set allBoundary = l.Shapes.All.CreateBoundary(0, 0, True, False)
'    Dim allShape As New ShapeRange
'
'
'    allShape.Add group1
'    allShape.Add backBoundary
'
'    Set allBoundary = allShape.CreateBoundary(0, 0, True, False)
'    allBoundary.Name = "allBoundary"
    
    'MsgBox Doc.Unit
    Doc.Unit = 1 ' reset unit
    'MsgBox Doc.Unit
    
    Exit Function
FontSetError:
    
    MsgBox "缺失字体: 仓耳舒圆体 W05"
    
    
End Function
'创建全部轮廓
Public Function GenOutLine()
    Dim p As Page
    Dim l As layer
    Dim allBoundary As shape
    Set p = Application.ActivePage
    Set l = p.Layers("图层 1")
    '同理
    'Set allBoundary = l.Shapes.All.CreateBoundary(0, 0, True, False)
    ' age 变成空字符串会报错，所以只用backBoundary创建轮廓（偷懒了哈哈）
    
    Dim ageTextShape As shape
    Set ageTextShape = l.Shapes.FindShape("age1")   'l.Shapes("age1")
    'If ageTextShape Is Nothing Then
    '    MsgBox "age1 is Nothing"
    'End If
    If ageTextShape Is Nothing Then
    ElseIf ageTextShape.Text.Story.Text = "" Then
        ageTextShape.Delete '直接删除，没那么多屁事了:D
        'MsgBox "删除age！"
        'Dim shape As shape
        'For Each shape In l.Shapes.All
        '    If shape.Name = "backBoundary" Then
        '        Set allBoundary = shape.CustomCommand("Boundary", "CreateBoundary")
        '    End If
        'Next shape
    End If
    Set allBoundary = l.Shapes.All.CustomCommand("Boundary", "CreateBoundary") '(0, 0, True, False)
    If allBoundary Is Nothing Then
        MsgBox "未找到 backBoundary 形状，描边失败"
        Exit Function
    End If
    
    Dim newLayer As layer
    Set newLayer = p.Layers.Find("轮廓图层")
    If newLayer Is Nothing Then
        Set newLayer = p.CreateLayer("轮廓图层")
    Else
        Set newLayer = p.Layers("轮廓图层")
    End If
    Dim oldBoundary As shape
    Set oldBoundary = newLayer.Shapes.FindShape("allBoundary")
    If oldBoundary Is Nothing Then
    Else
        oldBoundary.Delete     '清空原有的形状（轮廓）
    End If
    '不知为何 MoveToLayer 老是报错，可能是命名冲突，所以把修改名字放到移动后
    'MsgBox newLayer.Name
    allBoundary.MoveToLayer newLayer
    allBoundary.Name = "allBoundary"
    
End Function
