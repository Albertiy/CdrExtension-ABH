VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����ñ"
   ClientHeight    =   5040
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '����������
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


'���ڳ�ʼ�����ڼ���֮ǰ
Private Sub UserForm_Initialize()
    InitProductList
    InitColorList
    'ResourcePath = "D:\MyWorkspace\cdr\����ñ-��ͼƬ\"
    'MsgBox Application.Path  ' �� Draw Ϊֹ
    ResourcePath = Application.Path & "GMS\ABH-Resource\"
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
    
    If typeStr = "" Then
        MsgBox "��ѡ��һ����Ʒ"
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
        MsgBox "��Ч�Ĳ�Ʒ��"
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
    Dim Doc As Document
    Dim p As Page
    Dim l As layer
    Dim s As Shape
    
    templatePath = ResourcePath & typeStr & ".cdt"
    'MsgBox (Application.Path)
    Set Doc = Application.CreateDocumentFromTemplate(templatePath, True)
    'MsgBox (doc.Name)
    Set p = Doc.Pages(1) ' 0 is MainPage
    'MsgBox (p.Name)
    Set l = p.Layers("ͼ�� 1")
    'MsgBox (l.Name)
    
    Doc.Unit = cdrMillimeter ' set unit to mm
    
    Dim group1 As Shape
    Set group1 = l.Shapes("group1")
    
    Dim nameTextShape As Shape
    Dim ageTextShape As Shape
    
    Dim logoBack1 As Shape
    Set logoBack1 = l.Shapes("logoback1")
    
    Set nameTextShape = l.Shapes("name1")
    oldText = nameTextShape.Text.Story.Text '��ȡ�ַ���
    'MsgBox (oldText)
    nameTextShape.Text.Replace oldText, nameStr, True, ReplaceAll:=True '�滻����
    
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
    
    '�������� �ֶ���Բ�� W05 TsangerShuYuanT-W05 ����ɫ;
    '���嵥λ�� pt��1pt=0.353���� 40mm= 113.314
    
    On Error GoTo FontSetError '��ʵ���ᱨ��
    nameTextShape.Text.Story.Font = "�ֶ���Բ�� W05"
    ageTextShape.Text.Story.Font = "�ֶ���Բ�� W05"
    On Error GoTo 0
    
    nameTextShape.Text.Story.Fill.UniformColor = ColorList(colorStr)
    
    If typeStr = "type-4" Or typeStr = "type-5" Then
        ageTextShape.Text.Story.Fill.UniformColor = ColorList("yellow")
    Else
        ageTextShape.Text.Story.Fill.UniformColor = ColorList("black")
    End If
    
    '���ɱ߿�
    Dim nameEffect As Effect
    Dim nameSrFromContour As ShapeRange
    Dim nameUpperBack1 As Shape
    Dim nameBack1 As Shape
    
    Set nameEffect = nameTextShape.CreateContour(cdrContourOutside, 4, 1, OutlineColor:=ColorList("black"), FillColor:=ColorList("black"), CornerType:=cdrContourCornerRound)
    ' Not Suggested Sub : Shape.Separate; ��Ϊû�з���ֵ
    ' Suggested Function: Effect.Separate; ����ShapeRange
    Set nameSrFromContour = nameEffect.Separate
    Set nameUpperBack1 = nameSrFromContour(1)
    nameUpperBack1.Name = "nameupperback1"
    Set nameBack1 = nameUpperBack1.Duplicate
    nameBack1.Name = "nameback1"
    nameBack1.OrderBackOf logoBack1


    Dim ageEffect As Effect
    Dim ageSrFromContour As ShapeRange
    Dim ageUpperBack1 As Shape
    Dim ageBack1 As Shape
    
    If typeStr = "type-4" Or typeStr = "type-5" Then
        If ageStr <> "" Then
            Set ageEffect = ageTextShape.CreateContour(cdrContourOutside, 4, 1, OutlineColor:=ColorList("black"), FillColor:=ColorList("black"), CornerType:=cdrContourCornerRound)
            Set ageSrFromContour = ageEffect.Separate
            Set ageUpperBack1 = ageSrFromContour(1)
            ageUpperBack1.Name = "ageupperback1"
            Set ageBack1 = ageUpperBack1.Duplicate
            ageBack1.Name = "ageback1"
            ageBack1.OrderBackOf logoBack1
        End If
    End If
    
    Dim backSr As New ShapeRange
    Dim backBoundary As Shape
    
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
    
    Set backBoundary = backSr.CreateBoundary(0, 0, True, False)
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

    Doc.Unit = 1
        
    Exit Function
FontSetError:
    
    MsgBox "ȱʧ����: �ֶ���Բ�� W05"
    
    
End Function
'����ȫ������
Public Function GenOutLine()
    Dim p As Page
    Dim l As layer
    Dim allBoundary As Shape
    Set p = Application.ActivePage
    Set l = p.Layers("ͼ�� 1")
    Set allBoundary = l.Shapes.All.CreateBoundary(0, 0, True, False)
    
    Dim newLayer As layer
    Set newLayer = p.Layers.Find("����ͼ��")
    If newLayer Is Nothing Then
        Set newLayer = p.CreateLayer("����ͼ��")
    Else
        Set newLayer = p.Layers("����ͼ��")
    End If
    Dim oldBoundary As Shape
    Set oldBoundary = newLayer.Shapes.FindShape("allBoundary")
    If oldBoundary Is Nothing Then
    Else
        oldBoundary.Delete     '���ԭ�е���״��������
    End If
    '��֪Ϊ�� MoveToLayer ���Ǳ���������������ͻ�����԰��޸����ַŵ��ƶ���
    'MsgBox newLayer.Name
    allBoundary.MoveToLayer newLayer
    allBoundary.Name = "allBoundary"
    
End Function
