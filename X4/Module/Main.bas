Attribute VB_Name = "Main"
'Public Sub run()
'    Dim templatePath As String
'    templatePath = "D:\MyWorkspace\cdr\生日帽-单图片"
'
'    MsgBox (templatePath)
'    'import '
'    Dim iFilt As ImportFilter
'    Dim importProps As New StructImportOptions
'    Set iFilt = ActiveLayer.ImportEx(templatePath, cdrAutoSense, importProps)
'    iFilt.Finish
'    Dim currShape As Shape
'    Set currShape = ActiveShape
'    currShape.UngroupAll
'
'    Dim i As Integer
'    For i = 1 To ActivePage.Shapes.Count
'        Dim obj As Shape
'        Dim tag As String
'        Set obj = ActivePage.Shapes.Item(i)
'        tag = obj.URL.Address
'        If tag = "name" Then
'            obj.Text.Story.Text = "Albertiy"
'            obj.URL.Address = ""
'        End If
'        If tag = "job" Then
'            obj.Text.Story.Text = "Software Develop Enginner"
'        If tag = "phone" Then
'            obj.Text.Story.Text = "18936987963"
'            obj.URL.Address = ""
'        End If
'    Next i
'
'
'
'    ''
'    'Open fileContentPath For Input As #FreeNum'
'    '    Do Until EOF(FreeNum)'
'     '       Line Input #FreeNum, lineText'
'      '  Loop'
'        '必须关闭'
'    'Close FreeNum'
'    'ActiveDocument.addPages()'
'
'End Sub

'Sub Test()
''    Dim answer As Integer
''    answer = MsgBox("Hello?", vbQuestion + vbYesNoCancel, "Oh Title")
''
''    If answer = vbYes Then
''        MsgBox ("Yes!")
''    End If
''
''    If answer = vbNo Then
''        MsgBox ("No~")
''    End If
''
''    If answer = vbCancel Then
''        MsgBox ("It's fine if you don't want to say")
''    End If
''    Dim cdr As CorelDRAW.Application
''    Set cdr = CreateObject("CorelDRAW.Application")
''    Dim Doc As Document
''    Set Doc = Application.ActiveDocument
''    MsgBox (Doc)
'     Application.ImportWorkspace Application.Path & "GMS\ABH-Resource\ABH_Toolbar.cdws"
'End Sub
Sub Test()
    'MsgBox "Test"
    Dim l As layer
    Set l = Application.ActivePage.Layers("图层 1")
    Dim s As Shape
    'Set s = l.Shapes.All.CreateBoundary(0, 0, True, False)
    
    Application.ActiveDocument.ClearSelection
    l.Shapes.All.AddToSelection
    
    's.Name = "new Shape"
End Sub
