VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_Start()
    'MsgBox "Start"
    Application.ImportWorkspace Application.Path & "GMS\ABH-Resource\ABH_Toolbar.cdws"
End Sub
