VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Khoi dong ung dung khi mo file
    On Error Resume Next
    Application.EnableEvents = True
    
    ' Goi ham khoi dong tu modImport
    modImport.InitializeApplication
    
    ' Khoi tao event handler cho sheet MainMenu
    modMain.InitializeMainMenuSheet
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Xu ly khi sheet duoc kich hoat
    On Error Resume Next
    
    ' Neu sheet MainMenu duoc kich hoat
    If Sh.Name = SHEET_MAIN_MENU Then
        modMain.InitializeMainMenuSheet
    End If
End Sub
