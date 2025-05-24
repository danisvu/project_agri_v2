VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMainMenuSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================
' Class: clsMainMenuSheet
' Mo ta: Xu ly su kien cho sheet MainMenu
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 21/05/2025
' ======================================================

Private WithEvents wsMainMenu As Worksheet
Attribute wsMainMenu.VB_VarHelpID = -1
Private blnInitializing As Boolean

' Khoi tao class va gan sheet MainMenu
Public Sub Initialize(ByRef ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Gan sheet
    Set wsMainMenu = ws
    ' Dat co dang khoi tao
    blnInitializing = True
    
    Exit Sub
    
ErrorHandler:
    LogError "clsMainMenuSheet.Initialize", Err.Number, Err.Description
End Sub

Private Sub wsMainMenu_Activate()
    ' Tat xu ly loi de tranh hien thi thong bao
    On Error Resume Next
    
    ' Kiem tra sheet da duoc khoi tao chua
    If wsMainMenu Is Nothing Then Exit Sub
    
    ' Neu dang trong qua trinh khoi tao, chi can bo qua
    If blnInitializing Then
        blnInitializing = False
        Exit Sub
    End If
    
    ' Tam dung de tranh xu ly qua nhanh
    Application.Wait Now + TimeValue("00:00:00.3")
    
    ' Goi truc tiep form chu khong qua ham
    Dim frm As Object
    
    ' Kiem tra form MainMenu co ton tai khong
    On Error Resume Next
    Set frm = UserForms("frmMainMenu")
    
    ' Neu form chua duoc tao, thi load no
    If frm Is Nothing Then
        ' Load form truc tiep - cach an toan nhat
        VBA.UserForms.Add("frmMainMenu").Show
    ElseIf frm.Visible = False Then
        ' Neu form da ton tai nhung dang an, thi hien thi no
        frm.Show
    End If
End Sub

