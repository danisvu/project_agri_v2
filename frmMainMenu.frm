VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMainMenu 
   Caption         =   "Quan Ly Thong Tin Khach Hang Vay"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000.001
   OleObjectBlob   =   "frmMainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ======================================================
' Form: frmMainMenu
' Mo ta: Giao dien menu chinh cua ung dung
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 20/05/2025
' ======================================================

' Khoi tao form
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Thiet lap tieu de
    Me.Caption = TITLE_MAIN_MENU
    lblTitle.Caption = "HE THONG QUAN LY THONG TIN KHACH HANG VAY"
    lblTitle.ForeColor = AGRIBANK_RED_COLOR
    
    ' Thiet lap cac frame
    fraInfo.Caption = "THONG TIN DU LIEU"
    fraFunctions.Caption = "CHUC NANG CHINH"
    
    ' Thiet lap cac nhan
    lblDuNoTitle.Caption = "Du no:"
    lblTaiSanTitle.Caption = "Tai san dam bao:"
    lblTraGocTitle.Caption = "Tra goc:"
    lblTraLaiTitle.Caption = "Tra lai:"
    
    ' Mac dinh, cac label thong tin se duoc cap nhat trong RefreshMainMenuInfo
    lblDuNoInfo.Caption = ""
    lblTaiSanInfo.Caption = ""
    lblTraGocInfo.Caption = ""
    lblTraLaiInfo.Caption = ""
    
    lblDuNoCount.Caption = ""
    lblTaiSanCount.Caption = ""
    lblTraGocCount.Caption = ""
    lblTraLaiCount.Caption = ""
    
    ' Thiet lap cac nut chuc nang
    cmdKhachHang.Caption = "Tim kiem khach hang"
    cmdKhoanVay.Caption = "Quan ly khoan vay"
    cmdTaiSan.Caption = "Quan ly tai san dam bao"
    cmdTraNo.Caption = "Theo doi tra no"
    cmdCanhBao.Caption = "Canh bao rui ro"
    cmdBaoCao.Caption = "Bao cao & Thong ke"
    cmdImport.Caption = "Import lai du lieu"
    cmdClose.Caption = "Dong"
    
    ' Thiet lap tab index cho cac nut
    cmdKhachHang.TabIndex = 0
    cmdKhoanVay.TabIndex = 1
    cmdTaiSan.TabIndex = 2
    cmdTraNo.TabIndex = 3
    cmdCanhBao.TabIndex = 4
    cmdBaoCao.TabIndex = 5
    cmdImport.TabIndex = 6
    cmdClose.TabIndex = 7
    
    ' Trang thai he thong
    lblStatus.Caption = "Khoi tao he thong..."
    
    ' Tao hinh anh logo
    With imgLogo
        imgLogo.BackColor = AGRIBANK_RED_COLOR
    End With
    
    ' Cap nhat thong tin hien thi
    RefreshMainMenuInfo
    
    Exit Sub
    
ErrorHandler:
    LogError "frmMainMenu_Initialize", Err.Number, Err.Description
    MsgBox "Loi khi khoi tao form: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien khi form duoc kich hoat
Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    
    ' Cap nhat thong tin tren form
    RefreshMainMenuInfo
    
    Exit Sub
    
ErrorHandler:
    LogError "frmMainMenu_Activate", Err.Number, Err.Description
End Sub

' Xu ly su kien dong form
Private Sub UserForm_Terminate()
    On Error Resume Next
    
    ' Goi ham xu ly khi dong form
    OnMainMenuFormClosed
End Sub

' Xu ly su kien nut Tim kiem khach hang
Private Sub cmdKhachHang_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Tim kiem khach hang
    NavigateToFunction FUNC_KHACH_HANG
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdKhachHang_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Tim kiem khach hang: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Quan ly khoan vay
Private Sub cmdKhoanVay_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Quan ly khoan vay
    NavigateToFunction FUNC_KHOAN_VAY
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdKhoanVay_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Quan ly khoan vay: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Quan ly tai san
Private Sub cmdTaiSan_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Quan ly tai san
    NavigateToFunction FUNC_TAI_SAN
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdTaiSan_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Quan ly tai san: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Theo doi tra no
Private Sub cmdTraNo_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Theo doi tra no
    NavigateToFunction FUNC_TRA_NO
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdTraNo_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Theo doi tra no: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Canh bao rui ro
Private Sub cmdCanhBao_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Canh bao rui ro
    NavigateToFunction FUNC_CANH_BAO
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdCanhBao_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Canh bao rui ro: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Bao cao & Thong ke
Private Sub cmdBaoCao_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Bao cao & Thong ke
    NavigateToFunction FUNC_BAO_CAO
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdBaoCao_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Bao cao & Thong ke: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Import lai du lieu
Private Sub cmdImport_Click()
    On Error GoTo ErrorHandler
    
    ' Dieu huong den chuc nang Import du lieu
    NavigateToFunction FUNC_IMPORT_DATA
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdImport_Click", Err.Number, Err.Description
    MsgBox "Loi khi mo chuc nang Import du lieu: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien nut Dong
Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler
    
    ' Dong form
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdClose_Click", Err.Number, Err.Description
    ' Van dong form ngay ca khi co loi
    Unload Me
End Sub

' Xu ly khi nhan phim Escape
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    ' Dong form khi nhan Escape
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
