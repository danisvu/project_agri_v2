VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Import Du Lieu"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955.001
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ======================================================
' Form: frmImport
' Mo ta: Giao dien import du lieu tu file IPCAS
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 18/05/2025
' ======================================================

' --- Khoi tao form ---
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Thiet lap tieu de
    Me.Caption = "Import Du Lieu"
    lblTitle.Caption = "IMPORT DU LIEU"
    
    ' Thiet lap các label
    fraStatus.Caption = "Trang thai du lieu"
    fraButtons.Caption = "Thao tác"
    
    lblDuNo.Caption = "Du no:"
    lblTaiSan.Caption = "Tài san dam bao:"
    lblTraGoc.Caption = "Tra goc:"
    lblTraLai.Caption = "Tra lai:"
    lblComplete.Caption = "Trang thái:"
    
    ' Thiiet lap trang thai ban dau
    lblStatusDuNo.Caption = "Chua import"
    lblStatusDuNo.ForeColor = RGB(192, 0, 0) ' Màu do
    
    lblStatusTaiSan.Caption = "Chua import"
    lblStatusTaiSan.ForeColor = RGB(192, 0, 0) ' Màu do
    
    lblStatusTraGoc.Caption = "Chua import"
    lblStatusTraGoc.ForeColor = RGB(192, 0, 0) ' Màu do
    
    lblStatusTraLai.Caption = "Chua import"
    lblStatusTraLai.ForeColor = RGB(192, 0, 0) ' Màu do
    
    lblStatusComplete.Caption = "Can import du lieu"
    lblStatusComplete.ForeColor = RGB(192, 0, 0) ' Màu do
    
    ' Thiet lap các nút
    cmdImportDuNo.Caption = "Import Du no"
    cmdImportTaiSan.Caption = "Import Tài san dam bao"
    cmdImportTraGoc.Caption = "Import Tra goc"
    cmdImportTraLai.Caption = "Import Tra lai"
    cmdContinue.Caption = "Tiep tuc >"
    cmdCancel.Caption = "Ðóng"
    
    ' Mac dinh nút Tiep tuc không duoc kích hoat
    cmdContinue.Enabled = False
    
    ' Cap nhat trang thái form dua tren du lieu hien co
    UpdateImportFormStatus
    
    Exit Sub
    
ErrorHandler:
    LogError "frmImport_Initialize", Err.Number, Err.Description
    MsgBox "Loi khi khoi tao form: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Import Du no
Private Sub cmdImportDuNo_Click()
    On Error GoTo ErrorHandler
    
    ' Goi hàm import du lieu Du no
    ImportData FILE_TYPE_DU_NO
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdImportDuNo_Click", Err.Number, Err.Description
    MsgBox "Loi khi import Du no: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Import Tai san
Private Sub cmdImportTaiSan_Click()
    On Error GoTo ErrorHandler
    
    ' Goi hàm import du lieu Tài san
    ImportData FILE_TYPE_TAI_SAN
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdImportTaiSan_Click", Err.Number, Err.Description
    MsgBox "Loi khi import Tai san: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Import Tra goc
Private Sub cmdImportTraGoc_Click()
    On Error GoTo ErrorHandler
    
    ' Goi hàm import du lieu Tra goc
    ImportData FILE_TYPE_TRA_GOC
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdImportTraGoc_Click", Err.Number, Err.Description
    MsgBox "Loi khi import Tra goc: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Import Tra lai
Private Sub cmdImportTraLai_Click()
    On Error GoTo ErrorHandler
    
    ' Goi hàm import du lieu Tra lai
    ImportData FILE_TYPE_TRA_LAI
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdImportTraLai_Click", Err.Number, Err.Description
    MsgBox "Loi khi import Tra lai: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Tiep tuc
Private Sub cmdContinue_Click()
    On Error GoTo ErrorHandler
    
    Me.Hide
    
    ProcessAfterImport
    
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdContinue_Click", Err.Number, Err.Description
    MsgBox "Loi khi chuyen tiep: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Xu ly su kien click nut Dong
Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler
    
    ' Kiem tra xem da import du lieu chua
    If Not IsDataComplete() Then
        ' Hoi xác nhan neu chua import du
        If MsgBox("Ban chua import du du lieu. Ban co chac muon dong form?", _
                  vbQuestion + vbYesNo, TITLE_CONFIRMATION) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Ðóng form
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    LogError "cmdCancel_Click", Err.Number, Err.Description
    ' Van dong form ngay khi co loi
    Unload Me
End Sub

' Xu ly khi dong form
Private Sub UserForm_Terminate()
    On Error Resume Next
    
    ' Ghi log
    LogInfo "frmImport_Terminate", "Form Import da bi dong"
End Sub

' Xu ly khi form duoc hien thi
Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    
    ' Cap nhat trang thái form
    UpdateImportFormStatus
    
    Exit Sub
    
ErrorHandler:
    LogError "frmImport_Activate", Err.Number, Err.Description
End Sub

