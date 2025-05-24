Attribute VB_Name = "modImport"
Option Explicit

' ======================================================
' Module: modImport
' Mo ta: Xu ly import du lieu tu files IPCAS vao he thong
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 18/05/2025
' ======================================================

' Bien toan cuc theo doi trang thai import
Private blnIsImporting As Boolean

' Show form import du lieu
Public Sub ShowImportForm()
    On Error GoTo ErrorHandler
    
    ' Kiem tra neu Excel dang chay macro khac
    If blnIsImporting Then
        MsgBox "He thong dang xu ly import du lieu. Vui long cho...", vbInformation, TITLE_IMPORT
        Exit Sub
    End If
    
    ' Kiem tra neu file dang o che do chi doc
    If IsWorkbookReadOnly() Then
        MsgBox MSG_WORKBOOK_READ_ONLY, vbExclamation, TITLE_ERROR
        Exit Sub
    End If
    
    ' Kiem tra phien ban Excel
    If Not IsExcelVersionCompatible() Then
        MsgBox "Phien ban Excel khong tuong thich. Can Excel 2007 tro len.", vbExclamation, TITLE_ERROR
        Exit Sub
    End If
    
    ' Load form Import
    Load frmImport
    
    ' Cap nhat trang thai cac loai du lieu
    UpdateImportFormStatus
    
    ' Hien thi form
    frmImport.Show vbModeless
    
    Exit Sub
    
ErrorHandler:
    blnIsImporting = False
    LogError "ShowImportForm", Err.Number, Err.Description
    MsgBox "Loi khi mo form import: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Cap nhat trang thai du lieu tren form Import
Public Sub UpdateImportFormStatus()
    On Error GoTo ErrorHandler
    
    ' Kiem tra form da duoc load chua
    If frmImport Is Nothing Then Exit Sub
    
    ' Cap nhat trang thai cho tung loai du lieu
    
    ' Du no
    If sheetExists(SHEET_DU_NO) Then
        frmImport.lblStatusDuNo.Caption = "Da import - " & GetLastImportInfo(SHEET_DU_NO)
        frmImport.lblStatusDuNo.ForeColor = RGB(0, 128, 0) ' Mau xanh la
    Else
        frmImport.lblStatusDuNo.Caption = "Chua import"
        frmImport.lblStatusDuNo.ForeColor = RGB(192, 0, 0) ' Mau do
    End If
    
    ' Tai san
    If sheetExists(SHEET_TAI_SAN) Then
        frmImport.lblStatusTaiSan.Caption = "Da import - " & GetLastImportInfo(SHEET_TAI_SAN)
        frmImport.lblStatusTaiSan.ForeColor = RGB(0, 128, 0) ' Mau xanh la
    Else
        frmImport.lblStatusTaiSan.Caption = "Chua import"
        frmImport.lblStatusTaiSan.ForeColor = RGB(192, 0, 0) ' Mau do
    End If
    
    ' Tra goc
    If sheetExists(SHEET_TRA_GOC) Then
        frmImport.lblStatusTraGoc.Caption = "Da import - " & GetLastImportInfo(SHEET_TRA_GOC)
        frmImport.lblStatusTraGoc.ForeColor = RGB(0, 128, 0) ' Mau xanh la
    Else
        frmImport.lblStatusTraGoc.Caption = "Chua import"
        frmImport.lblStatusTraGoc.ForeColor = RGB(192, 0, 0) ' Mau do
    End If
    
    ' Tra lai
    If sheetExists(SHEET_TRA_LAI) Then
        frmImport.lblStatusTraLai.Caption = "Da import - " & GetLastImportInfo(SHEET_TRA_LAI)
        frmImport.lblStatusTraLai.ForeColor = RGB(0, 128, 0) ' Mau xanh la
    Else
        frmImport.lblStatusTraLai.Caption = "Chua import"
        frmImport.lblStatusTraLai.ForeColor = RGB(192, 0, 0) ' Mau do
    End If
    
    ' Cap nhat trang thai tong the - kiem tra du du lieu va dung ma chi nhanh
    Dim isDataValid As Boolean
    Dim hasAllData As Boolean
    
    ' Kiem tra co du 4 loai du lieu
    hasAllData = IsDataComplete()
    
    ' Neu co du du lieu, kiem tra ma chi nhanh
    If hasAllData Then
        isDataValid = ValidateImportedDataWithBranchCode()
        
        If isDataValid Then
            ' Du lieu day du va hop le
            frmImport.cmdContinue.Enabled = True
            frmImport.lblStatusComplete.Caption = MSG_READY_TO_CONTINUE
            frmImport.lblStatusComplete.ForeColor = RGB(0, 128, 0) ' Mau xanh la
        Else
            ' Du lieu day du nhung khong hop le
            frmImport.cmdContinue.Enabled = False
            frmImport.lblStatusComplete.Caption = "Data khong dung"
            frmImport.lblStatusComplete.ForeColor = RGB(192, 0, 0) ' Mau do
        End If
    Else
        ' Chua co du du lieu
        frmImport.cmdContinue.Enabled = False
        frmImport.lblStatusComplete.Caption = "Can import: " & GetMissingDataSheets()
        frmImport.lblStatusComplete.ForeColor = RGB(192, 0, 0) ' Mau do
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "UpdateImportFormStatus", Err.Number, Err.Description
End Sub

' Thuc hien import du lieu
Public Sub ImportData(ByVal fileType As String)
    On Error GoTo ErrorHandler
    
    ' Dat co hieu dang import
    blnIsImporting = True
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Dang import du lieu " & fileType & "..."
    
    ' Lay duong dan file can import
    Dim filePath As String
    filePath = GetFilePath(fileType)
    
    ' Kiem tra neu nguoi dung huy chon file
    If filePath = "" Then
        Application.StatusBar = "Da huy import du lieu."
        GoTo CleanUp
    End If
    
    ' Lay pattern tuong ung voi loai file
    Dim filePattern As String
    Select Case fileType
        Case FILE_TYPE_DU_NO
            filePattern = FILE_PATTERN_DU_NO
        Case FILE_TYPE_TAI_SAN
            filePattern = FILE_PATTERN_TAI_SAN
        Case FILE_TYPE_TRA_GOC
            filePattern = FILE_PATTERN_TRA_GOC
        Case FILE_TYPE_TRA_LAI
            filePattern = FILE_PATTERN_TRA_LAI
        Case Else
            filePattern = ""
    End Select
    
    ' Kiem tra ten file co phu hop khong
    If Not ValidateFileNamePattern(filePath, filePattern) Then
        Dim msgText As String
        msgText = Replace(MSG_FILE_NAME_INVALID, "{0}", """" & filePattern & """")
        MsgBox msgText, vbExclamation, TITLE_ERROR
        GoTo CleanUp
    End If
    
    ' Kiem tra tinh hop le cua file
    If Not ValidateImportFile(filePath, fileType) Then
        MsgBox MSG_FILE_INVALID, vbExclamation, TITLE_ERROR
        GoTo CleanUp
    End If
    
    ' Xac dinh ten sheet tuong ung
    Dim sheetName As String
    sheetName = GetSheetNameFromFileType(fileType)
    
    If sheetName = "" Then
        MsgBox "Loai file khong hop le!", vbExclamation, TITLE_ERROR
        GoTo CleanUp
    End If
    
    ' Tao hoac lay sheet
    Dim targetSheet As Worksheet
    
    If sheetExists(sheetName) Then
        Set targetSheet = ThisWorkbook.Worksheets(sheetName)
        
        ' Hoi nguoi dung co muon ghi de du lieu cu
        If Application.WorksheetFunction.CountA(targetSheet.UsedRange) > 0 Then
            If MsgBox(MSG_IMPORT_OVERWRITE, vbQuestion + vbYesNo, TITLE_CONFIRMATION) = vbNo Then
                GoTo CleanUp
            End If
        End If
        
        ' Xoa du lieu cu nhung giu lai thong tin o dong dau tien
        ClearWorksheetData targetSheet
    Else
        ' Tao sheet moi
        Set targetSheet = ThisWorkbook.Worksheets.Add
        targetSheet.Name = sheetName
    End If
    
    ' Ghi thong tin ngay gio import vao dong dau tien
    targetSheet.Cells(INFO_ROW, 1).Value = IMPORT_INFO_PREFIX & Now()
    targetSheet.Cells(INFO_ROW, 1).Font.Bold = True
    
    ' Mo workbook nguon
    Application.StatusBar = "Dang doc file nguon..."
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    
    ' Lay sheet dau tien
    Dim wsSource As Worksheet
    Set wsSource = wbSource.Worksheets(1)
    
    ' Copy toan bo du lieu tu nguon
    Application.StatusBar = "Dang sao chep du lieu..."
    wsSource.UsedRange.Copy
    
    ' Dan vao sheet dich tu dong 2 (duoi dong thong tin import)
    targetSheet.Cells(FIRST_DATA_ROW, 1).PasteSpecial xlPasteValues
    targetSheet.Cells(FIRST_DATA_ROW, 1).PasteSpecial xlPasteFormats
    
    ' Format worksheet
    Application.StatusBar = "Dang dinh dang du lieu..."
    FormatWorksheet targetSheet
    
    ' Dong workbook nguon
    wbSource.Close SaveChanges:=False
    
    ' Thong bao thanh cong
    Application.StatusBar = "Import du lieu " & fileType & " thanh cong."
    MsgBox MSG_IMPORT_SUCCESS, vbInformation, TITLE_SUCCESS
    
    ' Ghi log thanh cong
    LogInfo "ImportData", "Import " & fileType & " thanh cong tu " & filePath
    
    ' Cap nhat trang thai form
    UpdateImportFormStatus
    
CleanUp:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Bo co dang import
    blnIsImporting = False
    
    Exit Sub
    
ErrorHandler:
    ' Dong workbook nguon neu dang mo
    On Error Resume Next
    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If
    
    ' Ghi log loi
    LogError "ImportData", Err.Number, Err.Description
    
    ' Thong bao loi
    MsgBox MSG_IMPORT_FAILED & vbNewLine & "Chi tiet loi: " & Err.Description, vbExclamation, TITLE_ERROR
    
    ' Khoi phuc trang thai
    Resume CleanUp
End Sub

' Thuc hien cac xu ly sau khi import du lieu
Public Sub ProcessAfterImport()
    On Error GoTo ErrorHandler
    
    ' Kiem tra du lieu da du chua
    If Not IsDataComplete() Then
        MsgBox MSG_DATA_INCOMPLETE, vbExclamation, TITLE_WARNING
        Exit Sub
    End If
    
    ' Kiem tra du lieu co hop le khong
    If Not ValidateImportedDataWithBranchCode() Then
        MsgBox "Data khong dung", vbExclamation, TITLE_WARNING
        Exit Sub
    End If
    
    ' An form Import neu dang hien thi
    If Not frmImport Is Nothing Then
        Unload frmImport
    End If
    
    ' Tao sheet MainMenu neu chua ton tai va cap nhat thong tin
    modMain.CreateMainMenuSheet
    
    ' Tao tong quan du lieu
    GenerateDataSummary
    
    ' Cap nhat thong tin MainMenu
    modMain.UpdateMainMenuAfterImport
    
    ' Hien thi sheet MainMenu
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Cho 1 giay de dam bao cac thao tac truoc da hoan tat
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Hien thi form MainMenu
    Application.EnableEvents = True
    modMain.ShowMainMenuForm
    
    Exit Sub
    
ErrorHandler:
    LogError "ProcessAfterImport", Err.Number, Err.Description
    MsgBox "Loi khi xu ly sau import: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Ham nay se duoc goi boi su kien Workbook_Open
Public Sub InitializeApplication()
    On Error GoTo ErrorHandler
    
    ' Ghi log
    LogInfo "InitializeApplication", "Bat dau khoi dong ung dung"
    
    ' Tao cac sheet can thiet neu chua ton tai
    CreateSheetIfNotExists SHEET_LOG
    
    ' Kiem tra du lieu da duoc import chua
    If IsDataComplete() Then
        ' Tao sheet MainMenu neu chua ton tai
        If Not sheetExists(SHEET_MAIN_MENU) Then
            modMain.CreateMainMenuSheet
        End If
        
        ' Hien thi sheet MainMenu
        ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
        
        ' Mo form MainMenu
        Application.Wait Now + TimeValue("00:00:01")
        modMain.ShowMainMenuForm
        
        ' Khoi tao class sheet MainMenu
        modMain.InitializeMainMenuSheet
    Else
        ' Chua co du du lieu, can import
        ShowImportForm
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "InitializeApplication", Err.Number, Err.Description
    MsgBox "Loi khi khoi dong ung dung: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Them vao phan ThisWorkbook
' Private Sub Workbook_Open()
'     InitializeApplication
' End Sub

' Chuc nang test import du lieu
Public Sub TestImportData()
    On Error GoTo ErrorHandler
    
    ' Thong bao bat dau test
    MsgBox "Bat dau kiem tra chuc nang Import du lieu", vbInformation, "Test Import"
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    ' Ghi log bat dau test
    LogInfo "TestImportData", "Bat dau kiem tra chuc nang Import du lieu"
    
    ' Kiem tra co cac sheet du lieu chua
    Dim hasData As Boolean
    hasData = IsDataComplete()
    
    If hasData Then
        ' Tao ban sao du lieu truoc khi test
        BackupCurrentData
        
        ' Thong bao da backup du lieu
        Application.StatusBar = "Da tao ban sao du lieu hien tai"
    End If
    
    ' Mo form import
    ShowImportForm
    
    ' Ghi log ket thuc test
    LogInfo "TestImportData", "Ket thuc kiem tra chuc nang Import du lieu"
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log loi
    LogError "TestImportData", Err.Number, Err.Description
    
    ' Thong bao loi
    MsgBox "Loi khi test chuc nang import: " & Err.Description, vbExclamation, TITLE_ERROR
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub


' Cap nhat cac thong tin sau khi import thanh cong
Public Sub UpdateAfterSuccessfulImport()
    On Error GoTo ErrorHandler
    
    ' Kiem tra du lieu da du chua
    If Not IsDataComplete() Then
        Exit Sub
    End If
    
    ' Tao tong quan du lieu
    GenerateDataSummary
    
    ' Cap nhat hoac tao sheet MainMenu
    CreateMainMenuSheet
    
    ' An cac sheet du lieu de bao ve
    HideDataSheets
    
    ' Hien thi sheet MainMenu
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Ghi log
    LogInfo "UpdateAfterSuccessfulImport", "Cap nhat he thong sau khi import thanh cong"
    
    Exit Sub
    
ErrorHandler:
    LogError "UpdateAfterSuccessfulImport", Err.Number, Err.Description
End Sub
