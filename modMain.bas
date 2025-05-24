Attribute VB_Name = "modMain"
Option Explicit

' ======================================================
' Module: modMain
' Mo ta: Module chinh dieu khien ung dung va form MainMenu
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 20/05/2025
' ======================================================

' Bien toan cuc de theo doi trang thai
Private blnIsMainMenuShowing As Boolean
' Bien toan cuc MainMenu Sheet Class
Private objMainMenuSheet As clsMainMenuSheet

' Khoi tao class cho sheet MainMenu
Public Sub InitializeMainMenuSheet()
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet MainMenu co ton tai khong
    If Not sheetExists(SHEET_MAIN_MENU) Then Exit Sub
    
    ' Tao instance moi cua class
    If objMainMenuSheet Is Nothing Then
        Set objMainMenuSheet = New clsMainMenuSheet
        
        ' Khoi tao instance voi sheet MainMenu
        objMainMenuSheet.Initialize ThisWorkbook.Worksheets(SHEET_MAIN_MENU)
        
        ' Ghi log
        LogInfo "InitializeMainMenuSheet", "Da khoi tao class cho sheet MainMenu"
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "InitializeMainMenuSheet", Err.Number, Err.Description
End Sub

' Hien thi form MainMenu
Public Sub ShowMainMenuForm()
    On Error GoTo ErrorHandler
    
    ' Ghi log cho m?c dích g? l?i
    LogInfo "ShowMainMenuForm", "Bat dau mo form MainMenu"
    
    ' Kiem tra neu form dang duoc hien thi
    If blnIsMainMenuShowing Then
        Exit Sub
    End If
    
    ' Dat co hieu form dang hien thi
    blnIsMainMenuShowing = True
    
    ' Kiem tra du lieu da duoc import chua
    If Not IsDataComplete() Then
        MsgBox MSG_DATA_INCOMPLETE_FOR_MAIN, vbExclamation, TITLE_WARNING
        ShowImportForm
        blnIsMainMenuShowing = False
        Exit Sub
    End If
    
    ' Dam bao form cu da bi unload (neu ton tai)
    On Error Resume Next
    Unload frmMainMenu
    On Error GoTo ErrorHandler
    
    ' Load form moi
    Load frmMainMenu
    
    ' Cap nhat thong tin tren form
    RefreshMainMenuInfo
    
    ' Hien thi form
    frmMainMenu.Show vbModeless
    
    ' Ghi log thanh cong
    LogInfo "ShowMainMenuForm", "Da hien thi form MainMenu thanh cong"
    
    Exit Sub
    
ErrorHandler:
    blnIsMainMenuShowing = False
    LogError "ShowMainMenuForm", Err.Number, Err.Description
    MsgBox "Loi khi mo form Main Menu: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Cap nhat thong tin tren form MainMenu
Public Sub RefreshMainMenuInfo()
    On Error GoTo ErrorHandler
    
    ' Kiem tra form da duoc load chua
    If frmMainMenu Is Nothing Then Exit Sub
    
    ' Cap nhat thong tin import
    frmMainMenu.lblDuNoInfo.Caption = "Du no: " & GetLastImportInfo(SHEET_DU_NO)
    frmMainMenu.lblTaiSanInfo.Caption = "Tai san: " & GetLastImportInfo(SHEET_TAI_SAN)
    frmMainMenu.lblTraGocInfo.Caption = "Tra goc: " & GetLastImportInfo(SHEET_TRA_GOC)
    frmMainMenu.lblTraLaiInfo.Caption = "Tra lai: " & GetLastImportInfo(SHEET_TRA_LAI)
    
    ' Cap nhat thong tin so luong ban ghi
    frmMainMenu.lblDuNoCount.Caption = "So ban ghi: " & CountRecords(SHEET_DU_NO)
    frmMainMenu.lblTaiSanCount.Caption = "So ban ghi: " & CountRecords(SHEET_TAI_SAN)
    frmMainMenu.lblTraGocCount.Caption = "So ban ghi: " & CountRecords(SHEET_TRA_GOC)
    frmMainMenu.lblTraLaiCount.Caption = "So ban ghi: " & CountRecords(SHEET_TRA_LAI)
    
    ' Cap nhat trang thai he thong
    frmMainMenu.lblStatus.Caption = "He thong san sang! (Cap nhat: " & Format(Now, DATE_TIME_FORMAT) & ")"
    
    Exit Sub
    
ErrorHandler:
    LogError "RefreshMainMenuInfo", Err.Number, Err.Description
End Sub

' Tao va cap nhat sheet MainMenu
Public Sub CreateMainMenuSheet()
    On Error GoTo ErrorHandler
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Tao hoac lay sheet MainMenu
    Dim wsMainMenu As Worksheet
    
    If sheetExists(SHEET_MAIN_MENU) Then
        Set wsMainMenu = ThisWorkbook.Worksheets(SHEET_MAIN_MENU)
        wsMainMenu.Cells.Clear
    Else
        Set wsMainMenu = ThisWorkbook.Worksheets.Add
        wsMainMenu.Name = SHEET_MAIN_MENU
    End If
    
    ' Dinh dang sheet
    wsMainMenu.Tab.Color = AGRIBANK_RED_COLOR
    
    ' Tao tieu de
    With wsMainMenu.Range("A1")
        .Value = "HE THONG QUAN LY THONG TIN KHACH HANG VAY"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = AGRIBANK_RED_COLOR
    End With
    
    ' Them logo (su dung hinh chu nhat voi mau do Agribank)
    With wsMainMenu.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 30)
        .Fill.ForeColor.RGB = AGRIBANK_RED_COLOR
        .Line.Visible = msoFalse
    End With
    
    ' Thong tin import
    wsMainMenu.Range("A3").Value = "THONG TIN IMPORT DU LIEU:"
    wsMainMenu.Range("A3").Font.Bold = True
    
    wsMainMenu.Range("A4").Value = "Du no:"
    wsMainMenu.Range("B4").Value = GetLastImportInfo(SHEET_DU_NO)
    
    wsMainMenu.Range("A5").Value = "Tai san:"
    wsMainMenu.Range("B5").Value = GetLastImportInfo(SHEET_TAI_SAN)
    
    wsMainMenu.Range("A6").Value = "Tra goc:"
    wsMainMenu.Range("B6").Value = GetLastImportInfo(SHEET_TRA_GOC)
    
    wsMainMenu.Range("A7").Value = "Tra lai:"
    wsMainMenu.Range("B7").Value = GetLastImportInfo(SHEET_TRA_LAI)
    
    ' Tong quan so ban ghi
    wsMainMenu.Range("C3").Value = "SO LUONG BAN GHI:"
    wsMainMenu.Range("C3").Font.Bold = True
    
    wsMainMenu.Range("C4").Value = CountRecords(SHEET_DU_NO)
    wsMainMenu.Range("C5").Value = CountRecords(SHEET_TAI_SAN)
    wsMainMenu.Range("C6").Value = CountRecords(SHEET_TRA_GOC)
    wsMainMenu.Range("C7").Value = CountRecords(SHEET_TRA_LAI)
    
    ' Huong dan
    wsMainMenu.Range("A10").Value = "HUONG DAN SU DUNG:"
    wsMainMenu.Range("A10").Font.Bold = True
    
    wsMainMenu.Range("A11").Value = "1. Su dung nut 'Main Menu' duoi day de mo giao dien chinh"
    wsMainMenu.Range("A12").Value = "2. Giao dien chinh se cung cap quyen truy cap den tat ca cac chuc nang"
    wsMainMenu.Range("A13").Value = "3. Su dung nut 'Import Du Lieu' de cap nhat du lieu moi"
    
    ' Tao cac nut chuc nang
    Dim t As Single
    t = 100 ' Khoang cach tu tren xuong
    
    ' Nut Main Menu
    With wsMainMenu.Buttons.Add(100, t, 150, 30)
        .Name = "btnMainMenu"
        .Caption = "Main Menu"
        .OnAction = "modMain.ShowMainMenuForm"
    End With
    
    ' Nut Import Du Lieu
    With wsMainMenu.Buttons.Add(100, t + 40, 150, 30)
        .Name = "btnImportData"
        .Caption = "Import Du Lieu"
        .OnAction = "ShowImportForm"
    End With
    
    ' Dinh dang tong the
    wsMainMenu.Columns("A:D").AutoFit
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    LogError "CreateMainMenuSheet", Err.Number, Err.Description
    MsgBox "Loi khi tao sheet MainMenu: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Dieu huong den cac chuc nang
Public Sub NavigateToFunction(ByVal functionName As String)
    On Error GoTo ErrorHandler
    
    ' An form MainMenu (neu can thiet)
    If Not frmMainMenu Is Nothing Then
        frmMainMenu.Hide
    End If
    
    ' Dieu huong den chuc nang duoc chon
    Select Case functionName
        Case FUNC_KHACH_HANG
            ' Mo form Khach hang
            ' ShowKhachHangForm
            MsgBox "Chuc nang Tim kiem khach hang dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_KHOAN_VAY
            ' Mo form Khoan vay
            ' ShowKhoanVayForm
            MsgBox "Chuc nang Quan ly khoan vay dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_TAI_SAN
            ' Mo form Tai san
            ' ShowTaiSanForm
            MsgBox "Chuc nang Quan ly tai san dam bao dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_TRA_NO
            ' Mo form Tra no
            ' ShowTraNoForm
            MsgBox "Chuc nang Theo doi tra no dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_CANH_BAO
            ' Mo form Canh bao
            ' ShowCanhBaoForm
            MsgBox "Chuc nang Canh bao rui ro dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_BAO_CAO
            ' Mo form Bao cao
            ' ShowBaoCaoForm
            MsgBox "Chuc nang Bao cao & Thong ke dang duoc phat trien", vbInformation, TITLE_INFORMATION
            
        Case FUNC_IMPORT_DATA
            ' Mo form Import du lieu
            ShowImportForm
            
        Case Else
            ' Hien lai form MainMenu
            If Not frmMainMenu Is Nothing Then
                frmMainMenu.Show vbModeless
            End If
    End Select
    
    Exit Sub
    
ErrorHandler:
    LogError "NavigateToFunction", Err.Number, Err.Description
    
    ' Hien lai form MainMenu neu co loi
    If Not frmMainMenu Is Nothing Then
        frmMainMenu.Show vbModeless
    End If
End Sub

' Xu ly su kien khi dong form MainMenu
Public Sub OnMainMenuFormClosed()
    On Error Resume Next
    
    ' Dat co hieu form khong con hien thi
    blnIsMainMenuShowing = False
    
    ' Hien thi sheet MainMenu (neu co)
    If sheetExists(SHEET_MAIN_MENU) Then
        ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    End If
    
    ' Ghi log
    LogInfo "OnMainMenuFormClosed", "Form MainMenu da duoc dong"
End Sub

' Ham kiem tra trang thai san sang cua he thong
Public Function IsSystemReady() As Boolean
    On Error GoTo ErrorHandler
    
    ' Kiem tra du lieu da duoc import du
    If Not IsDataComplete() Then
        IsSystemReady = False
        Exit Function
    End If
    
    ' Kiem tra cac sheet du lieu co hop le
    If Not ValidateImportedData(SHEET_DU_NO) Or _
       Not ValidateImportedData(SHEET_TAI_SAN) Or _
       Not ValidateImportedData(SHEET_TRA_GOC) Or _
       Not ValidateImportedData(SHEET_TRA_LAI) Then
        IsSystemReady = False
        Exit Function
    End If
    
    ' Kiem tra sheet MainMenu ton tai
    If Not sheetExists(SHEET_MAIN_MENU) Then
        CreateMainMenuSheet
    End If
    
    ' He thong san sang
    IsSystemReady = True
    
    Exit Function
    
ErrorHandler:
    LogError "IsSystemReady", Err.Number, Err.Description
    IsSystemReady = False
End Function

' Cap nhat giao dien cua MainMenu sau khi import du lieu
Public Sub UpdateMainMenuAfterImport()
    On Error GoTo ErrorHandler
    
    ' Kiem tra du lieu da du
    If Not IsDataComplete() Then Exit Sub
    
    ' Cap nhat sheet MainMenu
    CreateMainMenuSheet
    
    ' Cap nhat thong tin tren form MainMenu (neu dang mo)
    RefreshMainMenuInfo
    
    ' Kich hoat sheet MainMenu
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Khoi tao class cho sheet MainMenu
    InitializeMainMenuSheet
    
    Exit Sub
    
ErrorHandler:
    LogError "UpdateMainMenuAfterImport", Err.Number, Err.Description
End Sub

' Khoi dong ung dung va hien thi MainMenu
Public Sub StartApplication()
    On Error GoTo ErrorHandler
    
    ' Khoi tao class cho sheet MainMenu
    InitializeMainMenuSheet
    
    ' Kiem tra trang thai san sang cua he thong
    If IsSystemReady() Then
        ' Hien thi form MainMenu
        ShowMainMenuForm
    Else
        ' Hien thi form Import neu chua co du lieu
        ShowImportForm
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "StartApplication", Err.Number, Err.Description
    MsgBox "Loi khi khoi dong ung dung: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Reset trang thai MainMenu
Public Sub ResetMainMenuStatus()
    On Error Resume Next
    blnIsMainMenuShowing = False
End Sub
