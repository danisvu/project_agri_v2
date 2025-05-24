Attribute VB_Name = "modTest"
Option Explicit

' ======================================================
' Module: modTest
' Mo ta: Chua cac thu tuc kiem thu he thong
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 18/05/2025
' ======================================================

' Kiem thu quy trinh import du lieu
Public Sub TestImportProcess()
    On Error GoTo ErrorHandler
    
    ' Thong bao bat dau test
    MsgBox "Bat dau kiem thu quy trinh import du lieu." & vbNewLine & _
           "Huong dan:" & vbNewLine & _
           "1. Form Import se duoc hien thi." & vbNewLine & _
           "2. Ban can chon file du lieu tuong ung cho tung loai." & vbNewLine & _
           "3. Sau khi import du 4 loai du lieu, nhan 'Tiep tuc'." & vbNewLine & _
           "4. Ket qua test se duoc ghi vao sheet Log.", vbInformation, "Test Import"
    
    ' Ghi log bat dau test
    LogInfo "TestImportProcess", "Bat dau kiem thu quy trinh import du lieu"
    
    ' Tao ban sao du lieu truoc khi test
    BackupCurrentData
    
    ' Tao sheet Log neu chua ton tai
    CreateSheetIfNotExists SHEET_LOG
    
    ' Hien thi form Import
    ShowImportForm
    
    ' Tham khao: Sau khi import, quy trinh se tu dong tiep tuc
    
    Exit Sub
    
ErrorHandler:
    LogError "TestImportProcess", Err.Number, Err.Description
    MsgBox "Loi khi kiem thu import: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Kiem tra he thong va form MainMenu
Public Sub TestMainMenuForm()
    On Error GoTo ErrorHandler
    
    ' Ghi log
    LogInfo "TestMainMenuForm", "Bat dau kiem tra form MainMenu"
    
    ' Kiem tra du lieu
    If Not IsDataComplete() Then
        MsgBox "Chua co du du lieu. Hay import day du du lieu truoc!", vbExclamation, "Test MainMenu"
        Exit Sub
    End If
    
    ' Tao sheet MainMenu neu chua ton tai
    If Not sheetExists(SHEET_MAIN_MENU) Then
        modMain.CreateMainMenuSheet
    End If
    
    ' Kich hoat sheet MainMenu
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Cho 1 giay de dam bao sheet da duoc hien thi
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Hien thi form MainMenu
    modMain.ShowMainMenuForm
    
    ' Ghi log ket thuc
    LogInfo "TestMainMenuForm", "Ket thuc kiem tra form MainMenu"
    
    Exit Sub
    
ErrorHandler:
    LogError "TestMainMenuForm", Err.Number, Err.Description
    MsgBox "Loi khi kiem tra form MainMenu: " & Err.Description, vbExclamation, "Test MainMenu"
End Sub
' Kiem thu tinh toan ven du lieu
Public Sub TestDataIntegrity()
    On Error GoTo ErrorHandler
    
    ' Kiem tra du lieu da du chua
    If Not IsDataComplete() Then
        MsgBox "Chua co du du lieu de kiem tra tinh toan ven. Vui long import du lieu truoc.", _
               vbExclamation, "Test Data Integrity"
        Exit Sub
    End If
    
    ' Thong bao bat dau test
    MsgBox "Bat dau kiem tra tinh toan ven du lieu...", vbInformation, "Test Data Integrity"
    
    ' Ghi log bat dau test
    LogInfo "TestDataIntegrity", "Bat dau kiem tra tinh toan ven du lieu"
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Kiem tra tung sheet
    Dim testResult As String
    testResult = ""
    
    ' Du no
    If ValidateImportedData(SHEET_DU_NO) Then
        testResult = testResult & "- Du no: OK" & vbNewLine
    Else
        testResult = testResult & "- Du no: FAILED" & vbNewLine
    End If
    
    ' Tai san
    If ValidateImportedData(SHEET_TAI_SAN) Then
        testResult = testResult & "- Tai san: OK" & vbNewLine
    Else
        testResult = testResult & "- Tai san: FAILED" & vbNewLine
    End If
    
    ' Tra goc
    If ValidateImportedData(SHEET_TRA_GOC) Then
        testResult = testResult & "- Tra goc: OK" & vbNewLine
    Else
        testResult = testResult & "- Tra goc: FAILED" & vbNewLine
    End If
    
    ' Tra lai
    If ValidateImportedData(SHEET_TRA_LAI) Then
        testResult = testResult & "- Tra lai: OK" & vbNewLine
    Else
        testResult = testResult & "- Tra lai: FAILED" & vbNewLine
    End If
    
    ' Tao tong quan so ban ghi
    testResult = testResult & vbNewLine & "So luong ban ghi:" & vbNewLine
    testResult = testResult & "- Du no: " & CountRecords(SHEET_DU_NO) & vbNewLine
    testResult = testResult & "- Tai san: " & CountRecords(SHEET_TAI_SAN) & vbNewLine
    testResult = testResult & "- Tra goc: " & CountRecords(SHEET_TRA_GOC) & vbNewLine
    testResult = testResult & "- Tra lai: " & CountRecords(SHEET_TRA_LAI) & vbNewLine
    
    ' Ghi log ket qua test
    LogInfo "TestDataIntegrity", "Ket qua kiem tra: " & vbNewLine & testResult
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Hien thi ket qua
    MsgBox "Ket qua kiem tra tinh toan ven du lieu:" & vbNewLine & vbNewLine & _
           testResult, vbInformation, "Test Data Integrity"
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Ghi log loi
    LogError "TestDataIntegrity", Err.Number, Err.Description
    MsgBox "Loi khi kiem tra tinh toan ven du lieu: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Tao du lieu mau de test
Public Sub GenerateSampleData()
    On Error GoTo ErrorHandler
    
    ' Xac nhan nguoi dung
    If MsgBox("Chuc nang nay se tao du lieu mau de test, ghi de len du lieu hien tai. Ban co chac chan muon tiep tuc?", _
              vbQuestion + vbYesNo, "Tao du lieu mau") = vbNo Then
        Exit Sub
    End If
    
    ' Thong bao bat dau
    MsgBox "Bat dau tao du lieu mau...", vbInformation, "Tao du lieu mau"
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Dang tao du lieu mau..."
    
    ' Backup du lieu hien tai
    BackupCurrentData
    
    ' Tao du lieu mau cho Du no
    Application.StatusBar = "Dang tao du lieu mau cho Du no..."
    Dim wsDuNo As Worksheet
    
    If sheetExists(SHEET_DU_NO) Then
        Set wsDuNo = ThisWorkbook.Worksheets(SHEET_DU_NO)
        wsDuNo.Cells.Clear
    Else
        Set wsDuNo = ThisWorkbook.Worksheets.Add
        wsDuNo.Name = SHEET_DU_NO
    End If
    
    ' Ghi thong tin ngay gio import vao dong dau tien
    wsDuNo.Cells(INFO_ROW, 1).Value = IMPORT_INFO_PREFIX & Now()
    wsDuNo.Cells(INFO_ROW, 1).Font.Bold = True
    
    ' Tao header
    wsDuNo.Cells(2, 1).Value = "custseq"
    wsDuNo.Cells(2, 2).Value = "custnm"
    wsDuNo.Cells(2, 3).Value = "apprseq"
    wsDuNo.Cells(2, 4).Value = "dsbsdt"
    wsDuNo.Cells(2, 5).Value = "dsbsmatdt"
    wsDuNo.Cells(2, 6).Value = "dsbsamt"
    wsDuNo.Cells(2, 7).Value = "dsbsbal"
    
    ' Tao du lieu mau
    Dim i As Long
    For i = 3 To 10
        wsDuNo.Cells(i, 1).Value = "KH" & Format(i - 2, "000")
        wsDuNo.Cells(i, 2).Value = "Khach hang " & (i - 2)
        wsDuNo.Cells(i, 3).Value = "KV" & Format(i - 2, "000")
        wsDuNo.Cells(i, 4).Value = DateAdd("d", -(i - 3) * 30, Date)
        wsDuNo.Cells(i, 5).Value = DateAdd("d", 365, DateAdd("d", -(i - 3) * 30, Date))
        wsDuNo.Cells(i, 6).Value = (i - 2) * 100000000
        wsDuNo.Cells(i, 7).Value = (i - 2) * 90000000
    Next i
    
    ' Dinh dang
    FormatWorksheet wsDuNo
    
    ' Tao du lieu mau cho Tai san
    Application.StatusBar = "Dang tao du lieu mau cho Tai san..."
    Dim wsTaiSan As Worksheet
    
    If sheetExists(SHEET_TAI_SAN) Then
        Set wsTaiSan = ThisWorkbook.Worksheets(SHEET_TAI_SAN)
        wsTaiSan.Cells.Clear
    Else
        Set wsTaiSan = ThisWorkbook.Worksheets.Add
        wsTaiSan.Name = SHEET_TAI_SAN
    End If
    
    ' Ghi thong tin ngay gio import vao dong dau tien
    wsTaiSan.Cells(INFO_ROW, 1).Value = IMPORT_INFO_PREFIX & Now()
    wsTaiSan.Cells(INFO_ROW, 1).Font.Bold = True
    
    ' Tao header
    wsTaiSan.Cells(2, 1).Value = "clno"
    wsTaiSan.Cells(2, 2).Value = "clcustno"
    wsTaiSan.Cells(2, 3).Value = "clcustnm"
    wsTaiSan.Cells(2, 4).Value = "cltpcd"
    wsTaiSan.Cells(2, 5).Value = "cldtltpcd"
    wsTaiSan.Cells(2, 6).Value = "clamt"
    
    ' Tao du lieu mau
    For i = 3 To 10
        wsTaiSan.Cells(i, 1).Value = "TS" & Format(i - 2, "000")
        wsTaiSan.Cells(i, 2).Value = "KH" & Format(i - 2, "000")
        wsTaiSan.Cells(i, 3).Value = "Khach hang " & (i - 2)
        
        If i Mod 2 = 0 Then
            wsTaiSan.Cells(i, 4).Value = "Bat dong san"
            wsTaiSan.Cells(i, 5).Value = "Quyen su dung dat"
        Else
            wsTaiSan.Cells(i, 4).Value = "Dong san"
            wsTaiSan.Cells(i, 5).Value = "Phuong tien van tai"
        End If
        
        wsTaiSan.Cells(i, 6).Value = (i - 2) * 150000000
    Next i
    
    ' Dinh dang
    FormatWorksheet wsTaiSan
    
    ' Tao du lieu mau cho Tra goc
    Application.StatusBar = "Dang tao du lieu mau cho Tra goc..."
    Dim wsTraGoc As Worksheet
    
    If sheetExists(SHEET_TRA_GOC) Then
        Set wsTraGoc = ThisWorkbook.Worksheets(SHEET_TRA_GOC)
        wsTraGoc.Cells.Clear
    Else
        Set wsTraGoc = ThisWorkbook.Worksheets.Add
        wsTraGoc.Name = SHEET_TRA_GOC
    End If
    
    ' Ghi thong tin ngay gio import vao dong dau tien
    wsTraGoc.Cells(INFO_ROW, 1).Value = IMPORT_INFO_PREFIX & Now()
    wsTraGoc.Cells(INFO_ROW, 1).Font.Bold = True
    
    ' Tao header
    wsTraGoc.Cells(2, 1).Value = "matdt"
    wsTraGoc.Cells(2, 2).Value = "custseqno"
    wsTraGoc.Cells(2, 3).Value = "custnm"
    wsTraGoc.Cells(2, 4).Value = "amt"
    wsTraGoc.Cells(2, 5).Value = "refrno"
    wsTraGoc.Cells(2, 6).Value = "processed"
    
    ' Tao du lieu mau
    For i = 3 To 18
        wsTraGoc.Cells(i, 1).Value = DateAdd("d", (i - 3) * 30, Date)
        wsTraGoc.Cells(i, 2).Value = "KH" & Format(((i - 3) Mod 8) + 1, "000")
        wsTraGoc.Cells(i, 3).Value = "Khach hang " & (((i - 3) Mod 8) + 1)
        wsTraGoc.Cells(i, 4).Value = ((i - 3) Mod 8 + 1) * 10000000
        wsTraGoc.Cells(i, 5).Value = "KV" & Format(((i - 3) Mod 8) + 1, "000")
        wsTraGoc.Cells(i, 6).Value = "N"
    Next i
    
    ' Dinh dang
    FormatWorksheet wsTraGoc
    
    ' Tao du lieu mau cho Tra lai
    Application.StatusBar = "Dang tao du lieu mau cho Tra lai..."
    Dim wsTraLai As Worksheet
    
    If sheetExists(SHEET_TRA_LAI) Then
        Set wsTraLai = ThisWorkbook.Worksheets(SHEET_TRA_LAI)
        wsTraLai.Cells.Clear
    Else
        Set wsTraLai = ThisWorkbook.Worksheets.Add
        wsTraLai.Name = SHEET_TRA_LAI
    End If
    
    ' Ghi thong tin ngay gio import vao dong dau tien
    wsTraLai.Cells(INFO_ROW, 1).Value = IMPORT_INFO_PREFIX & Now()
    wsTraLai.Cells(INFO_ROW, 1).Font.Bold = True
    
    ' Tao header
    wsTraLai.Cells(2, 1).Value = "matdt"
    wsTraLai.Cells(2, 2).Value = "custseqno"
    wsTraLai.Cells(2, 3).Value = "custnm"
    wsTraLai.Cells(2, 4).Value = "amt"
    wsTraLai.Cells(2, 5).Value = "refrno"
    wsTraLai.Cells(2, 6).Value = "processed"
    
    ' Tao du lieu mau
    For i = 3 To 18
        wsTraLai.Cells(i, 1).Value = DateAdd("d", (i - 3) * 30, Date)
        wsTraLai.Cells(i, 2).Value = "KH" & Format(((i - 3) Mod 8) + 1, "000")
        wsTraLai.Cells(i, 3).Value = "Khach hang " & (((i - 3) Mod 8) + 1)
        wsTraLai.Cells(i, 4).Value = ((i - 3) Mod 8 + 1) * 1500000
        wsTraLai.Cells(i, 5).Value = "KV" & Format(((i - 3) Mod 8) + 1, "000")
        wsTraLai.Cells(i, 6).Value = "N"
    Next i
    
    ' Dinh dang
    FormatWorksheet wsTraLai
    
    ' Tao tong quan du lieu
    Application.StatusBar = "Dang tao tong quan du lieu..."
    GenerateDataSummary
    
    ' Cap nhat hoac tao sheet MainMenu
    Application.StatusBar = "Dang tao giao dien chinh..."
    CreateMainMenuSheet
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Hien thi sheet MainMenu
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Thong bao thanh cong
    MsgBox "Da tao du lieu mau thanh cong!", vbInformation, "Tao du lieu mau"
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Ghi log loi
    LogError "GenerateSampleData", Err.Number, Err.Description
    MsgBox "Loi khi tao du lieu mau: " & Err.Description, vbExclamation, TITLE_ERROR
End Sub

' Test khoi dong ung dung
Public Sub TestApplicationStartup()
    On Error GoTo ErrorHandler
    
    ' Ghi log
    LogInfo "TestApplicationStartup", "Bat dau test khoi dong ung dung"
    
    ' Goi ham khoi dong
    modImport.InitializeApplication
    
    Exit Sub
    
ErrorHandler:
    LogError "TestApplicationStartup", Err.Number, Err.Description
    MsgBox "Loi khi test khoi dong: " & Err.Description, vbExclamation, "Test Startup"
End Sub

' Test sheet MainMenu va kich hoat su kien
Public Sub TestMainMenuSheet()
    On Error GoTo ErrorHandler
    
    ' Ghi log
    LogInfo "TestMainMenuSheet", "Bat dau test sheet MainMenu"
    
    ' Kiem tra du lieu
    If Not IsDataComplete() Then
        MsgBox "Chua co du du lieu. Hay import day du du lieu truoc!", vbExclamation, "Test MainMenu"
        Exit Sub
    End If
    
    ' Tao sheet MainMenu neu chua ton tai
    If Not sheetExists(SHEET_MAIN_MENU) Then
        modMain.CreateMainMenuSheet
    End If
    
    ' Thong bao
    MsgBox "Sap kich hoat sheet khac, sau do se chuyen ve sheet MainMenu. " & _
           "Form MainMenu se tu dong hien thi khi chuyen den sheet MainMenu.", _
           vbInformation, "Test MainMenu"
    
    ' Reset bien trang thai trong modMain neu co
    On Error Resume Next
    Application.Run "modMain.ResetMainMenuStatus"
    On Error GoTo ErrorHandler
    
    ' Kich hoat sheet khac truoc
    If sheetExists(SHEET_CONFIG) Then
        ThisWorkbook.Worksheets(SHEET_CONFIG).Activate
    ElseIf sheetExists(SHEET_LOG) Then
        ThisWorkbook.Worksheets(SHEET_LOG).Activate
    End If
    
    ' Cho 1 giay de dam bao sheet da duoc hien thi
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Kich hoat sheet MainMenu - su kien nay se tu dong goi ShowMainMenuForm
    ThisWorkbook.Worksheets(SHEET_MAIN_MENU).Activate
    
    ' Ghi log ket thuc
    LogInfo "TestMainMenuSheet", "Ket thuc kiem tra form MainMenu"
    
    Exit Sub
    
ErrorHandler:
    LogError "TestMainMenuSheet", Err.Number, Err.Description
    MsgBox "Loi khi kiem tra form MainMenu: " & Err.Description, vbExclamation, "Test MainMenu"
End Sub

' Ham export tat ca code VBA trong mot workbook ra thu muc
' Can tham chieu Microsoft Visual Basic for Applications Extensibility 5.3
Public Sub ExportVBACode()
    On Error GoTo ErrorHandler
    
    Dim VBProject As Object
    Dim VBComponent As Object
    Dim exportPath As String
    Dim fileName As String
    Dim componentCount As Long
    Dim i As Long
    Dim fileExtension As String
    
    ' Kiem tra tham chieu: Tools > References > Microsoft Visual Basic for Applications Extensibility
    ' Neu khong co se bao loi
    Set VBProject = ThisWorkbook.VBProject
    
    ' Tat thong bao cap nhat
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Tao thu muc export
    exportPath = ThisWorkbook.Path & "\VBA_Export_" & Format(Now, "yyyymmdd_hhmmss") & "\"
    
    ' Tao thu muc neu chua ton tai
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Tao thu muc con cho tung loai component
    MkDir exportPath & "Modules\"
    MkDir exportPath & "ClassModules\"
    MkDir exportPath & "Forms\"
    MkDir exportPath & "Others\"
    
    componentCount = VBProject.VBComponents.Count
    i = 0
    
    ' Duyet qua tung component va export
    For Each VBComponent In VBProject.VBComponents
        i = i + 1
        Application.StatusBar = "Dang export... " & i & "/" & componentCount & ": " & VBComponent.Name
        
        ' Xac dinh loai file extension dua vao loai component
        Select Case VBComponent.Type
            Case 1 ' Module
                fileExtension = ".bas"
                fileName = exportPath & "Modules\" & VBComponent.Name & fileExtension
                
            Case 2 ' Class Module
                fileExtension = ".cls"
                fileName = exportPath & "ClassModules\" & VBComponent.Name & fileExtension
                
            Case 3 ' Form
                fileExtension = ".frm"
                fileName = exportPath & "Forms\" & VBComponent.Name & fileExtension
                
            Case Else ' Document, etc.
                fileExtension = ".cls"
                fileName = exportPath & "Others\" & VBComponent.Name & fileExtension
        End Select
        
        ' Export component
        VBComponent.Export fileName
    Next VBComponent
    
    ' Tao file thong tin tong quat
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim infoFile As Object
    Set infoFile = fso.CreateTextFile(exportPath & "ExportInfo.txt", True)
    
    ' Ghi thong tin
    infoFile.WriteLine "THONG TIN EXPORT CODE VBA"
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Ten workbook: " & ThisWorkbook.Name
    infoFile.WriteLine "Ngay export: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    infoFile.WriteLine "So luong component: " & componentCount
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Danh sach component da export:"
    
    ' Liet ke component da export
    i = 0
    For Each VBComponent In VBProject.VBComponents
        i = i + 1
        infoFile.WriteLine i & ". " & VBComponent.Name & " (" & GetComponentTypeName(VBComponent.Type) & ")"
    Next VBComponent
    
    infoFile.Close
    Set infoFile = Nothing
    Set fso = Nothing
    
CleanUp:
    ' Khoi phuc thong bao
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Thong bao ket qua
    MsgBox "Da export thanh cong " & componentCount & " component VBA vao thu muc:" & vbCrLf & _
           exportPath, vbInformation, "Export VBA Code"
    
    ' Mo thu muc chua code da export
    Shell "explorer.exe """ & exportPath & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    
    Select Case Err.Number
        Case 32813
            errorMsg = "Ban chua bat References 'Microsoft Visual Basic for Applications Extensibility'." & vbCrLf & _
                       "Hay vao Tools > References trong VBA Editor va chon reference nay."
        Case 9
            errorMsg = "Khong tim thay component."
        Case 48, 76
            errorMsg = "Loi khi tao thu muc. Kiem tra lai quyen truy cap thu muc."
        Case 1004
            errorMsg = "Loi VBA Project bi khoa. Hay mo khoa VBA Project truoc khi export." & vbCrLf & _
                       "(Tools > VBAProject Properties > Protection tab)"
        Case Else
            errorMsg = "Loi: " & Err.Number & " - " & Err.Description
    End Select
    
    MsgBox errorMsg, vbCritical, "Loi Export VBA Code"
    Resume CleanUp
End Sub

' Ham phu: Lay ten loai component
Private Function GetComponentTypeName(typeNum As Integer) As String
    Select Case typeNum
        Case 1
            GetComponentTypeName = "Module"
        Case 2
            GetComponentTypeName = "Class Module"
        Case 3
            GetComponentTypeName = "UserForm"
        Case 100
            GetComponentTypeName = "Document Module"
        Case Else
            GetComponentTypeName = "Khac (" & typeNum & ")"
    End Select
End Function

' Ham export tat ca code VBA trong mot workbook ra thu muc, tao ban sao luu file Excel
' va tao thu muc General chua tat ca cac file
' Can tham chieu Microsoft Visual Basic for Applications Extensibility 5.3
Public Sub ExportVBACodeAndBackup()
    On Error GoTo ErrorHandler
    
    Dim VBProject As Object
    Dim VBComponent As Object
    Dim exportPath As String
    Dim fileName As String
    Dim generalFileName As String
    Dim componentCount As Long
    Dim i As Long
    Dim fileExtension As String
    Dim backupFileName As String
    Dim timeStamp As String
    Dim fso As Object
    
    ' Tao doi tuong FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Kiem tra tham chieu: Tools > References > Microsoft Visual Basic for Applications Extensibility
    ' Neu khong co se bao loi
    Set VBProject = ThisWorkbook.VBProject
    
    ' Tat thong bao cap nhat
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Tao timestamp cho ten thu muc va file backup
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    
    ' Tao thu muc export
    exportPath = ThisWorkbook.Path & "\VBA_Export_" & timeStamp & "\"
    
    ' Tao thu muc neu chua ton tai
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Tao thu muc con cho tung loai component
    MkDir exportPath & "Modules\"
    MkDir exportPath & "ClassModules\"
    MkDir exportPath & "Forms\"
    MkDir exportPath & "Others\"
    MkDir exportPath & "Backup\"
    MkDir exportPath & "General\"  ' Them thu muc General chua tat ca cac file
    
    ' TAO BAN SAO LUU CUA FILE EXCEL
    backupFileName = exportPath & "Backup\" & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & _
                    "_Backup_" & timeStamp & Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    
    ' Luu ban sao cua file Excel
    ThisWorkbook.SaveCopyAs backupFileName
    
    componentCount = VBProject.VBComponents.Count
    i = 0
    
    ' Duyet qua tung component va export
    For Each VBComponent In VBProject.VBComponents
        i = i + 1
        Application.StatusBar = "Dang export... " & i & "/" & componentCount & ": " & VBComponent.Name
        
        ' Xac dinh loai file extension dua vao loai component
        Select Case VBComponent.Type
            Case 1 ' Module
                fileExtension = ".bas"
                fileName = exportPath & "Modules\" & VBComponent.Name & fileExtension
                
            Case 2 ' Class Module
                fileExtension = ".cls"
                fileName = exportPath & "ClassModules\" & VBComponent.Name & fileExtension
                
            Case 3 ' Form
                fileExtension = ".frm"
                fileName = exportPath & "Forms\" & VBComponent.Name & fileExtension
                
            Case Else ' Document, etc.
                fileExtension = ".cls"
                fileName = exportPath & "Others\" & VBComponent.Name & fileExtension
        End Select
        
        ' Export component
        VBComponent.Export fileName
        
        ' Tao ban sao file vao thu muc General
        generalFileName = exportPath & "General\" & VBComponent.Name & fileExtension
        
        ' Kiem tra neu file dich da ton tai thi xoa truoc
        If fso.FileExists(generalFileName) Then
            fso.DeleteFile generalFileName, True
        End If
        
        ' Sao chep file vua export vao thu muc General
        fso.CopyFile fileName, generalFileName
    Next VBComponent
    
    ' Tao file thong tin tong quat
    Dim infoFile As Object
    Set infoFile = fso.CreateTextFile(exportPath & "ExportInfo.txt", True)
    
    ' Ghi thong tin
    infoFile.WriteLine "THONG TIN EXPORT CODE VBA"
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Ten workbook: " & ThisWorkbook.Name
    infoFile.WriteLine "Ngay export: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    infoFile.WriteLine "So luong component: " & componentCount
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Duong dan file backup: " & backupFileName
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Cau truc thu muc:"
    infoFile.WriteLine "- Modules: Chua cac module thuong (.bas)"
    infoFile.WriteLine "- ClassModules: Chua cac class module (.cls)"
    infoFile.WriteLine "- Forms: Chua cac user form (.frm)"
    infoFile.WriteLine "- Others: Chua cac loai module khac"
    infoFile.WriteLine "- General: Chua tat ca cac file tu cac thu muc tren"
    infoFile.WriteLine "- Backup: Chua ban sao luu cua file Excel"
    infoFile.WriteLine "-------------------------"
    infoFile.WriteLine "Danh sach component da export:"
    
    ' Liet ke component da export
    i = 0
    For Each VBComponent In VBProject.VBComponents
        i = i + 1
        infoFile.WriteLine i & ". " & VBComponent.Name & " (" & GetComponentTypeName(VBComponent.Type) & ")"
    Next VBComponent
    
    infoFile.Close
    Set infoFile = Nothing
    Set fso = Nothing
    
CleanUp:
    ' Khoi phuc thong bao
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Thong bao ket qua
    MsgBox "Da hoan thanh export:" & vbCrLf & _
           "- Export " & componentCount & " component VBA vao thu muc rieng biet" & vbCrLf & _
           "- Tao ban sao cua tat ca component trong thu muc General" & vbCrLf & _
           "- Tao ban sao luu file Excel" & vbCrLf & _
           "Tat ca duoc luu trong thu muc:" & vbCrLf & _
           exportPath, vbInformation, "Export VBA Code & Backup"
    
    ' Mo thu muc chua code da export
    Shell "explorer.exe """ & exportPath & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    
    Select Case Err.Number
        Case 32813
            errorMsg = "Ban chua bat References 'Microsoft Visual Basic for Applications Extensibility'." & vbCrLf & _
                       "Hay vao Tools > References trong VBA Editor va chon reference nay."
        Case 9
            errorMsg = "Khong tim thay component."
        Case 48, 76
            errorMsg = "Loi khi tao thu muc. Kiem tra lai quyen truy cap thu muc."
        Case 58, 52
            errorMsg = "Loi khi sao chep file vao thu muc General. File co the dang duoc su dung."
        Case 1004
            If InStr(Err.Description, "SaveCopyAs") > 0 Then
                errorMsg = "Khong the tao ban sao luu file Excel. Kiem tra quyen truy cap thu muc hoac file co dang mo khong."
            Else
                errorMsg = "Loi VBA Project bi khoa. Hay mo khoa VBA Project truoc khi export." & vbCrLf & _
                           "(Tools > VBAProject Properties > Protection tab)"
            End If
        Case Else
            errorMsg = "Loi: " & Err.Number & " - " & Err.Description
    End Select
    
    MsgBox errorMsg, vbCritical, "Loi Export VBA Code"
    Resume CleanUp
End Sub

