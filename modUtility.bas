Attribute VB_Name = "modUtility"
Option Explicit

' ======================================================
' Module: modUtility
' Mo ta: Chua cac ham tien ich ho tro cho ung dung
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 18/05/2025
' ======================================================

' Kiem tra xem sheet co ton tai hay khong
Public Function sheetExists(ByVal sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Kiem tra xem sheet co ton tai trong workbook hien tai
    sheetExists = False
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit Function
        End If
    Next ws
    
    Exit Function
    
ErrorHandler:
    ' Neu co loi, tra ve False
    sheetExists = False
    LogError "SheetExists", Err.Number, Err.Description
End Function

' Tao sheet moi neu chua ton tai
Public Function CreateSheetIfNotExists(ByVal sheetName As String) As Worksheet
    On Error GoTo ErrorHandler
    
    If Not sheetExists(sheetName) Then
        ' Tao sheet moi
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.Worksheets.Add
        newSheet.Name = sheetName
        Set CreateSheetIfNotExists = newSheet
    Else
        ' Tra ve sheet da ton tai
        Set CreateSheetIfNotExists = ThisWorkbook.Worksheets(sheetName)
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi tao sheet: " & Err.Description, vbExclamation, TITLE_ERROR
    LogError "CreateSheetIfNotExists", Err.Number, Err.Description
    Set CreateSheetIfNotExists = Nothing
End Function

' Lay duong dan file
Public Function GetFilePath(ByVal fileType As String) As String
    On Error GoTo ErrorHandler
    
    ' Khoi tao ham tra ve rong
    GetFilePath = ""
    
    ' Tao file dialog
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Chon file " & fileType
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        
        If .Show = -1 Then
            ' Nguoi dung da chon file
            GetFilePath = .SelectedItems(1)
        Else
            ' Nguoi dung huy
            GetFilePath = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi lay duong dan file: " & Err.Description, vbExclamation, TITLE_ERROR
    LogError "GetFilePath", Err.Number, Err.Description
    GetFilePath = ""
End Function

' Ghi log loi
Public Sub LogError(ByVal procedureName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    
    ' Tao sheet log neu chua ton tai
    Dim wsLog As Worksheet
    Set wsLog = CreateSheetIfNotExists(SHEET_LOG)
    
    ' Tim dong cuoi cung co du lieu
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    
    ' Neu sheet trong, them header
    If lastRow = 0 Then
        wsLog.Cells(1, 1).Value = "Thoi gian"
        wsLog.Cells(1, 2).Value = "Loai"
        wsLog.Cells(1, 3).Value = "Thu tuc"
        wsLog.Cells(1, 4).Value = "Ma loi"
        wsLog.Cells(1, 5).Value = "Mo ta loi"
        
        lastRow = 1
    End If
    
    ' Ghi thong tin loi
    wsLog.Cells(lastRow + 1, 1).Value = Now()
    wsLog.Cells(lastRow + 1, 2).Value = LOG_ERROR_PREFIX
    wsLog.Cells(lastRow + 1, 3).Value = procedureName
    wsLog.Cells(lastRow + 1, 4).Value = errorNumber
    wsLog.Cells(lastRow + 1, 5).Value = errorDescription
    
    ' Dinh dang thoi gian
    wsLog.Cells(lastRow + 1, 1).NumberFormat = DATE_TIME_FORMAT
    
    ' Tu dong dieu chinh do rong cot
    wsLog.Columns("A:E").AutoFit
End Sub

' Ghi log thong tin
Public Sub LogInfo(ByVal procedureName As String, ByVal infoMessage As String)
    On Error Resume Next
    
    ' Tao sheet log neu chua ton tai
    Dim wsLog As Worksheet
    Set wsLog = CreateSheetIfNotExists(SHEET_LOG)
    
    ' Tim dong cuoi cung co du lieu
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    
    ' Neu sheet trong, them header
    If lastRow = 0 Then
        wsLog.Cells(1, 1).Value = "Thoi gian"
        wsLog.Cells(1, 2).Value = "Loai"
        wsLog.Cells(1, 3).Value = "Thu tuc"
        wsLog.Cells(1, 4).Value = "Ma loi"
        wsLog.Cells(1, 5).Value = "Mo ta loi"
        
        lastRow = 1
    End If
    
    ' Ghi thong tin
    wsLog.Cells(lastRow + 1, 1).Value = Now()
    wsLog.Cells(lastRow + 1, 2).Value = LOG_INFO_PREFIX
    wsLog.Cells(lastRow + 1, 3).Value = procedureName
    wsLog.Cells(lastRow + 1, 4).Value = 0
    wsLog.Cells(lastRow + 1, 5).Value = infoMessage
    
    ' Dinh dang thoi gian
    wsLog.Cells(lastRow + 1, 1).NumberFormat = DATE_TIME_FORMAT
    
    ' Tu dong dieu chinh do rong cot
    wsLog.Columns("A:E").AutoFit
End Sub

' Kiem tra file co ton tai
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(filePath) = 0 Then
        FileExists = False
        Exit Function
    End If
    
    FileExists = (Dir(filePath) <> "")
    
    Exit Function
    
ErrorHandler:
    LogError "FileExists", Err.Number, Err.Description
    FileExists = False
End Function

' Lay thong tin import gan nhat cua sheet
Public Function GetLastImportInfo(ByVal sheetName As String) As String
    On Error GoTo ErrorHandler
    
    GetLastImportInfo = ""
    
    ' Kiem tra sheet co ton tai
    If Not sheetExists(sheetName) Then
        Exit Function
    End If
    
    ' Doc thong tin import o hang dau tien
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws.Cells(INFO_ROW, 1).Value <> "" Then
        GetLastImportInfo = ws.Cells(INFO_ROW, 1).Value
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "GetLastImportInfo", Err.Number, Err.Description
    GetLastImportInfo = ""
End Function

' Kiem tra xem du lieu da duoc import du chua
Public Function IsDataComplete() As Boolean
    On Error GoTo ErrorHandler
    
    ' Kiem tra ca 4 sheet du lieu co ton tai
    If Not sheetExists(SHEET_DU_NO) Then
        IsDataComplete = False
        Exit Function
    End If
    
    If Not sheetExists(SHEET_TAI_SAN) Then
        IsDataComplete = False
        Exit Function
    End If
    
    If Not sheetExists(SHEET_TRA_GOC) Then
        IsDataComplete = False
        Exit Function
    End If
    
    If Not sheetExists(SHEET_TRA_LAI) Then
        IsDataComplete = False
        Exit Function
    End If
    
    ' Kiem tra xem co du lieu trong moi sheet
    Dim ws As Worksheet
    
    ' Du no
    Set ws = ThisWorkbook.Worksheets(SHEET_DU_NO)
    If Application.WorksheetFunction.CountA(ws.UsedRange) <= 1 Then
        IsDataComplete = False
        Exit Function
    End If
    
    ' Tai san
    Set ws = ThisWorkbook.Worksheets(SHEET_TAI_SAN)
    If Application.WorksheetFunction.CountA(ws.UsedRange) <= 1 Then
        IsDataComplete = False
        Exit Function
    End If
    
    ' Tra goc
    Set ws = ThisWorkbook.Worksheets(SHEET_TRA_GOC)
    If Application.WorksheetFunction.CountA(ws.UsedRange) <= 1 Then
        IsDataComplete = False
        Exit Function
    End If
    
    ' Tra lai
    Set ws = ThisWorkbook.Worksheets(SHEET_TRA_LAI)
    If Application.WorksheetFunction.CountA(ws.UsedRange) <= 1 Then
        IsDataComplete = False
        Exit Function
    End If
    
    ' Neu tat ca deu co du lieu
    IsDataComplete = True
    
    Exit Function
    
ErrorHandler:
    LogError "IsDataComplete", Err.Number, Err.Description
    IsDataComplete = False
End Function

' Lay danh sach sheet thieu du lieu
Public Function GetMissingDataSheets() As String
    On Error GoTo ErrorHandler
    
    Dim missingSheets As String
    missingSheets = ""
    
    ' Kiem tra tung sheet
    If Not sheetExists(SHEET_DU_NO) Then
        missingSheets = missingSheets & FILE_TYPE_DU_NO & ", "
    Else
        ' Kiem tra co du lieu
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(SHEET_DU_NO)
        If Application.WorksheetFunction.CountA(ws.UsedRange) <= 1 Then
            missingSheets = missingSheets & FILE_TYPE_DU_NO & ", "
        End If
    End If
    
    If Not sheetExists(SHEET_TAI_SAN) Then
        missingSheets = missingSheets & FILE_TYPE_TAI_SAN & ", "
    Else
        ' Kiem tra co du lieu
        Dim wsTaiSan As Worksheet
        Set wsTaiSan = ThisWorkbook.Worksheets(SHEET_TAI_SAN)
        If Application.WorksheetFunction.CountA(wsTaiSan.UsedRange) <= 1 Then
            missingSheets = missingSheets & FILE_TYPE_TAI_SAN & ", "
        End If
    End If
    
    If Not sheetExists(SHEET_TRA_GOC) Then
        missingSheets = missingSheets & FILE_TYPE_TRA_GOC & ", "
    Else
        ' Kiem tra co du lieu
        Dim wsTraGoc As Worksheet
        Set wsTraGoc = ThisWorkbook.Worksheets(SHEET_TRA_GOC)
        If Application.WorksheetFunction.CountA(wsTraGoc.UsedRange) <= 1 Then
            missingSheets = missingSheets & FILE_TYPE_TRA_GOC & ", "
        End If
    End If
    
    If Not sheetExists(SHEET_TRA_LAI) Then
        missingSheets = missingSheets & FILE_TYPE_TRA_LAI & ", "
    Else
        ' Kiem tra co du lieu
        Dim wsTraLai As Worksheet
        Set wsTraLai = ThisWorkbook.Worksheets(SHEET_TRA_LAI)
        If Application.WorksheetFunction.CountA(wsTraLai.UsedRange) <= 1 Then
            missingSheets = missingSheets & FILE_TYPE_TRA_LAI & ", "
        End If
    End If
    
    ' Loai bo dau phay o cuoi
    If Len(missingSheets) > 0 Then
        missingSheets = Left(missingSheets, Len(missingSheets) - 2)
    End If
    
    GetMissingDataSheets = missingSheets
    
    Exit Function
    
ErrorHandler:
    LogError "GetMissingDataSheets", Err.Number, Err.Description
    GetMissingDataSheets = "Khong xac dinh"
End Function

' Chuan hoa ten sheet dua vao loai file
Public Function GetSheetNameFromFileType(ByVal fileType As String) As String
    On Error GoTo ErrorHandler
    
    Select Case fileType
        Case FILE_TYPE_DU_NO
            GetSheetNameFromFileType = SHEET_DU_NO
        Case FILE_TYPE_TAI_SAN
            GetSheetNameFromFileType = SHEET_TAI_SAN
        Case FILE_TYPE_TRA_GOC
            GetSheetNameFromFileType = SHEET_TRA_GOC
        Case FILE_TYPE_TRA_LAI
            GetSheetNameFromFileType = SHEET_TRA_LAI
        Case Else
            GetSheetNameFromFileType = ""
    End Select
    
    Exit Function
    
ErrorHandler:
    LogError "GetSheetNameFromFileType", Err.Number, Err.Description
    GetSheetNameFromFileType = ""
End Function

' Kiem tra tinh hop le cua file import
Public Function ValidateImportFile(ByVal filePath As String, ByVal fileType As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la khong hop le
    ValidateImportFile = False
    
    ' Kiem tra file co ton tai
    If Not FileExists(filePath) Then
        MsgBox MSG_FILE_NOT_FOUND, vbExclamation, TITLE_ERROR
        Exit Function
    End If
    
    ' Mo file de kiem tra
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    
    ' Kiem tra du lieu co phu hop voi loai file
    Dim isValid As Boolean
    isValid = False
    
    ' Lay sheet dau tien
    Dim wsSource As Worksheet
    Set wsSource = wbSource.Worksheets(1)
    ' Kiem tra sheet co ton tai
    If wsSource Is Nothing Then
        isValid = False
        Exit Function
    End If
    
    ' Phuong phap kiem tra phu thuoc vao loai file
    Select Case fileType
        Case FILE_TYPE_DU_NO
            ' Kiem tra cac truong du lieu cua file Du no
            Dim hasField1 As Boolean, hasField2 As Boolean, hasField3 As Boolean
            hasField1 = Not wsSource.Cells.Find("custseq", LookIn:=xlValues) Is Nothing
            hasField2 = Not wsSource.Cells.Find("custnm", LookIn:=xlValues) Is Nothing
            hasField3 = Not wsSource.Cells.Find("dsbsdt", LookIn:=xlValues) Is Nothing
            isValid = hasField1 And hasField2 And hasField3
                      
        Case FILE_TYPE_TAI_SAN
            ' Kiem tra cac truong du lieu cua file Tai san
            Dim hasField1TS As Boolean, hasField2TS As Boolean, hasField3TS As Boolean
            hasField1TS = Not wsSource.Cells.Find("clno", LookIn:=xlValues) Is Nothing
            hasField2TS = Not wsSource.Cells.Find("clcustnm", LookIn:=xlValues) Is Nothing
            hasField3TS = Not wsSource.Cells.Find("cltpcd", LookIn:=xlValues) Is Nothing
            isValid = hasField1TS And hasField2TS And hasField3TS

        Case FILE_TYPE_TRA_GOC
            ' Kiem tra cac truong du lieu cua file Tra goc
            Dim hasField1TG As Boolean, hasField2TG As Boolean, hasField3TG As Boolean
            hasField1TG = Not wsSource.Cells.Find("custseqno", LookIn:=xlValues) Is Nothing
            hasField2TG = Not wsSource.Cells.Find("custnm", LookIn:=xlValues) Is Nothing
            hasField3TG = Not wsSource.Cells.Find("matdt", LookIn:=xlValues) Is Nothing
            isValid = hasField1TG And hasField2TG And hasField3TG
        
        Case FILE_TYPE_TRA_LAI
            ' Kiem tra cac truong du lieu cua file Tra lai
            Dim hasCustSeqNoTL As Boolean
            Dim hasCustNmTL As Boolean
            Dim hasMatDtTL As Boolean
            
            hasCustSeqNoTL = Not wsSource.Cells.Find("custseqno", LookIn:=xlValues) Is Nothing
            hasCustNmTL = Not wsSource.Cells.Find("custnm", LookIn:=xlValues) Is Nothing
            hasMatDtTL = Not wsSource.Cells.Find("matdt", LookIn:=xlValues) Is Nothing
            
            isValid = hasCustSeqNoTL And hasCustNmTL And hasMatDtTL
        
        Case Else
            isValid = False
    End Select
    
    ' Dong workbook nguon
    wbSource.Close SaveChanges:=False
    
    ' Tra ve ket qua kiem tra
    ValidateImportFile = isValid
    
    Exit Function
    
ErrorHandler:
    ' Dong workbook nguon neu dang mo
    If Not wbSource Is Nothing Then
        On Error Resume Next
        wbSource.Close SaveChanges:=False
    End If
    
    LogError "ValidateImportFile", Err.Number, Err.Description
    ValidateImportFile = False
End Function

' Format worksheet sau khi import
Public Sub FormatWorksheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Lay dong cuoi cung co du lieu
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Neu worksheet trong, khong lam gi ca
    If ws.UsedRange.Rows.Count <= 1 Then
        Exit Sub
    End If
    
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' Format dong header (dong 2) neu co du lieu
    If lastRow >= 2 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    ' Them AutoFilter
    If lastRow > 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    End If
    
    ' Format cell ngay thang
    Dim i As Long, j As Long
    Dim cellValue As Variant
    
    For j = 1 To lastCol
        For i = 3 To lastRow ' Bat dau tu dong 3 (sau header)
            cellValue = ws.Cells(i, j).Value
            
            ' Kiem tra xem co phai la ngay thang khong
            If IsDate(cellValue) Then
                ws.Cells(i, j).NumberFormat = DATE_FORMAT
            End If
        Next i
    Next j
    
    ' Auto-fit tat ca cac cot
    ws.Columns("A:Z").AutoFit
    
    Exit Sub
    
ErrorHandler:
    LogError "FormatWorksheet", Err.Number, Err.Description
End Sub

' Clear du lieu cu trong worksheet
Public Sub ClearWorksheetData(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Luu thong tin o dong dau tien
    Dim firstRowInfo As String
    firstRowInfo = ws.Cells(INFO_ROW, 1).Value
    
    ' Xoa du lieu
    ws.Cells.Clear
    
    ' Khoi phuc thong tin dong dau tien
    ws.Cells(INFO_ROW, 1).Value = firstRowInfo
    
    Exit Sub
    
ErrorHandler:
    LogError "ClearWorksheetData", Err.Number, Err.Description
End Sub

' Kiem tra tuong thich phien ban Excel
Public Function IsExcelVersionCompatible() As Boolean
    On Error GoTo ErrorHandler
    
    ' Lay phien ban Excel
    Dim excelVersion As Double
    excelVersion = Val(Application.Version)
    
    ' Kiem tra phien ban (Excel 2007 tro len)
    IsExcelVersionCompatible = (excelVersion >= 12)
    
    Exit Function
    
ErrorHandler:
    LogError "IsExcelVersionCompatible", Err.Number, Err.Description
    IsExcelVersionCompatible = False
End Function

' Kiem tra neu workbook dang o che do chi doc
Public Function IsWorkbookReadOnly() As Boolean
    On Error GoTo ErrorHandler
    
    IsWorkbookReadOnly = ThisWorkbook.ReadOnly
    
    Exit Function
    
ErrorHandler:
    LogError "IsWorkbookReadOnly", Err.Number, Err.Description
    IsWorkbookReadOnly = True ' Gia su la chi doc neu co loi
End Function

' Tao ban sao du lieu truoc khi import
Public Sub BackupCurrentData()
    On Error GoTo ErrorHandler
    
    ' Tao thu muc backup neu chua ton tai
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backup"
    
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If
    
    ' Tao ten file backup voi timestamp
    Dim backupFileName As String
    backupFileName = backupFolder & "\Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    
    ' Luu ban sao cua workbook hien tai
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupFileName
    Application.DisplayAlerts = True
    
    ' Ghi log thong tin
    LogInfo "BackupCurrentData", "Da tao ban sao du lieu tai: " & backupFileName
    
    Exit Sub
    
ErrorHandler:
    LogError "BackupCurrentData", Err.Number, Err.Description
    ' Khong hien thong bao loi vi day la chuc nang ngam
End Sub

' Kiem tra tinh hop le cua du lieu sau khi import
Public Function ValidateImportedData(ByVal sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la hop le
    ValidateImportedData = True
    
    ' Kiem tra sheet co ton tai
    If Not sheetExists(sheetName) Then
        ValidateImportedData = False
        Exit Function
    End If
    
    ' Lay sheet can kiem tra
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' Kiem tra co du lieu hay khong
    If ws.UsedRange.Rows.Count <= 1 Then
        ValidateImportedData = False
        Exit Function
    End If
    
    ' Kiem tra cac cot bat buoc tuy theo loai du lieu
    Select Case sheetName
        Case SHEET_DU_NO
            ' Kiem tra cot bat buoc cho Du no
            If ws.Cells.Find("custseq", LookIn:=xlValues) Is Nothing Or _
               ws.Cells.Find("custnm", LookIn:=xlValues) Is Nothing Then
                ValidateImportedData = False
                Exit Function
            End If
            
        Case SHEET_TAI_SAN
            ' Kiem tra cot bat buoc cho Tai san
            If ws.Cells.Find("clno", LookIn:=xlValues) Is Nothing Or _
               ws.Cells.Find("clcustnm", LookIn:=xlValues) Is Nothing Then
                ValidateImportedData = False
                Exit Function
            End If
            
        Case SHEET_TRA_GOC
            ' Kiem tra cot bat buoc cho Tra goc
            If ws.Cells.Find("custseqno", LookIn:=xlValues) Is Nothing Or _
               ws.Cells.Find("custnm", LookIn:=xlValues) Is Nothing Then
                ValidateImportedData = False
                Exit Function
            End If
            
        Case SHEET_TRA_LAI
            ' Kiem tra cot bat buoc cho Tra lai
            If ws.Cells.Find("custseqno", LookIn:=xlValues) Is Nothing Or _
               ws.Cells.Find("custnm", LookIn:=xlValues) Is Nothing Then
                ValidateImportedData = False
                Exit Function
            End If
            
        Case Else
            ' Neu la sheet khong xac dinh, van coi la hop le
    End Select
    
    Exit Function
    
ErrorHandler:
    LogError "ValidateImportedData", Err.Number, Err.Description
    ValidateImportedData = False
End Function

' Dem so ban ghi trong sheet
Public Function CountRecords(ByVal sheetName As String) As Long
    On Error GoTo ErrorHandler
    
    ' Khoi tao
    CountRecords = 0
    
    ' Kiem tra sheet co ton tai
    If Not sheetExists(sheetName) Then
        Exit Function
    End If
    
    ' Lay sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' Kiem tra co du lieu hay khong
    If ws.UsedRange.Rows.Count <= 1 Then
        Exit Function
    End If
    
    ' Dem so ban ghi (tru dong header)
    CountRecords = ws.UsedRange.Rows.Count - 1
    
    Exit Function
    
ErrorHandler:
    LogError "CountRecords", Err.Number, Err.Description
    CountRecords = 0
End Function

' Tao tong quan du lieu sau khi import
Public Sub GenerateDataSummary()
    On Error GoTo ErrorHandler
    
    ' Kiem tra da co du lieu chua
    If Not IsDataComplete() Then
        Exit Sub
    End If
    
    ' Tao hoac lay sheet Config
    Dim wsConfig As Worksheet
    
    If Not sheetExists(SHEET_CONFIG) Then
        Set wsConfig = ThisWorkbook.Worksheets.Add
        wsConfig.Name = SHEET_CONFIG
    Else
        Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    End If
    
    ' Xoa du lieu cu
    wsConfig.Cells.Clear
    
    ' Tao tieu de
    wsConfig.Range("A1").Value = "TONG QUAN DU LIEU"
    wsConfig.Range("A1").Font.Bold = True
    wsConfig.Range("A1").Font.Size = 14
    
    ' Thong tin import
    wsConfig.Range("A3").Value = "THONG TIN IMPORT:"
    wsConfig.Range("A3").Font.Bold = True
    
    wsConfig.Range("A4").Value = "Du no:"
    wsConfig.Range("B4").Value = GetLastImportInfo(SHEET_DU_NO)
    
    wsConfig.Range("A5").Value = "Tai san:"
    wsConfig.Range("B5").Value = GetLastImportInfo(SHEET_TAI_SAN)
    
    wsConfig.Range("A6").Value = "Tra goc:"
    wsConfig.Range("B6").Value = GetLastImportInfo(SHEET_TRA_GOC)
    
    wsConfig.Range("A7").Value = "Tra lai:"
    wsConfig.Range("B7").Value = GetLastImportInfo(SHEET_TRA_LAI)
    
    ' Tong quan so luong ban ghi
    wsConfig.Range("A9").Value = "SO LUONG BAN GHI:"
    wsConfig.Range("A9").Font.Bold = True
    
    wsConfig.Range("A10").Value = "Du no:"
    wsConfig.Range("B10").Value = CountRecords(SHEET_DU_NO)
    
    wsConfig.Range("A11").Value = "Tai san:"
    wsConfig.Range("B11").Value = CountRecords(SHEET_TAI_SAN)
    
    wsConfig.Range("A12").Value = "Tra goc:"
    wsConfig.Range("B12").Value = CountRecords(SHEET_TRA_GOC)
    
    wsConfig.Range("A13").Value = "Tra lai:"
    wsConfig.Range("B13").Value = CountRecords(SHEET_TRA_LAI)
    
    ' Tong hop cac loai du lieu
    wsConfig.Range("A15").Value = "TONG HOP DU LIEU:"
    wsConfig.Range("A15").Font.Bold = True
    
    ' Format
    wsConfig.Columns("A:B").AutoFit
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateDataSummary", Err.Number, Err.Description
End Sub

' An tat ca cac sheet du lieu
Public Sub HideDataSheets()
    On Error GoTo ErrorHandler
    
    ' An cac sheet du lieu
    If sheetExists(SHEET_DU_NO) Then
        ThisWorkbook.Worksheets(SHEET_DU_NO).Visible = xlSheetHidden
    End If
    
    If sheetExists(SHEET_TAI_SAN) Then
        ThisWorkbook.Worksheets(SHEET_TAI_SAN).Visible = xlSheetHidden
    End If
    
    If sheetExists(SHEET_TRA_GOC) Then
        ThisWorkbook.Worksheets(SHEET_TRA_GOC).Visible = xlSheetHidden
    End If
    
    If sheetExists(SHEET_TRA_LAI) Then
        ThisWorkbook.Worksheets(SHEET_TRA_LAI).Visible = xlSheetHidden
    End If
    
    If sheetExists(SHEET_LOG) Then
        ThisWorkbook.Worksheets(SHEET_LOG).Visible = xlSheetHidden
    End If
    
    ' An sheet Config (ch? n?u không ph?i là developer mode)
    If sheetExists(SHEET_CONFIG) Then
        ThisWorkbook.Worksheets(SHEET_CONFIG).Visible = xlSheetHidden
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "HideDataSheets", Err.Number, Err.Description
End Sub

' Hien thi tat ca cac sheet du lieu
Public Sub ShowDataSheets()
    On Error GoTo ErrorHandler
    
    ' Hien thi cac sheet du lieu
    If sheetExists(SHEET_DU_NO) Then
        ThisWorkbook.Worksheets(SHEET_DU_NO).Visible = xlSheetVisible
    End If
    
    If sheetExists(SHEET_TAI_SAN) Then
        ThisWorkbook.Worksheets(SHEET_TAI_SAN).Visible = xlSheetVisible
    End If
    
    If sheetExists(SHEET_TRA_GOC) Then
        ThisWorkbook.Worksheets(SHEET_TRA_GOC).Visible = xlSheetVisible
    End If
    
    If sheetExists(SHEET_TRA_LAI) Then
        ThisWorkbook.Worksheets(SHEET_TRA_LAI).Visible = xlSheetVisible
    End If
    
    If sheetExists(SHEET_LOG) Then
        ThisWorkbook.Worksheets(SHEET_LOG).Visible = xlSheetVisible
    End If
    
    If sheetExists(SHEET_CONFIG) Then
        ThisWorkbook.Worksheets(SHEET_CONFIG).Visible = xlSheetVisible
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "ShowDataSheets", Err.Number, Err.Description
End Sub

' Kiem tra ten file co phu hop voi pattern
Public Function ValidateFileNamePattern(ByVal filePath As String, ByVal pattern As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Lay ten file (khong bao gom duong dan)
    Dim fileName As String
    fileName = LCase(Dir(filePath))
    
    ' Kiem tra xem ten file co chua pattern khong (khong phan biet hoa thuong)
    ValidateFileNamePattern = InStr(1, fileName, LCase(pattern)) > 0
    
    Exit Function
    
ErrorHandler:
    LogError "ValidateFileNamePattern", Err.Number, Err.Description
    ValidateFileNamePattern = False
End Function

' Kiem tra cac dieu kien brcd tren cac sheet
Public Function ValidateImportedDataWithBranchCode() As Boolean
    On Error GoTo ErrorHandler
    
    ' Khai bao bien luu ket qua kiem tra
    Dim Check_brcd_Tra_lai As Boolean
    Dim Check_brcd_Tra_goc As Boolean
    Dim Check_brcd_Du_no As Boolean
    
    ' Gia tri mac dinh la false
    Check_brcd_Tra_lai = False
    Check_brcd_Tra_goc = False
    Check_brcd_Du_no = False
    
    ' --- Kiem tra sheet Tra_lai ---
    If sheetExists(SHEET_TRA_LAI) Then
        Dim wsTraLai As Worksheet
        Set wsTraLai = ThisWorkbook.Worksheets(SHEET_TRA_LAI)
        
        ' Kiem tra so luong dong co du lieu
        If wsTraLai.UsedRange.Rows.Count >= 9 Then
            ' Tim vi tri cac cot can kiem tra
            Dim colCustSeqNoTL As Long, colAplNoTL As Long, colLnofcSeqTL As Long
            Dim colBrcdTL As Long, colBrcdlnTL As Long, colBrcdaplTL As Long
            
            ' Tim vi tri cac cot
            colCustSeqNoTL = GetColumnPosition(wsTraLai, "custseqno")
            colAplNoTL = GetColumnPosition(wsTraLai, "aplno")
            colLnofcSeqTL = GetColumnPosition(wsTraLai, "lnofcseq")
            colBrcdTL = GetColumnPosition(wsTraLai, "brcd")
            colBrcdlnTL = GetColumnPosition(wsTraLai, "brcdln")
            colBrcdaplTL = GetColumnPosition(wsTraLai, "brcdapl")
            
            ' Kiem tra neu tim thay tat ca cac cot
            If colCustSeqNoTL > 0 And colAplNoTL > 0 And colLnofcSeqTL > 0 And _
               colBrcdTL > 0 And colBrcdlnTL > 0 And colBrcdaplTL > 0 Then
                
                ' Kiem tra dieu kien dua tren du lieu o dong 9
                Dim custseqnoTL As String, aplnoTL As String, lnofcseqTL As String
                Dim brcdTL As String, brcdlnTL As String, brcdaplTL As String
                
                ' Lay gia tri cac cell
                custseqnoTL = Trim(CStr(wsTraLai.Cells(9, colCustSeqNoTL).Value))
                aplnoTL = Trim(CStr(wsTraLai.Cells(9, colAplNoTL).Value))
                lnofcseqTL = Trim(CStr(wsTraLai.Cells(9, colLnofcSeqTL).Value))
                brcdTL = Trim(CStr(wsTraLai.Cells(9, colBrcdTL).Value))
                brcdlnTL = Trim(CStr(wsTraLai.Cells(9, colBrcdlnTL).Value))
                brcdaplTL = Trim(CStr(wsTraLai.Cells(9, colBrcdaplTL).Value))
                
                ' Kiem tra dieu kien
                Dim check1TL As Boolean, check2TL As Boolean
                
                ' Kiem tra 4 ky tu dau tien cua custseqno, aplno, lnofcseq la '1902'
                check1TL = (Len(custseqnoTL) >= 4 And Left(custseqnoTL, 4) = "1902") And _
                           (Len(aplnoTL) >= 4 And Left(aplnoTL, 4) = "1902") And _
                           (Len(lnofcseqTL) >= 4 And Left(lnofcseqTL, 4) = "1902")
                           
                ' Kiem tra brcd, brcdln, brcdapl deu co gia tri la 1902
                check2TL = (brcdTL = "1902") And (brcdlnTL = "1902") And (brcdaplTL = "1902")
                
                ' Dieu kien sheet Tra_lai hop le
                Check_brcd_Tra_lai = check1TL And check2TL
            End If
        End If
    End If
    
    ' --- Kiem tra sheet Tra_goc ---
    If sheetExists(SHEET_TRA_GOC) Then
        Dim wsTraGoc As Worksheet
        Set wsTraGoc = ThisWorkbook.Worksheets(SHEET_TRA_GOC)
        
        ' Kiem tra so luong dong co du lieu
        If wsTraGoc.UsedRange.Rows.Count >= 8 Then
            ' Tim vi tri cac cot can kiem tra
            Dim colCustSeqNoTG As Long, colAplNoTG As Long, colLnofcSeqTG As Long
            Dim colBrcdTG As Long, colBrcdlnTG As Long, colBrcdaplTG As Long
            
            ' Tim vi tri cac cot
            colCustSeqNoTG = GetColumnPosition(wsTraGoc, "custseqno")
            colAplNoTG = GetColumnPosition(wsTraGoc, "aplno")
            colLnofcSeqTG = GetColumnPosition(wsTraGoc, "lnofcseq")
            colBrcdTG = GetColumnPosition(wsTraGoc, "brcd")
            colBrcdlnTG = GetColumnPosition(wsTraGoc, "brcdln")
            colBrcdaplTG = GetColumnPosition(wsTraGoc, "brcdapl")
            
            ' Kiem tra neu tim thay tat ca cac cot
            If colCustSeqNoTG > 0 And colAplNoTG > 0 And colLnofcSeqTG > 0 And _
               colBrcdTG > 0 And colBrcdlnTG > 0 And colBrcdaplTG > 0 Then
                
                ' Kiem tra dieu kien dua tren du lieu o dong 8
                Dim custseqnoTG As String, aplnoTG As String, lnofcseqTG As String
                Dim brcdTG As String, brcdlnTG As String, brcdaplTG As String
                
                ' Lay gia tri cac cell
                custseqnoTG = Trim(CStr(wsTraGoc.Cells(8, colCustSeqNoTG).Value))
                aplnoTG = Trim(CStr(wsTraGoc.Cells(8, colAplNoTG).Value))
                lnofcseqTG = Trim(CStr(wsTraGoc.Cells(8, colLnofcSeqTG).Value))
                brcdTG = Trim(CStr(wsTraGoc.Cells(8, colBrcdTG).Value))
                brcdlnTG = Trim(CStr(wsTraGoc.Cells(8, colBrcdlnTG).Value))
                brcdaplTG = Trim(CStr(wsTraGoc.Cells(8, colBrcdaplTG).Value))
                
                ' Kiem tra dieu kien
                Dim check1TG As Boolean, check2TG As Boolean
                
                ' Kiem tra 4 ky tu dau tien cua custseqno, aplno, lnofcseq la '1902'
                check1TG = (Len(custseqnoTG) >= 4 And Left(custseqnoTG, 4) = "1902") And _
                           (Len(aplnoTG) >= 4 And Left(aplnoTG, 4) = "1902") And _
                           (Len(lnofcseqTG) >= 4 And Left(lnofcseqTG, 4) = "1902")
                           
                ' Kiem tra brcd, brcdln, brcdapl deu co gia tri la 1902
                check2TG = (brcdTG = "1902") And (brcdlnTG = "1902") And (brcdaplTG = "1902")
                
                ' Dieu kien sheet Tra_goc hop le
                Check_brcd_Tra_goc = check1TG And check2TG
            End If
        End If
    End If
    
    ' --- Kiem tra sheet Du_no ---
    If sheetExists(SHEET_DU_NO) Then
        Dim wsDuNo As Worksheet
        Set wsDuNo = ThisWorkbook.Worksheets(SHEET_DU_NO)
        
        ' Kiem tra so luong dong co du lieu
        If wsDuNo.UsedRange.Rows.Count >= 8 Then
            ' Tim vi tri cac cot can kiem tra
            Dim colBrcdDN As Long, colApprSeqDN As Long, colDsbsSeqDN As Long, colCustSeqDN As Long
            
            ' Tim vi tri cac cot
            colBrcdDN = GetColumnPosition(wsDuNo, "brcd")
            colApprSeqDN = GetColumnPosition(wsDuNo, "apprseq")
            colDsbsSeqDN = GetColumnPosition(wsDuNo, "dsbsseq")
            colCustSeqDN = GetColumnPosition(wsDuNo, "custseq")
            
            ' Kiem tra neu tim thay tat ca cac cot
            If colBrcdDN > 0 And colApprSeqDN > 0 And colDsbsSeqDN > 0 And colCustSeqDN > 0 Then
                
                ' Kiem tra dieu kien dua tren du lieu o dong 8
                Dim brcdDN As String, apprseqDN As String, dsbsseqDN As String, custseqDN As String
                
                ' Lay gia tri cac cell
                brcdDN = Trim(CStr(wsDuNo.Cells(8, colBrcdDN).Value))
                apprseqDN = Trim(CStr(wsDuNo.Cells(8, colApprSeqDN).Value))
                dsbsseqDN = Trim(CStr(wsDuNo.Cells(8, colDsbsSeqDN).Value))
                custseqDN = Trim(CStr(wsDuNo.Cells(8, colCustSeqDN).Value))
                
                ' Kiem tra dieu kien
                Dim check1DN As Boolean, check2DN As Boolean
                
                ' Kiem tra 4 ky tu dau tien cua brcd, apprseq, dsbsseq la '1902'
                check1DN = (Len(brcdDN) >= 4 And Left(brcdDN, 4) = "1902") And _
                           (Len(apprseqDN) >= 4 And Left(apprseqDN, 4) = "1902") And _
                           (Len(dsbsseqDN) >= 4 And Left(dsbsseqDN, 4) = "1902")
                           
                ' Kiem tra custseq co gia tri la 1902
                check2DN = (custseqDN = "1902")
                
                ' Dieu kien sheet Du_no hop le
                Check_brcd_Du_no = check1DN And check2DN
            End If
        End If
    End If
    
    ' Ket qua cuoi cung: tat ca 3 sheet deu phai hop le
    ValidateImportedDataWithBranchCode = Check_brcd_Tra_lai And Check_brcd_Tra_goc And Check_brcd_Du_no
    
    Exit Function
    
ErrorHandler:
    LogError "ValidateImportedDataWithBranchCode", Err.Number, Err.Description
    ValidateImportedDataWithBranchCode = False
End Function

' Ham ho tro: Tim vi tri cot theo ten
Private Function GetColumnPosition(ws As Worksheet, columnName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Set cell = ws.Rows(FIRST_DATA_ROW).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cell Is Nothing Then
        GetColumnPosition = cell.Column
    Else
        GetColumnPosition = 0
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "GetColumnPosition", Err.Number, Err.Description
    GetColumnPosition = 0
End Function

