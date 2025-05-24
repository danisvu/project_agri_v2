Attribute VB_Name = "modConstants"
Option Explicit

' ======================================================
' Module: modConstants
' Mo ta: Chua cac hang so su dung trong ung dung
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 18/05/2025
' ======================================================

' --- Hang so cho ten cac sheet du lieu ---
Public Const SHEET_DU_NO As String = "Du_no"
Public Const SHEET_TAI_SAN As String = "Tai_san"
Public Const SHEET_TRA_GOC As String = "Tra_goc"
Public Const SHEET_TRA_LAI As String = "Tra_lai"
Public Const SHEET_MAIN_MENU As String = "MainMenu"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_CONFIG As String = "Config"

' --- Hang so cho thong bao form Import ---
Public Const MSG_IMPORT_SUCCESS As String = "Import du lieu thanh cong!"
Public Const MSG_IMPORT_FAILED As String = "Loi khi import du lieu!"
Public Const MSG_FILE_INVALID As String = "File khong hop le. Vui long chon dung loai file!"
Public Const MSG_DATA_INCOMPLETE As String = "Chua co du 4 loai du lieu. Vui long import cac file con thieu!"
Public Const MSG_FILE_NOT_FOUND As String = "Khong tim thay file. Vui long kiem tra duong dan!"
Public Const MSG_SELECT_FILE As String = "Vui long chon file"
Public Const MSG_READY_TO_CONTINUE As String = "Da import du 4 loai du lieu. Ban co the tiep tuc su dung he thong!"
Public Const MSG_WORKBOOK_READ_ONLY As String = "File dang duoc mo o che do chi doc. Khong the import du lieu!"
Public Const MSG_IMPORT_OVERWRITE As String = "Du lieu da ton tai. Ban co muon ghi de du lieu cu?"
Public Const MSG_SHEET_DOES_NOT_EXIST As String = "Sheet khong ton tai. He thong se tao sheet moi!"
Public Const MSG_CLOSE_WORKBOOK_ERROR As String = "Khong the dong workbook nguon. Vui long dong thu cong neu can!"

' --- Hang so cho tieu de thong bao ---
Public Const TITLE_IMPORT As String = "Import Du Lieu"
Public Const TITLE_ERROR As String = "Loi"
Public Const TITLE_SUCCESS As String = "Thanh cong"
Public Const TITLE_WARNING As String = "Chu y"
Public Const TITLE_CONFIRMATION As String = "Xac nhan"
Public Const TITLE_INFORMATION As String = "Thong tin"

' --- Hang so cho dinh dang thoi gian ---
Public Const DATE_TIME_FORMAT As String = "dd/MM/yyyy HH:mm:ss"
Public Const DATE_FORMAT As String = "dd/MM/yyyy"

' --- Hang so cho pattern kiem tra ten file ---
Public Const FILE_PATTERN_DU_NO As String = "du no"
Public Const FILE_PATTERN_TAI_SAN As String = "tai san"
Public Const FILE_PATTERN_TRA_GOC As String = "tra goc"
Public Const FILE_PATTERN_TRA_LAI As String = "tra lai"
Public Const MSG_FILE_NAME_INVALID As String = "Ten file khong phu hop. Vui long chon file co ten chua ""{0}""!"
Public Const TIME_FORMAT As String = "HH:mm:ss"

' --- Hang so cho prefix thong tin import ---
Public Const IMPORT_INFO_PREFIX As String = "Thoi gian import: "

' --- Hang so cho cac loai file ---
Public Const FILE_TYPE_DU_NO As String = "Du no"
Public Const FILE_TYPE_TAI_SAN As String = "Tai san"
Public Const FILE_TYPE_TRA_GOC As String = "Tra goc"
Public Const FILE_TYPE_TRA_LAI As String = "Tra lai"

' --- Hang so cho file dialog ---
Public Const FILE_FILTER As String = "Excel Files (*.xls;*.xlsx),*.xls;*.xlsx"

' --- Hang so cho log ---
Public Const LOG_ERROR_PREFIX As String = "LOI"
Public Const LOG_WARNING_PREFIX As String = "CANH BAO"
Public Const LOG_INFO_PREFIX As String = "THONG TIN"

' --- Hang so cho vi tri cell ---
Public Const FIRST_DATA_ROW As Long = 2 ' Dong bat dau du lieu (sau dong header)
Public Const INFO_ROW As Long = 1 ' Dong chua thong tin import

Public Const xlValues As Long = -4163 ' Gia tri hang so Excel

' --- Hang so cho Form MainMenu ---
Public Const TITLE_MAIN_MENU As String = "Quan Ly Thong Tin Khach Hang Vay"
Public Const MSG_DATA_INCOMPLETE_FOR_MAIN As String = "Chua co du du lieu de hien thi Main Menu. Vui long import day du du lieu truoc!"

' --- Hang so mau sac ---
Public Const AGRIBANK_RED_COLOR As Long = 4194367 ' RGB(174, 28, 63)

' --- Hang so cho cac chuc nang ---
Public Const FUNC_KHACH_HANG As String = "KhachHang"
Public Const FUNC_KHOAN_VAY As String = "KhoanVay"
Public Const FUNC_TAI_SAN As String = "TaiSan"
Public Const FUNC_TRA_NO As String = "TraNo"
Public Const FUNC_CANH_BAO As String = "CanhBao"
Public Const FUNC_BAO_CAO As String = "BaoCao"
Public Const FUNC_IMPORT_DATA As String = "ImportData"
