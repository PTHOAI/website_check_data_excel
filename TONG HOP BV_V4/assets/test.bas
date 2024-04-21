Public status_global As Boolean, thieu_file_pdf As Long, name_file_globle As String, link_TH_file_dich As String, check_name_file_exist As String, thieu_file_excel As Long, type_file_excel_thieu As String
Sub TH_FILE_CG()
    Dim isEx As Boolean
    Dim duong_dan_dau_vao As String
    Dim duong_dan_dich As String
    Dim arr_name_file() As String
    thieu_file_pdf = 0
    thieu_file_excel = 0
    type_file_excel_thieu = ""
    ThisWorkbook.Sheets(2).Range("A5:D500").Delete
    ThisWorkbook.Sheets(2).Tab.Color = RGB(217, 217, 217)
    duong_dan_dau_vao = ThisWorkbook.Sheets(1).Cells(7, 6)
    duong_dan_dau_ra = ThisWorkbook.Sheets(1).Cells(10, 6)
    link_TH_file_dich = duong_dan_dau_ra
    If duong_dan_dau_vao = "" Then
        MsgBox "Nhap duong dan dau ra truoc khi tong hop"
        Exit Sub
    End If
    If duong_dan_dau_ra = "" Then
        MsgBox "Nhap duong dan dau ra truoc khi thuc hien tong hop!"
        Exit Sub
    End If
    
    isEx = Check_Expiry(5242025)
    If isEx Then
        MsgBox "error: cannot access the system, please contact the writer to resolve", vbExclamation
        Exit Sub
    End If
    
    For I = 3 To ThisWorkbook.Sheets(1).Cells(4, 6) + 2
        name_file_globle = ThisWorkbook.Sheets(1).Cells(I, 1)
        arr_name_file = Split(ThisWorkbook.Sheets(1).Cells(I, 1), "/")
        check_name_file_exist = arr_name_file(3)
       TimKiemTenFileTrongThuMucCon arr_name_file(3), duong_dan_dau_vao
    Next I
    
    For I = 3 To ThisWorkbook.Sheets(1).Cells(4, 7) * 3 + 2
        name_file_globle = ThisWorkbook.Sheets(1).Cells(I, 2)
        arr_name_file = Split(ThisWorkbook.Sheets(1).Cells(I, 2), "/")
        check_name_file_exist = arr_name_file(3)
       TimKiemTenFileTrongThuMucCon arr_name_file(3), duong_dan_dau_vao
    Next I
    
    If thieu_file_pdf > 0 Or thieu_file_excel > 0 Then
        ThisWorkbook.Sheets(2).Tab.Color = RGB(255, 0, 0)
        MsgBox "TH File Thieu:___PDF: " & thieu_file_pdf & "EXCEL: " & thieu_file_excel
        Sheets(2).Select
        Else
        MsgBox "TH File PDF Thanh cong "
    End If
    
End Sub

Function Check_Expiry(value_Ex As Long) As Boolean

    Dim NgayHienTai As Date
    NgayHienTai = Date
    
    Dim Ngay As Integer
    Dim Thang As Integer
    Dim Nam As Integer
    Dim checkValue As Variant
    
    Ngay = Day(NgayHienTai)
    Thang = Month(NgayHienTai)
    Nam = Year(NgayHienTai)
    
    checkValue = CLng(Thang & Ngay & Nam)
    
    If checkValue > value_Ex Then
        Check_Expiry = True
        Exit Function
    End If
    
    Check_Expiry = False
    
End Function

Function TimKiemTenFileTrongThuMucCon(ten_file As String, link_input As String)
    status_global = False
    Dim FSO As Object
    Set FSO = CreateObject("scripting.FileSystemObject")
    Dim objFSO As New FileSystemObject
    
    Dim objFolder As Folder
    For Each objFolder In objFSO.GetFolder(link_input).SubFolders
        If FSO.FileExists(objFolder.Path & "\" & ten_file) Then
            Get_Data_TH objFolder.Path & "\" & ten_file
            Exit Function
        End If
        If status_global Then
            Exit Function
        End If
        TimKiemTenFileTrongThuMuc objFolder, ten_file
    Next objFolder
    Kiem_tra_loai_file_thieu
    Hien_thi_file_thieu
End Function


Function TimKiemTenFileTrongThuMuc(objFolder As Folder, strFileName As String)
Dim FSO As Object
Set FSO = CreateObject("scripting.FileSystemObject")

    For Each objFolder In objFolder.SubFolders
        If FSO.FileExists(objFolder.Path & "\" & strFileName) Then
            status_global = True
            Get_Data_TH objFolder.Path & "\" & strFileName
            Exit Function
        End If
        TimKiemTenFileTrongThuMuc objFolder, strFileName
    Next objFolder

End Function

Function Get_Data_TH(duong_dan_file_TH As String)
    Dim arr_name_file() As String
    Dim folder_level1 As String
    Dim folder_level2 As String
    Dim folder_level3 As String
    Dim FSO As Object
    Set FSO = CreateObject("scripting.FileSystemObject")
    arr_name_file = Split(name_file_globle, "/")
    folder_level1 = link_TH_file_dich & "\" & arr_name_file(0)
    folder_level2 = folder_level1 & "\" & arr_name_file(1)
    folder_level3 = folder_level2 & "\" & arr_name_file(2)
    If FSO.FolderExists(folder_level1) = False Then
        FSO.CreateFolder (folder_level1)
    End If
    If FSO.FolderExists(folder_level2) = False Then
        FSO.CreateFolder (folder_level2)
    End If
    If FSO.FolderExists(folder_level3) = False Then
        FSO.CreateFolder (folder_level3)
    End If
    FileCopy duong_dan_file_TH, folder_level3 & "\" & arr_name_file(3)
    Yes_excel arr_name_file(3)
End Function

Function Hien_thi_file_thieu()
    If Mid(check_name_file_exist, (Len(check_name_file_exist) + 1) - 4, 4) = ".pdf" Then
        ThisWorkbook.Sheets(2).Cells(thieu_file_pdf + 4, 1).value = check_name_file_exist
    End If
End Function

Function Kiem_tra_loai_file_thieu()
    Dim arr_name_file() As String
    If Mid(check_name_file_exist, (Len(check_name_file_exist) + 1) - 4, 4) = ".pdf" Then
        thieu_file_pdf = thieu_file_pdf + 1
    End If
    Count_exist_excel
End Function
Function Count_exist_excel()
    Dim arr_name_file() As String
    arr_name_file = Split(check_name_file_exist, ".")
    
    If Mid(check_name_file_exist, (Len(check_name_file_exist) + 1) - 5, 5) = ".xlsx" Or Mid(check_name_file_exist, (Len(check_name_file_exist) + 1) - 4, 4) = ".xls" Or Mid(check_name_file_exist, (Len(check_name_file_exist) + 1) - 5, 5) = ".xlsm" Then
        If arr_name_file(0) <> type_file_excel_thieu Then
            thieu_file_excel = thieu_file_excel + 1
            ThisWorkbook.Sheets(2).Cells(thieu_file_excel + 4, 2).value = arr_name_file(0) & ".xlsx"
            type_file_excel_thieu = arr_name_file(0)
        End If
    End If
End Function
Function Yes_excel(value As String)
    Dim arr_name_file() As String
    arr_name_file = Split(value, ".")
    
    If Mid(value, (Len(value) + 1) - 5, 5) = ".xlsx" Or Mid(value, (Len(value) + 1) - 4, 4) = ".xls" Or Mid(value, (Len(value) + 1) - 5, 5) = ".xlsm" Then
        If arr_name_file(0) = type_file_excel_thieu Or type_file_excel_thieu = "" Then
            ThisWorkbook.Sheets(2).Cells(thieu_file_excel + 4, 2).value = ""
            thieu_file_excel = thieu_file_excel - 1
            Else
            type_file_excel_thieu = arr_name_file(0)
        End If
    End If
End Function
