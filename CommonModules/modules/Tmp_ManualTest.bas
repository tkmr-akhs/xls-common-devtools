Attribute VB_Name = "Tmp_ManualTest"
Option Explicit
Public Sub ManualTestMain()
'    Dim FSSrv As New IFileSystemService
'    Debug.Print Join(FSSrv.GetFileList("C:\Users\kaosun\Desktop"), vbCrLf)
'    Debug.Print "-------------------------"
'    Debug.Print Join(FSSrv.GetDirectoryList("C:\Users\kaosun\Desktop"), vbCrLf)
'    Debug.Print "-------------------------"
'    Debug.Print FSSrv.IsDirectory("C:\Users\kaosun\Desktop\新しいフォルダー")
'    Debug.Print FSSrv.IsFile("C:\Users\kaosun\Desktop\新しいフォルダー")
'    Debug.Print FSSrv.PathExists("C:\Users\kaosun\Desktop\新しいフォルダー")
'    Call FSSrv.CreateDirectory("C:\Users\kaosun\Desktop\新しいフォルダー\1\2\3", Force:=True, Recursive:=True)
'    Call FSSrv.CreateDirectory("C:\Users\kaosun\Desktop\新しいフォルダー\2\2\3", Force:=True, Recursive:=True)
'    Call FSSrv.CopyDirectory("C:\Users\kaosun\Desktop\新しいフォルダー\1", "C:\Users\kaosun\Desktop\新しいフォルダー\2\")
'    Call FSSrv.MoveDirectory("C:\Users\kaosun\Desktop\新しいフォルダー\1", "C:\Users\kaosun\Desktop\新しいフォルダー\2\", Force:=True)
'    Call FSSrv.CopyDirectory("C:\Users\kaosun\Desktop\新しいフォルダー\1", "C:\Users\kaosun\Desktop\新しいフォルダー\2\", Force:=True)
'    Debug.Print FSSrv.GetNewestFile("C:\Users\kaosun\Desktop\新しいフォルダー", "test[0-9]\.txt")
'    Debug.Print FSSrv.CreateBackupFile("C:\Users\kaosun\Desktop\新しいフォルダー\test1.txt", ".old")
    
'    Call ThisWorkbook.Activate
'    ActiveWindow.Visible = True
'
'    Dim WBSrv As IWorkbookService
'    Set WBSrv = New WorkbookService
'    Dim book_list As ObjectList
'    Set book_list = New ObjectList
'
'    Call book_list.Add(WBSrv.OpenWorkbook)
'    Debug.Print "open: " & Join(WBSrv.GetAllWorkbook, ", ")
'
'    'Call book_list.Add(WBSrv.OpenWorkbook)
'    'Debug.Print "open: " & Join(WBSrv.GetAllWorkbook, ", ")
'
'    'Call book_list.Update(0, WBSrv.SaveWorkbook(book_list.Item(0), "C:\Users\kaosun\Desktop\新しいフォルダー\test.xlsm"))
'    'Debug.Print "save: " & Join(WBSrv.GetAllWorkbook, ", ")
'
'    Call book_list.Add(WBSrv.OpenWorkbook("C:\Users\kaosun\Desktop\新しいフォルダー\test.xlsm"))
'    Debug.Print "open: " & Join(WBSrv.GetAllWorkbook, ", ")
'    Debug.Print WBSrv.ExistsWorkbook(book_list.Item(0))
'
'    Call WBSrv.CloseWorkbook(book_list.Item(0))
'    Debug.Print "close: " & Join(WBSrv.GetAllWorkbook, ", ")
'    Debug.Print WBSrv.ExistsWorkbook(book_list.Item(0))
'
'    Call WBSrv.CloseWorkbook(book_list.Item(1))
'    Debug.Print "close: " & Join(WBSrv.GetAllWorkbook, ", ")
    
'    Dim TFSrv As ITextFileService
'    Set TFSrv = New TextFileService
'
'    Dim text_file As ITextFileEntity
'    Set text_file = TFSrv.GetTextFileEntity("C:\Users\kaosun\Desktop\新しいフォルダー\test5.txt")
'    Call text_file.OpenFile(AsWrite:=True, AsAppend:=True)
'    Call text_file.WriteLine("test1")
'    Call text_file.WriteLine("test2")
'    Call text_file.WriteLine("test3")
'    Debug.Print text_file.IsEndOfFile
'    Call text_file.CloseFile
'
'    Call text_file.OpenFile
'    Do While Not text_file.IsEndOfFile
'        Debug.Print text_file.ReadLine
'    Loop
'    Call text_file.CloseFile

'    Dim src_book As Workbook
'    Dim dst_book As Workbook
'
'    Dim src_sheet As Worksheet
'
'    Set src_book = Workbooks.Add
'    Set dst_book = Workbooks.Add
'
'    Call ThisWorkbook.Activate
'
'    Dim WBSrv As WorkbookService
'    Set WBSrv = New WorkbookService
'    Call WBSrv.CopyWorksheet(SourceWorkbookName:=src_book.Name, DestinationWorksheetName:="Test", DestinationWorkbookName:=dst_book.Name)

'    Dim WBSrv As WorkbookService
'    Set WBSrv = New WorkbookService
'    Dim result_arr() As WorksheetRangeBounds
'    result_arr() = WBSrv.Find("aaa")
'
'    Dim range_item As Variant 'WorksheetRangeBounds
'    For Each range_item In result_arr
'        Debug.Print range_item.ToString()
'    Next range_item
    Dim usr_input As UserInputSheet
    Set usr_input = New UserInputSheet
    Call usr_input.Initialize("Sheet1")
    
    Debug.Print usr_input.GetItemRange("Item 2", "Item 2-2").ToString()
End Sub
