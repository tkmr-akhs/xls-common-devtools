Attribute VB_Name = "Test_TextFileServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! TextFileServiceTestDouble クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private TUtl As New UnitTestUtils

Public Sub Test_GetTextFileEntity_WithStubValue_ReturnsStub(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble
    
    Dim file_path As String
    file_path = "C:\Data\myfile.txt"
    
    ' 新しい TextFileEntityTestDouble を用意しておき、
    ' Dictionary に「file_path」をキーにして格納しておく。
    Dim expect_obj As TextFileEntityTestDouble
    Set expect_obj = New TextFileEntityTestDouble
    
    Call TUtl.SetValue(service_stub.GetTextFileEntity_Values, expect_obj, file_path)
    
    ' Act
    Dim actual_obj As ITextFileEntity
    Set actual_obj = service_stub.GetTextFileEntity(file_path)
    
    ' Assert
    ' 取り出せたオブジェクトが、Dictionary で格納したものと同一であればOK
    Assert.Equals expect_obj, actual_obj
End Sub

Public Sub Test_GetTextFileEntity_MultipleEntries_ReturnsCorrectStub(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils
    
    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble
    
    Dim file_path1 As String
    file_path1 = "C:\Data\myfile1.txt"
    
    Dim file_path2 As String
    file_path2 = "C:\Data\myfile2.txt"
    
    ' 新しい TextFileEntityTestDouble を用意しておき、
    ' Dictionary に「file_path」をキーにして格納しておく。
    Dim expect_obj1 As TextFileEntityTestDouble
    Set expect_obj1 = New TextFileEntityTestDouble
    
    Call TUtl.SetValue(service_stub.GetTextFileEntity_Values, expect_obj1, file_path1)
    
    Dim expect_obj2 As TextFileEntityTestDouble
    Set expect_obj2 = New TextFileEntityTestDouble
    
    Call TUtl.SetValue(service_stub.GetTextFileEntity_Values, expect_obj2, file_path2)
    
    ' Act
    Dim actual_obj1 As ITextFileEntity
    Set actual_obj1 = service_stub.GetTextFileEntity(file_path1)
    Dim actual_obj2 As ITextFileEntity
    Set actual_obj2 = service_stub.GetTextFileEntity(file_path2)
    
    ' Assert
    ' 取り出せたオブジェクトが、Dictionary で格納したものと同一であればOK
    Assert.Equals expect_obj1, actual_obj1
    Assert.Equals expect_obj2, actual_obj2
End Sub

