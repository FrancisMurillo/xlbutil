Attribute VB_Name = "TestBookUtil"
Public Sub TestGetSheetNames()
    VaseAssert.AssertArraysEqual _
        BookUtil.GetSheetNames(ActiveWorkbook), _
        Array("Butil")
End Sub

'# This also tests OpenBook and CloseBook
Public Sub TestCheckBook()
    VaseAssert.AssertFalse _
        BookUtil.CheckBook("C:\Nonexistant.xlsx")
        
    VaseAssert.AssertTrue _
        BookUtil.CheckBook(ActiveWorkbook.Path & Application.PathSeparator & _
                           "xlbutil-src" & Application.PathSeparator & "xlbutil-RELEASE.xlsm")
End Sub
