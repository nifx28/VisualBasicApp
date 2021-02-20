Imports System.IO
Imports System.Reflection
Imports Microsoft.Office.Interop

Public Class ExcelForm

    Friend Shared Title As String

    Dim objApp As New Excel.Application
    Dim objBook As Excel._Workbook

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Title = Text

        Dim RootPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
        Dim FilePath As String = Path.Combine(RootPath, "活頁簿1.xlsx")

        If Not File.Exists(FilePath) Then

            ExcelFunctions.ShowError(Me, "找不到檔案！")

        Else

            objBook = objApp.Workbooks.Open(FilePath)

        End If

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If objBook IsNot Nothing Then
            objBook.Close()
        End If

    End Sub

    Private Sub ButtonRead_Click(sender As Object, e As EventArgs) Handles ButtonRead.Click

        If Not ExcelFunctions.ShowError(Me, objBook, "尚未載入檔案！") Then

            Return

        End If

        Dim objSheet As Excel._Worksheet
        Dim objRange As Excel.Range

        objSheet = objBook.Worksheets(1)
        objRange = CType(objSheet.Cells(1, 1), Excel.Range)

        ExcelFunctions.ShowInfo(Me, objRange.Text)

        objRange = Nothing
        objSheet = Nothing

    End Sub

    Private Sub ButtonWrite_Click(sender As Object, e As EventArgs) Handles ButtonWrite.Click

        If Not ExcelFunctions.ShowError(Me, objBook, "尚未載入檔案！") Then

            Return

        End If

        Dim objSheet As Excel._Worksheet
        Dim objRange As Excel.Range
        Dim objRandom As New Random

        objSheet = objBook.Worksheets(1)
        objRange = CType(objSheet.Cells(1, 1), Excel.Range)

        objRange.Value = "隨機數值 " & objRandom.Next()
        objBook.Save()

        objRange = Nothing
        objSheet = Nothing

    End Sub

End Class
