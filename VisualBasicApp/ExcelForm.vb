Imports System.IO
Imports System.Reflection
Imports Microsoft.Office.Interop

Public Class ExcelForm

    Friend Shared Title As String
    ReadOnly ExcelFile As String = "活頁簿2{0}.xlsm"

    Dim objApp As New Excel.Application
    Dim objBook As Excel.Workbook

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Title = Text
        Text &= " - " & ExcelFile

        Dim RootPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
        Dim FilePath As String = Path.Combine(RootPath, String.Format(ExcelFile, String.Empty))

        If Not File.Exists(FilePath) Then

            ExcelFunctions.ShowError(Me, "找不到檔案！")

        Else

            objBook = objApp.Workbooks.Open(FilePath)

        End If

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If objBook IsNot Nothing Then

            'objBook.Application.Visible = True
            objBook.Close(True)

        End If

    End Sub

    Private Sub ButtonRead_Click(sender As Object, e As EventArgs) Handles ButtonRead.Click

        If Not ExcelFunctions.ShowError(Me, objBook, "尚未載入檔案！") Then

            Return

        End If

        Dim objSheet As Excel.Worksheet
        Dim objRange As Excel.Range

        objSheet = objBook.Sheets(1)
        objRange = objSheet.Cells(1, 1)

        ExcelFunctions.ShowInfo(Me, "讀取 " & objRange.Text)

        objRange = Nothing
        objSheet = Nothing

    End Sub

    Private Sub ButtonWrite_Click(sender As Object, e As EventArgs) Handles ButtonWrite.Click

        If Not ExcelFunctions.ShowError(Me, objBook, "尚未載入檔案！") Then

            Return

        End If

        Dim objSheet As Excel.Worksheet
        Dim objRange As Excel.Range
        Dim objRandom As New Random

        objSheet = objBook.Sheets(1)
        objRange = objSheet.Cells(1, 1)

        objRange.Value = "隨機數值 " & objRandom.Next()
        objBook.Save()

        ExcelFunctions.ShowInfo(Me, "寫入 " & objRange.Text)

        objRange = Nothing
        objSheet = Nothing

    End Sub

    Private Sub ButtonTrigger_Click(sender As Object, e As EventArgs) Handles ButtonTrigger.Click

        If Not ExcelFunctions.ShowError(Me, objBook, "尚未載入檔案！") Then

            Return

        End If

        Dim objSheet As Excel.Worksheet
        Dim objShape As Excel.Shape

        objSheet = objBook.Sheets(1)
        objShape = objSheet.Shapes.Item("Button 1") ' 中文版顯示出來的為［按鈕1］

        If objShape.FormControlType <> Excel.XlFormControl.xlButtonControl Then

            ExcelFunctions.ShowError(Me, "此控制項並非為按鈕！")
            Return

        End If

        Dim objButton As Excel.Button

        objButton = CType(objShape.OLEFormat.Object, Excel.Button)

        Dim result As DialogResult = ExcelFunctions.ShowInfo(Me,
            "是否觸發［" & objButton.Text & "］按鈕的觸發巨集 """ & objButton.OnAction & """",
            MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then

            objSheet.Application.Run(objButton.OnAction)

            Dim sheetList As New List(Of String)

            For Each item As Excel.Worksheet In objBook.Sheets

                sheetList.Add(item.Name)

            Next

            ExcelFunctions.ShowInfo(Me, String.Join(Environment.NewLine, sheetList))

            'objBook.SaveCopyAs(String.Format(ExcelFile, "_已修改"))

        End If

    End Sub

End Class
