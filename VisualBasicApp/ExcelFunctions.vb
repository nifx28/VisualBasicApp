Friend Class ExcelFunctions

    Friend Shared Sub ShowInfo(Self As Form, Message As String)

        MessageBox.Show(Self,
                        Message,
                        ExcelForm.Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

    End Sub

    Friend Shared Sub ShowError(Self As Form, Message As String)

        MessageBox.Show(Self,
                        Message,
                        ExcelForm.Title,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

    End Sub

    Friend Shared Function ShowError(Self As Form, TestObj As Object, Message As String)

        If TestObj Is Nothing Then

            ShowError(Self,
                      Message)
            Return False

        End If

        Return True

    End Function

End Class
