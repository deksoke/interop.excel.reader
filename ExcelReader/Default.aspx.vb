Partial Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub btReadData_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btReadData.Click
        Dim ds As New DataSet
        Dim dt As System.Data.DataTable

        'Support File Extensions - xls, xlsx, csv
        Using ex As New ExcelReader("C:/20160401.xls")
            ds = ex.GetDataSet(False, True)
            dt = ds.Tables(0)
        End Using

        dgv.DataSource = dt
        dgv.DataBind()
    End Sub
End Class