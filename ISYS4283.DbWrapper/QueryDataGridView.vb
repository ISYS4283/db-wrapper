Imports System.Data.SqlClient
Imports System.Windows.Forms

Friend Class QueryDataGridView : Implements IQuery
    Protected Sub Fill(ByRef command As SqlCommand, ByRef dgv As DataGridView)
        Dim adapter As New SqlDataAdapter(command)
        Dim dataset As New DataSet

        adapter.Fill(dataset)

        ' fill the DataGridView with first query result set
        If dataset.Tables.Count > 0 Then
            dgv.Refresh()
            dgv.DataSource = dataset.Tables(0)
        End If
    End Sub

    Public Sub Run(ByRef command As SqlCommand, Optional ByRef obj As Object = Nothing) Implements IQuery.Run
        Fill(command, obj)
    End Sub
End Class
