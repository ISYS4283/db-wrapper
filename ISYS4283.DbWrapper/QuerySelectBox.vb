Imports System.Data.SqlClient

' used for filling multiple form controls, such as ComboBox and ListBox
Friend Class QuerySelectBox : Implements IQuery
    Public Sub Run(ByRef command As SqlCommand, Optional ByRef obj As Object = Nothing) Implements IQuery.Run
        Dim dataset As New DataSet
        Dim queryDataSet As New QueryDataSet

        queryDataSet.Run(command, dataset)

        If dataset.Tables.Count > 0 And dataset.Tables(0).Columns.Count > 1 Then
            obj.DataSource = dataset.Tables(0).DefaultView
            obj.DisplayMember = dataset.Tables(0).Columns(0).ToString
            obj.ValueMember = dataset.Tables(0).Columns(1).ToString
        End If
    End Sub
End Class
