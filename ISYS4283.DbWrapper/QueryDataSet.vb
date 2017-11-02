Imports System.Data.SqlClient
Imports System.Windows.Forms

Friend Class QueryDataSet : Implements IQuery
    Protected Sub Fill(ByRef command As SqlCommand, ByRef dataset As DataSet)
        Dim adapter As New SqlDataAdapter(command)
        adapter.Fill(dataset)
    End Sub

    Public Sub Run(ByRef command As SqlCommand, Optional ByRef obj As Object = Nothing) Implements IQuery.Run
        Fill(command, obj)
    End Sub
End Class
