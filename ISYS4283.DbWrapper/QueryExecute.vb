Imports System.Data.SqlClient

Friend Class QueryExecute : Implements IQuery
    Public Sub Run(ByRef command As SqlCommand, Optional ByRef obj As Object = Nothing) Implements IQuery.Run
        command.ExecuteNonQuery()
    End Sub
End Class
