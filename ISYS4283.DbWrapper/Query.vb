Imports System.Data.SqlClient

Public Interface IQuery
    Sub Run(ByRef command As SqlCommand, Optional ByRef obj As Object = Nothing)
End Interface
