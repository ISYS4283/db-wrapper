Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class Db
    ' login to database server
    Protected connection As New SqlConnection With {
        .ConnectionString = "Server=essql1.walton.uark.edu;Database=isys4283-2017fa;Trusted_Connection=yes;"
    }
    ' prepare a query
    Protected command As New SqlCommand With {.Connection = connection}

    ' set and get connection
    Public Property ConnectionString() As String
        Get
            Return connection.ConnectionString
        End Get

        Set(value As String)
            connection.ConnectionString = value
        End Set
    End Property

    ' set and get sql command
    Public Property Sql() As String
        Get
            Return command.CommandText
        End Get

        Set(value As String)
            command.CommandText = value
        End Set
    End Property

    ' prevent sql injection
    Public Sub Bind(ByRef parameter As String, ByRef value As Object)
        command.Parameters.AddWithValue(parameter, value)
    End Sub

    Public Sub Execute()
        Run(New RunDelegate(AddressOf RunExecute))
    End Sub

    Protected Sub RunExecute()
        command.ExecuteNonQuery()
    End Sub

    ' populate a data grid view
    Public Sub Fill(ByRef dgv As DataGridView)
        Run(New RunDelegate(AddressOf RunFill), dgv)
    End Sub

    Protected Sub RunFill(ByRef dgv As DataGridView)
        Dim adapter As New SqlDataAdapter(command)
        Dim dataset As New DataSet

        adapter.Fill(dataset)

        ' fill the DataGridView with first query result set
        If dataset.Tables.Count > 0 Then
            dgv.Refresh()
            dgv.DataSource = dataset.Tables(0)
        End If
    End Sub

    Protected Delegate Sub RunDelegate(ByRef obj As Object)

    Protected Sub Run(ByRef del As [Delegate], Optional ByRef obj As Object = Nothing)
        ' if anything goes wrong,
        ' then we still need to close the connection
        ' https://stackoverflow.com/a/28483789/4233593
        Try
            Try
                connection.Open()
                del.DynamicInvoke(obj)
            Catch exception As Exception
                Log(exception)
                Throw
            End Try
        Finally
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
        End Try

        ' reset for next query
        command.CommandText = Nothing
        command.Parameters.Clear()
    End Sub

    ' override this method for real logger
    Protected Sub Log(ByRef exception As Exception)
        MsgBox(exception.Message)
    End Sub
End Class