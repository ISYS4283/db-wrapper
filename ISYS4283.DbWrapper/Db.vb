Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class Db
    ' login to database server
    Protected connection As New SqlConnection With {
        .ConnectionString = "Server=essql1.walton.uark.edu;Database=isys4283-2017fa;Trusted_Connection=yes;"
    }
    ' prepare a query
    Protected command As New SqlCommand With {.Connection = connection}

    ' set and get sql command
    Public Property Sql() As String
        Get
            Return command.CommandText
        End Get

        Set(value As String)
            command.CommandText = value
        End Set
    End Property

    ' populate a data grid view
    Public Sub Fill(ByRef dgv As DataGridView)
        Dim adapter As New SqlDataAdapter(command)
        Dim dataset As New DataSet

        ' if anything goes wrong,
        ' then we still need to close the connection
        Try
            connection.Open()
            adapter.Fill(dataset)
        Catch ex As Exception
            ' simulate logger
            MsgBox(ex.Message)
            Throw ex
        Finally
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
        End Try

        ' fill the DataGridView with first query result set
        If dataset.Tables.Count > 0 Then
            dgv.Refresh()
            dgv.DataSource = dataset.Tables(0)
        End If
    End Sub
End Class