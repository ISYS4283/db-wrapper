Imports System.Data.SqlClient
Imports System.Windows.Forms

Public MustInherit Class Db
    ' login to database server
    Protected connection As New SqlConnection
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
    Public Sub Bind(ByVal parameter As String, ByVal value As Object)
        command.Parameters.AddWithValue(parameter, value)
    End Sub

    ' execute DML
    Public Sub Execute()
        Run(New QueryExecute)
    End Sub

    ' populate a generic data set
    Public Sub Fill(ByRef dataset As DataSet)
        Run(New QueryDataSet, dataset)
    End Sub

    Public Sub Fill(ByRef combobox As ComboBox)
        Run(New QuerySelectBox, combobox)
    End Sub

    Public Sub Fill(ByRef listbox As ListBox)
        Run(New QuerySelectBox, listbox)
    End Sub

    ' populate a data grid view
    Public Sub Fill(ByRef dgv As DataGridView)
        Run(New QueryDataGridView, dgv)
    End Sub

    Protected Sub Run(ByRef query As IQuery, Optional ByRef obj As Object = Nothing)
        ' if anything goes wrong,
        ' then we still need to close the connection
        ' https://stackoverflow.com/a/28483789/4233593
        Try
            Try
                connection.Open()
                query.Run(command, obj)
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
    Protected Overridable Sub Log(ByRef exception As Exception)
        MsgBox(exception.Message)
    End Sub
End Class
