Imports Oracle.ManagedDataAccess.Client


Public Class clsOracleDB


    Dim oraCN As OracleConnection, oraDA As OracleDataAdapter, oraCMD As OracleCommand

    'Dim intBlkRowsCopied As Long

    Dim objTrans As OracleTransaction

    Public Sub New(ByVal dataSource As String, ByVal userID As String, ByVal password As String)
        Try

            oraCN = New OracleConnection
            'oraCN.ConnectionString = "Data Source=" & strOracleDataSource & ";User Id=" & strOracleUserID & ";Password=" & strOraclePswrd
            oraCN.ConnectionString = "Data Source=" & dataSource & ";User Id=" & userID & ";Password=" & password
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Public Sub OpenDB()
        Try

            If oraCN.State = ConnectionState.Closed Then
                oraCN.Open()
            End If

        Catch ex As Exception
            Throw ex
            'Throw New Exception("Unable to connect to Oracle datasource: " & strOracleDataSource)
        End Try
    End Sub

    Public Sub CloseDB()
        Try

            If oraCN.State = ConnectionState.Open Then
                oraCN.Close()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ExecuteNonQuery(ByVal strQuery As String)
        Try
            oraCMD = New OracleCommand

            With oraCMD
                .Connection = oraCN
                .CommandType = CommandType.Text
                .CommandText = strQuery
                .CommandTimeout = 120
                .ExecuteNonQuery()
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ExecuteNonQuerySP(ByVal strQuery As String)
        Try
            oraCMD = New OracleCommand

            With oraCMD
                .Connection = oraCN
                .CommandType = CommandType.StoredProcedure
                .CommandText = strQuery
                .CommandTimeout = 120
                .ExecuteNonQuery()
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Public Function GetDataSet(ByVal strSQLQuery As String) As DataSet
        Try
            oraCMD = New OracleCommand
            Dim sqlDA As New OracleDataAdapter, ds As New DataSet

            With oraCMD
                .Connection = oraCN
                .CommandType = CommandType.Text
                .CommandText = strSQLQuery
                sqlDA.SelectCommand = oraCMD
                sqlDA.Fill(ds)
                Return ds
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetServerDate() As String
        Try
            oraCMD = New OracleCommand
            Dim sqlDA As New OracleDataAdapter, ds As New DataSet

            With oraCMD
                .Connection = oraCN
                .CommandType = CommandType.Text
                .CommandText = "select getdate()"
                sqlDA.SelectCommand = oraCMD
                sqlDA.Fill(ds)

                Return ds.Tables(0).Rows(0).Item(0)
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Public Sub BulkInsert(ByVal dt As DataTable, ByVal strTableName As String)
    '    Try

    '        blk = New SqlClient.SqlBulkCopy(oraCN)

    '        blk.BatchSize = 5000
    '        blk.NotifyAfter = 1

    '        blk.DestinationTableName = strTableName
    '        blk.WriteToServer(dt)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Sub blk_SqlRowsCopied(ByVal sender As Object, ByVal e As System.Data.SqlClient.SqlRowsCopiedEventArgs) Handles blk.SqlRowsCopied
    '    Try
    '        intBlkRowsCopied = e.RowsCopied
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub


    Public Sub BulkInsert(ByVal tableName As String, ByVal dataTable As System.Data.DataTable)

        'Dim sqlConnection As New OracleConnection(connectionString)
        Dim sqlCommand As New OracleCommand() With {.Connection = oraCN}

        Try
            'sqlConnection.Open()

            ' Build insert statement with placeholders for each column
            Dim sqlInsert As String = "INSERT INTO " & tableName & " ("
            For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                sqlInsert += dataTable.Columns(colIndex).ColumnName & ","
            Next

            sqlInsert = sqlInsert.TrimEnd(",") & ") VALUES ("
            For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                sqlInsert += ":" & dataTable.Columns(colIndex).ColumnName & ","
            Next
            sqlInsert = sqlInsert.TrimEnd(",") & ")"

            sqlCommand.CommandText = sqlInsert

            '' Create parameters for each column
            'Dim parameters As New List(OracleParameter)
            'For colIndex As Integer = 0 To dataTable.Columns.Count - 1
            '    parameters.Add(New OracleParameter(":" & dataTable.Columns(colIndex).ColumnName, dataTable.Columns(colIndex).DataType))
            'Next

            '' Add parameters to command
            'sqlCommand.Parameters.AddRange(parameters.ToArray())

            Dim param As OracleParameter
            For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                ' param = New OracleParameter(":" & dataTable.Columns(colIndex).ColumnName, dataTable.Columns(colIndex).DataType)
                'param = New OracleParameter(":" & dataTable.Columns(colIndex).ColumnName, OracleDbType.Varchar2)
                param = New OracleParameter(":" & dataTable.Columns(colIndex).ColumnName, OracleDataType(dataTable.Columns(colIndex).DataType.ToString))
                sqlCommand.Parameters.Add(param)

            Next


            ' Loop through each data row and execute insert statement
            For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                ' Set parameter values for each cell in the current row
                For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                    sqlCommand.Parameters(colIndex).Value = dataTable.Rows(rowIndex).Item(colIndex)
                Next

                sqlCommand.ExecuteNonQuery()
            Next

            'Console.WriteLine("Data imported successfully to table: " & tableName)
        Catch ex As Exception
            Throw ex
            'Finally
            'If sqlConnection.State = ConnectionState.Open Then
            '    sqlConnection.Close()
            'End If
        End Try
    End Sub

    Public Function OracleDataType(ByVal strSystemDataType As String) As OracleDbType
        Try
            Select Case strSystemDataType

                Case "System.Decimal"
                    Return OracleDbType.Decimal

                Case "System.DateTime"
                    Return OracleDbType.Date

                Case Else
                    Return OracleDbType.Varchar2

            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
