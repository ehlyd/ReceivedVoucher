Imports System.Data.SqlClient

Public Class clsSQLDB

    Dim sqlCN As SqlConnection, sqlDA As SqlDataAdapter, sqlCMD As SqlCommand
    Dim WithEvents blk As SqlClient.SqlBulkCopy
    Dim intBlkRowsCopied As Long

    Public Sub New()
        Try
            sqlCN = New SqlConnection
            sqlCN.ConnectionString = "Persist Security Info=False;Data Source=" & strServer & ";Initial Catalog=" & strDatabase & ";User ID=" & strUserID & ";Password=" & strPswrd

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub OpenDB()
        Try

            If sqlCN.State = ConnectionState.Closed Then
                sqlCN.Open()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub CloseDB()
        Try

            If sqlCN.State = ConnectionState.Open Then
                sqlCN.Close()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ExecuteNonQuery(ByVal strSQLQuery As String)
        Try
            sqlCMD = New SqlCommand

            With sqlCMD
                .Connection = sqlCN
                .CommandType = CommandType.Text
                .CommandText = strSQLQuery
                .CommandTimeout = 120
                .ExecuteNonQuery()
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function GetDataSet(ByVal strSQLQuery As String) As DataSet
        Try
            sqlCMD = New SqlCommand
            Dim sqlDA As New SqlDataAdapter, ds As New DataSet

            With sqlCMD
                .Connection = sqlCN
                .CommandType = CommandType.Text
                .CommandText = strSQLQuery
                sqlDA.SelectCommand = sqlCMD
                sqlDA.Fill(ds)
                Return ds
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetServerDate() As String
        Try
            sqlCMD = New SqlCommand
            Dim sqlDA As New SqlDataAdapter, ds As New DataSet

            With sqlCMD
                .Connection = sqlCN
                .CommandType = CommandType.Text
                .CommandText = "select getdate()"
                sqlDA.SelectCommand = sqlCMD
                sqlDA.Fill(ds)

                Return ds.Tables(0).Rows(0).Item(0)
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub BulkInsert(ByVal dt As DataTable, ByVal strTableName As String)
        Try

            blk = New SqlClient.SqlBulkCopy(sqlCN)

            blk.BatchSize = 5000
            blk.NotifyAfter = 1

            blk.DestinationTableName = strTableName
            blk.WriteToServer(dt)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub blk_SqlRowsCopied(ByVal sender As Object, ByVal e As System.Data.SqlClient.SqlRowsCopiedEventArgs) Handles blk.SqlRowsCopied
        Try
            intBlkRowsCopied = e.RowsCopied
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
