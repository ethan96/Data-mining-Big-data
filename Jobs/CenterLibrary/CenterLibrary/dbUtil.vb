Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Configuration
Imports System.IO

Public Class dbUtil
    Public Shared Function dbGetDataTable( _
    ByVal ConnectionName As String, _
    ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New SqlConnection(ConnectionName)
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(strSqlCmd, g_adoConn)
        da.SelectCommand.CommandTimeout = 5 * 60
        Try
            da.Fill(dt)
        Catch ex As Exception
            g_adoConn.Close() : Throw New Exception(ex.ToString() + vbTab + "sql:" + strSqlCmd)
        End Try
        g_adoConn.Close() : g_adoConn = Nothing
        Return dt
    End Function

    Public Shared Function dbGetReaderAsync( _
    ByVal ConnectionName As String, _
    ByVal strSqlCmd As String, ByRef g_adoConn As SqlConnection, ByRef dbCmd As SqlCommand) As IAsyncResult
        g_adoConn = New SqlConnection(ConnectionName)
        dbCmd = g_adoConn.CreateCommand()
        dbCmd.Connection = g_adoConn : dbCmd.CommandText = strSqlCmd
        dbCmd.CommandTimeout = 500 * 1000
        g_adoConn.Open()
        Return dbCmd.BeginExecuteReader()
    End Function

    Public Shared Function dbExecuteNoQuery2(ByVal ConnectionStringName As String, ByVal strSqlCmd As String, Optional ByVal Parameters As SqlParameter() = Nothing) As Integer
        Dim g_adoConn As New SqlConnection(ConnectionStringName)
        Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
        dbCmd.Connection = g_adoConn : dbCmd.CommandText = strSqlCmd
        dbCmd.CommandTimeout = 500 * 1000
        Dim retInt As Integer = -1
        If Parameters IsNot Nothing AndAlso Parameters.Length > 0 Then
            dbCmd.Parameters.AddRange(Parameters)
        End If
        For i As Integer = 0 To 3
            Try
                g_adoConn.Open()
                Exit For
            Catch ex As SqlException
                If i = 3 Then Throw ex
                Threading.Thread.Sleep(100)
            End Try
        Next
        Try
            retInt = dbCmd.ExecuteNonQuery()
        Catch ex As Exception
            g_adoConn.Close() : Throw New Exception(ex.ToString + vbTab + "sql:" + strSqlCmd)
        End Try
        g_adoConn.Close() : Return retInt
    End Function
    Public Shared Function dbExecuteNoQuery( _
    ByVal ConnectionStringName As String, _
    ByVal strSqlCmd As String) As Integer
        Dim g_adoConn As New SqlConnection(ConnectionStringName)
        Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
        dbCmd.Connection = g_adoConn : dbCmd.CommandText = strSqlCmd
        dbCmd.CommandTimeout = 500 * 1000
        Dim retInt As Integer = -1
        For i As Integer = 0 To 3
            Try
                g_adoConn.Open()
                Exit For
            Catch ex As SqlException
                If i = 3 Then Throw ex
                Threading.Thread.Sleep(100)
            End Try
        Next
        'Using tran As SqlTransaction = g_adoConn.BeginTransaction
        Try
            'dbCmd.Transaction = tran
            retInt = dbCmd.ExecuteNonQuery()
            'tran.Commit()
        Catch ex As Exception
            'tran.Rollback()
            g_adoConn.Close() : Throw New Exception(ex.ToString + " sql:" + strSqlCmd)
        End Try
        'End Using
        g_adoConn.Close() : Return retInt
    End Function
    Public Shared Function dbExecuteNoQueryAsync( _
    ByVal ConnectionStringName As String, _
    ByVal strSqlCmd As String, ByRef g_adoConn As SqlConnection, ByRef dbCmd As SqlCommand) As IAsyncResult
        g_adoConn = New SqlConnection(ConnectionStringName)
        dbCmd = g_adoConn.CreateCommand()
        dbCmd.Connection = g_adoConn : dbCmd.CommandText = strSqlCmd
        Dim ar As IAsyncResult = Nothing
        g_adoConn.Open()
        Return dbCmd.BeginExecuteNonQuery()
    End Function
    Public Shared Function dbExecuteScalar(ByVal ConnectionName As String, ByVal strSqlCmd As String) As Object
        Dim g_adoConn As New SqlConnection(ConnectionName)
        For i As Integer = 0 To 3
            Try
                g_adoConn.Open()
                Exit For
            Catch ex As SqlException
                If i = 3 Then Throw ex
                Threading.Thread.Sleep(100)
            End Try
        Next
        Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
        dbCmd.CommandType = CommandType.Text : dbCmd.CommandText = strSqlCmd : dbCmd.CommandTimeout = 5 * 60
        Dim retObj As Object = Nothing
        Try
            retObj = dbCmd.ExecuteScalar()
        Catch ex As Exception
            g_adoConn.Close() : Throw ex
        End Try
        g_adoConn.Close() : Return retObj
    End Function


End Class

Public Class WriteLog
    Public Solution As String
    Public FunctionName As String
    Public StepName As String
    Public Message As String

    Public Sub WriteSystemLog()
        dbUtil.dbExecuteNoQuery(DBConnection.MyLocal, "insert into MY_LOG (SOLUTION_NAME,FUNCTION_NAME,STEP_NAME,MESSAGE,TIMESTAMP) values ('" + Solution.Replace("'", "''") + "','" + FunctionName.Replace("'", "''") + "','" + StepName.Replace("'", "''") + "','" + Message.Replace("'", "''") + "',GetDate())")
    End Sub

    Public Sub WriteErrorLog()
        dbUtil.dbExecuteNoQuery(DBConnection.MyLocal, "insert into MY_ERR_LOG (APPID,URL,QSTRING,EXMSG,ERRDATE) values ('" + Solution.Replace("'", "''") + "','" + FunctionName.Replace("'", "''") + "','" + StepName.Replace("'", "''") + "','" + Message.Replace("'", "''") + "',GetDate())")
    End Sub
End Class