Imports System.Data.SqlClient

Module Module1
    Public ACLECAMP As String = CenterLibrary.DBConnection.MyLocal
    Sub Main()
        Dim tmpJob As Microsoft.SqlServer.Management.Smo.Agent.Job = FindSqlAgentJob("(DM)eCampaign_Contact_Open_Click")
        If tmpJob IsNot Nothing Then
            If tmpJob.CurrentRunStatus = Microsoft.SqlServer.Management.Smo.Agent.JobExecutionStatus.Executing Then
                tmpJob.Stop()
            End If
        End If
    End Sub

    Public Function FindSqlAgentJob(ByVal JobName As String) As Microsoft.SqlServer.Management.Smo.Agent.Job
        Dim srv As New Microsoft.SqlServer.Management.Smo.Server( _
        New Microsoft.SqlServer.Management.Common.ServerConnection(New SqlConnection(ACLECAMP)))
        If srv.JobServer.Jobs.Contains(JobName) Then
            Return srv.JobServer.Jobs(JobName)
        Else
            Return Nothing
        End If
    End Function
End Module
