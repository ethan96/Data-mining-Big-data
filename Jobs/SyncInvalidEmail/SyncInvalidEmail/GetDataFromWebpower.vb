Imports System.IO
Imports System.Xml

Public Class GetDataFromWebpower

    Public Shared Sub GetHardBounceMail(ByRef dt As DataTable)
        Try
            Dim ExcuteDate = Convert.ToDateTime(ConfigurationManager.AppSettings("ExcuteDate"))

            If (ExcuteDate >= DateTime.Now.Date) Then
                Return
            End If


            Dim client As New WebpowerService.DMdeliverySoapAPIPortClient()
            Dim user As New WebpowerService.DMdeliveryLoginType() With {.username = ConfigurationManager.AppSettings("WebpowerUser"), .password = ConfigurationManager.AppSettings("WebpowerPwd")}

            Dim results As WebpowerService.MailingBounceType() = client.getMailingBounce(user, 1, 5, "hard,soft", "email", ExcuteDate)

            For Each result In results
                Dim dr = dt.NewRow()
                dr("EMAIL") = result.field
                dr("INS_DATE") = result.log_date
                dr("EMAIL_BODY") = result.message
                dr("REASON_FLAG") = String.Format("{0} Bounced", result.type.ToString())
                dt.Rows.Add(dr)
            Next

            GetDataFromWebpower.ModifyConfig()

        Catch ex As Exception

        End Try

    End Sub

    Private Shared Sub ModifyConfig()
        Dim configFullPath = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
        Dim configMap = New ExeConfigurationFileMap() With {.ExeConfigFilename = configFullPath}
        Dim config As System.Configuration.Configuration = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None)
        Dim h = ConfigurationManager.AppSettings("ExcuteDate")
        config.AppSettings.Settings("ExcuteDate").Value = DateTime.Now.Date
        config.Save()
    End Sub

End Class
