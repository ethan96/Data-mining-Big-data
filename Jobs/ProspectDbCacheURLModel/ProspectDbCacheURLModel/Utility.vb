Imports System.Text.RegularExpressions

Public Class Utility
    Public Shared strMyConn As String = CenterLibrary.DBConnection.MyAdvantechGlobal
    Public Shared strProsDbConn As String = CenterLibrary.DBConnection.CurationPool
    Public Shared Function GetMatchedModelsFromText(ByVal strText As String, ByVal strModels As String) As ArrayList
        Dim RegExp As New Regex(strModels, RegexOptions.IgnoreCase)
        Dim mc As MatchCollection = RegExp.Matches(strText), mmAry As New ArrayList
        For Each m As Match In mc
            If Not mmAry.Contains(m.Value) Then mmAry.Add(m.Value)
        Next
        Return mmAry
    End Function

    Public Shared Function GetAllModelStrings() As String
        Dim dt As New DataTable
        Dim strSql As String = _
            " select distinct model_no from MyAdvantechGlobal.dbo.SAP_PRODUCT where MODEL_NO<>'' and MODEL_NO is not null " + _
            " and LEN(model_no)>=4 and model_no not like '#%' and model_no not like '%''%' and model_no like '%-%' " + _
            " and model_no not like '%(%' and model_no not like '%)%' and model_no<>'NONE' and MATERIAL_GROUP in ('PRODUCT','BTOS') order by MODEL_NO "
        Dim apt As New SqlClient.SqlDataAdapter(strSql, strMyConn)
        apt.Fill(dt)
        apt.SelectCommand.Connection.Close()
        Dim arr As New ArrayList
        For Each r As DataRow In dt.Rows
            arr.Add(Replace(r.Item("model_no"), "'", "''"))
        Next
        Return String.Join("|", arr.ToArray())
    End Function
End Class
