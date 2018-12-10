Imports RDotNet

Module Module1
    Dim CurationConnStr As String = "Data Source=ACLSQL6\SQl2008R2;Initial Catalog=CurationPool;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Dim ACLECDMConnStr As String = "Data Source=aclecampaign\MATEST;Initial Catalog=DataMining;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Dim eStoreConnStr As String = "Data Source=172.21.1.20;Initial Catalog=eStoreProduction;Persist Security Info=True;User ID=estore3test;Password=estore3test;async=true;Connect Timeout=180;pooling='true'"
    Sub Main()

        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim dirPath = System.IO.Path.GetDirectoryName(dirInfo.FullName)

        Dim AllForecasts As New DataTable

        REngine.SetEnvironmentVariables()
        Dim RE As REngine = REngine.GetInstance()
        RE.Evaluate("library(forecast)")

        Dim dtPNList As New DataTable
        Dim aptPNList As New SqlClient.SqlDataAdapter(
            " select distinct top 3 a.MODEL_NO, a.PART_NO " +
            " from MyAdvantechGlobal.dbo.SAP_PRODUCT a (nolock) inner join MyAdvantechGlobal.dbo.SAP_PRODUCT_ORG b (nolock) on a.PART_NO=b.PART_NO  " +
            " inner join MyAdvantechGlobal.dbo.SAP_PRODUCT_ABC c (nolock) on a.PART_NO=c.PART_NO  " +
            " where b.ORG_ID='EU10' and a.MATERIAL_GROUP in ('PRODUCT') and a.MODEL_NO<>'' and isnumeric(a.MODEL_NO)=0 and c.PLANT='EUH1' " +
            " and b.STATUS in ('A','S5','N') and LEFT(c.ABC_INDICATOR,1) in ('A','B','C','D','V') and isnumeric(a.MODEL_NO)=0 " +
            " order by a.MODEL_NO  ", CurationConnStr)
        'aptPNList.SelectCommand.Parameters.AddWithValue("MODEL", ModelNo)
        aptPNList.Fill(dtPNList)
        aptPNList.SelectCommand.Connection.Close()

        For Each PNRow As DataRow In dtPNList.Rows
            Dim HistoryFromDate As New Date(2014, 1, 1), ForeHead As Integer = 4, HistoryEndDate As New Date(2015, 8, 31)
            Dim HistoryMonths As Integer = DateDiff(DateInterval.Month, HistoryFromDate, HistoryEndDate)
            Dim PartNo As String = PNRow.Item("PART_NO")
            Dim ModelNo As String = PNRow.Item("MODEL_NO")

            Console.WriteLine("Model:{0} Part:{1}", ModelNo, PartNo)

            Dim OrderHistoryDt As DataTable = GetOrderHistory(PartNo)
            Dim QuoteHistory As DataTable = GetEQ(PartNo)
            Dim GoogleViewModelHistoryDt As DataTable = GetGoogleModelView(ModelNo)
            Dim eStoreViewModelDt As DataTable = GetEstoreModelView(ModelNo, PartNo)

            Dim HistoryFactRecords As New List(Of R_Arima_Forecast.HistoryFactRecord)
            Dim HistoryRegressionInputRecords As New List(Of R_Arima_Forecast.RegressionInputRecord)
            Dim FutureRegressionInputRecords As New List(Of R_Arima_Forecast.RegressionInputRecord)


            For w As Integer = 1 To HistoryMonths
                Dim hisRec As New R_Arima_Forecast.HistoryFactRecord()
                hisRec.HistoryValue = 0 : hisRec.HistoryDate = HistoryFromDate.Date
                Dim orderRows() As DataRow = OrderHistoryDt.Select(String.Format("year={0} and month={1}", hisRec.HistoryDate.Year, hisRec.HistoryDate.Month))
                If orderRows.Length > 0 Then
                    hisRec.HistoryValue = orderRows(0).Item("qty")
                End If
                HistoryFactRecords.Add(hisRec)

                Dim hisRegRec As New R_Arima_Forecast.RegressionInputRecord()
                hisRegRec.HistoryDate = HistoryFromDate.Date
                'hisRegRec.Year = HistoryFromDate.Year : hisRegRec.Week = DatePart(DateInterval.WeekOfYear, HistoryFromDate)

                Dim inputValue1 As New R_Arima_Forecast.ColNameValue("www_model_view", 0)

                Dim gr() As DataRow = GoogleViewModelHistoryDt.Select(String.Format("year={0} and month={1}", hisRegRec.HistoryDate.Year, hisRegRec.HistoryDate.Month))
                If gr.Length > 0 Then inputValue1.Value = gr(0).Item("v")
                hisRegRec.InputValues.Add(inputValue1)

                Dim inputValue2 As New R_Arima_Forecast.ColNameValue("quote_qty", 0)
                Dim qr() As DataRow = QuoteHistory.Select(String.Format("year={0} and month={1}", hisRegRec.HistoryDate.Year, hisRegRec.HistoryDate.Month))
                If qr.Length > 0 Then
                    inputValue2.Value = qr(0).Item("qty")
                End If
                hisRegRec.InputValues.Add(inputValue2)

                Dim inputValue3 As New R_Arima_Forecast.ColNameValue("estore_model_view", 0)
                Dim er() As DataRow = eStoreViewModelDt.Select(String.Format("year={0} and month={1}", hisRegRec.HistoryDate.Year, hisRegRec.HistoryDate.Month))
                If er.Length > 0 Then
                    inputValue3.Value = er(0).Item("v")
                End If
                hisRegRec.InputValues.Add(inputValue3)


                HistoryRegressionInputRecords.Add(hisRegRec)
                HistoryFromDate = DateAdd(DateInterval.Month, 1, HistoryFromDate)
            Next

            For w As Integer = 1 To ForeHead
                Dim fuRegRec As New R_Arima_Forecast.RegressionInputRecord()
                fuRegRec.HistoryDate = HistoryFromDate.Date
                fuRegRec.Year = HistoryFromDate.Year : fuRegRec.Week = DatePart(DateInterval.WeekOfYear, HistoryFromDate)
                Dim inputValue1 As New R_Arima_Forecast.ColNameValue("www_model_view", _
                                                                     Aggregate q In HistoryRegressionInputRecords Where q.HistoryDate.Month = fuRegRec.HistoryDate.Month Into Average(CDbl(q.InputValues(0).Value)))
                fuRegRec.InputValues.Add(inputValue1)
                Dim inputValue2 As New R_Arima_Forecast.ColNameValue("quote_qty", _
                                                                     Aggregate q In HistoryRegressionInputRecords Where q.HistoryDate.Month = fuRegRec.HistoryDate.Month Into Average(CDbl(q.InputValues(1).Value)))
                fuRegRec.InputValues.Add(inputValue2)
                Dim inputValue3 As New R_Arima_Forecast.ColNameValue("estore_model_view", _
                                                                     Aggregate q In HistoryRegressionInputRecords Where q.HistoryDate.Month = fuRegRec.HistoryDate.Month Into Average(CDbl(q.InputValues(2).Value)))
                fuRegRec.InputValues.Add(inputValue3)
                FutureRegressionInputRecords.Add(fuRegRec)
                HistoryFromDate = DateAdd(DateInterval.Month, 1, HistoryFromDate)
            Next

            Dim ForecastRecords = R_Arima_Forecast.InvokeR_ArimaForecast(RE, 2014, 1, HistoryFactRecords, HistoryRegressionInputRecords, FutureRegressionInputRecords, ForeHead)
            Dim dtFore As DataTable = ListHelper.ListToDataTable(ForecastRecords)
            dtFore.Columns.Add("PART_NO") : dtFore.Columns.Add("ForecastMonth")
            For Each r As DataRow In dtFore.Rows
                r.Item("PART_NO") = PartNo
                r.Item("ForecastMonth") = CDate(r.Item("ForecastDate")).ToString("yyyy-MM")
            Next
            AllForecasts.Merge(dtFore)
          
          
            'Threading.Thread.Sleep(500)

            RE.Evaluate("rm(list=ls(all=TRUE))")
            'RE.Evaluate("library(forecast)")
        Next


        RE.Dispose()

        With AllForecasts.Columns
            .Remove("Low80") : .Remove("High80") : .Remove("Low95") : .Remove("High95") : .Remove("ForecastDate")
            .Remove("Year") : .Remove("Week")
        End With


        NPOIXlsUtil.RenderDataTableToExcel(AllForecasts, dirPath + "\" + String.Format("AEU_ALL_ModelForecast_{0}.xls", Now.ToString("yyyyMMdd")))

        Console.WriteLine("done")
        Console.ReadKey()

    End Sub

    Function GetOrderHistory(PartNo As String) As DataTable
        Dim sql As String = _
            " select YEAR(a.order_date) as [year], DATEPART(month, a.order_date) as [month], cast(SUM(a.Qty) as int) as qty  " + _
            " from MyAdvantechGlobal.dbo.EAI_SALE_FACT a (nolock) inner join MyAdvantechGlobal.dbo.SAP_PRODUCT b (nolock) on a.item_no=b.part_no " + _
            " where a.org='EU10' " + _
            " and a.fact_1234=1 and a.order_date between '2014-1-1' and '2015-08-31' and a.Qty>0 " + _
            " and (a.item_no=@PN) and b.material_group='PRODUCT' " + _
            " group by YEAR(a.order_date), DATEPART(month, a.order_date) " + _
            " order by YEAR(a.order_date), DATEPART(month, a.order_date) "
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(sql, CurationConnStr)
        apt.SelectCommand.Parameters.AddWithValue("PN", PartNo)
        'apt.SelectCommand.Parameters.AddWithValue("MN", ModelNo)
        apt.Fill(dt)
        apt.SelectCommand.Connection.Close()
        Return dt
    End Function

    Function GetEQ(PartNo As String) As DataTable
        Dim sql As String = _
            " select YEAR(a.createdDate) as [year], DATEPART(month,a.createdDate) as [month], SUM(b.qty) as qty  " + _
            " from eQuotation.dbo.QuotationMaster a (nolock) inner join eQuotation.dbo.QuotationDetail b (nolock) on a.quoteId=b.quoteId " + _
            " inner join MyAdvantechGlobal.dbo.SAP_PRODUCT c (nolock) on b.PartNo=c.PART_NO " + _
            " where a.org='EU10' and a.createdDate>='2014-01-01' and a.Createddate<='2015-08-31' and c.material_group='PRODUCT' " + _
            " and (b.PartNo=@PN) " + _
            " group by YEAR(a.createdDate), DATEPART(month,a.createdDate) " + _
            " order by YEAR(a.createdDate), DATEPART(month,a.createdDate) "
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(sql, CurationConnStr)
        apt.SelectCommand.Parameters.AddWithValue("PN", PartNo) 'apt.SelectCommand.Parameters.AddWithValue("MN", ModelNo)
        apt.Fill(dt)
        apt.SelectCommand.Connection.Close()
        Return dt
    End Function

    Function GetGoogleModelView(ModelNo As String) As DataTable
        Dim sql As String = _
            " select YEAR(a.RequestTime) as [year], DATEPART(month,a.RequestTime) as [month], count(*) as v " + _
            " from GOOGLE_VIEW_MODEL_REC a (nolock) " + _
            " where a.Model_No=@MODEL and a.country in " + _
            " ('Albania','Andorra','Armenia','Austria','Azerbaijan','Belarus','Belgium','Bosnia and Herzegovina'," + _
            "'Bulgaria','Croatia','Cyprus','Czech Republic','Denmark','Estonia','Faroe Islands','Finland','France'," + _
            "'Georgia','Germany','Gibraltar','Greece','Guernsey','Hungary','Iceland','Ireland','Isle of Man'," + _
            "'Italy','Jersey','Kazakhstan','Kosovo','Latvia') " + _
            " and a.RequestTime>='2014-01-01' and a.RequestTime<='2015-08-31' " + _
            " group by YEAR(a.RequestTime), DATEPART(month,a.RequestTime) " + _
            " order by YEAR(a.RequestTime), DATEPART(month,a.RequestTime) "
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(sql, ACLECDMConnStr)
        apt.SelectCommand.Parameters.AddWithValue("MODEL", ModelNo)
        apt.Fill(dt)
        apt.SelectCommand.Connection.Close()
        Return dt
    End Function

    Function GetEstoreModelView(ModelNo As String, PartNo As String) As DataTable
        Dim sql As String = _
            " select YEAR(a.CreatedDate) as [year], DATEPART(month, a.CreatedDate) as [month], COUNT(a.SessionID) as v " + _
            " from UserActivityLog a (nolock) " + _
            " where a.CreatedDate>='2013-01-01' and a.StoreID='AUS' " + _
            " and a.RemoteAddr not like '172.%' and a.UserId not like '%@advantech%.%' " + _
            " and a.ProductID is not null and (a.ProductID like '" + Replace(ModelNo, "'", "''") + "%' or a.ProductID=@PN) " + _
            " group by YEAR(a.CreatedDate), DATEPART(month, a.CreatedDate) " + _
            " order by YEAR(a.CreatedDate), DATEPART(month, a.CreatedDate) "
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(sql, eStoreConnStr)
        apt.SelectCommand.Parameters.AddWithValue("PN", PartNo)
        apt.Fill(dt)
        apt.SelectCommand.Connection.Close()
        Return dt
    End Function

End Module
