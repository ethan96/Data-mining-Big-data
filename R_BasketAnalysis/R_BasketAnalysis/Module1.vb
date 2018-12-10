Imports System.Data.SqlClient
Imports RDotNet

Module Module1
    Public MyConnStr As String = "Data Source=ACLSQL6\SQL2008R2;Initial Catalog=MyAdvantechGlobal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!"
    Public MyLocalConnStr As String = "Data Source=aclecampaign\MATEST;Initial Catalog=MyLocal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"

    Sub Main()
        'GoogleAnalyticsSearchKeys()
        AlsoBuy()
        Console.WriteLine("done")
        Console.Read()
    End Sub

    Sub GoogleAnalyticsSearchKeys()
        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim dirPath = System.IO.Path.GetDirectoryName(dirInfo.FullName)
        Dim AnalysisName As String = "GASearchKeys"

        Dim myConn As New SqlConnection(MyLocalConnStr)
        Dim aptMy As New SqlDataAdapter( _
            " select distinct top 99999 REC_ID as TRAN_ID, KW as ITEM  " + _
            " from GOOGLE_WWW_SEARCH_KEYWORDS a (nolock)  " + _
            " where a.REC_ID in (select KEYID from GOOGLE_WWW_SEARCH_REC z (nolock) where z.dateHour>=GETDATE()-180) " + _
            " and a.REC_ID in (select REC_ID from GOOGLE_WWW_SEARCH_KEYWORDS group by REC_ID having COUNT(SEQ)>1) ", myConn)
        Dim dtOrderModel As New DataTable
        aptMy.Fill(dtOrderModel)
        aptMy.SelectCommand.Connection.Close()


        REngine.SetEnvironmentVariables()
        Dim RE As REngine = REngine.GetInstance()
        RE.Evaluate("library(arules)")

        Dim SoNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)
        Dim ModelNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)

        For rowIdx As Integer = 0 To dtOrderModel.Rows.Count - 1
            SoNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("TRAN_ID")
            ModelNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("ITEM")
        Next

        RE.SetSymbol("TID", SoNoVector) : RE.SetSymbol("item", ModelNoVector)

        RE.Evaluate("a_df3=data.frame(TID, item)")
        RE.Evaluate("trans4<-as(split(a_df3[,'item'], a_df3[,'TID']), 'transactions')")
        RE.Evaluate("rules <- apriori(trans4, parameter = list(minlen=2,supp =0.0003, conf = 0.0003))")
        'RE.Evaluate("rules <- apriori(trans4)")
        RE.Evaluate("rules_df<-as(rules, ""data.frame"")")
     

        Dim rules_df = RE.GetSymbol("rules_df").AsDataFrame()
        If rules_df.Count > 0 Then
            Dim rules = rules_df("rules").AsVector()
            Dim support = rules_df("support").AsVector()
            Dim confidence = rules_df("confidence").AsVector()
            Dim lift = rules_df("lift").AsVector()

            Dim RuleList As New List(Of DS.Apriori_Rule)
            Console.WriteLine("{0} rules", rules.Count)
            For idx As Integer = 0 To rules.Count - 1
                Dim arule1 As New DS.Apriori_Rule()
                Dim lrhs() = Split(rules(idx).ToString(), "=>")
                arule1.lhs = Trim(lrhs(0))
                arule1.lhs = arule1.lhs.Substring(1)
                arule1.lhs = arule1.lhs.Substring(0, arule1.lhs.Length - 1)
                arule1.rhs = Trim(lrhs(1))
                arule1.rhs = arule1.rhs.Substring(1)
                arule1.rhs = arule1.rhs.Substring(0, arule1.rhs.Length - 1)
                arule1.Support = support(idx)
                arule1.Confidence = confidence(idx)
                arule1.Lift = lift(idx)
                RuleList.Add(arule1)

            Next

            NPOIXlsUtil.RenderDataTableToExcel(ListHelper.ListToDataTable(RuleList), dirPath + "\" + String.Format("{0}_{1}.xls", AnalysisName, Now.ToString("yyyyMMdd")))
        Else
            Console.WriteLine("no rules found")
        End If




        RE.Dispose()
    End Sub

    Sub AlsoBuy()
        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim dirPath = System.IO.Path.GetDirectoryName(dirInfo.FullName)
        Dim AnalysisName As String = "AEU_IPPC-BTO_AlsoBuy"

        Dim myConn As New SqlConnection(MyConnStr)
        Dim aptMy As New SqlDataAdapter(
            " select distinct a.order_no, b.model_no  " +
            " from EAI_SALE_FACT a (nolock) inner join SAP_PRODUCT b (nolock) on a.item_no=b.PART_NO  " +
            " where a.org ='EU10' and a.efftive_date between '2012-01-01' and GETDATE()  " +
            " and a.Tran_Type='Shipment' and a.sector like '%AOnline%' " +
            " and b.MATERIAL_GROUP in ('PRODUCT','BTOS') and b.MODEL_NO<>'' and ISNUMERIC(b.MODEL_NO)=0 " +
            " and a.Qty>0 and a.item_no not like '%PS%ATX%' and a.egroup not in ('Advantech Global Services','Others','PAPS') " +
            " and a.order_no in (select order_no from eai_sale_fact z (nolock) where z.item_no like 'ARK%-BTO' group by order_no) " +
            " order by a.order_no ", myConn)

        aptMy.SelectCommand.CommandText =
            " select distinct a.order_no, b.model_no " +
            " from EAI_SALE_FACT a (nolock) inner join SAP_PRODUCT b (nolock) on a.item_no=b.PART_NO   " +
            " where a.org ='EU10' and a.efftive_date between '2012-01-01' and GETDATE()   " +
            " and a.Tran_Type='Shipment' and a.sector like '%AOnline%'  " +
            " and b.MATERIAL_GROUP in ('PRODUCT') and b.MODEL_NO<>'' and ISNUMERIC(b.MODEL_NO)=0  " +
            " and a.Qty>0 and a.item_no not like '%PS%ATX%' and a.egroup not in ('Advantech Global Services','Others','PAPS')  " +
            " and a.order_no in (select order_no from eai_sale_fact z (nolock) where z.item_no like 'IPPC%-BTO' and z.tr_line>=100 and z.tr_line % 100=0 group by order_no)  " +
            " and a.tr_line>100 " +
            " order by a.order_no "

        aptMy.SelectCommand.CommandTimeout = 99999
        Dim dtOrderModel As New DataTable
        aptMy.Fill(dtOrderModel)
        aptMy.SelectCommand.Connection.Close()


        REngine.SetEnvironmentVariables()
        Dim RE As REngine = REngine.GetInstance()
        RE.Evaluate("library(arules)")

        Dim SoNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)
        Dim ModelNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)

        For rowIdx As Integer = 0 To dtOrderModel.Rows.Count - 1
            SoNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("order_no")
            ModelNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("model_no")
        Next

        RE.SetSymbol("TID", SoNoVector) : RE.SetSymbol("item", ModelNoVector)

        RE.Evaluate("a_df3=data.frame(TID, item)")
        RE.Evaluate("trans4<-as(split(a_df3[,'item'], a_df3[,'TID']), 'transactions')")
        RE.Evaluate("rules <- apriori(trans4, parameter = list(minlen=2,supp =0.0065, conf = 0.007))")
        'RE.Evaluate("rules <- apriori(trans4)")
        RE.Evaluate("rules_df<-as(rules, ""data.frame"")")

        'RE.Evaluate("library(arulesViz)")
        'RE.Evaluate("jpeg('" + AnalysisName + ".jpg')")
        'RE.Evaluate("plot(rules, method=""graph"", control=list(type=""items""))")
        'RE.Evaluate("dev.off()")

        Dim rules_df = RE.GetSymbol("rules_df").AsDataFrame()
        If rules_df.Count > 0 Then
            Dim rules = rules_df("rules").AsVector()
            Dim support = rules_df("support").AsVector()
            Dim confidence = rules_df("confidence").AsVector()
            Dim lift = rules_df("lift").AsVector()

            Dim RuleList As New List(Of DS.Apriori_Rule)
            Console.WriteLine("{0} rules", rules.Count)
            For idx As Integer = 0 To rules.Count - 1
                Dim arule1 As New DS.Apriori_Rule()
                Dim lrhs() = Split(rules(idx).ToString(), "=>")
                arule1.lhs = Trim(lrhs(0))
                arule1.lhs = arule1.lhs.Substring(1)
                arule1.lhs = arule1.lhs.Substring(0, arule1.lhs.Length - 1)
                arule1.rhs = Trim(lrhs(1))
                arule1.rhs = arule1.rhs.Substring(1)
                arule1.rhs = arule1.rhs.Substring(0, arule1.rhs.Length - 1)
                arule1.Support = support(idx)
                arule1.Confidence = confidence(idx)
                arule1.Lift = lift(idx)
                RuleList.Add(arule1)


            Next

            NPOIXlsUtil.RenderDataTableToExcel(ListHelper.ListToDataTable(RuleList), dirPath + "\" + String.Format("{0}_{1}.xls", AnalysisName, Now.ToString("yyyyMMdd")))
        Else
            Console.WriteLine("no rules found")
        End If




        RE.Dispose()
    End Sub

    Sub AKR_FPM()
        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim dirPath = System.IO.Path.GetDirectoryName(dirInfo.FullName)
        Dim AnalysisName As String = "AKR_FPM-SYS"

        Dim myConn As New SqlConnection(MyConnStr)
        Dim aptMy As New SqlDataAdapter( _
            " select distinct a.SO_NO, a.PART_NO  " + _
            " from SAP_ORDER_HISTORY a (nolock) " + _
            " inner join " + _
            " ( " + _
            " 	select a.SO_NO, a.HIGHER_LEVEL  " + _
            " 	from SAP_ORDER_HISTORY a (nolock) " + _
            " 	where a.SALES_ORG='KR01' and a.PART_NO like 'FPM%' and a.HIGHER_LEVEL>0 " + _
            " 	group by a.SO_NO, a.HIGHER_LEVEL  " + _
            " ) b on a.SO_NO=b.SO_NO and (a.HIGHER_LEVEL=b.HIGHER_LEVEL or a.LINE_NO=b.HIGHER_LEVEL) " + _
            " inner join SAP_PRODUCT c (nolock) on a.PART_NO=c.PART_NO  " + _
            " where a.PART_NO not in ('PTRADE-BTO') and a.PART_NO not like 'AGS-CTOS-%' and a.PART_NO not like '%PS%ATX%' and c.material_group='PRODUCT' " + _
            " order by a.SO_NO, a.PART_NO  ", myConn)
        Dim dtOrderModel As New DataTable
        aptMy.Fill(dtOrderModel)
        aptMy.SelectCommand.Connection.Close()


        REngine.SetEnvironmentVariables()
        Dim RE As REngine = REngine.GetInstance()
        RE.Evaluate("library(arules)")

        Dim SoNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)
        Dim ModelNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)

        For rowIdx As Integer = 0 To dtOrderModel.Rows.Count - 1
            SoNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("SO_NO")
            ModelNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("PART_NO")
        Next

        RE.SetSymbol("TID", SoNoVector) : RE.SetSymbol("item", ModelNoVector)

        RE.Evaluate("a_df3=data.frame(TID, item)")
        RE.Evaluate("trans4<-as(split(a_df3[,'item'], a_df3[,'TID']), 'transactions')")
        RE.Evaluate("rules <- apriori(trans4, parameter = list(minlen=2,supp = 0.038, conf = 0.1))")
        RE.Evaluate("rules_df<-as(rules, ""data.frame"")")

        'RE.Evaluate("library(arulesViz)")
        'RE.Evaluate("jpeg('" + AnalysisName + ".jpg')")
        'RE.Evaluate("plot(rules, method=""graph"", control=list(type=""items""))")
        'RE.Evaluate("dev.off()")

        Dim rules_df = RE.GetSymbol("rules_df").AsDataFrame()

        If rules_df.RowCount = 0 Then
            Console.WriteLine("No rules found")
            'Console.Read()
        Else
            Dim rules = rules_df("rules").AsVector()
            Dim support = rules_df("support").AsVector()
            Dim confidence = rules_df("confidence").AsVector()
            Dim lift = rules_df("lift").AsVector()

            Dim RuleList As New List(Of DS.Apriori_Rule)
            Console.WriteLine("{0} rules", rules.Count)
            For idx As Integer = 0 To rules.Count - 1
                Dim arule1 As New DS.Apriori_Rule()
                Dim lrhs() = Split(rules(idx).ToString(), "=>")
                arule1.lhs = Trim(lrhs(0))
                arule1.lhs = arule1.lhs.Substring(1)
                arule1.lhs = arule1.lhs.Substring(0, arule1.lhs.Length - 1)
                arule1.rhs = Trim(lrhs(1))
                arule1.rhs = arule1.rhs.Substring(1)
                arule1.rhs = arule1.rhs.Substring(0, arule1.rhs.Length - 1)
                arule1.Support = support(idx)
                arule1.Confidence = confidence(idx)
                arule1.Lift = lift(idx)

                If arule1.lhs.Contains("FPM") Or arule1.rhs.Contains("FPM") Then
                    RuleList.Add(arule1)
                End If

            Next

            'NPOIXlsUtil.RenderDataTableToExcel(ListHelper.ListToDataTable(RuleList), dirPath + "\" + String.Format("{0}_{1}.xls", AnalysisName, Now.ToString("yyyyMMdd")))


            Dim SiebelContactDt As New DataTable
            'With SiebelContactDt.Columns
            '    .Add("WhenBuy") : .Add("AlsoBuy") : .Add("ERPID") : .Add("CompanyName") : .Add("Account Row Id") : .Add("Acocunt Status") : .Add("Email")
            'End With

            Dim cuApt As New SqlDataAdapter("", MyConnStr)
            cuApt.SelectCommand.CommandTimeout = 99999
            cuApt.SelectCommand.Connection.Open()
            For Each arule1 In RuleList
                Dim lhsPNs = Split(arule1.lhs, ","), rhsPNs = Split(arule1.rhs, ",")

                Dim lhsAry As New ArrayList, rhsAry As New ArrayList
                For Each l In lhsPNs
                    lhsAry.Add("'" + l + "'")
                Next

                For Each l In rhsPNs
                    rhsAry.Add("'" + l + "'")
                Next

                Dim lhsPNIn As String = String.Join(",", lhsAry.ToArray())
                Dim rhsPNIn As String = String.Join(",", rhsAry.ToArray())

                Dim sqlERPID As String = _
                    " select distinct '" + arule1.lhs + "' as WhenBuy, '" + arule1.rhs + "' as AlsoBuy, a.COMPANY_ID, b.COMPANY_NAME, c.ACCOUNT_ROW_ID, c.ACCOUNT_STATUS, c.EMAIL_ADDRESS, c.JOB_TITLE, c.WorkPhone    " + _
                    " from SAP_ORDER_HISTORY a (nolock) inner join SAP_DIMCOMPANY b (nolock) on a.COMPANY_ID=b.COMPANY_ID and b.ORG_ID='KR01' " + _
                    " left join SIEBEL_CONTACT c (nolock) on b.COMPANY_ID=c.ERPID and c.OrgID='AKR' " + _
                    " where a.SO_NO in " + _
                    " ( " + _
                    " 	select distinct so_no from SAP_ORDER_HISTORY (nolock) where PART_NO in (" + lhsPNIn + ") and SALES_ORG='KR01' " + _
                    " 	and SO_NO in (select distinct so_no from SAP_ORDER_HISTORY (nolock) where PART_NO in (" + rhsPNIn + ") and SALES_ORG='KR01') " + _
                    " ) " + _
                    " order by a.COMPANY_ID, c.EMAIL_ADDRESS   "

                Dim dtERPID As New DataTable
                cuApt.SelectCommand.CommandText = sqlERPID
                cuApt.Fill(dtERPID)
                SiebelContactDt.Merge(dtERPID)
            Next
            cuApt.SelectCommand.Connection.Close()

            NPOIXlsUtil.RenderDataTableToExcel(SiebelContactDt, dirPath + "\" + "AKR_FPM_SiebelContacts_20150805.xls")

        End If

        RE.Dispose()
    End Sub

    Sub AOLOrderBasketAnalysis()
        Dim myConn As New SqlConnection(MyConnStr)
        Dim aptMy As New SqlDataAdapter( _
            " select distinct a.order_no, a.MODEL_NO  " + _
            " from V_AOL_ORDER_V2 a (nolock)  " + _
            " where a.ORG_ID='US01' and a.ORDER_YEAR>=2014 " + _
            " order by a.order_no  ", myConn)
        Dim dtOrderModel As New DataTable
        aptMy.Fill(dtOrderModel)
        aptMy.SelectCommand.Connection.Close()

        REngine.SetEnvironmentVariables()
        Dim RE As REngine = REngine.GetInstance()
        RE.Evaluate("library(arules)")

        Dim SoNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)
        Dim ModelNoVector = RE.CreateCharacterVector(dtOrderModel.Rows.Count)

        For rowIdx As Integer = 0 To dtOrderModel.Rows.Count - 1
            SoNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("order_no")
            ModelNoVector(rowIdx) = dtOrderModel.Rows(rowIdx).Item("MODEL_NO")
        Next

        RE.SetSymbol("TID", SoNoVector) : RE.SetSymbol("item", ModelNoVector)

        RE.Evaluate("a_df3=data.frame(TID, item)")
        RE.Evaluate("trans4<-as(split(a_df3[,'item'], a_df3[,'TID']), 'transactions')")
        'RE.Evaluate("inspect(trans4)")
        RE.Evaluate("rules <- apriori(trans4, parameter = list(minlen=1, supp=0.005, conf=0.1))")
        RE.Evaluate("inspect(rules)")
        RE.Dispose()
    End Sub

End Module
