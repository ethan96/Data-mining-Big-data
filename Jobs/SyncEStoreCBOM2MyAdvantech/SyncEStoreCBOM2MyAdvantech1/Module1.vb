Imports System.IO

Module Module1

    Sub Main()

        Dim _DirectoryPath As String = My.Application.Info.DirectoryPath
        Dim _sqlfilepath As String = Path.Combine(_DirectoryPath, "SyncCtosSql.txt")
        Console.WriteLine(_sqlfilepath)
        Console.WriteLine("Does file exist? " & File.Exists(_sqlfilepath))
        'Console.ReadKey()
        'Exit Sub

        Try
            Dim sqlnew As New System.Text.StringBuilder
            sqlnew.AppendLine(" delete from EZ_CBOM_MAPPING where ismanual =0 ;")
            sqlnew.AppendLine(" INSERT INTO EZ_CBOM_MAPPING   SELECT newid() as row_id,   Product.DisplayPartno as number,  ")
            sqlnew.AppendLine(" Product_Ctos.BTONo as vnumber,0 as ismanual,SUBSTRING( Parts.StoreID , 2 , 2 ) AS ORG,GETDATE() as last_upd_date,'From Estore' as last_upd_by ")
            sqlnew.AppendLine(" FROM  eStoreProduction.dbo.Parts  ")
            sqlnew.AppendLine(" INNER JOIN eStoreProduction.dbo.Product ON Parts.StoreID = Product.StoreID  ")
            sqlnew.AppendLine(" AND Parts.SProductID = Product.SProductID  ")
            sqlnew.AppendLine(" INNER JOIN eStoreProduction.dbo.Product_Ctos ON Product.StoreID = Product_Ctos.StoreID ")
            sqlnew.AppendLine(" AND Product.SProductID = Product_Ctos.SProductID ")
            'sqlnew.AppendLine(" where  Product.[Status] != 'PHASED_OUT' order by Parts.StoreID,Product_Ctos.BTONo ")
            sqlnew.AppendLine(" where  Product.[Status] not in ('PHASED_OUT','INACTIVE','INACTIVE_AUTO','DELETED','SOLUTION_ONLY','TOBEREVIEW')   AND Product.PublishStatus = 1  order by Parts.StoreID,Product_Ctos.BTONo ")
            dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, sqlnew.ToString())
        Catch ex As Exception
            dbUtil.SendEmail("myadvantech@advantech.com", "ebiz.aeu@advantech.eu", "Sync New ESTORE_BTOS From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.ToString(), False)
            Exit Sub
        End Try

        Dim ESTORE_BTOS_CATEGORY As New System.Text.StringBuilder
        'ESTORE_BTOS_CATEGORY.AppendLine("Truncate table ESTORE_BTOS_CATEGORY;")
        ESTORE_BTOS_CATEGORY.AppendLine(" with CTECategory([Storeid] ,[CategoryID] ,[CategoryPath] ,[CategoryName] ,[LocalCategoryName] ,[ParentCategoryID] ,[CreatedDate],[CreatedBy],[CategoryType],[DisplayType] ")
        ESTORE_BTOS_CATEGORY.AppendLine(",[ImageURL],[Description] ,[ExtendedDescription],[Keywords] ,[Publish] ,[Sequence] ,[PageTitle],[PageDescription],[ProductDivision],[RuleSetId],[level],[catalog],CategoryNameRoot")
        ESTORE_BTOS_CATEGORY.AppendLine(" ) as (SELECT  [Storeid] ,[CategoryID] ,[CategoryPath],[CategoryName] ,[LocalCategoryName] ,[ParentCategoryID] ,[CreatedDate] ,[CreatedBy] ,[CategoryType] ,[DisplayType] ")
        ESTORE_BTOS_CATEGORY.AppendLine("  ,[ImageURL],[Description],[ExtendedDescription] ,[Keywords],[Publish],[Sequence],[PageTitle],[PageDescription],[ProductDivision] ,[RuleSetId] ")
        ESTORE_BTOS_CATEGORY.AppendLine(" ,0 as [level],[CategoryPath] as [catalog],CategoryName  AS CategoryNameRoot ")
        ESTORE_BTOS_CATEGORY.AppendLine("  FROM [ProductCategory] where  ParentCategoryID is null and Publish=1 and CategoryType='CTOSCategory' ")
        ESTORE_BTOS_CATEGORY.AppendLine(" union all ")
        ESTORE_BTOS_CATEGORY.AppendLine(" SELECT a.[Storeid] ,a.[CategoryID],a.[CategoryPath] ,a.[CategoryName] ,a.[LocalCategoryName],a.[ParentCategoryID] ,a.[CreatedDate] ")
        ESTORE_BTOS_CATEGORY.AppendLine("  ,a.[CreatedBy],a.[CategoryType] ,a.[DisplayType] ,a.[ImageURL] ,a.[Description] ,a.[ExtendedDescription] ,a.[Keywords] ")
        ESTORE_BTOS_CATEGORY.AppendLine(" ,a.[Publish],a.[Sequence] ,a.[PageTitle],a.[PageDescription] ,a.[ProductDivision] ,a.[RuleSetId],b.[level]+1 as [level] ")
        ESTORE_BTOS_CATEGORY.AppendLine(" ,b.[catalog] as [catalog],b.CategoryNameRoot FROM [ProductCategory] a inner join CTECategory b  ")
        ESTORE_BTOS_CATEGORY.AppendLine(" on a.[ParentCategoryID]=b. [CategoryID] and a.storeid=b.Storeid  ")
        ESTORE_BTOS_CATEGORY.AppendLine(" where   b.[level]<20 and   a.Publish=1 ) ")
        ESTORE_BTOS_CATEGORY.AppendLine(" SELECT distinct Product.StoreID,   CategoryNameRoot as CategoryName,Product.SProductID,Product_Ctos.BTONo,Product.DisplayPartno,GETDATE() as last_upd_date ")
        ESTORE_BTOS_CATEGORY.AppendLine("  FROM         Product_Ctos INNER JOIN ")
        ESTORE_BTOS_CATEGORY.AppendLine(" Product ON Product_Ctos.StoreID = Product.StoreID AND Product_Ctos.SProductID = Product.SProductID INNER JOIN ")
        ESTORE_BTOS_CATEGORY.AppendLine(" ProductCategroyMapping ON Product.StoreID = ProductCategroyMapping.StoreID AND Product.SProductID = ProductCategroyMapping.SProductID inner JOIN ")
        ESTORE_BTOS_CATEGORY.AppendLine(" CTECategory on CTECategory.CategoryID=ProductCategroyMapping.CategoryID and CTECategory.Storeid= ProductCategroyMapping.StoreID ")
        ESTORE_BTOS_CATEGORY.AppendLine("  where  Product.PublishStatus=1  and Product.StoreID in ('AUS','AAU','AJP','ATW','AKR') AND ") '2015/4/10 Add ATW
        ESTORE_BTOS_CATEGORY.AppendLine(" (dbo.Product.Status <> 'INACTIVE_AUTO') AND (dbo.Product.Status <> 'deleted') AND (dbo.Product.Status <> 'INACTIVE') AND (dbo.Product.Status <> 'PHASED_OUT') AND (dbo.Product.Status <> 'SOLUTION_ONLY') AND (dbo.Product.Status <> 'TOBEREVIEW') and (Product_Ctos.Assembly is null or Product_Ctos.Assembly=1) ")
        'ESTORE_BTOS_CATEGORY.AppendLine(" (dbo.Product.Status <> 'INACTIVE_AUTO') AND (dbo.Product.Status <> 'deleted') AND (dbo.Product.Status <> 'INACTIVE') AND (dbo.Product.Status <> 'PHASED_OUT') and (Product_Ctos.Assembly is null or Product_Ctos.Assembly=1) or (Product.SProductID='21431' and Product.StoreID='AUS') ")

        Try
            dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.eStore, "update CTOSComponentDetail set qty=1 where qty=0;") 'ICC To prevent qty = 0 CTOS bom in eStore. Update qty first.
            Dim ESTORE_BTOS_CATEGORYDT As DataTable = dbUtil.dbGetDataTable(CenterLibrary.DBConnection.eStore, ESTORE_BTOS_CATEGORY.ToString)
            If ESTORE_BTOS_CATEGORYDT.Rows.Count = 0 Then
                dbUtil.SendEmail("myadvantech@advantech.com", "ebiz.aeu@advantech.eu", "Sync ESTORE_BTOS_CATEGORY From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
                Exit Sub
            End If
            Dim DeleteSQL As String = "Truncate table ESTORE_BTOS_CATEGORY"
            Dim ReturnInt As Integer = dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, DeleteSQL)
            For Each DR As DataRow In ESTORE_BTOS_CATEGORYDT.Rows

                Dim Insertsql As New System.Text.StringBuilder
                With Insertsql
                    .AppendFormat(" INSERT INTO ESTORE_BTOS_CATEGORY ")
                    .AppendFormat(" (storeid,CategoryName,SProductID,BTONo,DisplayPartno,last_upd_date) ")
                    .AppendFormat(" VALUES ('{0}',N'{1}','{2}','{3}','{4}',{5}) ", DR.Item("storeid"),
                                  DR.Item("CategoryName"),
                                  DR.Item("SProductID"),
                                  DR.Item("BTONo"),
                                  DR.Item("DisplayPartno"),
                                  "GETDATE()")

                End With

                Console.WriteLine(String.Format("{0};", Insertsql.ToString))
                Try
                    dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, Insertsql.ToString)
                Catch ex As Exception
                    dbUtil.SendEmail("MyAdvantech@advantech.com", "MyAdvantech@advantech.com", "Sync CBOM From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.ToString(), False)
                End Try

            Next
            '''''''

        Catch ex As Exception
            dbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync ESTORE_BTOS_CATEGORY From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.ToString(), False)
            Exit Sub
        End Try
        ' MING TEST
        'dbUtil.SendEmail("tc.chen@advantech.com.tw,ming.zhao@advantech.com.cn", "ebiz.aeu@advantech.eu", "Sync CBOM From Estore OK " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
        'Exit Sub
        Dim srFile As StreamReader = Nothing
        Dim sql As String = String.Empty
        Try

           
            Console.WriteLine(_DirectoryPath)
            'srFile = New StreamReader("E:\Scheduled_Programs\SyncEStoreCBOM2MyAdvantech\SyncEStoreCBOM2MyAdvantech1\SyncCtosSql.txt", System.Text.Encoding.[Default])
            srFile = New StreamReader(_sqlfilepath, System.Text.Encoding.[Default])
            sql = srFile.ReadToEnd()
        Catch ex As Exception
            dbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync CBOM From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.Message.ToString(), False)
        Finally
            If srFile IsNot Nothing Then
                srFile.Dispose()
                srFile.Close()
            End If
        End Try
        'Console.WriteLine(msg)
        'Console.ReadLine()
        Dim EstoreCBOM As DataTable = dbUtil.dbGetDataTable(CenterLibrary.DBConnection.eStore, sql.ToString)
        If EstoreCBOM.Rows.Count > 0 Then
            Dim Delete_SQL As String = "DELETE from CBOM_CATALOG_CATEGORY WHERE (ORG='US' OR ORG ='AU' OR ORG ='JP' OR ORG ='TW' OR ORG ='KR' ) and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming'" '2015/4/10 Add TW
            Dim ReturnInt As Integer = dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, Delete_SQL)
            If ReturnInt <> -1 Then
                ''Dim intforemail As Integer = 0
                'For i As Integer = 0 To EstoreCBOM.Rows.Count - 1

                '    Dim DEFAULT_FLAG As Integer = 0
                '    If Boolean.TryParse(EstoreCBOM.Rows(i).Item("DEFAULT_FLAG"), True) AndAlso EstoreCBOM.Rows(i).Item("DEFAULT_FLAG") = True Then
                '        EstoreCBOM.Rows(i).Item("DEFAULT_FLAG") = "1"
                '    Else
                '        EstoreCBOM.Rows(i).Item("DEFAULT_FLAG") = "0"
                '    End If
                '    Dim SHOW_HIDE As Integer = 0
                '    If Boolean.TryParse(EstoreCBOM.Rows(i).Item("SHOW_HIDE"), True) AndAlso EstoreCBOM.Rows(i).Item("SHOW_HIDE") = True Then
                '        EstoreCBOM.Rows(i).Item("SHOW_HIDE") = "1"
                '    Else
                '        EstoreCBOM.Rows(i).Item("SHOW_HIDE") = "0"
                '    End If
                '    'Console.WriteLine(String.Format("{0};", Insertsql.ToString))
                'Next
                'EstoreCBOM.AcceptChanges()
                Try
                    Dim conn As SqlClient.SqlConnection = New SqlClient.SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
                    Dim bk As New SqlClient.SqlBulkCopy(conn)
                    bk.DestinationTableName = "CBOM_CATALOG_CATEGORY"
                    bk.BulkCopyTimeout = 300    'ICC 2015/4/28 Change bulk copy time out to 5 minutes
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CATEGORY_ID", "CATEGORY_ID"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CATEGORY_NAME", "CATEGORY_NAME"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CATEGORY_TYPE", "CATEGORY_TYPE"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("PARENT_CATEGORY_ID", "PARENT_CATEGORY_ID"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CATALOG_ID", "CATALOG_ID"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CATEGORY_DESC", "CATEGORY_DESC"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("DISPLAY_NAME", "DISPLAY_NAME"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("IMAGE_ID", "IMAGE_ID"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("EXTENDED_DESC", "EXTENDED_DESC"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CREATED", "CREATED"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CREATED_BY", "CREATED_BY"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("LAST_UPDATED", "LAST_UPDATED"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("LAST_UPDATED_BY", "LAST_UPDATED_BY"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("SEQ_NO", "SEQ_NO"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("PUBLISH_STATUS", "PUBLISH_STATUS"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("DEFAULT_FLAG", "DEFAULT_FLAG"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("CONFIGURATION_RULE", "CONFIGURATION_RULE"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("NOT_EXPAND_CATEGORY", "NOT_EXPAND_CATEGORY"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("SHOW_HIDE", "SHOW_HIDE"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("EZ_FLAG", "EZ_FLAG"))
                    'bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("UID", "UID"))
                    bk.ColumnMappings.Add(New SqlClient.SqlBulkCopyColumnMapping("ORG", "ORG"))
                    If conn.State <> ConnectionState.Open Then conn.Open()
                    bk.WriteToServer(EstoreCBOM)
                    'dbUtil.dbExecuteNoQuery("my", Insertsql.ToString)
                Catch ex As Exception
                    'intforemail += 1
                    'If intforemail < 3 Then
                    dbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync CBOM From Estore Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.ToString(), False)
                    'Else
                    Exit Sub
                    'End If

                End Try

                dbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync CBOM From Estore OK " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
                ' Console.Read()
            End If

            '            Dim STR As String = "update CBOM_CATALOG_CATEGORY set  " + _
            '" CONFIGURATION_RULE='REQUIRED'  " + _
            '" WHERE (ORG='US' OR ORG ='AU' OR ORG ='JP') and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='category' " + _
            '" AND (select count(category_id) FROM CBOM_CATALOG_CATEGORY B where B.PARENT_CATEGORY_ID=CBOM_CATALOG_CATEGORY.CATEGORY_ID AND B.category_id<>'None' and B.DEFAULT_FLAG=1 AND B.ORG='US')<>0"
            'Frank 2015/02/11:If there is a component named Nono in a category, then this category should not be a required category
            Dim STR As String = "update CBOM_CATALOG_CATEGORY set  " + _
" CONFIGURATION_RULE='REQUIRED'  " + _
" WHERE (ORG='US' OR ORG ='AU' OR ORG ='JP' OR ORG='TW') and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='category' " + _
" AND ( (select count(category_id) FROM CBOM_CATALOG_CATEGORY B where B.PARENT_CATEGORY_ID=CBOM_CATALOG_CATEGORY.CATEGORY_ID AND B.category_id<>'None' and B.DEFAULT_FLAG=1 AND B.ORG='US')<>0" + _
"   AND (select count(category_id) FROM CBOM_CATALOG_CATEGORY C where C.PARENT_CATEGORY_ID=CBOM_CATALOG_CATEGORY.CATEGORY_ID AND C.category_id='None' AND C.ORG='US')=0 )"

            'dbUtil.dbExecuteNoQuery("MY", STR)

            'Ming 20150915 No need's Parent category is not required  
            Dim RTR As New Text.StringBuilder()
            RTR.Append(" ;update   CBOM_CATALOG_CATEGORY set CONFIGURATION_RULE ='' WHERE ")
            RTR.Append(" (ORG='US' ) and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='category' ")
            RTR.Append(" and CATEGORY_ID in ( ")
            RTR.Append(" select PARENT_CATEGORY_ID from  CBOM_CATALOG_CATEGORY WHERE (ORG='US' )  ")
            RTR.Append(" and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='Component' and  CATEGORY_ID='no need') ")
            dbUtil.dbExecuteNoQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, STR + RTR.ToString())


            'Dim dt As DataTable = dbUtil.dbGetDataTable("MY", "select CATEGORY_ID,org from CBOM_CATALOG_CATEGORY WHERE (ORG='US' OR ORG ='AU' OR ORG ='JP') and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='category'")
            'If dt.Rows.Count > 0 Then
            '    Dim objnum As Object = Nothing
            '    For Each dr As DataRow In dt.Rows
            '        Dim dt2 As DataTable = dbUtil.dbGetDataTable("my", "select  CATEGORY_ID  from  CBOM_CATALOG_CATEGORY where CATEGORY_ID='None' and PARENT_CATEGORY_ID='" + dr.Item("CATEGORY_ID") + "' and org='" + dr.Item("org") + "'")
            '        If dt2.Rows.Count > 0 Then
            '        Else
            '            sql = "update CBOM_CATALOG_CATEGORY set  CONFIGURATION_RULE='REQUIRED'  WHERE ORG ='" + dr.Item("org") + "' and EZ_FLAG='2' AND IMAGE_ID ='created_by_ming' and CATEGORY_TYPE='category' and CATEGORY_ID='" + dr.Item("CATEGORY_ID") + "'"
            '            dbUtil.dbExecuteNoQuery("my", sql)
            '            Console.WriteLine(String.Format("{0};", sql) + vbNewLine)
            '        End If
            '    Next
            'End If
        End If
    End Sub



End Module
