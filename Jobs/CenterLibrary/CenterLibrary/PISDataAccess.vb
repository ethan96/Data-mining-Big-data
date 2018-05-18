Imports System.Text
Imports System.Data.SqlClient

Public Class PISDataAccess
    Public Function GetCATEGORY_HIERARCHY() As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Select [model_no],[parent_category_id1],[category_name1],[category_type1],[parent_category_id2] ")
        _sql.AppendLine(" ,[category_name2],[category_type2],[parent_category_id3],[category_name3],[category_type3] ")
        _sql.AppendLine(" ,[parent_category_id4],[category_name4],[category_type4],[parent_category_id5],[category_name5] ")
        _sql.AppendLine(" ,[category_type5],[parent_category_id6],[category_name6],[category_type6],[catalog_id] ")
        _sql.AppendLine(" FROM [CATEGORY_HIERARCHY] ")
        _sql.AppendLine(" where catalog_id <>'736d1add-de2e-47ef-a7c4-00625ad8d53f' ")
        _sql.AppendLine(" order by seq ")
        Dim _db_type As String = DBConnection.PIS
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function


    Public Function Delete_MODELCATEGORY_INTERESTEDPRODUCT_MAPPING() As Integer
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Delete From MODELCATEGORY_INTERESTEDPRODUCT_MAPPING ")
        Dim _db_type As String = DBConnection.PIS
        Return dbUtil.dbExecuteNoQuery(_db_type, _sql.ToString)
    End Function


    Public Function Delete_MyAdvantechGlobal_PIS_InterestedProduct_ProductGroup() As Integer
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Delete From PIS_InterestedProduct_ProductGroup ")
        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        Return dbUtil.dbExecuteNoQuery(_db_type, _sql.ToString)
    End Function

    Public Function Get_InterestedProduct_ProductGroup() As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" SELECT INTERESTED_PRODUCT_CATEGOEY_ID,INTERESTED_PRODUCT_DISPLAY_NAME,PRODUCT_GROUP_CATEGOEY_ID,PRODUCT_GROUP_DISPLAY_NAME ")
        _sql.AppendLine(" FROM MODELCATEGORY_INTERESTEDPRODUCT_MAPPING ")
        _sql.AppendLine(" WHERE INTERESTED_PRODUCT_CATEGOEY_ID<>'' and PRODUCT_GROUP_CATEGOEY_ID<>'' ")
        _sql.AppendLine(" GROUP BY INTERESTED_PRODUCT_CATEGOEY_ID,INTERESTED_PRODUCT_DISPLAY_NAME,PRODUCT_GROUP_CATEGOEY_ID,PRODUCT_GROUP_DISPLAY_NAME ")
        Dim _db_type As String = DBConnection.PIS
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function




    Public Function GetModelDisplayName(ByVal _ModelName As String, Optional ByVal _LangID As String = "ENU") As String
        Dim _sql As New StringBuilder


        Select Case _LangID
            Case "ENU"
                _sql.AppendLine(" Select DISPLAY_NAME From Model Where MODEL_NAME='" & _ModelName & "'")

            Case Else

        End Select

        Dim _db_type As String = DBConnection.PIS
        Dim _dt As DataTable = dbUtil.dbGetDataTable(_db_type, _sql.ToString)

        If _dt IsNot Nothing AndAlso _dt.Rows.Count > 0 Then Return _dt.Rows(0).Item("DISPLAY_NAME").ToString

        Return String.Empty
    End Function

    Public Function GetLiteratureByLiterId(ByVal _LiterId As String) As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" SELECT LITERATURE_ID,SIEBEL_FILENAME,LIT_NAME,LIT_DESC ")
        _sql.AppendLine(" ,LIT_TYPE,[FILE_NAME],FILE_EXT,FILE_SIZE,FILE_LOCATION ")
        _sql.AppendLine(" ,PRIMARY_ORG_ID,PRIMARY_BU,PRIMARY_LEVEL,PRIMARY_SDU ")
        _sql.AppendLine(" ,CREATED,CREATED_BY,LAST_UPDATED,LAST_UPDATED_BY ")
        _sql.AppendLine(" ,[START_DATE],END_DATE,INT_FLG,LANG ")
        _sql.AppendLine(" FROM LITERATURE ")
        _sql.AppendLine(" Where LITERATURE_ID='" & _LiterId & "' ")
        Dim _db_type As String = DBConnection.PIS
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function

    Public Function GetPRODUCT_LINEByModelName(ByVal model_name As String) As String

        Dim _conStr As String = DBConnection.PISBackend
        'GetConnectionString( ConnectionManager.DatabaseConnection.PISBackend_Readonly

        Dim query As String = " SELECT top 1 b.PRODUCT_LINE FROM model_product a left join PIS.dbo.PRODUCT_LOGISTICS_NEW b on a.part_no=b.PART_NO "
        query &= " Where a.model_name=@modelname and a.relation='product' and a.[status]='active' "
        query &= " Order by Last_update_date desc "
        Using con As New SqlConnection(_conStr)
            Using cmd As New SqlCommand(query, con)
                con.Open()
                cmd.Parameters.AddWithValue("@modelname", model_name)
                Dim _dt As DataTable = New DataTable
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.HasRows() Then
                        _dt.Load(dr)
                        If _dt IsNot Nothing AndAlso _dt.Rows.Count > 0 Then
                            Return _dt.Rows(0).Item("PRODUCT_LINE").ToString
                        End If
                    End If
                End Using
            End Using
        End Using
        Return ""
    End Function


    Public Function GetTodayModifyInterestedProductList(Optional BeforeDayCount As Integer = 1) As DataTable

        Dim _conStr As String = DBConnection.PISBackend
        Dim _BeforeDateStr As String = Format(DateAdd(DateInterval.Day, -BeforeDayCount, Now), "yyyy/MM/dd")
        Dim _sql As New StringBuilder
        _sql.AppendLine(" SELECT [userid],[action],[model_partno],[functionname] ")
        _sql.AppendLine(" ,[inserttime],[OldData],[NewData],[notifydate] ")
        _sql.AppendLine(" FROM PISlog a ")
        _sql.AppendLine(" Where CONVERT(VARCHAR(10), a.inserttime, 111)>='" & _BeforeDateStr & "' ")
        _sql.AppendLine(" And a.[action] like '%interested product%' ")
        _sql.AppendLine(" And a.Lang_id='ENU' ")
        Using con As New SqlConnection(_conStr)
            Using cmd As New SqlCommand(_sql.ToString, con)
                con.Open()
                Dim _dt As DataTable = New DataTable
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.HasRows() Then
                        _dt.Load(dr)
                        Return _dt
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function


    Public Function GetTodayModifyModelList(Optional BeforeDayCount As Integer = 1) As DataTable

        Dim _conStr As String = DBConnection.PISBackend
        Dim _BeforeDateStr As String = Format(DateAdd(DateInterval.Day, -BeforeDayCount, Now), "yyyy/MM/dd")
        Dim _sql As New StringBuilder
        _sql.AppendLine(" SELECT distinct a.model_partno as Model_Name ")
        _sql.AppendLine(" FROM PISlog a inner join model b on a.model_partno=b.MODEL_NAME ")
        _sql.AppendLine(" Where CONVERT(VARCHAR(10), a.inserttime, 111)>='" & _BeforeDateStr & "' ")
        _sql.AppendLine(" And a.model_partno<>'' ")
        Using con As New SqlConnection(_conStr)
            Using cmd As New SqlCommand(_sql.ToString, con)
                con.Open()
                Dim _dt As DataTable = New DataTable
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.HasRows() Then
                        _dt.Load(dr)
                        Return _dt
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function



    Public Function GetAllModelList() As DataTable

        Dim _conStr As String = DBConnection.PISBackend

        Dim _sql As New StringBuilder
        _sql.AppendLine(" Select m.Model_name as Model_Name,mpu.Active_FLG as Model_Active_FLG From model m ")
        _sql.AppendLine(" left join Category_Model cm on cm.Model_name=m.Model_name ")
        _sql.AppendLine(" left join CATALOG_CATEGORY cc on cc.Category_id=cm.Category_id ")
        _sql.AppendLine(" left join Model_Publish mpu on mpu.model_name=m.model_name and mpu.Site_ID ='ACL' ")
        _sql.AppendLine(" Where cc.CATALOGID IN ('1-2MLAX2','1-2JKBQD') ")
        _sql.AppendLine(" Order by m.LAST_UPDATED desc ")
        Using con As New SqlConnection(_conStr)
            Using cmd As New SqlCommand(_sql.ToString, con)
                con.Open()
                Dim _dt As DataTable = New DataTable
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.HasRows() Then
                        _dt.Load(dr)
                        Return _dt
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function

#Region "MyA"


    Public Function GetUnprocessPackageIDList() As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Select Distinct PackageID ")
        _sql.AppendLine(" FROM PACKAGE_PRODUCT_DATASHEET ")
        _sql.AppendLine(" Where IsPackage=0 ")
        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function

    Public Function GetUnprocessPackageIDByPackageID(ByVal _PackageID As String) As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Select Distinct PackageID")
        _sql.AppendLine(" FROM PACKAGE_PRODUCT_DATASHEET ")
        _sql.AppendLine(" Where IsPackage=0 and PackageID='" & _PackageID & "' ")
        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function


    Public Function GetPackageFileListByPackageId(ByVal _PackageId As String) As DataTable
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Select PackageID,PartNo,Model_Name,PisLiterID,Email,UploadTime,IsPackage ")
        _sql.AppendLine(" FROM PACKAGE_PRODUCT_DATASHEET ")
        _sql.AppendLine(" Where PackageID='" & _PackageId & "' ")
        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        Return dbUtil.dbGetDataTable(_db_type, _sql.ToString)
    End Function

    Public Function UpdatePackageStatusByPackageId(ByVal _PackageId As String, ByVal _IsProcessed As Boolean) As Integer
        Dim _sql As New StringBuilder
        _sql.AppendLine(" Update PACKAGE_PRODUCT_DATASHEET ")
        If _IsProcessed Then
            _sql.AppendLine(" Set IsPackage=1 ")
        Else
            _sql.AppendLine(" Set IsPackage=0 ")
        End If
        _sql.AppendLine(" Where PackageID='" & _PackageId & "' ")

        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        Return dbUtil.dbExecuteNoQuery(_db_type, _sql.ToString)
    End Function

    Public Sub SavePackageDatasheetFile(ByVal _PackageID As String, ByVal _File_bytes As Byte())
        Dim _MyCon As New SqlConnection(DBConnection.MyAdvantechGlobal)
        If _MyCon.State <> ConnectionState.Open Then _MyCon.Open()



        Dim cmd As New SqlClient.SqlCommand()
        cmd.Connection = _MyCon

        Dim _DeleteSQL As String = String.Format("Delete From PACKAGE_PRODUCT_DATASHEET_CACHE Where PackageID='{0}'", _PackageID)

        Dim _db_type As String = DBConnection.MyAdvantechGlobal
        dbUtil.dbExecuteNoQuery(_db_type, _DeleteSQL)


        'cmd.CommandText = "insert into THUMBNAIL_CACHE (SOURCE_TYPE, SOURCE_ID, THUMBNAIL_BYTES) values(@SOURCE_TYPE,@SOURCE_ID,@THUMBNAIL_BYTES)"
        cmd.CommandText = _DeleteSQL & ";insert into PACKAGE_PRODUCT_DATASHEET_CACHE (PackageID, File_Bytes, Cached_Date) values(@PackageID,@File_Bytes,@Cached_Date)"
        cmd.Parameters.AddWithValue("PackageID", _PackageID) : cmd.Parameters.AddWithValue("File_Bytes", _File_bytes) : cmd.Parameters.AddWithValue("Cached_Date", Now)
        cmd.ExecuteNonQuery()

    End Sub

#End Region
End Class
