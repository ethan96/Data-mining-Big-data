Public Class DBConnection
    Public Shared MyAdvantechGlobal As String = "Data Source=ACLSTNR12;Initial Catalog=MyAdvantechGlobal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Public Shared CurationPool As String = "Data Source=ACLSTNR12;Initial Catalog=CurationPool;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Public Shared SSO As String = "Data Source=172.21.1.106;Initial Catalog=membership;Persist Security Info=True;User ID=estore3test;Password=estore3test;async=true;Connect Timeout=300;pooling='true'"
    Public Shared CRMDB As String = "Data Source=ACLSQL9;Initial Catalog=aclcrmdb;Password=q32j80qkut;User ID=SYSMYADVANTECH;Connect Timeout=3000;pooling='true'"
    Public Shared PIS As String = "Data Source=ACLSTNR12;database=PIS;User Id =pisdbsa;Password =piss@;MultipleActiveResultSets=true;Application Name=SalesPortal_PIS"
    Public Shared MyLocal As String = "Data Source=ACLECAMPAIGN\MATEST;Initial Catalog=MyLocal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Public Shared eStore As String = "Data Source=172.21.1.106;Initial Catalog=eStoreProduction;Persist Security Info=True;User ID=estore3test;Password=estore3test;async=true;Connect Timeout=180;pooling='true'"
    Public Shared eStoreBB As String = "Data Source=172.21.1.106;Initial Catalog=BBeStore;Persist Security Info=True;User ID=estore3test;Password=estore3test;async=true;Connect Timeout=180;pooling='true'"
    Public Shared PISBackend As String = "Data Source=ACLSTNR12;Initial Catalog=PISBackend;Persist Security Info=True;User ID=pisdbsa;Password=piss@"
    Public Shared SAP_PRD As String = "user id=ebiz;password=ebiz;data source=(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)(HOST=172.20.1.166)(PORT=1527))(CONNECT_DATA=(SERVICE_NAME=RDP)))"
    Public Shared Elearning As String = "Data Source=ACLSTNR16;Initial Catalog=eLearning_Advantech;persist security info=True;User ID=eLear2MyAdv;password=el1@r2my@dv;async=true;Connect Timeout=300;pooling='true'"
    Public Shared Membership As String = "Data Source=172.21.1.106;Initial Catalog=membership;Persist Security Info=True;User ID=estore3test;Password=estore3test;async=true;Connect Timeout=300;pooling='true'"
End Class

Public Class AppConfig
    Public Shared SMTPServerIP As String = "172.20.0.76"
    Public Shared SAP_PRD As String = "CLIENT=168 USER=b2baeu PASSWD=ebizaeu ASHOST=172.20.1.88 SYSNR=0"
End Class
