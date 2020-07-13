Option Explicit On
Option Strict On

Public NotInheritable Class DBConnString

    'VPN  7.191.194.14    
    ''Public Shared strConn2 As String = "Data Source=7.191.194.14\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=sys0500"

    '  DataBaseTestz 
    'Public Shared strConn3 As String = "Data Source=118.175.36.66,1433\SQLexpress;Initial Catalog=db2012n; User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=192.168.1.3,1433\SQLexpress;Initial Catalog=db2012n; User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=192.168.1.3\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    'Public Shared strConn4 As String = "Data Source=192.168.1.3,1433\SQLexpress;Initial Catalog=db2012n; User ID=sa;Password=$y$05000"

    Public Shared strConn1 As String = "Data Source=192.168.1.3\SQLEXPRESS;Initial Catalog=newZone;User ID=sa;Password=$y$05000"
    Public Shared strConn2 As String = "Data Source=EDP01\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    Public Shared strConn3 As String = "Data Source=58.97.96.57\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"
    Public Shared strConn4 As String = "Data Source=192.168.1.3\SQLEXPRESS;Initial Catalog=newZone;User ID=sa;Password=$y$05000"
    Public Shared strConn5 As String = "Data Source=58.97.96.57\SQLEXPRESS;Initial Catalog=newZone;User ID=sa;Password=$y$05000"

    'Public Shared strConn2 As String = "Data Source=COM-PC\SQLEXPRESS;Initial Catalog=db2012; User ID=sa;Password=$y$05000"
    'Public Shared strConn2 As String = "Data Source=ict-kritpon\SQLEXPRESS;Initial Catalog=db2012;User ID=sa;Password=$y$05000"

    'strConNet = "server=COM-PC\SQLEXPRESS;Persist Security Info=True;"
    'strConNet = strConNet & "database=db2012;User ID=sa;password=$y$05000;"

    '===================================================================================
    '===================================================================================

    'Public Shared strConn2 As String = "Data Source=192.168.1.13\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=sys0500"
    'Public Shared strConn2 As String = "Data Source=192.168.1.3\SQLEXPRESS;Initial Catalog=DB2012;User ID=sa;Password=$y$05000"

    'strConNet = "server=192.168.1.13\SQLEXPRESS;database=DB2006;Persist Security Info=True;"
    'strConNet = strConNet & "User ID=sa;password=sys0500;"

    Public Shared UserName As String = ""


End Class
