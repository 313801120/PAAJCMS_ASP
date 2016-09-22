<%
'************************************************************
'作者：云孙World(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-09-22
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
Public MDBPath, accessPass                                     '数据库路径与密码 


'检测SqlServer链接是否成功
function checkSqlServer(byval connStr)
	dim conn
	checkSqlServer=true
	on error resume next
	Set conn = CreateObject("Adodb.Connection")  
	conn.Open connStr 
	if err<>0 then
		checkSqlServer=false
	end if
end function

'打开数据库
Public Sub openConn()
    Dim connStr 
    If Request("MDBPath") <> "" And isNul(Request("MDBPath")) = False Then
        MDBPath = handlePath(Request("MDBPath")) 
    End If 
    'call eerr("",MDBPath)
	
	'链接sqlserver数据库
	if databaseType = "sqlserver" then
        connStr = " Password = " & sqlServerPassword & "; user id =" & sqlServerUsername & "; Initial Catalog =" & sqlServerDatabaseName & "; data source =" & sqlServerHostIP & ";Provider = sqloledb;" 
		if checkSqlServer(connStr)=false then
			call eerr("SqlServer数据库链接出错", "<a href='/inc/install.asp'>点击创建数据库</a>")
		end if
        Set conn = CreateObject("Adodb.Connection")
		conn.Open connStr 
	
    '连接MMD数据库
	else
		if MDBPath="" or checkFile(MDBPath)=false then
			call eerr("Access数据库不存在", "<a href='/inc/install.asp'>点击创建数据库</a>")
		end if
        If InStr(MDBPath, ":") = 0 Then MDBPath = Server.MapPath(Replace(MDBPath, "//", "/")) 
		 
		
        Set conn = CreateObject("Adodb.Connection")
		'call eerr("MDBPath",MDBPath)
        conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '" & accessPass & "';Data Source = " & MDBPath 
    End If

    '加个判断目的是为了重复设置，调用时会出错 如GetRecordCount(Table, AddSql)   20150313
    If TypeName(rs) <> "Recordset" Then
		Set rs = CreateObject("Adodb.RecordSet")
		Set rsx = CreateObject("Adodb.RecordSet")
		Set rss = CreateObject("Adodb.RecordSet")
		Set rst = CreateObject("Adodb.RecordSet")
		Set rsd = CreateObject("Adodb.RecordSet")
		Set tempRs = CreateObject("Adodb.RecordSet")
		Set tempRs2 = CreateObject("Adodb.RecordSet")
		Set rsTemp = CreateObject("Adodb.RecordSet")
    End If
End Sub 
'关闭数据库
Function closeConn()
    On Error Resume Next 
    conn.Close 
    Set conn = Nothing 
End Function 
'获得数据库句柄  ThisConn=GetAccessConn(ThisConn,ThisMDBPath,"")
Function getAccessConn(accessConn, thisMDBPath, thisAccessPass)
    thisMDBPath = handlePath(Request("ThisMDBPath")) 
    Set accessConn = CreateObject("Adodb.Connection")
        accessConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database PassWord = '" & thisAccessPass & "';Data Source = " & thisMDBPath 
        getAccessConn = accessConn 
End Function
%> 

