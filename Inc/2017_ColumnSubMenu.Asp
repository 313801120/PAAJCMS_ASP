<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-13
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'无级菜单操作类 20171002
'获得栏目名称
function handleGetColumnName(tableName,sID)
	handleGetColumnName=""
	if tableName="" then
		tableName="webcolumn"
	end if
    if sID <> "" then
	'【@是jsp显示@】try{
        rsx.open "Select * from " & db_PREFIX & tableName & " where id=" & sID, conn, 1, 1 
        if not rsx.EOF then
            handleGetColumnName = rsx("columnName") 
        end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    end if 
end function '获得栏目名称
function handleGetColumnID(tableName,columnName)
	handleGetColumnID=-1
	if tableName="" then
		tableName="webcolumn"
	end if
    if columnName <> "" then
	'【@是jsp显示@】try{
        rsx.open "Select * from " & db_PREFIX & tableName & " where columnName='"&columnName&"'", conn, 1, 1 
        if not rsx.EOF then
            handleGetColumnID = rsx("id") 
        end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    end if 
end function 
%>