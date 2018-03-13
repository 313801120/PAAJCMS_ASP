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
Sub bootstrapPageControl(iCount,pagecount,page,table_style,font_style)
'生成上一页下一页链接
    Dim action,query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next

    Response.Write(" <ul class=""pagination"">" & vbCrLf )    
        
    if page<=1 then
		Response.Write("<li class=""disabled""><a href=""javascript:;"">&laquo;</a></li>" & vbcrlf)
    else        
		Response.Write("<li><a href="""& action & "?" & temp & "Page=1"">&laquo;</a></li>" & vbcrlf) 
    end if
	
	
	    dim n ,i,nDispalyOK,nDisplay,sTemp
		nDisplay=6
		nDispalyOK=0
    if page=0 then
		page=1
	end if
    n =(page - 3) 
    'call echo("n=" & n, "nPage=" & nPage)
	
    '翻页循环
    for i = n to pagecount
        if i >= 1 then
            nDispalyOK = nDispalyOK + 1 
            'call echo(i,nPage)
            if i = page then 
				response.Write(" <li class=""disabled""><a href=""javascript:;"">"& i &"</a></li>" & vbcrlf)
            else
                sTemp = cstr(i) 
                if i <= 1 then
                    sTemp = "" 
                end if 
				response.Write(" <li><a href="""& action & "?" & temp & "Page="& i &""">"& i &"</a></li>" & vbcrlf)
            end if 
            if nDispalyOK > nDisplay then
                exit for 
            end if 
        end if 
    next

    if page>=pagecount then
    	Response.Write("<li class=""disabled""><a href=""javascript:;"">&raquo;</a></li>" & vbcrlf)          
    else 
		Response.Write("<li><a href="""& action & "?" & temp & "Page="& pagecount &""">&raquo;</a></li>" & vbcrlf)          
    end if
 
End Sub
  %>