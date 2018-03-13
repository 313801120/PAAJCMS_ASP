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
'ASP和PHP通用函数

'获得导航id 根据大类与小类 20170404  call getWebColumnDidSid(did,sid,true)
function getWebColumnDidSid(byval did,byval sid, byval isAdd)
	dim i,parentid,id
	rs1.open"select * from " & db_PREFIX & "WebColumn where parentid=-1 and columnName='"& did &"'",conn,1,1
	if not rs1.eof then
		id=rs1(0)
		parentId=id
	end if:rs1.close
	if id="" then
		exit function
	end if 
	rs1.open"select * from " & db_PREFIX & "WebColumn where parentid="& id &" and columnName='"& sid &"'",conn,1,3
	if not rs1.eof then
		id=rs1(0)
	else
		id=""
	end if:rs1.close
	'添加
	if id="" and isAdd=true then	
		rs1.open"select * from " & db_PREFIX & "WebColumn",conn,1,3
		rs1.addnew
		rs1("parentid")=parentId
		rs1("columnName")=sid
		rs1.update:rs1.close
	end if
	
	getWebColumnDidSid=id
end function
'处理主栏目信息
function handleColumnRoot(byval id)
	dim i,parentid,rs1
	Set rs1 = CreateObject("Adodb.RecordSet")
	for i =1 to 9
		rs1.open"select * from " & db_PREFIX & "WebColumn where id="& id &"",conn,1,1
		if not rs1.eof then
			parentid=cstr(rs1("parentid"))
			if parentid<>"-1" then
				id=rs1("parentid")
			else
				glb_columnRootName=rs1("columnname")
				glb_columnRootEnName=rs1("columnenname")
				glb_columnRootId=rs1("id")
				rs1.close
				exit for		
			end if
		else
			exit for
		end if:rs1.close
	next
end function

'获得栏目主id
function getColumnRootId(byval id)
	dim parentid,i
	for i =1 to 9
		rs2.open"select * from " & db_PREFIX & "WebColumn where id="& id &"",conn,1,1
		if not rs2.eof then
			parentid=cstr(rs2("parentid"))
			if parentid<>"-1" then
				id=rs2("parentid")
			else
				exit for		
			end if
		else
			exit for
		end if:rs2.close
	next
	getColumnRootId=id
end function
'获得栏目上一级id
function getParentColumnId(parentID) 
    getParentColumnId = "-1" 
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "webcolumn where id=" & parentID & "", conn, 1, 1 
    if not rsx.EOF then
        getParentColumnId = rsx("parentID") 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function
'获得栏目上一级名称
function getParentColumnName(parentID)  
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "webcolumn where id=" & parentID & "", conn, 1, 1 
    if not rsx.EOF then
        getParentColumnName = rsx("columnName") 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 
'获得子栏目id列表 
function getColumnIdList(byval id, sType)
	dim parentid,c,splstr,s
	if sType="addthis" then
		c=id
	end if
	rs3.open"select * from " & db_PREFIX & "WebColumn where parentid="& id &"",conn,1,1
	while not rs3.eof
		id=rs3("id")
		if instr("," & c & ",", ","& id &",")=false then
			if c<>"" then
				c=c & ","
			end if
			c=c & id
		end if
	rs3.movenext:wend:rs3.close
	splstr=split(c,",")
	for each id in splstr
		if id<>"" then
			s=getColumnIdList(id,"")
			if s<>"" and instr(","& c &",", ","& s &",")=false then
				c=c & "," & s
			end if
		end if
	next 
	getColumnIdList=c
end function

'文章相关标签 组  待改进
function aritcleRelatedTags(relatedTags)
    dim c, splStr, s, url 
    splStr = split(relatedTags, ",") 
    for each s in splStr
        if s <> "" then
            if c <> "" then
                c = c & "," 
            end if 
            url = getColumnUrl(s, "name") 
            c = c & "<a href=""" & url & """ rel=""category tag"" class=""ablue"">" & s & "</a>" & vbCrLf 
        end if 
    next 

    c = "<footer class=""articlefooter"">" & vbCrLf & "标签： " & c & "</footer>" & vbCrLf 
    aritcleRelatedTags = c 
end function 


'获得随机文章id列表
function getRandArticleId(addSql, nTopNumb)
    dim splStr, s, c, nIndex 
	'【@是jsp显示@】try{	
    rs.open "select * from " & db_PREFIX & "articledetail " & addSql, conn, 1, 1 
    while not rs.EOF
        if c <> "" then c = c & "," 
        c = c & rs("id") 
    rs.moveNext : wend : rs.close 
	'【@是jsp显示@】}catch(Exception e){}
    getRandArticleId = randomShow(c, ",", 4) 
    splStr = split(c, ",") : c = "" : nIndex = 0 
    for each s in splStr
        if c <> "" then c = c & "," 
        c = c & s 
        nIndex = nIndex + 1 
        if nIndex >= nTopNumb then exit for 
    next 
    getRandArticleId = c 
end function 
'获得网站栏目排序SQL
function getWebColumnSortSql(id)
    dim sql 
	'【@是jsp显示@】try{
    tempRs2.open "select * from " & db_PREFIX & "webcolumn where id=" & id, conn, 1, 1 
    if not tempRs2.eof then
        sql = cstr(tempRs2("sortsql")) 
    end if : tempRs2.close 
	'【@是jsp显示@】}catch(Exception e){}
    getWebColumnSortSql = sql 
end function 

'上一篇文章 这里面的sortrank(排序)也可以改为id,在引用的时候就要用id
function upArticle(parentid, lableName, lableValue, ascOrDesc,url)
    upArticle = handleUpDownArticle("上一篇：", "uppage", parentid, lableName, lableValue, ascOrDesc,url) 
end function 
'下一篇文章
function downArticle(parentid, lableName, lableValue, ascOrDesc,url)
    downArticle = handleUpDownArticle("下一篇：", "downpage", parentid, lableName, lableValue, ascOrDesc,url) 
end function 
'处理上下页  url就给上页值用的
function handleUpDownArticle(lableTitle, sType, parentid, lableName, lableValue, ascOrDesc,url)
    dim c, target, targetStr 

    dim sql 
    if lableName = "adddatetime" then
        lableValue = "#" & lableValue & "#" 
    end if 
    '位置互换
    if ascOrDesc = "desc" then
        if sType = "uppage" then
            sType = "downpage" 
        else
            sType = "uppage"  
        end if 
    end if 
    if sType = "uppage" then 
        sql = "select * from " & db_PREFIX & "articledetail where parentid=" & parentid & " and " & lableName & "<" & lableValue & " order by " & lableName & " desc" 
    else
        sql = "select * from " & db_PREFIX & "articledetail where parentid=" & parentid & " and " & lableName & ">" & lableValue & " order by " & lableName & " asc" 
    end if 

    'call echo("sql",sql) 
	'【@是jsp显示@】try{
    rsx.open sql, conn, 1, 1 
    if not rsx.EOF then
        target = rsx("target") 
        if target <> "" then
            targetStr = " target=""" & target & """" 
        end if 
        if isMakeHtml = true then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
        else
            if rsx("customAUrl") = "" then
                url = handleWebUrl("?act=detail&id=" & rsx("id")) 
            else
                url = handleWebUrl(rsx("customAUrl")) 
            end if 
        end if 
        c = "<a href=""" & url & """" & targetStr & ">" & lableTitle & rsx("title") & "</a>" 
    else
        c = lableTitle & "没有" 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    handleUpDownArticle = c 
end function 
'获得RS网址 配置上一页 下一页
function getRsUrl(byVal fileName, customAUrl, defaultFileName)
    dim url 
    '用默认文件名称
    if fileName = "" then
        fileName = defaultFileName 
    end if 
    '网址
    if fileName <> "" then
        fileName = lCase(fileName)                                                      '让文件名称小写20160315
        url = fileName 
        if inStr(lCase(url), ".html") = 0 and right(url, 1) <> "/" then
            url = url & ".html" 
        end if 
    end if 
    if trim(customAUrl) <> "" then
        url = trim(customAUrl) 
    end if 
	'追加这个是为了在生成静态文件时，可以获得首页的文件名称，好让index.html#about  出现 20160728
	if url="/" then
		url="/index.html"
	end if
    if inStr(cfg_flags, "|addwebsite|") > 0 then 
        'url = replaceGlobleVariable(url)   '替换全局变量
        if inStr(url, "$cfg_websiteurl$") = false and inStr(url, "{$GetColumnUrl ") = false and inStr(url, "{$GetArticleUrl ") = false and inStr(url, "{$GetOnePageUrl ") = false then
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        end if 
    end if 
    getRsUrl = url 
end function 
'获得处理后RS网址
function getHandleRsUrl(fileName, customAUrl, defaultFileName)
    dim url 
    url = getRsUrl(fileName, customAUrl, defaultFileName) 
    '因为URL如果为自定义的则需要处理下全局变量，这样程序运行又会变慢，不就可以使用生成HTML方法解决这个问题，20160308
    url = replaceGlobleVariable(url) 
    getHandleRsUrl = url 
end function 

'获得单页url 20160114
function getOnePageUrl(title)
    dim url 
	'【@是jsp显示@】try{
    rsx.open "select * from " & db_PREFIX & "onepage where title='" & title & "'", conn, 1, 1 
    if not rsx.EOF then
        if isMakeHtml = true then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/page" & rsx("id")) 
        else
            url = handleWebUrl("?act=onepage&id=" & rsx("id")) 
            if rsx("customAUrl") <> "" then
                url = rsx("customAUrl") 
            end if 
        end if 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}

    getOnePageUrl = url 
end function 
'获得文章URL
function getArticleUrl(title)
    dim url 
	'【@是jsp显示@】try{
    rsx.open "select * from " & db_PREFIX & "articledetail where title='" & title & "'", conn, 1, 1 
    if not rsx.EOF then
        if isMakeHtml = true then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/" & rsx("id")) 
        else
            url = handleWebUrl("?act=article&id=" & rsx("id")) 
            if rsx("customAUrl") <> "" then
                url = rsx("customAUrl") 
            end if 
        end if 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}

    getArticleUrl = url 
end function 
'获得栏目URL 20160114 getColumnUrl("首页","type")   
function getColumnUrl(columnNameOrId, sType)
    dim url, addSql 

    columnNameOrId = replaceGlobleVariable(columnNameOrId)                          '处理动作 <a href="{$GetColumnUrl columnname='[$glb_columnName$]' $}" >更多图片</a>

    if sType = "name" then
        addSql = " where columnname='" & replace(columnNameOrId, "'", "''") & "'"       '对'号处理，要不然sql查询出错20160716
    elseif sType = "type" then
        addSql = " where columntype='"& columnNameOrId &"'"       '对'号处理，要不然sql查询出错20160716
    else
        addSql = " where id=" & columnNameOrId & "" 
    end if 
	'【@是jsp显示@】try{
    rsx.open "select * from " & db_PREFIX & "webcolumn" & addSql, conn, 1, 1 
    if not rsx.EOF then
        if isMakeHtml = true then
            url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
			'call echo(rsx("columnName"),url)
        else
            'url = handleWebUrl("?act=nav&columnName=" & rsx("columnname")) 		'这个在jquery.mobile里会有问题
            url = handleWebUrl("?act=nav&id=" & rsx("id")) 
            if rsx("customAUrl") <> "" then
                url = rsx("customAUrl") 
            end if 
        end if 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}

    getColumnUrl = url 
end function 

'获得文章标题对应的id
function getArticleId(title)
    title = replace(title, "'", "")                                                 '注意，这个不能留
    getArticleId = "-1" 
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "ArticleDetail where title='" & title & "'", conn, 1, 1 
    if not rsx.EOF then
        getArticleId = rsx("id")
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 

'获得栏目id
function getColumnId(columnName)
    'columnName = Replace(columnName, "'", "")           '注意，这个不能留  因为sql里已经处理了 20160716 home 程序写得越来越深，逻辑越多
    getColumnId = "-1" 
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "webcolumn where columnName='" & columnName & "'", conn, 1, 1 
    if not rsx.EOF then
        getColumnId = rsx("id") 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function
'获得栏目父窗口id
function getWebColumnParentId(s,sType)
	dim sql,addsql,rsx
	Set rsx = CreateObject("Adodb.RecordSet")
	if sType="id" then
		addsql=" where id=" & s & ""
		getWebColumnParentId="-1"
	else
		addsql=" where columnName='" & s & "'"
	end if
	sql="Select * from " & db_PREFIX & "webcolumn"&addsql
	'call eerr("sql",sql)
    rsx.open sql, conn, 1, 1 
    if not rsx.EOF then
        getWebColumnParentId = rsx("parentId") 
    end if : rsx.close  
end function


'获得栏目名称
function getColumnName(sID)
	getColumnName=""
    if sID <> "" then
	'【@是jsp显示@】try{ 
        rsx.open "Select * from " & db_PREFIX & "webcolumn where id=" & sID, conn, 1, 1 
        if not rsx.EOF then
            getColumnName = rsx("columnName") 
        end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    end if 
end function 

'获得栏目类型
function getColumnType(columnID)
	getColumnType=""
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "webcolumn where id=" & columnID, conn, 1, 1 
    if not rsx.EOF then
        getColumnType = rsx("columnType") 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 
 
'获得栏目内容
function getColumnBodyContent(columnID)
	getColumnBodyContent=""
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "webcolumn where id=" & columnID, conn, 1, 1 
    if not rsx.EOF then
        getColumnBodyContent = rsx("bodyContent") 
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 


'获得后台菜单名称 
function getListMenuId(title)
    title = replace(title, "'", "")                                                 '注意，这个不能留
    getListMenuId = "-1" 
	'【@是jsp显示@】try{
    rsx.open "Select * from " & db_PREFIX & "listmenu where title='" & title & "'", conn, 1, 1 
    if not rsx.EOF then
        getListMenuId = rsx("id")
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 

'获得后台菜单ID 
function getListMenuName(sID) 
	getListMenuName=""
    if sID <> "" then
		'【@是jsp显示@】try{
        rsx.open "Select * from " & db_PREFIX & "listmenu where id=" & sID, conn, 1, 1 
        if not rsx.EOF then
            getListMenuName = rsx("title") 
        end if : rsx.close 
		'【@是jsp显示@】}catch(Exception e){}
    end if 
end function  

'网站统计2014
function webStat(folderPath)
    dim sDateTime, content, splStr 
    dim thisUrl, goToUrl, caiShu, c, fileName, co, ie, xp 
	call createDirFolder(folderPath)		'生成统计指定文件夹
    goToUrl = request.serverVariables("HTTP_REFERER") 
    thisUrl = "http://" & request.serverVariables("HTTP_HOST") & request.serverVariables("SCRIPT_NAME") 
    caiShu = request.serverVariables("QUERY_STRING") 
    if caiShu <> "" then
        thisUrl = thisUrl & "?" & caiShu 
    end if 
    goToUrl = request("GoToUrl") 
    thisUrl = request("ThisUrl") 
    co = request.queryString("co") 
    sDateTime = now() 
    content = request.serverVariables("HTTP_USER_AGENT") 
    content = replace(content, "MSIE", "Internet Explorer") 
    content = replace(content, "NT 5.0", "2000") 
    content = replace(content, "NT 5.1", "XP") 
    content = replace(content, "NT 5.2", "2003") 

    splStr = split(content & ";;;;", ";") 
    ie = splStr(1) 
    xp = trim(splStr(2)) 
    if right(xp, 1) = ")" then xp = mid(xp, 1, len(xp) - 1) 
    c = "来访" & goToUrl & vbCrLf 
    c = c & "当前：" & thisUrl & vbCrLf 
    c = c & "时间：" & sDateTime & vbCrLf 
    c = c & "IP:" & getIP() & vbCrLf 
    c = c & "IE:" & getBrType("") & vbCrLf 
    c = c & "Cookies=" & co & vbCrLf 
    c = c & "XP=" & xp & vbCrLf 
    c = c & "Screen=" & request("screen") & vbCrLf                                  '屏幕分辨率
    c = c & "用户信息=" & request.serverVariables("HTTP_USER_AGENT") & vbCrLf       '用户信息

    c = c & "-------------------------------------------------" & vbCrLf 
    'c=c & "CaiShu=" & CaiShu & vbcrlf
    fileName = folderPath & format_Time(now(), 2) & ".txt" 
    call createAddFile(fileName, c) 
    c = c & vbCrLf & fileName 
    c = replace(c, vbCrLf, "\n") 
    c = replace(c, """", "\""") 
    'Response.Write("eval(""var MyWebStat=\""" & C & "\"""")")

    dim splxx, nIP, nPV, ipList, s, ip 
    '判断是否显示回显记录
    if request("stype") = "display" then
        content = getftext(fileName) 
        splxx = split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
        nIP = 0 
        nPV = 0 
        ipList = "" 
        for each s in splxx
            if inStr(s, "当前：") > 0 then
                s = vbCrLf & s & vbCrLf 
                ip = aDSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
                nPV = nPV + 1 
                if inStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = false then
                    ipList = ipList & ip & vbCrLf 
                    nIP = nIP + 1 
                end if 
            end if 
        next 
        call rw("document.write('网长统计 | 今日IP[" & nIP & "] | 今日PV[" & nPV & "] ')") 
    end if 
    webStat = c 
end function 

'判断传值是否相等
function checkFunValue(action, funName)
    checkFunValue =IIF(left(action, len(funName)) = funName, true, false) 
end function 
'HTML标签参数自动添加(target|title|alt|id|class|style|)    辅助类
function setHtmlParam(content, paramList)
    dim splStr, startStr, endStr, c, paramValue, replaceStartStr 
    endStr = "'" 
    splStr = split(paramList, "|") 
    for each startStr in splStr 
        startStr = trim(startStr) 
        if startStr <> "" then
            '替换开始字符   因为开始字符类型可变 不同
            replaceStartStr = startStr 
            if left(replaceStartStr, 3) = "img" then
                replaceStartStr = mid(replaceStartStr, 4) 
            elseIf left(replaceStartStr, 1) = "a" then
                replaceStartStr = mid(replaceStartStr, 2) 
            elseIf inStr("|ul|li|", "|" & left(replaceStartStr, 2) & "|") > 0 then
                replaceStartStr = mid(replaceStartStr, 3) 
            end if 
            replaceStartStr = " " & replaceStartStr & "='" 

            startStr = " " & startStr & "='" 
            if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                paramValue = strCut(content, startStr, endStr, 2) 
                paramValue = handleInModule(paramValue, "end")                       '处理内部模块
                c = c & replaceStartStr & paramValue & endStr 
            end if 
        end if 
    next 
    setHtmlParam = c 
end function 
%>       

