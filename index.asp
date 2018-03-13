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
<!--#Include File = "Inc/Config.Asp"-->
<!--#Include File = "inc/admin_function.asp"--> 
<% 
'asp服务器   的发生的111
call openconn()    
'========= 
dim makeHtmlFileSetCode:makeHtmlFileSetCode="gb2312"			'生成html文件编码  

'处理动作   ReplaceValueParam为控制字符显示方式
function handleAction(content)
    dim startStr, endStr, actionList, splStr, action, s, isHnadle 
    startStr = "{$" : endStr = "$}" 
    actionList = getArrayNew(content, startStr, endStr, true, true) 
    'Call echo("ActionList ", ActionList)
    splStr = split(actionList, "$Array$") 
    for each s in splStr
        action = trim(s) 
        action = handleInModule(action, "start")                                        '处理\'替换掉
        if action <> "" then
            action = trim(mid(action, 3, len(action) - 4)) & " " 
            'call echo("s",s)
            isHnadle = true                                                                 '处理为真
            '{VB #} 这种是放在图片路径里，目的是为了在VB里不处理这个路径
            if checkFunValue(action, "# ") = true then
                action = "" 
            '测试
            elseIf checkFunValue(action, "GetLableValue ") = true then
                action = XY_getLableValue(action) 
            '标题在搜索引擎里列表
            elseIf checkFunValue(action, "TitleInSearchEngineList ") = true then
                action = XY_TitleInSearchEngineList(action) 

            '加载文件
            elseIf checkFunValue(action, "Include ") = true then
                action = XY_Include(action) 
            '栏目列表
            elseIf checkFunValue(action, "ColumnList ") = true then
                action = XY_AP_ColumnList(action) 
            '文章列表
            elseIf checkFunValue(action, "ArticleList ") = true or checkFunValue(action, "CustomInfoList ") = true then
                action = XY_AP_ArticleList(action) 
            '评论列表
            elseIf checkFunValue(action, "CommentList ") = true then
                action = XY_AP_CommentList(action) 
            '搜索统计列表
            elseIf checkFunValue(action, "SearchStatList ") = true then
                action = XY_AP_SearchStatList(action) 
            '友情链接列表
            elseIf checkFunValue(action, "Links ") = true then
                action = XY_AP_Links(action) 
            '单页列表 20170518
            elseIf checkFunValue(action, "GetOnePageList ") = true then
                action = XY_AP_GetOnePageList(action) 
            '会员列表 20170518
            elseIf checkFunValue(action, "GetMemberList ") = true then
                action = XY_AP_GetMemberList(action) 
            '自定义表列表 20170603
            elseIf checkFunValue(action, "CustomTableList ") = true then
                action = XY_AP_CustomTableList(action) 
				
				
				
				
            '获得分割列表 20170518
            elseIf checkFunValue(action, "GetSplitList ") = true then
                action = XY_AP_GetSplitList(action) 

            '显示单页内容
            elseIf checkFunValue(action, "GetOnePageBody ") = true or checkFunValue(action, "MainInfo ") = true then
                action = XY_AP_GetOnePageBody(action) 
            '显示文章内容
            elseIf checkFunValue(action, "GetArticleBody ") = true then
                action = XY_AP_GetArticleBody(action) 
            '显示栏目内容
            elseIf checkFunValue(action, "GetColumnBody ") = true then
                action = XY_AP_GetColumnBody(action) 

            '获得栏目URL
            elseIf checkFunValue(action, "GetColumnUrl ") = true then
                action = XY_GetColumnUrl(action) 
            '获得栏目ID
            elseIf checkFunValue(action, "GetColumnId ") = true then
                action = XY_GetColumnId(action) 
            '获得文章URL
            elseIf checkFunValue(action, "GetArticleUrl ") = true then
                action = XY_GetArticleUrl(action) 
            '获得单页URL
            elseIf checkFunValue(action, "GetOnePageUrl ") = true then
                action = XY_GetOnePageUrl(action) 
            '获得导航URL
            elseIf checkFunValue(action, "GetNavUrl ") = true then
                action = XY_GetNavUrl(action) 



                '------------------- 模板模块区 -----------------------
            '显示包裹块 作用不大
            elseIf checkFunValue(action, "DisplayWrap ") = true then
                action = XY_DisplayWrap(action) 
            '显示布局
            elseIf checkFunValue(action, "Layout ") = true then
                action = XY_Layout(action) 
            '显示模块
            elseIf checkFunValue(action, "Module ") = true then
                action = XY_Module(action) 
            '读模块内容
            elseIf checkFunValue(action, "ReadTemplateModule ") = true then
                action = XY_ReadTemplateModule(action) 
            '获得内容模块 20150108
            elseIf checkFunValue(action, "GetContentModule ") = true then
                action = XY_ReadTemplateModule(action) 
            '读模板样式并设置标题与内容   软件里有个栏目Style进行设置
            elseIf checkFunValue(action, "ReadColumeSetTitle ") = true then
                action = XY_ReadColumeSetTitle(action) 
            '读模板样式并设置标题与内容 顶部
            elseIf checkFunValue(action, "ReadTop ") = true then
                action = XY_ReadTop(action) 
            '读模板样式并设置标题与内容 导航
            elseIf checkFunValue(action, "ReadNav ") = true then
                action = XY_ReadNav(action) 
            '读模板样式并设置标题与内容 底部
            elseIf checkFunValue(action, "ReadFoot ") = true then
                action = XY_ReadFoot(action) 


                '------------------- 其它区 -----------------------
            '显示JS渲染ASP/PHP/VB等程序的编辑器
            elseIf checkFunValue(action, "displayEditor ") = true then
                action = displayEditor(action) 
            'Js版网站统计
            elseIf checkFunValue(action, "JsWebStat ") = true then
                action = XY_JsWebStat(action) 
            '获得表指定值
            elseIf checkFunValue(action, "NewGetFieldValue ") = true then
                action = XY_AP_NewGetFieldValue(action) 
            '获得子类表指定值
            elseIf checkFunValue(action, "SubGetFieldValue ") = true then
                action = XY_AP_SubGetFieldValue(action) 


 
                '------------------- 链接区 -----------------------
            '普通链接A
            elseIf checkFunValue(action, "HrefA ") = true then
                action = XY_HrefA(action) 
            '栏目菜单(引用后台栏目程序)
            elseIf checkFunValue(action, "ColumnMenu ") = true then
                action = XY_AP_ColumnMenu(action) 

                '------------------- 循环处理 -----------------------
            'For循环处理
            elseIf checkFunValue(action, "ForArray ") = true then
                action = XY_ForArray(action) 

                '------------------- 待分区 -----------------------
            '网站底部
            elseIf checkFunValue(action, "WebSiteBottom ") = true or checkFunValue(action, "WebBottom ") = true then
                action = XY_AP_WebSiteBottom(action) 
            '显示网站栏目 20160331
            elseIf checkFunValue(action, "DisplayWebColumn ") = true then
                action = XY_DisplayWebColumn(action) 
            'URL加密
            elseIf checkFunValue(action, "escape ") = true then
                action = XY_escape(action) 
            'URL解密
            elseIf checkFunValue(action, "unescape ") = true then
                action = XY_unescape(action) 
            'asp与php版本
            elseIf checkFunValue(action, "EDITORTYPE ") = true then
                action = XY_EDITORTYPE(action) 

            '获得网址
            elseIf checkFunValue(action, "getUrl ") = true then
                action = XY_getUrl(action) 

            '文章位置显示信息{}为有动作的
            elseIf checkFunValue(action, "detailPosition ") = true then
                action = XY_detailPosition(action) 

			

                '------------------- 会员 -----------------------
			'关注 
            elseIf checkFunValue(action, "thisArticleFollow ") = true then
                action = XY_thisArticleFollow(action) 
			
                '------------------- 特殊 -----------------------
			'关注 
            elseIf checkFunValue(action, "Request ") = true then
                action = XY_Request(action)  

            '暂时不屏蔽
            elseIf checkFunValue(action, "copyTemplateMaterial ") = true then
                action = "" 
            elseIf checkFunValue(action, "clearCache ") = true then
                action = "" 
            else
                isHnadle = false                                                                '处理为假
            end if 
            '注意这样，有的则不显示 晕 And IsNul(action)=False
            if isNul(action) = true then action = "" 
            if isHnadle = true then
                content = replace(content, s, action) 
            end if 
        end if 
    next 
    handleAction = content 
end function 

'显示网站栏目 新版 把之前网站 导航程序改进过来的
function XY_DisplayWebColumn(action)
    dim i, c, s, url, sql, dropDownMenu, focusType, addSql 
    dim isConcise                                                                   '简洁显示20150212
    dim styleId, styleValue                                                         '样式ID与样式内容
    dim cssNameAddId 
    dim shopnavidwrap                                                               '是否显示栏目ID包
	
	isConcise=false		'简洁显示为假
    styleId = PHPTrim(rParam(action, "styleID")) 
    styleValue = PHPTrim(rParam(action, "styleValue")) 
    addSql = PHPTrim(rParam(action, "addSql")) 
    shopnavidwrap = PHPTrim(rParam(action, "shopnavidwrap")) 
    'If styleId <> "" Then
    'Call ReadNavCSS(styleId, styleValue)
    'End If

    '为数字类型 则自动提取样式内容  20150615
    if checkStrIsNumberType(styleValue) then
        cssNameAddId = "_" & styleValue                                                 'Css名称追加Id编号
    end if 
    sql = "select * from " & db_PREFIX & "webcolumn" 
    '追加sql
    if addSql <> "" then
        sql = getWhereAnd(sql, addSql) 
    end if 
    if checkSql(sql) = false then 
		call eerr("Sql", sql) 
	end if
	'【@是jsp显示@】try{
    rs.open sql, conn, 1, 1 
    dropDownMenu = LCase(rParam(action, "DropDownMenu")) 
    focusType = LCase(rParam(action, "FocusType")) 
    isConcise = IIF(LCase(rParam(action, "isConcise")) = "true", false, true) 

    if isConcise = true then 
		c = c & copyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
	end if
    for i = 1 to rs.recordCount

        '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善
        url = getColumnUrl(rs("columnname"), "name") 
        if rs("columnName") = glb_columnName then
            if focusType = "a" then
                s = copyStr(" ", 8) & "<li class=focus><a href=""" & url & """>" & rs("columname") & "</a>" 
            else
                s = copyStr(" ", 8) & "<li class=focus>" & rs("columnname") 
            end if 
        else
            s = copyStr(" ", 8) & "<li><a href=""" & url & """>" & rs("columnname") & "</a>" 
        end if 
        '网站栏目没有page位置处理
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '处理是否添加在线修改管理器

        c = c & s 

        '小类

        c = c & copyStr(" ", 8) & "</li>" & vbCrLf 

        if isConcise = true  then
			c = c & copyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    	end if
	rs.moveNext : next : rs.close 
	'【@是jsp显示@】}catch(Exception e){} 
    if isConcise = true then 
		c = c & copyStr(" ", 8) & "<li class=right></li>" & vbCrLf 
	end if
    if styleId <> "" then
        c = "<ul class='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</ul>" & vbCrLf 
    end if 
    if shopnavidwrap = "1" or shopnavidwrap = "true" then
        c = "<div id='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</div>" & vbCrLf 
    end if 

    XY_DisplayWebColumn = c 
end function 

'替换全局变量 {$cfg_websiteurl$}
function replaceGlobleVariable(byVal content)
    content = handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl)              '网址
    content = handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate)            '模板
    content = handleRGV(content, "{$cfg_webImages$}", cfg_webImages)                '图片路径
    content = handleRGV(content, "{$cfg_webCss$}", cfg_webCss)                      'css路径
    content = handleRGV(content, "{$cfg_webJs$}", cfg_webJs)                        'js路径
    content = handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle)                  '网站标题
    content = handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords)            '网站关键词
    content = handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription)      '网站描述

    content = handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom)        '网站底部内容

    content = handleRGV(content, "{$glb_columnId$}", glb_columnId)                  '栏目Id
    content = handleRGV(content, "{$glb_columnRootId$}", glb_columnRootId)                  '栏目主Id
    content = handleRGV(content, "{$glb_columnParentId$}", glb_columnParentId)                  '上一级栏目Id
    content = handleRGV(content, "{$glb_columnRootName$}", glb_columnRootName)                  '栏目主名称
    content = handleRGV(content, "{$glb_columnRootEnName$}", glb_columnRootEnName)                  '栏目主英文名称

	
	
    content = handleRGV(content, "{$glb_columnName$}", glb_columnName)              '栏目名称
    content = handleRGV(content, "{$glb_columnEnName$}", glb_columnEnName)              '栏目英文名称
    content = handleRGV(content, "{$glb_columnType$}", glb_columnType)              '栏目类型
    content = handleRGV(content, "{$glb_columnENType$}", glb_columnENType)          '栏目英文类型

    content = handleRGV(content, "{$glb_Table$}", glb_table)                        '表
    content = handleRGV(content, "{$glb_Id$}", glb_id)                              'id
	
	'会员
    content = handleRGV(content, "{$member_id$}", getsession("member_id"))          '会员id
    content = handleRGV(content, "{$member_user$}", getsession("member_user"))          '会员账号 
	

    content = handleRGV(content, "[$模块目录$]", "Module/")                    'Module


    '兼容旧版本 渐渐把它去掉
    content = handleRGV(content, "{$WebImages$}", cfg_webImages)                    '图片路径
    content = handleRGV(content, "{$WebCss$}", cfg_webCss)                          'css路径
    content = handleRGV(content, "{$WebJs$}", cfg_webJs)                            'js路径
    content = handleRGV(content, "{$Web_Title$}", cfg_webTitle) 
    content = handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords) 
    content = handleRGV(content, "{$Web_Description$}", cfg_webDescription) 


    content = handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE)                      '后缀
    content = handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                    '首页显示网址
    '文章用到
    content = handleRGV(content, "{$glb_articleAuthor$}", glb_articleAuthor)        '文章作者
    content = handleRGV(content, "{$glb_articleAdddatetime$}", glb_articleAdddatetime) '文章添加时间
    content = handleRGV(content, "{$glb_articlehits$}", glb_articlehits)            '文章点击次数

    content = handleRGV(content, "{$glb_upArticle$}", glb_upArticle)                '上一篇文章
    content = handleRGV(content, "{$glb_downArticle$}", glb_downArticle)            '下一篇文章
    content = handleRGV(content, "{$glb_upArticleUrl$}", glb_upArticleUrl)                '上一篇文章链接
    content = handleRGV(content, "{$glb_downArticleUrl$}", glb_downArticleUrl)            '下一篇文章链接
	
    content = handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags) '文章标签组
    content = handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage)    '文章大图
    content = handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage) '文章小图
    content = handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord)        '首页显示网址
	
	
	
    content = handleRGV(content, "{$member_expiredatetime$}", getSession("member_expiredatetime"))        '会员超时
    content = handleRGV(content, "{$member_user$}", getSession("member_user"))        '会员账号
	
	
	
    content = handleRGV(content, "{$pageInfo$}", gbl_PageInfo)        '显示翻页信息
    content = handleRGV(content, "{$detailTitle$}", glb_detailTitle)        '首页详细页标题
    content = handleRGV(content, "{$detailContent$}", glb_bodyContent)        '显示详细页内容
	
	
    content = handleRGV(content, "{$glb_detailTitle$}", glb_detailTitle)        '文章细节标题
    content = handleRGV(content, "{$glb_flags$}", glb_flags)        '旗
    content = handleRGV(content, "{$glb_smallimage$}", glb_smallimage)        '小图
    content = handleRGV(content, "{$glb_bigimage$}", glb_bigimage)        '大图
    content = handleRGV(content, "{$glb_bannerimage$}", glb_bannerimage)        '大图
	
	
    content = handleRGV(content, "{$glb_author$}", glb_author)        '作者
    content = handleRGV(content, "{$glb_sortrank$}", glb_sortrank)        '排序
    content = handleRGV(content, "{$glb_aboutcontent$}", glb_aboutcontent)        '介绍
    content = handleRGV(content, "{$glb_bodyContent$}", glb_bodyContent)        '内容
 
 
    replaceGlobleVariable = content 
end function 

'处理替换
function handleRGV(byVal content, findStr, replaceStr)
    dim lableName 
    '对[$$]处理
    lableName = mid(findStr, 3, len(findStr) - 4) & " " 
    lableName = mid(lableName, 1, inStr(lableName, " ") - 1) 
    content = replaceValueParam(content, lableName, replaceStr) 
    content = replaceValueParam(content, LCase(lableName), replaceStr) 
    '直接替换{$$}这种方式，兼容之前网站
    content = replace(content, findStr, replaceStr) 
    content = replace(content, LCase(findStr), replaceStr) 
    handleRGV = content 
end function 

'加载网站配置信息
sub loadWebConfig()
    dim templatedir 
    call openconn() 
	'【@是jsp显示@】try{
    rs.open "select * from " & db_PREFIX & "website", conn, 1, 1 
    if not rs.EOF then
        cfg_webSiteUrl = phptrim(rs("webSiteUrl"))                                      '网址
        cfg_webTemplate = webDir & phptrim(rs("webTemplate"))                           '模板路径
        cfg_webImages = webDir & phptrim(rs("webImages"))                               '图片路径
        cfg_webCss = webDir & phptrim(rs("webCss"))                                     'css路径
        cfg_webJs = webDir & phptrim(rs("webJs"))                                       'js路径
        cfg_webTitle = rs("webTitle")                                                   '网址标题
        cfg_webKeywords = rs("webKeywords")                                             '网站关键词
        cfg_webDescription = rs("webDescription")                                       '网站描述
        cfg_webSiteBottom = rs("webSiteBottom")                                         '网站地底
        cfg_flags = rs("flags")                                                         '旗
		templatesetcode=rs("templatesetcode")											'读出模板编码 
		isMemberVerification=IIF(rs("isMemberVerification")=1,true,false)				'会员检测判断
		

        '改换模板20160202
        if request("templatedir") <> "" then
            '删除绝对目录前面的目录，不需要那个东西20160414
            templatedir = replace(handlePath(request("templatedir")), handlePath("/"), "/") 
            'call eerr("templatedir",templatedir)

            if(inStr(templatedir, ":") > 0 or inStr(templatedir, "..") > 0) and getIP() <> "127.0.0.1" then
                call eerr("提示", "模板目录有非法字符") 
            end if 
            templatedir = handlehttpurl(replace(templatedir, handlePath("/"), "/")) 

            cfg_webImages = replace(cfg_webImages, cfg_webTemplate, templatedir) 
            cfg_webCss = replace(cfg_webCss, cfg_webTemplate, templatedir) 
            cfg_webJs = replace(cfg_webJs, cfg_webTemplate, templatedir) 
            cfg_webTemplate = templatedir 
        end if  
    end if : rs.close  
	'【@是jsp显示@】}catch(Exception e){} 
end sub 



'显示管理列表
function getDetailList(action, content, actionName, lableTitle, byVal fieldNameList, nPageSize, sPage, addSql)
    call openconn() 
    dim defaultStr, i, s, c, tableName, j, splxx, sql, nPage 
    dim nX, url, nCount 
    dim pageInfo, nModI, startStr, endStr 

    dim fieldName                                                                   '字段名称
    dim splFieldName                                                                '分割字段

    dim replaceStr                                                                  '替换字符
    tableName = LCase(actionName)                                                   '表名称
    dim listFileName                                                                '列表文件名称
    listFileName = rParam(action, "listFileName") 
    dim abcolorStr                                                                  'A加粗和颜色
    dim atargetStr                                                                  'A链接打开方式
    dim atitleStr                                                                   'A链接的title20160407
    dim anofollowStr                                                                'A链接的nofollow

    dim id, idPage 
    id = rq("id") 
    call checkIDSQL(request("id")) 

    if fieldNameList = "*" then
        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    end if 

    fieldNameList = specialStrReplace(fieldNameList)                                '特殊字符处理
    splFieldName = split(fieldNameList, ",")                                        '字段分割成数组


    defaultStr = getStrCut(content, "<!--#body start#-->", "<!--#body end#-->", 2) 



    pageInfo = getStrCut(content, "[page]", "[/page]", 1) 
    if pageInfo <> "" then
        content = replace(content, pageInfo, "") 
    end if 
    'call eerr("pageInfo",pageInfo)

    sql = "select * from " & db_PREFIX & tableName & " " & addSql 
    '检测SQL
    if checksql(sql) = false then
        call errorLog("出错提示：<br>sql=" & sql & "<br>") 
        exit function 
    end if 
	'【@是jsp显示@】try{
    rs.open sql, conn, 1, 1 
    '【PHP】删除rs
    nCount = rs.recordCount 

    '为动态翻页网址
    if isMakeHtml = true then
        url = "" 
        if len(listFileName) > 5 then
            url = mid(listFileName, 1, len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        end if 
    else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 
	gbl_PageInfo=webPageControl(nCount, nPageSize, cstr(sPage), url, pageInfo)
    content = replace(content, "[$pageInfo$]", gbl_PageInfo) 

    if EDITORTYPE = "asp" then
        nX = getRsPageNumber(rs, nCount, nPageSize, sPage)                              '【@不是asp屏蔽@】   获得Rs页数      记录总数
   

	elseif EDITORTYPE = "aspx" then 	
			
		'【@是.netc显示@】int  nCountPage = getCountPage(nCount, nPageSize);
		'【@是.netc显示@】if(nPage<=1){
		'【@是.netc显示@】	nX=nPageSize; 
		'【@是.netc显示@】	if(nX>nCount){
		'【@是.netc显示@】		nX=nCount;
		'【@是.netc显示@】	} 
		'【@是.netc显示@】}else{
		'【@是.netc显示@】	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
		'【@是.netc显示@】		rs.Read();
		'【@是.netc显示@】	}
		'【@是.netc显示@】	if(nPage<nCountPage){
		'【@是.netc显示@】		nX=nPageSize;
		'【@是.netc显示@】	}else{
		'【@是.netc显示@】		nX=nCount-nPageSize*(nPage-1);
		'【@是.netc显示@】	}
		'【@是.netc显示@】} 
   
	elseif EDITORTYPE = "jsp" then
	
		'【@是jsp显示@】int  nCountPage = getCountPage(nCount, nPageSize);
		'【@是jsp显示@】rs = Conn.executeQuery(sql);			
		'【@是jsp显示@】if(nPage<=1){
		'【@是jsp显示@】	nX=nPageSize;
		'【@是jsp显示@】	if(nX>nCount){
		'【@是jsp显示@】		nX=nCount;
		'【@是jsp显示@】	} 
		'【@是jsp显示@】}else{
		'【@是jsp显示@】	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
		'【@是jsp显示@】		rs.next();
		'【@是jsp显示@】	}
		'【@是jsp显示@】	if(nPage<nCountPage){
		'【@是jsp显示@】		nX=nPageSize;
		'【@是jsp显示@】	}else{
		'【@是jsp显示@】		nX=nCount-nPageSize*(nPage-1);
		'【@是jsp显示@】	}
		'【@是jsp显示@】}  

	else									 
		if nPage<>0 then'【@是.netc屏蔽@】'【@是jsp屏蔽@】
			nPage = nPage - 1 '【@是.netc屏蔽@】'【@是jsp屏蔽@】
		end if '【@是.netc屏蔽@】'【@是jsp屏蔽@】
		sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize '【@是.netc屏蔽@】'【@是jsp屏蔽@】
		rs.open sql, conn, 1, 1 '【@是.netc屏蔽@】'【@是jsp屏蔽@】
		'【PHP】删除rs
		nX = rs.recordCount '【@是.netc屏蔽@】'【@是jsp屏蔽@】
    end if 
    'call echo("sql",sql)
    for i = 1 to nX
        '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善
		'【@是.netc显示@】rs.Read();
		'【@是jsp显示@】rs.next();
        startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 

        '在最后时排序当前交点20160202
        if i = nX then
            startStr = "[list-end]" : endStr = "[/list-end]" 
        end if 

        '例[list-mod2]  [/list-mod2]    20150112
        for nModI = 6 to 2 step - 1
            if inStr(defaultStr, startStr) = false and i mod nModI = 0 then
                startStr = "[list-mod" & nModI & "]" : endStr = "[/list-mod" & nModI & "]" 
                if inStr(defaultStr, startStr) > 0 then
                    exit for 
                end if 
            end if 
        next 

        '没有则用默认
        if inStr(defaultStr, startStr) = false or startStr = "" then
            startStr = "[list]" : endStr = "[/list]" 
        end if 

        if inStr(defaultStr, startStr) > 0 and inStr(defaultStr, endStr) > 0 then
            s = strCut(defaultStr, startStr, endStr, 2) 

            's = defaultStr
            s = replace(s, "[$id$]", rs("id")) 
            for j = 0 to uBound(splFieldName)
                if splFieldName(j) <> "" then
                    splxx = split(splFieldName(j) & "|||", "|") 
                    fieldName = splxx(0) 
                    replaceStr = rs(fieldName) & "" 
                    s = replaceValueParam(s, fieldName, replaceStr) 
                end if 

                if isMakeHtml = true then
                    url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
                else
                    url = handleWebUrl("?act=detail&id=" & rs("id")) 
                    if rs("customAUrl") <> "" then
                        url = rs("customAUrl") 
                    end if 
                end if 

                'A链接添加颜色
                abcolorStr = "" 
                if inStr(fieldNameList, ",titlecolor,") > 0 then
                    'A链接颜色
                    if rs("titlecolor") <> "" then
                        abcolorStr = "color:" & rs("titlecolor") & ";" 
                    end if 
                elseif inStr(fieldNameList, ",fontcolor,") > 0 then
                    'A链接颜色
                    if rs("fontcolor") <> "" then
                        abcolorStr = "color:" & rs("fontcolor") & ";" 
                    end if 
                end if  
                if inStr(fieldNameList, ",flags,") > 0 then
                    'A链接加粗
                    if inStr(rs("flags"), "|b|") > 0 then
                        abcolorStr = abcolorStr & "font-weight:bold;" 
                    end if 
                end if 
                if abcolorStr <> "" then
                    abcolorStr = " style=""" & abcolorStr & """" 
                end if 

                '打开方式2016
                if inStr(fieldNameList, ",target,") > 0 then
                    atargetStr = IIF(rs("target") <> "", " target=""" & rs("target") & """", "") 
                end if 

                'A的title
                if inStr(fieldNameList, ",title,") > 0 then
                    atitleStr = IIF(rs("title") <> "", " title=""" & rs("title") & """", "") 
                end if 

                'A的nofollow
                if inStr(fieldNameList, ",nofollow,") > 0 then
                    anofollowStr = IIF(rs("nofollow") <> 0, " rel=""nofollow""", "") 
                end if 



				s = replaceValueParam(s, "i", i)                                                '循环编号
				s = replaceValueParam(s, "编号", i)                                             '循环编号
			
                s = replaceValueParam(s, "url", url) 
                s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A链接加颜色与加粗
                s = replaceValueParam(s, "atitle", atitleStr)                                   'A链接title
                s = replaceValueParam(s, "anofollow", anofollowStr)                             'A链接nofollow
                s = replaceValueParam(s, "atarget", atargetStr)                                 'A链接打开方式


            next 
        end if 
		if EDITORTYPE <> "jsp" then
        idPage = getThisIdPage(db_PREFIX & tableName, rs("id"), 10) 
        '【留言】
        if tableName = "guestbook" then
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=GuestBook&lableTitle=留言&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" & idPage & "&id=" & rs("id") & "&n=" & getRnd(11) 
        '【招聘】
        elseif tableName = "job" then
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=Job&lableTitle=招聘&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" & idPage & "&id=" & rs("id") & "&n=" & getRnd(11) 
        '【默认显示文章】
        else
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=" & idPage & "&parentid=" & rs("parentid") & "&id=" & rs("id") & "&n=" & getRnd(11) 
        end if 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span") 
		end if
        c = c & s 
    rs.moveNext : next : rs.close 
	'【@是jsp显示@】}catch(Exception e){} 
    content = replace(content, "<!--#body start#-->" & defaultStr & "<!--#body end#-->", c) 

    if isMakeHtml = true then
        url = "" 
        if len(listFileName) > 5 then
            url = mid(listFileName, 1, len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        end if 
    else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 

    getDetailList = content 
end function 


'****************************************************
'默认列表模板
function defaultListTemplate(sType, sName)
    dim c, templateHtml, listTemplate, startStr, endStr, lableName 

    templateHtml=readFile(cfg_webTemplate & "/" & templateName,templatesetcode)
    'templateHtml = getFText(cfg_webTemplate & "/" & templateName) 
    '从栏目名称搜索，到栏目类型，到默认20160630
    lableName = sName & "list" 
    startStr = "<!--#" & lableName & " start#-->" 
    endStr = "<!--#" & lableName & " end#-->" 
    if inStr(templateHtml, startStr) = false or inStr(templateHtml, endStr) = false then
        lableName = sType & "list" 
        startStr = "<!--#" & lableName & " start#-->" 
        endStr = "<!--#" & lableName & " end#-->" 
    end if 
    if inStr(templateHtml, startStr) = false or inStr(templateHtml, endStr) = false then
        lableName = "list" 
        startStr = "<!--#" & lableName & " start#-->" 
        endStr = "<!--#" & lableName & " end#-->" 
    end if 

    'call rwend(templateHtml)
    if inStr(templateHtml, startStr) > 0 and inStr(templateHtml, endStr) > 0 then
        listTemplate = strCut(templateHtml, startStr, endStr, 2) 
    else
        startStr = "<!--#" & lableName 
        endStr = "#-->" 
        if inStr(templateHtml, startStr) > 0 and inStr(templateHtml, endStr) > 0 then
            listTemplate = strCut(templateHtml, startStr, endStr, 2) 
        end if 
    end if 
    if listTemplate = "" then
        c = "<ul class=""list""><!--#body start#-->" & vbCrLf 
        c = c & "[list]    <li><a href=""[$url$]""[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class=""time"">[$adddatetime format_time='1'$]</span></li>" & vbCrLf 
        c = c & "[/list]11111111111<!--#body end#--> " & vbCrLf 
        c = c & "</ul>" & vbCrLf 
        c = c & "<div class=""clear10""></div>" & vbCrLf 
        c = c & "<div>[$pageInfo$]</div>" & vbCrLf 
        listTemplate = c 
    end if 
    'call rwend(listTemplate)

    defaultListTemplate = listTemplate 
end function 

call loadRun()                                                                  '【@是.netc屏蔽@】'【@是jsp屏蔽@】
'加载就运行
sub loadRun()
	dim c
    '这是为了给.net使用的，因为在.net里面全局变量不能有变量
	WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE)	
	
    '缓冲处理20160622
    cacheHtmlFilePath = "/cache/html/" & setFileName(getThisUrlFileParam()) & ".html" 
    '启用缓冲
    if request("cache") <> "false" and isOnCacheHtml = true then
        if checkFile(cacheHtmlFilePath) = true then
            'call echo("读取缓冲文件","OK")
            call rwend(getftext(cacheHtmlFilePath)) 
        end if 
    end if 

    '记录表前缀
    if request("db_PREFIX") <> "" then
        db_PREFIX = request("db_PREFIX") 
    elseIf getsession("db_PREFIX") <> "" then
        db_PREFIX = getsession("db_PREFIX") 
    end if 
    '加载网址配置
    call loadWebConfig() 
    isMakeHtml = false                                                              '默认生成HTML为关闭
    if request("isMakeHtml") = "1" or request("isMakeHtml") = "true" then
        isMakeHtml = true 
    end if 
    templateName = request("templateName")                                          '模板名称


    '保存数据处理页
	if request("act")= "savedata" then
		call saveData(request("stype")) : response.end()                   '保存数据
	''站长统计 | 今日IP[653] | 今日PV[9865] | 当前在线[65]')
	elseif request("act")= "webstat" then 
		call webStat(adminDir & "/Data/Stat/") : response.end()             '网站统计
	
	elseif request("act")= "saveSiteMap" then
		isMakeHtml = true : call saveSiteMap() : response.end()         '保存sitemap.xml
	
	elseif request("act")= "member" then
		call callMember()	:call eerr("","sdf")											'调用Member文件函数
	
	elseif request("act")= "handleAction" then
		if request("ishtml") = "1" then
			isMakeHtml = true 
		end if
		call rwend(handleAction(request("content")))                                         '处理动作 
    '生成html
    elseif request("act") = "makehtml" then
        call echo("makehtml", "makehtml") 
        isMakeHtml = true 
        call makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' id='" & request("id") & "' template='"& request("template") &"' ") 
        call writeToFile("index.html", code,makeHtmlFileSetCode) 

    '复制Html到网站
    elseIf request("act") = "copyHtmlToWeb" then
        call copyHtmlToWeb() 
    '全部生成
    elseIf request("act") = "makeallhtml" then
        call makeAllHtml("", "", request("id")) 

    '生成当前页面
    elseIf request("isMakeHtml") <> "" and request("isSave") <> "" then

        call handlePower("生成当前HTML页面")                                            '管理权限处理
        call writeSystemLog("", "生成当前HTML页面")                                     '系统日志

        isMakeHtml = true 


        call checkIDSQL(request("id")) 
        call rw(makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "' template='"& request("template") &"' ")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" then
            glb_filePath = glb_filePath & "index.html" 
        elseIf glb_filePath = "" and glb_columnType = "首页" then
            glb_filePath = "index.html" 
        end if 
        '文件不为空  并且开启生成html
        if glb_filePath <> "" and glb_isonhtml = true then
            call createDirFolder(getFileAttr(glb_filePath, "1")) 
            call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
            if request("act") = "detail" then
                conn.execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & request("id")) 
            elseIf request("act") = "nav" then
                if request("id") <> "" then
                    conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & request("id")) 
                else
                    conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where columnname='" & request("columnName") & "'") 
                end if 
            end if 
            call echo("生成文件路径", "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>") 

            '新闻则批量生成 20160216
            if glb_columnType = "新闻" then
                call makeAllHtml("", "", glb_columnId) 
            end if 

        end if 

    '全部生成
    elseIf request("act") = "Search" then
        call rw(makeWebHtml("actionType='Search' npage='"& IIF(request("page")="","1",request("page")) &"'  template='"& request("template") &"'")) 
    else
        if LCase(request("issave")) = "1" then
            call makeAllHtml(request("columnType"), request("columnName"), request("columnId")) 
        else
            call checkIDSQL(request("id")) 
			c=makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "'  template='"& request("template") &"'")
            
			'测试用
			if host()="http://aa/" then
				'call echo("",host())
				c=replace(c,"http://sharembweb.com/asptoLanguage/","")
				c=replace(c,"http://sharembweb.com/AspToLanguage/","")
				c=replace(c,"http://sharembweb.com/toolslist/","/toolslist/")
				c=replace(c,"http://sharembweb.com/Tools/FormattingTools/","/函数/")
			end if
			call rw(c) 
        end if 
    end if 
  
    '开启缓冲html
    if isOnCacheHtml = true then
        call createFile(cacheHtmlFilePath, code)                                        '保存到缓冲文件里20160622
    end if 
end sub 
'检测ID是否SQL安全
function checkIDSQL(id)
    if checkNumber(id) = false and id <> "" then
        call eerr("提示", "id中有非法字符") 
    end if 
end function 
'http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
'http://127.0.0.1/aspweb.asp?act=detail&id=75
'生成html静态页
function makeWebHtml(action)
    dim actionType, npagesize, npage, s, url, addSql, sortSql, sortFieldName, ascOrDesc,thisTable,thisSql,customTemplate
    dim parentid                                                  '追加于20160716 home
	dim isOK,relatedtags,sortFieldValue,sql
	isOK=false			'真假
    actionType = rParam(action, "actionType") 
    s = rParam(action, "npage") 
    s = getnumber(s) 
    if s = "" then
        npage = 1 
    else
        npage = CInt(s) 
    end if
	
   	customTemplate = rParam(action, "template") 					'自定义模板
	'call echo("action",action)
	
    '导航
    if actionType = "nav" then
        glb_columnType = rParam(action, "columnType") 
        glb_columnName = rParam(action, "columnName") 
        glb_columnId = rParam(action, "columnId") 
        if glb_columnId = "" then
            glb_columnId = rParam(action, "id") 
        end if 
        if glb_columnType <> "" then
            addSql = "where columnType='" & glb_columnType & "'" 
        end if 
        if glb_columnName <> "" then
            addSql = getWhereAnd(addSql, "where columnName='" & glb_columnName & "'") 
        end if 
        if glb_columnId <> "" then
            addSql = getWhereAnd(addSql, "where id=" & glb_columnId & "") 
        end if 		
		'【@是jsp显示@】try{
		thisTable="webcolumn"
		thisSql="Select * from " & db_PREFIX & thisTable & " " & addSql
        rs.open thisSql, conn, 1, 1
        if not rs.EOF then
            glb_columnId = rs("id")
            glb_columnName = rs("columnname")
			glb_columnEnName=rs("columnenname")
            glb_columnType = rs("columntype") 
            glb_bodyContent = rs("bodycontent") 
            glb_detailTitle = glb_columnName 
			glb_memberusercheck=rs("memberusercheck")									'会员检测
            glb_flags = rs("flags") 
            npagesize = rs("npagesize")                                                     '每页显示条数
            glb_isonhtml = IIF(rs("isonhtml") = 0, false, true)                          '是否生成静态网页
            sortSql = " " & rs("sortsql")                                                   '排序SQL
			glb_bigimage=rs("bigimage")														'大图
			glb_smallimage=rs("smallimage")													'小图
			glb_bannerimage=rs("bannerimage")												'banner图
			
            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            end if 
            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                elseIf rs("columntype") <> "首页" then
                    templateName = getDateilTemplate(rs("id"), "List") 
                end if 
            end if 
			
        end if : rs.close 
		call handleColumnRoot(glb_columnId) 	'获得栏目主ID与主名称
		
		'【@是jsp显示@】}catch(Exception e){} 
        glb_columnENType = handleColumnType(glb_columnType) 
        glb_url = getColumnUrl(glb_columnName, "name") 

        '文章类列表
        if inStr("|产品|新闻|视频|下载|案例|", "|" & glb_columnType & "|") > 0 then
		
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "栏目列表", "*", npagesize, cstr(npage), "where parentid in("& getColumnIdList(glb_columnId,"addthis") &") " & sortSql)  
			
        '留言类列表
        elseIf inStr("|留言|", "|" & glb_columnType & "|") > 0 then
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "GuestBook", "留言列表", "*", npagesize, cstr(npage), " where isthrough<>0 " & sortSql) 
        '招聘类列表
        elseIf inStr("|招聘|", "|" & glb_columnType & "|") > 0 then 
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "Job", "招聘列表", "*", npagesize, cstr(npage), " where isthrough<>0 " & sortSql) 
        elseIf glb_columnType = "文本" then
            '航行栏目加管理
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & glb_columnId & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span")

        end if 
    '细节
    elseIf actionType = "detail" then
        glb_locationType = "detail" 
		thisTable="articledetail"		
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id")
		'【@是jsp显示@】try{
        rs.open thisSql, conn, 1, 1 
        if not rs.EOF then
			isOK=true
			
            glb_columnId = rs("parentid")  
			
			if glb_columnRootId="" then
				glb_columnRootId=getColumnRootId(rs("id"))					'获得主栏目ID 为空时处理下
			end if
			glb_columnParentId=getParentColumnId(rs("parentid"))		'获得上一级栏目ID
			glb_columnParentName=getParentColumnName(rs("parentid"))		'获得上一级栏目名称
			
			
			
            glb_detailTitle = rs("title") 
            glb_flags = rs("flags") 
            glb_smallimage = rs("smallimage") 
            glb_bigimage = rs("bigimage") 
            glb_author = rs("author") 
            glb_sortrank = rs("sortrank")
            glb_aboutcontent = rs("aboutcontent")
            glb_bodyContent = rs("bodycontent")
			 
            glb_isonhtml =IIF(rs("isonhtml") = 0, false, true)   
            glb_id = rs("id")                                                               '文章ID
            if isMakeHtml = true then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            end if 

            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            end if 


            glb_articleAuthor = rs("author") 
            glb_articleAdddatetime = rs("adddatetime")
			parentid=rs("parentid")
			relatedtags=rs("relatedtags")
			
			glb_relatedTags=relatedtags
            glb_aritcleSmallImage = rs("smallimage") 
            glb_aritcleBigImage = rs("bigimage") 
            glb_articlehits = rs("hits") 
 
            glb_bodyContent = rs("bodycontent") 
			 
            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                else
                    templateName = getDateilTemplate(rs("parentid"), "Detail") 
                end if 
            end if  
        end if : rs.close 
		if isOK=true then
			call handleColumnRoot(glb_columnId) 	'获得栏目主ID与主名称
		
            '改进20160628
            sortFieldName = "id" 
            ascOrDesc = "asc" 
            addSql = trim(getWebColumnSortSql(parentid)) 
            if addSql <> "" then
                sortFieldName = trim(replace(replace(replace(addSql, "order by", ""), " desc", ""), " asc", "")) 
                if instr(addSql, " desc") > 0 then
                    ascOrDesc = "desc" 
                end if 
            end if
			'获得排序对应字符的值
			rs.open thisSql,conn,1,1
			if not rs.eof then
				sortFieldValue=rs(sortFieldName) 
			end if:rs.close
			
            glb_columnName = getColumnName(parentid) 
            glb_upArticle = upArticle(parentid, sortFieldName, sortFieldValue, ascOrDesc,glb_upArticleUrl) 
            glb_downArticle = downArticle(parentid, sortFieldName, sortFieldValue, ascOrDesc,glb_downArticleUrl) 
            glb_aritcleRelatedTags = aritcleRelatedTags(relatedtags)
            conn.execute("update " & db_PREFIX & "articledetail set hits=hits+1 where id=" & glb_id) '更新点击数

            '文章详细加控制
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & rParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 
			
		end if
		'【@是jsp显示@】}catch(Exception e){} 
	
    '单页
    elseIf actionType = "onepage" then
		'【@是jsp显示@】try{
		
		thisTable="onepage"		
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id") 
        rs.open thisSql, conn, 1, 1 
        if not rs.EOF then
            glb_detailTitle = rs("title") 
            glb_isonhtml = IIF(rs("isonhtml") = 0, false, true)                   '是否生成静态网页
            if isMakeHtml = true then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/page/page" & rs("id")) 
            else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            end if 

            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '网址标题
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '网站关键词
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '网站描述
            end if 
            '内容
            glb_bodyContent = rs("bodycontent") 


            '文章详细加控制
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=分类信息&nPageSize=10&page=&parentid=&id=" & rParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 


            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                else
                    templateName = "Main_Model.html" 
                end if 
            end if 

        end if : rs.close 
		'【@是jsp显示@】}catch(Exception e){}

    '搜索
    elseIf actionType = "Search" then
		templateName="Main_Model.html"
		if request("template") <>"" then
			templateName=request("template") & ".html"
		end if 
        searchKeyWordName = request("keywordname") 
		glb_searchFieldName= request("searchfieldname")		'搜索字段名称
        parentid = request("parentid") 
        if searchKeyWordName = "" then
            searchKeyWordName = "wd" 
        end if
		if glb_searchFieldName="" then
			glb_searchFieldName="title"
		end if
		'call echo("searchKeyWordName",searchKeyWordName)
        glb_searchKeyWord = replace(request(searchKeyWordName), "<", "&lt;") 
        addSql = "" 
        if parentid <> "" then
            addSql = " where parentid=" & parentid
		elseif request("subparentid")<>"" then
			glb_columnRootId=trim(request("subparentid"))		'记录下，给搜索用，恩
			addSql = " where parentid in(" & getColumnIdList(request("subparentid"),"addthis") & ")"
			'call eerr("addSql",addSql)
        end if 
        addSql = getWhereAnd(addSql, " where "& glb_searchFieldName &" like '%" & glb_searchKeyWord & "%'")
		'call echo("addsql",addsql)
        npagesize = 20  
        glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "网站栏目", "*", npagesize, cstr(npage), addSql) 
        positionEndStr = " >> 搜索内容”" & glb_searchKeyWord & "“" 
    '加载等待
    elseIf actionType = "loading" then
        call rwend("页面正在加载中。。。") 
	'表 index.asp?act=table_HuoYuan&id=1
    elseIf left(actionType,6) = "table_" then
		thisTable=mid(actionType,7)
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id") 
 
    end if 
    '20170603追加
	if customTemplate<>"" then
		templateName=customTemplate & ".html"
	'模板为空，则用默认首页模板
    elseif templateName = "" then
        templateName = "Index_Model.html"                                               '默认模板
    end if 
	'处理会员验证20170522
	templateName=getMemberVerificationTemplateName(templateName) 
	
    '检测当前路径是否有模板
    if inStr(templateName, "/") = false then
        templateName = cfg_webTemplate & "/" & templateName 
    end if 
    'call echo("templateName",templateName)
    if checkFile(templateName) = false then
        call eerr("未找到模板文件", templateName) 
    end if
	
	code=readFile(templateName,templatesetcode)			'模板可自定义编码 
    'code = getftext(templateName) 

    code = handleAction(code)                                                       '处理动作
    'code = thisPosition(code)                                                       '位置
    code = replaceGlobleVariable(code)                                              '替换全局标签
    code = handleAction(code)                                                       '处理动作    '再来一次，处理数据内容里动作

    'call die(code)
    code = handleAction(code)                                                       '处理动作
    code = handleAction(code)                                                       '处理动作
    'code = thisPosition(code)                                                       '位置
    code = replaceGlobleVariable(code)                                              '替换全局标签
    code = delTemplateMyNote(code)                                                  '删除无用内容
    code = handleAction(code)                                                       '处理动作
	
	if thisTable <>"" then
		'call eerr(thisTable,thisSql)
		code=handleReplaceTableFieldList(code,thisTable,thisSql,"this_glb_","")
	end if
 
'call echo("templateName",templateName)
'call eerr("customTemplate",customTemplate)

    '格式化HTML
    if inStr(cfg_flags, "|formattinghtml|") > 0 then
        'code = HtmlFormatting(code)        '简单
        code = handleHtmlFormatting(code, false, 0, "删除空行")                         '自定义
    '格式化HTML第二种
    elseIf inStr(cfg_flags, "|formattinghtmltow|") > 0 then
        code = htmlFormatting(code)                                                     '简单
        code = handleHtmlFormatting(code, false, 0, "删除空行")                         '自定义
    '压缩HTML
    elseIf inStr(cfg_flags, "|ziphtml|") > 0 then
        code = ziphtml(code) 

    end if 
    '闭合标签
    if inStr(cfg_flags, "|labelclose|") > 0 then
        code = handleCloseHtml(code, true, "")                                          '图片自动加alt  "|*|",
    end if 

    '在线编辑20160127
    if rq("gl") = "edit" then
        if inStr(code, "<head>") > 0 then
            if inStr(lcase(code), "jquery.min.js") = false then
                code = replace(code, "<head>", "<head><script src=""/Jquery/jquery.Min.js""></script>") 
            end if 
		end if
		if inStr(code, "</head>") > 0 then
            code = replace(code, "</head>", "<script src=""/Jquery/Callcontext_menu.js""></script></head>") 
        end if 
        if inStr(code, "<body>") > 0 then
        'Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        end if 
    end if 
    'call echo(templateName,templateName)
    makeWebHtml = code 
end function 

'获得默认细节模板页
function getDateilTemplate(parentid, templateType)
    dim templateName,tempS
    templateName = "Main_Model.html" 
	'【@是jsp显示@】try{
    rsx.open "select * from " & db_PREFIX & "webcolumn where id=" & parentid, conn, 1, 1 
    if not rsx.EOF then
        'call echo("columntype",rsx("columntype"))
        if rsx("columntype") = "新闻" then
            '新闻细节页
            if checkFile(cfg_webTemplate & "/News_" & templateType & ".html") = true then
                templateName = "News_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "产品" then
            '产品细节页
            if checkFile(cfg_webTemplate & "/Product_" & templateType & ".html") = true then
                templateName = "Product_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "下载" then
            '下载细节页
            if checkFile(cfg_webTemplate & "/Down_" & templateType & ".html") = true then
                templateName = "Down_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "招聘" then
            '下载细节页
            if checkFile(cfg_webTemplate & "/Job_" & templateType & ".html") = true then
                templateName = "Job_" & templateType & ".html" 
            end if 

        elseIf rsx("columntype") = "视频" then
            '视频细节页
            if checkFile(cfg_webTemplate & "/Video_" & templateType & ".html") = true then
                templateName = "Video_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "留言" then
            '视频细节页
            if checkFile(cfg_webTemplate & "/GuestBook_" & templateType & ".html") = true then
                templateName = "GuestBook_" & templateType & ".html" 
            end if 
        
		elseIf rsx("columntype") = "文本" then
            '视频细节页
            if checkFile(cfg_webTemplate & "/Page_" & templateType & ".html") = true then
                templateName = "Page_" & templateType & ".html" 
            end if 
        end if  
		if rsx("columntype") <> "文本" and (templateType="List" Or templateType="Detail") and templateName="Main_Model.html"  then			 
			tempS="Default_" & templateType & ".html" 
			if checkFile(cfg_webTemplate & "/" & tempS) = true then
				templateName =tempS
			end if
			'call echo("templateName",templateName) 
		end if
		'call echo("templateType",templateType)
    end if : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    getDateilTemplate = templateName 

end function 


'生成全部html页面
sub makeAllHtml(columnType, columnName, columnId)
    dim action, s, i, nPageSize, nCountSize, nPage, addSql, url, articleSql 
    call handlePower("生成全部HTML页面")                                            '管理权限处理
    call writeSystemLog("", "生成全部HTML页面")                                     '系统日志

    isMakeHtml = true 
    '栏目
    call echo("栏目", columnName) 
    if columnType <> "" then
        addSql = "where columnType='" & columnType & "'" 
    end if 
    if columnName <> "" then
        addSql = getWhereAnd(addSql, "where columnName='" & columnName & "'") 
    end if 
    if columnId <> "" then
        addSql = getWhereAnd(addSql, "where id in(" & columnId & ")") 
    end if 
	'【@是jsp显示@】try{
    rss.open "select * from " & db_PREFIX & "webcolumn " & addSql & " order by sortrank asc", conn, 1, 1 
    while not rss.EOF
        glb_columnName = "" 
        '开启生成html
        if rss("isonhtml") <>0 then
            if inStr("|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单|", "|" & rss("columntype") & "|") > 0 then
                if rss("columntype") = "留言" then
                    nCountSize = getRecordCount(db_PREFIX & "guestbook", "")                        '记录数
                elseif rss("columntype") = "招聘" then
                    nCountSize = getRecordCount(db_PREFIX & "job", "")                        '记录数
                else
                    nCountSize = getRecordCount(db_PREFIX & "articledetail", " where parentid=" & rss("id")) '记录数
                end if 
                nPageSize = rss("npagesize") 
                nPage = getPageNumb(CInt(nCountSize), CInt(nPageSize)) 
                if nPage <= 0 then
                    nPage = 1 
                end if 
                for i = 1 to nPage
                    url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/nav" & rss("id")) 
                    glb_filePath = replace(url, cfg_webSiteUrl, "") 
                    if right(glb_filePath, 1) = "/" or glb_filePath = "" then
                        glb_filePath = glb_filePath & "index.html" 
                    end if 
                    'call echo("glb_filePath",glb_filePath)
                    action = " action actionType='nav' columnName='" & rss("columnname") & "' npage='" & i & "' listfilename='" & glb_filePath & "' " 
                    'call echo("action",action)
                    call makeWebHtml(action) 
                    if i > 1 then
                        glb_filePath = mid(glb_filePath, 1, len(glb_filePath) - 5) & i & ".html" 
                    end if 
                    s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                    call echo(action, s) 
                    if glb_filePath <> "" then
                        call createDirFolder(getFileAttr(glb_filePath, "1")) 
                        call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                    end if 
                    doevents() 
                    templateName = ""                                                               '清空模板文件名称
                next 
            else
                action = " action actionType='nav' columnName='" & rss("columnname") & "'" 
                call makeWebHtml(action) 
                glb_filePath = replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
                if right(glb_filePath, 1) = "/" or glb_filePath = "" then
                    glb_filePath = glb_filePath & "index.html" 
                end if 
                s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                call echo(action, s) 
                if glb_filePath <> "" then
                    call createDirFolder(getFileAttr(glb_filePath, "1")) 
                    call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                end if 
                doevents() 
                templateName = "" 
            end if 
            conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & rss("id")) '更新导航为生成状态
        end if 
    rss.moveNext : wend : rss.close 
	'【@是jsp显示@】}catch(Exception e){} 

    '单独处理指定栏目对应文章
    if columnId <> "" then
        articleSql = "select * from " & db_PREFIX & "articledetail where parentid=" & columnId & " order by sortrank asc" 
    '批量处理文章
    elseIf addSql = "" then
        articleSql = "select * from " & db_PREFIX & "articledetail order by sortrank asc" 
    end if 
    if articleSql <> "" then
        '文章
        call echo("文章", "") 
		'【@是jsp显示@】try{
        rss.open articleSql, conn, 1, 1 
        while not rss.EOF
            glb_columnName = "" 
            action = " action actionType='detail' columnName='" & rss("parentid") & "' id='" & rss("id") & "'" 
            'call echo("action",action)
            call makeWebHtml(action) 
            glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
            if right(glb_filePath, 1) = "/" then
                glb_filePath = glb_filePath & "index.html" 
            end if 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            call echo(action, s) 
            '文件不为空  并且开启生成html
            if glb_filePath <> "" and rss("isonhtml") <>0 then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                conn.execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & rss("id")) '更新文章为生成状态
            end if 
            templateName = ""                                                               '清空模板文件名称
        rss.moveNext : wend : rss.close 
		'【@是jsp显示@】}catch(Exception e){} 
    end if 

    if addSql = "" then
        '单页
        call echo("单页", "") 
		'【@是jsp显示@】try{
        rss.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        while not rss.EOF
            glb_columnName = "" 
            action = " action actionType='onepage' id='" & rss("id") & "'" 
            'call echo("action",action)
            call makeWebHtml(action) 
            glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
            if right(glb_filePath, 1) = "/" then
                glb_filePath = glb_filePath & "index.html" 
            end if 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            call echo(action, s) 
            '文件不为空  并且开启生成html
            if glb_filePath <> "" and rss("isonhtml") <>0 then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                conn.execute("update " & db_PREFIX & "onepage set ishtml=true where id=" & rss("id")) '更新单页为生成状态
            end if 
            templateName = ""                                                               '清空模板文件名称
        rss.moveNext : wend : rss.close 
		'【@是jsp显示@】}catch(Exception e){} 

    end if 
end sub 

'复制html到网站
sub copyHtmlToWeb()
    dim webDir, toWebDir, toFilePath, filePath, fileName, fileList, splStr, content, s, s1, c, webImages, webCss, webJs, splJs 
    dim webFolderName, jsFileList, setFileCode, nErrLevel, jsFilePath, url,isSetFileNameType,isPinYin

    setFileCode = request("setcode")                                                '设置文件保存编码
	if setFileCode="" then
		setFileCode="gb2312"	'默认
	end if
	
	isSetFileNameType=request("isSetFileNameType")									'生成后的文件名称 是否把-转_
	isPinYin=request("isPinYin")													'文件名转拼音

    call handlePower("复制生成HTML页面")                                            '管理权限处理
    call writeSystemLog("", "复制生成HTML页面")                                     '系统日志

    webFolderName = cfg_webTemplate 
    if left(webFolderName, 1) = "/" then
        webFolderName = mid(webFolderName, 2) 
    end if 
    if right(webFolderName, 1) = "/" then
        webFolderName = mid(webFolderName, 1, len(webFolderName) - 1) 
    end if 
    if inStr(webFolderName, "/") > 0 then
        webFolderName = mid(webFolderName, inStr(webFolderName, "/") + 1) 
    end if 
    webDir = "/htmladmin/" & webFolderName & "/" 
    toWebDir = "/htmlw" & "eb/viewweb/" 
    call createDirFolder(toWebDir) 
    toWebDir = toWebDir & pinYin2(webFolderName) & "/" 

    call deleteFolder(toWebDir)                                                     '删除
    call createFolder("/htmlweb/web")                                               '创建文件夹 防止web文件夹不存在20160504
    call deleteFolder(webDir) 
    call createDirFolder(webDir) 
    webImages = webDir & "Images/" 
    webCss = webDir & "Css/" 
    webJs = webDir & "Js/" 
    call copyFolder(cfg_webImages, webImages) 
    call copyFolder(cfg_webCss, webCss) 
    call createFolder(webJs)                                                        '创建Js文件夹


    '处理Js文件夹
    splJs = split(getDirJsList(webJs), vbCrLf) 
    for each filePath in splJs
        if filePath <> "" then
            toFilePath = webJs & getFileName(filePath) 
            call echo("js", filePath) 
            call moveFile(filePath, toFilePath) 
        end if 
    next 
    '处理Css文件夹
    splStr = split(getDirCssList(webCss), vbCrLf) 
    for each filePath in splStr
        if filePath <> "" then
            content=readFile(filePath,templatesetcode)
			'content = getftext(filePath)
			 
            content = replace(content, cfg_webImages, "../Images") 

            content = deleteCssNote(content) 
            content = phptrim(content) 
            '设置为utf-8编码 20160527
            if lcase(setFileCode) = "utf-8" then
                content = replace(content, "gb2312", "utf-8") 
            end if 
            call writeToFile(filePath, content, setFileCode) 
            call echo("css", cfg_webImages) 
        end if 
    next 
    '复制栏目HTML
    isMakeHtml = true 
	'【@是jsp显示@】try{
    rss.open "select * from " & db_PREFIX & "webcolumn where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_filePath = replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 

        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("导航", glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
    '复制文章HTML
    rss.open "select * from " & db_PREFIX & "articledetail where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("文章" & rss("title"), glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
    '复制单面HTML
    rss.open "select * from " & db_PREFIX & "onepage where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/page/page" & rss("id")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("单页" & rss("title"), glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
	'【@是jsp显示@】}catch(Exception e){}
	
	 
    '批量处理html文件列表
    'call echo(cfg_webSiteUrl,cfg_webTemplate)
    'call rwend(fileList)
    dim sourceUrl, replaceUrl 
    splStr = split(fileList, vbCrLf) 
    for each filePath in splStr
        if filePath <> "" then
            filePath =  replace(filePath, "/", "_") 
			if isSetFileNameType="1" then
				filePath = replace(filePath, "-", "_") 
			end if
			if isPinYin="1" then
				filePath=pinYin2(filePath)			'注意
			end if
            filePath = webDir & filePath
            call echo("filePath", filePath) 
            content = getftext(filePath) 
            for each s in splStr
                s1 = s 
                if right(s1, 11) = "/index.html" then
                    s1 = left(s1, len(s1) - 11) & "/" 
                end if 
                sourceUrl = cfg_webSiteUrl & s1 
				s=replace(s, "/", "_") 
				if isSetFileNameType="1" then
					s = replace(s, "-", "_") 
				end if
				if isPinYin="1" then
					s=pinYin2(s)
				end if
                replaceUrl = cfg_webSiteUrl & s
				
                'Call echo(sourceUrl, replaceUrl)                             '屏蔽  否则大量显示20160613
                content = replace(content, sourceUrl, replaceUrl) 
            next 
            content = replace(content, cfg_webSiteUrl, "")                                  '删除网址
            content = replace(content, cfg_webTemplate & "/", "")                           '删除模板路径 记
			
            content = replace(replace(content, "\Inc/YZM_7.asp", "Images/YZM_7.jpg"), "/Inc/YZM_7.asp", "Images/YZM_7.jpg")  '替换验证码20160916
			content = replace(content, "/inc/yzm_7.asp", "Images/YZM_7.jpg")
			
			
            'content=nullLinkAddDefaultName(content)
            for each s in splJs
                if s <> "" then
                    fileName = getFileName(s) 
                    content = replace(content, "Images/" & fileName, "Js/" & fileName) 
                end if 
            next 
            if inStr(content, "/Jquery/Jquery.Min.js") > 0 then
                content = replace(content, "/Jquery/Jquery.Min.js", "Js/Jquery.Min.js") 
                call copyfile("/Jquery/Jquery.Min.js", webJs & "/Jquery.Min.js") 
            end if 
            content = replace(content, "<a href="""" ", "<a href=""index.html"" ")    '让首页加index.html

            call writeToFile(filePath, content,makeHtmlFileSetCode)  
			
        end if 
    next 

    '把复制网站夹下的images/文件夹下的js移到js/文件夹下  20160315
    dim htmlFileList, splHtmlFile, splJsFile, htmlFilePath, jsFileName 
    jsFileList = getDirJsNameList(webImages) 
    htmlFileList = getDirHtmlList(webDir) 
    splHtmlFile = split(htmlFileList, vbCrLf) 
    splJsFile = split(jsFileList, vbCrLf) 
    for each htmlFilePath in splHtmlFile
        content = getftext(htmlFilePath) 
        for each jsFileName in splJsFile
            content = regExp_Replace(content, "Images/" & jsFileName, "Js/" & jsFileName) 
        next 

        nErrLevel = 0 
        content = handleHtmlFormatting(content, false, nErrLevel, "|删除空行|")         '|删除空行|
        content = handleCloseHtml(content, true, "")                                    '闭合标签
        'nErrLevel = checkHtmlFormatting(content) 
        if checkHtmlFormatting(content) = false then
            call echored(htmlFilePath & "(格式化错误)", nErrLevel)                          '注意
        end if 
        '设置为utf-8编码
        if lcase(setFileCode) = "utf-8" then
            content = replace(content, "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />", "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />") 
        end if 
        content = phptrim(content) 
        call writeToFile(htmlFilePath, content, setFileCode) 
    next 
    'images下js移动到js下
    for each jsFileName in splJsFile
        jsFilePath = webImages & jsFileName 
        content = getftext(jsFilePath) 
        content = phptrim(content) 
        call writeToFile(webJs & jsFileName, content, setFileCode) 
        call deleteFile(jsFilePath) 
    next 

    call copyFolder(webDir, toWebDir) 
    '使htmlWeb文件夹用php压缩
    if request("isMakeZip") = "1" then
        call makeHtmlWebToZip(webDir) 
    end if 
    '使网站用xml打包20160612
    if request("isMakeXml") = "1" then
        call makeHtmlWebToXmlZip("/htmladmin/", webFolderName) 
    end if 
    '浏览地址
    url = "http://10.10.10.57/" & toWebDir 
    call echo("浏览", "<a href='" & url & "' target='_blank'>" & url & "</a>") 
    call echo("浏览", "<a href='" & toWebDir & "' target='_blank'>" & toWebDir & "</a>") 
end sub 
'使htmlWeb文件夹用php压缩
function makeHtmlWebToZip(webDir)
    dim content, splStr, filePath, c, arrayFile, fileName, fileType, isTrue 
    dim webFolderName 
    dim cleanFileList 
    splStr = split(webDir, "/") 
    webFolderName = splStr(2) 
    'content = getFileFolderList(webDir, true, "全部", "", "全部文件夹", "", "") 		'屏蔽这种
	content=getDirAllFileList(webDir,"")
    splStr = split(content, vbCrLf) 
    for each filePath in splStr
        if checkfolder(filePath) = false then
            arrayFile = handleFilePathArray(filePath) 
            fileName = LCase(arrayFile(2)) 
            fileType = LCase(arrayFile(4)) 
            fileName = remoteNumber(fileName) 
            isTrue = true 

            if inStr("|" & cleanFileList & "|", "|" & fileName & "|") > 0 and fileType = "html" then
                isTrue = false 
            end if 
            if isTrue = true then
                'call echo(fileType,fileName)
                if c <> "" then c = c & "|" 
                c = c & replace(filePath, handlePath("/"), "") 
                cleanFileList = cleanFileList & fileName & "|" 
            end if 
        end if 
    next 
    call rw(c) 
    c = c & "|||||" 
    call writeToFile("htmlweb/1.txt", c,makeHtmlFileSetCode) 
    call echo("<hr>cccccccccccc", c) 
    '先判断这个文件存在20160309
    if checkFile("/myZIP.php") = true then
        call echo("", XMLPost(getHost() & "/myZIP.php?webFolderName=" & webFolderName, "content=" & escape(c))) 
    end if 

end function 
'使网站用xml打包20160612
function makeHtmlWebToXmlZip(sNewWebDir, rootDir)
    dim xmlFileName, xmlSize 
    xmlFileName = setFileName(getIP()) & "_update.xml"                              '获得ip有可能为空:: 创建时会有问题

    'sNewWebDir="\Templates2015\"
    'rootDir="\sharembweb\"

    dim objXmlZIP : set objXmlZIP = new xmlZIP
        call objXmlZIP.callRun(handlePath(sNewWebDir), handlePath(sNewWebDir & rootDir), false, xmlFileName) 
        call echo(handlePath(sNewWebDir), handlePath(sNewWebDir & rootDir)) 
    set objXmlZIP = nothing 
    doevents 
    xmlSize = getFSize(xmlFileName) 
    xmlSize = printSpaceValue(xmlSize) 
    call echo("下载xml打包文件", "<a href=/tools/downfile.asp?act=download&downfile=" & xorEnc("/" & xmlFileName, 31380) & " title='点击下载'>点击下载" & xmlFileName & "(" & xmlSize & ")</a>") 
end function 


'生成更新sitemap.xml 20160118
sub saveSiteMap()
    dim isWebRunHtml                                                                '是否为html方式显示网站
    dim changefreg                                                                  '更新频率
    dim priority                                                                    '优先级
    dim s, c, url,sql 
    call handlePower("修改生成SiteMap")                                             '管理权限处理

    changefreg = request("changefreg") 
    priority = request("priority") 
    call loadWebConfig()                                                            '加载配置
    'call eerr("cfg_flags",cfg_flags)
    if inStr(cfg_flags, "|htmlrun|") > 0 then
        isWebRunHtml = true 
    else
        isWebRunHtml = false 
    end if 

    c = c & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf 
    c = c & vbTab & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbCrLf 
    dim rsx : set rsx = createObject("Adodb.RecordSet")
        '栏目
		'【@是jsp显示@】try{
        rsx.open "select * from " & db_PREFIX & "webcolumn where isonhtml<>0 order by sortrank asc", conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = escape("?act=nav&columnName=" & rsx("columnname")) 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("栏目2", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 
		'【@是jsp显示@】}catch(Exception e){} 

        '文章
		sql="select * from " & db_PREFIX & "articledetail  where isonhtml<>0 order by sortrank asc"
		'【@是jsp显示@】try{
        rsx.open sql, conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = "?act=detail&id=" & rsx("id") 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("文章", "<a href=""" & url & """>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 
		'【@是jsp显示@】}catch(Exception e){} 

        '单页
		'【@是jsp显示@】try{
        rsx.open "select * from " & db_PREFIX & "onepage where isonhtml<>0 order by sortrank asc", conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = "?act=onepage&id=" & rsx("id") 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("单页", "<a href=""" & url & """>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 
		'【@是jsp显示@】}catch(Exception e){} 
        c = c & vbTab & "</urlset>" & vbCrLf 
        call loadWebConfig() 
        call createFile("sitemap.xml", c) 
        call echo("生成sitemap.xml文件成功", "<a href='/sitemap.xml' target='_blank'>点击预览sitemap.xml</a>") 


        '判断是否生成sitemap.html
        if request("issitemaphtml") = "1" then
            c = "" 
            '第二种
            '栏目
			'【@是jsp显示@】try{
            rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
            while not rsx.EOF
                if rsx("nofollow") = 0 then
                    if isWebRunHtml = true then
                        url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
                        url = handleAction(url) 
                    else
                        url = escape("?act=nav&columnName=" & rsx("columnname")) 
                    end if 
                    url = urlAddHttpUrl(cfg_webSiteUrl, url) 

                    '判断是否生成html
                    if rsx("isonhtml") <>0 then
                        s = "<a href=""" & url & """>" & rsx("columnname") & "</a>" 
                    else
                        s = "<span>" & rsx("columnname") & "</span>" 
                    end if 
                    c = c & "<li style=""width:20%;"">" & s & vbCrLf & "<ul>" & vbCrLf 

                    '文章
					sql="select * from " & db_PREFIX & "articledetail where parentId=" & rsx("id") & " order by sortrank asc"
                    rss.open sql, conn, 1, 1 
                    while not rss.EOF
                        if rss("nofollow") = 0 then
                            if isWebRunHtml = true then
                                url = getRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
                                url = handleAction(url) 
                            else
                                url = "?act=detail&id=" & rss("id") 
                            end if 
                            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                            '判断是否生成html
                            if rss("isonhtml") <>0 then
                                s = "<a href=""" & url & """>" & rss("title") & "</a>" 
                            else
                                s = "<span>" & rss("title") & "</span>" 
                            end if 
                            c = c & "<li style=""width:20%;"">" & s & "</li>" & vbCrLf 
                        end if 
                    rss.moveNext : wend : rss.close 
                    c = c & "</ul>" & vbCrLf & "</li>" & vbCrLf 


                end if 
            rsx.moveNext : wend : rsx.close 
			'【@是jsp显示@】}catch(Exception e){} 

            '单面
            c = c & "<li style=""width:20%;""><a href=""javascript:;"">单面列表</a>" & vbCrLf & "<ul>" & vbCrLf 
			'【@是jsp显示@】try{
            rsx.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
            while not rsx.EOF
                if rsx("nofollow") = 0 then
                    c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                    if isWebRunHtml = true then
                        url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
                        url = handleAction(url) 
                    else
                        url = "?act=onepage&id=" & rsx("id") 
                    end if 
                    '判断是否生成html
                    if rsx("isonhtml") <>0 then
                        s = "<a href=""" & url & """>" & rsx("title") & "</a>" 
                    else
                        s = "<span>" & rsx("title") & "</span>" 
                    end if 

                    c = c & "<li style=""width:20%;"">" & s & "</li>" & vbCrLf                'target=""_blank""  去掉
                end if 
            rsx.moveNext : wend : rsx.close 
			'【@是jsp显示@】}catch(Exception e){} 

            c = c & "</ul>" & vbCrLf & "</li>" & vbCrLf 

            dim templateContent 
            templateContent = getftext(adminDir & "/template_SiteMap.html") 


            templateContent = replace(templateContent, "{$content$}", c) 
            templateContent = replace(templateContent, "{$Web_Title$}", cfg_webTitle) 


            call createFile("sitemap.html", templateContent) 
            call echo("生成sitemap.html文件成功", "<a href='/sitemap.html' target='_blank'>点击预览sitemap.html</a>") 
        end if 
        call writeSystemLog("", "保存sitemap.xml")                                      '系统日志
end sub
 
'获得会员验证后模板 20170422
function getMemberVerificationTemplateName(templateName) 
	templateName=getCheckMemberLoginTemplate(templateName)
	getMemberVerificationTemplateName=templateName
end function
%>   