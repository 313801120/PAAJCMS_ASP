<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-02-27
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'dim nCount,nPageSize,maxpage,page,x,i,PageControl
function pageControl(nCount, nPageCount, page, table_style, font_style)
    '生成上一页下一页链接
    dim query, a, x, temp, action, i, c 
    action = "http://" & request.serverVariables("HTTP_HOST") & request.serverVariables("SCRIPT_NAME") 
    query = split(request.serverVariables("QUERY_STRING"), "&") 
    for each x in query
        a = split(x, "=") 
        if strComp(a(0), "page", vbTextCompare) <> 0 then 
            temp = temp & a(0) & "=" & a(1) & "&" 
        end if 
    next 

    c = c & "<table width=100% border=0 cellpadding=0 cellspacing=0 class=hui>" & vbCrLf 
    c = c & "<TD align=center height=35 class='font'>" & vbCrLf 
    c = c & font_style & vbCrLf 
    if page <= 1 then
        c = c & "首 页 " & vbCrLf 
        c = c & "上一页 " & vbCrLf 
    else
        c = c & "<A HREF=" & action & "?" & temp & "Page=1 class='huilianjie'>首 页</A> " & vbCrLf 
        c = c & "<A HREF=" & action & "?" & temp & "Page=" &(page - 1) & " class='huilianjie'>上一页</A> " & vbCrLf 
    end if 
    if page >= nPageCount then
        c = c & "下一页 " & vbCrLf 
        c = c & "尾 页 " & vbCrLf 
    else
        c = c & "<A HREF=" & action & "?" & temp & "Page=" &(page + 1) & " class='huilianjie'>下一页</A> " & vbCrLf 
        c = c & "<A HREF=" & action & "?" & temp & "Page=" & nPageCount & " class='huilianjie'>尾 页</A> " & vbCrLf 
    end if 
    c = c & " 页次：<b><font color=red>" & page & "</font>/" & nPageCount & "</b>页" & vbCrLf 
    c = c & " 共有<b>" & nCount & "</b>个记录" & vbCrLf 
    if nPageCount < 10 then
        c = c & "<SELECT onchange='location.href(this.options[this.selectedIndex].value)'>" 
        for i = 1 to nPageCount
            if page = i then
                c = c & "<option value=?" & temp & "page=" & i & " selected> 第" & i & "页</option>" 
            else
                c = c & "<option value=?" & temp & "page=" & i & "> 第" & i & "页</option>" 
            end if 
        next 
        c = c & "</select>" 
    else
        temp = replace(temp, "\", "\\") 
        c = c & "<input name=""GotoPage"" value=""" & page & """ type=""text"" id=""GotoPage"" size=""5"" style=""text-align:center;"" /> &nbsp;" 
        c = c & "<input type=""button"" name=""button"" id=""button"" value=""跳转"" onclick=""location.href('?" & temp & "Page=' + document.all.GotoPage.value)"" />" 
    end if 
    c = c & "<!--用时--></TD>" & vbCrLf 
    c = c & "</TR>" & vbCrLf 
    c = c & "</table>" & vbCrLf 
    pageControl = c 
end function 

'页控制  记录总数  每页显示数  当前面 (2015117)   webPageControl(59,12,1,"http://www.baidu.com")
function webPageControl(nRecrodCount, nPageSize, sPage, configPageUrl, action)
    dim c, nCountPage, i, nDisplay, nDispalyOK, sTemp, cPages, url, selStr,nPage,s
    dim nPreviousPage, nNextPage                                                      '定义上一页，下一页
    dim isDisplayTip                                                                '是否显示提示翻页信息
    isDisplayTip = true 
	if sPage="0" then
		sPage=1
	end if
	'call echo(typename(sPage),sPage)
	
    dim sPageStart, sPageEnd, sHomePage, sHomePageFocus, sUpPage, sUpPageFocus, sNextPage, sNextPageFocus, sForPage, sForPageFocus, sTailPage, sTailPageFocus 
    if action <> "" then
        sPageStart = getStrCut(action, "[sPageStart]", "[/sPageStart]", 2)    '页头部分
        sPageEnd = getStrCut(action, "[sPageEnd]", "[/sPageEnd]", 2)          '页尾部分
        sHomePage = getStrCut(action, "[sHomePage]", "[/sHomePage]", 2)       '首页
        sHomePageFocus = getStrCut(action, "[sHomePageFocus]", "[/sHomePageFocus]", 2) '首页交点
        sUpPage = getStrCut(action, "[sUpPage]", "[/sUpPage]", 2)             '上一页
        sUpPageFocus = getStrCut(action, "[sUpPageFocus]", "[/sUpPageFocus]", 2) '上一页交点
        sNextPage = getStrCut(action, "[sNextPage]", "[/sNextPage]", 2)       '下一页
        sNextPageFocus = getStrCut(action, "[sNextPageFocus]", "[/sNextPageFocus]", 2) '下一页交点
        sForPage = getStrCut(action, "[sForPage]", "[/sForPage]", 2)          '循环页
        sForPageFocus = getStrCut(action, "[sForPageFocus]", "[/sForPageFocus]", 2) '循环页交点
        sTailPage = getStrCut(action, "[sTailPage]", "[/sTailPage]", 2)       '最后页
        sTailPageFocus = getStrCut(action, "[sTailPageFocus]", "[/sTailPageFocus]", 2) '最后页交点


    end if 
    '页头部分
    if sPageStart = "" then
        sPageStart = "<ul class=""pagecontrolwrap"">" & vbCrLf & "<li class=""pageinfo"">共[$nRecrodCount$]条 [$nPage$]/[$nCountPage$]页</li>" & vbCrLf 
    end if 
    '页尾部分
    if sPageEnd = "" then
        sPageEnd = "</ul><div class=""clear""></div>" & vbCrLf 
    end if 
    '首页
    if sHomePage = "" then
        sHomePage = "<li><a href=""[$url$]"">首页</a></li>" & vbCrLf 
    end if 
    '首页交点
    if sHomePageFocus = "" then
        sHomePageFocus = "<li class=""pageli"">首页</li>" & vbCrLf 
    end if 
    '上一页
    if sUpPage = "" then
        sUpPage = "<li><a href=""[$url$]"">上一页</a></li>" & vbCrLf 
    end if 
    '上一页交点
    if sUpPageFocus = "" then
        sUpPageFocus = "<li class=""pageli"">上一页</li>" & vbCrLf 
    end if 
    '下一页
    if sNextPage = "" then
        sNextPage = "<li><a href=""[$url$]"">下一页</a></li>" & vbCrLf 
    end if 
    '下一页交点
    if sNextPageFocus = "" then
        sNextPageFocus = "<li class=""pageli"">下一页</li>" & vbCrLf 
    end if 
    '循环页
    if sForPage = "" then
        sForPage = "<li><a href=""[$url$]"">[$i$]</a></li>" & vbCrLf 
    end if 
    '循环页交点
    if sForPageFocus = "" then
        sForPageFocus = "<li class=""pagefocus"">[$i$]</li>" & vbCrLf 
    end if 
    '最后页
    if sTailPage = "" then
        sTailPage = "<li><a href=""[$url$]"">末页</a></li>" & vbCrLf 
    end if 
    '最后页交点
    if sTailPageFocus = "" then
        sTailPageFocus = "<li class=""pageli"">末页</li>" & vbCrLf 
    end if 
    '测试时用到20160630
    if 1 = 2 then
        c = "[sPageStart]" & vbCrLf & sPageStart & "[/sPageStart]" & vbCrLf & vbCrLf 
        c = c & "[sHomePage]" & vbCrLf & sHomePage & "[/sHomePage]" & vbCrLf & vbCrLf 
        c = c & "[sHomePageFocus]" & vbCrLf & sHomePageFocus & "[/sHomePageFocus]" & vbCrLf & vbCrLf 

        c = c & "[sUpPage]" & vbCrLf & sUpPage & "[/sUpPage]" & vbCrLf & vbCrLf 
        c = c & "[sUpPageFocus]" & vbCrLf & sUpPageFocus & vbCrLf & "[/sUpPageFocus]" & vbCrLf 


        c = c & "[sForPage]" & vbCrLf & sForPage & "[/sForPage]" & vbCrLf & vbCrLf 
        c = c & "[sForPageFocus]" & vbCrLf & sForPageFocus & "[/sForPageFocus]" & vbCrLf & vbCrLf 


        c = c & "[sNextPage]" & vbCrLf & sNextPage & "[/sNextPage]" & vbCrLf & vbCrLf 
        c = c & "[sNextPageFocus]" & vbCrLf & sNextPageFocus & "[/sNextPageFocus]" & vbCrLf & vbCrLf 

        c = c & "[sTailPage]" & vbCrLf & sTailPage & "[/sTailPage]" & vbCrLf & vbCrLf 
        c = c & "[sTailPageFocus]" & vbCrLf & sTailPageFocus & "[/sTailPageFocus]" & vbCrLf & vbCrLf 
        c = c & "[sPageEnd]" & vbCrLf & sPageEnd & "[/sPageEnd]" & vbCrLf 
        call rwend("[page]" & vbCrLf & vbCrLf & c & vbCrLf & "[/page]") 
    end if 
    '配置页为空则
    if configPageUrl = "" then
        configPageUrl = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 

    nDisplay = 6                                                                    '显示数
    nDispalyOK = 0                                                                  '显示成功数
    s = handleNumberType(sPage) 
    if s = "" then
        nPage = 1 
    else
        nPage = CInt(s) 
    end if 
    '获得总页数
    nCountPage = getCountPage(nRecrodCount, nPageSize) 



    nPreviousPage = nPage - 1 
    nNextPage = nPage + 1 

    '处理上一页
    if nPreviousPage <= 0 then
        nPreviousPage = -1
    end if 
    '处理下一页
    if nNextPage > nCountPage then
        nNextPage = -1
    end if 

    '页开始
    c = sPageStart 
    '首页  
    if nPage > 1 then
        c = c & replace(sHomePage, "[$url$]", replace(configPageUrl, "[id]", "")) 
    elseIf isDisplayTip = true then
        c = c & sHomePageFocus 
    end if 
    '上一页
    if nPreviousPage <> -1 then
        sTemp = cstr(nPreviousPage) 
        if nPreviousPage <= 1 then
            sTemp = "" 
        end if 
        c = c & replace(sUpPage, "[$url$]", replace(configPageUrl, "[id]", sTemp)) 
    elseIf isDisplayTip = true then
        c = c & sUpPageFocus 
    end if 


    dim n 
    if nPage=0 then
		nPage=1
	end if
    n =(nPage - 3) 
    'call echo("n=" & n, "nPage=" & nPage)
	
    '翻页循环
    for i = n to nCountPage
        if i >= 1 then
            nDispalyOK = nDispalyOK + 1 
            'call echo(i,nPage)
            if i = nPage then
                c = c & replace(sForPageFocus, "[$i$]", i) 
            else
                sTemp = cstr(i) 
                if i <= 1 then
                    sTemp = "" 
                end if 
                c = c & replace(replace(sForPage, "[$url$]", replace(configPageUrl, "[id]", sTemp)), "[$i$]", i) 
            end if 
            if nDispalyOK > nDisplay then
                exit for 
            end if 
        end if 
    next 
    '下一页
    if nCountPage > nPage then
        c = c & replace(sNextPage, "[$url$]", replace(configPageUrl, "[id]", nNextPage)) 
    elseIf isDisplayTip = true then
        c = c & sNextPageFocus 
    end if 
    '末页
    if nCountPage > nPage then
        c = c & replace(sTailPage, "[$url$]", replace(configPageUrl, "[id]", nCountPage)) 
    elseIf isDisplayTip = true then
        c = c & sTailPageFocus 
    end if 

    c = c & sPageEnd 


    c = replaceValueParam(c, "nRecrodCount", nRecrodCount) 
    c = replaceValueParam(c, "nPage", nPage) 
    if nCountPage = 0 then
        nCountPage = 1 
    end if 
    c = replaceValueParam(c, "nCountPage", nCountPage) 

    if inStr(c, "[$page-select-openlist$]") > 0 then
        for i = 1 to nCountPage
            url = replace(configPageUrl, "[id]", i) 
            selStr = "" 
            if i = nPage then
                selStr = " selected" 
            end if 
            cPages = cPages & "<option value=""" & url & """" & selStr & ">" & i & "</option>" & vbCrLf 
        next 
        c = replace(c, "[$page-select-openlist$]", cPages) 
    end if 

    webPageControl = c & vbCrLf 
end function 


'获得Rs页数
function getRsPageNumber(rs, nCount, nPageSize, page)
    dim maxPage 
    dim x 
    x = 0 
    rs.pageSize = nPageSize                                                         '每页记录条数 可自定义
    maxPage = rs.pageCount 
    if maxPage > 0 then
        if not isNumeric(page) or page = "" then
            page = 1 
        else
            page = CInt(page) 
        end if 
        if page < 1 then
            page = 1 
        elseIf page > maxPage then
            page = maxPage 
        end if 
        rs.absolutePage = page 
        if page = maxPage then
            x = nCount - (maxPage - 1) * nPageSize 
        else
            x = nPageSize 
        end if 
    end if 
    getRsPageNumber = x 
end function 
%>   
