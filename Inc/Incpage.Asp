<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-04
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'dim nCount,nPageSize,maxpage,page,x,i,PageControl
function pageControl(nCount, nPageCount, page, table_style, font_style)
    '������һҳ��һҳ����
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
        c = c & "�� ҳ " & vbCrLf 
        c = c & "��һҳ " & vbCrLf 
    else
        c = c & "<A HREF=" & action & "?" & temp & "Page=1 class='huilianjie'>�� ҳ</A> " & vbCrLf 
        c = c & "<A HREF=" & action & "?" & temp & "Page=" &(page - 1) & " class='huilianjie'>��һҳ</A> " & vbCrLf 
    end if 
    if page >= nPageCount then
        c = c & "��һҳ " & vbCrLf 
        c = c & "β ҳ " & vbCrLf 
    else
        c = c & "<A HREF=" & action & "?" & temp & "Page=" &(page + 1) & " class='huilianjie'>��һҳ</A> " & vbCrLf 
        c = c & "<A HREF=" & action & "?" & temp & "Page=" & nPageCount & " class='huilianjie'>β ҳ</A> " & vbCrLf 
    end if 
    c = c & " ҳ�Σ�<b><font color=red>" & page & "</font>/" & nPageCount & "</b>ҳ" & vbCrLf 
    c = c & " ����<b>" & nCount & "</b>����¼" & vbCrLf 
    if nPageCount < 10 then
        c = c & "<SELECT onchange='location.href(this.options[this.selectedIndex].value)'>" 
        for i = 1 to nPageCount
            if page = i then
                c = c & "<option value=?" & temp & "page=" & i & " selected> ��" & i & "ҳ</option>" 
            else
                c = c & "<option value=?" & temp & "page=" & i & "> ��" & i & "ҳ</option>" 
            end if 
        next 
        c = c & "</select>" 
    else
        temp = replace(temp, "\", "\\") 
        c = c & "<input name=""GotoPage"" value=""" & page & """ type=""text"" id=""GotoPage"" size=""5"" style=""text-align:center;"" /> &nbsp;" 
        c = c & "<input type=""button"" name=""button"" id=""button"" value=""��ת"" onclick=""location.href('?" & temp & "Page=' + document.all.GotoPage.value)"" />" 
    end if 
    c = c & "<!--��ʱ--></TD>" & vbCrLf 
    c = c & "</TR>" & vbCrLf 
    c = c & "</table>" & vbCrLf 
    pageControl = c 
end function 

'ҳ����  ��¼����  ÿҳ��ʾ��  ��ǰ�� (2015117)   webPageControl(59,12,1,"http://www.baidu.com")
function webPageControl(nRecrodCount, nPageSize, sPage, configPageUrl, action)
    dim c, nCountPage, i, nDisplay, nDispalyOK, sTemp, cPages, url, selStr,nPage,s
    dim nPreviousPage, nNextPage                                                      '������һҳ����һҳ
    dim isDisplayTip                                                                '�Ƿ���ʾ��ʾ��ҳ��Ϣ
    isDisplayTip = true 

    dim sPageStart, sPageEnd, sHomePage, sHomePageFocus, sUpPage, sUpPageFocus, sNextPage, sNextPageFocus, sForPage, sForPageFocus, sTailPage, sTailPageFocus 
    if action <> "" then
        sPageStart = getStrCut(action, "[sPageStart]", "[/sPageStart]", 2)    'ҳͷ����
        sPageEnd = getStrCut(action, "[sPageEnd]", "[/sPageEnd]", 2)          'ҳβ����
        sHomePage = getStrCut(action, "[sHomePage]", "[/sHomePage]", 2)       '��ҳ
        sHomePageFocus = getStrCut(action, "[sHomePageFocus]", "[/sHomePageFocus]", 2) '��ҳ����
        sUpPage = getStrCut(action, "[sUpPage]", "[/sUpPage]", 2)             '��һҳ
        sUpPageFocus = getStrCut(action, "[sUpPageFocus]", "[/sUpPageFocus]", 2) '��һҳ����
        sNextPage = getStrCut(action, "[sNextPage]", "[/sNextPage]", 2)       '��һҳ
        sNextPageFocus = getStrCut(action, "[sNextPageFocus]", "[/sNextPageFocus]", 2) '��һҳ����
        sForPage = getStrCut(action, "[sForPage]", "[/sForPage]", 2)          'ѭ��ҳ
        sForPageFocus = getStrCut(action, "[sForPageFocus]", "[/sForPageFocus]", 2) 'ѭ��ҳ����
        sTailPage = getStrCut(action, "[sTailPage]", "[/sTailPage]", 2)       '���ҳ
        sTailPageFocus = getStrCut(action, "[sTailPageFocus]", "[/sTailPageFocus]", 2) '���ҳ����


    end if 
    'ҳͷ����
    if sPageStart = "" then
        sPageStart = "<ul class=""pagecontrolwrap"">" & vbCrLf & "<li class=""pageinfo"">��[$nRecrodCount$]�� [$nPage$]/[$nCountPage$]ҳ</li>" & vbCrLf 
    end if 
    'ҳβ����
    if sPageEnd = "" then
        sPageEnd = "</ul><div class=""clear""></div>" & vbCrLf 
    end if 
    '��ҳ
    if sHomePage = "" then
        sHomePage = "<li><a href=""[$url$]"">��ҳ</a></li>" & vbCrLf 
    end if 
    '��ҳ����
    if sHomePageFocus = "" then
        sHomePageFocus = "<li class=""pageli"">��ҳ</li>" & vbCrLf 
    end if 
    '��һҳ
    if sUpPage = "" then
        sUpPage = "<li><a href=""[$url$]"">��һҳ</a></li>" & vbCrLf 
    end if 
    '��һҳ����
    if sUpPageFocus = "" then
        sUpPageFocus = "<li class=""pageli"">��һҳ</li>" & vbCrLf 
    end if 
    '��һҳ
    if sNextPage = "" then
        sNextPage = "<li><a href=""[$url$]"">��һҳ</a></li>" & vbCrLf 
    end if 
    '��һҳ����
    if sNextPageFocus = "" then
        sNextPageFocus = "<li class=""pageli"">��һҳ</li>" & vbCrLf 
    end if 
    'ѭ��ҳ
    if sForPage = "" then
        sForPage = "<li><a href=""[$url$]"">[$i$]</a></li>" & vbCrLf 
    end if 
    'ѭ��ҳ����
    if sForPageFocus = "" then
        sForPageFocus = "<li class=""pagefocus"">[$i$]</li>" & vbCrLf 
    end if 
    '���ҳ
    if sTailPage = "" then
        sTailPage = "<li><a href=""[$url$]"">ĩҳ</a></li>" & vbCrLf 
    end if 
    '���ҳ����
    if sTailPageFocus = "" then
        sTailPageFocus = "<li class=""pageli"">ĩҳ</li>" & vbCrLf 
    end if 
    '����ʱ�õ�20160630
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
    '����ҳΪ����
    if configPageUrl = "" then
        configPageUrl = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 

    nDisplay = 6                                                                    '��ʾ��
    nDispalyOK = 0                                                                  '��ʾ�ɹ���
    s = handleNumberType(sPage) 
    if s = "" then
        nPage = 1 
    else
        nPage = CInt(s) 
    end if 
    '�����ҳ��
    nCountPage = getCountPage(nRecrodCount, nPageSize) 



    nPreviousPage = nPage - 1 
    nNextPage = nPage + 1 

    '������һҳ
    if nPreviousPage <= 0 then
        nPreviousPage = -1
    end if 
    '������һҳ
    if nNextPage > nCountPage then
        nNextPage = -1
    end if 

    'ҳ��ʼ
    c = sPageStart 
    '��ҳ
    if nPage > 1 then
        c = c & replace(sHomePage, "[$url$]", replace(configPageUrl, "[id]", "")) 
    elseIf isDisplayTip = true then
        c = c & sHomePageFocus 
    end if 
    '��һҳ
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
	
    '��ҳѭ��
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
    '��һҳ
    if nCountPage > nPage then
        c = c & replace(sNextPage, "[$url$]", replace(configPageUrl, "[id]", nNextPage)) 
    elseIf isDisplayTip = true then
        c = c & sNextPageFocus 
    end if 
    'ĩҳ
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


'���Rsҳ��
function getRsPageNumber(rs, nCount, nPageSize, page)
    dim maxPage 
    dim x 
    x = 0 
    rs.pageSize = nPageSize                                                         'ÿҳ��¼���� ���Զ���
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
