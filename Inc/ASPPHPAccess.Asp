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
'ASP PHP���ݲ���ͨ���ļ�

'�ж�׷��Sql�Ǽ�Where ����And   sql = getWhereAnd(sql,addSql)        �޸���20141007 ��ǿ��
function getWhereAnd(byVal sql, byVal addSql)
    dim LCaseAddSql, addType, s 
    '׷��SQlΪ�����˳�
    if trim(addSql) = "" then getWhereAnd = sql : exit function 
    if inStr(LCase(sql), " where ") > 0 then
        addType = " And " 
    else
        addType = " Where " 
    end if 
    if addSql <> "" then
        addSql = trim(addSql) 
        LCaseAddSql = LCase(addSql) 
        if left(LCaseAddSql, 6) = "order " or left(LCaseAddSql, 6) = "group " then
            getWhereAnd = sql & " " & addSql : exit function '�Ľ�����ӿո���Ϊǰ���Ѿ�ɾ����20160115
        elseIf left(LCaseAddSql, 6) = "where " then
            addSql = mid(addSql, 7) 
        elseIf left(LCaseAddSql, 4) = "and " then
            addSql = mid(addSql, 5) 
        end if 
        '��where �Ľ�   20160623
        s = lcase(addSql) 
        if s <> "and" and s <> "or" and s <> "where" then
            sql = sql & addType & addSql 
        end if 
    end if 
    getWhereAnd = sql 
end function 
'�����ѯ Or �� And        �����޸���20140703
function orAndSearch(addSql, seectField, searchValue)
    dim splStr, s, c 
    searchValue = regExp_Replace(searchValue, " or ", " Or ") 
    searchValue = regExp_Replace(searchValue, " and ", " And ") 
    if inStr(searchValue, " Or ") > 0 then
        splStr = split(searchValue, " Or ") 
        for each s in splStr
            if s <> "" then
                if c <> "" then c = c & " Or " 
                c = c & " " & seectField & " Like '%" & s & "%'" 
            end if 
        next 
    elseIf inStr(searchValue, " And ") > 0 then
        splStr = split(searchValue, " And ") 
        for each s in splStr
            if s <> "" then
                if c <> "" then c = c & " And " 
                c = c & " " & seectField & " Like '%" & s & "%'" 
            end if 
        next 
    elseIf searchValue <> "" then
        splStr = split(searchValue, " And ") 
        for each s in splStr
            if s <> "" then
                if c <> "" then c = c & " And " 
                c = c & " " & seectField & " Like '%" & s & "%'" 
            end if 
        next 
    end if 
    if c <> "" then
        if inStr(LCase(addSql), " where ") = 0 then
            c = " Where " & c 
        else
            c = " And " & c 
        end if 
        addSql = addSql & c 
    end if 
    orAndSearch = addSql 
end function 



'��õ�ǰid��ǰҳ�� Ĭ��ÿҳ��ʾ10�� 20160716
function getThisIdPage(tableName, id, nPageSize)
    dim nCount 
    '��php����������Լ������
    getThisIdPage = "1" 
    if id = "" then
        exit function 
    end if
'    if nPageSize = "" then        nPageSize = 10     end if 		'Ϊ���Ͻ�����Ҫ������Ժ�Ҫ��.netc������ͨ�������Բ�����
    nCount = nGetRecordCount( tableName ," where id<=" & id) 
    getThisIdPage = cstr(getCountPage(cint(nCount), nPageSize)) 
'call echo("tableName=" & tableName & "id=" & id &",ncount=" & ncount,npagesize & "               ," & getThisIdPage)
end function
'����ܼ�¼
function nGetRecordCount(tableName,addsql)
	nGetRecordCount=0
	rsx.open"select * from " & tableName & " " & addsql,conn,1,1
	if not rsx.eof then
		nGetRecordCount=cint(rsx.recordcount)
	end if:rsx.close
end function


'����SqlServer�����﷨(AccessתSqlServer)
'SqlServer���Ĵ���Ҫ��Ƚϸߣ��ո�ֻ��Ϊ''��������""���о�����ֵ���ַ���������
function handleSqlServer(byVal content)
    dim i, s, tempS, tempS2, c, splStr 
    splStr = split(content, ",") 
    for each content in splStr

        content = replace(content, """""", "''") 
        content = replace(content, "''''", "''") 
        content = replace(content, "Date()", "GetDate()") 
        content = replace(content, "Now()", "GetDate()") 
        'Content = Replace(Content,"YesNo Default Yes"," Int Default 1")
        'Content = Replace(Content,"YesNo Default No"," Int Default 0")
        'Content = Replace(Content,"YesNo,"""," Int Default 1,""")

        content = replace(content, "YesNo Default Yes", " Bit Default 'true'") 
        content = replace(content, "YesNo Default No", " Bit Default 'false'") 
        content = replace(content, " YesNo", "  Bit Default 'true'") 
        if inStr(content, "YesNo") > 0 then
            call eerr("Content", content) 
        end if 
        for i = 1 to len(content)
            s = mid(content, i, 1) 
            if s = " " then
                tempS = LCase(mid(content, i + 1, 7)) 
                tempS2 = LCase(mid(content, i + 1, 8)) 
                if tempS = "varchar" and tempS2 <> "varchar(" then
                    c = c & " VarChar(255)" 
                    i = i + 7 
                else
                    c = c & s 
                end if 
            else
                c = c & s 
            end if 
        next 
        c = c & "," 
    next 
    if c <> "" then c = left(c, len(c) - 1) 
    handleSqlServer = c 
end function 
%> 
