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
'Form

'�����ύ
sub formSubmit()
    dim tableName, splStr, s, fieldName, fieldContent, fieldList, YZM  
    call OpenConn() 
    splStr = split(request.form(), "&") 
    tableName = rf("TableName") 
    YZM = trim(rf("YZM")) 

    if YZM <> "" then
        if session("YZM") <> YZM then
            call Javascript("����", "��֤�벻��ȷ", "") 
            response.end() 
        end if 
    end if 

    fieldList = lCase(getFieldList(tableName)) 
    'Call Echo("FieldList",FieldList)
    'Call Echo("TableName", TableName)
    rs.open "Select * From [" & tableName & "]", conn, 1, 3 
    rs.addNew 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
        'FieldContent = Mid(S,InStr(S,"=")+1)
        fieldContent = rf(fieldName) 
        if inStr("," & fieldList & ",", "," & fieldName & ",") > 0 then
            rs(fieldName) = fieldContent '��@����asp����@��  
        end if 
    'Call Echo(FieldName,FieldContent & "," & unescape(FieldContent)) 
    next 
    rs.update : rs.close 
    'Call Echo("DialogTitle",Rf("DialogTitle"))
    'Call Die("��������")

    call Javascript("����", "�ύ" & rf("DialogTitle") & "�ɹ�", "") 
end sub 

'���POST�ֶ������б� 20160226
function getFormFieldList()
    dim s, c, splStr, fieldName 
    splStr = split(request.form(), "&") 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
        if c <> "" then c = c & "|" 
        c = c & fieldName 
    next 
    getFormFieldList = c 
end function 

%> 










