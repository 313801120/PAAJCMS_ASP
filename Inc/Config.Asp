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
<% option explicit%> 
<!--#Include file = "Common.Asp"-->  
<!--#Include file = "Function.Asp"-->  
<!--#Include file = "FSO.Asp"-->  
<!--#Include file = "AdodbStream.Asp"-->  
<!--#Include file = "XML.Asp"-->  
<!--#Include file = "Time.Asp"-->  
<!--#Include file = "RegExp.Asp"-->  
<!--#Include file = "Cai.Asp"-->  
<!--#Include file = "StringNumber.Asp"-->  
<!--#Include file = "EncDec.Asp"-->  
<!--#Include file = "Dictionary.Asp"-->  
<!--#Include file = "Access.Asp"-->  
<!--#Include file = "Incpage.Asp"-->  
<!--#Include file = "Check.Asp"-->  
<!--#Include file = "GetWord.Asp"-->  
<!--#Include file = "Url.Asp"-->  
<!--#Include file = "Conn.Asp"-->  
<!--#Include file = "Fun_Table.Asp"-->  
<!--#Include file = "FunHTML.Asp"-->   
<!--#Include file = "Html.Asp"-->   
<!--#Include file = "Img.Asp"-->  
<!--#Include file = "PinYin.Asp"--> 
<!--#Include file = "Md5.Asp"-->   
<!--#Include file = "IE.Asp"-->   
<!--#Include file = "ASPPHPAccess.Asp"-->     
<!--#Include file = "2014_Array.Asp"-->  
<!--#Include file = "2014_Js.Asp"-->  
<!--#Include file = "2014_Css.Asp"-->  
<!--#Include file = "2014_Template.Asp"-->  
<!--#Include file = "2014_GBUTF.Asp"-->    
<!--#Include file = "2014_Form.Asp"-->     
<!--#Include file = "2014_Doc.Asp"--> 
<!--#Include file = "2014_Xls.Asp"--> 
<!--#Include file = "2014_Cache.Asp"--> 
<!--#Include file = "2014_Cookies.Asp"--> 
<!--#Include file = "2014_Author.Asp"--> 
<!--#Include file = "2015_Formatting.Asp"-->  
<!--#Include file = "2015_Color.Asp"-->  
<!--#Include file = "2015_Param.Asp"--> 
<!--#Include file = "2015_ASP.Asp"--> 
<!--#Include file = "2015_SqlIn.Asp"-->   
<!--#Include file = "2015_Json.Asp"--> 
<!--#Include file = "2015_PHP.Asp"-->    
<!--#Include file = "2015_APGeneral.Asp"-->  
<!--#Include file = "2016_Log.Asp"-->   
<!--#Include file = "SystemInfo.Asp"-->   
<!--#Include file = "2015_Param.Asp"-->    
<!--#Include file = "2016_WebControl.Asp"-->   
<!--#Include file = "2015_NewWebFunction.Asp"--> 
<!--#Include file = "2016_SaveData.Asp"-->   
<!--#Include file = "2015_ToMyPHP.Asp"-->   
 
<% 


server.scriptTimeout = 9999                                                     '��ʱ
session.timeout = 1440                                                          '���ֵΪ24Сʱ��Ҳ����˵��������session.timeout=1440��1441���ǲ�����
response.addheader "Content-Type", "text/html; charset=gb2312" 
'��ֹ����
response.buffer = true 
response.expires = 0 
response.expiresAbsolute = now() - 1 
response.cacheControl = "no-cache" 
response.addHeader "Expires", date() 
response.addHeader "Pragma", "no-cache" 
response.addHeader "Cache-Control", "private, no-cache, must-revalidate" 
dim conn, rs, rsx, rss, rst, rsd, tempRs, tempRs2, rsTemp 
dim pubTimer : pubTimer = now()                                                 'Timer()��Ҫ�����(20150728)

'Note:��������ԭ���ܼ򵥣����Բ����ӣ����޸�html�ģ����Բ��޸�asp����ɵ��ʽ�����Բ�Ҫ����(20160121)
dim WEBCOLUMNTYPE : WEBCOLUMNTYPE = "��ҳ|�ı�|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����" 
dim WEB_VIEWURL : WEB_VIEWURL = "/index.asp"                                   'ǰ̨��ʾ
dim WEB_ADMINURL : WEB_ADMINURL = "/admin/index.asp"                                  '��̨��ʾ
dim EDITORTYPE : EDITORTYPE = "asp"                                             '�������ͣ���ASP,��PHP,��jSP,��.NET
dim webDir : webDir = "" 
MDBPath = "" 
if checkFile(webDir & MDBPath) = false then
    webDir = "/" & getUrlDir(remoteUrlFileName(getThisUrl())) 
    webDir = mid(webDir, 1, len(webDir) - 1) 
    webDir = replace(LCase(webDir), "admin/", "") 
'call echo(webDir,mdbpath)
end if
dim isOnCacheHtml : isOnCacheHtml = false                                       '�Ƿ�������html

dim databaseType : databaseType = ""                                      '���ݿ����ͣ�Access��SqlServer֮��
dim sqlServerHostIP : sqlServerHostIP = "localhost"                          'SQL����IP��ַ��������������2008�汾ʹ��
dim sqlServerUsername : sqlServerUsername = "sa"                                'SQL���ݿ��û���
dim sqlServerPassword : sqlServerPassword = ""                                'SQL���ݿ��û�����
dim sqlServerDatabaseName : sqlServerDatabaseName = "webdata"                   'SQL���ݿ���
dim db_PREFIX : db_PREFIX = "x2asdy_"                                               '��ǰ׺
'=========



 
dim sArraySysColor 
sArraySysColor = array("#999", "#666") 

dim makeHtmlFileToLCase : makeHtmlFileToLCase = true                            '����HTML�ļ�תСд
dim isWebLabelClose : isWebLabelClose = true                                    '�պϱ�ǩ(20150831)
dim handleisCache : handleisCache = false                                       '�����Ƿ�����

dim isCacheTip : isCacheTip = false                                             '�Ƿ���������ʾ
dim adminDir : adminDir = "/admin/"                                               '��̨Ŀ¼


dim openErrorLog : openErrorLog = true                                          '����������־
dim openWriteSystemLog : openWriteSystemLog = "|txt|database|"                  '����дϵͳ��־ txtд���ı� databaseд�����ݿ�
dim isTestEcho : isTestEcho = true                                              '���ز��Ի���
dim webVersion : webVersion = "PAAJCMS v1.5"                                  '��վ�汾


'dim WEB_CACHEFile : WEB_CACHEFile=webDir & "/web/"& EDITORTYPE &"cachedata.txt"                                '�����ļ�
dim WEB_CACHEFile : WEB_CACHEFile = "[adminDir]/[EDITORTYPE]cachedata.txt" '�����ļ�
dim WEB_CACHEContent : WEB_CACHEContent = ""                                    '�����ļ�����

dim language : language = "#en-us"                                              'en-us  | zh-cn | zh-tw
dim moduleReplaceArray : moduleReplaceArray = array("", "")                     '�滻����20160630

 
'ǰ���õ�
dim code, pubCode                                                               'html����
dim templateName                                                                'ģ������
dim cfg_webSiteUrl, cfg_webTemplate, cfg_webImages, cfg_webCss, cfg_webJs, cfg_webTitle, cfg_webKeywords, cfg_webDescription, cfg_webSiteBottom, cfg_flags 
dim glb_columnName, glb_columnId, glb_id, glb_columnType, glb_columnENType, glb_table, glb_detailTitle, glb_flags 
'dim webTemplate                                                                 '��վģ��·��
dim glb_url, glb_filePath                                                       '��ǰ������ַ,���ļ�·��
dim glb_isonhtml                                                                '�Ƿ����ɾ�̬��ҳ
dim glb_locationType                                                            'λ������

dim glb_bodyContent                                                             '��������
dim glb_articleAuthor                                                           '��������
dim glb_articleAdddatetime                                                      '��������ʱ��
dim glb_articlehits                                                             '���������
dim glb_upArticle                                                               '��һƪ����
dim glb_downArticle                                                             '��һƪ����
dim glb_aritcleRelatedTags                                                      '���±�ǩ��
dim glb_aritcleSmallImage, glb_aritcleBigImage                                  '����Сͼ�����´�ͼ
dim glb_searchKeyWord                                                           '�����ؼ���
dim cacheHtmlFilePath                                                           '����html�ļ�·��
dim positionEndStr                                                              '�����β��׷������

dim isMakeHtml                                                                  '�Ƿ�������ҳ










'*****************************************  ����
'�����滻����
function batchReplaceL(content, str)
    dim splStr, s, i 
    splStr = split(str, "|*|") 
    for i = 0 to uBound(splStr)
        s = splStr(i) 
        if s <> "" then
            content = replaceL(content, s) 
        end if 
    next 
    batchReplaceL = content 
end function 
'�滻����
function replaceL(content, str)
    replaceL = replace(content, str, setL(str)) 
end function 
'����
function setL(str)
    dim c 
    c = str 
    if language = "en-us" then
        c = languageEN(c) 
    end if 
    setL = c 
end function 
'����������  c=handleDisplayLanguage(c,"loginok")
function handleDisplayLanguage(c, sType)
    '�����ֱ��ת���ˣ���Ҫ��һ��һ��ת�ˣ�
    if language = "zh-tw" then
        handleDisplayLanguage = simplifiedTransfer(c) 
        exit function 
    end if 
    if sType = "login" then

        c = batchReplaceL(c, "�벻Ҫ���������ַ�|*|������ȷ|*|�û�����������ĸ|*|�û�����������ĸ|*|�����û���Ϊ��|*|�����������ĸ|*|��������Ϊ��") 
        c = batchReplaceL(c, "��¼��̨|*|����Ա��¼|*|��������ǹ���Ա|*|������ֹͣ���ĵ�½��Ϊ|*|�û���|*|��") 
        c = batchReplaceL(c, "��&nbsp;��|*|����|*|������|*|�� ¼|*|��¼|*|�� ��|*|����") 


    elseif sType = "loginok" then
        c = batchReplaceL(c, "��̨��ͼ|*|�������|*|��������Ա|*|��ǰλ��|*|����Ա��Ϣ|*|�޸�����|*|���·ÿ���Ϣ|*|�����Ŷ�|*|��Ȩ����|*|������֧���Ŷ�") 
        c = batchReplaceL(c, "���������޸�ģʽ|*|ϵͳ��Ϣ|*|��ѿ�Դ��|*|��Ȩ��Ϣ|*|����������|*|�������汾|*|����Ⱥ|*|�������|*|��¼��̨") 
        c = batchReplaceL(c, "�û���|*|��ǰ׺|*|����|*|�˳�|*|����|*|��ҳ|*|Ȩ��|*|�˿�|*|����|*|����|*|��|*|�쳾����") 
        c = batchReplaceL(c, "ϵͳ����|*|�ҵ����|*|��Ŀ|*|ģ��|*|��Ա|*|����|*|��������") 

        c = batchReplaceL(c, "վ������|*|��վͳ��|*|����|*|��̨������־|*|��̨����Ա|*|��վ��Ŀ|*|������Ϣ|*|����|*|����ͳ��|*|��ҳ|*|��������|*|��Ƹ����") 
        c = batchReplaceL(c, "����|*|����|*|��Ա����|*|���۴�|*|��վ����|*|��վģ��|*|��̨�˵�|*|ִ��|*|��վ") 


    end if 
    handleDisplayLanguage = c 
end function 

'ΪӢ��
function languageEN(str)
    dim c 
    if str = "��¼�ɹ������ڽ����̨..." then
        c = "Login successfully, is entering the background..." 
    elseif str = "�˺��������<br>��¼����Ϊ " then
        c = "Account password error <br> login " 
    elseif str = "��¼�ɹ������ڽ����̨..." then
        c = "Login successfully, is entering the background..." 
    elseif str = "�˳��ɹ�" then
        c = "Exit success" 
    elseif str = "�˳��ɹ������ڽ����¼����..." then
        c = "Quit successfully, is entering the login screen..." 
    elseif str = "���������ɣ����ڽ����̨����..." then
        c = "Clear buffer finish, is entering the background interface..." 
    elseif str = "��ʾ��Ϣ" then
        c = "Prompt info" 
    elseif str = "������������û���Զ���ת����������" then
        c = "If your browser does not automatically jump, please click here" 
    elseif str = "����ʱ" then
        c = "Countdown " 
    elseif str = "��̨��ͼ" then
        c = "Admin map" 
    elseif str = "�������" then
        c = "Clear buffer" 
    elseif str = "��������Ա" then
        c = "Super administrator" 
    elseif str = "��ǰλ��" then
        c = "current location" 
    elseif str = "����Ա��Ϣ" then
        c = "Admin info" 
    elseif str = "�޸�����" then
        c = "Modify password" 
    elseif str = "�û���" then
        c = "username" 
    elseif str = "��ǰ׺" then
        c = "Table Prefix" 
    elseif str = "���������޸�ģʽ" then
        c = "online modification" 
    elseif str = "ϵͳ��Ϣ" then
        c = "system info" 
    elseif str = "��Ȩ��Ϣ" then
        c = "Authorization information" 
    elseif str = "��ѿ�Դ��" then
        c = "Free open source" 
    elseif str = "����������" then
        c = "Server name" 
    elseif str = "�������汾" then
        c = "Server version" 
    elseif str = "���·ÿ���Ϣ" then
        c = "visitor info" 
    elseif str = "�����Ŷ�" then
        c = "team info" 
    elseif str = "��Ȩ����" then
        c = "copyright" 
    elseif str = "������֧���Ŷ�" then
        c = "Develop and support team" 
    elseif str = "����Ⱥ" then
        c = "Exchange group" 
    elseif str = "�������" then
        c = "Related links" 
    elseif str = "ϵͳ����" then
        c = "System" 
    elseif str = "�ҵ����" then
        c = "My panel" 
    elseif str = "��Ŀ" then
        c = "Column" 
    elseif str = "ģ��" then
        c = "Template" 
    elseif str = "��Ա" then
        c = "Member" 
    elseif str = "����" then
        c = "Generation" 
    elseif str = "��������" then
        c = "More settings" 


    elseif str = "��¼��̨" then
        c = "Login management background" 
    elseif str = "����Ա��¼" then
        c = "Administrator login " 
    elseif str = "��������ǹ���Ա" then
        c = "If you are not an administrator" 
    elseif str = "������ֹͣ���ĵ�½��Ϊ" then
        c = "Please stop your login immediately" 
    elseif str = "��&nbsp;��" or str = "����" then
        c = "password" 
    elseif str = "������" then
        c = "Please input" 
    elseif str = "�� ¼" or str = "��¼" then
        c = "login" 
    elseif str = "�� ��" or str = "����" then
        c = "reset" 
    elseif str = "�벻Ҫ���������ַ�" then
        c = "Please do not enter special characters" 
    elseif str = "������ȷ" then
        c = "Input correct" 
    elseif str = "�û�����������ĸ" then
        c = "Username use " 
    elseif str = "�����û���Ϊ��" then
        c = "Your username is empty" 
    elseif str = "�����������ĸ" then
        c = "Passwords use " 
    elseif str = "��������Ϊ��" then
        c = "Your password is empty" 
    elseif str = "վ������" then
        c = "Site configuration" 
    elseif str = "��վͳ��" then
        c = "Website statistics" 
    elseif str = "��̨������־" then
        c = "Admin log " 
    elseif str = "��̨����Ա" then
        c = "Background manager" 
    elseif str = "��վ��Ŀ" then
        c = "Website column" 
    elseif str = "������Ϣ" then
        c = "Classified information" 
    elseif str = "����ͳ��" then
        c = "Search statistics" 
    elseif str = "��ҳ" then
        c = "Single page" 
    elseif str = "��������" then
        c = "Friendship link" 
    elseif str = "��Ƹ����" then
        c = "Recruitment management" 
    elseif str = "����" then
        c = "Feedback management" 
    elseif str = "����" then
        c = "message management" 
    elseif str = "��Ա����" then
        c = "Member allocation" 
    elseif str = "���۴�" then
        c = "Bidding words" 
    elseif str = "��վ����" then
        c = "Website layout" 
    elseif str = "��վģ��" then
        c = "Website module" 
    elseif str = "��̨�˵�" then
        c = "Background menu" 
    elseif str = "��վ" then
        c = "Template website " 

    elseif str = "11111" then
        c = "1111111" 



    elseif str = "ִ��" then
        c = "implement " 
    elseif str = "����" then
        c = "comment " 
    elseif str = "����" then
        c = "generate " 
    elseif str = "Ȩ��" then
        c = "jurisdiction" 
    elseif str = "����" then
        c = "Help" 
    elseif str = "�˳�" then
        c = "sign out" 
    elseif str = "����" then
        c = "hello" 
    elseif str = "��ҳ" then
        c = "home" 
    elseif str = "�˿�" then
        c = "port" 
    elseif str = "����" then
        c = "official website" 
    elseif str = "����" then
        c = "Emai" 
    elseif str = "�쳾����" then
        c = "Cloud" 

    elseif str = "��" then
        c = " edition"  

    else
        c = str 
    end if 
    languageEN = c 
end function 
%> 

















