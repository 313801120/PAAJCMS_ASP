<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-15
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'��ַ�ɼ���

'dim t:  set t=new class_scanUrl

class class_scanUrl
	dim thisFormatObj,thisContent,httpUrl,webSite,webSiteName,otherWebSiteList
	dim htmlSaveDir,htmlSaveFilePath,htmlSaveFileCharset,htmlSaveFileSize,htmlSaveOpenSpeed
	dim nmails,mails						'�����б�
	dim nlinks							'��������
	 '���캯�� ��ʼ��
    Sub Class_Initialize() 
		set thisFormatObj=new class_formatting			'��ʽ����
		htmlSaveDir="E:\E��\WEB��վ\��վ�ɼ�����\"
		nlinks=0
		nmails=0
		htmlSaveOpenSpeed=0
	end sub
    '�������� ����ֹ
    Sub Class_Terminate()
        'HtmlFolder=nothing
        'HtmlFilename=nothing
        'HtmlContent=nothing
        'Urlname=nothing
    End Sub
	'��������
	sub handleContent(content)
		'thisContent=thisFormatObj.handleFormatting(content) 
	end sub
	'�����ļ�����
	sub handleFileContent(byval filePath) 
		htmlSaveFilePath=handlePath(filePath)
		htmlSaveFileCharset=getFileCharset(htmlSaveFilePath)
		htmlSaveFileSize=getFSize(htmlSaveFilePath)
		thisContent=readFile(htmlSaveFilePath,"")
		
		mails=""
		nlinks=0
		
        thisContent = Left(thisContent, 302400)         'ֻ����300K����Ϊ̫�������������
		call handleContent(thisContent)
	end sub
	'������ַ
	sub setHttpUrl(url)
		httpUrl=url
		webSite=getWebSite(url)
		webSiteName=lcase(getWebSiteName(url))
	end sub
	'��ô��������б�
	sub getHandleWebSiteNameList()
		dim formatObj,content,splstr,s,c,url,sUrl,title,nACount,thisWebSiteName,sql,sWebSite
		'content=thisFormatObj.handleLabelContent("a","*","","get+label")			'���ַ������ܶԴ��ļ�����
		'splstr=split(content,"$Array$")
		content=getContentAUrlList(thisContent)
		
		splstr=split(content,vbcrlf)
		nACount=ubound(splstr)							'��������
		call echoRedB("����","��������" & nACount & "��")
		for each url in splstr
			url=phpTrim(url)
			if url <>"" then
				sUrl=lcase(phptrim(url))
				'call echo("url",url)
				if not (sUrl="#" or sUrl="" or sUrl="/" or sUrl="\" or left(sUrl,11)="javascript:") and instr(sUrl,"@")=false and instr(sUrl,"mailto:")=false then
					thisWebSiteName=lcase(getWebSiteName(url))
					sWebSite=getWebSite(url)
					if thisWebSiteName<>webSiteName and sWebSite<>"" and instr(vbcrlf & otherWebSiteList & vbcrlf, vbcrlf & sWebSite & vbcrlf)=false and instr(thisWebSiteName,".")=false then
						otherWebSiteList=otherWebSiteList & sWebSite & vbcrlf
						nlinks=nlinks+1
					end if
				elseif left(sUrl,7)="mailto:" and instr(sUrl,"@")>0 then
					sUrl=mid(url,8)
					if instr(sUrl,"&")>0 then
						sUrl=mid(sUrl,1,instr(sUrl,"&")-1)
					end if
					if getWebSite(sUrl)<>"" then
						if mails<>"" then
							mails=mails & vbcrlf
						end if
						mails=mails & replace(sUrl,"/","")
						nmails=nmails+1 
					end if
				end if
			end if
		next
	end sub
	'���µ�ǰ��������
	sub updateThisWebSiteInfo(webstate)
		dim sql
		call openconn() 
		if getRecordCount( db_PREFIX & "webdomain", " where website='"& website &"'" )=0 then
			call connInsertInto("insert into " & db_PREFIX & "webdomain (website,isdomain) values('"& ADSql(website) &"',1)")
		end if
		sql="isthrough=0,nlinks="&nlinks &",webstate='"& webstate &"',charset='"& htmlSaveFileCharset &"',websize="& htmlSaveFileSize &""
		sql=sql & ",openspeed="& htmlSaveOpenSpeed &",nmails="& nmails &",mails='"& mails &"'"
		connUpdate("update " & db_PREFIX & "webdomain set "& sql &" where website='"& website &"'") 
	end sub
	'�����һ��������ַ
	function getNextWebSite() 
		call openconn()
		getNextWebSite=getFieldValue(db_PREFIX & "webdomain", "website", " where isthrough<>0 order by id") 
	end function
	'���������ַ�б�
	function getWebSiteList(nTopNumb) 
		getWebSiteList=getFieldValueList("Select top "& nTopNumb &" * From "& db_PREFIX & "webdomain where isthrough<>0  order by id", "website")
	end function
	'����д������
	function batchAddWebSite()
		dim splstr,website,nAddOK
		call openconn()
		call echoRedB("����","��������д�������б�" & otherWebSiteList)
		nAddOK=0					'��ӳɹ���
		splstr=split(otherWebSiteList,vbcrlf)
		for each website in splstr
			if website<>"" then
				if getRecordCount( db_PREFIX & "webdomain", " where website='"& website &"'" )=0 then
				  call connInsertInto("insert into " & db_PREFIX & "webdomain (website,isthrough,isdomain) values('"& ADSql(website) &"',1,1)")
				  nAddOK=nAddOK+1
				  call echo("��� " & nAddOK, website)
				end if
			end if
		next
	end function
	'��ӡ
	function print()
		call rw( thisFormatObj.printWrite() )	
	end function
	'�����ַҪ���·��
	function getUrlSavePath(url) 
		dim fileDir 
		fileDir=left(lcase(getWebSiteName(url)),1)
		if fileDir="" then
			getUrlSavePath=""
			exit function
		end if 
		fileDir=handlePath(htmlSaveDir  & "\" & fileDir)
		if checkFolder(fileDir)=false then
			call createFolder(fileDir)
		end if
		
		getUrlSavePath=fileDir & "\" & setFileName(url)
	end function
	'������ַ
	function scanHttpUrl(url)
	
	
	end function
	
end class 


%>