<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-16
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'�ɼ���

'dim t:  set t=new class_cai 

class class_cai
	dim thisFormatObj,thisContent,httpUrl,webSite,webSiteName,otherWebSiteList
	dim htmlSaveDir,htmlSaveFilePath,htmlSaveFileCharset,htmlSaveFileSize,htmlSaveOpenSpeed 
	dim nPageCount				'ҳ��
	dim splListArray			'�б�����
	 '���캯�� ��ʼ��
    Sub Class_Initialize() 
		set thisFormatObj=new class_formatting			'��ʽ����
		htmlSaveDir="E:\E��\WEB��վ\�ɼ����ݴ����\��ͼ\html\"
		nPageCount=0
	end sub
    '�������� ����ֹ
    Sub Class_Terminate() 
    End Sub
	'��������
	sub handleContent(content)
		thisContent=thisFormatObj.handleFormatting(content) 
	end sub
	'�����ļ�����
	sub handleFileContent(byval filePath) 
		htmlSaveFilePath=handlePath(filePath)
		htmlSaveFileCharset=getFileCharset(htmlSaveFilePath)
		htmlSaveFileSize=getFSize(htmlSaveFilePath)
		
		thisContent=readFile(htmlSaveFilePath,"") 
        thisContent = Left(thisContent, 302400)         'ֻ����300K����Ϊ̫�������������
		call handleContent(thisContent)
	end sub
	'��õ�ǰҳ��
	function getPageCount(labelName,sType)
		dim content,splstr,s,c,cf2,sPage
		content=thisFormatObj.handleLabelContent(labelName,"*","","get")			'��ҳ 
		 
		'����ҳ����
		if sType="ĩҳ" then
			set cf2=new class_formatting
			call cf2.handleFormatting(content)
			c=cf2.handleLabelContent("a","*","","get+label") 
			splstr=split(c,"$Array$")
			for each s in splstr
				if instr(s,"ĩҳ") then
					sPage=getNumber(s)
					exit for
				end if
			next 
  		ElseIf sType = "��ϸҳli0" Then
            '����ҳ����
            cf2 = New class_formatting
            Call cf2.handleFormatting(content)
            c = cf2.handleLabelContent("li:eq(0)", "*", "", "get+label")
            sPage = getNumber(c) 
		end if
		if sPage="" then
			nPageCount=0
		else
			nPageCount=cInt(sPage)
		end if
		getPageCount=nPageCount
	end function
	'���ҳ��ַ�б�   .htm[$split$]_*.htm
	function getUrlPageList(byval url,pageConfig, byval nStartPage, byval nEndPage)
		dim urlHead,urlFoot,nLen,i,s,c,splxx
		if right(url,1)="/" then
			urlHead=url
			urlFoot=pageConfig
		else
			nLen=instrrev(url,"/")
			urlHead=mid(url,1,nLen)
			urlFoot=mid(url,nLen+1)
			if getnumber(urlFoot)="1" then
				urlFoot=replace(urlFoot,"1","*")
			end if
			'����ҳ��Ϊ��
			if pageConfig<>"" then
				if instr(pageConfig,"[$split$]")>0 then
					splxx=split(pageConfig,"[$split$]")
					urlFoot=replace(urlFoot,splxx(0),splxx(1))
				else
					urlFoot=pageConfig
				end if
			end if
			call echo("urlHead",urlHead)
			call echo("urlFoot",urlFoot)
		end if
		if nEndPage>nPageCount or nPageCount=-1 then
			nEndPage=nPageCount
		end if
		for i=nStartPage to nPageCount
			s=urlHead & replace(urlFoot,"*",i)
			if c<>"" then
				c=c & vbcrlf
			end if
			c=c & s
		next
		getUrlPageList=c
	end function
	'����б�����
	function getListArray(labelName,splType)
		dim content,splstr,s,c,cf2
		content=thisFormatObj.handleLabelContent(labelName,"*","","get")			'��ҳ 
		
		set cf2=new class_formatting
		call cf2.handleFormatting(content)
		c=cf2.handleLabelContent(splType,"*","","get+label")
		splListArray=split(c,"$Array$")
		getListArray=splListArray
	end function
	 
	'������ַ
	sub setHttpUrl(url)
		httpUrl=url
		webSite=getWebSite(url)
		webSiteName=lcase(getWebSiteName(url))
	end sub 
	'��ӡ
	function print()
		call rw( thisFormatObj.printWrite() )	
	end function
	'�����ַҪ���·��
	function getUrlSavePath(url)   
		getUrlSavePath=handlePath( htmlSaveDir & "\" & setFileName(url) )
	end function
end class 


%>