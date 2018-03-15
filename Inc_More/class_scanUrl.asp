<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-15
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'网址采集类

'dim t:  set t=new class_scanUrl

class class_scanUrl
	dim thisFormatObj,thisContent,httpUrl,webSite,webSiteName,otherWebSiteList
	dim htmlSaveDir,htmlSaveFilePath,htmlSaveFileCharset,htmlSaveFileSize,htmlSaveOpenSpeed
	dim nmails,mails						'邮箱列表
	dim nlinks							'外链总数
	 '构造函数 初始化
    Sub Class_Initialize() 
		set thisFormatObj=new class_formatting			'格式化类
		htmlSaveDir="E:\E盘\WEB网站\网站采集数据\"
		nlinks=0
		nmails=0
		htmlSaveOpenSpeed=0
	end sub
    '析构函数 类终止
    Sub Class_Terminate()
        'HtmlFolder=nothing
        'HtmlFilename=nothing
        'HtmlContent=nothing
        'Urlname=nothing
    End Sub
	'处理内容
	sub handleContent(content)
		'thisContent=thisFormatObj.handleFormatting(content) 
	end sub
	'处理文件内容
	sub handleFileContent(byval filePath) 
		htmlSaveFilePath=handlePath(filePath)
		htmlSaveFileCharset=getFileCharset(htmlSaveFilePath)
		htmlSaveFileSize=getFSize(htmlSaveFilePath)
		thisContent=readFile(htmlSaveFilePath,"")
		
		mails=""
		nlinks=0
		
        thisContent = Left(thisContent, 302400)         '只处理300K，因为太大慢慢会很慢的
		call handleContent(thisContent)
	end sub
	'设置网址
	sub setHttpUrl(url)
		httpUrl=url
		webSite=getWebSite(url)
		webSiteName=lcase(getWebSiteName(url))
	end sub
	'获得处理域名列表
	sub getHandleWebSiteNameList()
		dim formatObj,content,splstr,s,c,url,sUrl,title,nACount,thisWebSiteName,sql,sWebSite
		'content=thisFormatObj.handleLabelContent("a","*","","get+label")			'这种方法不能对大文件处理
		'splstr=split(content,"$Array$")
		content=getContentAUrlList(thisContent)
		
		splstr=split(content,vbcrlf)
		nACount=ubound(splstr)							'链接总数
		call echoRedB("操作","共有链接" & nACount & "条")
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
	'更新当前域名操作
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
	'获得下一个域名网址
	function getNextWebSite() 
		call openconn()
		getNextWebSite=getFieldValue(db_PREFIX & "webdomain", "website", " where isthrough<>0 order by id") 
	end function
	'获得域名网址列表
	function getWebSiteList(nTopNumb) 
		getWebSiteList=getFieldValueList("Select top "& nTopNumb &" * From "& db_PREFIX & "webdomain where isthrough<>0  order by id", "website")
	end function
	'批量写入域名
	function batchAddWebSite()
		dim splstr,website,nAddOK
		call openconn()
		call echoRedB("操作","正在批量写入域名列表" & otherWebSiteList)
		nAddOK=0					'添加成功数
		splstr=split(otherWebSiteList,vbcrlf)
		for each website in splstr
			if website<>"" then
				if getRecordCount( db_PREFIX & "webdomain", " where website='"& website &"'" )=0 then
				  call connInsertInto("insert into " & db_PREFIX & "webdomain (website,isthrough,isdomain) values('"& ADSql(website) &"',1,1)")
				  nAddOK=nAddOK+1
				  call echo("添加 " & nAddOK, website)
				end if
			end if
		next
	end function
	'打印
	function print()
		call rw( thisFormatObj.printWrite() )	
	end function
	'获得网址要存放路径
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
	'搜索网址
	function scanHttpUrl(url)
	
	
	end function
	
end class 


%>