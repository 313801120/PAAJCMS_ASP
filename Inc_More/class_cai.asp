<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-16
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'采集类

'dim t:  set t=new class_cai 

class class_cai
	dim thisFormatObj,thisContent,httpUrl,webSite,webSiteName,otherWebSiteList
	dim htmlSaveDir,htmlSaveFilePath,htmlSaveFileCharset,htmlSaveFileSize,htmlSaveOpenSpeed 
	dim nPageCount				'页数
	dim splListArray			'列表数组
	 '构造函数 初始化
    Sub Class_Initialize() 
		set thisFormatObj=new class_formatting			'格式化类
		htmlSaveDir="E:\E盘\WEB网站\采集数据存放区\美图\html\"
		nPageCount=0
	end sub
    '析构函数 类终止
    Sub Class_Terminate() 
    End Sub
	'处理内容
	sub handleContent(content)
		thisContent=thisFormatObj.handleFormatting(content) 
	end sub
	'处理文件内容
	sub handleFileContent(byval filePath) 
		htmlSaveFilePath=handlePath(filePath)
		htmlSaveFileCharset=getFileCharset(htmlSaveFilePath)
		htmlSaveFileSize=getFSize(htmlSaveFilePath)
		
		thisContent=readFile(htmlSaveFilePath,"") 
        thisContent = Left(thisContent, 302400)         '只处理300K，因为太大慢慢会很慢的
		call handleContent(thisContent)
	end sub
	'获得当前页数
	function getPageCount(labelName,sType)
		dim content,splstr,s,c,cf2,sPage
		content=thisFormatObj.handleLabelContent(labelName,"*","","get")			'翻页 
		 
		'处理页总数
		if sType="末页" then
			set cf2=new class_formatting
			call cf2.handleFormatting(content)
			c=cf2.handleLabelContent("a","*","","get+label") 
			splstr=split(c,"$Array$")
			for each s in splstr
				if instr(s,"末页") then
					sPage=getNumber(s)
					exit for
				end if
			next 
  		ElseIf sType = "详细页li0" Then
            '处理页总数
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
	'获得页网址列表   .htm[$split$]_*.htm
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
			'配置页不为空
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
	'获得列表数组
	function getListArray(labelName,splType)
		dim content,splstr,s,c,cf2
		content=thisFormatObj.handleLabelContent(labelName,"*","","get")			'翻页 
		
		set cf2=new class_formatting
		call cf2.handleFormatting(content)
		c=cf2.handleLabelContent(splType,"*","","get+label")
		splListArray=split(c,"$Array$")
		getListArray=splListArray
	end function
	 
	'设置网址
	sub setHttpUrl(url)
		httpUrl=url
		webSite=getWebSite(url)
		webSiteName=lcase(getWebSiteName(url))
	end sub 
	'打印
	function print()
		call rw( thisFormatObj.printWrite() )	
	end function
	'获得网址要存放路径
	function getUrlSavePath(url)   
		getUrlSavePath=handlePath( htmlSaveDir & "\" & setFileName(url) )
	end function
end class 


%>