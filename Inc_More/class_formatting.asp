<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-02-27
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'格式化html  神作20170829
'dim t:  set t=new class_formatting
'call echo("",t.handleFormatting) 
class class_formatting 
	dim pubContent				'内容
	dim nCountLen				'字符长度
	dim beforeStr				'上一个字符
	dim afterStr				'下一个字符
	dim isDisplayEcho			'是否回显
	dim islabel					'是否为标签
	dim labelLevelName(29999,11)		'标签等级名称    标签名称  左边位置 字符长度  右边位置  标签内容 标签上面内容，标签下面内容 9为动作
	dim nLevel,tempNLevel                 	'级别数
	dim nLabelIndex				'标签索引
	dim imgList					'图片列表
	dim cssList					'CSS列表
	dim jsList					'JS列表		
	dim swfList					'SWF列表	
	dim sourceList				'source列表 视频 
	dim videoList				'video列表 视频 
	dim aList					'a列表 链接地址20171207 
		 
	dim urlList					'总素材网址列表
    dim nRowCount				'行总数

	 '构造函数 初始化
    Sub Class_Initialize() 
		isDisplayEcho=false
		islabel=false			'是否为标签 默认为假
	 	nLevel = 0              '级别数
		nLabelIndex=0			'标签索引默认为零
	end sub
    '析构函数 类终止
    Sub Class_Terminate() 
    End Sub 
	'处理格式化
	function handleFormatting(content)
		pubContent=content
		dim i,s,s2,labelS,labelName,parentLableName,endCode,nLeft,nRight,nLen,isOK,parentObjId,nGetLeft,lableC,lableNRow 
		dim nNoteLeft,nNoteRight,isHtmlNote:isHtmlNote=false				'是否为html注释
		dim isHtmlStyle:isHtmlStyle=false			'是否为style样式
		dim isHtmlStyleNote:isHtmlStyleNote=false	'是否为style样式注释
		dim isHtmlJs:isHtmlJs=false					'是否为js
		dim isHtmlJsNote:isHtmlJsNote=false			'是否为js注释
		dim jsWriteLabel			    			'为js输入标签名称  如'或"
		dim brStr									'换行值
		brStr=chr(10)				'用这个，则不用vbcrlf  这能通用于只有chr(10)内容
		'if instr(content,vbcrlf)>0 then
		'	brStr=vbcrlf
		'end if
		'brStr=chr(10)
		'call echo("brStr",len(brStr))
		
		nCountLen=len(content)
		for i = 1 to nCountLen
			s = mid(content, i, 1) 
			'为html注释 运行时
			if isHtmlNote=true then
				if mid(content,i,3)="-->" then
					isHtmlNote=false				'html注释为假
					i=i+2
					'注释也放到标签数据组里20161024
					nNoteRight=i-nNoteLeft+1
					labelS=mid(content, nNoteLeft, nNoteRight) 
					nLabelIndex=nLabelIndex+1
					labelLevelName(nLabelIndex,0)=nLevel					
					labelLevelName(nLabelIndex,1)="!--htmlnote--"
					labelLevelName(nLabelIndex,2)=nNoteLeft
					labelLevelName(nLabelIndex,3)=len(labelS)
					labelLevelName(nLabelIndex,4)=i+1
					labelLevelName(nLabelIndex,5)=labelS
					labelLevelName(nLabelIndex,6)=""			'标签前面内容
					labelLevelName(nLabelIndex,7)=""			'标签后面内容
					labelLevelName(nLabelIndex,9)=""			'标签动作 
				end if
			'判断是否为html 注释
			elseif mid(content,i,4)="<!--"  then
				nNoteLeft=i
				isHtmlNote=true				'html注释为真 
			'样式注释开始
			elseif isHtmlStyle=true and isHtmlStyleNote=false and s="/" and mid(content & " ", i+1, 1)="*" then
				i=i+1
				isHtmlStyleNote=true
			'样式注释结束
			elseif isHtmlStyle=true and isHtmlStyleNote=true and s="*" and mid(content & " ", i+1, 1)="/" then
				i=i+1
				isHtmlStyleNote=false 
			elseif isHtmlStyle=true and isHtmlStyleNote=true  then
				'为css样式里的注释，不处理了
				
			'JS 单行注释
			elseif isHtmlJs=true and isHtmlJsNote=false and s="/" and mid(content & " ", i+1, 1)="/" and jsWriteLabel="" then
				s2=mid(content,i)
				i=i+instr(s2,brStr)  
			'JS 多行注释开始
			elseif isHtmlJs=true and isHtmlJsNote=false and s="/" and mid(content & " ", i+1, 1)="*" and jsWriteLabel="" then
				i=i+1
				isHtmlJsNote=true 
			'JS 多行注释结束
			elseif isHtmlJs=true and isHtmlJsNote=true and s="*" and mid(content & " ", i+1, 1)="/" then
				i=i+1
				isHtmlJsNote=false 
			elseif isHtmlJs=true and isHtmlJsNote=true  then
				'JS 多行注释里的注释，不处理了
				
			'js里输出内容处理20161115
			elseif isHtmlJs=true and instr("\'""",s)>0 then
				'输入标签结果
				if s="\" then
					i=i+1
				elseif jsWriteLabel<>"" and jsWriteLabel=s then
					jsWriteLabel=""
				elseif s="'" or s="""" then
					jsWriteLabel=s 
				end if 
			elseif s="<"  then 
				endCode=mid(content,i)
				'最后一个标签没有>则加上，要不出错，晕，2016116   程序写到最后自己都晕了，只能把代码让以后的人来改了
				if instr(endCode,">")=false then
					endCode=endCode & ">"
				end if
				labelS=mid(endCode,1,instr(endCode,">"))
				nLeft=i
				nLen=len(labelS)
				nRight=i+nLen
				nLabelIndex=nLabelIndex+1
				labelName=getLabelName(labelS)				'标签名称
				'上一个标签名称
				parentLableName=""
				if nLabelIndex>0 then
					parentLableName=labelLevelName(nLabelIndex-1,1)		
				end if
				
				i=i+len(labelS)-1
				isOK=true
				if labelName="style" then
					isHtmlStyle=true
				elseif labelName="/style" then
					isHtmlStyle=false
				elseif isHtmlStyle=true then 
					isOK=false
					nLabelIndex=nLabelIndex-1
				
				elseif labelName="script" then
					isHtmlJs=true
					jsWriteLabel=""			'js输入标签，查找用到
				elseif labelName="/script" then
					isHtmlJs=false
				elseif isHtmlJs=true then
					isOK=false
					nLabelIndex=nLabelIndex-1
				end if
				if isOK=true then
					labelLevelName(nLabelIndex,0)=nLevel
					if left(labelS,2)="</" then
						nLevel=nLevel-1
						labelLevelName(nLabelIndex,0)=nLevel
						'不等于上一级标签，则再减一
						if labelName <> "/" & parentLableName  and nLevel>0 then		'and left(parentLableName,1)<>"/"  不要判断上一级是否是</
							tempNLevel=getEndLabelToStartLable(nLabelIndex, nLevel, labelName,parentObjId)		'获得上一级等级数
							'call echoyellowb("tempNLevel",tempNLevel)
							if tempNLevel<>-1 then
								nLevel=tempNLevel 
								nGetLeft=labelLevelName(parentObjId,2)
								lableC=mid(content,nGetLeft,nRight-nGetLeft)
								lableNRow=ubound(split(lableC,brStr))
								
								labelLevelName(nLabelIndex,8)=lableNRow			'标签中间内容有多少行 结束标签
								labelLevelName(parentObjId,8)=lableNRow			'标签中间内容有多少行 开始标签 
							else  
								nLevel=nLevel-1
							end if
							labelLevelName(nLabelIndex,0)=nLevel
						'加这个判断是为了收集开始标签与结束标签之间内容的行数
						elseif left(labelName,1)="/" then
							tempNLevel=getEndLabelToStartLable(nLabelIndex, nLevel, labelName,parentObjId)		'获得上一级等级数 
							if tempNLevel<>-1 then 
								nGetLeft=labelLevelName(parentObjId,2)
								lableC=mid(content,nGetLeft,nRight-nGetLeft)
								lableNRow=ubound(split(lableC,brStr))
								
								labelLevelName(nLabelIndex,8)=lableNRow			'标签中间内容有多少行 结束标签
								labelLevelName(parentObjId,8)=lableNRow			'标签中间内容有多少行 开始标签 
							end if	
						end if
						
					elseif right(labelS,2)<>"/>" then
						if ucase(labelName)<>"!DOCTYPE" then
							nLevel=nLevel+1
						end if
						'call echo("加" & nLevel,replace(labelS,"<","&lt;"))
					end if
					
					labelLevelName(nLabelIndex,1)=labelName
					labelLevelName(nLabelIndex,2)=nLeft
					labelLevelName(nLabelIndex,3)=nLen
					labelLevelName(nLabelIndex,4)=nRight
					labelLevelName(nLabelIndex,5)=labelS
					labelLevelName(nLabelIndex,6)=""			'标签前面内容
					labelLevelName(nLabelIndex,7)=""			'标签后面内容
					labelLevelName(nLabelIndex,9)=""			'标签动作
				end if
			end if 
			
			
		next
		
		handleFormatting=content
	end function
	'获得标签正对出现的开始位置
	function getEndLabelToStartLable(byval nIndex, byval findNLevel, byval findLabelName, parentObjId)
		dim i,labelName,nLevel,nLeft,nLen,nRight,s,nI
		parentObjId=-1
		findLabelName=lcase(findLabelName)
		for i = 1 to nIndex
			nI=nIndex-i
			nLevel=labelLevelName(nI,0)
			labelName=lcase(labelLevelName(nI,1))
			nLeft=labelLevelName(nI,2)
			nLen=labelLevelName(nI,3)
			nRight=labelLevelName(nI,4)
			
			if findLabelName = "/" & labelName and findNLevel>=nLevel then
				parentObjId=nI						'当前索引，也就是ID
				getEndLabelToStartLable=nLevel
				exit function
			end if 
			'call echoYellowB(i, findLabelName & "," & labelName)
		next
		getEndLabelToStartLable=-1
	end function 
	'获得标签名称
	function getLabelName(byval labelS)
		dim labelName
		labelName=replace(replace(labelS,">"," "),vbtab," ")
		labelName=phptrim(mid(labelName,2,instr(labelName," ")-1)) 
		getLabelName=labelName
	end function
	'打印
	function printWrite()
		dim i,labelName,nLevel,nLeft,nLen,nRight,s,lableNRow
		for i =1 to nLabelIndex
			nLevel=labelLevelName(i,0)
			labelName=lcase(labelLevelName(i,1))
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			s=labelLevelName(i,5)
			lableNRow=labelLevelName(i,8)
			s=replace(s,"<","&lt;")
			if instr("|div|span|", "|" & labelName & "|")>0 then
				call rw("<"& labelName &">")
			end if
			s=s & "("& lableNRow &")行"
			call echoredb( copyStr("&nbsp;&nbsp;",nLevel) & nLevel & "|" & labelName, nLeft & "," & nRight & " , " & s)
			if instr("|/div|/span|", "|" & labelName & "|")>0 then
				call rw("<"& labelName &">")
			end if
		next
		call rw("<script src='/Jquery/Jquery.Min.js' type='text/javascript'></script>")
		call rw("<script src='/Inc_More/z_formatting.js' type='text/javascript'></script>")
		call rw(copyStr("<br>" & vbcrlf ,10))
	end function
	
	'获得格式化HTML 
	function getFormattingHtml(byval sHtmlType, byval sJsType, byval sCssType)
		dim i,labelName,parentLabelName,nLevel,nLeft,nLen,nRight,s,c,labelBeforeStr,labelAfterStr,parentNRight,labelS,s2,labelNRow,nLen2
		dim action,parentAction			'当前动作与上一个标签动作
		dim isPre:isPre=false			'默认Pre为假
		sHtmlType="|"& sHtmlType &"|"			'html操作类型
		sJsType="|"& sJsType &"|"			'js操作类型
		sCssType="|"& sCssType &"|"			'css操作类型
		for i =1 to nLabelIndex
			parentLabelName=labelName
			nLevel=labelLevelName(i,0)
			labelName=labelLevelName(i,1)
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			labelS=labelLevelName(i,5)
			labelBeforeStr=labelLevelName(i,6) 	'标签前面内容
			labelAfterStr=labelLevelName(i,7) 		'标签后面内容
			labelNRow=labelLevelName(i,8)			'标签内容行数
			action = "|" & labelLevelName(i,9) & "|" 		'当前动作  
			parentAction = "|" & labelLevelName(i-1,9) & "|" 		'上一个标签动作  
			parentNRight=labelLevelName(i-1,4) 						'上一个右边字符位置
			 
			if parentLabelName="pre" then
				isPre=true 
			elseif parentLabelName="/pre" then
				isPre=false
			end if
			
			
			'和一个标签
			if i=1 then
				if nLeft>1 then
					s=mid(pubContent,1,nLeft-1)
					'内容清除两边空格
					if instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 then
						s=phptrim(s)
					end if
					c=c & s
				end if
				'动作不是删除，则处理
				if instr("|del|",action)=false then
					c=c & labelBeforeStr			'标签前面的内容
					c=c & labelS				'标题本身
					c=c & labelAfterStr			'标签后面的内容
				end if
				 
			else 
				if instr("|replace|del|",parentAction)>0 then
					'为删除与替换，不处理 
				else 
					s2=mid(pubContent,parentNRight,nLeft-parentNRight)
					'压缩JS
					if  instr(sJsType, "|zip|")>0  and labelName="/script" then 
						s2=pubJsObj.handleJsContent(s2, sJsType)		'删除js里注释
					'压缩CSS
					elseif  instr(sCssType,"|zip|")>0   and labelName="/style" then
						s2=pubCssObj.handleCssContent("css",s2,"*","","","", sCssType)
					'解缩JS
					elseif   instr(sJsType, "|unzip|")>0  and labelName="/script" then
						s2=pubJsObj.handleJsContent(s2, sJsType)
					'解缩CSS
					elseif instr(sCssType, "|unzip|")>0  and labelName="/style" then
						s2=pubCssObj.handleCssContent("css",s2,"*","","","", sCssType)
						
						
					'内容清除两边空格
					elseif ( instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 ) and isPre=false and labelName<>"/pre" then 
						s2=phptrim(s2)
					end if
					'解压缩进
					if instr(sHtmlType, "|unzip|")>0 and isPre=false then
						 
						if labelNRow>0 then
							labelS=vbcrlf & "" &  copyStr("    ",nLevel ) & labelS
						elseif len(s2)=0 and "/" & parentLabelName<>labelName and instr("|br|hr|","|"& labelName &"|")=false then
							'call echoRedB(parentLabelName,labelName)
							labelS=vbcrlf & "" &  copyStr("    ",nLevel ) & labelS  
						end if
						if labelName="/style" and labelName="/script" then
							s2=handleHtmlTab(s2,nLevel)			'处理Html的Tab退格
						end if
					end if
					'删除HTML注释
					if instr(sHtmlType, "|delnote|")>0 and labelName="!--htmlnote--" then
						labelS=""
					end if
					c=c & s2 					'上一个标签的内容
				end if				
				c=c & labelBeforeStr			'标签前面的内容
				'当前标签不为删除
				if instr("|del|",action)=false then
					c=c & labelS				'标题本身
				end if 
				c=c & labelAfterStr			'标签后面的内容 
			end if
			'最后标签，追加最后剩下的内容
			if i = nLabelIndex then 
				nLen2=nCountLen-nRight+1
				if nLen2<0 then nLen2=0 			'不能为负数 20171113  对这个内容 a<b 
				s=mid(pubContent,nRight,nLen2)
				'内容清除两边空格
					if instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 then
					s=phptrim(s)
				end if
				c=c & s
			end if
			
		next
		getFormattingHtml=c
	end function
	'处理Html的Tab退格 20161025
	function handleHtmlTab(byval content, byval nLevel)			
		dim i,splstr,s,c
		splstr=split(content,vbcrlf)
		for each s in splstr
			s=phptrim(s)
			if s<>"" then
				if c<>"" then
					c=c & vbcrlf & copyStr("    ",nLevel )
				end if
				c=c & s
			end if
		next
		handleHtmlTab=c
	end function
	'判断标签内容
	function checkLabelHtml(labelNameList) 
		checkLabelHtml=handleLabelContent(labelNameList,"*","","check")
	end function
	'获得标签内容
	function getLabelHtml(labelNameList) 
		getLabelHtml=handleLabelContent(labelNameList,"*","","get") 
	end function
	'设置标签内容 并格式化
	function setLabelHtml(labelNameList,htmlValue)  
		call  handleLabelContent(labelNameList,"*",htmlValue,"set")	'修改
		setLabelHtml=getFormattingHtml("","","")		'获得格式化后的html
	end function
	'设置标签内容 并格式化
	function delLabelHtml(labelNameList)  
		call  handleLabelContent(labelNameList,"*","","del")	'修改
		delLabelHtml=getFormattingHtml("","","")		'获得格式化后的html
	end function
'before() - 在被选元素之前插入 after() - 在被选元素之后插入 prepend() - 在被选元素的开头插入 append() - 在被选元素的结尾插入
	
	'处理标签指定内容   sType='set' 为替换html内容 
	function handleLabelContent(byval labelNameList, byval paramName, byval htmlValue, byval sType) 
		dim i,labelName,nLevel,nLeft,nLen,nRight,s,c,labelBeforeStr,labelAfterStr,parentNRight,labelS,lCaseLabelS,s2
		dim LeftS,RightS,nEQ,nCount,findLabelLevel(999,9),id,actionList,cList,action,tempHtmlValue
		dim nLevelEQ			'等级EQ
		dim isSetAddUrlEnc:isSetAddUrlEnc=false		'处理图片无加密
		 
		'处理图片无加密
		if instr(htmlValue,"setaddurlenc(*)")>0 then
			htmlValue=replace(htmlValue,"setaddurlenc(*)","")
			isSetAddUrlEnc=true
		end if
		'截取索引位置
		nLevelEQ=getStrCut(labelNameList,":leveleq(",")",0)
		if nLevelEQ<>"" then
			labelNameList=replace(labelNameList,":leveleq("& nLevelEQ &")","")
			nLevelEQ=handleNumber(nLevelEQ)
			if nLevelEQ="" then
				nLevelEQ=-1
			else
				nLevelEQ=cint(nLevelEQ)
			end if
		else
			nLevelEQ=-1
		end if 
		nCount=-1 
		'动作
		actionList=getStrCut(labelNameList,"[","]",0)
		if actionList<>"" then
			labelNameList=replace(labelNameList,"["& actionList &"]","")
		end if
		'截取索引位置
		nEQ=getStrCut(labelNameList,":eq(",")",0)
		if nEQ<>"" then
			labelNameList=replace(labelNameList,":eq("& nEQ &")","")
			nEQ=handleNumber(nEQ)
			if nEQ="" then
				nEQ=-1
			else
				nEQ=cint(nEQ)
			end if
		else
			nEQ=-1
		end if 
		if sType="check" then
			handleLabelContent=false
		else
			handleLabelContent=""
		end if
		for i =1 to nLabelIndex
			nLevel=labelLevelName(i,0)
			if nLevel<0 then
				nLevel=0
			end if
			labelName=labelLevelName(i,1)
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			labelS=labelLevelName(i,5)  : lCaseLabelS=labelS
			labelBeforeStr=labelLevelName(i,6) 	'标签前面内容
			labelAfterStr=labelLevelName(i,7) 		'标签后面内容
			parentNRight=labelLevelName(i-1,4)
			action="|" & labelLevelName(i,9) & "|"
			'在替换内容时，这里处理不好，因为再有个搜索，那样就会很慢，矣，怎么办？
			if instr(action,"|del|delthis|")>0 then
				'为删除则不处理了
			elseif findLabelLevel(nLevel,4)="OK" then
				'call echoRedB(nLevel,findLabelLevel(nLevel,0))
				'call echoYellowB(labelName , "/" & findLabelLevel(nLevel,1))
				'call echo(nCount,nEQ)
				if nLevel=findLabelLevel(nLevel,0) and ( labelName="/" & findLabelLevel(nLevel,1) or labelNameList="*") then					 
					s=mid(pubContent,findLabelLevel(nLevel,2),nLeft-findLabelLevel(nLevel,2))
					id=findLabelLevel(nLevel,5)	 
					if sType="1" then
						c=LeftS & s & labelS
					elseif sType="label" then
						c=LeftS
					elseif sType="set" or sType="del" or sType="delthis" then		 
						if sType="set"then
							labelLevelName(id,9)="replace"
							labelLevelName(id,7)=htmlValue 
						elseif sType="del" or sType="delthis" then
							labelLevelName(id,9)="del" 
							labelLevelName(i,9)="del"
							if sType="del"then 
								labelLevelName(id,6)="" 
								labelLevelName(id,7)="" 
								labelLevelName(i,6)="" 
								labelLevelName(i,7)=""
							end if 
							'设置子标签为删除动作 开始位置+1是因为排除自身   结束位置-1是排除自身 
							call setLabelAction(id+1,i-1,"del") 
						end if
					
					'before() - 在被选元素之前插入
					elseif sType="before" then
						id=findLabelLevel(nLevel,5)
						labelLevelName(id,6)=htmlValue
						
					'after() - 在被选元素之后插入
					elseif sType="after" then
						labelLevelName(i,7)=htmlValue
						
					'prepend() - 在被选元素的开头插入 
					elseif sType="prepend" then
						id=findLabelLevel(nLevel,5)
						labelLevelName(id,7)=htmlValue
					 
					'append() - 在被选元素的结尾插入
					elseif sType="append" then
						labelLevelName(i,6)=htmlValue 
					
					'检测
					elseif sType="check" then
						c=true
					elseif sType="get+label" then
						c= findLabelLevel(nLevel,3) &  s & labelS
						
					elseif sType="handleimg"  then 
						
						s2=mid(pubContent,parentNRight,nLeft-parentNRight) '可以加src了，因为我处理好了
						c=pubCssObj.handleCssContent("css",s2,"*","background||background-image||src","*",htmlValue,"images")					
						if instr(vbcrlf & imgList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
							if imgList<>"" then
								imgList=imgList & vbcrlf
							end if
							imgList=imgList & c
						end if
						tempHtmlValue=htmlValue
						if isSetAddUrlEnc=true then
							tempHtmlValue=tempHtmlValue & "setaddurlenc(*)"
						end if
						c= pubCssObj.handleCssContent("css",s2,"*","background||background-image||src","*",tempHtmlValue ,"set")	  
						labelLevelName(id,9)="replace" 
						labelLevelName(id,7)=c 
					elseif sType="html"  then 			'标签内容
						c=s
					else
						c=s
						'call echo("找到","")
					end if
					'为数组
					if cList<>"" then
						cList=cList & "$Array$"
					end if
					cList=cList & c
					handleLabelContent=cList
					findLabelLevel(nLevel,4)=""
					'为指定一条
					if nEQ<>-1 or sType="check" then
						exit function  
					end if
				end if 
				
				
			elseif   (instr(labelNameList,labelName)>0 or labelNameList="*") then 
				'call echo(replace(labelS,"<","&lt;"), pubHtmlObj.handleHtmlLabel(labelS,labelNameList,paramName,actionList,"check","") )
				
				if sType="getimg"  then
					if labelName="img" then
						'call echo("labelS",replace(labelS,"<","&lt;"))
						c=pubHtmlObj.handleHtmlLabel(labelS,"img","src","","get",htmlValue)
						if c<>"" then
							'为数组
							if cList<>"" then
								cList=cList & vbcrlf
							end if
							cList=cList & c
							handleLabelContent=cList
						end if
					end if
				elseif sType="handleimg"  then 
					tempHtmlValue=htmlValue
					if isSetAddUrlEnc=true then
						tempHtmlValue=tempHtmlValue & "setaddurlenc(*)"
					end if 
					if labelName="img" or labelName="input"  then
						'type=application/x-shockwave-flash 为flash判断
						'加个判断是为了加快 问：有必需吗？答：待测试呀
						if instr(lCaseLabelS,"src")>0 then
							c=pubHtmlObj.handleHtmlLabel(labelS,"*","src","","get",htmlValue) 
							labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"*","src","","set",tempHtmlValue)
							if instr(vbcrlf & imgList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
								if imgList<>"" then
									imgList=imgList & vbcrlf
								end if
								imgList=imgList & c
							end if
						end if
					
					'swf
					elseif labelName="embed" then 
						if instr(lCaseLabelS,"src")>0 then
							c=pubHtmlObj.handleHtmlLabel(labelS,"*","src","","get",htmlValue)
							labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"*","src","","set",tempHtmlValue)
							if instr(vbcrlf & swfList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
								if swfList<>"" then
									swfList=swfList & vbcrlf
								end if
								swfList=swfList & c
							end if
						end if
					'swf
					elseif labelName="param" then
						if instr(lCaseLabelS,"value")>0 then
							c=pubHtmlObj.handleHtmlLabel(labelS,"*","value","name=movie","get",htmlValue)
							labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"*","value","name=movie","set",tempHtmlValue)
							if instr(vbcrlf & swfList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
								if swfList<>"" then
									swfList=swfList & vbcrlf
								end if
								swfList=swfList & c
							end if 
						end if  
						
						
					elseif labelName="link" then  
						c=pubHtmlObj.handleHtmlLabel(labelS,"link","href","rel=shortcut icon, ||type=image/ico","get",htmlValue) 
						labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"link","href","rel=shortcut icon, ||type=image/ico","set",tempHtmlValue)
						'call echo("link",replace(labelLevelName(i,5),"<","&lt;"))
						if instr(vbcrlf & imgList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
							if imgList<>"" then
								imgList=imgList & vbcrlf
							end if
							imgList=imgList & c
						end if
						'css  第二次处理  不能用labelS 因为它在上面有可能处理了  明白？
						c=pubHtmlObj.handleHtmlLabel(labelS,"link","href","rel=stylesheet","get",htmlValue)
						labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelLevelName(i,5),"link","href","rel=stylesheet","set",tempHtmlValue) 
						if instr(vbcrlf & cssList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
							if cssList<>"" then
								cssList=cssList & vbcrlf
							end if
							cssList=cssList & c
						end if 
					elseif labelName="script" then
						if instr(lCaseLabelS,"src")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"script","src","","get",htmlValue)
							'说明不存在这个src标签属性
							if c<>"" then 
								labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"script","src","","set",tempHtmlValue)
								if instr(vbcrlf & jsList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
									if jsList<>"" then
										jsList=jsList & vbcrlf
									end if
									jsList=jsList & c
								end if
							end if
						'js写在内部
						else
							'call eerr("555555555555",labelLevelName(i+1,5))
							'call eerr("显示1111111",mid(pubContent,labelLevelName(i,4),labelLevelName(i+1,4))) 
						end if
					'追加于20171013
					elseif labelName="source" then
						if instr(lCaseLabelS,"src")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"source","src","","get",htmlValue)
							'说明不存在这个src标签属性
							if c<>"" then 
								labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"source","src","","set",tempHtmlValue)
								if instr(vbcrlf & sourceList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
									if sourceList<>"" then
										sourceList=sourceList & vbcrlf
									end if
									sourceList=sourceList & c
								end if
							end if
						end if
						
					'追加于20171019
					elseif labelName="video" then
						if instr(lCaseLabelS,"src")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"video","src","","get",htmlValue)
							'说明不存在这个src标签属性
							if c<>"" then 
								labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"video","src","","set",tempHtmlValue)
								if instr(vbcrlf & videoList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
									if videoList<>"" then
										videoList=videoList & vbcrlf
									end if
									videoList=videoList & c
								end if
							end if
						end if
					'追加于20171207
					elseif labelName="a" then
						if instr(lCaseLabelS,"href")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"a","href","","get",htmlValue) 
							'说明不存在这个src标签属性
							if c<>"" then  
								if instr(vbcrlf & aList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
									if aList<>"" then
										aList=aList & vbcrlf
									end if
									aList=aList & c
								end if
							end if
						end if
					elseif labelName="style" then  
						findLabelLevel(nLevel,0)=nLevel
						findLabelLevel(nLevel,1)=labelName
						findLabelLevel(nLevel,2)=nRight
						findLabelLevel(nLevel,3)=labelS
						nCount=nCount+1					'搜索到的值累加  
						findLabelLevel(nLevel,4)="OK" 
						findLabelLevel(nLevel,5)=i			'索引
					end if 
					'这里加个style判断，非常有必要的，真的
					if instr(lCaseLabelS,"style")>0 then
						c=pubHtmlObj.handleHtmlLabel(labelS,"*","style","style>background","get",htmlValue)
						labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelLevelName(i,5),"*","style","style>background","set",tempHtmlValue)
						if instr(vbcrlf & imgList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
							if imgList<>"" then
								imgList=imgList & vbcrlf
							end if
							imgList=imgList & c
						end if
					end if
					
					
				'判断 重要
				elseif pubHtmlObj.handleHtmlLabel(labelS,labelNameList,paramName,actionList,"check","")=true then 
					'call echo("labelName",labelName)
					
					findLabelLevel(nLevel,0)=nLevel
					findLabelLevel(nLevel,1)=labelName
					findLabelLevel(nLevel,2)=nRight
					findLabelLevel(nLevel,3)=labelS
					nCount=nCount+1					'搜索到的值累加 
					if (nCount=nEQ or nEQ=-1) and (nLevelEQ=nLevel or nLevelEQ=-1) then		'追加了等级判断
						findLabelLevel(nLevel,4)="OK"
					end if 
					findLabelLevel(nLevel,5)=i			'索引
					
					'单标签，暂时只处理删除
					if findLabelLevel(nLevel,4)="OK" and instr("|img|meta|link|param|hr|br|", "|" & labelName & "|")>0 then
						if sType="del" then		   
							labelLevelName(i,9)="del"
						
						'对单标签值获得  暂时只加获得
						elseif sType="get" then 
							'为数组
							if cList<>"" then
								cList=cList & "$Array$"
							end if
							cList=cList & labelS
							handleLabelContent=cList 
						end if 
						findLabelLevel(nLevel,4)=""
					end if 
				end if
				
			end if
			
		next
		'总数
		if sType="count" then
			handleLabelContent=nCount
		end if
	end function
	'设置指定数组标签动作
	function setLabelAction(nStart,nEnd,actionStr)
		dim i
		for i = nStart to nEnd
			labelLevelName(i,9)=actionStr							
		next						
	end function 
	
	'获得图片/css/js地址列表
	function getUrlList()
		dim splstr,s,c
		urlList=""
		splstr=split(imgList & vbcrlf & cssList & vbcrlf & jsList,vbcrlf & sourceList & vbcrlf  & videoList & vbcrlf )
		for each s in splstr
			if s<>"" then
				if instr(vbcrlf & urlList & vbcrlf, vbcrlf & s & vbcrlf)=false then
					if urlList<>"" then
						urlList=urlList & vbcrlf
					end if
					urlList=urlList & s 
				end if 
			end if
		next
		
		getUrlList=urlList
	end function
	'清除数组
	function clearArray() 
		dim i
		nLabelIndex=0
		for i = 0 to 9999
			if labelLevelName(i,1)="" then
				exit for
			end if
			labelLevelName(i,0)=""
			labelLevelName(i,1)=""
			labelLevelName(i,2)=""
			labelLevelName(i,3)=""
			labelLevelName(i,4)=""
			labelLevelName(i,5)=""
			labelLevelName(i,6)=""
			labelLevelName(i,7)=""
			labelLevelName(i,8)=""
			labelLevelName(i,9)=""
			labelLevelName(i,10)="" 
		next 
	end function
	'更新
	function updateHtml()
		dim c
		c=getFormattingHtml("","","")
		call clearArray()
		c=handleFormatting(c)
		updateHtml=c 
	end function
	
	
	'删除HTML，待写，这种方法不太好
	function del()
		dim i,labelName,nLevel,nLeft,nLen,nRight,s,c  
		for i =1 to nLabelIndex
			nLevel=labelLevelName(i,0)
			labelName=labelLevelName(i,1)
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			if nLevel="" then
				exit for
			end if	
			s=mid(pubContent,nLeft,nLen) 
			call echo(pubHtmlObj.handleHtmlLabel(s,"class","aa","","check",""),replace(s,"<","&lt;"))
	 
		next 
	end function
	 
	'处理上一个字符与下一个字符
	sub handleNRowCountAndBeforeStrAndAfterStr(content,s,i)
		if s=chr(10)   then
			nRowCount=nRowCount+1
		end if
		beforeStr=""
		if i>1 then
			beforeStr = mid(content, i-1, 1)                        '上一个字符
		end if
		afterStr =""
		if i+1<len(content) then
			afterStr = mid(content, i+1, 1)       				'下一个字符  
		end if
	end sub
	
	
	
	'*******************************************************************************
	'找到指定位置的标签
	function findLabelContent(findNLevel)
		dim i,labelName,parentLabelName,nLevel,nLeft,nLen,nRight,s,c,labelBeforeStr,labelAfterStr,parentNRight,labelS,s2,labelNRow
		dim action,parentAction,labelListArray(99,3),j,nBigLabelArray,bigLabelArrayName
		for i =1 to nLabelIndex
			parentLabelName=labelName
			nLevel=labelLevelName(i,0)
			labelName=labelLevelName(i,1)
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			labelS=labelLevelName(i,5)
			labelBeforeStr=labelLevelName(i,6) 	'标签前面内容
			labelAfterStr=labelLevelName(i,7) 		'标签后面内容
			labelNRow=labelLevelName(i,8)			'标签内容行数
			action = "|" & labelLevelName(i,9) & "|" 		'当前动作  
			parentAction = "|" & labelLevelName(i-1,9) & "|" 		'上一个标签动作  
			parentNRight=labelLevelName(i-1,4) 						'上一个右边字符位置
			
			if nLevel=findNLevel and left(labelName,1)="/" then
				c=c & labelName & "|"
				for j=0 to ubound(labelListArray)
					if labelListArray(j,0)="" then
						labelListArray(j,0)=labelName
						labelListArray(j,1)=1
					elseif labelListArray(j,0)=labelName then
						labelListArray(j,1)=labelListArray(j,1)+1
					end if
				next
			end if
		next
		nBigLabelArray=0
		for j=0 to ubound(labelListArray)
			if labelListArray(j,0)="" then
				exit for
			end if
			if labelListArray(j,1)>nBigLabelArray then
				nBigLabelArray=labelListArray(j,1)
				bigLabelArrayName= labelListArray(j,0)
			end if 
		next
		bigLabelArrayName=mid(bigLabelArrayName,2)
		findLabelContent=bigLabelArrayName
	end function
end class


%>