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
'js 压缩与解压类型

'dim jsObj:  set jsObj=new class_js
'call rw(jsObj.handleJsContent(c))
class class_js
	dim isDisplayEcho			'是否回显  
	dim jsErrorList,jsErrorCount,httpurl
	dim imgList					'图片列表
	dim cssList					'CSS列表
	dim jsList					'JS列表		
	dim swfList					'SWF列表	
	dim sourceList				'source列表 视频 
	dim videoList				'video列表 视频 

	 '构造函数 初始化
    Sub Class_Initialize() 
		httpurl="http://www.sharembweb.com/"
		isDisplayEcho=false
	end sub
    '析构函数 类终止
    Sub Class_Terminate() 
    End Sub 
	 
	'处理JS程序 sType=source  为源码不处理的那种
	function handleJsContent(byVal content, byval sType)
		dim splStr, sx, tempSx, s, clearS, c, i, j, nSYHCount, s1, s2, tempZc, rowC, tempRow, noteStr, downRow, downWord, nc, isLoop, beforeStr, afterStr, endCode, tempEndCode ,leftStr,errS
		dim upWord, upWord2, yesOK,isJS 
		dim wc                                                                          '输入文本存储内容
		dim wcType                                                                      '输入文本类型，如 " 或 '
		dim zc                                                                          '字母文件存储内容
		dim tempS 
		dim RSpaceStr, RSpaceNumb, addToRSpaceNumb                                      '行空格与空格数,追加行数
		dim isDanHaoNote																'单行注释为真
		dim multiLineComment                                                            '多行注释
		multiLineComment = ""                                                           '多行注释默认为空
		RSpaceStr = "    " : RSpaceNumb = 0                                             '行空格值 与 空格数 设置初值
		nSYHCount = 0                                                                   '双引号默认为0
		dim isAddToSYH                                                                  '是否累加双引号
		isJS=true
		jsErrorCount=0																	'出错总数
		
		sType="|"& sType &"|"
		
		splStr = split(content, vbCrLf)                                                 '分割行
		'循环分行
		for j = 0 to uBound(splStr)
			s = splStr(j)
			s = phptrim(s)
			leftStr=left(splStr(j),len(splStr(j))-len(phpLTrim(splStr(j))))				'获得左边被清空的字符 保持原js原样
			s = replace(replace(s, chr(10), ""), chr(13), "") '奇怪为什么 s里会有 chr(10)与chr(13) 呢？
			clearS = phpTrim(s)                                                             '清除两边空格与退格
			tempS = s                                                                       '暂存S
			rowC = "" : tempRow = ""                                                        '清空每行ASP代码和暂存完整行代码
			noteStr = ""                                                                    '默认注释代码为空
			downRow = ""                                                                    '下一行代码
			downWord = ""                                                                   '下一行单词
			addToRSpaceNumb = 0                                                             '清空追加行数
			isDanHaoNote=false		'单行注释为假
			if(j + 1) <= uBound(splStr) then
				downRow = splStr(j + 1) 
				downWord = getBeforeNStr(downRow, "全部", 1) 
			end if 
			nc = ""                                                                         '数字为空
			isLoop = true                                                                   '循环字符为真
			'循环每个字符
			for i = 1 to len(s)
				sx = mid(s, i, 1) : tempSx = sx 
				beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '上一个字符
				afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '下一个字符
				endCode = mid(s, i + 1)                                                         '当前字符往后面代码 一行
				'输入文本
				if(sx = """" or sx = "'" and wcType = "") or sx = wcType or wc <> "" then
					isAddToSYH = true 
					'这是一种简单的方法，等完善(20150914)
					if isAddToSYH = true and beforeStr = "\" then	
						if len(wc) >= 1 then
							if isStrTransferred(wc) = true then  '为转义字符为真 
								isAddToSYH = false 
							end if 
						else
							isAddToSYH = false 
						end if 
					'call echo(wc,isAddToSYH)
					end if 
					if wc = "" then
						wcType = sx 
					end if 
	
					'双引号累加
					if sx = wcType and isAddToSYH = true then nSYHCount = nSYHCount + 1         '排除上一个字符为\这个转义字符(20150914)
	
	
					'判断是否"在最后
					if nSYHCount mod 2 = 0 and beforeStr <> "\" then
						if mid(s, i + 1, 1) <> wcType then
							wc = wc & sx
							'call echo(wcType,replace(replace(wc,"<","&lt;"),"\" & wcType, wcType))   '对JS里内容图片处理，等实现
							'call echo("wcType=" & wcType,wc)
							if instr("|"& sType &"|","|html|")>0 then 
								wc=handleJSContentHtml(wcType,wc)					'处理JS里的html内容 20171113
								'call echo(wc,"eeee")
							end if
							
							rowC = rowC & wc                     '行代码累加
							'call echo("wc",wc)
							nSYHCount = 0 : wc = ""              '清除
							wcType = "" 
						else
							wc = wc & sx 
						end if 
					else
						wc = wc & sx 
					end if 
	
				'多行注释
				elseIf(sx = "/" and afterStr = "*") or multiLineComment <> "" then
					noteStr = mid(s, i) 
					'call echo("多行注释",noteStr)
					if multiLineComment <> "" then multiLineComment = multiLineComment & vbCrLf 
					multiLineComment = multiLineComment & noteStr 
					if inStr(noteStr, "*/") > 0 then
						'删除注释
						if instr("|"& sType &"|","|delnote|")>0 then 
							multiLineComment=""
						end if
						rowC = rowC & multiLineComment 
						multiLineComment = "" 
					end if 
					exit for 
				'注释则退出 单选注释
				elseIf sx = "/" and afterStr = "/" then
					noteStr = mid(s, i)
					'为压缩时 单行注释加换行
					if instr("|"& sType &"|","|zip|")>0 then
						isDanHaoNote=true		'单行注释为真
					end if
					'删除注释
					if instr("|"& sType &"|","|delnote|")>0 then 
						noteStr=""
					end if
					rowC = rowC & noteStr 
					exit for 
	
				'+-*\=&   字符特殊处理
				elseIf inStr(".&=,+-*/:()><;", sx) > 0 and nc = "" then
					if zc <> "" then
						tempZc = zc 
						upWord2 = upWord                  '记录上上一个变量
						upWord = LCase(tempZc)            '记录上一个变量名称   为小写
						rowC = rowC & zc 
						zc = ""                           '清空字母内容
					end if 
					'清除如if(1=1){;;;;;;;;}   20151224
					if sx = ";" and instr("{;" , right(trim(rowC), 1)) > 0 then			'判断上一个有效字符不是;{
						jsErrorCount=jsErrorCount+1			'找到一处错误
						errS=j & "行，多余:号  " & s  & vbcrlf
						if instr(vbcrlf & jsErrorList & vbcrlf, vbcrlf & errS & vbcrlf)=false then
							jsErrorList=jsErrorList & errS
						end if
					
						sx = "" 
					end if 
					rowC = rowC & sx 
					upWord2 = upWord                                                            '记录上上一个变量
					upWord = sx 
				'字母
				elseIf checkABC(sx) = true or sx = "_" or zc <> "" then
					zc = zc & sx 
					yesOK = true 
					s1 = mid(s & " ", i + 1, 1) 
					s2 = mid(zc, 1, 1) 
					if checkABC(s1) <> true and s1 <> "_" then
						yesOK = false 
					end if 
					'允许变量后面是数字
					if checkNumber(s1) = true and checkABC(s2) = true then
						yesOK = true 
					end if 
					if yesOK = false then 
						tempZc = zc 
	
						upWord2 = upWord                  '记录上上一个变量
						upWord = LCase(tempZc)            '记录上一个变量名称   为小写
						rowC = rowC & zc 
						zc = ""                           '清空字母内容
					end if 
				'为数字
				elseIf checkNumber(sx) = true or nc <> "" then
					nc = nc & sx 
					if afterStr <> "." and checkNumber(afterStr) = false then
						rowC = rowC & nc 
						nc = "" 
					end if 
				else
					'JS或PHP程序{}处理
					if sx = "{" then
						if phptrim(rowC) <> "" then
							addToRSpaceNumb = 1 
						else
							RSpaceNumb = RSpaceNumb + 1 
						end if 
					elseIf sx = "}" then
						if RSpaceNumb > 0 then
							if phptrim(rowC) <> "" then
								'删除后台注释20151224
								tempEndCode = endCode 
								if instr(tempEndCode, "//") > 0 then
									tempEndCode = mid(tempEndCode, 1, instr(tempEndCode, "//") - 1) 
								end if 
								'修复if(a==b){}  方法
								if phptrim(tempEndCode) <> "" then
									addToRSpaceNumb = -1 
								else
									addToRSpaceNumb = 0 
								end if 
							else
								RSpaceNumb = RSpaceNumb - 1 
							end if 
						end if 
					end if 
	
					yesOK = true 
					if sx = " " and i > 1 then
						if mid(s, i - 1, 1) = " " then
							yesOK = false 
						end if 
					end if 
					if yesOK = true then
						rowC = rowC & sx 
					'call echo("sx",sx)
					end if 
				end if 
				'暂存每行内容
				tempRow = tempRow & sx 
				if isLoop = false then exit for 
			next 
			if rowC = ";}" then
				rowC = "}" 
			end if 
			if instr("|"& sType &"|","|source|")>0 then
				rowC=leftStr & rowC
			elseif rowC <> "" and RSpaceNumb > 0 then
				rowC = copyStrNumb(RSpaceStr, RSpaceNumb) & replace(rowC, vbCrLf, vbCrLf & copyStrNumb(RSpaceStr, RSpaceNumb)) '修进20150902
			
			end if 
			RSpaceNumb = RSpaceNumb + addToRSpaceNumb 
			if multiLineComment = "" then                                                   '修进20150902
				'压缩
				if instr("|"& sType &"|","|zip|")>0 then 
					rowC=phpTRim(rowC)
					c=c & rowC
					if isDanHaoNote=true then
						c=c & vbcrlf
					elseif rowC<>"" and right(rowC,1)<>"}" and right(rowC,1)<>"{" and right(rowC,1)<>";" then
						c = c & ";" 'vbCrLf
					end if 
				else					
					c = c & rowC & vbCrLf 
				end if
			end if 
		next 
		handleJsContent = c 
	end function
	
	'处理JS里的html内容 
	function handleJSContentHtml(byval wcType,byval wc)
		dim thisFormatObj:set thisFormatObj = new class_formatting                                        '初始化格式化HTML类 
		dim content
		content=mid(wc,2,len(wc)-2) 
		if wcType="""" then
			content=replace(replace(content,"\""","'"),"\'","'")
		else
			content=replace(replace(content,"\""",""""),"\'","""")
		end if
		
		if instr(content,"<")=false or instr(content,">")=false then
			handleJSContentHtml=wc:exit function
		end if
		'格式化HTML
		call thisFormatObj.handleFormatting(content) 
		call thisFormatObj.handleLabelContent("*", "*", "setaddurlenc(*)isimg(*)fullurl(*)" & httpurl, "handleimg") 
		'获得格式化的HTML
		content = thisFormatObj.getFormattingHtml("","","") 
		if wcType="""" then
			content=replace(content,"""","\""")
		else
			content=replace(content,"'","\'")
		end if
		content=wcType & content & wcType 
		
		if imgList<>"" then
			imgList=imgList & vbcrlf
		end if
		imgList=imgList & thisFormatObj.imgList
		if swfList<>"" then
			swfList=swfList & vbcrlf
		end if
		swfList=swfList & thisFormatObj.swfList
		if sourceList<>"" then
			sourceList=sourceList & vbcrlf
		end if
		sourceList=sourceList & thisFormatObj.sourceList
		if videoList<>"" then
			videoList=videoList & vbcrlf
		end if
		videoList=videoList & thisFormatObj.videoList
		if jsList<>"" then
			jsList=jsList & vbcrlf
		end if
		jsList=jsList & thisFormatObj.jsList
		if cssList<>"" then
			cssList=cssList & vbcrlf
		end if
		cssList=cssList & thisFormatObj.cssList 
		'call echo("thisFormatObj.jsList",thisFormatObj.jsList) 
		handleJSContentHtml=content
	
	end function
	
end class 


%>