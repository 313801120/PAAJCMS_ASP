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

'dim t:  set t=new class_css
'call echo("",t.id) 
class class_css 
	dim isDisplayEcho

	 '构造函数 初始化
    Sub Class_Initialize() 
		isDisplayEcho=false
	end sub
    '析构函数 类终止
    Sub Class_Terminate() 
    End Sub 
	 
	'获得CSS内容里URL列表
	function getCssContentUrlList(content)
		getCssContentUrlList=handleCssContent("css",content ,"*","background||background-image","trim(*)*","","images")
	end function 
	'get,set,images,error,errorcount,labellist,paramlist
	'这么变态的操作CSS函数，也能写出来，可想我有多闲，时间多的实在太富裕了，可自己经济穷得可怜。一个爱好编程小子
	'处理CSS样式  太复杂了，一定要写好注释   默认获得内容 image获得背景图片列表  排除了注释里内容  <space>  为清空
	function handleCssContent(htmlOrCssOrStyle,byval content,byval sCssLabelName,byval sCssParamName,byval findSCssParamValue,byval sCssParamValue,byval sType)
		dim splstr,i,s,c,rowC,nRowCount,saveRowC,splxx,isCssStyle,endCode,nLen
		dim cssLabelName,tempCssLabelName,cssParamName,tempCssParamName
		dim cssLabelValue,cssParamValue,yunCssParamValue,tempCssParamValue,noteC
		dim isCssLabel,isCssParam,isNote,isMaoHao,url,urlList
		dim beforeStr,afterStr,nCountLen,isKuHu
		dim findLabel,findParam,findValue,isLabelOK,isParamOK,isValueOK,str1,str2
		dim isLabelNameLCase,isFindValueTrim,isFindValueLCase,isSetValueTrim,isSetValueLCase
		dim cssErrorList,cssErrorCount,errS				'css错误列表与错误总数
		dim labelList,paramList,isParamList			'样式名称列表，参数名称列表
		dim isSCssParamValueImg,isImg
		dim isUrlEnc:isUrlEnc=false						'url加密   如"[##"& url &"##]"
		isSCssParamValueImg=false
		dim isParamValueFullForceUrl:isParamValueFullForceUrl=false					'强制网址完整使用当前网址 
		dim s2,c2
		dim splCssParamValue,nCssImgIndex		'以 src(1.jpg),src(1.jpg) 处理
		
		sType="|"& sType &"|"			'操作类型
		
		'让获得url 强制完整
		if instr(sCssParamValue,"fullforceurl(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"fullforceurl(*)","") 
			isParamValueFullForceUrl=true 
		end if
		 
		'判断是否为图片判断
		if instr(sCssParamValue,"setaddurlenc(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"setaddurlenc(*)","") 
			isUrlEnc=true
		end if
		
		'判断是否为图片判断
		if instr(sCssParamValue,"isimg(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"isimg(*)","") 
			isSCssParamValueImg=true
		end if
 
		'检测内容是否一样，这个有用吗        没用，待改进吧
        if instr(sCssParamValue, "[#check#]") > 0 then  
            splxx = split(sCssParamValue, "[#check#]") 
            findSCssParamValue = splxx(0) 
            sCssParamValue = splxx(1)  
        end if   
		
		isLabelNameLCase=false		'标签名转成小写 
		isFindValueTrim=false		'参数值清除两边空格
		isFindValueLCase=false		'参数值转成小写 
		isSetValueTrim=false		'获得值清除两边空格
		isSetValueLCase=false		'获得值转小写
		'清除获得参数值 去除两边空格   让获得样式名称转小写  不太有啥用的
		if instr(sCssLabelName,"lcase(*)")>0 then
			sCssLabelName=replace(sCssLabelName,"lcase(*)","")
			isLabelNameLCase=true
		end if
		
		'获得参数值 处理成小写
		if instr(findSCssParamValue,"lcase(*)")>0 then
			findSCssParamValue=replace(findSCssParamValue,"lcase(*)","")
			isFindValueLCase=true
		end if 
		'清除获得参数值 去除两边空格
		if instr(findSCssParamValue,"trim(*)")>0 then
			findSCssParamValue=replace(findSCssParamValue,"trim(*)","")
			isFindValueTrim=true
		end if
		
		'获得设置值 和赋值 处理成小写
		if instr(sCssParamValue,"lcase(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"lcase(*)","")
			sCssParamValue=lcase(sCssParamValue)
			isSetValueLCase=true
		end if 
		'清除获得设置值 和赋值 去除两边空格
		if instr(sCssParamValue,"trim(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"trim(*)","")
			sCssParamValue=phptrim(sCssParamValue)						'对赋值处理，也对获得值处理，只是这里可以在这里处理，快
			isSetValueTrim=true
		end if

		dim isSCssParamValueFullUrl,isSCssParamValueReplaceDir
		isSCssParamValueFullUrl=false
		isSCssParamValueReplaceDir=false
		'让获得url 完整
		if instr(sCssParamValue,"fullurl(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"fullurl(*)","") 
			isSCssParamValueFullUrl=true
		end if
		'让获得换本地目录
		if instr(sCssParamValue,"dir(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"dir(*)","") 
			isSCssParamValueReplaceDir=true
		end if
		
		'样式参数名 都为小写
		sCssParamName=lcase(sCssParamName)
		
		
		cssErrorCount=0				'css错误总数
		isParamList=false			'收集参数列表
		
		isCssLabel=false	'样式名为假
		isCssParam=false	'样式参数为假
		isNote=false		'注释为假
		isMaoHao=false			'是否为冒号
		isCssStyle=true				'默认为CSS
		isKuHu=false			'()括弧
		'对<div style='color:red'>  对div里style处理 因为它没有{，默认让{为真 
		if lcase(htmlOrCssOrStyle)="style" then
			isCssLabel=true 
			content=content & " "			'后面必需加个空格  才能获得值，有Bug以后再改了，现在太累了20161010
		elseif lcase(htmlOrCssOrStyle)="html" then
			isCssStyle=false
		end if
		nRowCount=1					'行总数默认清零
		nCountLen=len(content)
		for i = 1 to nCountLen
			s = mid(content, i, 1)
			if s=chr(10)   then
				nRowCount=nRowCount+1
			end if
			beforeStr=""
			if i>1 then
				beforeStr = mid(content, i-1, 1)                        '上一个字符
			end if
			afterStr = mid(content & " ", i+1, 1)       				'下一个字符 
			if isCssStyle=false then
				if s="<" then
					endCode=lcase(mid(content,i))
					if left(endCode,7)="<style>" or left(endCode,7)="<style " then 
						nLen=instr(endCode,">")
						c=c & mid(content,i,nLen)
						i=i+nLen 
						isCssStyle=true			'CSS样式为真
					else
						c=c & s
					end if
				elseif s<>"}" then
					c=c & s
				end if
			elseif isCssStyle=true and lcase(mid(content,i,8))="</style>" then
				isCssStyle=false			'CSS样式为假
				c=c & s
			'注释开始
			elseif s="/" and afterStr="*"  then
				isNote=true									'开启注释
				c=c & rowC
				rowC=""
				noteC=s										'注释赋值
				s=""										'清空s
			'注释运行  与结束
			elseif isNote=true then
				noteC=noteC & s								'注释累加
				if s="/" and beforeStr="*" then				'注释结果
					saveRowC=noteC							'保存赋值 
					noteC=""								'清空注释
					if instr(sType,"|delnote|")>0 then
						c=phptrim(c)
					 	saveRowC=""
					end if
					isNote=false 							'关闭注释
				end if 						
				s=""										'清空s
			'css开始标签
			elseif s="{" then 
				cssLabelName=phptrim(rowC)				'清空两边空格与TAB
				tempCssLabelName=lcase(cssLabelName)	'暂时CSS标签名转小写   
				saveRowC = rowC							'即时输出
				if instr(sType,"|zip|")>0 then
					saveRowC=phpTrim(saveRowC)
				elseif instr(sType,"|unzip|")>0 then
					saveRowC=vbcrlf & phpTrim(saveRowC)
				end if
				isCssLabel=true							'开启CSS标签
			'css结果标签
			elseif s="}" then
				if isCssParam=true then					'参数值为真
					cssParamValue=rowC					'记录上面值，一般在;里操作
				end if
				saveRowC = rowC 						'即时输出
				if instr(sType,"|zip|")>0 then
					saveRowC=phpTrim(saveRowC) 
				elseif instr(sType,"|unzip|")>0 then
					saveRowC=saveRowC & vbcrlf 
				end if 
				isCssLabel=false						'关闭CSS标签
				isCssParam=false						'关闭样式参数
				isMaoHao=false							'冒号关闭
				'当参数列表为真，并没有保存内容了，则退出
				if isParamList=true and phptrim(saveRowC)="" then
					exit for
				end if
			elseif s="(" then
				isKuHu=true
				rowC=rowC & s
			elseif s=")" then
				isKuHu=false
				rowC=rowC & s
			elseif isKuHu=true then 
				rowC=rowC & s
			'参数开始标签
			elseif isCssLabel=true and  s=":"  then
				if isMaoHao=false then
					cssParamName=rowC						'参数名称
					saveRowC = rowC							'即时输出
					if instr(sType,"|zip|")>0 then
						saveRowC=phpTrim(saveRowC) 
					'解密处理
					elseif instr(sType,"|unzip|")>0 then
						saveRowC= vbcrlf & "    " & phpTrim(saveRowC) 
					end if
					isCssParam=true							'开启参数
					isMaoHao=true							'冒号开启
					'call echo("cssParamName",cssParamName)
				else
					'rowC=rowC & s 							'这样注意，需要累加吗，如color::::::red;      暂时不需要，但是保留
					cssErrorCount=cssErrorCount+1			'找到一处错误
					errS=nRowCount & "行，多余:号  " & phptrim(cssLabelName) & ">>" & phptrim(cssParamName) & vbcrlf
					if instr(vbcrlf & cssErrorList & vbcrlf, vbcrlf & errS & vbcrlf)=false then
						cssErrorList=cssErrorList & errS
					end if
				end if
			'参数结束标签
			elseif isCssLabel=true and ( s=";" or i=nCountLen  ) and left(phptrim(lcase(rowc)),9)<>"@charset " then		'这里需要用i=nCountLen嘛,当时怎么想的	加个 @charset 排除，暂时20171003
				
				'判断当前 ; 的前一个字符不等于 ;20171028
				if s=";" and right(phptrim(mid(content,1,i-1)),1)<>";" then 
					 
					cssParamValue=rowC 						'获得参数值
					saveRowC = rowC							'即时输出
					isCssParam=false						'关闭参数
					isMaoHao=true							'冒号关闭
				end if
			else
				rowC=rowC & s							'累加字符
				saveRowC=""								'清空即时输入
			end if 
			'处理保存
			if saveRowC<>"" then   
				 if (s=";" or s="}" or i=nCountLen ) and cssParamName<>"" or (s="}" and instr(sType,"|labellist|")>0 ) then 		'当在最后则处理参数
					tempCssParamName=lcase(phptrim(cssParamName))					'处理成小写和清除两边空格
					tempCssParamValue=replace(replace(lcase(phptrim(cssParamValue))," ",""),vbtab,"")	'参数清空
					
					yunCssParamValue=cssParamValue
					'背景图片处理
					if (tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src") and instr(tempCssParamValue,"url(")>0 then 
						'call echo("cssParamValue1",cssParamValue)
						cssParamValue=batchGetKuHuoValue(cssParamValue)								'参数名称转小写与除两边空格
						'call echo("cssParamValue2",cssParamValue)
					end if
					
					isLabelOK=false
					isParamOK=false
					isValueOK=false
					findLabel=getStrCut(sCssLabelName,"find(",")",2)
					findParam=getStrCut(sCssParamName,"find(",")",2)
					findValue=getStrCut(findSCssParamValue,"find(",")",2)
					'css样式名
					if isLabelNameLCase=true then
						cssLabelName=lcase(cssLabelName)
					end if 
					if findLabel<>"" then
						isLabelOK=IIF(instr(cssLabelName,findLabel)>0,true,false)
					elseif instr(sCssLabelName,"||")>0 then 
						splxx=split(sCssLabelName,"||")
						for each str2 in splxx  
							if str2=cssLabelName then
								isLabelOK=true 
								exit for
							end if
						next
					elseif sCssLabelName=cssLabelName or sCssLabelName="*" then
						isLabelOK=true
					end if
					'css样式参数名
					if findParam<>"" then
						isParamOK=IIF(instr(tempCssParamName,findParam)>0,true,false)
					elseif instr(sCssParamName,"||")>0 then 
						splxx=split(sCssParamName,"||")
						for each str2 in splxx 
							if str2=tempCssParamName then
								isParamOK=true 
								exit for
							end if
						next											
					elseif sCssParamName=tempCssParamName or sCssParamName="*" then
						isParamOK=true
					end if
					'css样式参数值
					if isFindValueTrim=true then
						cssParamValue=phptrim(cssParamValue)
					end if 
					if isFindValueLCase=true then
						cssParamValue=lCase(cssParamValue)
					end if 
					if findValue<>"" then
						isValueOK=IIF(instr(cssParamValue,findValue)>0,true,false)
					elseif instr(findSCssParamValue,"||")>0 then 
						splxx=split(findSCssParamValue,"||")
						for each str2 in splxx  
							if str2=cssParamValue then
								isValueOK=true 
								exit for
							end if
						next
					elseif findSCssParamValue=cssParamValue or findSCssParamValue="*" then
						isValueOK=true
					end if  
					if isDisplayEcho=true then
						call echo("sCssLabelName",sCssLabelName)
						call echo("sCssParamName",sCssParamName)
						call echo("findSCssParamValue",findSCssParamValue)
						call echo("findLabel",findLabel)
						call echo("findParam",findParam)
						call echo("findValue",findValue) 
						call hr() 
					end if
					'清除两边空格
					if isSetValueTrim=true then
						cssParamValue=phptrim(cssParamValue)
					end if
					if isSetValueLCase=true then
						cssParamValue=lcase(cssParamValue)
					end if  
					isImg=true
					'是否为图片地址
					if isSCssParamValueImg=true then
						if phptrim(cssParamValue)="#" or phptrim(cssParamValue)="/" or phptrim(cssParamValue)="\" or lcase(left(phptrim(cssParamValue),11))="data:image/" then
							isImg=false
						end if
					end if 
					 
					'网址完整
					if isSCssParamValueFullUrl=true or isParamValueFullForceUrl=true then 
						'排除这种不对它设置20170824
						if left(phptrim(cssParamValue),5)<>"rgba("  and left(phptrim(cssParamValue),8)<>"-webkit-" and left(phptrim(cssParamValue),6)<>"alpha(" and left(phptrim(cssParamValue),11)<>"data:image/" then
							'清除两边空格
							if isSetValueTrim=true then
								cssParamValue=phptrim(cssParamValue)
							end if
						
						
							'强制使用当前网址
							if isParamValueFullForceUrl=true and cssParamValue<>"" then  
								cssParamValue=getFileAttr(cssParamValue,"name") 
							end if
						
							cssParamValue=batchFullHttpUrl(sCssParamValue,cssParamValue)		'批量获得，加于20171124对，处理
						else
							isParamOK=false				'加个这个不处理，让错误吵出错
						end if
					'替换目录
					elseif isSCssParamValueReplaceDir=true then					
						'排除这种不对它设置20170824   instr(phptrim(cssParamValue),"url(")以后用这个
						if left(phptrim(cssParamValue),5)<>"rgba(" and left(phptrim(cssParamValue),8)<>"-webkit-" and left(phptrim(cssParamValue),6)<>"alpha(" and left(phptrim(cssParamValue),11)<>"data:image/" then
							 
							'批量处理，对url(1.jpg),url(2.jpg)这种 20171211
							splCssParamValue=split(cssParamValue,vbcrlf) 
							cssParamValue=""
							for each s2 in splCssParamValue 
								if instr(s2,"?")>0 then
									s2=mid(s2,1,instr(s2,"?")-1)
								end if 
								str1= getFileAttr(s2,"2")
								'清除两边空格
								if isSetValueTrim=true then
									str1=phptrim(str1)
								end if
								s2=sCssParamValue & str1
								if cssParamValue<>"" then cssParamValue=cssParamValue & vbcrlf
								cssParamValue=cssParamValue & s2
							next
						else
							isParamOK=false				'加个这个不处理，让错误吵出错
						end if
					end if  
					c2=""		'不回这个就会把上一级的图片加密累加到下一级图片里20171203
					'URL加密 如 "[##"& url &"##]"   20161017
					if isUrlEnc=true and cssParamValue<>"" then 
						'call echo("cssParamValue",cssParamValue)
						splxx=split(cssParamValue,vbcrlf)
						for each s2 in splxx
							if s2<>"" then
								s2=hRFileName(s2)
								if c2<>"" then c2=c2 &vbcrlf
								c2=c2&s2
							end if
						next
						cssParamValue=c2
						'call echo("cssParamValue2",cssParamValue) 
					end if
					
					'设置值
					if instr(sType,"|set|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then  
							'对背景图片特殊处理
							if tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src"  then		'src加于20171124			
								'background: red;  排除这种
								if instr(saveRowC,"(")>0 then
									if isImg=true then 
										'call echo("yunCssParamValue",yunCssParamValue)
										splxx=split(yunCssParamValue,",")		'处理字体用的20171124
										splCssParamValue=split(cssParamValue,vbcrlf)
										nCssImgIndex=-1
										for each s2 in splxx
											'call echo("s2",s2)
											str1=getStrCut(s2,"(",")",1)
											nCssImgIndex=nCssImgIndex+1 
											'处理当前值
											if isParamValueFullForceUrl=true or isSCssParamValueFullUrl=true or isSCssParamValueReplaceDir=true then
												'call echo(saveRowC,str1)
												'call echo(nCssImgIndex,splCssParamValue(nCssImgIndex))
												'call hr()
												'call echo(nCssImgIndex,splCssParamValue(nCssImgIndex))
												'call echo(yunCssParamValue,saveRowC)
												if nCssImgIndex<=ubound(splCssParamValue) then
													saveRowC=replace(saveRowC,str1,"("& splCssParamValue(nCssImgIndex) &")") 'cssParamValue 区分下面，这是多级20171211						 
												end if
												'call echo(saveRowC & " =>> " & str1,splCssParamValue(nCssImgIndex)) 
											'替换值
											else 
												saveRowC=replace(yunCssParamValue,str1,"("& sCssParamValue &")") 
												'call echo("替换",str1)
											end if
										next
									else
										'saveRowC="url("& yunCssParamValue &")" 
									end if
								end if 
							else 
								saveRowC=sCssParamValue
							end if 
						end if 
					'获得值
					elseif instr(sType,"|getimage|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then 
				 
							handleCssContent=cssParamValue
							
							exit function 
						end if
					elseif instr(sType,"|get|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then 
			 
							handleCssContent=cssParamValue				'yunCssParamValue不能用这个，因为它没有处理过
							exit function 
						end if
					'样式名称列表 
					elseif instr(sType,"|labellist|")>0 then
						if isLabelOK=true and  s="}" then
							labelList=labelList & phptrim(cssLabelName) & vbcrlf
						end if
					'样式名称参数列表
					elseif instr(sType,"|paramlist|")>0 then
						if isLabelOK=true and isParamOK=true then
							paramList=paramList & tempCssParamName & vbcrlf
							isParamList=true
							if s="}" then
								exit for
							end if
						end if
					'获得图片列表
					elseif (instr(sType,"|images|")>0 or instr(sType,"|allimages|")>0) and (tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src") and instr(tempCssParamValue,"(")>0 then		'判断是否为背景图片 
						if isLabelOK=true and isParamOK=true and isValueOK=true and isImg=true then 
							url = cssParamValue 
							'排除 rgba(  这种 20170824
							if url<>"" and ( (instr(sType,"|images|")>0 and left(phptrim(url),11)<>"data:image/"   and left(phptrim(url),5)<>"rgba("and left(phptrim(url),8)<>"-webkit-") or instr(sType,"|allimages|")>0) then  
								if urlList<>"" then
									urlList=urlList & vbcrlf
								end if
								urlList=urlList & url
							end if
						end if
					end if
					cssParamValue=""
					cssParamName=""										'清空参数名称
					tempCssParamName=""
					isMaoHao=false							'冒号开启
				 end if
				'call echo("s",s)
				c=c & saveRowC & s						'累加最后输出内容
				saveRowC=""								'清空保存值
				rowC="" 								'清空当前值
			elseif s="}" then
				if instr(sType,"|unzip|")>0 then
					s= vbcrlf & s & vbcrlf
				end if
				c=c & s								'解决  width:12p;}  这种，因为 }前面没有值，这个}就输出不出来，所以要这里要判断下20161002
			end if
			 
		next
		c=c & saveRowC
		handleCssContent=c
		if instr(sType,"|images|")>0 or instr(sType,"|allimages|")>0 then 
			handleCssContent=urlList
		elseif instr(sType,"|error|")>0 then
			handleCssContent=cssErrorList
		elseif instr(sType,"|errorcount|")>0 then
			handleCssContent=cssErrorCount
		elseif instr(sType,"|labellist|")>0 then
			handleCssContent=labelList
		elseif instr(sType,"|paramlist|")>0 then
			handleCssContent=paramList
		'get最后是没有获得值，则返回空
		elseif instr(sType,"|get|")>0 or instr(sType,"|getimage|")>0 then
			handleCssContent=""
		
		'暂存，删除多余行20161024
		elseif instr(sType,"|zip|")>0 or instr(sType,"|unzip|")>0 then
			for i = 1  to 120
				if instr(handleCssContent,vbcrlf & vbcrlf)=false then
					exit for
				end if
				handleCssContent=replace(handleCssContent,vbcrlf & vbcrlf, vbcrlf)
			next
		end if 
	end function
end class 


%>