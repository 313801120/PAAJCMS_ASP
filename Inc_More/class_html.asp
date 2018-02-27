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
'处理html标签
'dim jsObj:  set jsObj=new class_html
'call rw(jsObj.handleJsContent(c))
class class_html
    dim isDisplayEcho                                                               '是否回显


    '构造函数 初始化
    sub class_Initialize()
        isDisplayEcho = false 
    end sub 
    '析构函数 类终止
    sub class_Terminate()
    end sub 

    '获得标签值
    function getLabelParam(byval content, paramName)
        getLabelParam = handleHtmlLabelParam(content, paramName, "", "get") 
    end function 
    '设置标签值
    function setLabelParam(byval content, paramName, paramValue)
        setLabelParam = handleHtmlLabelParam(content, paramName, paramValue, "set") 
    end function 
    '修改标签值
    function editLabelParam(byval content, paramName, paramValue)
        editLabelParam = handleHtmlLabelParam(content, paramName, paramValue, "edit") 
    end function 
    '检测标签值
    function checkLabelParam(byval content, paramName)
        checkLabelParam = handleHtmlLabelParam(content, paramName, "", "check") 
    end function 
    '处理标签参数 20160929   
    function handleHtmlLabelParam(byval content, byval paramName, byval paramValue, byval sType) 
        dim i, tempContent, tempContent2, findStr, startStr, endStr, s,s1, s2, nLen, isTrue,splxx
        dim nLeft, nRight,isParamValueTrim,isParamValueLCase,isParamValueFullUrl,isParamValueReplaceDir,isParamValueImg,isParamValueUrl,findParamValue,isParamValueFullForceUrl,paramValueFindValue,paramValueReplaceValue
		dim isCheck					'找值相等的
		dim isUrlEnc:isUrlEnc=false						'url加密   如"[##"& url &"##]"
		dim isReplace:isReplace=false			'当替换为真时则必需要有这个标题参数才可以替换20170825
		'判断是否为图片加密判断
		if instr(paramValue,"setaddurlenc(*)")>0 then
			paramValue=replace(paramValue,"setaddurlenc(*)","") 
			isUrlEnc=true
		end if
		'判断是否为图片判断
		if instr(paramValue,"isreplace(*)")>0 then
			paramValue=replace(paramValue,"isreplace(*)","") 
			isReplace=true
		end if
		
		dim isNoEMail:isNoEMail=false				'不为邮箱
		if instr(paramValue,"isnoemail(*)")>0 then
			paramValue=replace(paramValue,"isnoemail(*)","") 
			isNoEMail=true
		end if
		dim isEMail:isEMail=false				'为邮箱
		if instr(paramValue,"isemail(*)")>0 then
			paramValue=replace(paramValue,"isemail(*)","") 
			isEMail=true
		end if
		dim isHandleEMail:isHandleEMail=false				'为邮箱
		if instr(paramValue,"ishandleemail(*)")>0 then
			paramValue=replace(paramValue,"ishandleemail(*)","") 
			isHandleEMail=true
		end if
		
		isCheck=false
		if sType="check" and paramName="*" then
			handleHtmlLabelParam=true
			exit function
		end if
		'检测内容是否一样，这个有用吗        没用，待改进吧
        if instr(paramValue, "[#check#]") > 0 then  
            splxx = split(paramValue, "[#check#]") 
            findParamValue = splxx(0) 
            paramValue = splxx(1) 
            isCheck = true  
        end if  
		'清除获得参数值 去除两边空格
		isParamValueLCase=false
		if instr(paramValue,"lcase(*)")>0 then
			paramValue=replace(paramValue,"lcase(*)","")
			paramValue=lcase(paramValue)					'如果是符值 可以免得再次处理
			isParamValueLCase=true
		end if
		'清除获得参数值 去除两边空格
		isParamValueTrim=false
		if instr(paramValue,"trim(*)")>0 then
			paramValue=replace(paramValue,"trim(*)","")
			paramValue=phptrim(paramValue)					'如果是符值 可以免得再次处理
			isParamValueTrim=true
		end if 
		'让获得url 完整
		isParamValueFullUrl=false
		if instr(paramValue,"fullurl(*)")>0 then
			paramValue=replace(paramValue,"fullurl(*)","") 
			isParamValueFullUrl=true
		end if
		'让获得url 强制完整
		isParamValueReplaceDir=false
		if instr(paramValue,"fullforceurl(*)")>0 then
			paramValue=replace(paramValue,"fullforceurl(*)","") 
			isParamValueFullForceUrl=true 
		end if 
		'让获得换本地目录
		isParamValueFullForceUrl=false
		if instr(paramValue,"dir(*)")>0 then
			paramValue=replace(paramValue,"dir(*)","") 
			isParamValueReplaceDir=true
		end if
		'判断是否为图片判断
		isParamValueImg=false
		if instr(paramValue,"isimg(*)")>0 then
			paramValue=replace(paramValue,"isimg(*)","") 
			isParamValueImg=true
		end if
		'判断是否为网址  与图片判断一样
		isParamValueUrl=false
		if instr(paramValue,"isurl(*)")>0 then
			paramValue=replace(paramValue,"isurl(*)","") 
			isParamValueUrl=true
		end if
		'判断是否替换参数内容  
		if instr(paramValue,"replace[")>0 then
			s=getStrCut(paramValue,"replace[","]",0)
			splxx=split(s,"=")
			paramValueFindValue=splxx(0):paramValueReplaceValue=splxx(1) 
			paramValue=replace(paramValue,"replace["& s &"]","")
		end if
		
		 
		
        '处理等于
        for i = 1 to 30
            content = replace(replace(content, "= ", "="), "= ", "=") 
            content = replace(replace(content, " =", "="), " =", "=") 
            if instr(content, "= ") = false and instr(content, " =") = false then
                exit for 
            end if 
        next 
        tempContent = LCase(content) 

        for i = 1 to 4
            if i = 1 then
                '双引号
                startStr = paramName & "=""" 
                endStr = """" 
            elseif i = 2 then
                '单引号
                startStr = paramName & "='" 
                endStr = "'" 
            elseif i = 3 then
                '其它
                startStr = paramName & "=" 
                endStr = " " 
            elseif i = 4 then
                '其它
                startStr = paramName & "=" 
                endStr = ">" 
            end if 

            nLeft = 0 : nRight = 0 
            isTrue = true 
            nLeft = inStr(tempContent, startStr) 
            if nLeft = 0 then
                isTrue = false 
            end if 
            if isTrue = true then
                tempContent2 = mid(tempContent, nLeft + len(startStr)) 

                nRight = inStr(tempContent2, endStr) 
                if nRight = 0 then
                    isTrue = false 
                end if 
            end if 
 
            if isTrue = true then
                findStr = mid(content, nLeft + len(startStr), nRight - len(endStr))  
				'值小写
				if isParamValueLCase=true then
					findStr=lcase(findStr)
				end if
				'值清空两边空格
				if isParamValueTrim=true then
					findStr=phptrim(findStr)
				end if
				
				
				'是否为图片地址
				if isParamValueImg=true then
					if phptrim(findStr)="#" or phptrim(findStr)="" or phptrim(findStr)="/" or phptrim(findStr)="\" or lcase(left(phptrim(findStr),11))="data:image/" then
						if sType <> "set" and sType <> "edit" then 'addto20171115
							findStr=""
						end if		 	
					end if
				end if 
				if isParamValueUrl=true or isParamValueFullForceUrl=true then 
					if phptrim(findStr)="#" or phptrim(findStr)="" or phptrim(findStr)="/" or phptrim(findStr)="\" or lcase(left(phptrim(findStr),11))="javascript:" then
						findStr=""
					end if
				end if
				'排除邮箱
				if isnoemail=true and lcase(left(phptrim(findStr),7))="mailto:" then	'追加于20171208
					findStr=""
				end if
				'提取邮箱
				if isemail=true and lcase(left(phptrim(findStr),7))<>"mailto:" then	'追加于20171208
					findStr=""
				elseif isemail=true  and isHandleEMail=true then
					findStr=mid(findStr,8)
				end if
				
				
				'call echo("findStr",findStr)
					
				'网址完整 排除base64类型(addto20171115)
				if isParamValueFullUrl=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					findStr=fullHttpUrl(paramValue,findStr)
				'强制网址完整  
				elseif isParamValueFullForceUrl=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					'call echo(paramValue,findStr)
					findStr=getFileAttr(findStr,"name")
					'call echo(paramValue,findStr)
					findStr=fullHttpUrl(paramValue,findStr) 
					
				'替换目录
				elseif isParamValueReplaceDir=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					if instr(findStr,"?")>0 then
						findStr=mid(findStr,1,instr(findStr,"?")-1)
					end if
					findStr=paramValue & getFileAttr(findStr,"2")
				end if

				'URL加密 如 "[##"& url &"##]"   20161017
				if isUrlEnc=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					findStr=hRFileName(findStr)
				end if

						
				'设置的时间 isParamValueFullForceUrl  可能用不上，暂时加上吧20170825
                if sType = "set" or sType = "edit" then 
					if (isCheck = true and findStr = findParamValue) or isCheck = false then
						if isParamValueFullUrl=false and isParamValueFullForceUrl=false and  isParamValueReplaceDir=false then 
							findStr = paramValue 
						end if
					end if
                end if 
                '清空
                if paramValue = "clear(*)" then
                    '强强强旱替换
                    nLen = inStr(tempContent, startStr) 
                    s1 = mid(content, 1, instr(tempContent, startStr) - 1) 
                    s2 = right(content, len(content) - nLen + 1 - len(startStr)) 
                    s2 = mid(s2, instr(s2, endStr) + 1) 
                    s1 = rTrim(s1)                                                                  '去右边空格
                    'call rwend(s1 & s2) 
					content=s1 & s2
				'替换20171109
				elseif paramValueFindValue<>"" then
					if left(paramValueFindValue,2)="**" then
						paramValueFindValue=mid(paramValueFindValue,3) 
						content=replaceNoULCase(content,paramValueFindValue,paramValueReplaceValue)
					else
						content=replace(content,paramValueFindValue,paramValueReplaceValue)
					end if
                else
                    '强强强旱替换
                    nLen = inStr(tempContent, startStr) - 1 + len(startStr)                         '注意用转了小写的tempContent
                    s2 = right(content, len(content) - nLen) 
                    s2 = mid(s2, inStr(s2, endStr)) 
                    s1 = left(content, nLen) 
					
					
                    content = s1 & findStr & s2 
                end if 
                exit for                                                                        '找到了退出
            end if 
        next 
		
		
		 
        if sType = "get" then 
			'call echo(isParamValueImg,findstr)
			 if (isCheck = true and findStr = findParamValue) or isCheck = false then 
			 else
			 	findStr=""
			 end if
			
            handleHtmlLabelParam = findStr 
		elseif sType = "check" then
			handleHtmlLabelParam=isTrue
        else 
			'call echo("content",content):call echo("paramName",paramName):call echo("paramValue",paramValue)
			'追加参数
            if isTrue = false and paramValue<>"clear(*)" and sType="set" then
				if instr(content,paramName & "=")>0 or isReplace=false then
					if instr(content,"/>")>0 then
						content = replace(content, "/>", " " & paramName & "=""" & findStr & paramValue & """/>")
					else
						content = replace(content, ">", " " & paramName & "=""" & findStr & paramValue & """>")  
					end if
				end if
            end if
            handleHtmlLabelParam = content 
        end if  
    end function 




    '获得Img图片列表
    function getImgSrcList(content)
        getImgSrcList = getHtmlLabelParam(content, "img", "src", "") 
    end function 
    '获得background-image图片列表
    function getBackgroundImage(content)
        getBackgroundImage = getHtmlLabelParam(content, "*", "style", "style>background") 
    end function 
    '获得style-image图片列表
    function getStyleImage(content)
        getStyleImage = pubCssObj.getCssContentUrlList(getHtmlStyleContent(content)) 
    end function 
    '获得Link图版链接列表
    function getLinkImageList(content)
        getLinkImageList = getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico") 
    end function 
    '获得SWF链接列表
    function getSwfSrcList(content)
        dim swf1,swf2,splstr,s,c 
		swf1=pubHtmlObj.handleHtmlLabel(content,"embed","src","type=application/x-shockwave-flash","get","trim(*)")
		swf2=pubHtmlObj.handleHtmlLabel(content,"param","value","name=movie","get","trim(*)") 
        splstr=split(swf1 & vbcrlf & swf2,vbcrlf)
		for each s in splstr
			if s <>"" and s<>"#" and s<>"/" and instr(vbcrlf & c & vbcrlf, vbcrlf & s & vbcrlf)=false then
				if c<>"" then
					c=c & vbcrlf
				end if
				c=c & s
			end if
		next
        getSwfSrcList = c 
    end function 
    '获得A链接列表
    function getAHrefList(content)
        getAHrefList = getHtmlLabelParam(content, "a", "href", "") 
    end function 
    '获得Link链接列表
    function getLinkHrefList(content)
        getLinkHrefList = getHtmlLabelParam(content, "link", "href", "rel=stylesheet") 
    end function 
    '获得script链接列表
    function getScriptSrcList(content)
        getScriptSrcList = getHtmlLabelParam(content, "script", "src", "") 
    end function 


    '获得HTML标签值 带判断
    function getHtmlLabelParam(content, labelNameList, paramName, actionList)
        getHtmlLabelParam = handleHtmlLabel(content, labelNameList, paramName, actionList, "get", "") 
    end function 
    '设置HTML标签值 带判断
    function setHtmlLabelParam(content, labelNameList, paramName, paramValue, actionList)
        setHtmlLabelParam = handleHtmlLabel(content, labelNameList, paramName, actionList, "set", paramValue) 
    end function 
	
	
    '处理HTML标签
    function handleHtmlLabel(byval content, byval labelNameList, byval paramName, byval actionList, byval sType, byval paramValue)
        dim i, s, c, yuanStr, tempS, lalType, nLen, lalStr, splxx, labelType, AHrefList, s1, s2, splLabel, labelName, splStr 
        dim isAction, action, actionLeft, actionRight, actionMode, actionTrueArray(99), nIndex, url, isAnd 
        dim findParamValue, temp_ParamValue,objCss
		
		
		'当标签名称列表与参数名称为空，则退出     
		if labelNameList="" or paramName="" then
			handleHtmlLabel=""
			exit function
		end if
		
        labelType = "|*|" 
        splStr = split(actionList, ",")                                                 '分割动作
        splLabel = split(labelNameList, "|")                                            '分割标签名称
 
        for i = 1 to len(content)
            s = mid(content, i, 1) 
            if s = "<" then
                yuanStr = mid(content, i) 
                tempS = LCase(replace(yuanStr, vbTab, " ")) 
                tempS = replace(replace(replace(tempS, chr(10), " " & vbCrLf), chr(13), " " & vbCrLf),">"," ")  '<div>对这种也处理

                lalStr = mid(yuanStr, 1, inStr(yuanStr, ">")) 
                lalType = LCase(mid(tempS, 1, inStr(tempS, " "))) 
                for each labelName in splLabel
                    if labelNameList = "*" then
                        if instr(tempS, " ") > 0 then
                            labelName = mid(tempS, 2, instr(tempS, " ") - 2) 
                        else
                            labelName = "#NO无名标签 不处理#" 
                        end if 
                    end if 
												'labelType它为|*| 等于处理全部，这个判断无意义，晕，当时自己怎么想的，猪
                    if lalType = "<" & labelName & " " and(inStr(labelType, "|" & labelName & "|") > 0 or inStr(labelType, "|*|") > 0) then
                        nLen = len(lalStr) - 1 
                        i = i + nLen 

                        nLen = instr(lalStr, "<" & labelName & vbTab) 
                        if nLen > 0 then
                            lalStr = "<" & labelName & " " & mid(lalStr, nLen + len("<" & labelName & vbTab)) 
                        end if  
						url = cStr(handleHtmlLabelParam(lalStr, paramName, paramValue, sType)) 
                        '排除#号，不处理
                        if left(actionList, 1) <> "#" then
                            '动作处理
                            isAction = true 
                            nIndex = 0 
                            for each action in splStr
                                action = ltrim(action) 
                                if instr(action, "=") > 0 then
                                    actionLeft = mid(action, 1, instr(action, "=") - 1) 
                                    actionRight = mid(action, instr(action, "=") + 1) 
                            		isAnd = true 
                                    if left(actionLeft, 2) = "||" then
                                        isAnd = false 
                                        actionLeft = mid(actionLeft, 3)
									'并且 默认就是这个
                                    elseif left(actionLeft, 2) = "&&" then 
                                        actionLeft = mid(actionLeft, 3)
                                    end if 

                                    actionMode = "=" 
                                    if left(actionLeft, 1) = "!" then
                                        actionMode = "!=" 
                                        actionLeft = mid(actionLeft, 2) 
                                    elseif left(actionLeft, 2) = "**" then
                                        actionMode = "*=*" 
                                        actionLeft = mid(actionLeft, 3) 
                                    elseif left(actionLeft, 1) = "*" then
                                        actionMode = "*=" 
                                        actionLeft = mid(actionLeft, 2) 
                                    end if  
                                    s1 = phpTrim(getLabelParam(lalStr, actionLeft)) 
                                    nIndex = nIndex + 1 
                                    if actionMode = "!=" then
                                        actionTrueArray(nIndex) =(s1 <> actionRight) 
                                    elseif actionMode = "*=*" then
                                        actionTrueArray(nIndex) = IIF(instr(lcase(s1), lcase(actionRight)), true, false) 
                                    elseif actionMode = "*=" then
                                        actionTrueArray(nIndex) = IIF(instr(s1, actionRight), true, false) 
                                    elseif actionMode = "=" then
                                        actionTrueArray(nIndex) =(s1 = actionRight) 
                                    end if 

                                end if 
                            next 
							'call echo(nIndex,isAnd)
                            if nIndex = 1 then
                                isAction = actionTrueArray(1) 
                            elseif nIndex = 2 then
                                if isAnd = false then
                                    isAction = actionTrueArray(1) or actionTrueArray(2) 
                                else
                                    isAction = actionTrueArray(1) and actionTrueArray(2) 
                                end if 
                            elseif nIndex = 3 then
                                isAction = actionTrueArray(1) and actionTrueArray(2) and actionTrueArray(3) 
                            elseif nIndex = 4 then
                                isAction = actionTrueArray(1) and actionTrueArray(2) and actionTrueArray(3) and actionTrueArray(4) 

                            end if 
                            if isAction = false then
                                url = "" 
                            end if 
                        end if 
                        '背景图片 特殊
                        if actionList = "style>background" and isAction=true then
							url= handleHtmlLabelParam(lalStr, paramName, "", "get")			'获得参数值
							if url<>"" then
								'call echo(" pubCssObj.handleCssContent(""style"","""& url &""",""*"",""background||background-image"",""*"",""" & paramValue & """,""images"")",url)
							end if
							'获得处理后style样式
                            url = pubCssObj.handleCssContent("style",url,"*","background||background-image","*",paramValue,"images")
							 
                        end if 
                        '排除单独运行javascript
                        if left(LCase(url), 11) = "javascript:" then
                            url = "" 
                        end if 
						'排除空值   与设置<div style里背景图片
                        if url <> ""   then  
                            
							if sType = "check" then
								if url="True" then 				'这样写是为了配置vb.net
									handleHtmlLabel=true
									exit function
								end if
							elseif sType = "set" or sType = "edit" then
								'call echo(url , findParamValue)
                                '替换style被图片
                                if instr(actionList, "style>background") > 0 then
									url= handleHtmlLabelParam(lalStr, paramName, "", "get")			'获得参数值
									'call echo(" pubCssObj.handleCssContent(""style"","""& url &""",""*"",""background||background-image"",""*"",""" & paramValue & """,""set"")",url)
									url = pubCssObj.handleCssContent("style",url,"*","background||background-image","*",paramValue,"set")   '替换html中<div style=''里图片
									 c = c & handleHtmlLabelParam(lalStr, paramName, url, sType)   
									   
                                elseif paramValue<>"" then
                                    'call echo(sType,replace(handleHtmlLabelParam(lalStr, paramName, paramValue, sType)  ,"<","&lt;")) 
									c = c & handleHtmlLabelParam(lalStr, paramName, paramValue, sType)  

                                else
                                    c = c & lalStr 
                                end if 
                            else
                                if AHrefList <> "" then
                                    AHrefList = AHrefList & vbCrLf 
                                end if 
                                AHrefList = AHrefList & url 
                            end if 
                        else
                            c = c & lalStr 
                        end if 
                    else
                        c = c & s 
                    end if 
                next 
            else
                c = c & s 
            end if 
            doEvents 
        next  
        if sType = "set" or sType = "edit" then 
            handleHtmlLabel = c 
        else
            handleHtmlLabel = AHrefList 
        end if 
    end function 


end class 


%>