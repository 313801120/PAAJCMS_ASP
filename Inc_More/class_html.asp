<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-02-27
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'����html��ǩ
'dim jsObj:  set jsObj=new class_html
'call rw(jsObj.handleJsContent(c))
class class_html
    dim isDisplayEcho                                                               '�Ƿ����


    '���캯�� ��ʼ��
    sub class_Initialize()
        isDisplayEcho = false 
    end sub 
    '�������� ����ֹ
    sub class_Terminate()
    end sub 

    '��ñ�ǩֵ
    function getLabelParam(byval content, paramName)
        getLabelParam = handleHtmlLabelParam(content, paramName, "", "get") 
    end function 
    '���ñ�ǩֵ
    function setLabelParam(byval content, paramName, paramValue)
        setLabelParam = handleHtmlLabelParam(content, paramName, paramValue, "set") 
    end function 
    '�޸ı�ǩֵ
    function editLabelParam(byval content, paramName, paramValue)
        editLabelParam = handleHtmlLabelParam(content, paramName, paramValue, "edit") 
    end function 
    '����ǩֵ
    function checkLabelParam(byval content, paramName)
        checkLabelParam = handleHtmlLabelParam(content, paramName, "", "check") 
    end function 
    '�����ǩ���� 20160929   
    function handleHtmlLabelParam(byval content, byval paramName, byval paramValue, byval sType) 
        dim i, tempContent, tempContent2, findStr, startStr, endStr, s,s1, s2, nLen, isTrue,splxx
        dim nLeft, nRight,isParamValueTrim,isParamValueLCase,isParamValueFullUrl,isParamValueReplaceDir,isParamValueImg,isParamValueUrl,findParamValue,isParamValueFullForceUrl,paramValueFindValue,paramValueReplaceValue
		dim isCheck					'��ֵ��ȵ�
		dim isUrlEnc:isUrlEnc=false						'url����   ��"[##"& url &"##]"
		dim isReplace:isReplace=false			'���滻Ϊ��ʱ�����Ҫ�������������ſ����滻20170825
		'�ж��Ƿ�ΪͼƬ�����ж�
		if instr(paramValue,"setaddurlenc(*)")>0 then
			paramValue=replace(paramValue,"setaddurlenc(*)","") 
			isUrlEnc=true
		end if
		'�ж��Ƿ�ΪͼƬ�ж�
		if instr(paramValue,"isreplace(*)")>0 then
			paramValue=replace(paramValue,"isreplace(*)","") 
			isReplace=true
		end if
		
		dim isNoEMail:isNoEMail=false				'��Ϊ����
		if instr(paramValue,"isnoemail(*)")>0 then
			paramValue=replace(paramValue,"isnoemail(*)","") 
			isNoEMail=true
		end if
		dim isEMail:isEMail=false				'Ϊ����
		if instr(paramValue,"isemail(*)")>0 then
			paramValue=replace(paramValue,"isemail(*)","") 
			isEMail=true
		end if
		dim isHandleEMail:isHandleEMail=false				'Ϊ����
		if instr(paramValue,"ishandleemail(*)")>0 then
			paramValue=replace(paramValue,"ishandleemail(*)","") 
			isHandleEMail=true
		end if
		
		isCheck=false
		if sType="check" and paramName="*" then
			handleHtmlLabelParam=true
			exit function
		end if
		'��������Ƿ�һ�������������        û�ã����Ľ���
        if instr(paramValue, "[#check#]") > 0 then  
            splxx = split(paramValue, "[#check#]") 
            findParamValue = splxx(0) 
            paramValue = splxx(1) 
            isCheck = true  
        end if  
		'�����ò���ֵ ȥ�����߿ո�
		isParamValueLCase=false
		if instr(paramValue,"lcase(*)")>0 then
			paramValue=replace(paramValue,"lcase(*)","")
			paramValue=lcase(paramValue)					'����Ƿ�ֵ ��������ٴδ���
			isParamValueLCase=true
		end if
		'�����ò���ֵ ȥ�����߿ո�
		isParamValueTrim=false
		if instr(paramValue,"trim(*)")>0 then
			paramValue=replace(paramValue,"trim(*)","")
			paramValue=phptrim(paramValue)					'����Ƿ�ֵ ��������ٴδ���
			isParamValueTrim=true
		end if 
		'�û��url ����
		isParamValueFullUrl=false
		if instr(paramValue,"fullurl(*)")>0 then
			paramValue=replace(paramValue,"fullurl(*)","") 
			isParamValueFullUrl=true
		end if
		'�û��url ǿ������
		isParamValueReplaceDir=false
		if instr(paramValue,"fullforceurl(*)")>0 then
			paramValue=replace(paramValue,"fullforceurl(*)","") 
			isParamValueFullForceUrl=true 
		end if 
		'�û�û�����Ŀ¼
		isParamValueFullForceUrl=false
		if instr(paramValue,"dir(*)")>0 then
			paramValue=replace(paramValue,"dir(*)","") 
			isParamValueReplaceDir=true
		end if
		'�ж��Ƿ�ΪͼƬ�ж�
		isParamValueImg=false
		if instr(paramValue,"isimg(*)")>0 then
			paramValue=replace(paramValue,"isimg(*)","") 
			isParamValueImg=true
		end if
		'�ж��Ƿ�Ϊ��ַ  ��ͼƬ�ж�һ��
		isParamValueUrl=false
		if instr(paramValue,"isurl(*)")>0 then
			paramValue=replace(paramValue,"isurl(*)","") 
			isParamValueUrl=true
		end if
		'�ж��Ƿ��滻��������  
		if instr(paramValue,"replace[")>0 then
			s=getStrCut(paramValue,"replace[","]",0)
			splxx=split(s,"=")
			paramValueFindValue=splxx(0):paramValueReplaceValue=splxx(1) 
			paramValue=replace(paramValue,"replace["& s &"]","")
		end if
		
		 
		
        '�������
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
                '˫����
                startStr = paramName & "=""" 
                endStr = """" 
            elseif i = 2 then
                '������
                startStr = paramName & "='" 
                endStr = "'" 
            elseif i = 3 then
                '����
                startStr = paramName & "=" 
                endStr = " " 
            elseif i = 4 then
                '����
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
				'ֵСд
				if isParamValueLCase=true then
					findStr=lcase(findStr)
				end if
				'ֵ������߿ո�
				if isParamValueTrim=true then
					findStr=phptrim(findStr)
				end if
				
				
				'�Ƿ�ΪͼƬ��ַ
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
				'�ų�����
				if isnoemail=true and lcase(left(phptrim(findStr),7))="mailto:" then	'׷����20171208
					findStr=""
				end if
				'��ȡ����
				if isemail=true and lcase(left(phptrim(findStr),7))<>"mailto:" then	'׷����20171208
					findStr=""
				elseif isemail=true  and isHandleEMail=true then
					findStr=mid(findStr,8)
				end if
				
				
				'call echo("findStr",findStr)
					
				'��ַ���� �ų�base64����(addto20171115)
				if isParamValueFullUrl=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					findStr=fullHttpUrl(paramValue,findStr)
				'ǿ����ַ����  
				elseif isParamValueFullForceUrl=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					'call echo(paramValue,findStr)
					findStr=getFileAttr(findStr,"name")
					'call echo(paramValue,findStr)
					findStr=fullHttpUrl(paramValue,findStr) 
					
				'�滻Ŀ¼
				elseif isParamValueReplaceDir=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					if instr(findStr,"?")>0 then
						findStr=mid(findStr,1,instr(findStr,"?")-1)
					end if
					findStr=paramValue & getFileAttr(findStr,"2")
				end if

				'URL���� �� "[##"& url &"##]"   20161017
				if isUrlEnc=true and findStr<>"" and instr(lcase(findStr),";base64,")=false then  
					findStr=hRFileName(findStr)
				end if

						
				'���õ�ʱ�� isParamValueFullForceUrl  �����ò��ϣ���ʱ���ϰ�20170825
                if sType = "set" or sType = "edit" then 
					if (isCheck = true and findStr = findParamValue) or isCheck = false then
						if isParamValueFullUrl=false and isParamValueFullForceUrl=false and  isParamValueReplaceDir=false then 
							findStr = paramValue 
						end if
					end if
                end if 
                '���
                if paramValue = "clear(*)" then
                    'ǿǿǿ���滻
                    nLen = inStr(tempContent, startStr) 
                    s1 = mid(content, 1, instr(tempContent, startStr) - 1) 
                    s2 = right(content, len(content) - nLen + 1 - len(startStr)) 
                    s2 = mid(s2, instr(s2, endStr) + 1) 
                    s1 = rTrim(s1)                                                                  'ȥ�ұ߿ո�
                    'call rwend(s1 & s2) 
					content=s1 & s2
				'�滻20171109
				elseif paramValueFindValue<>"" then
					if left(paramValueFindValue,2)="**" then
						paramValueFindValue=mid(paramValueFindValue,3) 
						content=replaceNoULCase(content,paramValueFindValue,paramValueReplaceValue)
					else
						content=replace(content,paramValueFindValue,paramValueReplaceValue)
					end if
                else
                    'ǿǿǿ���滻
                    nLen = inStr(tempContent, startStr) - 1 + len(startStr)                         'ע����ת��Сд��tempContent
                    s2 = right(content, len(content) - nLen) 
                    s2 = mid(s2, inStr(s2, endStr)) 
                    s1 = left(content, nLen) 
					
					
                    content = s1 & findStr & s2 
                end if 
                exit for                                                                        '�ҵ����˳�
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
			'׷�Ӳ���
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




    '���ImgͼƬ�б�
    function getImgSrcList(content)
        getImgSrcList = getHtmlLabelParam(content, "img", "src", "") 
    end function 
    '���background-imageͼƬ�б�
    function getBackgroundImage(content)
        getBackgroundImage = getHtmlLabelParam(content, "*", "style", "style>background") 
    end function 
    '���style-imageͼƬ�б�
    function getStyleImage(content)
        getStyleImage = pubCssObj.getCssContentUrlList(getHtmlStyleContent(content)) 
    end function 
    '���Linkͼ�������б�
    function getLinkImageList(content)
        getLinkImageList = getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico") 
    end function 
    '���SWF�����б�
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
    '���A�����б�
    function getAHrefList(content)
        getAHrefList = getHtmlLabelParam(content, "a", "href", "") 
    end function 
    '���Link�����б�
    function getLinkHrefList(content)
        getLinkHrefList = getHtmlLabelParam(content, "link", "href", "rel=stylesheet") 
    end function 
    '���script�����б�
    function getScriptSrcList(content)
        getScriptSrcList = getHtmlLabelParam(content, "script", "src", "") 
    end function 


    '���HTML��ǩֵ ���ж�
    function getHtmlLabelParam(content, labelNameList, paramName, actionList)
        getHtmlLabelParam = handleHtmlLabel(content, labelNameList, paramName, actionList, "get", "") 
    end function 
    '����HTML��ǩֵ ���ж�
    function setHtmlLabelParam(content, labelNameList, paramName, paramValue, actionList)
        setHtmlLabelParam = handleHtmlLabel(content, labelNameList, paramName, actionList, "set", paramValue) 
    end function 
	
	
    '����HTML��ǩ
    function handleHtmlLabel(byval content, byval labelNameList, byval paramName, byval actionList, byval sType, byval paramValue)
        dim i, s, c, yuanStr, tempS, lalType, nLen, lalStr, splxx, labelType, AHrefList, s1, s2, splLabel, labelName, splStr 
        dim isAction, action, actionLeft, actionRight, actionMode, actionTrueArray(99), nIndex, url, isAnd 
        dim findParamValue, temp_ParamValue,objCss
		
		
		'����ǩ�����б����������Ϊ�գ����˳�     
		if labelNameList="" or paramName="" then
			handleHtmlLabel=""
			exit function
		end if
		
        labelType = "|*|" 
        splStr = split(actionList, ",")                                                 '�ָ��
        splLabel = split(labelNameList, "|")                                            '�ָ��ǩ����
 
        for i = 1 to len(content)
            s = mid(content, i, 1) 
            if s = "<" then
                yuanStr = mid(content, i) 
                tempS = LCase(replace(yuanStr, vbTab, " ")) 
                tempS = replace(replace(replace(tempS, chr(10), " " & vbCrLf), chr(13), " " & vbCrLf),">"," ")  '<div>������Ҳ����

                lalStr = mid(yuanStr, 1, inStr(yuanStr, ">")) 
                lalType = LCase(mid(tempS, 1, inStr(tempS, " "))) 
                for each labelName in splLabel
                    if labelNameList = "*" then
                        if instr(tempS, " ") > 0 then
                            labelName = mid(tempS, 2, instr(tempS, " ") - 2) 
                        else
                            labelName = "#NO������ǩ ������#" 
                        end if 
                    end if 
												'labelType��Ϊ|*| ���ڴ���ȫ��������ж������壬�Σ���ʱ�Լ���ô��ģ���
                    if lalType = "<" & labelName & " " and(inStr(labelType, "|" & labelName & "|") > 0 or inStr(labelType, "|*|") > 0) then
                        nLen = len(lalStr) - 1 
                        i = i + nLen 

                        nLen = instr(lalStr, "<" & labelName & vbTab) 
                        if nLen > 0 then
                            lalStr = "<" & labelName & " " & mid(lalStr, nLen + len("<" & labelName & vbTab)) 
                        end if  
						url = cStr(handleHtmlLabelParam(lalStr, paramName, paramValue, sType)) 
                        '�ų�#�ţ�������
                        if left(actionList, 1) <> "#" then
                            '��������
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
									'���� Ĭ�Ͼ������
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
                        '����ͼƬ ����
                        if actionList = "style>background" and isAction=true then
							url= handleHtmlLabelParam(lalStr, paramName, "", "get")			'��ò���ֵ
							if url<>"" then
								'call echo(" pubCssObj.handleCssContent(""style"","""& url &""",""*"",""background||background-image"",""*"",""" & paramValue & """,""images"")",url)
							end if
							'��ô����style��ʽ
                            url = pubCssObj.handleCssContent("style",url,"*","background||background-image","*",paramValue,"images")
							 
                        end if 
                        '�ų���������javascript
                        if left(LCase(url), 11) = "javascript:" then
                            url = "" 
                        end if 
						'�ų���ֵ   ������<div style�ﱳ��ͼƬ
                        if url <> ""   then  
                            
							if sType = "check" then
								if url="True" then 				'����д��Ϊ������vb.net
									handleHtmlLabel=true
									exit function
								end if
							elseif sType = "set" or sType = "edit" then
								'call echo(url , findParamValue)
                                '�滻style��ͼƬ
                                if instr(actionList, "style>background") > 0 then
									url= handleHtmlLabelParam(lalStr, paramName, "", "get")			'��ò���ֵ
									'call echo(" pubCssObj.handleCssContent(""style"","""& url &""",""*"",""background||background-image"",""*"",""" & paramValue & """,""set"")",url)
									url = pubCssObj.handleCssContent("style",url,"*","background||background-image","*",paramValue,"set")   '�滻html��<div style=''��ͼƬ
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