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
'js ѹ�����ѹ����

'dim jsObj:  set jsObj=new class_js
'call rw(jsObj.handleJsContent(c))
class class_js
	dim isDisplayEcho			'�Ƿ����  
	dim jsErrorList,jsErrorCount,httpurl
	dim imgList					'ͼƬ�б�
	dim cssList					'CSS�б�
	dim jsList					'JS�б�		
	dim swfList					'SWF�б�	
	dim sourceList				'source�б� ��Ƶ 
	dim videoList				'video�б� ��Ƶ 

	 '���캯�� ��ʼ��
    Sub Class_Initialize() 
		httpurl="http://www.sharembweb.com/"
		isDisplayEcho=false
	end sub
    '�������� ����ֹ
    Sub Class_Terminate() 
    End Sub 
	 
	'����JS���� sType=source  ΪԴ�벻���������
	function handleJsContent(byVal content, byval sType)
		dim splStr, sx, tempSx, s, clearS, c, i, j, nSYHCount, s1, s2, tempZc, rowC, tempRow, noteStr, downRow, downWord, nc, isLoop, beforeStr, afterStr, endCode, tempEndCode ,leftStr,errS
		dim upWord, upWord2, yesOK,isJS 
		dim wc                                                                          '�����ı��洢����
		dim wcType                                                                      '�����ı����ͣ��� " �� '
		dim zc                                                                          '��ĸ�ļ��洢����
		dim tempS 
		dim RSpaceStr, RSpaceNumb, addToRSpaceNumb                                      '�пո���ո���,׷������
		dim isDanHaoNote																'����ע��Ϊ��
		dim multiLineComment                                                            '����ע��
		multiLineComment = ""                                                           '����ע��Ĭ��Ϊ��
		RSpaceStr = "    " : RSpaceNumb = 0                                             '�пո�ֵ �� �ո��� ���ó�ֵ
		nSYHCount = 0                                                                   '˫����Ĭ��Ϊ0
		dim isAddToSYH                                                                  '�Ƿ��ۼ�˫����
		isJS=true
		jsErrorCount=0																	'��������
		
		sType="|"& sType &"|"
		
		splStr = split(content, vbCrLf)                                                 '�ָ���
		'ѭ������
		for j = 0 to uBound(splStr)
			s = splStr(j)
			s = phptrim(s)
			leftStr=left(splStr(j),len(splStr(j))-len(phpLTrim(splStr(j))))				'�����߱���յ��ַ� ����ԭjsԭ��
			s = replace(replace(s, chr(10), ""), chr(13), "") '���Ϊʲô s����� chr(10)��chr(13) �أ�
			clearS = phpTrim(s)                                                             '������߿ո����˸�
			tempS = s                                                                       '�ݴ�S
			rowC = "" : tempRow = ""                                                        '���ÿ��ASP������ݴ������д���
			noteStr = ""                                                                    'Ĭ��ע�ʹ���Ϊ��
			downRow = ""                                                                    '��һ�д���
			downWord = ""                                                                   '��һ�е���
			addToRSpaceNumb = 0                                                             '���׷������
			isDanHaoNote=false		'����ע��Ϊ��
			if(j + 1) <= uBound(splStr) then
				downRow = splStr(j + 1) 
				downWord = getBeforeNStr(downRow, "ȫ��", 1) 
			end if 
			nc = ""                                                                         '����Ϊ��
			isLoop = true                                                                   'ѭ���ַ�Ϊ��
			'ѭ��ÿ���ַ�
			for i = 1 to len(s)
				sx = mid(s, i, 1) : tempSx = sx 
				beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '��һ���ַ�
				afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '��һ���ַ�
				endCode = mid(s, i + 1)                                                         '��ǰ�ַ���������� һ��
				'�����ı�
				if(sx = """" or sx = "'" and wcType = "") or sx = wcType or wc <> "" then
					isAddToSYH = true 
					'����һ�ּ򵥵ķ�����������(20150914)
					if isAddToSYH = true and beforeStr = "\" then	
						if len(wc) >= 1 then
							if isStrTransferred(wc) = true then  'Ϊת���ַ�Ϊ�� 
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
	
					'˫�����ۼ�
					if sx = wcType and isAddToSYH = true then nSYHCount = nSYHCount + 1         '�ų���һ���ַ�Ϊ\���ת���ַ�(20150914)
	
	
					'�ж��Ƿ�"�����
					if nSYHCount mod 2 = 0 and beforeStr <> "\" then
						if mid(s, i + 1, 1) <> wcType then
							wc = wc & sx
							'call echo(wcType,replace(replace(wc,"<","&lt;"),"\" & wcType, wcType))   '��JS������ͼƬ������ʵ��
							'call echo("wcType=" & wcType,wc)
							if instr("|"& sType &"|","|html|")>0 then 
								wc=handleJSContentHtml(wcType,wc)					'����JS���html���� 20171113
								'call echo(wc,"eeee")
							end if
							
							rowC = rowC & wc                     '�д����ۼ�
							'call echo("wc",wc)
							nSYHCount = 0 : wc = ""              '���
							wcType = "" 
						else
							wc = wc & sx 
						end if 
					else
						wc = wc & sx 
					end if 
	
				'����ע��
				elseIf(sx = "/" and afterStr = "*") or multiLineComment <> "" then
					noteStr = mid(s, i) 
					'call echo("����ע��",noteStr)
					if multiLineComment <> "" then multiLineComment = multiLineComment & vbCrLf 
					multiLineComment = multiLineComment & noteStr 
					if inStr(noteStr, "*/") > 0 then
						'ɾ��ע��
						if instr("|"& sType &"|","|delnote|")>0 then 
							multiLineComment=""
						end if
						rowC = rowC & multiLineComment 
						multiLineComment = "" 
					end if 
					exit for 
				'ע�����˳� ��ѡע��
				elseIf sx = "/" and afterStr = "/" then
					noteStr = mid(s, i)
					'Ϊѹ��ʱ ����ע�ͼӻ���
					if instr("|"& sType &"|","|zip|")>0 then
						isDanHaoNote=true		'����ע��Ϊ��
					end if
					'ɾ��ע��
					if instr("|"& sType &"|","|delnote|")>0 then 
						noteStr=""
					end if
					rowC = rowC & noteStr 
					exit for 
	
				'+-*\=&   �ַ����⴦��
				elseIf inStr(".&=,+-*/:()><;", sx) > 0 and nc = "" then
					if zc <> "" then
						tempZc = zc 
						upWord2 = upWord                  '��¼����һ������
						upWord = LCase(tempZc)            '��¼��һ����������   ΪСд
						rowC = rowC & zc 
						zc = ""                           '�����ĸ����
					end if 
					'�����if(1=1){;;;;;;;;}   20151224
					if sx = ";" and instr("{;" , right(trim(rowC), 1)) > 0 then			'�ж���һ����Ч�ַ�����;{
						jsErrorCount=jsErrorCount+1			'�ҵ�һ������
						errS=j & "�У�����:��  " & s  & vbcrlf
						if instr(vbcrlf & jsErrorList & vbcrlf, vbcrlf & errS & vbcrlf)=false then
							jsErrorList=jsErrorList & errS
						end if
					
						sx = "" 
					end if 
					rowC = rowC & sx 
					upWord2 = upWord                                                            '��¼����һ������
					upWord = sx 
				'��ĸ
				elseIf checkABC(sx) = true or sx = "_" or zc <> "" then
					zc = zc & sx 
					yesOK = true 
					s1 = mid(s & " ", i + 1, 1) 
					s2 = mid(zc, 1, 1) 
					if checkABC(s1) <> true and s1 <> "_" then
						yesOK = false 
					end if 
					'�����������������
					if checkNumber(s1) = true and checkABC(s2) = true then
						yesOK = true 
					end if 
					if yesOK = false then 
						tempZc = zc 
	
						upWord2 = upWord                  '��¼����һ������
						upWord = LCase(tempZc)            '��¼��һ����������   ΪСд
						rowC = rowC & zc 
						zc = ""                           '�����ĸ����
					end if 
				'Ϊ����
				elseIf checkNumber(sx) = true or nc <> "" then
					nc = nc & sx 
					if afterStr <> "." and checkNumber(afterStr) = false then
						rowC = rowC & nc 
						nc = "" 
					end if 
				else
					'JS��PHP����{}����
					if sx = "{" then
						if phptrim(rowC) <> "" then
							addToRSpaceNumb = 1 
						else
							RSpaceNumb = RSpaceNumb + 1 
						end if 
					elseIf sx = "}" then
						if RSpaceNumb > 0 then
							if phptrim(rowC) <> "" then
								'ɾ����̨ע��20151224
								tempEndCode = endCode 
								if instr(tempEndCode, "//") > 0 then
									tempEndCode = mid(tempEndCode, 1, instr(tempEndCode, "//") - 1) 
								end if 
								'�޸�if(a==b){}  ����
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
				'�ݴ�ÿ������
				tempRow = tempRow & sx 
				if isLoop = false then exit for 
			next 
			if rowC = ";}" then
				rowC = "}" 
			end if 
			if instr("|"& sType &"|","|source|")>0 then
				rowC=leftStr & rowC
			elseif rowC <> "" and RSpaceNumb > 0 then
				rowC = copyStrNumb(RSpaceStr, RSpaceNumb) & replace(rowC, vbCrLf, vbCrLf & copyStrNumb(RSpaceStr, RSpaceNumb)) '�޽�20150902
			
			end if 
			RSpaceNumb = RSpaceNumb + addToRSpaceNumb 
			if multiLineComment = "" then                                                   '�޽�20150902
				'ѹ��
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
	
	'����JS���html���� 
	function handleJSContentHtml(byval wcType,byval wc)
		dim thisFormatObj:set thisFormatObj = new class_formatting                                        '��ʼ����ʽ��HTML�� 
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
		'��ʽ��HTML
		call thisFormatObj.handleFormatting(content) 
		call thisFormatObj.handleLabelContent("*", "*", "setaddurlenc(*)isimg(*)fullurl(*)" & httpurl, "handleimg") 
		'��ø�ʽ����HTML
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