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

'dim t:  set t=new class_css
'call echo("",t.id) 
class class_css 
	dim isDisplayEcho

	 '���캯�� ��ʼ��
    Sub Class_Initialize() 
		isDisplayEcho=false
	end sub
    '�������� ����ֹ
    Sub Class_Terminate() 
    End Sub 
	 
	'���CSS������URL�б�
	function getCssContentUrlList(content)
		getCssContentUrlList=handleCssContent("css",content ,"*","background||background-image","trim(*)*","","images")
	end function 
	'get,set,images,error,errorcount,labellist,paramlist
	'��ô��̬�Ĳ���CSS������Ҳ��д�������������ж��У�ʱ����ʵ��̫��ԣ�ˣ����Լ�������ÿ�����һ�����ñ��С��
	'����CSS��ʽ  ̫�����ˣ�һ��Ҫд��ע��   Ĭ�ϻ������ image��ñ���ͼƬ�б�  �ų���ע��������  <space>  Ϊ���
	function handleCssContent(htmlOrCssOrStyle,byval content,byval sCssLabelName,byval sCssParamName,byval findSCssParamValue,byval sCssParamValue,byval sType)
		dim splstr,i,s,c,rowC,nRowCount,saveRowC,splxx,isCssStyle,endCode,nLen
		dim cssLabelName,tempCssLabelName,cssParamName,tempCssParamName
		dim cssLabelValue,cssParamValue,yunCssParamValue,tempCssParamValue,noteC
		dim isCssLabel,isCssParam,isNote,isMaoHao,url,urlList
		dim beforeStr,afterStr,nCountLen,isKuHu
		dim findLabel,findParam,findValue,isLabelOK,isParamOK,isValueOK,str1,str2
		dim isLabelNameLCase,isFindValueTrim,isFindValueLCase,isSetValueTrim,isSetValueLCase
		dim cssErrorList,cssErrorCount,errS				'css�����б����������
		dim labelList,paramList,isParamList			'��ʽ�����б����������б�
		dim isSCssParamValueImg,isImg
		dim isUrlEnc:isUrlEnc=false						'url����   ��"[##"& url &"##]"
		isSCssParamValueImg=false
		dim isParamValueFullForceUrl:isParamValueFullForceUrl=false					'ǿ����ַ����ʹ�õ�ǰ��ַ 
		dim s2,c2
		dim splCssParamValue,nCssImgIndex		'�� src(1.jpg),src(1.jpg) ����
		
		sType="|"& sType &"|"			'��������
		
		'�û��url ǿ������
		if instr(sCssParamValue,"fullforceurl(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"fullforceurl(*)","") 
			isParamValueFullForceUrl=true 
		end if
		 
		'�ж��Ƿ�ΪͼƬ�ж�
		if instr(sCssParamValue,"setaddurlenc(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"setaddurlenc(*)","") 
			isUrlEnc=true
		end if
		
		'�ж��Ƿ�ΪͼƬ�ж�
		if instr(sCssParamValue,"isimg(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"isimg(*)","") 
			isSCssParamValueImg=true
		end if
 
		'��������Ƿ�һ�������������        û�ã����Ľ���
        if instr(sCssParamValue, "[#check#]") > 0 then  
            splxx = split(sCssParamValue, "[#check#]") 
            findSCssParamValue = splxx(0) 
            sCssParamValue = splxx(1)  
        end if   
		
		isLabelNameLCase=false		'��ǩ��ת��Сд 
		isFindValueTrim=false		'����ֵ������߿ո�
		isFindValueLCase=false		'����ֵת��Сд 
		isSetValueTrim=false		'���ֵ������߿ո�
		isSetValueLCase=false		'���ֵתСд
		'�����ò���ֵ ȥ�����߿ո�   �û����ʽ����תСд  ��̫��ɶ�õ�
		if instr(sCssLabelName,"lcase(*)")>0 then
			sCssLabelName=replace(sCssLabelName,"lcase(*)","")
			isLabelNameLCase=true
		end if
		
		'��ò���ֵ �����Сд
		if instr(findSCssParamValue,"lcase(*)")>0 then
			findSCssParamValue=replace(findSCssParamValue,"lcase(*)","")
			isFindValueLCase=true
		end if 
		'�����ò���ֵ ȥ�����߿ո�
		if instr(findSCssParamValue,"trim(*)")>0 then
			findSCssParamValue=replace(findSCssParamValue,"trim(*)","")
			isFindValueTrim=true
		end if
		
		'�������ֵ �͸�ֵ �����Сд
		if instr(sCssParamValue,"lcase(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"lcase(*)","")
			sCssParamValue=lcase(sCssParamValue)
			isSetValueLCase=true
		end if 
		'����������ֵ �͸�ֵ ȥ�����߿ո�
		if instr(sCssParamValue,"trim(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"trim(*)","")
			sCssParamValue=phptrim(sCssParamValue)						'�Ը�ֵ����Ҳ�Ի��ֵ����ֻ��������������ﴦ����
			isSetValueTrim=true
		end if

		dim isSCssParamValueFullUrl,isSCssParamValueReplaceDir
		isSCssParamValueFullUrl=false
		isSCssParamValueReplaceDir=false
		'�û��url ����
		if instr(sCssParamValue,"fullurl(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"fullurl(*)","") 
			isSCssParamValueFullUrl=true
		end if
		'�û�û�����Ŀ¼
		if instr(sCssParamValue,"dir(*)")>0 then
			sCssParamValue=replace(sCssParamValue,"dir(*)","") 
			isSCssParamValueReplaceDir=true
		end if
		
		'��ʽ������ ��ΪСд
		sCssParamName=lcase(sCssParamName)
		
		
		cssErrorCount=0				'css��������
		isParamList=false			'�ռ������б�
		
		isCssLabel=false	'��ʽ��Ϊ��
		isCssParam=false	'��ʽ����Ϊ��
		isNote=false		'ע��Ϊ��
		isMaoHao=false			'�Ƿ�Ϊð��
		isCssStyle=true				'Ĭ��ΪCSS
		isKuHu=false			'()����
		'��<div style='color:red'>  ��div��style���� ��Ϊ��û��{��Ĭ����{Ϊ�� 
		if lcase(htmlOrCssOrStyle)="style" then
			isCssLabel=true 
			content=content & " "			'�������Ӹ��ո�  ���ܻ��ֵ����Bug�Ժ��ٸ��ˣ�����̫����20161010
		elseif lcase(htmlOrCssOrStyle)="html" then
			isCssStyle=false
		end if
		nRowCount=1					'������Ĭ������
		nCountLen=len(content)
		for i = 1 to nCountLen
			s = mid(content, i, 1)
			if s=chr(10)   then
				nRowCount=nRowCount+1
			end if
			beforeStr=""
			if i>1 then
				beforeStr = mid(content, i-1, 1)                        '��һ���ַ�
			end if
			afterStr = mid(content & " ", i+1, 1)       				'��һ���ַ� 
			if isCssStyle=false then
				if s="<" then
					endCode=lcase(mid(content,i))
					if left(endCode,7)="<style>" or left(endCode,7)="<style " then 
						nLen=instr(endCode,">")
						c=c & mid(content,i,nLen)
						i=i+nLen 
						isCssStyle=true			'CSS��ʽΪ��
					else
						c=c & s
					end if
				elseif s<>"}" then
					c=c & s
				end if
			elseif isCssStyle=true and lcase(mid(content,i,8))="</style>" then
				isCssStyle=false			'CSS��ʽΪ��
				c=c & s
			'ע�Ϳ�ʼ
			elseif s="/" and afterStr="*"  then
				isNote=true									'����ע��
				c=c & rowC
				rowC=""
				noteC=s										'ע�͸�ֵ
				s=""										'���s
			'ע������  �����
			elseif isNote=true then
				noteC=noteC & s								'ע���ۼ�
				if s="/" and beforeStr="*" then				'ע�ͽ��
					saveRowC=noteC							'���渳ֵ 
					noteC=""								'���ע��
					if instr(sType,"|delnote|")>0 then
						c=phptrim(c)
					 	saveRowC=""
					end if
					isNote=false 							'�ر�ע��
				end if 						
				s=""										'���s
			'css��ʼ��ǩ
			elseif s="{" then 
				cssLabelName=phptrim(rowC)				'������߿ո���TAB
				tempCssLabelName=lcase(cssLabelName)	'��ʱCSS��ǩ��תСд   
				saveRowC = rowC							'��ʱ���
				if instr(sType,"|zip|")>0 then
					saveRowC=phpTrim(saveRowC)
				elseif instr(sType,"|unzip|")>0 then
					saveRowC=vbcrlf & phpTrim(saveRowC)
				end if
				isCssLabel=true							'����CSS��ǩ
			'css�����ǩ
			elseif s="}" then
				if isCssParam=true then					'����ֵΪ��
					cssParamValue=rowC					'��¼����ֵ��һ����;�����
				end if
				saveRowC = rowC 						'��ʱ���
				if instr(sType,"|zip|")>0 then
					saveRowC=phpTrim(saveRowC) 
				elseif instr(sType,"|unzip|")>0 then
					saveRowC=saveRowC & vbcrlf 
				end if 
				isCssLabel=false						'�ر�CSS��ǩ
				isCssParam=false						'�ر���ʽ����
				isMaoHao=false							'ð�Źر�
				'�������б�Ϊ�棬��û�б��������ˣ����˳�
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
			'������ʼ��ǩ
			elseif isCssLabel=true and  s=":"  then
				if isMaoHao=false then
					cssParamName=rowC						'��������
					saveRowC = rowC							'��ʱ���
					if instr(sType,"|zip|")>0 then
						saveRowC=phpTrim(saveRowC) 
					'���ܴ���
					elseif instr(sType,"|unzip|")>0 then
						saveRowC= vbcrlf & "    " & phpTrim(saveRowC) 
					end if
					isCssParam=true							'��������
					isMaoHao=true							'ð�ſ���
					'call echo("cssParamName",cssParamName)
				else
					'rowC=rowC & s 							'����ע�⣬��Ҫ�ۼ�����color::::::red;      ��ʱ����Ҫ�����Ǳ���
					cssErrorCount=cssErrorCount+1			'�ҵ�һ������
					errS=nRowCount & "�У�����:��  " & phptrim(cssLabelName) & ">>" & phptrim(cssParamName) & vbcrlf
					if instr(vbcrlf & cssErrorList & vbcrlf, vbcrlf & errS & vbcrlf)=false then
						cssErrorList=cssErrorList & errS
					end if
				end if
			'����������ǩ
			elseif isCssLabel=true and ( s=";" or i=nCountLen  ) and left(phptrim(lcase(rowc)),9)<>"@charset " then		'������Ҫ��i=nCountLen��,��ʱ��ô���	�Ӹ� @charset �ų�����ʱ20171003
				
				'�жϵ�ǰ ; ��ǰһ���ַ������� ;20171028
				if s=";" and right(phptrim(mid(content,1,i-1)),1)<>";" then 
					 
					cssParamValue=rowC 						'��ò���ֵ
					saveRowC = rowC							'��ʱ���
					isCssParam=false						'�رղ���
					isMaoHao=true							'ð�Źر�
				end if
			else
				rowC=rowC & s							'�ۼ��ַ�
				saveRowC=""								'��ռ�ʱ����
			end if 
			'������
			if saveRowC<>"" then   
				 if (s=";" or s="}" or i=nCountLen ) and cssParamName<>"" or (s="}" and instr(sType,"|labellist|")>0 ) then 		'��������������
					tempCssParamName=lcase(phptrim(cssParamName))					'�����Сд��������߿ո�
					tempCssParamValue=replace(replace(lcase(phptrim(cssParamValue))," ",""),vbtab,"")	'�������
					
					yunCssParamValue=cssParamValue
					'����ͼƬ����
					if (tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src") and instr(tempCssParamValue,"url(")>0 then 
						'call echo("cssParamValue1",cssParamValue)
						cssParamValue=batchGetKuHuoValue(cssParamValue)								'��������תСд������߿ո�
						'call echo("cssParamValue2",cssParamValue)
					end if
					
					isLabelOK=false
					isParamOK=false
					isValueOK=false
					findLabel=getStrCut(sCssLabelName,"find(",")",2)
					findParam=getStrCut(sCssParamName,"find(",")",2)
					findValue=getStrCut(findSCssParamValue,"find(",")",2)
					'css��ʽ��
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
					'css��ʽ������
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
					'css��ʽ����ֵ
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
					'������߿ո�
					if isSetValueTrim=true then
						cssParamValue=phptrim(cssParamValue)
					end if
					if isSetValueLCase=true then
						cssParamValue=lcase(cssParamValue)
					end if  
					isImg=true
					'�Ƿ�ΪͼƬ��ַ
					if isSCssParamValueImg=true then
						if phptrim(cssParamValue)="#" or phptrim(cssParamValue)="/" or phptrim(cssParamValue)="\" or lcase(left(phptrim(cssParamValue),11))="data:image/" then
							isImg=false
						end if
					end if 
					 
					'��ַ����
					if isSCssParamValueFullUrl=true or isParamValueFullForceUrl=true then 
						'�ų����ֲ���������20170824
						if left(phptrim(cssParamValue),5)<>"rgba("  and left(phptrim(cssParamValue),8)<>"-webkit-" and left(phptrim(cssParamValue),6)<>"alpha(" and left(phptrim(cssParamValue),11)<>"data:image/" then
							'������߿ո�
							if isSetValueTrim=true then
								cssParamValue=phptrim(cssParamValue)
							end if
						
						
							'ǿ��ʹ�õ�ǰ��ַ
							if isParamValueFullForceUrl=true and cssParamValue<>"" then  
								cssParamValue=getFileAttr(cssParamValue,"name") 
							end if
						
							cssParamValue=batchFullHttpUrl(sCssParamValue,cssParamValue)		'������ã�����20171124�ԣ�����
						else
							isParamOK=false				'�Ӹ�����������ô��󳳳���
						end if
					'�滻Ŀ¼
					elseif isSCssParamValueReplaceDir=true then					
						'�ų����ֲ���������20170824   instr(phptrim(cssParamValue),"url(")�Ժ������
						if left(phptrim(cssParamValue),5)<>"rgba(" and left(phptrim(cssParamValue),8)<>"-webkit-" and left(phptrim(cssParamValue),6)<>"alpha(" and left(phptrim(cssParamValue),11)<>"data:image/" then
							 
							'����������url(1.jpg),url(2.jpg)���� 20171211
							splCssParamValue=split(cssParamValue,vbcrlf) 
							cssParamValue=""
							for each s2 in splCssParamValue 
								if instr(s2,"?")>0 then
									s2=mid(s2,1,instr(s2,"?")-1)
								end if 
								str1= getFileAttr(s2,"2")
								'������߿ո�
								if isSetValueTrim=true then
									str1=phptrim(str1)
								end if
								s2=sCssParamValue & str1
								if cssParamValue<>"" then cssParamValue=cssParamValue & vbcrlf
								cssParamValue=cssParamValue & s2
							next
						else
							isParamOK=false				'�Ӹ�����������ô��󳳳���
						end if
					end if  
					c2=""		'��������ͻ����һ����ͼƬ�����ۼӵ���һ��ͼƬ��20171203
					'URL���� �� "[##"& url &"##]"   20161017
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
					
					'����ֵ
					if instr(sType,"|set|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then  
							'�Ա���ͼƬ���⴦��
							if tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src"  then		'src����20171124			
								'background: red;  �ų�����
								if instr(saveRowC,"(")>0 then
									if isImg=true then 
										'call echo("yunCssParamValue",yunCssParamValue)
										splxx=split(yunCssParamValue,",")		'���������õ�20171124
										splCssParamValue=split(cssParamValue,vbcrlf)
										nCssImgIndex=-1
										for each s2 in splxx
											'call echo("s2",s2)
											str1=getStrCut(s2,"(",")",1)
											nCssImgIndex=nCssImgIndex+1 
											'����ǰֵ
											if isParamValueFullForceUrl=true or isSCssParamValueFullUrl=true or isSCssParamValueReplaceDir=true then
												'call echo(saveRowC,str1)
												'call echo(nCssImgIndex,splCssParamValue(nCssImgIndex))
												'call hr()
												'call echo(nCssImgIndex,splCssParamValue(nCssImgIndex))
												'call echo(yunCssParamValue,saveRowC)
												if nCssImgIndex<=ubound(splCssParamValue) then
													saveRowC=replace(saveRowC,str1,"("& splCssParamValue(nCssImgIndex) &")") 'cssParamValue �������棬���Ƕ༶20171211						 
												end if
												'call echo(saveRowC & " =>> " & str1,splCssParamValue(nCssImgIndex)) 
											'�滻ֵ
											else 
												saveRowC=replace(yunCssParamValue,str1,"("& sCssParamValue &")") 
												'call echo("�滻",str1)
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
					'���ֵ
					elseif instr(sType,"|getimage|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then 
				 
							handleCssContent=cssParamValue
							
							exit function 
						end if
					elseif instr(sType,"|get|")>0 then
						if isLabelOK=true and isParamOK=true and isValueOK=true then 
			 
							handleCssContent=cssParamValue				'yunCssParamValue�������������Ϊ��û�д����
							exit function 
						end if
					'��ʽ�����б� 
					elseif instr(sType,"|labellist|")>0 then
						if isLabelOK=true and  s="}" then
							labelList=labelList & phptrim(cssLabelName) & vbcrlf
						end if
					'��ʽ���Ʋ����б�
					elseif instr(sType,"|paramlist|")>0 then
						if isLabelOK=true and isParamOK=true then
							paramList=paramList & tempCssParamName & vbcrlf
							isParamList=true
							if s="}" then
								exit for
							end if
						end if
					'���ͼƬ�б�
					elseif (instr(sType,"|images|")>0 or instr(sType,"|allimages|")>0) and (tempCssParamName="background" or tempCssParamName="background-image" or tempCssParamName="src") and instr(tempCssParamValue,"(")>0 then		'�ж��Ƿ�Ϊ����ͼƬ 
						if isLabelOK=true and isParamOK=true and isValueOK=true and isImg=true then 
							url = cssParamValue 
							'�ų� rgba(  ���� 20170824
							if url<>"" and ( (instr(sType,"|images|")>0 and left(phptrim(url),11)<>"data:image/"   and left(phptrim(url),5)<>"rgba("and left(phptrim(url),8)<>"-webkit-") or instr(sType,"|allimages|")>0) then  
								if urlList<>"" then
									urlList=urlList & vbcrlf
								end if
								urlList=urlList & url
							end if
						end if
					end if
					cssParamValue=""
					cssParamName=""										'��ղ�������
					tempCssParamName=""
					isMaoHao=false							'ð�ſ���
				 end if
				'call echo("s",s)
				c=c & saveRowC & s						'�ۼ�����������
				saveRowC=""								'��ձ���ֵ
				rowC="" 								'��յ�ǰֵ
			elseif s="}" then
				if instr(sType,"|unzip|")>0 then
					s= vbcrlf & s & vbcrlf
				end if
				c=c & s								'���  width:12p;}  ���֣���Ϊ }ǰ��û��ֵ�����}�����������������Ҫ����Ҫ�ж���20161002
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
		'get�����û�л��ֵ���򷵻ؿ�
		elseif instr(sType,"|get|")>0 or instr(sType,"|getimage|")>0 then
			handleCssContent=""
		
		'�ݴ棬ɾ��������20161024
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