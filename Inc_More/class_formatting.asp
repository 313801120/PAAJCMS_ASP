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
'��ʽ��html  ����20170829
'dim t:  set t=new class_formatting
'call echo("",t.handleFormatting) 
class class_formatting 
	dim pubContent				'����
	dim nCountLen				'�ַ�����
	dim beforeStr				'��һ���ַ�
	dim afterStr				'��һ���ַ�
	dim isDisplayEcho			'�Ƿ����
	dim islabel					'�Ƿ�Ϊ��ǩ
	dim labelLevelName(29999,11)		'��ǩ�ȼ�����    ��ǩ����  ���λ�� �ַ�����  �ұ�λ��  ��ǩ���� ��ǩ�������ݣ���ǩ�������� 9Ϊ����
	dim nLevel,tempNLevel                 	'������
	dim nLabelIndex				'��ǩ����
	dim imgList					'ͼƬ�б�
	dim cssList					'CSS�б�
	dim jsList					'JS�б�		
	dim swfList					'SWF�б�	
	dim sourceList				'source�б� ��Ƶ 
	dim videoList				'video�б� ��Ƶ 
	dim aList					'a�б� ���ӵ�ַ20171207 
		 
	dim urlList					'���ز���ַ�б�
    dim nRowCount				'������

	 '���캯�� ��ʼ��
    Sub Class_Initialize() 
		isDisplayEcho=false
		islabel=false			'�Ƿ�Ϊ��ǩ Ĭ��Ϊ��
	 	nLevel = 0              '������
		nLabelIndex=0			'��ǩ����Ĭ��Ϊ��
	end sub
    '�������� ����ֹ
    Sub Class_Terminate() 
    End Sub 
	'�����ʽ��
	function handleFormatting(content)
		pubContent=content
		dim i,s,s2,labelS,labelName,parentLableName,endCode,nLeft,nRight,nLen,isOK,parentObjId,nGetLeft,lableC,lableNRow 
		dim nNoteLeft,nNoteRight,isHtmlNote:isHtmlNote=false				'�Ƿ�Ϊhtmlע��
		dim isHtmlStyle:isHtmlStyle=false			'�Ƿ�Ϊstyle��ʽ
		dim isHtmlStyleNote:isHtmlStyleNote=false	'�Ƿ�Ϊstyle��ʽע��
		dim isHtmlJs:isHtmlJs=false					'�Ƿ�Ϊjs
		dim isHtmlJsNote:isHtmlJsNote=false			'�Ƿ�Ϊjsע��
		dim jsWriteLabel			    			'Ϊjs�����ǩ����  ��'��"
		dim brStr									'����ֵ
		brStr=chr(10)				'�����������vbcrlf  ����ͨ����ֻ��chr(10)����
		'if instr(content,vbcrlf)>0 then
		'	brStr=vbcrlf
		'end if
		'brStr=chr(10)
		'call echo("brStr",len(brStr))
		
		nCountLen=len(content)
		for i = 1 to nCountLen
			s = mid(content, i, 1) 
			'Ϊhtmlע�� ����ʱ
			if isHtmlNote=true then
				if mid(content,i,3)="-->" then
					isHtmlNote=false				'htmlע��Ϊ��
					i=i+2
					'ע��Ҳ�ŵ���ǩ��������20161024
					nNoteRight=i-nNoteLeft+1
					labelS=mid(content, nNoteLeft, nNoteRight) 
					nLabelIndex=nLabelIndex+1
					labelLevelName(nLabelIndex,0)=nLevel					
					labelLevelName(nLabelIndex,1)="!--htmlnote--"
					labelLevelName(nLabelIndex,2)=nNoteLeft
					labelLevelName(nLabelIndex,3)=len(labelS)
					labelLevelName(nLabelIndex,4)=i+1
					labelLevelName(nLabelIndex,5)=labelS
					labelLevelName(nLabelIndex,6)=""			'��ǩǰ������
					labelLevelName(nLabelIndex,7)=""			'��ǩ��������
					labelLevelName(nLabelIndex,9)=""			'��ǩ���� 
				end if
			'�ж��Ƿ�Ϊhtml ע��
			elseif mid(content,i,4)="<!--"  then
				nNoteLeft=i
				isHtmlNote=true				'htmlע��Ϊ�� 
			'��ʽע�Ϳ�ʼ
			elseif isHtmlStyle=true and isHtmlStyleNote=false and s="/" and mid(content & " ", i+1, 1)="*" then
				i=i+1
				isHtmlStyleNote=true
			'��ʽע�ͽ���
			elseif isHtmlStyle=true and isHtmlStyleNote=true and s="*" and mid(content & " ", i+1, 1)="/" then
				i=i+1
				isHtmlStyleNote=false 
			elseif isHtmlStyle=true and isHtmlStyleNote=true  then
				'Ϊcss��ʽ���ע�ͣ���������
				
			'JS ����ע��
			elseif isHtmlJs=true and isHtmlJsNote=false and s="/" and mid(content & " ", i+1, 1)="/" and jsWriteLabel="" then
				s2=mid(content,i)
				i=i+instr(s2,brStr)  
			'JS ����ע�Ϳ�ʼ
			elseif isHtmlJs=true and isHtmlJsNote=false and s="/" and mid(content & " ", i+1, 1)="*" and jsWriteLabel="" then
				i=i+1
				isHtmlJsNote=true 
			'JS ����ע�ͽ���
			elseif isHtmlJs=true and isHtmlJsNote=true and s="*" and mid(content & " ", i+1, 1)="/" then
				i=i+1
				isHtmlJsNote=false 
			elseif isHtmlJs=true and isHtmlJsNote=true  then
				'JS ����ע�����ע�ͣ���������
				
			'js��������ݴ���20161115
			elseif isHtmlJs=true and instr("\'""",s)>0 then
				'�����ǩ���
				if s="\" then
					i=i+1
				elseif jsWriteLabel<>"" and jsWriteLabel=s then
					jsWriteLabel=""
				elseif s="'" or s="""" then
					jsWriteLabel=s 
				end if 
			elseif s="<"  then 
				endCode=mid(content,i)
				'���һ����ǩû��>����ϣ�Ҫ�������Σ�2016116   ����д������Լ������ˣ�ֻ�ܰѴ������Ժ����������
				if instr(endCode,">")=false then
					endCode=endCode & ">"
				end if
				labelS=mid(endCode,1,instr(endCode,">"))
				nLeft=i
				nLen=len(labelS)
				nRight=i+nLen
				nLabelIndex=nLabelIndex+1
				labelName=getLabelName(labelS)				'��ǩ����
				'��һ����ǩ����
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
					jsWriteLabel=""			'js�����ǩ�������õ�
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
						'��������һ����ǩ�����ټ�һ
						if labelName <> "/" & parentLableName  and nLevel>0 then		'and left(parentLableName,1)<>"/"  ��Ҫ�ж���һ���Ƿ���</
							tempNLevel=getEndLabelToStartLable(nLabelIndex, nLevel, labelName,parentObjId)		'�����һ���ȼ���
							'call echoyellowb("tempNLevel",tempNLevel)
							if tempNLevel<>-1 then
								nLevel=tempNLevel 
								nGetLeft=labelLevelName(parentObjId,2)
								lableC=mid(content,nGetLeft,nRight-nGetLeft)
								lableNRow=ubound(split(lableC,brStr))
								
								labelLevelName(nLabelIndex,8)=lableNRow			'��ǩ�м������ж����� ������ǩ
								labelLevelName(parentObjId,8)=lableNRow			'��ǩ�м������ж����� ��ʼ��ǩ 
							else  
								nLevel=nLevel-1
							end if
							labelLevelName(nLabelIndex,0)=nLevel
						'������ж���Ϊ���ռ���ʼ��ǩ�������ǩ֮�����ݵ�����
						elseif left(labelName,1)="/" then
							tempNLevel=getEndLabelToStartLable(nLabelIndex, nLevel, labelName,parentObjId)		'�����һ���ȼ��� 
							if tempNLevel<>-1 then 
								nGetLeft=labelLevelName(parentObjId,2)
								lableC=mid(content,nGetLeft,nRight-nGetLeft)
								lableNRow=ubound(split(lableC,brStr))
								
								labelLevelName(nLabelIndex,8)=lableNRow			'��ǩ�м������ж����� ������ǩ
								labelLevelName(parentObjId,8)=lableNRow			'��ǩ�м������ж����� ��ʼ��ǩ 
							end if	
						end if
						
					elseif right(labelS,2)<>"/>" then
						if ucase(labelName)<>"!DOCTYPE" then
							nLevel=nLevel+1
						end if
						'call echo("��" & nLevel,replace(labelS,"<","&lt;"))
					end if
					
					labelLevelName(nLabelIndex,1)=labelName
					labelLevelName(nLabelIndex,2)=nLeft
					labelLevelName(nLabelIndex,3)=nLen
					labelLevelName(nLabelIndex,4)=nRight
					labelLevelName(nLabelIndex,5)=labelS
					labelLevelName(nLabelIndex,6)=""			'��ǩǰ������
					labelLevelName(nLabelIndex,7)=""			'��ǩ��������
					labelLevelName(nLabelIndex,9)=""			'��ǩ����
				end if
			end if 
			
			
		next
		
		handleFormatting=content
	end function
	'��ñ�ǩ���Գ��ֵĿ�ʼλ��
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
				parentObjId=nI						'��ǰ������Ҳ����ID
				getEndLabelToStartLable=nLevel
				exit function
			end if 
			'call echoYellowB(i, findLabelName & "," & labelName)
		next
		getEndLabelToStartLable=-1
	end function 
	'��ñ�ǩ����
	function getLabelName(byval labelS)
		dim labelName
		labelName=replace(replace(labelS,">"," "),vbtab," ")
		labelName=phptrim(mid(labelName,2,instr(labelName," ")-1)) 
		getLabelName=labelName
	end function
	'��ӡ
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
			s=s & "("& lableNRow &")��"
			call echoredb( copyStr("&nbsp;&nbsp;",nLevel) & nLevel & "|" & labelName, nLeft & "," & nRight & " , " & s)
			if instr("|/div|/span|", "|" & labelName & "|")>0 then
				call rw("<"& labelName &">")
			end if
		next
		call rw("<script src='/Jquery/Jquery.Min.js' type='text/javascript'></script>")
		call rw("<script src='/Inc_More/z_formatting.js' type='text/javascript'></script>")
		call rw(copyStr("<br>" & vbcrlf ,10))
	end function
	
	'��ø�ʽ��HTML 
	function getFormattingHtml(byval sHtmlType, byval sJsType, byval sCssType)
		dim i,labelName,parentLabelName,nLevel,nLeft,nLen,nRight,s,c,labelBeforeStr,labelAfterStr,parentNRight,labelS,s2,labelNRow,nLen2
		dim action,parentAction			'��ǰ��������һ����ǩ����
		dim isPre:isPre=false			'Ĭ��PreΪ��
		sHtmlType="|"& sHtmlType &"|"			'html��������
		sJsType="|"& sJsType &"|"			'js��������
		sCssType="|"& sCssType &"|"			'css��������
		for i =1 to nLabelIndex
			parentLabelName=labelName
			nLevel=labelLevelName(i,0)
			labelName=labelLevelName(i,1)
			nLeft=labelLevelName(i,2)
			nLen=labelLevelName(i,3)
			nRight=labelLevelName(i,4)
			labelS=labelLevelName(i,5)
			labelBeforeStr=labelLevelName(i,6) 	'��ǩǰ������
			labelAfterStr=labelLevelName(i,7) 		'��ǩ��������
			labelNRow=labelLevelName(i,8)			'��ǩ��������
			action = "|" & labelLevelName(i,9) & "|" 		'��ǰ����  
			parentAction = "|" & labelLevelName(i-1,9) & "|" 		'��һ����ǩ����  
			parentNRight=labelLevelName(i-1,4) 						'��һ���ұ��ַ�λ��
			 
			if parentLabelName="pre" then
				isPre=true 
			elseif parentLabelName="/pre" then
				isPre=false
			end if
			
			
			'��һ����ǩ
			if i=1 then
				if nLeft>1 then
					s=mid(pubContent,1,nLeft-1)
					'����������߿ո�
					if instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 then
						s=phptrim(s)
					end if
					c=c & s
				end if
				'��������ɾ��������
				if instr("|del|",action)=false then
					c=c & labelBeforeStr			'��ǩǰ�������
					c=c & labelS				'���Ȿ��
					c=c & labelAfterStr			'��ǩ���������
				end if
				 
			else 
				if instr("|replace|del|",parentAction)>0 then
					'Ϊɾ�����滻�������� 
				else 
					s2=mid(pubContent,parentNRight,nLeft-parentNRight)
					'ѹ��JS
					if  instr(sJsType, "|zip|")>0  and labelName="/script" then 
						s2=pubJsObj.handleJsContent(s2, sJsType)		'ɾ��js��ע��
					'ѹ��CSS
					elseif  instr(sCssType,"|zip|")>0   and labelName="/style" then
						s2=pubCssObj.handleCssContent("css",s2,"*","","","", sCssType)
					'����JS
					elseif   instr(sJsType, "|unzip|")>0  and labelName="/script" then
						s2=pubJsObj.handleJsContent(s2, sJsType)
					'����CSS
					elseif instr(sCssType, "|unzip|")>0  and labelName="/style" then
						s2=pubCssObj.handleCssContent("css",s2,"*","","","", sCssType)
						
						
					'����������߿ո�
					elseif ( instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 ) and isPre=false and labelName<>"/pre" then 
						s2=phptrim(s2)
					end if
					'��ѹ����
					if instr(sHtmlType, "|unzip|")>0 and isPre=false then
						 
						if labelNRow>0 then
							labelS=vbcrlf & "" &  copyStr("    ",nLevel ) & labelS
						elseif len(s2)=0 and "/" & parentLabelName<>labelName and instr("|br|hr|","|"& labelName &"|")=false then
							'call echoRedB(parentLabelName,labelName)
							labelS=vbcrlf & "" &  copyStr("    ",nLevel ) & labelS  
						end if
						if labelName="/style" and labelName="/script" then
							s2=handleHtmlTab(s2,nLevel)			'����Html��Tab�˸�
						end if
					end if
					'ɾ��HTMLע��
					if instr(sHtmlType, "|delnote|")>0 and labelName="!--htmlnote--" then
						labelS=""
					end if
					c=c & s2 					'��һ����ǩ������
				end if				
				c=c & labelBeforeStr			'��ǩǰ�������
				'��ǰ��ǩ��Ϊɾ��
				if instr("|del|",action)=false then
					c=c & labelS				'���Ȿ��
				end if 
				c=c & labelAfterStr			'��ǩ��������� 
			end if
			'����ǩ��׷�����ʣ�µ�����
			if i = nLabelIndex then 
				nLen2=nCountLen-nRight+1
				if nLen2<0 then nLen2=0 			'����Ϊ���� 20171113  ��������� a<b 
				s=mid(pubContent,nRight,nLen2)
				'����������߿ո�
					if instr(sHtmlType, "|zip|")>0 or instr(sHtmlType, "|unzip|")>0 or instr(sHtmlType, "|htmlzip|")>0 or instr(sHtmlType, "|htmlunzip|")>0 then
					s=phptrim(s)
				end if
				c=c & s
			end if
			
		next
		getFormattingHtml=c
	end function
	'����Html��Tab�˸� 20161025
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
	'�жϱ�ǩ����
	function checkLabelHtml(labelNameList) 
		checkLabelHtml=handleLabelContent(labelNameList,"*","","check")
	end function
	'��ñ�ǩ����
	function getLabelHtml(labelNameList) 
		getLabelHtml=handleLabelContent(labelNameList,"*","","get") 
	end function
	'���ñ�ǩ���� ����ʽ��
	function setLabelHtml(labelNameList,htmlValue)  
		call  handleLabelContent(labelNameList,"*",htmlValue,"set")	'�޸�
		setLabelHtml=getFormattingHtml("","","")		'��ø�ʽ�����html
	end function
	'���ñ�ǩ���� ����ʽ��
	function delLabelHtml(labelNameList)  
		call  handleLabelContent(labelNameList,"*","","del")	'�޸�
		delLabelHtml=getFormattingHtml("","","")		'��ø�ʽ�����html
	end function
'before() - �ڱ�ѡԪ��֮ǰ���� after() - �ڱ�ѡԪ��֮����� prepend() - �ڱ�ѡԪ�صĿ�ͷ���� append() - �ڱ�ѡԪ�صĽ�β����
	
	'�����ǩָ������   sType='set' Ϊ�滻html���� 
	function handleLabelContent(byval labelNameList, byval paramName, byval htmlValue, byval sType) 
		dim i,labelName,nLevel,nLeft,nLen,nRight,s,c,labelBeforeStr,labelAfterStr,parentNRight,labelS,lCaseLabelS,s2
		dim LeftS,RightS,nEQ,nCount,findLabelLevel(999,9),id,actionList,cList,action,tempHtmlValue
		dim nLevelEQ			'�ȼ�EQ
		dim isSetAddUrlEnc:isSetAddUrlEnc=false		'����ͼƬ�޼���
		 
		'����ͼƬ�޼���
		if instr(htmlValue,"setaddurlenc(*)")>0 then
			htmlValue=replace(htmlValue,"setaddurlenc(*)","")
			isSetAddUrlEnc=true
		end if
		'��ȡ����λ��
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
		'����
		actionList=getStrCut(labelNameList,"[","]",0)
		if actionList<>"" then
			labelNameList=replace(labelNameList,"["& actionList &"]","")
		end if
		'��ȡ����λ��
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
			labelBeforeStr=labelLevelName(i,6) 	'��ǩǰ������
			labelAfterStr=labelLevelName(i,7) 		'��ǩ��������
			parentNRight=labelLevelName(i-1,4)
			action="|" & labelLevelName(i,9) & "|"
			'���滻����ʱ�����ﴦ���ã���Ϊ���и������������ͻ�������ӣ���ô�죿
			if instr(action,"|del|delthis|")>0 then
				'Ϊɾ���򲻴�����
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
							'�����ӱ�ǩΪɾ������ ��ʼλ��+1����Ϊ�ų�����   ����λ��-1���ų����� 
							call setLabelAction(id+1,i-1,"del") 
						end if
					
					'before() - �ڱ�ѡԪ��֮ǰ����
					elseif sType="before" then
						id=findLabelLevel(nLevel,5)
						labelLevelName(id,6)=htmlValue
						
					'after() - �ڱ�ѡԪ��֮�����
					elseif sType="after" then
						labelLevelName(i,7)=htmlValue
						
					'prepend() - �ڱ�ѡԪ�صĿ�ͷ���� 
					elseif sType="prepend" then
						id=findLabelLevel(nLevel,5)
						labelLevelName(id,7)=htmlValue
					 
					'append() - �ڱ�ѡԪ�صĽ�β����
					elseif sType="append" then
						labelLevelName(i,6)=htmlValue 
					
					'���
					elseif sType="check" then
						c=true
					elseif sType="get+label" then
						c= findLabelLevel(nLevel,3) &  s & labelS
						
					elseif sType="handleimg"  then 
						
						s2=mid(pubContent,parentNRight,nLeft-parentNRight) '���Լ�src�ˣ���Ϊ�Ҵ������
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
					elseif sType="html"  then 			'��ǩ����
						c=s
					else
						c=s
						'call echo("�ҵ�","")
					end if
					'Ϊ����
					if cList<>"" then
						cList=cList & "$Array$"
					end if
					cList=cList & c
					handleLabelContent=cList
					findLabelLevel(nLevel,4)=""
					'Ϊָ��һ��
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
							'Ϊ����
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
						'type=application/x-shockwave-flash Ϊflash�ж�
						'�Ӹ��ж���Ϊ�˼ӿ� �ʣ��б����𣿴𣺴�����ѽ
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
						'css  �ڶ��δ���  ������labelS ��Ϊ���������п��ܴ�����  ���ף�
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
							'˵�����������src��ǩ����
							if c<>"" then 
								labelLevelName(i,5)=pubHtmlObj.handleHtmlLabel(labelS,"script","src","","set",tempHtmlValue)
								if instr(vbcrlf & jsList & vbcrlf, vbcrlf & c & vbcrlf)=false and c<>"" then
									if jsList<>"" then
										jsList=jsList & vbcrlf
									end if
									jsList=jsList & c
								end if
							end if
						'jsд���ڲ�
						else
							'call eerr("555555555555",labelLevelName(i+1,5))
							'call eerr("��ʾ1111111",mid(pubContent,labelLevelName(i,4),labelLevelName(i+1,4))) 
						end if
					'׷����20171013
					elseif labelName="source" then
						if instr(lCaseLabelS,"src")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"source","src","","get",htmlValue)
							'˵�����������src��ǩ����
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
						
					'׷����20171019
					elseif labelName="video" then
						if instr(lCaseLabelS,"src")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"video","src","","get",htmlValue)
							'˵�����������src��ǩ����
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
					'׷����20171207
					elseif labelName="a" then
						if instr(lCaseLabelS,"href")>0 then  
							c=pubHtmlObj.handleHtmlLabel(labelS,"a","href","","get",htmlValue) 
							'˵�����������src��ǩ����
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
						nCount=nCount+1					'��������ֵ�ۼ�  
						findLabelLevel(nLevel,4)="OK" 
						findLabelLevel(nLevel,5)=i			'����
					end if 
					'����Ӹ�style�жϣ��ǳ��б�Ҫ�ģ����
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
					
					
				'�ж� ��Ҫ
				elseif pubHtmlObj.handleHtmlLabel(labelS,labelNameList,paramName,actionList,"check","")=true then 
					'call echo("labelName",labelName)
					
					findLabelLevel(nLevel,0)=nLevel
					findLabelLevel(nLevel,1)=labelName
					findLabelLevel(nLevel,2)=nRight
					findLabelLevel(nLevel,3)=labelS
					nCount=nCount+1					'��������ֵ�ۼ� 
					if (nCount=nEQ or nEQ=-1) and (nLevelEQ=nLevel or nLevelEQ=-1) then		'׷���˵ȼ��ж�
						findLabelLevel(nLevel,4)="OK"
					end if 
					findLabelLevel(nLevel,5)=i			'����
					
					'����ǩ����ʱֻ����ɾ��
					if findLabelLevel(nLevel,4)="OK" and instr("|img|meta|link|param|hr|br|", "|" & labelName & "|")>0 then
						if sType="del" then		   
							labelLevelName(i,9)="del"
						
						'�Ե���ǩֵ���  ��ʱֻ�ӻ��
						elseif sType="get" then 
							'Ϊ����
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
		'����
		if sType="count" then
			handleLabelContent=nCount
		end if
	end function
	'����ָ�������ǩ����
	function setLabelAction(nStart,nEnd,actionStr)
		dim i
		for i = nStart to nEnd
			labelLevelName(i,9)=actionStr							
		next						
	end function 
	
	'���ͼƬ/css/js��ַ�б�
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
	'�������
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
	'����
	function updateHtml()
		dim c
		c=getFormattingHtml("","","")
		call clearArray()
		c=handleFormatting(c)
		updateHtml=c 
	end function
	
	
	'ɾ��HTML����д�����ַ�����̫��
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
	 
	'������һ���ַ�����һ���ַ�
	sub handleNRowCountAndBeforeStrAndAfterStr(content,s,i)
		if s=chr(10)   then
			nRowCount=nRowCount+1
		end if
		beforeStr=""
		if i>1 then
			beforeStr = mid(content, i-1, 1)                        '��һ���ַ�
		end if
		afterStr =""
		if i+1<len(content) then
			afterStr = mid(content, i+1, 1)       				'��һ���ַ�  
		end if
	end sub
	
	
	
	'*******************************************************************************
	'�ҵ�ָ��λ�õı�ǩ
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
			labelBeforeStr=labelLevelName(i,6) 	'��ǩǰ������
			labelAfterStr=labelLevelName(i,7) 		'��ǩ��������
			labelNRow=labelLevelName(i,8)			'��ǩ��������
			action = "|" & labelLevelName(i,9) & "|" 		'��ǰ����  
			parentAction = "|" & labelLevelName(i-1,9) & "|" 		'��һ����ǩ����  
			parentNRight=labelLevelName(i-1,4) 						'��һ���ұ��ַ�λ��
			
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