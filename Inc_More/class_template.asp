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
'����html��ģ��
'dim t:set t=new class_template 
class class_template
	dim sTest
	dim DefaultTemplateFolder			'Ĭ��ģ���ļ���·��
	dim DefaultTemplateHttpUrl
	 '���캯�� ��ʼ��
    Sub Class_Initialize()
		sTest=now() 
		DefaultTemplateFolder="\DataDir\VBģ��\������\Template\"
		DefaultTemplateHttpUrl="http://vb/DataDir/VBģ��/������/Template/"
	end sub
    '�������� ����ֹ
    Sub Class_Terminate()
        'HtmlFolder=nothing
        'HtmlFilename=nothing
        'HtmlContent=nothing
        'Urlname=nothing
    End Sub
	
	'�������������б�
	function batchHandleWebNavLayout(byval content, byval dataList)
		dim splstr,s
		splstr=split(dataList,"$Array$")
		for each s in splstr
			s=phptrim(s)
			if s<>"" then
				content=handleWebNavLayout(content,s) 
			end if
		next
		batchHandleWebNavLayout=content
	end function
	'navC=pubFormatObj.handleLabelContent("ul[**class=naV]","*","","get+label"))
	'���������б�  
	function handleWebNavLayout(byval content, byval navC) 
		dim sLabelName,listS,listC,s,nIndex,splxx,replaceId, delIdList,nTopNumb
		dim fObj : set fObj=new class_formatting
		 
		navC=fObj.handleFormatting(navC) 			'��ʽ����������
		sLabelName=fObj.findLabelContent(1)			'���ָ��λ���ϵı�ǩ�б�
		listS=fObj.handleLabelContent("li:eq(0)","*","","get+label")
		listC=fObj.handleLabelContent("li","*","","get+label")
		nTopNumb=0
		listS = handleLi(listC,replaceId, delIdList,nTopNumb)
		
		dim f2Obj : set f2Obj=new class_formatting 
		listS=f2Obj.handleFormatting(listS) 
		call f2Obj.setLabelHtml("a","[$columnname$]")		'����A���ӵ�html
		listS=f2Obj.getFormattingHtml( "","","" )		'���list����
		
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","href","","edit","[$url$]")		'�滻href
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","title","","edit","[$columnname$]")	'�滻title
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","alt","","edit","[$columnname$]")		'�޸�alt
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","target","","edit","[$target$]")		'�޸�target
		'call echoBlue("listS",replace(listS,"<","&lt;"))
		
		splxx=split(delIdList,"|")
		for each nIndex in splxx
			if nIndex <>"" then
				call fObj.handleLabelContent("li:eq("& nIndex &")","*","","del")	 'ɾ��li
			end if
		next
		if replaceId<>"" then	
			call fObj.handleLabelContent("li:eq("& replaceId &")","*","","delthis")	 'ɾ��li	
			call fObj.handleLabelContent("li:eq("& replaceId &")","*",getNavTemplateLayout(listS),"before")		'�滻ģ������
		end if
		'call fObj.handleLabelContent("li:leveleq(1)","*","","delthis")							'ɾ���ȼ�Ϊ1��ȫ��li
		'call fObj.handleLabelContent("li:eq(0)","*",getNavTemplateLayout(listS),"before")		'�滻ģ������
		s=fObj.getFormattingHtml( "","","" )			'��ø�ʽ��������� 
		handleWebNavLayout=replace(content,navC,s)
		
		'call echo("replaceId",replaceId)
		'call echo("delIdList",delIdList)
	
	end function
	'�������������б�
	function batchHandleWebNewsLayout(byval content, byval dataList)
		dim splstr,s
		splstr=split(dataList,"$Array$")
		for each s in splstr
			s=phptrim(s)
			if s<>"" then
				content=handleWebNewsLayout(content,s) 
			end if
		next
		batchHandleWebNewsLayout=content
	end function
	'navC=pubFormatObj.handleLabelContent("div[**class=columnlist2]","*","","get+label"))
	'���������б�
	function handleWebNewsLayout(byval content, byval navC) 
		dim sLabelName,listS,listC,s,nIndex,splxx,replaceId, delIdList,nTopNumb,columnName
		dim fObj : set fObj=new class_formatting
		splxx=split(phpTrim(delHtml(navC)),vbcrlf)
		columnName=splxx(0) 
		navC=fObj.handleFormatting(navC) 			'��ʽ����������
		sLabelName=fObj.findLabelContent(1)			'���ָ��λ���ϵı�ǩ�б�
		'call echo("sLabelName",sLabelName)
		'listS=fObj.handleLabelContent("li:eq(0)","*","","get+label")
		listC=fObj.handleLabelContent("li","*","","get+label")
		nTopNumb=0
		listS = handleLi(listC,replaceId, delIdList,nTopNumb)
		
		dim f2Obj : set f2Obj=new class_formatting 
		listS=f2Obj.handleFormatting(listS) 
		call f2Obj.setLabelHtml("a","[$title$]")		'����A���ӵ�html
		listS=f2Obj.getFormattingHtml( "","","" )		'���list����
		
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","href","","edit","[$url$]")		'�滻href
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","title","","edit","[$title$]")	'�滻title
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","alt","","edit","[$title$]")		'�޸�alt
		listS=pubHtmlObj.handleHtmlLabel(listS,"*","target","","edit","[$target$]")		'�޸�target
		'call echoBlue("listS",replace(listS,"<","&lt;"))
		
		splxx=split(delIdList,"|")
		for each nIndex in splxx
			if nIndex <>"" then
				call fObj.handleLabelContent("li:eq("& nIndex &")","*","","del")	 'ɾ��li
			end if
		next
		if replaceId<>"" then	
			call fObj.handleLabelContent("li:eq("& replaceId &")","*","","delthis")	 'ɾ��li	
			call fObj.handleLabelContent("li:eq("& replaceId &")","*",getNewsTemplateLayout(listS,columnName,nTopNumb),"before")		'�滻ģ������
		end if
		'call fObj.handleLabelContent("li:leveleq(1)","*","","delthis")							'ɾ���ȼ�Ϊ1��ȫ��li
		'call fObj.handleLabelContent("li:eq(0)","*",getNavTemplateLayout(listS),"before")		'�滻ģ������
		s=fObj.getFormattingHtml( "","","" )			'��ø�ʽ��������� 
		
		handleWebNewsLayout=replace(content,navC,s)
		'handleWebNewsLayout=s
		'call echo("replaceId",replaceId)
		'call echo("delIdList",delIdList)
		'call echoRedB("nTopNumb",nTopNumb)
	
	end function
	
	'����LI
	function handleLi(listC, replaceId, delIdList,nTopNumb)
		dim splstr,i,s,s2,c,listS,html,isLi,liEndS
		isLi=true 
		splstr=split(listC,"$Array$")
		for i=0 to ubound(splstr)
			s=splstr(i)
			s2=phpTrim(s)
			html=phpTrim(replace(delHtml(s),"&nbsp;",""))
			if (i=0 or i=ubound(splstr)) and html="" then
				'call echoRedB("��Ҫɾ��", replace(s,"<","&lt;")) 
			else
				
				if liEndS=s2 or html<>"" then
					delIdList=delIdList & i & "|"
					if html<>"" then
						nTopNumb=nTopNumb+1
					end if
				else
				end if
				'����б�li
				if listS="" then
					listS=s
					replaceId=i
					delIdList=""
					delIdList=delIdList & i & "|"
				elseif listS<>"" and isLi=true then
					if html="" then
						listS=listS & s  
					end if
					liEndS=s2			'li�����ַ�
					isLi=false 
					delIdList=delIdList & i & "|"
				end if
				'call echo(i & "��" & html, replace(s,"<","&lt;")) 
			end if
			
		next
		'call echoYellowB("delIdList",delIdList) 
		handleLi=listS
	end function
	
	'��õ���ģ������
	function getNavTemplateLayout(listS)
		dim c,id
		id=getRand(5) & " " & format_Time(now(),12)
		c="{$ColumnList flags='top' addsql='parentid=-1 order by sortrank'  default='[_"& id &"]'$}" & vbcrlf
		c=c & "<!--#[_"& id &"]" & vbcrlf
		c=c & "[list]" & listS & vbcrlf
		c=c & "[/list]" & vbcrlf
		'c=c & "[list-focus]<li><a href='[$url$]' title='[$labletitle$]'[$atarget$] class='focus'>[$columnname$]</a></li>" & vbcrlf
		'c=c & "[/list-focus]" & vbcrlf
		c=c & "#-->" & vbcrlf
		
		getNavTemplateLayout=c
	end function
	'�������ģ������
	function getNewsTemplateLayout(listS,columnName,nTopNumb)
		dim c,id,columnNameSql
		if columnName<>"" then
			columnNameSql=" columnname='"& columnName &"'"
		end if
		id=getRand(5) & " " & format_Time(now(),12)
		c="{$ArticleList"& columnNameSql &" topnumb='"& nTopNumb &"' default='[_"& id &"]'$}" & vbcrlf
		c=c & "<!--#[_"& id &"]" & vbcrlf
		c=c & "[list]" & listS & vbcrlf
		c=c & "[/list]" & vbcrlf 
		c=c & "#-->" & vbcrlf 
		getNewsTemplateLayout=c
	end function	
 
 


'============================== VB���� =================================
 
 
 
 '���Top��ʽ
Sub GetTopCss(id, cssStyle, DivHtml)
    Dim c, path, ImgWidth, ImgHeight, FFolder, fName
    'Css����
    cssStyle = cssStyle & "#header" & id & "{" & vbCrLf
    cssStyle = cssStyle & "    width:auto;" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Top\" & id & "\Index"
    fName = findDirFileName(FFolder, "Logo")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Top/" & id & "/Index/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "}" & vbCrLf
    'Html����
    DivHtml = DivHtml & "    <div id=header" & id & ">" & vbCrLf
    DivHtml = DivHtml & "    </div>" & vbCrLf
End Sub
'��ʾĬ�ϵ���
Function NavDefaultValue()
    Dim splStr, i, s, c
    c = c & "<li><a href=javascript:;>��ҳ</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��˾���</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��Ʒչʾ</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��������</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��������</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��˾��Ƹ</a></li><li class=line></li>" & vbCrLf
    c = c & "<li><a href=javascript:;>��ϵ����</a></li><li class=line></li>" & vbCrLf
    NavDefaultValue = c
End Function
'���ָ��������������ʽ 20150626
Function getNavCssStyleCount(id)
    Dim content, splStr, FFolder, nOK, s
    FFolder = DefaultTemplateFolder & "\Nav\" & id
    content = getFText(FFolder & "/Nav.txt")
    splStr = Split(content, vbCrLf)
    nOK = -1
    For Each s In splStr
        If phpTrim(s) <> "" Then
            nOK = nOK + 1
        End If
    Next
    getNavCssStyleCount = nOK
End Function
'���ָ��������������ʽ 20150626
Function getNavIDCssStyleContent(id, parameter)
    Dim content, splStr, FFolder, nOK, s
    FFolder = DefaultTemplateFolder & "\Nav\" & id
    content = getFText(FFolder & "/Nav.txt")
    splStr = Split(content, vbCrLf)
    nOK = 0
    For Each s In splStr
        If phpTrim(s) <> "" Then
            If nOK = parameter Then
                getNavIDCssStyleContent = s
                Exit Function
            End If
            nOK = nOK + 1
        End If
    Next
End Function

'���Nav��ʽ IDΪ��ʽ���� ParameterΪ����
Public Sub GetNavCss(id, cssStyle, DivHtml, ByVal stylevalue)
    Dim c, FFolder, fName, ImgWidth, ImgHeight, content, AWidth, AHeight, NavBackground, NavListBgColor, NavFontHeight, AMargin
    Dim splStr, DefaultColor, MoveColor, FontSize, FontFamily, FontWeight, SidLiBorderTop, tempS
    Dim Apadding, NavValueRepeat, CssLeftPadding
    Dim CopyOK                          '�����ļ��Ƿ�
    Dim NavClassHeightYes               '������ʽ���Ƿ�Ϊ�棬Ϊ��������Զ���߶�   20150317
    Dim cssNameAddId                    'Css����׷��Id���
    
    DefaultColor = "#000000"            'Ĭ��������ɫ
    MoveColor = "#FFFFFF"               '�ƶ�������ɫ
    FontSize = 12                       'Ĭ�����ִ�С
    AWidth = 0                          'Ĭ�����ӿ�
    AHeight = 0                         'Ĭ�����Ӹ�
    NavBackground = ""                  '��������
    NavListBgColor = ""                 '�����б�����
    NavFontHeight = ""                  '��������߶�
    AMargin = ""                        '����LIֵ
    NavValueRepeat = ""                 '������ǰѡ�����ݱ����Ƿ��ظ� 20150213
    CssLeftPadding = ""                   'CssLeft Padding��ʽ����
     
    
    FFolder = DefaultTemplateFolder & "\Nav\" & id
    
    'Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    If checkStrIsNumberType(stylevalue) Then
        cssNameAddId = "_" & stylevalue                      'Css����׷��Id���
        stylevalue = getNavIDCssStyleContent(id, CInt(stylevalue))
    End If

    
    'Ĭ����ɫ
    tempS = rParam(stylevalue, "DefaultColor")
    If tempS <> "" Then DefaultColor = tempS
    '������ɫ
    tempS = rParam(stylevalue, "MoveColor")
    If tempS <> "" Then MoveColor = tempS
    '���ִ�С
    tempS = rParam(stylevalue, "FontSize")
    If tempS <> "" Then FontSize = tempS
    '��������
    tempS = rParam(stylevalue, "font-family")
    If tempS <> "" Then FontFamily = tempS
    '����������Ӵ�
    tempS = rParam(stylevalue, "FontWeight")
    If tempS <> "" Then FontWeight = tempS
    '���ӿ���
    tempS = rParam(stylevalue, "AWidth")
    If tempS <> "" Then AWidth = tempS
    '��������
    tempS = rParam(stylevalue, "NavBackground")
    If tempS <> "" Then NavBackground = tempS
    '�����б���ɫ
    tempS = rParam(stylevalue, "NavListBgColor")
    If tempS <> "" Then NavListBgColor = tempS
    '�����ı��߶�
    tempS = rParam(stylevalue, "NavFontHeight")
    If tempS <> "" Then NavFontHeight = tempS
    '��������A�߿�ֵ
    tempS = rParam(stylevalue, "AMargin")
    If tempS <> "" Then AMargin = tempS
    '��������A�߿�ֵ
    tempS = rParam(stylevalue, "Apadding")
    If tempS <> "" Then Apadding = tempS
    '�˵�С��Li�����߿�
    tempS = rParam(stylevalue, "SidLiBorderTop")
    If tempS <> "" Then SidLiBorderTop = tempS
    '����ѡ�б����ײ�������ʾ
    tempS = rParam(stylevalue, "navvaluerepeat")
    If tempS <> "" Then NavValueRepeat = tempS
    'CssLeft padding����
    tempS = rParam(stylevalue, "cssleftpadding")
    If tempS <> "" Then CssLeftPadding = tempS
    
     
    '�ɰ治���ã�����
    'SplStr = Split(styleValue, "|")
    'If UBound(SplStr) >= 0 Then DefaultColor = IIf(SplStr(0) <> "", SplStr(0), DefaultColor)
    'If UBound(SplStr) >= 1 Then MoveColor = IIf(SplStr(1) <> "", SplStr(1), MoveColor)
    'If UBound(SplStr) >= 2 Then FontSize = IIf(SplStr(2) <> "", SplStr(2), FontSize)
    'If UBound(SplStr) >= 3 Then AWidth = IIf(SplStr(3) <> "", SplStr(3), AWidth)
    'If UBound(SplStr) >= 4 Then NavBackground = IIf(SplStr(4) <> "", SplStr(4), NavBackground)
    'If UBound(SplStr) >= 5 Then NavListBgColor = IIf(SplStr(5) <> "", SplStr(5), NavListBgColor)
    'If UBound(SplStr) >= 6 Then NavFontHeight = IIf(SplStr(6) <> "", SplStr(6), NavFontHeight)
    

    
    
    
    CopyOK = True
    If checkNumber(id) = False Then
        FFolder = handlePath(id) & "\"
        id = ""
        CopyOK = False
    End If
    'MsgBox (FFolder & CheckFolder(FFolder))
    
    '��վ������
    cssStyle = cssStyle & "/*����*/" & vbCrLf
    
    '�������
    cssStyle = cssStyle & "#nav" & id & cssNameAddId & "{" & vbCrLf
    
    
    fName = findSubDirFileName(FFolder, "NavLeft|NavRight")
    '������ұ�û����ʽ��
    If fName = "" Then
        If NavBackground <> "" Then cssStyle = cssStyle & "background-color:" & NavBackground & ";" & vbCrLf
        fName = findDirFileName(FFolder, "NavBg")
        If fName <> "" Then
            cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
            cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf     'XY���ظ� ���Զ��е���Ҳ����νѽ
        End If
    End If
    

    
    cssStyle = cssStyle & "}" & vbCrLf
    
    '��������
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & "{" & vbCrLf
    cssStyle = cssStyle & "    width:100%;" & vbCrLf
    '��������
    fName = findDirFileName(FFolder, "NavBg")
    NavClassHeightYes = False           '�������͸�Ĭ��Ϊ��
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf     'XY���ظ� ���Զ��е���Ҳ����νѽ
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
        cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
        NavClassHeightYes = True
    Else
        'û�б����Զ�Ѱ�ҵ�����
        fName = findSubDirFileName(FFolder, "NavMove|NavValue}NavLine")
        If fName <> "" Then
            Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
            If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
            cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
            cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
            NavClassHeightYes = True
        End If
        '�����б���
        If NavBackground <> "" Then
            cssStyle = cssStyle & "    background-color: " & NavBackground & ";" & vbCrLf
        End If
    End If
    '��ʱ20150317
    If NavFontHeight <> "" And NavClassHeightYes = False Then
        cssStyle = cssStyle & "    height:" & NavFontHeight & "px;" & vbCrLf
        cssStyle = cssStyle & "    line-height:" & NavFontHeight & "px;" & vbCrLf
    End If
    
    cssStyle = cssStyle & "    font-weight:" & FontWeight & ";" & vbCrLf                '�����������ǼӴ�
    cssStyle = cssStyle & "}" & vbCrLf
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li{float:left;position:relative;}" & vbCrLf   '���˸���Զ�λ20141111
    '�������
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .left{" & vbCrLf
 
    fName = findDirFileName(FFolder, "NavLeft")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    '����left���ص�
    Else
        cssStyle = cssStyle & "    width:0px;" & vbCrLf
        cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
        
        If CssLeftPadding <> "" Then
            cssStyle = cssStyle & "    padding:" & CssLeftPadding & vbCrLf
        End If
    End If
    cssStyle = cssStyle & "}" & vbCrLf
    '�����ָ���
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .line{" & vbCrLf
    
    fName = findDirFileName(FFolder, "NavLine")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    '����left���ص�
    Else
        cssStyle = cssStyle & "    width:0px;" & vbCrLf
        cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
    End If
    cssStyle = cssStyle & "}" & vbCrLf

    '�����ұ�
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .right{" & vbCrLf
    
    fName = findDirFileName(FFolder, "NavRight")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    '����left���ص�
    Else
        cssStyle = cssStyle & "    width:0px;" & vbCrLf
        cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:right;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��������
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " a{" & vbCrLf
    
    fName = findSubDirFileName(FFolder, "NavValue")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        If AWidth = 0 Then AWidth = ImgWidth                                     '�������ӿ�����
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        If FontFamily <> "" Then cssStyle = cssStyle & "    font-family:" & FontFamily & ";" & vbCrLf
    End If
    '��õ������
    fName = findSubDirFileName(FFolder, "NavValue|NavMove")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        If AWidth = 0 Then AWidth = ImgWidth                                        '�������ӿ�����
    End If
    
    cssStyle = cssStyle & "    font-size:" & FontSize & "px;" & vbCrLf
    cssStyle = cssStyle & "    color:" & DefaultColor & ";" & vbCrLf
    cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
    If AWidth <> "" And CStr(AWidth) <> "0" And rParam(stylevalue, "AWidth") <> "" Then
        cssStyle = cssStyle & "    width:" & minMaxBetween(90, 220, AWidth) & "px;" & vbCrLf      'һ��ʼ�жϴ��ˣ����Ǹ���ѽ
    End If
    cssStyle = cssStyle & "    height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
    '��������A�߿�ֵ
    If AMargin <> "" Then
        cssStyle = cssStyle & "    margin:" & AMargin & ";" & vbCrLf
    End If
    '��������A���ֵ
    If Apadding <> "" Then
        cssStyle = cssStyle & "    padding:" & Apadding & ";" & vbCrLf
    End If
    '����0
    If AHeight > 0 Or NavFontHeight <> "" Then
        cssStyle = cssStyle & "    line-height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    display:block;" & vbCrLf
    cssStyle = cssStyle & "    text-align:center;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '������������
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " a:hover{" & vbCrLf
    
    fName = findDirFileName(FFolder, "NavMove")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        '20150213
        If NavValueRepeat = "false" Then
            cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
            cssStyle = cssStyle & "    background-position: center bottom;" & vbCrLf
        'ͼƬ��С��20��ƽ��
        ElseIf ImgWidth < 20 Then
            cssStyle = cssStyle & "    background-repeat: repeat-x;" & vbCrLf
        Else
            cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        End If
        If FontFamily <> "" Then cssStyle = cssStyle & "    font-family:" & FontFamily & ";" & vbCrLf
    '������Ŀ�б��б���
    ElseIf NavListBgColor <> "" Then
        cssStyle = cssStyle & "    background-color: " & NavListBgColor & ";" & vbCrLf
    End If
    cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
    cssStyle = cssStyle & "    color:" & MoveColor & ";" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '���е�ǰ��ť ������20140905
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .focus{" & vbCrLf
    cssStyle = cssStyle & "    font-size:" & FontSize & "px;" & vbCrLf
    cssStyle = cssStyle & "    text-align:center;" & vbCrLf
    cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
    cssStyle = cssStyle & "    color:" & MoveColor & ";" & vbCrLf
    If AWidth <> "" And CStr(AWidth) <> "0" And rParam(stylevalue, "AWidth") <> "" Then
        cssStyle = cssStyle & "    width:" & minMaxBetween(90, 200, AWidth) & "px;" & vbCrLf
    End If
    If FontFamily <> "" Then cssStyle = cssStyle & "    font-family:" & FontFamily & ";" & vbCrLf
    cssStyle = cssStyle & "    height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
    '��������A�߿�ֵ
    If AMargin <> "" Then
        cssStyle = cssStyle & "    margin:" & AMargin & ";" & vbCrLf
    End If
    '��������A���ֵ
    If Apadding <> "" Then
        cssStyle = cssStyle & "    padding:" & Apadding & ";" & vbCrLf
    End If
    '����0
    If AHeight > 0 Or NavFontHeight <> "" Then
        cssStyle = cssStyle & "    line-height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
    End If
     
    
    fName = findSubDirFileName(FFolder, "NavMove|NavValue")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
        
        '20150213
        If NavValueRepeat = "false" Then
            cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
            cssStyle = cssStyle & "    background-position: center bottom;" & vbCrLf
        'ͼƬ��С��20��ƽ��
        ElseIf ImgWidth < 20 Then
            cssStyle = cssStyle & "    background-repeat: repeat-x;" & vbCrLf
        Else
            cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        End If
    '������Ŀ�б��б���
    ElseIf NavListBgColor <> "" Then
        cssStyle = cssStyle & "    background-color: " & NavListBgColor & ";" & vbCrLf
    End If
    cssStyle = cssStyle & "}" & vbCrLf
    '���е�ǰ��ť ��������ʽ 20141217
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .focus a{" & vbCrLf
    cssStyle = cssStyle & "    font-size:" & FontSize & "px;" & vbCrLf
    cssStyle = cssStyle & "    text-align:center;" & vbCrLf
    cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
    cssStyle = cssStyle & "    color:" & MoveColor & ";" & vbCrLf
    cssStyle = cssStyle & "    height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
    cssStyle = cssStyle & "    padding:0px;" & vbCrLf
    cssStyle = cssStyle & "    margin:0px;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    
    
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li ul {display:none;" & vbCrLf
    If NavBackground <> "" Then cssStyle = cssStyle & "background-color:" & NavBackground & ";" & vbCrLf
        cssStyle = cssStyle & "margin-top:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;}" & vbCrLf
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li:hover ul {display: block; position: absolute; top:0px;left:0;} " & vbCrLf
    
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li ul li {"
    '�˵�С��Li�����߿�
    If SidLiBorderTop <> "" Then cssStyle = cssStyle & "border-top:" & SidLiBorderTop & ";" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li:hover ul li a {display:block; width:110px; text-align:center;}" & vbCrLf
    cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li:hover ul li a:hover {}" & vbCrLf
    
    
    
    cssStyle = cssStyle & "/*End*/" & vbCrLf

    
    DivHtml = DivHtml & "<div id=""nav" & id & cssNameAddId & """>" & vbCrLf
    DivHtml = DivHtml & "    <div class=""navbox"">    " & vbCrLf
    DivHtml = DivHtml & "        <ul class=""nav" & id & cssNameAddId & """>" & vbCrLf
    DivHtml = DivHtml & "{$Nav$}" & vbCrLf
    DivHtml = DivHtml & "        </ul>" & vbCrLf
    DivHtml = DivHtml & "        <div class=""clear""></div> " & vbCrLf
    DivHtml = DivHtml & "    </div>" & vbCrLf
    DivHtml = DivHtml & "</div>" & vbCrLf
 
End Sub


'���ָ��Left��������ʽ 20150626
Function getLeftIDCssStyleContent(id, parameter)
    Dim content, splStr, FFolder, nOK, s
    FFolder = DefaultTemplateFolder & "\Left\" & id
    content = getFText(FFolder & "/Left.txt")
    splStr = Split(content, "------------------------" & vbCrLf)
    If UBound(splStr) >= parameter Then
        getLeftIDCssStyleContent = splStr(parameter)
    End If
End Function
'���Left��ʽ ParameterΪ����
Public Function GetLeftCss(id, cssStyle, DivHtml, stylevalue)
    Dim c, FFolder, fName, ImgWidth, ImgHeight, content, AHeight
    Dim startStr, endStr, cssStr, CustomCssStyle
    Dim TLeftYes, TIconYes, TLineYes, TMoreYes, TRightYes, CLeftYes, CRightYes, FLeftYes, FLineYes, FRightYes
    Dim splStr, cssNameAddId
    TLeftYes = False         '�������ͼƬΪ��
    TIconYes = False         '����ICOͼƬΪ��
    TLineYes = False         '������ͼƬΪ��
    TMoreYes = False         '��������ͼƬΪ��
    TRightYes = False        '�����ұ�ͼƬΪ��
    CLeftYes = False         '�м����ͼƬΪ��
    CRightYes = False         '�м��ұ�ͼƬΪ��
    FLeftYes = False         '�ײ����ͼƬΪ��
    FLineYes = False         '�ײ���ͼƬΪ��
    FRightYes = False         '�ײ��ұ�ͼƬΪ��
    

    
    
    AHeight = 0                'Ĭ�����ӿ���߳�ֵ
    cssStr = ""         'Css����
    FFolder = DefaultTemplateFolder & "\Left\" & id
    
    
    
    'Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    If checkStrIsNumberType(stylevalue) Then
        cssNameAddId = "_" & stylevalue                      'Css����׷��Id���
        stylevalue = getLeftIDCssStyleContent(id, CInt(stylevalue))
    End If
     
    '��Ŀ����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tbg{" & vbCrLf
    startStr = ".tbg{": endStr = "}"
    If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
        cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
    End If
    
    

    
    fName = findDirFileName(FFolder, "L_Bg")
    If fName <> "" And InStr(cssStr, "background-image:") = False Then      '�������� ����û���Զ��屳��ͼƬ
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
    End If
    'Left��ʽ��
    fName = findSubDirFileName(FFolder, "L_Bg|L_Left|L_Right|L_Line")
    If fName <> "" And InStr(cssStr, "height:") = False Then      '��ʽͼƬ���� ����û���Զ����
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStr = cssStr & "    height:" & ImgHeight & "px;" & vbCrLf
        cssStr = cssStr & "    line-height:" & ImgHeight & "px;" & vbCrLf
    End If
    
    cssStyle = cssStyle & cssStr & "}" & vbCrLf
    '��Ŀ�����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tleft{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Left")
    If fName <> "" Then
        TLeftYes = True         '�������ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:left;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀͼ��
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .ticon{" & vbCrLf
    cssStr = ""         'Css����
    startStr = ".ticon{": endStr = "}"
    If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
        cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
    End If
     
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Icon")
    If fName <> "" And InStr(cssStr, "background-image:") = False Then        '�������� ����û���Զ��屳��ͼƬ
        TIconYes = True         '����ICOͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStr = cssStr & "    background-repeat: no-repeat;" & vbCrLf
        If InStr(cssStr, "background-position:") = False Then cssStr = cssStr & "    background-position: -0px;" & vbCrLf
        If InStr(cssStr, "width:") = False Then cssStr = cssStr & "    width:" & ImgWidth & "px;" & vbCrLf
        If InStr(cssStr, "height:") = False Then cssStr = cssStr & "    height:" & AHeight & "px;" & vbCrLf
    End If
    cssStr = cssStr & "    float:left;" & vbCrLf
    cssStyle = cssStyle & cssStr & "}" & vbCrLf
    '��Ŀ����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tvalue{" & vbCrLf
    
    cssStr = ""         'Css����
    startStr = ".tvalue{": endStr = "}"
    If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
        cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
    End If
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Value")
    If fName <> "" And InStr(cssStr, "background-image:") = False Then       '�������� ����û���Զ��屳��ͼƬ
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
    End If
    If InStr(cssStr, "width:") = False Then cssStr = cssStr & "    width:110px;" & vbCrLf
    If InStr(cssStr, "height:") = False Then cssStr = cssStr & "    height:" & minMaxBetween(24, 120, AHeight) & "px;" & vbCrLf
    If InStr(cssStr, "line-height:") = False Then cssStr = cssStr & "    line-height:" & minMaxBetween(24, 120, AHeight) & "px;" & vbCrLf
    If InStr(cssStr, "padding-left:") = False Then cssStr = cssStr & "    padding-left:10px;" & vbCrLf
    
    
    cssStr = cssStr & "    float:left;" & vbCrLf
    cssStyle = cssStyle & cssStr & "}" & vbCrLf
    '��Ŀ�ָ���
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tline{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Line")
    If fName <> "" Then
        TLineYes = True         '������ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:left;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tmore{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_More")
    If fName <> "" Then
        TMoreYes = True         '��������ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:right;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�ұ�
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .tright{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Right")
    If fName <> "" Then
        TRightYes = True         '�����ұ�ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:right;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�м䱳��
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .cbg{" & vbCrLf
    
    cssStr = ""         'Css����
    startStr = ".cbg{": endStr = "}"
    If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
        cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
    End If
    
    
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_CBg")
    If fName <> "" And InStr(cssStr, "background-image:") = False Then       '�������� ����û���Զ��屳��ͼƬ
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
    End If

    cssStyle = cssStyle & cssStr & "}" & vbCrLf
    '��Ŀ�����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .cleft{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_CLeft")
    If fName <> "" Then
        CLeftYes = True         '�м����ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    height:200px;" & vbCrLf
    cssStyle = cssStyle & "    float:left;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�м�����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .ccontent{" & vbCrLf
    cssStr = ""         'Css����
    startStr = ".ccontent{": endStr = "}"
    If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
        cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
    End If
    
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_Ccontent")
    If fName <> "" And InStr(cssStr, "background-image:") = False Then      '�������� ����û���Զ��屳��ͼƬ
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
    End If
    
    If InStr(cssStr, "font-size:") = False Then cssStr = cssStr & "    font-size:12px;" & vbCrLf
    If InStr(cssStr, "line-height:") = False Then cssStr = cssStr & "    line-height:24px;" & vbCrLf
    If InStr(cssStr, "margin") = False Then cssStr = cssStr & "    margin:8px;" & vbCrLf
    'If InStr(CssStr, "width:") = False Then CssStr = CssStr & "    width:94%;" & vbCrLf        '94% ȥ�������ƿ�20150108
    cssStr = cssStr & "    float:left;" & vbCrLf
    
    
    cssStyle = cssStyle & cssStr & "}" & vbCrLf
    '��Ŀ���ұ�
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .cright{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_CRight")
    If fName <> "" Then
        CRightYes = True         '�м��ұ�ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    height:200px;" & vbCrLf
    cssStyle = cssStyle & "    float:right;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�ײ�����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .fbg{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_FBg")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
        cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�����
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .fleft{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_FLeft")
    If fName <> "" Then
        FLeftYes = True         '�ײ����ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:left;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ�ָ���
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .fline{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_FLine")
    If fName <> "" Then
        FLineYes = True         '�ײ���ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
        cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:left;" & vbCrLf
    cssStyle = cssStyle & "}" & vbCrLf
    '��Ŀ���ұ�
    cssStyle = cssStyle & ".columnlist" & id & cssNameAddId & " .fright{" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Left\" & id
    fName = findDirFileName(FFolder, "L_FRight")
    If fName <> "" Then
        FRightYes = True         '�ײ��ұ�ͼƬΪ��
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        If AHeight < ImgHeight Then AHeight = ImgHeight         '�����ı���
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "    float:right;" & vbCrLf               '��ո���
    cssStyle = cssStyle & "}" & vbCrLf
    
    '��Ŀ��ʽ
    c = "/*��Ŀ�б�*/" & vbCrLf
    c = c & ".columnlist" & id & cssNameAddId & "{" & vbCrLf
    c = c & "    width:auto;" & vbCrLf
    c = c & "    height:auto;" & vbCrLf
    c = c & "}" & vbCrLf
    '����������ʽ
    c = c & ".columnlist" & id & cssNameAddId & " a{" & vbCrLf
    c = c & "    height:" & AHeight & "px;" & vbCrLf
    c = c & "    line-height:" & AHeight & "px;" & vbCrLf
    c = c & "}" & vbCrLf
    cssStyle = c & cssStyle & vbCrLf
    
    DivHtml = DivHtml & "    <div class=""columnlist" & id & cssNameAddId & """>            " & vbCrLf
    DivHtml = DivHtml & "        <div class=""tbg"">" & vbCrLf
    If TLeftYes = True Then DivHtml = DivHtml & "            <div class=""tleft""></div>    " & vbCrLf
    If TIconYes = True Then DivHtml = DivHtml & "            <div class=""ticon""></div>    " & vbCrLf
    DivHtml = DivHtml & "            <div class=""tvalue"">��Ŀ����</div>    " & vbCrLf
    If TLineYes = True Then DivHtml = DivHtml & "            <div class=""tline""></div>" & vbCrLf
    If TRightYes = True Then DivHtml = DivHtml & "            <div class=""tright""></div>   " & vbCrLf
    If TMoreYes = True Then DivHtml = DivHtml & "            <div class=""tmore""></div>" & vbCrLf
    
    DivHtml = DivHtml & "<!--#AMore#-->" & vbCrLf                    'A���Ӵ�����
    DivHtml = DivHtml & "            <div class=""clear""></div> " & vbCrLf                '�������������ҳ���֪��Ϊʲô��20141124��  ����û�������������Ǹ�DIV����һ����
    DivHtml = DivHtml & "        </div>" & vbCrLf
    DivHtml = DivHtml & "        " & vbCrLf
    
    DivHtml = DivHtml & "        <div class=""cbg"">" & vbCrLf
    '�м����ͼƬ
    If CLeftYes = True Then DivHtml = DivHtml & "            <div class=""cleft""></div>" & vbCrLf
    DivHtml = DivHtml & "            <div class=""ccontent"">��Ŀ����</div>" & vbCrLf
    '�м��ұ�ͼƬ
    If CRightYes = True Then DivHtml = DivHtml & "            <div class=""cright""></div>" & vbCrLf
    DivHtml = DivHtml & "            <div class=""clear""></div> " & vbCrLf             '��ո���
    DivHtml = DivHtml & "        </div>" & vbCrLf
    DivHtml = DivHtml & "        " & vbCrLf
    
    If FLeftYes = True Or FLineYes = True Or FRightYes = True Then
        DivHtml = DivHtml & "        <div class=""fbg"">" & vbCrLf
        If FLeftYes = True Then DivHtml = DivHtml & "            <div class=""fleft""></div>            " & vbCrLf
        If FLineYes = True Then DivHtml = DivHtml & "            <div class=""fline""></div>            " & vbCrLf
        If FRightYes = True Then DivHtml = DivHtml & "            <div class=""fright""></div>" & vbCrLf
        DivHtml = DivHtml & "            <div class=""clear""></div> " & vbCrLf
        DivHtml = DivHtml & "        </div>" & vbCrLf
    End If
     

    
    DivHtml = DivHtml & "    </div>" & vbCrLf
    GetLeftCss = DivHtml
End Function
'���Foot��ʽ
Sub GetFootCss(id, cssStyle, DivHtml)
    Dim c, path, ImgWidth, ImgHeight, FFolder, fName
    'Css����
    cssStyle = cssStyle & ".foot" & id & "{" & vbCrLf
    cssStyle = cssStyle & "    width:auto;" & vbCrLf
    FFolder = DefaultTemplateFolder & "\Foot\" & id & "\Index"
    fName = findDirFileName(FFolder, "FootBg")
    If fName <> "" Then
        Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
        cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Foot/" & id & "/Index/" & fName & ");" & vbCrLf
        cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
        cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
    End If
    cssStyle = cssStyle & "}" & vbCrLf
    'Html����
    DivHtml = DivHtml & "    <div class=foot" & id & " id=foot>{$WebBottom$}" & vbCrLf
    DivHtml = DivHtml & "    </div>" & vbCrLf
End Sub

end class 


%>