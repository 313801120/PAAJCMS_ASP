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
'处理html到模板
'dim t:set t=new class_vbTemplate 
class class_vbTemplate 
	dim DefaultTemplateFolder			'默认模板文件夹路径
	dim DefaultTemplateRootHttpUrl		'主网址
	dim DefaultTemplateHttpUrl
	 '构造函数 初始化
    Sub Class_Initialize() 
		DefaultTemplateFolder="\DataDir\VB模块\服务器\Template\"
		DefaultTemplateRootHttpUrl="http://sharembweb.com/img.asp?img="
		DefaultTemplateHttpUrl=DefaultTemplateRootHttpUrl' & "/DataDir/VB模块/服务器/Template/"
		MDBPath=DefaultTemplateFolder & "Database\WebStyle.mdb"
		accessPass=""
		call openconn()
	end sub
    '析构函数 类终止
    Sub Class_Terminate() 
    End Sub 
	
	
	function showBigClassList(did,sid)
		dim c,sel
		c="<form name='form1' method='get'><input type='hidden' name='act' value='handleVBTemplate' />"
		c=c&"大类<select name='BigClassName'>"
		c=c&"<option value="""">选择大类</option>"&vbcrlf
	    rs.Open "Select * From [BigClass] Order By Sort Asc", conn, 1, 1
		While Not rs.EOF
			sel=IIF(did=rs("bigclassname")," selected","")
			c=c&"<option value="""&rs("BigClassName")&""""&sel&">"&rs("BigClassName")&"</option>"&vbcrlf
		rs.MoveNext: Wend: rs.Close 
		c=c&"</select>"
		c=c&"小类<select name='SmallClassName'>"
		c=c&"<option value="""">选择小类</option>"&vbcrlf
		rs.Open "Select * From [SmallClass] Where BigClassName='" & did & "' Order By Sort Asc", conn, 1, 1
		While Not rs.EOF
			sel=IIF(sid=rs("SmallClassName")," selected","")
			c=c&"<option value="""&rs("SmallClassName")&""""&sel&">"&rs("SmallClassName")&"</option>"&vbcrlf
		rs.MoveNext: Wend: rs.Close
		c=c&"</select>"
		
		c=c&"<input type='submit' name='button' value='提交' /></form>"
		showBigClassList=c
	end function
	
	'显示处理后导航内容
	Sub ShowHandleWebPage(cssStyle, DivHtml)
		Dim filePath, content, c
		c = "<!DOCTYPE html>" & vbCrLf & "<head>" & vbCrLf
		c = c & "<style>" & vbCrLf
		c = c & "body, div, p,img,dl, dt, dd, ul, ol, li, h1, h2, h3, h4, h5, h6, pre, form, fieldset, input, textarea, blockquote { padding:0px; margin:0px; }" & vbCrLf
		c = c & "li { list-style-type:none; }" & vbCrLf
		c = c & ".navbox{width:980px;  margin:0 auto 0 auto}" & vbCrLf
		c = c & ".clear { clear:both; }" & vbCrLf
		c = c & cssStyle & vbCrLf & "</style>" & vbCrLf
		c = c & "<body>" & vbCrLf
		'If BigClassName.Text <> "Nav" Then C = C & "<div style=""width:980px;height:auto;margin:30px auto 0 auto"">" & vbCrLf
		c = c & DivHtml & vbCrLf
		c = Replace(c, "{$Nav$}", getFText("VB工程\Template\ShowDefaultNav.html"))         '替换显示默认导航内容
		'定义宽度
		'If BigClassName.Text <> "Nav" Then C = C & "</div>"
		
		c = c & vbCrLf & "</body></html>"
		c = Replace(c, vbTab, "    ")
		response.Write(c)
	End Sub
	'============================== VB引用 ================================= 
	 
	 '获得Top样式
	Sub GetTopCss(id, cssStyle, DivHtml)
		Dim c, path, ImgWidth, ImgHeight, FFolder, fName
		'Css部分
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
		'Html部分
		DivHtml = DivHtml & "    <div id=header" & id & ">" & vbCrLf
		DivHtml = DivHtml & "    </div>" & vbCrLf
	End Sub
	'显示默认导航
	Function NavDefaultValue()
		Dim splStr, i, s, c
		c = c & "<li><a href=javascript:;>首页</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>公司简介</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>产品展示</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>新闻中心</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>在线留言</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>公司招聘</a></li><li class=line></li>" & vbCrLf
		c = c & "<li><a href=javascript:;>联系我们</a></li><li class=line></li>" & vbCrLf
		NavDefaultValue = c
	End Function
	'获得指定导航多少种样式 20150626
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
	'获得指定导航多少种样式 20150626
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
	
	'获得Nav样式 ID为样式编码 Parameter为参数
	Public Sub GetNavCss(id, cssStyle, DivHtml, ByVal stylevalue)
		Dim c, FFolder, fName, ImgWidth, ImgHeight, content, AWidth, AHeight, NavBackground, NavListBgColor, NavFontHeight, AMargin
		Dim splStr, DefaultColor, MoveColor, FontSize, FontFamily, FontWeight, SidLiBorderTop, tempS
		Dim Apadding, NavValueRepeat, CssLeftPadding
		Dim CopyOK                          '复制文件是否
		Dim NavClassHeightYes               '导航样式高是否为真，为假则调用自定义高度   20150317
		Dim cssNameAddId                    'Css名称追加Id编号
		
		DefaultColor = "#000000"            '默认链接颜色
		MoveColor = "#FFFFFF"               '移动链接颜色
		FontSize = 12                       '默认文字大小
		AWidth = 0                          '默认链接宽
		AHeight = 0                         '默认链接高
		NavBackground = ""                  '导航背景
		NavListBgColor = ""                 '导航列表背景
		NavFontHeight = ""                  '导航字体高度
		AMargin = ""                        '导航LI值
		NavValueRepeat = ""                 '导航当前选中内容背景是否重复 20150213
		CssLeftPadding = ""                   'CssLeft Padding样式设置
		 
		
		FFolder = DefaultTemplateFolder & "\Nav\" & id
		
		'为数字类型 则自动提取样式内容  20150615
		If checkStrIsNumberType(stylevalue) Then
			cssNameAddId = "_" & stylevalue                      'Css名称追加Id编号
			stylevalue = getNavIDCssStyleContent(id, CInt(stylevalue))
		End If
	
		
		'默认颜色
		tempS = rParam(stylevalue, "DefaultColor")
		If tempS <> "" Then DefaultColor = tempS
		'放上颜色
		tempS = rParam(stylevalue, "MoveColor")
		If tempS <> "" Then MoveColor = tempS
		'文字大小
		tempS = rParam(stylevalue, "FontSize")
		If tempS <> "" Then FontSize = tempS
		'文字字体
		tempS = rParam(stylevalue, "font-family")
		If tempS <> "" Then FontFamily = tempS
		'文字正常与加粗
		tempS = rParam(stylevalue, "FontWeight")
		If tempS <> "" Then FontWeight = tempS
		'链接宽度
		tempS = rParam(stylevalue, "AWidth")
		If tempS <> "" Then AWidth = tempS
		'导航背景
		tempS = rParam(stylevalue, "NavBackground")
		If tempS <> "" Then NavBackground = tempS
		'导航列表颜色
		tempS = rParam(stylevalue, "NavListBgColor")
		If tempS <> "" Then NavListBgColor = tempS
		'导航文本高度
		tempS = rParam(stylevalue, "NavFontHeight")
		If tempS <> "" Then NavFontHeight = tempS
		'导航链接A边框值
		tempS = rParam(stylevalue, "AMargin")
		If tempS <> "" Then AMargin = tempS
		'导航链接A边框值
		tempS = rParam(stylevalue, "Apadding")
		If tempS <> "" Then Apadding = tempS
		'菜单小类Li顶部边框
		tempS = rParam(stylevalue, "SidLiBorderTop")
		If tempS <> "" Then SidLiBorderTop = tempS
		'导航选中背景底部居中显示
		tempS = rParam(stylevalue, "navvaluerepeat")
		If tempS <> "" Then NavValueRepeat = tempS
		'CssLeft padding设置
		tempS = rParam(stylevalue, "cssleftpadding")
		If tempS <> "" Then CssLeftPadding = tempS
		
		 
		'旧版不好用，摈弃
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
		
		'网站导航区
		cssStyle = cssStyle & "/*导航*/" & vbCrLf
		
		'导航框架
		cssStyle = cssStyle & "#nav" & id & cssNameAddId & "{" & vbCrLf
		
		
		fName = findSubDirFileName(FFolder, "NavLeft|NavRight")
		'左边与右边没有样式则
		If fName = "" Then
			If NavBackground <> "" Then cssStyle = cssStyle & "background-color:" & NavBackground & ";" & vbCrLf
			fName = findDirFileName(FFolder, "NavBg")
			If fName <> "" Then
				cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
				cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf     'XY都重复 所以多行导航也无所谓呀
			End If
		End If
		
	
		
		cssStyle = cssStyle & "}" & vbCrLf
		
		'导航背景
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & "{" & vbCrLf
		cssStyle = cssStyle & "    width:100%;" & vbCrLf
		'导航背景
		fName = findDirFileName(FFolder, "NavBg")
		NavClassHeightYes = False           '导航类型高默认为假
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf     'XY都重复 所以多行导航也无所谓呀
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
			cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
			NavClassHeightYes = True
		Else
			'没有背景自动寻找导航高
			fName = findSubDirFileName(FFolder, "NavMove|NavValue}NavLine")
			If fName <> "" Then
				Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
				If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
				cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
				cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
				NavClassHeightYes = True
			End If
			'导航有背景
			If NavBackground <> "" Then
				cssStyle = cssStyle & "    background-color: " & NavBackground & ";" & vbCrLf
			End If
		End If
		'暂时20150317
		If NavFontHeight <> "" And NavClassHeightYes = False Then
			cssStyle = cssStyle & "    height:" & NavFontHeight & "px;" & vbCrLf
			cssStyle = cssStyle & "    line-height:" & NavFontHeight & "px;" & vbCrLf
		End If
		
		cssStyle = cssStyle & "    font-weight:" & FontWeight & ";" & vbCrLf                '文字正常还是加粗
		cssStyle = cssStyle & "}" & vbCrLf
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " li{float:left;position:relative;}" & vbCrLf   '加了个相对定位20141111
		'导航左边
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .left{" & vbCrLf
	 
		fName = findDirFileName(FFolder, "NavLeft")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		'导航left隐藏掉
		Else
			cssStyle = cssStyle & "    width:0px;" & vbCrLf
			cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
			
			If CssLeftPadding <> "" Then
				cssStyle = cssStyle & "    padding:" & CssLeftPadding & vbCrLf
			End If
		End If
		cssStyle = cssStyle & "}" & vbCrLf
		'导航分割线
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .line{" & vbCrLf
		
		fName = findDirFileName(FFolder, "NavLine")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		'导航left隐藏掉
		Else
			cssStyle = cssStyle & "    width:0px;" & vbCrLf
			cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
		End If
		cssStyle = cssStyle & "}" & vbCrLf
	
		'导航右边
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " .right{" & vbCrLf
		
		fName = findDirFileName(FFolder, "NavRight")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		'导航left隐藏掉
		Else
			cssStyle = cssStyle & "    width:0px;" & vbCrLf
			cssStyle = cssStyle & "    overflow: hidden;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:right;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'导航链接
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " a{" & vbCrLf
		
		fName = findSubDirFileName(FFolder, "NavValue")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			If AWidth = 0 Then AWidth = ImgWidth                                     '导航链接宽等于
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			If FontFamily <> "" Then cssStyle = cssStyle & "    font-family:" & FontFamily & ";" & vbCrLf
		End If
		'获得导航项宽
		fName = findSubDirFileName(FFolder, "NavValue|NavMove")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			If AWidth = 0 Then AWidth = ImgWidth                                        '导航链接宽等于
		End If
		
		cssStyle = cssStyle & "    font-size:" & FontSize & "px;" & vbCrLf
		cssStyle = cssStyle & "    color:" & DefaultColor & ";" & vbCrLf
		cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
		If AWidth <> "" And CStr(AWidth) <> "0" And rParam(stylevalue, "AWidth") <> "" Then
			cssStyle = cssStyle & "    width:" & minMaxBetween(90, 220, AWidth) & "px;" & vbCrLf      '一开始判断错了，我是个猪呀
		End If
		cssStyle = cssStyle & "    height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
		'导航链接A边框值
		If AMargin <> "" Then
			cssStyle = cssStyle & "    margin:" & AMargin & ";" & vbCrLf
		End If
		'导航链接A填充值
		If Apadding <> "" Then
			cssStyle = cssStyle & "    padding:" & Apadding & ";" & vbCrLf
		End If
		'大于0
		If AHeight > 0 Or NavFontHeight <> "" Then
			cssStyle = cssStyle & "    line-height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    display:block;" & vbCrLf
		cssStyle = cssStyle & "    text-align:center;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'导航链接移上
		cssStyle = cssStyle & ".nav" & id & cssNameAddId & " a:hover{" & vbCrLf
		
		fName = findDirFileName(FFolder, "NavMove")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			'20150213
			If NavValueRepeat = "false" Then
				cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
				cssStyle = cssStyle & "    background-position: center bottom;" & vbCrLf
			'图片宽小于20则平铺
			ElseIf ImgWidth < 20 Then
				cssStyle = cssStyle & "    background-repeat: repeat-x;" & vbCrLf
			Else
				cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			End If
			If FontFamily <> "" Then cssStyle = cssStyle & "    font-family:" & FontFamily & ";" & vbCrLf
		'导航项目列表有背景
		ElseIf NavListBgColor <> "" Then
			cssStyle = cssStyle & "    background-color: " & NavListBgColor & ";" & vbCrLf
		End If
		cssStyle = cssStyle & "    text-decoration:none;" & vbCrLf
		cssStyle = cssStyle & "    color:" & MoveColor & ";" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'点中当前按钮 添加于20140905
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
		'导航链接A边框值
		If AMargin <> "" Then
			cssStyle = cssStyle & "    margin:" & AMargin & ";" & vbCrLf
		End If
		'导航链接A填充值
		If Apadding <> "" Then
			cssStyle = cssStyle & "    padding:" & Apadding & ";" & vbCrLf
		End If
		'大于0
		If AHeight > 0 Or NavFontHeight <> "" Then
			cssStyle = cssStyle & "    line-height:" & IIf(NavFontHeight <> "", NavFontHeight, AHeight) & "px;" & vbCrLf
		End If
		 
		
		fName = findSubDirFileName(FFolder, "NavMove|NavValue")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Nav/" & id & "/" & fName & ");" & vbCrLf
			
			'20150213
			If NavValueRepeat = "false" Then
				cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
				cssStyle = cssStyle & "    background-position: center bottom;" & vbCrLf
			'图片宽小于20则平铺
			ElseIf ImgWidth < 20 Then
				cssStyle = cssStyle & "    background-repeat: repeat-x;" & vbCrLf
			Else
				cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			End If
		'导航项目列表有背景
		ElseIf NavListBgColor <> "" Then
			cssStyle = cssStyle & "    background-color: " & NavListBgColor & ";" & vbCrLf
		End If
		cssStyle = cssStyle & "}" & vbCrLf
		'点中当前按钮 里链接样式 20141217
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
		'菜单小类Li顶部边框
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
	
	
	'获得指定Left多少种样式 20150626
	Function getLeftIDCssStyleContent(id, parameter)
		Dim content, splStr, FFolder, nOK, s
		FFolder = DefaultTemplateFolder & "\Left\" & id
		content = getFText(FFolder & "/Left.txt")
		splStr = Split(content, "------------------------" & vbCrLf)
		If UBound(splStr) >= parameter Then
			getLeftIDCssStyleContent = splStr(parameter)
		End If
	End Function
	'获得Left样式 Parameter为参数
	Public Function GetLeftCss(id, cssStyle, DivHtml, stylevalue)
		Dim c, FFolder, fName, ImgWidth, ImgHeight, content, AHeight
		Dim startStr, endStr, cssStr, CustomCssStyle
		Dim TLeftYes, TIconYes, TLineYes, TMoreYes, TRightYes, CLeftYes, CRightYes, FLeftYes, FLineYes, FRightYes
		Dim splStr, cssNameAddId
		TLeftYes = False         '顶部左边图片为假
		TIconYes = False         '顶部ICO图片为假
		TLineYes = False         '顶部线图片为假
		TMoreYes = False         '顶部更多图片为假
		TRightYes = False        '顶部右边图片为假
		CLeftYes = False         '中间左边图片为假
		CRightYes = False         '中间右边图片为假
		FLeftYes = False         '底部左边图片为假
		FLineYes = False         '底部线图片为假
		FRightYes = False         '底部右边图片为假
		
	
		
		
		AHeight = 0                '默认链接宽与高初值
		cssStr = ""         'Css内容
		FFolder = DefaultTemplateFolder & "\Left\" & id
		
		
		
		'为数字类型 则自动提取样式内容  20150615
		If checkStrIsNumberType(stylevalue) Then
			cssNameAddId = "_" & stylevalue                      'Css名称追加Id编号
			stylevalue = getLeftIDCssStyleContent(id, CInt(stylevalue))
		End If
		 
		'栏目背景
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tbg{" & vbCrLf
		startStr = ".tbg{": endStr = "}"
		If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
			cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
		End If
		
		
	
		
		fName = findDirFileName(FFolder, "L_Bg")
		If fName <> "" And InStr(cssStr, "background-image:") = False Then      '背景存在 并且没有自定义背景图片
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
		End If
		'Left样式高
		fName = findSubDirFileName(FFolder, "L_Bg|L_Left|L_Right|L_Line")
		If fName <> "" And InStr(cssStr, "height:") = False Then      '样式图片存在 并且没有自定义高
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStr = cssStr & "    height:" & ImgHeight & "px;" & vbCrLf
			cssStr = cssStr & "    line-height:" & ImgHeight & "px;" & vbCrLf
		End If
		
		cssStyle = cssStyle & cssStr & "}" & vbCrLf
		'栏目上左边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tleft{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Left")
		If fName <> "" Then
			TLeftYes = True         '顶部左边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:left;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目图标
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .ticon{" & vbCrLf
		cssStr = ""         'Css内容
		startStr = ".ticon{": endStr = "}"
		If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
			cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
		End If
		 
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Icon")
		If fName <> "" And InStr(cssStr, "background-image:") = False Then        '背景存在 并且没有自定义背景图片
			TIconYes = True         '顶部ICO图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStr = cssStr & "    background-repeat: no-repeat;" & vbCrLf
			If InStr(cssStr, "background-position:") = False Then cssStr = cssStr & "    background-position: -0px;" & vbCrLf
			If InStr(cssStr, "width:") = False Then cssStr = cssStr & "    width:" & ImgWidth & "px;" & vbCrLf
			If InStr(cssStr, "height:") = False Then cssStr = cssStr & "    height:" & AHeight & "px;" & vbCrLf
		End If
		cssStr = cssStr & "    float:left;" & vbCrLf
		cssStyle = cssStyle & cssStr & "}" & vbCrLf
		'栏目内容
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tvalue{" & vbCrLf
		
		cssStr = ""         'Css内容
		startStr = ".tvalue{": endStr = "}"
		If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
			cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
		End If
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Value")
		If fName <> "" And InStr(cssStr, "background-image:") = False Then       '背景存在 并且没有自定义背景图片
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
		End If
		If InStr(cssStr, "width:") = False Then cssStr = cssStr & "    width:110px;" & vbCrLf
		If InStr(cssStr, "height:") = False Then cssStr = cssStr & "    height:" & minMaxBetween(24, 120, AHeight) & "px;" & vbCrLf
		If InStr(cssStr, "line-height:") = False Then cssStr = cssStr & "    line-height:" & minMaxBetween(24, 120, AHeight) & "px;" & vbCrLf
		If InStr(cssStr, "padding-left:") = False Then cssStr = cssStr & "    padding-left:10px;" & vbCrLf
		
		
		cssStr = cssStr & "    float:left;" & vbCrLf
		cssStyle = cssStyle & cssStr & "}" & vbCrLf
		'栏目分割线
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tline{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Line")
		If fName <> "" Then
			TLineYes = True         '顶部线图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:left;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目更多
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tmore{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_More")
		If fName <> "" Then
			TMoreYes = True         '顶部更多图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:right;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目右边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .tright{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Right")
		If fName <> "" Then
			TRightYes = True         '顶部右边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:right;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目中间背景
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .cbg{" & vbCrLf
		
		cssStr = ""         'Css内容
		startStr = ".cbg{": endStr = "}"
		If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
			cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
		End If
		
		
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_CBg")
		If fName <> "" And InStr(cssStr, "background-image:") = False Then       '背景存在 并且没有自定义背景图片
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStr = cssStr & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStr = cssStr & "    background-repeat: repeat;" & vbCrLf
		End If
	
		cssStyle = cssStyle & cssStr & "}" & vbCrLf
		'栏目中左边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .cleft{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_CLeft")
		If fName <> "" Then
			CLeftYes = True         '中间左边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    height:200px;" & vbCrLf
		cssStyle = cssStyle & "    float:left;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目中间内容
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .ccontent{" & vbCrLf
		cssStr = ""         'Css内容
		startStr = ".ccontent{": endStr = "}"
		If InStr(stylevalue, startStr) > 0 And InStr(stylevalue, endStr) > 0 Then
			cssStr = cssStr & strCut(stylevalue, startStr, endStr, 2)
		End If
		
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_Ccontent")
		If fName <> "" And InStr(cssStr, "background-image:") = False Then      '背景存在 并且没有自定义背景图片
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
		End If
		
		If InStr(cssStr, "font-size:") = False Then cssStr = cssStr & "    font-size:12px;" & vbCrLf
		If InStr(cssStr, "line-height:") = False Then cssStr = cssStr & "    line-height:24px;" & vbCrLf
		If InStr(cssStr, "margin") = False Then cssStr = cssStr & "    margin:8px;" & vbCrLf
		'If InStr(CssStr, "width:") = False Then CssStr = CssStr & "    width:94%;" & vbCrLf        '94% 去掉，限制宽20150108
		cssStr = cssStr & "    float:left;" & vbCrLf
		
		
		cssStyle = cssStyle & cssStr & "}" & vbCrLf
		'栏目中右边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .cright{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_CRight")
		If fName <> "" Then
			CRightYes = True         '中间右边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    height:200px;" & vbCrLf
		cssStyle = cssStyle & "    float:right;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目底部背景
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .fbg{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_FBg")
		If fName <> "" Then
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
			cssStyle = cssStyle & "    line-height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目底左边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .fleft{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_FLeft")
		If fName <> "" Then
			FLeftYes = True         '底部左边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:left;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目分割线
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .fline{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_FLine")
		If fName <> "" Then
			FLineYes = True         '底部线图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: no-repeat;" & vbCrLf
			cssStyle = cssStyle & "    background-position: -0px;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:left;" & vbCrLf
		cssStyle = cssStyle & "}" & vbCrLf
		'栏目中右边
		cssStyle = cssStyle & ".c" & id & cssNameAddId & " .fright{" & vbCrLf
		FFolder = DefaultTemplateFolder & "\Left\" & id
		fName = findDirFileName(FFolder, "L_FRight")
		If fName <> "" Then
			FRightYes = True         '底部右边图片为真
			Call getImageWidthHeight(FFolder & "\" & fName, ImgWidth, ImgHeight)
			If AHeight < ImgHeight Then AHeight = ImgHeight         '链接文本高
			cssStyle = cssStyle & "    background-image: url(" & DefaultTemplateHttpUrl & "/Left/" & id & "/" & fName & ");" & vbCrLf
			cssStyle = cssStyle & "    background-repeat: repeat;" & vbCrLf
			cssStyle = cssStyle & "    width:" & ImgWidth & "px;" & vbCrLf
			cssStyle = cssStyle & "    height:" & ImgHeight & "px;" & vbCrLf
		End If
		cssStyle = cssStyle & "    float:right;" & vbCrLf               '清空浮动
		cssStyle = cssStyle & "}" & vbCrLf
		
		'栏目样式
		c = "/*栏目列表*/" & vbCrLf
		c = c & ".c" & id & cssNameAddId & "{" & vbCrLf
		c = c & "    width:auto;" & vbCrLf
		c = c & "    height:auto;" & vbCrLf
		c = c & "}" & vbCrLf
		'标上链接样式
		c = c & ".c" & id & cssNameAddId & " a{" & vbCrLf
		c = c & "    height:" & AHeight & "px;" & vbCrLf
		c = c & "    line-height:" & AHeight & "px;" & vbCrLf
		c = c & "}" & vbCrLf
		cssStyle = c & cssStyle & vbCrLf
		
		DivHtml = DivHtml & "    <div class=""c" & id & cssNameAddId & """>            " & vbCrLf
		DivHtml = DivHtml & "        <div class=""tbg"">" & vbCrLf
		If TLeftYes = True Then DivHtml = DivHtml & "            <div class=""tleft""></div>    " & vbCrLf
		If TIconYes = True Then DivHtml = DivHtml & "            <div class=""ticon""></div>    " & vbCrLf
		DivHtml = DivHtml & "            <div class=""tvalue"">栏目标题</div>    " & vbCrLf
		If TLineYes = True Then DivHtml = DivHtml & "            <div class=""tline""></div>" & vbCrLf
		If TRightYes = True Then DivHtml = DivHtml & "            <div class=""tright""></div>   " & vbCrLf
		If TMoreYes = True Then DivHtml = DivHtml & "            <div class=""tmore""></div>" & vbCrLf
		
		DivHtml = DivHtml & "<!--#AMore#-->" & vbCrLf                    'A链接代码区
		DivHtml = DivHtml & "            <div class=""clear""></div> " & vbCrLf                '会出错，在内容页里，不知道为什么？20141124改  这里没错，错在上面那个DIV加了一个高
		DivHtml = DivHtml & "        </div>" & vbCrLf
		DivHtml = DivHtml & "        " & vbCrLf
		
		DivHtml = DivHtml & "        <div class=""cbg"">" & vbCrLf
		'中间左边图片
		If CLeftYes = True Then DivHtml = DivHtml & "            <div class=""cleft""></div>" & vbCrLf
		DivHtml = DivHtml & "            <div class=""ccontent"">栏目内容</div>" & vbCrLf
		'中间右边图片
		If CRightYes = True Then DivHtml = DivHtml & "            <div class=""cright""></div>" & vbCrLf
		DivHtml = DivHtml & "            <div class=""clear""></div> " & vbCrLf             '清空浮动
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
	'获得Foot样式
	Sub GetFootCss(id, cssStyle, DivHtml)
		Dim c, path, ImgWidth, ImgHeight, FFolder, fName,content
		'Css部分
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
		
		'Html部分
		DivHtml = DivHtml & "    <div class=foot" & id & " id=foot>{$WebBottom$}" & vbCrLf
		DivHtml = DivHtml & "    </div>" & vbCrLf
		
		'新版
		content=readFile("DataDir\VB模块\服务器\Template\Foot\"&id&"\foot.html","")
		cssStyle=getStrCut(content,"<style>","</style>",0)
		'cssStyle=replace(cssStyle,"http://127.0.0.1/","http://http://sharembweb.com/img.asp?img=")
		cssStyle=replace(cssStyle,"http://127.0.0.1/DataDir/VB模块/服务器/Template//",DefaultTemplateRootHttpUrl)
	End Sub

end class 


%>