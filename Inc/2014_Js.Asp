<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-04
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'Js

'Asp代码混淆处理 20160624 
function jsCodeConfusion(content)
    dim splStr, i, yesJs, yesWord, sx, s, wc, zc, s1, aspCode, upWord 
    dim upWordn, tempS, dimList 
    dim yesFunction                                                                 '函数是否为真
    dim isStartFunction                                                               '开始函数 目的是为了让function default 处理函数后面没有()   20150218
    dim wcType                                                                      '输入文本类型，如 " 或 '
    dim isAddToSYH                                                                  '是否累加双引号
    dim beforeStr, afterStr, endCode, nSYHCount 
	
    isStartFunction = false                                                           '默认开始函数为假
    'If nType="" Then  nType  = 0
    yesJs = false                                                                   '是ASP 默认为假
    yesFunction = false                                                             '是函数 默认为假
    yesWord = false                                                                 '是单词 默认为假
    nSYHCount = 0                                                                    '双引号默认为0
    splStr = split(content, vbCrLf)                                                 '分割行
    '循环分行
    for each s in splStr
        '循环每个字符
        for i = 1 to len(s)
            sx = mid(s, i, 1) 
            'Asp开始
            if sx = "<" and wc = "" then                                                 '输出文本必需为空 Wc为输出内容 如"<%" 排除 修改于20140412
                if mid(s, i + 1, 6) = "script" then
                    yesJs = true                                                                    'ASP为真
                    i = i + 1                                                                       '加1而不能加2，要不然<%function Test() 就截取不到
                    sx = mid(s, i, 1) 
                    aspCode = aspCode & "<" 
                end if 
            'ASP结束
            elseIf sx = "<" and mid(s, i + 1, 8) = "/script>" and wc = "" then           'Wc为输出内容
                yesJs = false                                                                   'ASP为假
                i = i + 1                                                                       '不能加2，只能加1，因为这里定义ASP为假，它会在下一次显示上面的 'ASP运行为假
                sx = mid(s, i, 8) 
                aspCode = aspCode & "/script>" 
            end if 
            if yesJs = true then

                beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '上一个字符
                afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '下一个字符
                endCode = mid(s, i + 1)                                                         '当前字符往后面代码 一行
                '输入文本
                if(sx = """" or sx = "'" and wcType = "") or sx = wcType or wc <> "" then
                    isAddToSYH = true 
                    '这是一种简单的方法，等完善(20150914)
                    if isAddToSYH = true and beforeStr = "\" then

                        if len(wc) >= 1 then
                            if isStrTransferred(wc) = true then                     '为转义字符为真
                                'call echo(wc,isStrTransferred(wc))
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
                    if sx = wcType and isAddToSYH = true then nSYHCount = nSYHCount + 1             '排除上一个字符为\这个转义字符(20150914)


                    '判断是否"在最后
                    if nSYHCount mod 2 = 0 and beforeStr <> "\" then
                        if mid(s, i + 1, 1) <> wcType then
                            wc = wc & sx 
                            aspCode = aspCode & wc                                  '行代码累加
                            'call echo("wc",wc)
                            nSYHCount = 0 : wc = ""                                 '清除
                            wcType = "" 
                        else
                            wc = wc & sx 
                        end if 
                    else
                        wc = wc & sx 
                    end if 

                elseIf sx = "'" then                                                            '注释则退出
                    aspCode = aspCode & mid(s, i) 
                    exit for 
                '字母
                elseIf checkABC(sx) = true or(sx = "_" and zc <> "") or zc <> "" then
                    zc = zc & sx 
                    s1 = LCase(mid(s & " ", i + 1, 1)) 
                    if inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1) = 0 and not(s1 = "_" and zc <> "") then'最简单判断
                        tempS = mid(s, i + 1) 
  
                        if inStr("|function|sub|", "|" & LCase(zc) & "|")>0 then
                            '函数开始
                            if yesFunction = false and LCase(upWord) <> "end" then
                                yesFunction = true 
                                dimList = getFunDimName(tempS) 
                                isStartFunction = true 
                            elseIf yesFunction = true and LCase(upWord) = "end" then'获得上一个单词
                                yesFunction = false 
                            end if 
                        elseIf yesFunction = true and LCase(zc) = "var" then
                            dimList = dimList & "," & getVarName(tempS) 
                        elseIf yesFunction = true then
                            '排除函数后面每一个名称
                            if isStartFunction = false then
                                zc = replaceDim2(dimList, zc) 
                            end if 
                            isStartFunction = false 
                        end if 
                        upWord = zc                                                                     '记住当前单词
                        aspCode = aspCode & zc  
                        zc = "" 
                    end if 
                else
                    aspCode = aspCode & sx 
                end if 
            else
                aspCode = aspCode & sx 
            end if 
            doEvents 
        next 
        aspCode = rTrim(aspCode)                                                        '去除右边空格
        aspCode = aspCode & vbCrLf                                                      'Asp换行
        doEvents 
    next 
    jsCodeConfusion = aspCode 
end function 


'删除JS注释 20160602
function delJsNote(content)
    dim splStr, s, c, isMultiLineNote, s2 
    isMultiLineNote = false                                                         '多行注释默认为假
    splStr = split(content, vbCrLf) 
    for each s in splStr
        s2 = phptrim(s) 
        if isMultiLineNote = true then
            if len(s2) >= 2 then
                if right(s2, 2) = "*/" then
                    isMultiLineNote = false 
                end if 
            end if 
            s = "" 
        else
            if left(s2, 2) = "/*" then
                if right(s2, 2) <> "*/" then
                    isMultiLineNote = true 
                end if 
                s = "" 
            elseif left(s2, 2) = "//" then
                s = "" 
            end if 
        end if 
        c = c & s & vbCrLf 
    next 
    delJsNote = c 
end function 

'JS转换，引用别人
function jsEncode__(byVal s)
    if isNul(s) then 
		jsEncode__ = ""
		exit function 
	end if
    dim arr1, arr2, i, j, c, p, t 
    arr1 = array(chr(34), chr(92), chr(47), chr(8), chr(12), chr(10), chr(13), chr(9)) '34|",92|\,47|/,8|,12|,10| ,13| ,9|    ,
    arr2 = array(chr(34), chr(92), chr(47), chr(98), chr(102), chr(110), chr(114)) '34|",92|\,47|/,98|b,102|f,110|n,114|r,1865|,
    for i = 1 to len(s)
        p = true 
        c = mid(s, i, 1) 
        for j = 0 to uBound(arr1)
            if c = arr1(j) then
                t = t & "\" & arr2(j) 
                p = false 
                exit for 
            end if 
        next 
        if p then t = t & c 
    next 
    jsEncode__ = t 
end function 


'远程网站会员统计2010330
'<script>document.writeln("<script src=\'http://127.0.0.1/web_soft/R.Asp?act=Stat&GoToUrl="+escape(document.referrer)+"&ThisUrl="+escape(window.location.href)+"&screen="+escape(window.screen.width+"x"+window.screen.height)+"&co="+escape(document.cookie)+" \'><\/script>");<'/script>
function showStatJSCode(url)
    showStatJSCode = "<script>document.writeln(""<script src=\'" & url & "act=Stat&GoToUrl=""+escape(document.referrer)+""&ThisUrl=""+escape(window.location.href)+""&screen=""+escape(window.screen.width+""x""+window.screen.height)+""&co=""+escape(document.cookie)+"" \'><\/script>"");</"&"script>" 
end function 


'Js定时跳转 Timing = 定时 时间测定 例：Call Rw("账号或密码错误，" & JsTiming("返回", 5))
function jsTiming(url, nSeconds)
    dim c 
    c = c & "<span id=mytimeidboyd>倒计时</span>" & vbCrLf 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "//配置Config" & vbCrLf 
    c = c & "var coutnumb" & vbCrLf 
    c = c & "coutnumb=" & nSeconds & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "//定时跳转" & vbCrLf 
    c = c & "function Countdown(){" & vbCrLf                                        'Countdown=倒数计秒
    c = c & "    coutnumb-=1" & vbCrLf 
    c = c & "    mytimeidboyd.innerHTML=""倒计时<font color=#000000>""+coutnumb+""</font>""" & vbCrLf 
    c = c & "    if(coutnumb<1){    " & vbCrLf 

    if url = "back" or url = "返回" then                                            '当Action为back是返回上页
        c = c & "        history.back();" & vbCrLf 
    else
        c = c & "        location.href='" & url & "';" & vbCrLf 
    end if 


    c = c & "    }else{" & vbCrLf 
    c = c & "        setTimeout(""Countdown()"",1000);" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}setTimeout(""Countdown()"",1)" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    jsTiming = c 
end function 
'JS弹窗 Call Javascript("返回", "操作成功", "")
function javascript(action, msg, url)
    if msg <> "" then msg = "alert('" & msg & "');"                                 '当Msg不为空则弹出信息
    if action = "back" or action = "返回" then                                      '当Action为back是返回上页
        response.write("<script>" & msg & "history.back();</"&"script>") 
    elseIf url <> "" then                                                           '当Url不为空
        response.write("<script>" & msg & "location.href='" & url & "';</"&"script>") '跳转Url页
    else
        response.write("<script>" & msg & "</"&"script>") 
    end if 
    response.end() 
end function 
'创建Ajax对象实例
function createAjax()
    dim c 
    c = "<script language=""javascript"">" & vbCrLf 
    c = c & "//AjAX XMLHTTP对象实例" & vbCrLf 
    c = c & "function createAjax() { " & vbCrLf 
    c = c & "    var _xmlhttp;" & vbCrLf 
    c = c & "    try {    " & vbCrLf 
    c = c & "        _xmlhttp=new ActiveXObject(""Microsoft.XMLHTTP"");    //IE的创建方式" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    catch (e) {" & vbCrLf 
    c = c & "        try {" & vbCrLf 
    c = c & "            _xmlhttp=new XMLHttpRequest();    //FF等浏览器的创建方式" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "        catch (e) {" & vbCrLf 
    c = c & "            _xmlhttp=false;        //如果创建失败，将返回false" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    return _xmlhttp;    //返回xmlhttp对象实例" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "//Ajax" & vbCrLf 
    c = c & "function Ajax(URL,ShowID) {  " & vbCrLf 
    c = c & "    var xmlhttp=createAjax();" & vbCrLf 
    c = c & "    if (xmlhttp) {" & vbCrLf 
    c = c & "        URL+= ""&n=""+Math.random() " & vbCrLf 
    c = c & "        xmlhttp.open('post', URL, true);//基本方法" & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""cache-control"",""no-cache""); " & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""Content-Type"", ""application/x-www-form-urlencoded"");         " & vbCrLf 
    c = c & "        xmlhttp.onreadystatechange=function() {        " & vbCrLf 
    c = c & "            if (xmlhttp.readyState==4 && xmlhttp.status==200) {     " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""操作完成""// unescape(xmlhttp.responseText); " & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "            else {                " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""正在加载中...""" & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "        }" & vbCrLf 
    'c=c & "alert(document.all.TEXTContent.value)" & vbcrlf
    c = c & "        xmlhttp.send(""Content=""+escape(document.all.TEXTContent.value)+"""");    " & vbCrLf 
    c = c & "        //alert(""网络错误"");" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function GetIDHTML(Root){" & vbCrLf 
    c = c & "    alert(document.all[Root].innerHTML)" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    createAjax = c 
end function 
'JS在线编辑
function onLineEditJS()
    dim c 
    c = c & "<script language=""javascript"">" & vbCrLf 
    c = c & "//显示与编辑内容，但不修改，ASP代码分析器用 创作于2013,10,5" & vbCrLf 
    c = c & "function TestInput(Root){ " & vbCrLf 
    c = c & "    var TempContent" & vbCrLf 
    c = c & "    TempContent="""" " & vbCrLf 
    c = c & "    " & vbCrLf 
    c = c & "    document.all[Root].title=""""" & vbCrLf 
    c = c & "    if(document.all[Root].innerHTML.indexOf(""<TEXTAREA"")==-1){" & vbCrLf 
    c = c & "            TempContent=document.all[Root].innerHTML" & vbCrLf 
    c = c & "            TempContent=TempContent.replace(/<BR><BR>/g,""<BR>"");     " & vbCrLf 
    c = c & "            TempContent=TempContent.replace(/<BR>/g,""\n"");     " & vbCrLf 
    c = c & "            if(TempContent==""&nbsp;""){TempContent=""""}" & vbCrLf 
    c = c & "            document.all[Root].innerHTML=""<textarea name=TEXT""+Root+"" style='width:50%;height:50%' onblur=if(this.value!=''){document.all.""+Root+"".title='点击可编辑';document.all.""+Root+"".innerHTML=ReplaceNToBR(this.value)}else{document.all.""+Root+"".innerHTML='&nbsp;'};>"" + TempContent + ""</textarea>"";" & vbCrLf 
    c = c & "            document.all[""TEXT""+Root].focus();" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function ReplaceNToBR(Content){" & vbCrLf 
    c = c & "    return Content.replace(/\n/g,""<BR>"")" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    onLineEditJS = c 
end function 
'在线编辑
function editTXT(content, jsId)
    content = IIF(content = "", "&nbsp;", content) 
    editTXT = "<span id='" & jsId & "' onClick=""TestInput('" & jsId & "');"" title='点击可编辑'>" & content & "</span>" 
end function 
'在线编辑  (辅助)
function onLineEdit(content, jsId)
    onLineEdit = editTXT(content, jsId) 
end function 
'****************************************************
'函数名：JSGoTo
'作  用：显示文本
'时  间：2013年12月14日
'参  数：Url
'*       SetTime
'返回值：字符串
'调  试：Call Echo("测试函数 JSGoTo", JSGoTo("", "",""))
'****************************************************
function jsGoTo(title, url, nSetTime)
    dim c 
    if title = "" then title = "添加成功" 
    'if nSetTime = "" then nSetTime = 4                                                '默认为4秒
    c = c & "<script>" & vbCrLf 
    c = c & "//通用定时器 如：MyTimer('Show', 'alert(1+1)', 5)" & vbCrLf 
    c = c & "var StopTimer = """"" & vbCrLf 
    c = c & "function MyTimer(ID, ActionStr,TimeNumb){" & vbCrLf 
    c = c & "    if(StopTimer == ""停止"" || StopTimer == ""停止定时器""){" & vbCrLf 
    c = c & "        StopTimer = """"" & vbCrLf 
    c = c & "        return false" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    TimeNumb--" & vbCrLf 
    c = c & "    document.all[ID].innerHTML = ""倒计时："" + TimeNumb" & vbCrLf 
    c = c & "    if(TimeNumb<1){" & vbCrLf 
    c = c & "        setTimeout(ActionStr,100);" & vbCrLf 
    c = c & "    }else{" & vbCrLf 
    c = c & "        setTimeout(""MyTimer('""+ID+""', '""+ActionStr+""',""+TimeNumb+"")"",1000);" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function GotoURL(){" & vbCrLf 
    c = c & "    location.href='" & url & "'" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    c = c & "<div id=""Show"">Loading...</div><script>MyTimer('Show', 'GotoURL()', " & nSetTime & ")</"&"script>" & vbCrLf 
    jsGoTo = c 
end function 

'JS图片滚动
function jsPhotoScroll(id, width, height)
    dim c 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "    var marqueeB = new Marquee(""" & id & """)    " & vbCrLf 
    c = c & "    marqueeB.Direction =2;" & vbCrLf 
    c = c & "    marqueeB.Step = 1;" & vbCrLf 
    c = c & "    marqueeB.Width = " & width & ";" & vbCrLf 
    c = c & "    marqueeB.Height = " & height & ";" & vbCrLf 
    c = c & "    marqueeB.Timer = 1;" & vbCrLf 
    c = c & "    marqueeB.DelayTime = 0;" & vbCrLf 
    c = c & "    marqueeB.WaitTime = 0;" & vbCrLf 
    c = c & "    marqueeB.ScrollStep = 20;" & vbCrLf 
    c = c & "    marqueeB.Start();    " & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    jsPhotoScroll = c 
end function 
'图片向左滚动（暂不用）
function photoLeftScroll(demo, demo1, demo2)
    dim c 
    c = c & "<!--图片向左轮番滚动-->" & vbCrLf 
    c = c & "<script language=""javascript"">" & vbCrLf 
    c = c & "var speed=30" & vbCrLf 
    c = c & "" & demo2 & ".innerHTML=" & demo1 & ".innerHTML" & vbCrLf 
    c = c & "function Marquee(){" & vbCrLf 
    c = c & "    if(" & demo2 & ".offsetWidth-" & demo & ".scrollLeft<=0)" & vbCrLf 
    c = c & "        " & demo & ".scrollLeft-=" & demo1 & ".offsetWidth" & vbCrLf 
    c = c & "    else{" & vbCrLf 
    c = c & "        " & demo & ".scrollLeft++" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "var MyMar=setInterval(Marquee,speed)" & vbCrLf 
    c = c & "" & demo & ".onmouseover=function() {clearInterval(MyMar)}" & vbCrLf 
    c = c & "" & demo & ".onmouseout=function() {MyMar=setInterval(Marquee,speed)}" & vbCrLf 
    c = c & "</"&"script> " & vbCrLf 
    photoLeftScroll = c 
end function 
%>  


