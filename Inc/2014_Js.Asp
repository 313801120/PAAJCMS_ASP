<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-04
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'Js

'Asp����������� 20160624 
function jsCodeConfusion(content)
    dim splStr, i, yesJs, yesWord, sx, s, wc, zc, s1, aspCode, upWord 
    dim upWordn, tempS, dimList 
    dim yesFunction                                                                 '�����Ƿ�Ϊ��
    dim isStartFunction                                                               '��ʼ���� Ŀ����Ϊ����function default ������������û��()   20150218
    dim wcType                                                                      '�����ı����ͣ��� " �� '
    dim isAddToSYH                                                                  '�Ƿ��ۼ�˫����
    dim beforeStr, afterStr, endCode, nSYHCount 
	
    isStartFunction = false                                                           'Ĭ�Ͽ�ʼ����Ϊ��
    'If nType="" Then  nType  = 0
    yesJs = false                                                                   '��ASP Ĭ��Ϊ��
    yesFunction = false                                                             '�Ǻ��� Ĭ��Ϊ��
    yesWord = false                                                                 '�ǵ��� Ĭ��Ϊ��
    nSYHCount = 0                                                                    '˫����Ĭ��Ϊ0
    splStr = split(content, vbCrLf)                                                 '�ָ���
    'ѭ������
    for each s in splStr
        'ѭ��ÿ���ַ�
        for i = 1 to len(s)
            sx = mid(s, i, 1) 
            'Asp��ʼ
            if sx = "<" and wc = "" then                                                 '����ı�����Ϊ�� WcΪ������� ��"<%" �ų� �޸���20140412
                if mid(s, i + 1, 6) = "script" then
                    yesJs = true                                                                    'ASPΪ��
                    i = i + 1                                                                       '��1�����ܼ�2��Ҫ��Ȼ<%function Test() �ͽ�ȡ����
                    sx = mid(s, i, 1) 
                    aspCode = aspCode & "<" 
                end if 
            'ASP����
            elseIf sx = "<" and mid(s, i + 1, 8) = "/script>" and wc = "" then           'WcΪ�������
                yesJs = false                                                                   'ASPΪ��
                i = i + 1                                                                       '���ܼ�2��ֻ�ܼ�1����Ϊ���ﶨ��ASPΪ�٣���������һ����ʾ����� 'ASP����Ϊ��
                sx = mid(s, i, 8) 
                aspCode = aspCode & "/script>" 
            end if 
            if yesJs = true then

                beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '��һ���ַ�
                afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '��һ���ַ�
                endCode = mid(s, i + 1)                                                         '��ǰ�ַ���������� һ��
                '�����ı�
                if(sx = """" or sx = "'" and wcType = "") or sx = wcType or wc <> "" then
                    isAddToSYH = true 
                    '����һ�ּ򵥵ķ�����������(20150914)
                    if isAddToSYH = true and beforeStr = "\" then

                        if len(wc) >= 1 then
                            if isStrTransferred(wc) = true then                     'Ϊת���ַ�Ϊ��
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

                    '˫�����ۼ�
                    if sx = wcType and isAddToSYH = true then nSYHCount = nSYHCount + 1             '�ų���һ���ַ�Ϊ\���ת���ַ�(20150914)


                    '�ж��Ƿ�"�����
                    if nSYHCount mod 2 = 0 and beforeStr <> "\" then
                        if mid(s, i + 1, 1) <> wcType then
                            wc = wc & sx 
                            aspCode = aspCode & wc                                  '�д����ۼ�
                            'call echo("wc",wc)
                            nSYHCount = 0 : wc = ""                                 '���
                            wcType = "" 
                        else
                            wc = wc & sx 
                        end if 
                    else
                        wc = wc & sx 
                    end if 

                elseIf sx = "'" then                                                            'ע�����˳�
                    aspCode = aspCode & mid(s, i) 
                    exit for 
                '��ĸ
                elseIf checkABC(sx) = true or(sx = "_" and zc <> "") or zc <> "" then
                    zc = zc & sx 
                    s1 = LCase(mid(s & " ", i + 1, 1)) 
                    if inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1) = 0 and not(s1 = "_" and zc <> "") then'����ж�
                        tempS = mid(s, i + 1) 
  
                        if inStr("|function|sub|", "|" & LCase(zc) & "|")>0 then
                            '������ʼ
                            if yesFunction = false and LCase(upWord) <> "end" then
                                yesFunction = true 
                                dimList = getFunDimName(tempS) 
                                isStartFunction = true 
                            elseIf yesFunction = true and LCase(upWord) = "end" then'�����һ������
                                yesFunction = false 
                            end if 
                        elseIf yesFunction = true and LCase(zc) = "var" then
                            dimList = dimList & "," & getVarName(tempS) 
                        elseIf yesFunction = true then
                            '�ų���������ÿһ������
                            if isStartFunction = false then
                                zc = replaceDim2(dimList, zc) 
                            end if 
                            isStartFunction = false 
                        end if 
                        upWord = zc                                                                     '��ס��ǰ����
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
        aspCode = rTrim(aspCode)                                                        'ȥ���ұ߿ո�
        aspCode = aspCode & vbCrLf                                                      'Asp����
        doEvents 
    next 
    jsCodeConfusion = aspCode 
end function 


'ɾ��JSע�� 20160602
function delJsNote(content)
    dim splStr, s, c, isMultiLineNote, s2 
    isMultiLineNote = false                                                         '����ע��Ĭ��Ϊ��
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

'JSת�������ñ���
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


'Զ����վ��Աͳ��2010330
'<script>document.writeln("<script src=\'http://127.0.0.1/web_soft/R.Asp?act=Stat&GoToUrl="+escape(document.referrer)+"&ThisUrl="+escape(window.location.href)+"&screen="+escape(window.screen.width+"x"+window.screen.height)+"&co="+escape(document.cookie)+" \'><\/script>");<'/script>
function showStatJSCode(url)
    showStatJSCode = "<script>document.writeln(""<script src=\'" & url & "act=Stat&GoToUrl=""+escape(document.referrer)+""&ThisUrl=""+escape(window.location.href)+""&screen=""+escape(window.screen.width+""x""+window.screen.height)+""&co=""+escape(document.cookie)+"" \'><\/script>"");</"&"script>" 
end function 


'Js��ʱ��ת Timing = ��ʱ ʱ��ⶨ ����Call Rw("�˺Ż��������" & JsTiming("����", 5))
function jsTiming(url, nSeconds)
    dim c 
    c = c & "<span id=mytimeidboyd>����ʱ</span>" & vbCrLf 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "//����Config" & vbCrLf 
    c = c & "var coutnumb" & vbCrLf 
    c = c & "coutnumb=" & nSeconds & "" & vbCrLf 
    c = c & "" & vbCrLf 
    c = c & "//��ʱ��ת" & vbCrLf 
    c = c & "function Countdown(){" & vbCrLf                                        'Countdown=��������
    c = c & "    coutnumb-=1" & vbCrLf 
    c = c & "    mytimeidboyd.innerHTML=""����ʱ<font color=#000000>""+coutnumb+""</font>""" & vbCrLf 
    c = c & "    if(coutnumb<1){    " & vbCrLf 

    if url = "back" or url = "����" then                                            '��ActionΪback�Ƿ�����ҳ
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
'JS���� Call Javascript("����", "�����ɹ�", "")
function javascript(action, msg, url)
    if msg <> "" then msg = "alert('" & msg & "');"                                 '��Msg��Ϊ���򵯳���Ϣ
    if action = "back" or action = "����" then                                      '��ActionΪback�Ƿ�����ҳ
        response.write("<script>" & msg & "history.back();</"&"script>") 
    elseIf url <> "" then                                                           '��Url��Ϊ��
        response.write("<script>" & msg & "location.href='" & url & "';</"&"script>") '��תUrlҳ
    else
        response.write("<script>" & msg & "</"&"script>") 
    end if 
    response.end() 
end function 
'����Ajax����ʵ��
function createAjax()
    dim c 
    c = "<script language=""javascript"">" & vbCrLf 
    c = c & "//AjAX XMLHTTP����ʵ��" & vbCrLf 
    c = c & "function createAjax() { " & vbCrLf 
    c = c & "    var _xmlhttp;" & vbCrLf 
    c = c & "    try {    " & vbCrLf 
    c = c & "        _xmlhttp=new ActiveXObject(""Microsoft.XMLHTTP"");    //IE�Ĵ�����ʽ" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    catch (e) {" & vbCrLf 
    c = c & "        try {" & vbCrLf 
    c = c & "            _xmlhttp=new XMLHttpRequest();    //FF��������Ĵ�����ʽ" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "        catch (e) {" & vbCrLf 
    c = c & "            _xmlhttp=false;        //�������ʧ�ܣ�������false" & vbCrLf 
    c = c & "        }" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    return _xmlhttp;    //����xmlhttp����ʵ��" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "//Ajax" & vbCrLf 
    c = c & "function Ajax(URL,ShowID) {  " & vbCrLf 
    c = c & "    var xmlhttp=createAjax();" & vbCrLf 
    c = c & "    if (xmlhttp) {" & vbCrLf 
    c = c & "        URL+= ""&n=""+Math.random() " & vbCrLf 
    c = c & "        xmlhttp.open('post', URL, true);//��������" & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""cache-control"",""no-cache""); " & vbCrLf 
    c = c & "        xmlhttp.setRequestHeader(""Content-Type"", ""application/x-www-form-urlencoded"");         " & vbCrLf 
    c = c & "        xmlhttp.onreadystatechange=function() {        " & vbCrLf 
    c = c & "            if (xmlhttp.readyState==4 && xmlhttp.status==200) {     " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""�������""// unescape(xmlhttp.responseText); " & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "            else {                " & vbCrLf 
    c = c & "                document.getElementById(ShowID).innerHTML = ""���ڼ�����...""" & vbCrLf 
    c = c & "            }" & vbCrLf 
    c = c & "        }" & vbCrLf 
    'c=c & "alert(document.all.TEXTContent.value)" & vbcrlf
    c = c & "        xmlhttp.send(""Content=""+escape(document.all.TEXTContent.value)+"""");    " & vbCrLf 
    c = c & "        //alert(""�������"");" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function GetIDHTML(Root){" & vbCrLf 
    c = c & "    alert(document.all[Root].innerHTML)" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    createAjax = c 
end function 
'JS���߱༭
function onLineEditJS()
    dim c 
    c = c & "<script language=""javascript"">" & vbCrLf 
    c = c & "//��ʾ��༭���ݣ������޸ģ�ASP����������� ������2013,10,5" & vbCrLf 
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
    c = c & "            document.all[Root].innerHTML=""<textarea name=TEXT""+Root+"" style='width:50%;height:50%' onblur=if(this.value!=''){document.all.""+Root+"".title='����ɱ༭';document.all.""+Root+"".innerHTML=ReplaceNToBR(this.value)}else{document.all.""+Root+"".innerHTML='&nbsp;'};>"" + TempContent + ""</textarea>"";" & vbCrLf 
    c = c & "            document.all[""TEXT""+Root].focus();" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "function ReplaceNToBR(Content){" & vbCrLf 
    c = c & "    return Content.replace(/\n/g,""<BR>"")" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</"&"script>" & vbCrLf 
    onLineEditJS = c 
end function 
'���߱༭
function editTXT(content, jsId)
    content = IIF(content = "", "&nbsp;", content) 
    editTXT = "<span id='" & jsId & "' onClick=""TestInput('" & jsId & "');"" title='����ɱ༭'>" & content & "</span>" 
end function 
'���߱༭  (����)
function onLineEdit(content, jsId)
    onLineEdit = editTXT(content, jsId) 
end function 
'****************************************************
'��������JSGoTo
'��  �ã���ʾ�ı�
'ʱ  �䣺2013��12��14��
'��  ����Url
'*       SetTime
'����ֵ���ַ���
'��  �ԣ�Call Echo("���Ժ��� JSGoTo", JSGoTo("", "",""))
'****************************************************
function jsGoTo(title, url, nSetTime)
    dim c 
    if title = "" then title = "���ӳɹ�" 
    'if nSetTime = "" then nSetTime = 4                                                'Ĭ��Ϊ4��
    c = c & "<script>" & vbCrLf 
    c = c & "//ͨ�ö�ʱ�� �磺MyTimer('Show', 'alert(1+1)', 5)" & vbCrLf 
    c = c & "var StopTimer = """"" & vbCrLf 
    c = c & "function MyTimer(ID, ActionStr,TimeNumb){" & vbCrLf 
    c = c & "    if(StopTimer == ""ֹͣ"" || StopTimer == ""ֹͣ��ʱ��""){" & vbCrLf 
    c = c & "        StopTimer = """"" & vbCrLf 
    c = c & "        return false" & vbCrLf 
    c = c & "    }" & vbCrLf 
    c = c & "    TimeNumb--" & vbCrLf 
    c = c & "    document.all[ID].innerHTML = ""����ʱ��"" + TimeNumb" & vbCrLf 
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

'JSͼƬ����
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
'ͼƬ����������ݲ��ã�
function photoLeftScroll(demo, demo1, demo2)
    dim c 
    c = c & "<!--ͼƬ�����ַ�����-->" & vbCrLf 
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

