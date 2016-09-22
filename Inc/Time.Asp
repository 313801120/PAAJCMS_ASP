<%
'************************************************************
'作者：云孙World(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-09-22
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'Time 时间操作类 (2013,9,27)

'时间处理    
function format_Time(timeStr, nType)
    dim y, m, d, h, mi, s ,c,nWeek
    if isDate(timeStr) = false then exit function 
    y = CStr(year(timeStr)) 
    m = CStr(month(timeStr))  
    if len(m) = 1 then m = "0" & m 
    d = CStr(day(timeStr))                                                          '在vb.net里要这样用  D = CStr(CDate(timeStr).Day)
	'周
	nWeek=fix(day(timeStr)/7)
	if day(timeStr) mod 7>0 then
		nWeek=nWeek+1
	end if
	
    if len(d) = 1 then d = "0" & d 
    h = CStr(hour(timeStr))  
    if len(h) = 1 then h = "0" & h 
    mi = CStr(minute(timeStr)) 
    if len(mi) = 1 then mi = "0" & mi 
    s = CStr(second(timeStr)) 
    if len(s) = 1 then s = "0" & s
    select case nType
        case 1:
            'yyyy-mm-dd hh:mm:ss
            c = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
        case 2:
            'yyyy-mm-dd
            c = y & "-" & m & "-" & d 
        case 3:
            'hh:mm:ss
            c = h & ":" & mi & ":" & s 
        case 4:
            'yyyy年mm月dd日
            c = y & "年" & m & "月" & d & "日" 
        case 5:
            'yyyymmdd
            c = y & m & d 
        case 6:
            'yyyymmddhhmmss
            c = y & m & d & h & mi & s 
        case 7:
            'mm-dd
            c = m & "-" & d 
        case 8:
            'yyyy年mm月dd日
            c = y & "年" & m & "月" & d & "日" & " " & h & ":" & mi & ":" & s 
        case 9:
            'yyyy年mm月dd日H时mi分S秒 早上
            c = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" & s & "秒，" & getDayStatus(h, 1) 
        case 10:
            'yyyy年mm月dd日H时
            c = y & "年" & m & "月" & d & "日" & h & "时" 
        case 11:
            'yyyy年mm月dd日H时mi分S秒
            c = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" & s & "秒" 
        case 12:
            'yyyy年mm月dd日H时mi分
            c = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" 
        case 13:
            'yyyy年mm月dd日H时mi分 早上
            c = m & "月" & d & "日" & " " & h & ":" & mi & " " & getDayStatus(h, 0) 
        case 14:
            'yyyy年mm月dd日
            c = y & "/" & m & "/" & d 
        case 15:
            'yyyy年mm月 第1周
            c = y & "年" & m & "月 第" & nWeek & "周" 
    end select
	format_Time=c
end function 
'获得当前时期还可以计算  默认为加天数
function getHandleDate(numb)
    'if numb = "" then
    '    numb = 0 
    'end if
    'getHandleDate = format_Time(now() + numb, 2)
	getHandleDate=dateAdd("d",numb,now())
end function 
'获得当前天状态
function getDayStatus(h, nType)
    dim c,nH
    if left(h, 1) = "0" then
        h = right(h, 1) 
    end if 
    nH = CInt(h) 
    if nH >= 0 and nH <= 5 then
        c = "凌晨" 
    elseIf nH >= 6 and nH <= 8 then
        c = "早上" 
    elseIf nH >= 9 and nH <= 12 then
        c = "上午" 
    elseIf nH >= 13 and nH <= 18 then
        c = "下午" 
    elseIf nH >= 19 and nH <= 24 then
        c = "晚上" 
    else
        c = "深夜" 
    end if 
    if nType = 1 then c = "<b>" & c & "</b>" 
    getDayStatus = c 
end function 
'时间计算
function printTimeValue(byval nV)
    on error resume next 
    dim nVTemp, c, n 
	nVTemp=nV
    if nV >= 86400 then
        n = nV mod 86400 
        nV = fix(nV / 60 / 60 / 24)        '这里面是除24，不是60，因为一天有24个小时，小云你这个傻X  20150119
        c = c & nV & "天" 
        nV = n 
    end if 
    if nV >= 3600 then
        n = nV mod 3600 
        nV = fix(nV / 60 / 60) 
        c = c & nV & "小时" 
        nV = n 
    end if 

    if nV >= 60 then
        n = nV mod 60 
        nV = fix(nV / 60) 
        c = c & nV & "分" 
        nV = n 
    end if 
    if nV > 0 then
        c = c & nV & "秒" 
    end if 
    if err then call doError(err.description, "printTimeValue 时间计算，nV=" & nVTemp) 
    printTimeValue = c 
end function 
'计算整时  (问答以小时或分钟显示)
function printAskTime(nV)
    on error resume next 
    dim  c, n  
    if nV >= 3600 then 
        n = nV mod 3600 
        nV = fix(nV / 60 / 60) 
        c = c & nV & "小时" 
        nV = n 
        printAskTime = c : exit function 
    end if 
    if nV >= 60 then
        n = nV mod 60 
        nV = fix(nV / 60) 
        c = c & nV & "分钟" 
        nV = n 
        printAskTime = c : exit function 
    end if 
    if nV >= 0 then
        c = c & nV & "秒钟"  
        printAskTime = c : exit function 
    end if    
end function  
'获得时间
function getTimerSet()
    getTimerSet = calculationTimer(pubTimer,now)
end function  
'计算时间  
function calculationTimer(startTime,endTime)   
    dim n 
    'n = formatNumber((timer() - pubTimer) * 1000, 2, - 1) / 1000 
    'calculationTimer = toNumber(n, 3)   
	calculationTimer=dateDiff("s", startTime,endTime)
end function  

'获得时间
function getTimer()
    getTimer = "<br>------------------------------------<br>页面执行时间：" & getTimerSet() & " 秒" 
end function 
'获得时间
function vbGetTimer()
    VBGetTimer = vbCrLf & "------------------------------------" & vbCrLf & "运行时间：" & calculationTimer(pubTimer,now())
end function 
'获得第三种
function vbEchoTimer()
    VBEchoTimer = "运行时间：" & calculationTimer(pubTimer,now()) & " 秒"
end function 
'获得时间第四种
function vbRunTimer(startTime)
    VBRunTimer = "运行时间：" & calculationTimer(pubTimer,now()) & " 秒"
end function



'添加时间
function sAddTime(timeObj,sType,nValue)
	dim s
	'秒
	if sType="s" then
		s=timeObj+0.00001*nValue
	'分钟
	elseif sType="n" then
		s=timeObj+0.00060*nValue
	'小时
	elseif sType="h" then
		s=timeObj+0.00001*nValue*60*60
	'日
	elseif sType="d" then
		s=timeObj+0.00001*nValue*60*60*24
	'周
	elseif sType="w" then
		s=timeObj+0.00001*nValue*60*60*24*7
	'月
	elseif sType="m" then
		s=timeObj+0.00001*nValue*60*60*24*30
	'季节
	elseif sType="q" then
		s=timeObj+0.00001*nValue*60*60*24*120
	'年
	elseif sType="y" then
		s=timeObj+0.00001*nValue*60*60*24*366
	end if
	sAddTime=s
end function
%>   

