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
'Time ʱ������� (2013,9,27)

'ʱ�䴦��    
function format_Time(timeStr, nType)
    dim y, m, d, h, mi, s ,c,nWeek
    if isDate(timeStr) = false then exit function 
    y = CStr(year(timeStr)) 
    m = CStr(month(timeStr))  
    if len(m) = 1 then m = "0" & m 
    d = CStr(day(timeStr))                                                          '��vb.net��Ҫ������  D = CStr(CDate(timeStr).Day)
	'��
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
            'yyyy��mm��dd��
            c = y & "��" & m & "��" & d & "��" 
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
            'yyyy��mm��dd��
            c = y & "��" & m & "��" & d & "��" & " " & h & ":" & mi & ":" & s 
        case 9:
            'yyyy��mm��dd��Hʱmi��S�� ����
            c = y & "��" & m & "��" & d & "��" & " " & h & "ʱ" & mi & "��" & s & "�룬" & getDayStatus(h, 1) 
        case 10:
            'yyyy��mm��dd��Hʱ
            c = y & "��" & m & "��" & d & "��" & h & "ʱ" 
        case 11:
            'yyyy��mm��dd��Hʱmi��S��
            c = y & "��" & m & "��" & d & "��" & " " & h & "ʱ" & mi & "��" & s & "��" 
        case 12:
            'yyyy��mm��dd��Hʱmi��
            c = y & "��" & m & "��" & d & "��" & " " & h & "ʱ" & mi & "��" 
        case 13:
            'yyyy��mm��dd��Hʱmi�� ����
            c = m & "��" & d & "��" & " " & h & ":" & mi & " " & getDayStatus(h, 0) 
        case 14:
            'yyyy��mm��dd��
            c = y & "/" & m & "/" & d 
        case 15:
            'yyyy��mm�� ��1��
            c = y & "��" & m & "�� ��" & nWeek & "��" 
    end select
	format_Time=c
end function 
'��õ�ǰʱ�ڻ����Լ���
function getHandleDate(numb)
    if numb = "" then
        numb = 0 
    end if 
    getHandleDate = format_Time(now() + numb, 2) 
end function 
'��õ�ǰ��״̬
function getDayStatus(h, nType)
    dim c,nH
    if left(h, 1) = "0" then
        h = right(h, 1) 
    end if 
    nH = CInt(h) 
    if nH >= 0 and nH <= 5 then
        c = "�賿" 
    elseIf nH >= 6 and nH <= 8 then
        c = "����" 
    elseIf nH >= 9 and nH <= 12 then
        c = "����" 
    elseIf nH >= 13 and nH <= 18 then
        c = "����" 
    elseIf nH >= 19 and nH <= 24 then
        c = "����" 
    else
        c = "��ҹ" 
    end if 
    if nType = 1 then c = "<b>" & c & "</b>" 
    getDayStatus = c 
end function 
'ʱ�����
function printTimeValue(byval nV)
    on error resume next 
    dim nVTemp, c, n 
	nVTemp=nV
    if nV >= 86400 then
        n = nV mod 86400 
        nV = fix(nV / 60 / 60 / 24)        '�������ǳ�24������60����Ϊһ����24��Сʱ��С�������ɵX  20150119
        c = c & nV & "��" 
        nV = n 
    end if 
    if nV >= 3600 then
        n = nV mod 3600 
        nV = fix(nV / 60 / 60) 
        c = c & nV & "Сʱ" 
        nV = n 
    end if 

    if nV >= 60 then
        n = nV mod 60 
        nV = fix(nV / 60) 
        c = c & nV & "��" 
        nV = n 
    end if 
    if nV > 0 then
        c = c & nV & "��" 
    end if 
    if err then call doError(err.description, "printTimeValue ʱ����㣬nV=" & nVTemp) 
    printTimeValue = c 
end function 
'������ʱ  (�ʴ���Сʱ�������ʾ)
function printAskTime(nV)
    on error resume next 
    dim  c, n  
    if nV >= 3600 then 
        n = nV mod 3600 
        nV = fix(nV / 60 / 60) 
        c = c & nV & "Сʱ" 
        nV = n 
        printAskTime = c : exit function 
    end if 
    if nV >= 60 then
        n = nV mod 60 
        nV = fix(nV / 60) 
        c = c & nV & "����" 
        nV = n 
        printAskTime = c : exit function 
    end if 
    if nV >= 0 then
        c = c & nV & "����"  
        printAskTime = c : exit function 
    end if    
end function 
'���ʱ��
function getTimerSet()
    getTimerSet = calculationTimer(pubTimer,now)
end function  
'����ʱ��  
function calculationTimer(startTime,endTime)   
    dim n 
    'n = formatNumber((timer() - pubTimer) * 1000, 2, - 1) / 1000 
    'calculationTimer = toNumber(n, 3)   
	calculationTimer=dateDiff("s", startTime,endTime)
end function  

'���ʱ��
function getTimer()
    getTimer = "<br>------------------------------------<br>ҳ��ִ��ʱ�䣺" & getTimerSet() & " ��" 
end function 
'���ʱ��
function vbGetTimer()
    VBGetTimer = vbCrLf & "------------------------------------" & vbCrLf & "����ʱ�䣺" & calculationTimer(pubTimer,now())
end function 
'��õ�����
function vbEchoTimer()
    VBEchoTimer = "����ʱ�䣺" & calculationTimer(pubTimer,now()) & " ��"
end function 
'���ʱ�������
function vbRunTimer(startTime)
    VBRunTimer = "����ʱ�䣺" & calculationTimer(pubTimer,now()) & " ��"
end function



'����ʱ��
function sAddTime(timeObj,sType,nValue)
	dim s
	'��
	if sType="s" then
		s=timeObj+0.00001*nValue
	'����
	elseif sType="n" then
		s=timeObj+0.00060*nValue
	'Сʱ
	elseif sType="h" then
		s=timeObj+0.00001*nValue*60*60
	'��
	elseif sType="d" then
		s=timeObj+0.00001*nValue*60*60*24
	'��
	elseif sType="w" then
		s=timeObj+0.00001*nValue*60*60*24*7
	'��
	elseif sType="m" then
		s=timeObj+0.00001*nValue*60*60*24*30
	'����
	elseif sType="q" then
		s=timeObj+0.00001*nValue*60*60*24*120
	'��
	elseif sType="y" then
		s=timeObj+0.00001*nValue*60*60*24*366
	end if
	sAddTime=s
end function
%>   

