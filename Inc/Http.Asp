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
'�������url״̬
function getHttpUrlState(httpurl)
    getHttpUrlState = handleGetHttpPage(httpurl, "")(1) 
end function 
'�������url�ķ���������
function getHttpUrlServerName(httpurl)
    getHttpUrlServerName = ""                                                       '����php���õ�
end function 

'��òɼ����� 
function getHttpPage(byVal httpurl, byVal sCharSet)
    getHttpPage = handleGetHttpPage(httpurl, sCharSet)(0)   
end function


'��òɼ����� (����)
function getURLContent(httpurl, sCharSet)
    getURLContent = getHttpPage(httpurl, sCharSet) 
end function 
'��òɼ����� (����)
function getHttpUrl(httpurl, sCharSet)
    getHttpUrl = getHttpPage(httpurl, sCharSet) 
end function 
function bytesToBstr(byteArr, cset)
    dim objStream 
    if isNul(byteArr) then exit function                                               'Ϊ�����˳�
    set objStream = createObject("ADODB.Stream")
        objStream.type = 1 
        objStream.mode = 3 
        objStream.open 
        call objStream.write(byteArr) 
        objStream.position = 0 
        objStream.type = 2 
        objStream.charset = cset 
        bytesToBstr = objStream.readText 
        objStream.close 
    set objStream = nothing 
end function 
'������òɼ����� 
function handleGetHttpPage(byVal httpurl, byVal sCharSet)
    on error resume next 
    dim startTime, nStatus, nTime, content, sSplType, webFileSize 
    startTime = now() 
    sSplType = "|-*ypia=|" 
    sCharSet = handleStrCharSet(sCharSet) 
    if isNull(httpurl) = true or len(httpurl) < 7 or httpurl = "$False$" then
        handleGetHttpPage = split("" & sSplType & - 1 & sSplType & nTime & sSplType & webFileSize, sSplType) 
        exit function 
    end if 
 
    dim http 
    set http = createObject("MSXML2.XMLHTTP")
        call http.open("GET", httpurl, false) 
        http.send 
        if http.readyState <> 4 then
            set http = nothing 
            handleGetHttpPage = split("" & sSplType & http.readyState & sSplType & nTime & sSplType & webFileSize, sSplType) 
            exit function 
        end if 

        content = bytesToBstr(http.responseBody, sCharSet) 
        nStatus = http.status 
        nTime = dateDiff("s", startTime, now()) 
        webFileSize = strLength(content) 

        handleGetHttpPage = split(content & sSplType & nStatus & sSplType & nTime & sSplType & webFileSize, sSplType) 

    set http = nothing 
end function 

%>