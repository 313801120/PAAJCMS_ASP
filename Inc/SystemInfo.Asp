<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-05
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'系统信息  (2014,05,27)



'操作系统版本
function operationSystem()
    dim httpAgent, systemVer 
    httpAgent = request.serverVariables("HTTP_USER_AGENT") 
    if inStr(httpAgent, "NT 5.2") > 0 then
        systemVer = "Windows Server 2003" 
    elseIf inStr(httpAgent, "NT 5.1") > 0 then
        systemVer = "Windows XP" 
    elseIf inStr(httpAgent, "NT 5") > 0 then
        systemVer = "Windows 2000" 
    elseIf inStr(httpAgent, "NT 4") > 0 then
        systemVer = "Windows NT4" 
    elseIf inStr(httpAgent, "4.9") > 0 then
        systemVer = "Windows ME" 
    elseIf inStr(httpAgent, "98") > 0 then
        systemVer = "Windows 98" 
    elseIf inStr(httpAgent, "95") > 0 then
        systemVer = "Windows 95" 
    else
        systemVer = httpAgent 
    end if 
    operationSystem = httpAgent 
end function 
'检测是否为手机
function checkMobile()
    checkMobile = false 
    if request.serverVariables("HTTP_X_WAP_PROFILE") <> "" then
        checkMobile = true 
    end if 
end function 

'获得IIS版本号
function getIISVersion()
    getIISVersion = request.serverVariables("SERVER_SOFTWARE") 
end function 

%> 


