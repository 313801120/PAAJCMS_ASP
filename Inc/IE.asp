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
'����� 20160203

'��ȡ���������(�����ж�:47�������;GoogLe,Grub,MSN,Yahoo!֩��;ʮ�ֳ���IE���)  Response.Write getBrType("")
function getBrType(theInfo)
    dim strType, tmp1, s 
    s = "Other Unknown" 
    if theInfo = "" then
        theInfo = uCase(request.serverVariables("HTTP_USER_AGENT")) 
    end if 
    if inStr(theInfo, uCase("mozilla")) > 0 then s = "Mozilla" 
    if inStr(theInfo, uCase("icab")) > 0 then s = "iCab" 
    if inStr(theInfo, uCase("lynx")) > 0 then s = "Lynx" 
    if inStr(theInfo, uCase("links")) > 0 then s = "Links" 
    if inStr(theInfo, uCase("elinks")) > 0 then s = "ELinks" 
    if inStr(theInfo, uCase("jbrowser")) > 0 then s = "JBrowser" 
    if inStr(theInfo, uCase("konqueror")) > 0 then s = "konqueror" 
    if inStr(theInfo, uCase("wget")) > 0 then s = "wget" 
    if inStr(theInfo, uCase("ask jeeves")) > 0 or inStr(theInfo, uCase("teoma")) > 0 then s = "Ask Jeeves/Teoma" 
    if inStr(theInfo, uCase("wget")) > 0 then s = "wget" 
    if inStr(theInfo, uCase("opera")) > 0 then s = "opera" 
    if inStr(theInfo, uCase("NOKIAN")) > 0 then s = "NOKIAN(ŵ�����ֻ�)" 
    if inStr(theInfo, uCase("SPV")) > 0 then s = "SPV(���մ��ֻ�)" 
    if inStr(theInfo, uCase("Jakarta Commons")) > 0 then s = "Jakarta Commons-HttpClient" 
    if inStr(theInfo, uCase("Gecko")) > 0 then
        strType = "[Gecko] " 
        s = "Mozilla Series" 
        if inStr(theInfo, uCase("aol")) > 0 then s = "AOL" 
        if inStr(theInfo, uCase("netscape")) > 0 then s = "Netscape" 
        if inStr(theInfo, uCase("firefox")) > 0 then s = "FireFox" 
        if inStr(theInfo, uCase("chimera")) > 0 then s = "Chimera" 
        if inStr(theInfo, uCase("camino")) > 0 then s = "Camino" 
        if inStr(theInfo, uCase("galeon")) > 0 then s = "Galeon" 
        if inStr(theInfo, uCase("k-meleon")) > 0 then s = "K-Meleon" 
        s = strType & s 
    end if 
    if inStr(theInfo, uCase("bot")) > 0 or inStr(theInfo, uCase("crawl")) > 0 then
        strType = "[Bot/Crawler]" 
        if inStr(theInfo, uCase("grub")) > 0 then s = "Grub" 
        if inStr(theInfo, uCase("googlebot")) > 0 then s = "GoogleBot" 
        if inStr(theInfo, uCase("msnbot")) > 0 then s = "MSN Bot" 
        if inStr(theInfo, uCase("slurp")) > 0 then s = "Yahoo! Slurp" 
        s = strType & s 
    end if 
    if inStr(theInfo, uCase("applewebkit")) > 0 then
        strType = "[AppleWebKit]" 
        s = "" 
        if inStr(theInfo, uCase("omniweb")) > 0 then s = "OmniWeb" 
        if inStr(theInfo, uCase("safari")) > 0 then s = "Safari" 
        s = strType & s 
    end if 
    if inStr(theInfo, uCase("msie")) > 0 then
        strType = "[MSIE" 
        tmp1 = mid(theInfo,(inStr(theInfo, uCase("MSIE")) + 4), 6) 
        tmp1 = left(tmp1, inStr(tmp1, ";") - 1) 
        strType = strType & tmp1 & "]" 
        s = "Internet Explorer"  
        s = strType & s 
    end if 
    if inStr(theInfo, uCase("msn")) > 0 then s = "MSN" 
    if inStr(theInfo, uCase("aol")) > 0 then s = "AOL" 
    if inStr(theInfo, uCase("webtv")) > 0 then s = "WebTV" 
    if inStr(theInfo, uCase("myie2")) > 0 then s = "MyIE2" 
    if inStr(theInfo, uCase("maxthon")) > 0 then s = "Maxthon(���������)" 
    if inStr(theInfo, uCase("gosurf")) > 0 then s = "GoSurf(���˸��������)" 
    if inStr(theInfo, uCase("netcaptor")) > 0 then s = "NetCaptor" 
    if inStr(theInfo, uCase("sleipnir")) > 0 then s = "Sleipnir" 
    if inStr(theInfo, uCase("avant browser")) > 0 then s = "AvantBrowser" 
    if inStr(theInfo, uCase("greenbrowser")) > 0 then s = "GreenBrowser" 
    if inStr(theInfo, uCase("slimbrowser")) > 0 then s = "SlimBrowser" 
    if inStr(theInfo, uCase("360SE")) > 0 then s = s & "-360SE(360��ȫ�����)" 
    if inStr(theInfo, uCase("QQDownload")) > 0 then s = s & "-QQDownload(QQ������)" 
    if inStr(theInfo, uCase("TheWorld")) > 0 then s = s & "-TheWorld(����֮�������)" 
    if inStr(theInfo, uCase("icafe8")) > 0 then s = s & "-icafe8(��ά��ʦ���ɹ�����)" 
    if inStr(theInfo, uCase("TencentTraveler")) > 0 then s = s & "-TencentTraveler(��ѶTT�����)" 
    if inStr(theInfo, uCase("baiduie8")) > 0 then s = s & "-baiduie8(�ٶ�IE8.0)" 
    if inStr(theInfo, uCase("iCafeMedia")) > 0 then s = s & "-iCafeMedia(������ý���Ʋ��)" 
    if inStr(theInfo, uCase("DigExt")) > 0 then s = s & "-DigExt(IE5�����ѻ��Ķ�ģʽ������)" 
    if inStr(theInfo, uCase("baiduds")) > 0 then s = s & "-baiduds(�ٶ�Ӳ������)" 
    if inStr(theInfo, uCase("CNCDialer")) > 0 then s = s & "-CNCDialer(���ز���)" 
    if inStr(theInfo, uCase("NOKIAN85")) > 0 then s = s & "-NOKIAN85(ŵ�����ֻ�)" 
    if inStr(theInfo, uCase("SPV_C600")) > 0 then s = s & "-SPV_C600(���մ�C600)" 
    if inStr(theInfo, uCase("Smartphone")) > 0 then s = s & "-Smartphone(Windows Mobile for Smartphone Edition ����ϵͳ�������ֻ�)" 
    getBrType = s 
end function 
%>  


