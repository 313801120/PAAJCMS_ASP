<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-16
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
class GPS2
 dim aso 
    sub class_Initialize()
        set aso = createObject("ADODB.Stream")
            aso.mode = 3 
            aso.type = 1 
            aso.open 
    end sub
    sub class_Terminate()
        set aso = nothing 
    end sub 
    function bin2Str(bin)
        dim g, str, clow 
		call echo("", ubound(bin))
        for g = 1 to lenB(bin)
            clow = midB(bin, g, 1) 
            if ascB(clow) < 128 then
                str = str & chr(ascB(clow)) 
            else
                g = g + 1 
                if g <= lenB(bin) then str = str & chr(ascW(midB(bin, g, 1) & clow)) 
            end if 
        next 
        bin2Str = str 
    end function 
    function binVal(bin)
        dim ret 
        dim h 
        ret = 0 
        for h = lenB(bin) to 1 step - 1
            ret = ret * 256 + ascB(midB(bin, h, 1)) 
        next 
        binVal = ret 
    end function 
    function binVal2(bin)
        dim ret, g 
        ret = 0 
        for g = 1 to lenB(bin)
            ret = ret * 256 + ascB(midB(bin, g, 1)) 
        next 
        binVal2 = ret 
    end function 
    '///以下是调用代码/// 获得图片文件大小
    function getImageSize(filespec)
        dim bFlag, p1 
        if checkFile(filespec) = false then exit function                               '图片文件不存 则退出

        aso.loadFromFile(filespec) 
        'On Error Resume Next
        dim ret(3) 
        'call echo("filespec",filespec)
        bFlag = aso.read(3) 
        'Call Echo("bFlag1",Hex(binVal(bFlag)))
		call echo(binVal(bFlag), hex(binVal(bFlag))) 
        select case hex(binVal(bFlag))
            case "4E5089"
                aso.read(15) 
                ret(0) = "PNG" 
                ret(1) = binVal2(aso.read(2)) 
                aso.read(2) 
                ret(2) = binVal2(aso.read(2)) 
            case "464947"
                aso.read(3) 
                ret(0) = "GIF" 
                ret(1) = binVal(aso.read(2)) 
                ret(2) = binVal(aso.read(2)) 
            case "FFD8FF"
                do
                do : p1 = binVal(aso.read(1)) : loop while p1 = 255 and not aso.EOS
                if p1 > 191 and p1 < 196 then exit do else aso.read(binval2(aso.read(2)) - 2) 
                do : p1 = binVal(aso.read(1)) : loop while p1 < 255 and not aso.EOS 
                loop while true 
                aso.read(3) 
                ret(0) = "JPG" 
                ret(2) = binval2(aso.read(2)) 
                ret(1) = binval2(aso.read(2)) 
            case else
                if left(bin2Str(bFlag), 2) = "BM" then
                    aso.read(15) 
                    ret(0) = "BMP" 
                    ret(1) = binval(aso.read(4)) 
                    ret(2) = binval(aso.read(4)) 
                else
                    ret(0) = "" 
                end if
        end select
        ret(3) = "width=""" & ret(1) & """ height=""" & ret(2) & """" 
        getImageSize = ret 
    end function 
end class 
 

%>