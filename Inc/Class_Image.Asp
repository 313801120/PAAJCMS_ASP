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


'���ɣ�����ͼƬ�ȱ������ŷ��� ��=213,��=160  ���ź��=100,��=75   100��213�õ�С������λ����������
'document.write("\n<br>��ʾ" + 100/213)
'document.write("\n<br>����" + 160*0.47)


'�÷�
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,""))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"��"))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"��"))
'���ͼƬ����    getFileContentType("/aa.jpg")
 
'psd���У��ȼ�
class GPS
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

    '///�����ǵ��ô���/// ���ͼƬ�ļ���С
    function getImageSize(filespec)
        dim bFlag, p1 
        if checkFile(filespec) = false or getFSize(filespec)=0 then exit function                               'ͼƬ�ļ����� ���˳�

        aso.loadFromFile(filespec) 
        'On Error Resume Next
        dim ret(3) 
        'call echo("filespec",filespec)
        bFlag = aso.read(3) 
        'Call Echo("bFlag1",Hex(binVal(bFlag)))
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