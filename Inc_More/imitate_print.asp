<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-13
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
'��ӡ��վ��Ϣ
 

Function getImitate_print(id,filePath,jsList,cssList,imgList,swfList,otherList,urlList,errInfoC)
	dim c
	
	jsList=deleteRepeatContent(jsList,vbcrlf)		'ɾ��JS�б��ظ�
	cssList=deleteRepeatContent(cssList,vbcrlf)		'ɾ��CSS�б��ظ�
	imgList=deleteRepeatContent(imgList,vbcrlf)		'ɾ��IMG�б��ظ�
	swfList=deleteRepeatContent(swfList,vbcrlf)		'ɾ��SWF�б��ظ�
	otherList=deleteRepeatContent(otherList,vbcrlf)		'ɾ�������б��ظ�
	urlList=deleteRepeatContent(urlList,vbcrlf)		'ɾ��A��ַ�б��ظ� 
	errInfoC=deleteRepeatContent(errInfoC,vbcrlf)		'ɾ��������Ϣ�б��ظ� 
	
	
    c=c & "<div onClick=""showHide('main"& id &"')"" class=""laout_title noselect"">�ļ�"& filePath &"<span id=""main"& id &"_ico"">-</span></div>" & vbCrlf 
    c=c & "<div id=""main"& id &""" class=""laout_content"">" & vbCrlf
	if jsList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &"')"">JS�ļ��б�("& getArrayCount(jsList,vbcrlf) &")<span id=""js_content"& id &"_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &"""><pre>"& jsList &"</pre></div>" & vbCrlf
    end if
	
	if cssList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".2')"">CSS�ļ��б�"& getArrayCount(cssList,vbcrlf) &"<span id=""js_content"& id &".2_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".2""><pre>"& cssList &"</div>" & vbCrlf
    end if
	
	if imgList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".3')"">Images�ļ��б�"& getArrayCount(imgList,vbcrlf) &"<span id=""js_content"& id &".3_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".3""><pre>"& imgList &"</pre></div>" & vbCrlf
    end if
	
	if swfList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".a')"">SWF�ļ��б�"& getArrayCount(swfList,vbcrlf) &"<span id=""js_content"& id &".a_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".a""><pre>"& swfList &"</pre></div>" & vbCrlf
    end if
	
    if otherList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".b')"">�����ļ��б�"& getArrayCount(otherList,vbcrlf) &"<span id=""js_content"& id &".b_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".b""><pre>"& otherList &"</pre></div>" & vbCrlf
    end if
	
	if urlList<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".g')"">A����URL�б�"& getArrayCount(urlList,vbcrlf) &"<span id=""js_content"& id &".g_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".g""><pre>"& urlList &"</pre></div>" & vbCrlf
	end if
	
	if errInfoC<>"" then
    c=c & "    <div class=""content_title"" onclick=""showHide('js_content"& id &".h')""><font color=blue>������Ϣ�б�"& getArrayCount(errInfoC,vbcrlf) &"</font><span id=""js_content"& id &".h_ico"">-</span></div>" & vbCrlf
    c=c & "    <div class=""content_content"" id=""js_content"& id &".h""><pre>"& errInfoC &"</pre></div>" & vbCrlf
	end if
	
	
	
    c=c & "</div>" & vbCrlf
	getImitate_print=c
end function

Function imitate_config(content)
    Dim splStr, i, s, c
    c=c & "<!DOCTYPE html>" & vbCrlf
    c=c & "<html>" & vbCrlf
    c=c & "<head>" & vbCrlf
    c=c & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbCrlf
    c=c & "<title>��ӡ������Ϣ</title>" & vbCrlf
    c=c & "<style>" & vbCrlf
    c=c & ".header_title{" & vbCrlf
    c=c & "    color: #003300;" & vbCrlf
    c=c & "    line-height: 69px;" & vbCrlf
    c=c & "    font-size: 34px;" & vbCrlf
    c=c & "    font-weight: bold;" & vbCrlf
    c=c & "    text-align: center;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".laout_title{" & vbCrlf
    c=c & "	font-size:14pxpx;" & vbCrlf
    c=c & "	line-height:32px;" & vbCrlf
    c=c & "	font-weight:bold;" & vbCrlf
    c=c & "	border:1px solid #999900; " & vbCrlf
    c=c & "	padding:0px 0 0 10px;" & vbCrlf
    c=c & "	background-color:#333333;" & vbCrlf
    c=c & "	font-weight:bold;" & vbCrlf
    c=c & "	position:relative;" & vbCrlf
    c=c & "	color:#fff;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".laout_title span{" & vbCrlf
    c=c & "	position:absolute;" & vbCrlf
    c=c & "	top:2px;" & vbCrlf
    c=c & "	right:12px;" & vbCrlf
    c=c & "	font-size:18px;" & vbCrlf
    c=c & "	line-height:32px;" & vbCrlf
    c=c & "	color:red;" & vbCrlf
    c=c & "	font-weight:bold;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".laout_content{" & vbCrlf
    c=c & "	font-size:14pxpx;" & vbCrlf
    c=c & "	line-height:22px;" & vbCrlf
    c=c & "	border:1px solid #999900;" & vbCrlf
    c=c & "	border-top:0px;" & vbCrlf
    c=c & "	margin:0 0 10px 0;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".content_title{" & vbCrlf
    c=c & "	padding:6px;" & vbCrlf
    c=c & "	border-bottom:1px solid #000;" & vbCrlf
    c=c & "	position:relative;" & vbCrlf
    c=c & "	background-color:#999999;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".content_title span{" & vbCrlf
    c=c & "	position:absolute;" & vbCrlf
    c=c & "	top:2px;" & vbCrlf
    c=c & "	right:12px;" & vbCrlf
    c=c & "	font-size:18px;" & vbCrlf
    c=c & "	line-height:32px;" & vbCrlf
    c=c & "	color:#000000;" & vbCrlf
    c=c & "	font-weight:bold;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".content_content{" & vbCrlf
    c=c & "	padding:6px; " & vbCrlf
    c=c & "	line-height:22px;" & vbCrlf
    c=c & "	color:#000000;" & vbCrlf
    c=c & "	border-bottom:1px solid #000;" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & ".noselect{-moz-user-select:none;-webkit-user-select:none;-ms-user-select:none;-khtml-user-select:none;user-select:none}" & vbCrlf
    c=c & "</style>" & vbCrlf
    c=c & "</head>" & vbCrlf
    c=c & "<body>" & vbCrlf
    c=c & "<div class=""header_title"">��ӡ������Ϣ</div>" & vbCrlf
    c=c & "    " & vbCrlf
   
    c=c & content & vbcrlf
   
    c=c & "<div style='height:600px;'></div>" & vbcrlf
    c=c & "<script language=""javascript"">" & vbCrlf
    c=c & "//��ʾ����DIV" & vbCrlf
    c=c & "function showHide(ID){ " & vbCrlf
    c=c & "	//Ϊ������ʾ" & vbCrlf
    c=c & "	if(document.all[ID].style.display == ""block"" || document.all[ID].style.display == """"){" & vbCrlf
    c=c & "		document.all[ID].style.display = ""none""" & vbCrlf
    c=c & "		document.all[ID+""_ico""].innerHTML=""+""" & vbCrlf
    c=c & "	}else{" & vbCrlf
    c=c & "		document.all[ID].style.display = ""block""" & vbCrlf
    c=c & "		document.all[ID+""_ico""].innerHTML=""-""" & vbCrlf
    c=c & "	}	" & vbCrlf
    c=c & "	" & vbCrlf
    c=c & "}" & vbCrlf
    c=c & "</script>" & vbCrlf
    c=c & "</body>" & vbCrlf
    c=c & "</html>" & vbCrlf

    imitate_config=c
End Function 


 
'�÷� c=  readBinary("1.jpg",0)& "|" & imageAddMyInfo("1") call decSaveBinary("1.1.jpg",c,0):response.End()
'ͼƬ�����ҵ���Ϣ �������վ���� ����վsoeasy
function imageAddMyInfo(sType)
	dim c
	if sType="1" then
	c="80|75|3|4|10|0|0|0|0|0|188|173|76|73|33|182|146|130|91|0|0|0|91|0|0|0|1|0|0|0|97|212|198|203|239|183|194|213|190|214|250|202|214|32|32|183|194|213|190|115|111|101|97|115|121|13|10|215|247|213|223|163|186|212|198|203|239|13|10|81|81|163|186|51|49|51|56|48|49|49|50|48|13|10|200|186|163|186|51|53|57|49|53|49|48|48|13|10|185|217|205|248|163|186|119|119|119|46|115|104|97|114|101|109|98|119|101|98|46|99|111|109|80|75|1|2|63|0|10|0|0|0|0|0|188|173|76|73|33|182|146|130|91|0|0|0|91|0|0|0|1|0|36|0|0|0|0|0|0|0|32|0|0|0|0|0|0|0|97|10|0|32|0|0|0|0|0|1|0|24|0|39|224|8|244|142|36|210|1|14|220|240|27|191|28|210|1|14|220|240|27|191|28|210|1|80|75|5|6|0|0|0|0|1|0|1|0|83|0|0|0|122|0|0|0|0|0"
	'������
	else
	c= "80|75|3|4|10|0|0|0|0|0|177|173|76|73|33|182|146|130|91|0|0|0|91|0|0|0|28|0|0|0|212|198|203|239|183|194|213|190|214|250|202|214|32|32|183|194|213|190|115|111|101|97|115|121|46|116|120|116|212|198|203|239|183|194|213|190|214|250|202|214|32|32|183|194|213|190|115|111|101|97|115|121|13|10|215|247|213|223|163|186|212|198|203|239|13|10|81|81|163|186|51|49|51|56|48|49|49|50|48|13|10|200|186|163|186|51|53|57|49|53|49|48|48|13|10|185|217|205|248|163|186|119|119|119|46|115|104|97|114|101|109|98|119|101|98|46|99|111|109|80|75|1|2|63|0|10|0|0|0|0|0|177|173|76|73|33|182|146|130|91|0|0|0|91|0|0|0|28|0|36|0|0|0|0|0|0|0|32|0|0|0|0|0|0|0|212|198|203|239|183|194|213|190|214|250|202|214|32|32|183|194|213|190|115|111|101|97|115|121|46|116|120|116|10|0|32|0|0|0|0|0|1|0|24|0|146|157|32|231|142|36|210|1|143|73|59|227|188|28|210|1|255|161|245|230|187|28|210|1|80|75|5|6|0|0|0|0|1|0|1|0|110|0|0|0|149|0|0|0|0|0"
	end if
	imageAddMyInfo=c
end function


'================================  Ϊ��վ�õ���������ʱ����

'����ͼƬ�б�
function handleDownUrlImageList(content,downFileTypeList,isUTF8Down)
	dim splUrl,splxx,url,urlList,fileName,fileType,c,s
	if content="" then
		content=getftext("/newweb\hkgtsj_99.com/��ַ�б�.txt")
	end if	
	if downFileTypeList="" then
		downFileTypeList="jpg|gif|png"
	end if
	downFileTypeList="|"& lcase(trim(downFileTypeList)) &"|"
 
	splUrl=split(content,vbcrlf)
	for each s in splUrl	
		splxx=split(s & "��|��","��|��")
		url=splxx(0)
		'20180921
		if isUTF8Down=true then
			url=toUTFChar(url)			'תutf������ַ
		end if
		if url<>"" and instr(vbcrlf & urlList & vbcrlf ,vbcrlf & url & vbcrlf)=false then
			fileName=getFileAttr(url,"2")
			fileType=lcase(getFileAttr(url,"4"))
			if instr(downFileTypeList,"|"& fileType &"|")>0 then		
			
				urlList=urlList & url & vbcrlf
				c=c & url & vbcrlf			'& "��|��" & fileName &
			end if
		end if
	
	next
	handleDownUrlImageList=c
end function
%>