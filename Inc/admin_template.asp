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
<!--#Include File = "../Inc/Config.asp"-->
<!--#Include File = "function.asp"--> 
<!DOCTYPE html> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /> 
<title>ģ���ļ�����</title> 
</head> 
<body> 
<style type="text/css"> 
<!-- 
body { 
    margin-left: 0px; 
    margin-top: 0px; 
    margin-right: 0px; 
    margin-bottom: 0px; 
} 
a:link,a:visited,a:active { 
    color: #000000; 
    text-decoration: none; 
} 
a:hover { 
    color: #666666; 
    text-decoration: none; 
} 
.tableline{ 
    border: 1px solid #999999; 
} 
body,td,th { 
    font-size: 12px; 
} 
a { 
    font-size: 12px; 
} 
--> 
</style> 
<script language="javascript"> 
function checkDel() 
{ 
    if(confirm("ȷ��Ҫɾ����ɾ���󽫲��ɻָ���")) 
    return true; 
    else 
    return false; 
} 
</script> 
<% 

If Session("adminusername") = "" Then
    Call eerr("��ʾ", "δ��¼�����ȵ�¼") 
End If 
 
call openconn()
Select Case Request("act")
    Case "templateFileList" : displayTemplateDirDialog(Request("dir")) : templateFileList(Request("dir"))'ģ���б�
    Case "delTemplateFile" : Call delTemplateFile(Request("dir"), Request("fileName")) : displayTemplateDirDialog(Request("dir")) : templateFileList(Request("dir"))		'ɾ��ģ���ļ�
    Case "addEditFile" : displayTemplateDirDialog(Request("dir")) : Call addEditFile(Request("dir"), Request("fileName"))'��ʾ�����޸��ļ�
    Case Else : displayTemplateDirDialog(Request("dir"))                                        '��ʾģ��Ŀ¼���
End Select

'ģ���ļ��б�
Sub templateFileList(dir)
    Dim content, splStr, fileName, s,fileType ,folderName,filePath
	
	if  Session("adminusername") = "PAAJCMS"  then 
		content = getDirFolderNameList(dir,"")
    	splStr = Split(content, vbCrLf) 	
		For Each folderName In splStr
			s="<a href='?act=templateFileList&dir="& dir & "/" & folderName &"'>"& folderName &"</a>" 
			Call echo("<img src='Images/file/folder.gif'>",s)
		next
		content = getDirFileNameList(dir,"")
	else
	    content = getDirHtmlNameList(dir)
	end if
    splStr = Split(content, vbCrLf) 
    For Each fileName In splStr
		if fileName<>"" then
			fileType=lcase(getFileAttr(fileName,4))
			filePath=dir & "/" & filename
			if instr("|asa|asp|aspx|bat|bmp|cfm|cmd|com|css|db|default|dll|doc|exe|fla|folder|gif|h|htm|html|inc|ini|jpg|js|jtbc|log|mdb|mid|mp3|php|png|rar|real|rm|swf|txt|wav|xls|xml|zip|","|"& fileType &"|")=false then
				fileType="default"			
			end if
			 
			s = "<a href=""../aspweb.asp?templatedir=" & escape(dir) & "&templateName=" & fileName & """ target='_blank'>Ԥ��</a> " 
			Call echo("<img src='Images/file/"& fileType &".gif'>" & fileName & "��"& printSpaceValue(getFSize(filePath)) &"��", s & "| <a href='?act=addEditFile&dir=" & dir & "&fileName=" & fileName & "'>�޸�</a> | <a href='?act=delTemplateFile&dir=" & Request("dir") & "&fileName=" & fileName & "' onclick='return checkDel()'>ɾ��</a>") 
		end if
    Next 
	
	
	
End Sub
 
'ɾ��ģ���ļ�
Sub delTemplateFile(dir, fileName)
    Dim filePath 
	
	call handlePower("ɾ��ģ���ļ�")						'����Ȩ�޴���
	
    filePath = dir & "/" & fileName 
    Call deleteFile(filePath) 
    Call echo("ɾ���ļ�", filePath) 
End Sub

'��ʾ�����ʽ�б�
function displayPanelList(dir)
	dim content,splstr,s,c
	content=getDirFolderNameList(dir)
	splstr=split(content,vbcrlf)
	c="<select name='selectLeftStyle'>"
	for each s in splstr
		s="<option value=''>"& s &"</option>"
		c=c & s & vbcrlf
	next
	displayPanelList = c & "</select>"	
end function
 

'�����޸��ļ�
Function addEditFile(dir, fileName)
    Dim filePath,promptMsg
	
    If Right(LCase(fileName), 5) <> ".html" and Session("adminusername") <> "PAAJCMS" Then
        fileName = fileName & ".html" 
    End If 
    filePath = dir & "/" & fileName
	
	if checkFile(filePath)=false then
		call handlePower("����ģ���ļ�")						'����Ȩ�޴���
	else
		call handlePower("�޸�ģ���ļ�")						'����Ȩ�޴���	
	end if
	 
    '��������
    If Request("issave") = "true" Then
        Call createfile(filePath, Request("content")) 
		promptMsg="����ɹ�"
    End If 
%> 
<form name="form1" method="post" action="?act=addEditFile&issave=true"> 
  <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline"> 
    <tr> 
      <td height="30">Ŀ¼<% =dir%><br> 
      <input name="dir" type="hidden" id="dir" value="<% =dir%>" /></td> 
    </tr> 
    <tr> 
      <td>�ļ����� 
      <input name="fileName" type="text" id="fileName" value="<% =fileName%>" size="40">&nbsp;<input type="submit" name="button" id="button" value=" ���� " /><%=promptMsg%>
      <br> 
      <textarea name="content"  style="width:99%;height:480px;"id="content"><%call rw(getFText(filePath))%></textarea></td> 
    </tr>  
  </table> 
</form>
<% End Function
'�ļ�������
Function displayTemplateDirDialog(dir)
	dim folderPath
%> 
<form name="form2" method="post" action="?act=templateFileList"> 
  <table width="99%" border="0" cellspacing="0" cellpadding="0" class="tableline"> 
    <tr> 
      <td height="30"><input name="dir" type="text" id="dir" value="<% =dir%>" size="60" /> 
        <input type="submit" name="button2" id="button2" value=" ���� " /><%
		folderPath=dir & "/images/column/"
		if checkFolder(folderPath) then
			call rw("�����ʽ" & displayPanelList(folderPath))
		end if
		folderPath=dir & "/images/nav/"
		if checkFolder(folderPath) then
			call rw("������ʽ" & displayPanelList(folderPath))
		end if
		%></td> 
    </tr> 
  </table> 
</form> 
<% End Function%>
