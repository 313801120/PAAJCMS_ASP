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
<!--#Include File = "Config.Asp"-->  
<%  
'ѡ�����б�
Select Case Request("act") 
	Case "SaveFile" : SaveFile( Request("FilePath"))			'�����ļ�
	Case "GetFunctionCode" : GetFunctionCode(Request("Lable"))	'��ú�������
	Case "SaveQQDialogue" : Call SaveQQDialogue()				'������QQ�Ի���Ϣ
	Case "UserLogin" : Call UserLogin()						 	'��Ա��¼
	Case "UserReg" : Call UserReg()								'��Աע��
	Case "MobileArticleList" : Call MobileArticleList()			'�ֻ��ļ��б�
	Case "ShowArticleNumb" : Call ShowArticleNumb()				'��ʾ��������� �ۼ�
	Case "SaveAJAXEdit" : Call SaveAJAXEdit()					'Ajax��������
	Case "SaveAJAXTEXTEdit" : Call SaveAJAXTEXTEdit()					'Ajax����TXT�ļ�����
	
	Case "DeleteAction" : Call DeleteAction()					'Ajaxɾ������
	
	Case "showArticleList" : Call showArticleList()					'Ajaxɾ������
End Select

'��ʾ�����б� http://127.0.0.1/inc/ajax.asp?act=showArticleList&did=�Ķ�&numb=4
function showArticleList()
	call openconn()
	dim did,sql,i,nNumb
	did=Rq("did")
	nNumb=HandleNumber(Rq("Numb")) 	
	if nNumb<>"" then
		sql = "select top "& nNumb &" * from [product]"
	else	
		sql = "select * from [product]"
	end if
	if did<>"" then
		sql=sql & " where BigClassName='"& did &"'"
	end if
	rs.open sql,conn,1,1
	Dim json : Set json = New Cls_AB_JSON        '����JSON��
	'call rw(json.toJSON(dataArr))
	call rw(json.toJSON(rs))
	'call rw(json.fieldToJson(rs,"|title|"))
	rs.close 
end function

function showArticleList2()
	call openconn()
	dim did,sql,i,nNumb
	did=Rq("did")
	nNumb=HandleNumber(Rq("Numb")) 	
	if nNumb<>"" then
		sql = "select top "& nNumb &" * from [product]"
	else	
		sql = "select * from [product]"
	end if
	if did<>"" then
		sql=sql & " where BigClassName='"& did &"'"
	end if
	rs.open sql,conn,1,1
	Dim json : Set json = New Cls_AB_JSON        '����JSON��
	'call rw(json.toJSON(dataArr))
	call rw(json.toJSON(rs))
	rs.close
end function

'Ajaxɾ������
Function DeleteAction() 
	Dim C	
	C=C & "MDBPath=" & Rq("MDBPath") & vbCrlf
	C=C & "TableName=" & Rq("TableName") & vbCrlf
	C=C & "FieldName=" & Rq("FieldName") & vbCrlf 
	C=C & "ID=" & Rq("ID") & vbCrlf
	'Call Rw("<pre>"& C &"</pre>"):Doevents : Response.end
	MdbPath=Request("MDBPath")
	Call OpenConn()
	Dim Sql
	Sql = "Select * From ["& Request("TableName") &"] Where Id=" & Request("Id")
	'Call Rw(MDBPath & "<hr>" & Sql)
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof Then
		Rs.Delete
		Rs.Update
		Call Rw("ɾ��ID��"& Request("Id") &"���ɹ�")
	Else		
		Call Rw("ID��"& Request("Id") &"�������ڣ�ɾ��ʧ��")
	End If:Rs.Close
End Function
'Ajax��������
Function SaveAJAXEdit() 
	Dim C
	
	C=C & "MDBPath=" & Rq("MDBPath") & "("& IsNul(Rq("MDBPath")) &")" & vbCrlf
	C=C & "MDBPath=" & handlePath(Rq("MDBPath")) & vbCrlf
	C=C & "TableName=" & Rq("TableName") & vbCrlf
	C=C & "FieldName=" & Rq("FieldName") & vbCrlf
	C=C & "Content=" & Rq("Content") & vbCrlf
	C=C & "ID=" & Rq("ID") & vbCrlf
	Call Rw("<pre>"& C &"</pre>"):Doevents
	'MdbPath=Request("MDBPath")						' �������Ҫ�ˣ���Ϊ��Conn.Asp���Ѿ��д�����δ�����20150311
	Call OpenConn()
	Dim Sql
	Sql = "Select * From ["& Request("TableName") &"] Where Id=" & Request("Id")
	'Call Rw(MDBPath & "<hr>" & Sql)
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof Then
		Rs(Rq("FieldName")) = Request("Content")
		Rs.Update
	End If:Rs.Close
End Function
'Ajax�����ļ�����
Function SaveAJAXTEXTEdit() 
	Dim C	
	C=C & "FilePath=" & Rq("FilePath") & vbCrlf 
	
	
	C=C & "content=" & Request.Form("content") & vbCrlf 
	C=C & "count=" & Request.Form("count") & vbCrlf 
	Call Rw("<pre>"& C &"</pre>"):Doevents
	Call CreateFile(Request("FilePath"),Request("content"))
End Function

'��ʾ��������� �ۼ�
Sub ShowArticleNumb()
	Dim Id
	Id = Request("Id")
	Call OpenConn()
	Rs.Open"Select * From [Product] Where Id=" & Id,Conn,1,3
	If Not Rs.Eof Then
		Rs("Hits") = Rs("Hits")+1
		Rs.Update
		Call Rw(Rs("Hits"))
	End If:Rs.Close
End Sub

'�ֻ��ļ��б�
Sub MobileArticleList()
	'Call Rw(GetFText("/2.xml"))
	Dim C,Sql,nCount,I
	Call OpenConn()
	If Session("ShowMobileArticleNumb") = "" Then Session("ShowMobileArticleNumb")=12
	nCount = 0
    C=C & "<?xml version=""1.0"" encoding=""gb2312"" ?> " & vbCrlf
    C=C & "<clists> " & vbCrlf
	Sql = "Select * From [Product] Where BigClassName='��Ϣ����' Order By Id Desc"
	Rs.Open Sql,Conn,1,1
	For I =1 To Rs.RecordCount
		If nCount>10 Then Exit For
		If I > Session("ShowMobileArticleNumb") then
			nCount=nCount+1
		    C=C & "<clist id="""& Rs("ID") &""" title="""& Rs("Title") &""" image=""0"" time="""& Rs("AddDateTime") &"""/>" & vbCrlf
		End If
	Rs.MoveNext:Next:Rs.Close
	Session("ShowMobileArticleNumb")=Session("ShowMobileArticleNumb")+nCount 
    C=C & "</clists>" & vbCrlf
	Call Rw(C)
End Sub
'��Աע��
Sub UserReg()
	Dim UserName,PassWord,Email
	UserName=Replace(Request("UserName"),"'","")
	PassWord=Replace(Request("PassWord"),"'","")
	If Request("P")="" Then
		PassWord = MD5(PassWord,2) & MD5(PassWord,4)
	End If
	Email=Replace(Request("Email"),"'","")	 
	Response.Cookies("UserName")=UserName
	Response.Cookies("UserName").Expires=Date+10	
	Call OpenConn()
	Rs.Open"Select * From [Member] Where UserName='"& UserName &"'",Conn,1,3
	If Not Rs.Eof Then
		Call Rw("�û����ѱ�ע��")
	Else
		Rs.Close
		Rs.Open"Select * From [Member] Where Email='"& Email &"'",Conn,1,3
		If Not Rs.Eof Then
			Call Rw("�����ѱ�ע��")
		Else
			Rs.AddNew
			Rs("UserName")=UserName
			Rs("Pwd")=PassWord
			Rs("Email")=Email
			Rs("MyMoney") = 100000
			Rs.Update
			Response.Cookies("PassWord")=PassWord
			Response.Cookies("PassWord").Expires=date+10
			Session("Email")=Rs("Email")
			Session("MyMoney")=CLng(Rs("MyMoney"))
			Call Rw("OK")
		End If
		Response.End()
	End If:Rs.Close
End Sub
'��Ա��¼
Sub UserLogin()
	Dim UserName,PassWord,autologin
	UserName=Replace(Request("UserName"),"'","")
	PassWord=Replace(Request("PassWord"),"'","")
	If Request("P")="" Then
		PassWord = MD5(PassWord,2) & MD5(PassWord,4)
	End If
	autologin=Replace(Request("autologin"),"'","")
	 
	Response.Cookies("UserName")=UserName
	Response.Cookies("UserName").Expires=date+10
	
	Call OpenConn()
'	Call Rw("UserName=" & UserName & "|PassWord=" & PassWord & "|autologin=" & autologin)
	Rs.Open"Select * From [Member] Where UserName='"& UserName &"' Or Email='"& UserName &"'",Conn,1,3
	If Rs.Eof Then
		Call Rw("�û���������")
	Else
		If Rs("Pwd")<>PassWord Then
			Call Rw("���벻��ȷ")
		Else
			Response.Cookies("PassWord")=PassWord
			Response.Cookies("PassWord").Expires=date+10
			Session("Email")=Rs("Email")
			Session("MyMoney")=CLng(Rs("MyMoney"))
			Call Rw("OK")
		End If
	End If:Rs.Close
End Sub

'������QQ�Ի���Ϣ
Sub SaveQQDialogue()
	Dim QQ,ClickStyle
	QQ = Request("QQ")
	ClickStyle = Request("ClickStyle")
	Call CreateAddFile("/QQ/QQ.txt","����QQ=" & QQ & "����ʽ=" & ClickStyle & "��ʱ��=" & Format_Time(Now(),8))
	Call Rw("����QQ�Ի��ɹ�")
End Sub 

'��ú�������(��ʾ������Ϣ����)
Sub GetFunctionCode(Lable)
	Dim Content
	Content = GetFText("FSO.Asp")
	Content = StrCut(Content,"Function " & Lable,"End Function",1)
	
	Call Rw(ShowAspCodeColor(Content))
End Sub
%>