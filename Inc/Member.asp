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
 

'����function2�ļ�����
function callMember()
    dim sType 
    sType = request("stype") 
    if sType = "login" then
        call member_Login()                               '��Ա��¼
    elseif sType = "outlogin" then
        call member_OutLogin()                            '��Ա�˳�
    elseif sType = "addFollow" then
        call member_addFollow()                           '��ӹ�ע
    elseif sType = "delFollow" then
        call member_delFollow()                           'ɾ����ע
    else
        call eerr("memberҳ��û�ж���", request("stype")) 
    end if 
end function


'��Ա��¼
sub member_Login() 
	dim user,pass,pwd
	user=replace(request("user"),"'","")
	pass=replace(request("pass"),"'","")
	pwd=myMD5(pass)
	'��@��jsp��ʾ@��try{
    rs.Open "select * from " & db_PREFIX & "member where username='" & user & "' and pwd='" & pwd & "'", conn, 1, 1 
    If Not rs.EOF Then 
		if dateDiff("d", rs("expireDateTime"), now())>=1 then
			conn.execute("delete * from  " & db_PREFIX & "member where id=" & rs("id") & " or parentid=" & rs("id"))
			
			call rwend("��Ա���ڣ���ǰ��Ա���¼���Ա����ɾ��")
		else
			call setsession("member_user",user)
			call setsession("member_id",rs("id"))
			call setsession("member_expiredatetime",rs("expiredatetime"))
			call setsession("member_followIdList",rs("followIdList"))		'��ע����ID�б�
			call rwend("ok")
		end if	 
	else
		call rwend("�˺��������")
	end if:rs.close
	'��@��jsp��ʾ@��}catch(Exception e){} 
end sub     

'��Ա�˳�
sub member_OutLogin()                         
	call setsession("member_user","")
	call setsession("member_id","")
	call setsession("member_expiredatetime","")
	call setsession("followIdList","")
	call rr("?")

end sub
'��ע
function XY_thisArticleFollow(action)
	dim followTrue,followFalse,s,defaultStr
    defaultStr = getDefaultValue(action)                                            '���Ĭ������
	followTrue = getStrCut(defaultStr, "[list-true]", "[/list-true]", 2) 
	followFalse = getStrCut(defaultStr, "[list-false]", "[/list-false]", 2) 
	'call echo(session("member_followIdList"),glb_id)
	if instr("," & getsession("member_followIdList") &",", ","& glb_id &",")>0 then
		s=followTrue
	else
		s=followFalse
	end if
	XY_thisArticleFollow=s 
end function

'��ӹ�ע
function member_addFollow()
	dim id,sql,splstr,s,c,isAdd
	id=request("id")
	isAdd=true
	splstr=split(getsession("member_followIdList"),",")
	for each s in splstr
		if s=id then
			isAdd=false
		end if
		if c<>"" then
			c=c & ","
		end if
		c=c & s
	next
	if isAdd=true then
		if c<>"" then
			c=c & ","
		end if
		c=c & id
	end if
	call setsession("member_followIdList",c)
	sql="update " & db_PREFIX & "member set followidlist='"& getsession("member_followIdList") &"' where id=" & getsession("member_id")
	'call echo("sql",sql):doevents
	conn.execute(sql)
	call rwend("ok")
end function
'ɾ����ע
function member_delFollow()
	dim id,sql,splstr,s,c,isAdd
	id=request("id")
	isAdd=true
	splstr=split(getsession("member_followIdList"),",")
	for each s in splstr
		if s<>id then			 
			if c<>"" then
				c=c & ","
			end if
			c=c & s
		end if
	next 
	call setsession("member_followIdList",c)
	sql="update " & db_PREFIX & "member set followidlist='"& getsession("member_followIdList") &"' where id=" & getsession("member_id")
	'call echo("sql",sql):doevents
	conn.execute(sql)
	call rwend("ok")
end function
%>