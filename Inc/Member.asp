<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-02-27
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%  
 

'调用function2文件函数
function callMember()
    dim sType 
    sType = request("stype") 
    if sType = "login" then
        call member_Login()                               '会员登录
    elseif sType = "outlogin" then
        call member_OutLogin()                            '会员退出
    elseif sType = "addFollow" then
        call member_addFollow()                           '添加关注
    elseif sType = "delFollow" then
        call member_delFollow()                           '删除关注
    else
        call eerr("member页里没有动作", request("stype")) 
    end if 
end function


'会员登录
sub member_Login() 
	dim user,pass,pwd
	user=replace(request("user"),"'","")
	pass=replace(request("pass"),"'","")
	pwd=myMD5(pass)
	'【@是jsp显示@】try{
    rs.Open "select * from " & db_PREFIX & "member where username='" & user & "' and pwd='" & pwd & "'", conn, 1, 1 
    If Not rs.EOF Then 
		if dateDiff("d", rs("expireDateTime"), now())>=1 then
			conn.execute("delete * from  " & db_PREFIX & "member where id=" & rs("id") & " or parentid=" & rs("id"))
			
			call rwend("会员到期，当前会员及下级会员将被删除")
		else
			call setsession("member_user",user)
			call setsession("member_id",rs("id"))
			call setsession("member_expiredatetime",rs("expiredatetime"))
			call setsession("member_followIdList",rs("followIdList"))		'关注文章ID列表
			call rwend("ok")
		end if	 
	else
		call rwend("账号密码错误")
	end if:rs.close
	'【@是jsp显示@】}catch(Exception e){} 
end sub     

'会员退出
sub member_OutLogin()                         
	call setsession("member_user","")
	call setsession("member_id","")
	call setsession("member_expiredatetime","")
	call setsession("followIdList","")
	call rr("?")

end sub
'关注
function XY_thisArticleFollow(action)
	dim followTrue,followFalse,s,defaultStr
    defaultStr = getDefaultValue(action)                                            '获得默认内容
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

'添加关注
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
'删除关注
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