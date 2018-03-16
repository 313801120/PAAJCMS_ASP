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
'会员操作函数 20170527

function getCheckMemberLoginTemplate(templateName)
	'是会员验证 
	if isMemberVerification=true and glb_memberusercheck=1 and request("usertype")<>"show"then
		if getsession("member_id")="" and getCookie("member_id")<>"" then 
			call setsession("member_id",getCookie("member_id"))
		end if
		if getsession("member_id")<>"" then
			call openConn()
			rs.open "Select * from " & db_PREFIX & "Member where id=" & getsession("member_id"), conn, 1,3
			if not rs.EOF then
				call setsession("member_user",rs("username"))
				call setsession("member_id",rs("id"))
				call setCookie("member_user",rs("username"),86400)
				call setCookie("member_id",rs("id"),86400)
			else
				templateName="login.html" 
				call clearMemberLoginSessionCookie()
			end if:rs.close 
		else
			templateName="login.html" 
			call clearMemberLoginSessionCookie()
		end if
	end if
	'call eerr("templateName",templateName)
	getCheckMemberLoginTemplate=templateName
end function
'清空会员登录
function clearMemberLoginSessionCookie()
	call deleteSession("member_user")
	call deleteSession("member_id")
	call removeCookie("member_user")
	call removeCookie("member_id") 
end function
%>