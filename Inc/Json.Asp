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
'json处理

'字典转Json
function dictionaryToJson(obj)
	dim s,c,key
	For Each s In aspD
		key=s'phpno
		s=aspD(key)'phpno
		'call echo(key, s)
		c=c & key & ":"""& JsEncode__(s) &""""
		if c<>"" then
			c=c & ","
		end if
	next
	c="[{"& c &"}]"
	dictionaryToJson=c
end function

%>