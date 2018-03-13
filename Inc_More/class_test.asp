<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-13
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%

'dim t:  set t=new testClass
'call echo("",t.id)
't.id=333
'call echo("",t.id)

class testClass
	public id 
	 '构造函数 初始化
    Sub Class_Initialize()
		id=3
	end sub
    '析构函数 类终止
    Sub Class_Terminate()
        'HtmlFolder=nothing
        'HtmlFilename=nothing
        'HtmlContent=nothing
        'Urlname=nothing
    End Sub
	
	function getID()
		getID=id
	end function
	sub setID(idStr)
		id=idStr
	end sub
end class 


%>