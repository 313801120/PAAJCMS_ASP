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
'新网站函数

'特殊字符替换
Function specialStrReplace(ByVal content)
    content = Replace(content, "\|", "[$特殊字符A]$") 
    content = Replace(content, "\-", "[$特殊字符B]$") 
    content = Replace(content, "\,", "[$特殊字符C]$") 
    content = Replace(content, "\'", "[$特殊字符D]$") 
    content = Replace(content, "\""", "[$特殊字符E]$") 
    specialStrReplace = content 
End Function 
'解特殊字符替换
Function unSpecialStrReplace(ByVal content, startStr)
    content = Replace(content, "[$特殊字符A]$", startStr & "|") 
    content = Replace(content, "[$特殊字符B]$", startStr & "-") 
    content = Replace(content, "[$特殊字符C]$", startStr & ",") 
    content = Replace(content, "[$特殊字符D]$", startStr & "'") 
    content = Replace(content, "[$特殊字符E]$", startStr & """") 
    unSpecialStrReplace = content 
End Function 

'栏目类型处理 首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单
function handleColumnType(columnName)
    dim s 
    if columnName = "首页" then
        s = "home" 
    elseif columnName = "文本" then
        s = "text" 
    elseif columnName = "产品" then
        s = "product" 
    elseif columnName = "新闻" then
        s = "news" 
    elseif columnName = "视频" then
        s = "video" 
    elseif columnName = "下载" then
        s = "download" 
    elseif columnName = "案例" then
        s = "case" 
    elseif columnName = "留言" then
        s = "message" 
    elseif columnName = "反馈" then
        s = "feedback" 
    elseif columnName = "招聘" then
        s = "job" 
    elseif columnName = "订单" then
        s = "order" 
    end if 
    handleColumnType = s 
end function

%> 
