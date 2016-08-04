<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-04
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'doc文件操作

'读Doc文件内容 20141218  操作docx 要自己安装一个组件，麻烦
'例：Call Rw( GetDocText("1.docx"))
function getDocText(byVal fileName)
    on error resume next 
    dim fso, fText, openFile 
    'GetFileText = ""   '它默认返回的就是空， 这个是多此一举 (2013,9,30)
    call handlePath(fileName)                '获得完整路径
    dim doc 
    set doc = createObject("AxWsc.Word.1.00")
        getDocText = doc.getDocContent(fileName) 
        if err then doError err.description, "GetDocText 读取文件内容 函数出错，FileName=" & fileName 
end function
%> 

