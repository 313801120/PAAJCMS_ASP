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
'颜色20152027

'获得随机颜色20152027
Function getRandColor()
    Dim splStr  
    splStr = Split("#990000,#999900,#333366,#663300,#669966,#FF6600,#CC33CC,#993366,#FF0099,#669900,#336699,#99CCCC,#CC3366,#FF9900,#9933FF,#669900,#6699FF,#333366,#99CC66,#996600,#000033,#003300,#330000,#660000,#000099,#330099,#6600FF,#990000,#CC0000,#FF0000,#990066,#FF0066,#CC00CC", ",") 
    getRandColor = splStr(phpRand(0, UBound(splStr))) 
End Function 
%> 
