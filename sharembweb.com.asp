<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-05
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%
function sharembweb_com(str)
	dim c
	c="欢迎访问 sharembweb.com<hr>"
	c=c&"学习交流可联系作者：313801120或群35915100<hr>"
	sharembweb_com=c
end function
'作者信息
Function authorInfo(FileInfo)
    Dim c 
    c = "'************************************************************" & vbCrLf 
    If FileInfo <> "" Then c = c & "'  文件：" & FileInfo & vbCrLf 
    c = c & aspS & "作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)" & vbCrLf 
    c = c & "'版权：源代码公开，各种用途均可免费使用。 " & vbCrlf
    c = c & "'创建：20160111" & vbCrLf 
    c = c & "'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com" & vbCrlf
    c = c & "'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得" & vbCrlf
    c = c & aspS & "*                                    Powered by PAAJCMS " & vbCrLf 
    c = c & "'************************************************************" & vbCrlf
    authorInfo = c 
End Function  
Function authorInfo2()
    Dim c 
    c = "                '''" & vbCrLf 
    c = c & "               (0 0)" & vbCrLf 
    c = c & "   +-----oOO----(_)------------+" & vbCrLf 
    c = c & "   |                           |" & vbCrLf 
    c = c & "   |    让我们一起来体验       |" & vbCrLf 
    c = c & "   |    QQ:313801120           |" & vbCrLf 
    c = c & "   |    sharembweb.com         |" & vbCrLf 
    c = c & "   |                           |" & vbCrLf 
    c = c & "   +------------------oOO------+" & vbCrLf 
    c = c & "              |__|__|" & vbCrLf 
    c = c & "               || ||" & vbCrLf 
    c = c & "              ooO Ooo" & vbCrLf 

    authorInfo2 = c 
End Function 
response.Write(sharembweb_com(""))
response.Write("<pre>" & authorInfo("") & authorInfo2())
%>


