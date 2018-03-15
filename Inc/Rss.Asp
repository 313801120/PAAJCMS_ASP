<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-15
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%@ codePage = 65001%>  
<!--#Include file = "Config.Asp"-->  
<% 
''时间转换函数
function return_RFC822_Date(byVal myDate, byVal timeZone)
    dim myDay, myDays, myMonth, myYear 
    dim myHours, myMinutes, mySeconds 
    myDate = CDate(myDate) 
    myDay = enWeekDayName(myDate) 
    myDays = right("00" & day(myDate), 2) 
    myMonth = enMonthName(myDate) 
    myYear = year(myDate) 
    myHours = right("00" & hour(myDate), 2) 
    myMinutes = right("00" & minute(myDate), 2) 
    mySeconds = right("00" & second(myDate), 2) 
    return_RFC822_Date = myDay & ", " & _ 
    myDays & " " & _ 
    myMonth & " " & _ 
    myYear & " " & _ 
    myHours & ":" & _ 
    myMinutes & ":" & _ 
    mySeconds & " " & _ 
     " " & timeZone 
end function 
''星期转换函数
function enWeekDayName(inputDate)
    dim result 
    select case weekday(inputDate, 1)
        case 1 : result = "Sun"
        case 2 : result = "Mon"
        case 3 : result = "Tue"
        case 4 : result = "Wed"
        case 5 : result = "Thu"
        case 6 : result = "Fri"
        case 7 : result = "Sat"
    end select
    enWeekDayName = result 
end function 
''月份转换函数
function enMonthName(inputDate)
    dim result 
    select case month(inputDate)
        case 1 : result = "Jan"
        case 2 : result = "Feb"
        case 3 : result = "Mar"
        case 4 : result = "Apr"
        case 5 : result = "May"
        case 6 : result = "Jun"
        case 7 : result = "Jul"
        case 8 : result = "Aug"
        case 9 : result = "Sep"
        case 10 : result = "Oct"
        case 11 : result = "Nov"
        case 12 : result = "Dec"
    end select
    enMonthName = result 
end function 
dim sql, XMLContent, rsc, web_Name, web_Url, categoryName, url, toPath 
web_Name = "南京麦思德餐饮有限公司" 
web_Url = host() 
categoryName = "餐饮培训" 
call OpenConn() 
''获取当前域名,其中rss.asp为你的文件名
sql = "Select Top 50 * From [Product] Where OnHtml=True Order By [UpDateTime] Desc" 
''根据自己实际修改
set rs = createObject("Adodb.RecordSet")
    rs.open sql, conn, 1, 1 
    'Response.ContentType = "text/xml"
    XMLContent = XMLContent & "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf 
    XMLContent = XMLContent & "<?xml-stylesheet type=""text/xsl"" href=""/rss/rss.xslt"" version=""1.0"" ?>" & vbCrLf 
    XMLContent = XMLContent & "<rss version=""2.0"">" & vbCrLf 
    XMLContent = XMLContent & "<channel>" & vbCrLf 
    XMLContent = XMLContent & "<title>" & web_Name & " 的最新动态</title>" & vbCrLf 
    XMLContent = XMLContent & "<link>" & web_Url & "</link>" & vbCrLf 
    XMLContent = XMLContent & "<language>zh-cn</language>" & vbCrLf 
    XMLContent = XMLContent & "<copyright>RSS Feed By www.hitux.com</copyright>" & vbCrLf 
    while not rs.eof
        XMLContent = XMLContent & "<item>" & vbCrLf 
        XMLContent = XMLContent & "<title><![CDATA[" & rs("title") & "]]></title>" & vbCrLf 
        url = web_Url & getHandleUrl(rs("FileName"))                                    '处理URL
        XMLContent = XMLContent & "<link>" & url & "</link>" & vbCrLf 
        XMLContent = XMLContent & "<description><![CDATA[" & rs("WebDescription") & "]]></description>" & vbCrLf 
        XMLContent = XMLContent & "<pubDate>" & return_RFC822_Date(rs("UpDateTime"), "08:00") & "</pubDate>" & vbCrLf 
        XMLContent = XMLContent & "<author>" & web_Name & "</author>" & vbCrLf 
        XMLContent = XMLContent & "<category>" & categoryName & "</category>" & vbCrLf 
        XMLContent = XMLContent & "</item>" & vbCrLf 
        rs.moveNext 
    wend 
    XMLContent = XMLContent & "</channel>" & vbCrLf 
    XMLContent = XMLContent & "</rss>" & vbCrLf 
    rs.close 
set rs = nothing 
call createFolder("../Rss") 
toPath = handlePath("../Rss/Feed.Xml") 
call WriteToFile(toPath, XMLContent, "utf-8") 
url = "<a " & AHref(web_Url & "/Rss/Feed.Xml", "点击打开", "_blank") & ">" & web_Url & "Rss/Feed.Xml" & "</a>" 

call echo("Create Rss/Feed.Xml OK", url) 
'Response.Write XMLContent
%>  



