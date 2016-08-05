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
<% option explicit%> 
<!--#Include file = "Common.Asp"-->  
<!--#Include file = "Function.Asp"-->  
<!--#Include file = "FSO.Asp"-->  
<!--#Include file = "AdodbStream.Asp"-->  
<!--#Include file = "XML.Asp"-->  
<!--#Include file = "Time.Asp"-->  
<!--#Include file = "RegExp.Asp"-->  
<!--#Include file = "Cai.Asp"-->  
<!--#Include file = "StringNumber.Asp"-->  
<!--#Include file = "EncDec.Asp"-->  
<!--#Include file = "Dictionary.Asp"-->  
<!--#Include file = "Access.Asp"-->  
<!--#Include file = "Incpage.Asp"-->  
<!--#Include file = "Check.Asp"-->  
<!--#Include file = "GetWord.Asp"-->  
<!--#Include file = "Url.Asp"-->  
<!--#Include file = "Conn.Asp"-->  
<!--#Include file = "Fun_Table.Asp"-->  
<!--#Include file = "FunHTML.Asp"-->   
<!--#Include file = "Html.Asp"-->   
<!--#Include file = "Img.Asp"-->  
<!--#Include file = "PinYin.Asp"--> 
<!--#Include file = "Md5.Asp"-->   
<!--#Include file = "IE.Asp"-->   
<!--#Include file = "ASPPHPAccess.Asp"-->     
<!--#Include file = "2014_Array.Asp"-->  
<!--#Include file = "2014_Js.Asp"-->  
<!--#Include file = "2014_Css.Asp"-->  
<!--#Include file = "2014_Template.Asp"-->  
<!--#Include file = "2014_GBUTF.Asp"-->    
<!--#Include file = "2014_Form.Asp"-->     
<!--#Include file = "2014_Doc.Asp"--> 
<!--#Include file = "2014_Xls.Asp"--> 
<!--#Include file = "2014_Cache.Asp"--> 
<!--#Include file = "2014_Cookies.Asp"--> 
<!--#Include file = "2014_Author.Asp"--> 
<!--#Include file = "2015_Formatting.Asp"-->  
<!--#Include file = "2015_Color.Asp"-->  
<!--#Include file = "2015_Param.Asp"--> 
<!--#Include file = "2015_ASP.Asp"--> 
<!--#Include file = "2015_SqlIn.Asp"-->   
<!--#Include file = "2015_Json.Asp"--> 
<!--#Include file = "2015_PHP.Asp"-->    
<!--#Include file = "2015_APGeneral.Asp"-->  
<!--#Include file = "2016_Log.Asp"-->   
<!--#Include file = "SystemInfo.Asp"-->   
<!--#Include file = "2015_Param.Asp"-->    
<!--#Include file = "2016_WebControl.Asp"-->   
<!--#Include file = "2015_NewWebFunction.Asp"--> 
<!--#Include file = "2016_SaveData.Asp"-->   
<!--#Include file = "2015_ToMyPHP.Asp"-->   
 
<% 


server.scriptTimeout = 9999                                                     '超时
session.timeout = 1440                                                          '最大值为24小时，也就是说你最大可以session.timeout=1440，1441都是不可以
response.addheader "Content-Type", "text/html; charset=gb2312" 
'禁止缓冲
response.buffer = true 
response.expires = 0 
response.expiresAbsolute = now() - 1 
response.cacheControl = "no-cache" 
response.addHeader "Expires", date() 
response.addHeader "Pragma", "no-cache" 
response.addHeader "Cache-Control", "private, no-cache, must-revalidate" 
dim conn, rs, rsx, rss, rst, rsd, tempRs, tempRs2, rsTemp 
dim pubTimer : pubTimer = now()                                                 'Timer()不要用这个(20150728)

'Note:开发程序原则，能简单，绝对不复杂，能修改html的，绝对不修改asp，能傻瓜式，绝对不要用脑(20160121)
dim WEBCOLUMNTYPE : WEBCOLUMNTYPE = "首页|文本|产品|新闻|视频|下载|案例|留言|反馈|招聘|订单" 
dim WEB_VIEWURL : WEB_VIEWURL = "/index.asp"                                   '前台显示
dim WEB_ADMINURL : WEB_ADMINURL = "/admin/index.asp"                                  '后台显示
dim EDITORTYPE : EDITORTYPE = "asp"                                             '程序类型，是ASP,或PHP,或jSP,或.NET
dim webDir : webDir = "" 
MDBPath = "" 
if checkFile(webDir & MDBPath) = false then
    webDir = "/" & getUrlDir(remoteUrlFileName(getThisUrl())) 
    webDir = mid(webDir, 1, len(webDir) - 1) 
    webDir = replace(LCase(webDir), "admin/", "") 
'call echo(webDir,mdbpath)
end if
dim isOnCacheHtml : isOnCacheHtml = false                                       '是否开启缓冲html

dim databaseType : databaseType = ""                                      '数据库类型，Access与SqlServer之分
dim sqlServerHostIP : sqlServerHostIP = "localhost"                          'SQL主机IP地址（服务器名）在2008版本使用
dim sqlServerUsername : sqlServerUsername = "sa"                                'SQL数据库用户名
dim sqlServerPassword : sqlServerPassword = ""                                'SQL数据库用户密码
dim sqlServerDatabaseName : sqlServerDatabaseName = "webdata"                   'SQL数据库名
dim db_PREFIX : db_PREFIX = "x2asdy_"                                               '表前缀
'=========



 
dim sArraySysColor 
sArraySysColor = array("#999", "#666") 

dim makeHtmlFileToLCase : makeHtmlFileToLCase = true                            '生成HTML文件转小写
dim isWebLabelClose : isWebLabelClose = true                                    '闭合标签(20150831)
dim handleisCache : handleisCache = false                                       '缓冲是否处理了

dim isCacheTip : isCacheTip = false                                             '是否开启缓冲提示
dim adminDir : adminDir = "/admin/"                                               '后台目录


dim openErrorLog : openErrorLog = true                                          '开启错误日志
dim openWriteSystemLog : openWriteSystemLog = "|txt|database|"                  '开启写系统日志 txt写入文本 database写入数据库
dim isTestEcho : isTestEcho = true                                              '开关测试回显
dim webVersion : webVersion = "PAAJCMS v1.5"                                  '网站版本


'dim WEB_CACHEFile : WEB_CACHEFile=webDir & "/web/"& EDITORTYPE &"cachedata.txt"                                '缓冲文件
dim WEB_CACHEFile : WEB_CACHEFile = "[adminDir]/[EDITORTYPE]cachedata.txt" '缓冲文件
dim WEB_CACHEContent : WEB_CACHEContent = ""                                    '缓冲文件内容

dim language : language = "#en-us"                                              'en-us  | zh-cn | zh-tw
dim moduleReplaceArray : moduleReplaceArray = array("", "")                     '替换数组20160630

 
'前端用到
dim code, pubCode                                                               'html代码
dim templateName                                                                '模板名称
dim cfg_webSiteUrl, cfg_webTemplate, cfg_webImages, cfg_webCss, cfg_webJs, cfg_webTitle, cfg_webKeywords, cfg_webDescription, cfg_webSiteBottom, cfg_flags 
dim glb_columnName, glb_columnId, glb_id, glb_columnType, glb_columnENType, glb_table, glb_detailTitle, glb_flags 
'dim webTemplate                                                                 '网站模板路径
dim glb_url, glb_filePath                                                       '当前链接网址,和文件路径
dim glb_isonhtml                                                                '是否生成静态网页
dim glb_locationType                                                            '位置类型

dim glb_bodyContent                                                             '主体内容
dim glb_articleAuthor                                                           '文章作者
dim glb_articleAdddatetime                                                      '文章添加时间
dim glb_articlehits                                                             '文章浏览器
dim glb_upArticle                                                               '上一篇文章
dim glb_downArticle                                                             '下一篇文章
dim glb_aritcleRelatedTags                                                      '文章标签组
dim glb_aritcleSmallImage, glb_aritcleBigImage                                  '文章小图与文章大图
dim glb_searchKeyWord                                                           '搜索关键词
dim cacheHtmlFilePath                                                           '缓冲html文件路径
dim positionEndStr                                                              '面包线尾部追加内容

dim isMakeHtml                                                                  '是否生成网页










'*****************************************  语言
'批量替换语言
function batchReplaceL(content, str)
    dim splStr, s, i 
    splStr = split(str, "|*|") 
    for i = 0 to uBound(splStr)
        s = splStr(i) 
        if s <> "" then
            content = replaceL(content, s) 
        end if 
    next 
    batchReplaceL = content 
end function 
'替换语言
function replaceL(content, str)
    replaceL = replace(content, str, setL(str)) 
end function 
'语言
function setL(str)
    dim c 
    c = str 
    if language = "en-us" then
        c = languageEN(c) 
    end if 
    setL = c 
end function 
'处理显语言  c=handleDisplayLanguage(c,"loginok")
function handleDisplayLanguage(c, sType)
    '繁体就直接转换了，不要再一个一个转了，
    if language = "zh-tw" then
        handleDisplayLanguage = simplifiedTransfer(c) 
        exit function 
    end if 
    if sType = "login" then

        c = batchReplaceL(c, "请不要输入特殊字符|*|输入正确|*|用户名可以用字母|*|用户名可以用字母|*|您的用户名为空|*|密码可以用字母|*|您的密码为空") 
        c = batchReplaceL(c, "登录后台|*|管理员登录|*|如果您不是管理员|*|请立即停止您的登陆行为|*|用户名|*|版") 
        c = batchReplaceL(c, "密&nbsp;码|*|密码|*|请输入|*|登 录|*|登录|*|重 置|*|重置") 


    elseif sType = "loginok" then
        c = batchReplaceL(c, "后台地图|*|清除缓冲|*|超级管理员|*|当前位置|*|管理员信息|*|修改密码|*|最新访客信息|*|开发团队|*|版权所有|*|开发与支持团队") 
        c = batchReplaceL(c, "进入在线修改模式|*|系统信息|*|免费开源版|*|授权信息|*|服务器名称|*|服务器版本|*|交流群|*|相关链接|*|登录后台") 
        c = batchReplaceL(c, "用户名|*|表前缀|*|帮助|*|退出|*|您好|*|首页|*|权限|*|端口|*|邮箱|*|官网|*|版|*|红尘云孙") 
        c = batchReplaceL(c, "系统管理|*|我的面板|*|栏目|*|模板|*|会员|*|生成|*|更多设置") 

        c = batchReplaceL(c, "站点配置|*|网站统计|*|生成|*|后台操作日志|*|后台管理员|*|网站栏目|*|分类信息|*|评论|*|搜索统计|*|单页|*|友情链接|*|招聘管理") 
        c = batchReplaceL(c, "反馈|*|留言|*|会员配置|*|竞价词|*|网站布局|*|网站模块|*|后台菜单|*|执行|*|仿站") 


    end if 
    handleDisplayLanguage = c 
end function 

'为英文
function languageEN(str)
    dim c 
    if str = "登录成功，正在进入后台..." then
        c = "Login successfully, is entering the background..." 
    elseif str = "账号密码错误<br>登录次数为 " then
        c = "Account password error <br> login " 
    elseif str = "登录成功，正在进入后台..." then
        c = "Login successfully, is entering the background..." 
    elseif str = "退出成功" then
        c = "Exit success" 
    elseif str = "退出成功，正在进入登录界面..." then
        c = "Quit successfully, is entering the login screen..." 
    elseif str = "清除缓冲完成，正在进入后台界面..." then
        c = "Clear buffer finish, is entering the background interface..." 
    elseif str = "提示信息" then
        c = "Prompt info" 
    elseif str = "如果您的浏览器没有自动跳转，请点击这里" then
        c = "If your browser does not automatically jump, please click here" 
    elseif str = "倒计时" then
        c = "Countdown " 
    elseif str = "后台地图" then
        c = "Admin map" 
    elseif str = "清除缓冲" then
        c = "Clear buffer" 
    elseif str = "超级管理员" then
        c = "Super administrator" 
    elseif str = "当前位置" then
        c = "current location" 
    elseif str = "管理员信息" then
        c = "Admin info" 
    elseif str = "修改密码" then
        c = "Modify password" 
    elseif str = "用户名" then
        c = "username" 
    elseif str = "表前缀" then
        c = "Table Prefix" 
    elseif str = "进入在线修改模式" then
        c = "online modification" 
    elseif str = "系统信息" then
        c = "system info" 
    elseif str = "授权信息" then
        c = "Authorization information" 
    elseif str = "免费开源版" then
        c = "Free open source" 
    elseif str = "服务器名称" then
        c = "Server name" 
    elseif str = "服务器版本" then
        c = "Server version" 
    elseif str = "最新访客信息" then
        c = "visitor info" 
    elseif str = "开发团队" then
        c = "team info" 
    elseif str = "版权所有" then
        c = "copyright" 
    elseif str = "开发与支持团队" then
        c = "Develop and support team" 
    elseif str = "交流群" then
        c = "Exchange group" 
    elseif str = "相关链接" then
        c = "Related links" 
    elseif str = "系统管理" then
        c = "System" 
    elseif str = "我的面板" then
        c = "My panel" 
    elseif str = "栏目" then
        c = "Column" 
    elseif str = "模板" then
        c = "Template" 
    elseif str = "会员" then
        c = "Member" 
    elseif str = "生成" then
        c = "Generation" 
    elseif str = "更多设置" then
        c = "More settings" 


    elseif str = "登录后台" then
        c = "Login management background" 
    elseif str = "管理员登录" then
        c = "Administrator login " 
    elseif str = "如果您不是管理员" then
        c = "If you are not an administrator" 
    elseif str = "请立即停止您的登陆行为" then
        c = "Please stop your login immediately" 
    elseif str = "密&nbsp;码" or str = "密码" then
        c = "password" 
    elseif str = "请输入" then
        c = "Please input" 
    elseif str = "登 录" or str = "登录" then
        c = "login" 
    elseif str = "重 置" or str = "重置" then
        c = "reset" 
    elseif str = "请不要输入特殊字符" then
        c = "Please do not enter special characters" 
    elseif str = "输入正确" then
        c = "Input correct" 
    elseif str = "用户名可以用字母" then
        c = "Username use " 
    elseif str = "您的用户名为空" then
        c = "Your username is empty" 
    elseif str = "密码可以用字母" then
        c = "Passwords use " 
    elseif str = "您的密码为空" then
        c = "Your password is empty" 
    elseif str = "站点配置" then
        c = "Site configuration" 
    elseif str = "网站统计" then
        c = "Website statistics" 
    elseif str = "后台操作日志" then
        c = "Admin log " 
    elseif str = "后台管理员" then
        c = "Background manager" 
    elseif str = "网站栏目" then
        c = "Website column" 
    elseif str = "分类信息" then
        c = "Classified information" 
    elseif str = "搜索统计" then
        c = "Search statistics" 
    elseif str = "单页" then
        c = "Single page" 
    elseif str = "友情链接" then
        c = "Friendship link" 
    elseif str = "招聘管理" then
        c = "Recruitment management" 
    elseif str = "反馈" then
        c = "Feedback management" 
    elseif str = "留言" then
        c = "message management" 
    elseif str = "会员配置" then
        c = "Member allocation" 
    elseif str = "竞价词" then
        c = "Bidding words" 
    elseif str = "网站布局" then
        c = "Website layout" 
    elseif str = "网站模块" then
        c = "Website module" 
    elseif str = "后台菜单" then
        c = "Background menu" 
    elseif str = "仿站" then
        c = "Template website " 

    elseif str = "11111" then
        c = "1111111" 



    elseif str = "执行" then
        c = "implement " 
    elseif str = "评论" then
        c = "comment " 
    elseif str = "生成" then
        c = "generate " 
    elseif str = "权限" then
        c = "jurisdiction" 
    elseif str = "帮助" then
        c = "Help" 
    elseif str = "退出" then
        c = "sign out" 
    elseif str = "您好" then
        c = "hello" 
    elseif str = "首页" then
        c = "home" 
    elseif str = "端口" then
        c = "port" 
    elseif str = "官网" then
        c = "official website" 
    elseif str = "邮箱" then
        c = "Emai" 
    elseif str = "红尘云孙" then
        c = "Cloud" 

    elseif str = "版" then
        c = " edition"  

    else
        c = str 
    end if 
    languageEN = c 
end function 
%> 


















