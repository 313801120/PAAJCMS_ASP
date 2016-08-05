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
<!--#Include file = "Config.Asp"-->  
<% 
'选择处理
select case request("act")
    case "CreateVBDialog", "vb1" : createVBDialog()                                 '创建VB需要的面板
end select
response.end() 
sub createVBDialog()
    dim splStr, i, s, c, content, rootFolder, zFormFolder, dataFolder, path, toPath, formName, basList, frmList, toolsMenu, frmListStr, openFrom, nTools, toolsClick, menuTitle 
    dim yesSortPlug, tempStr, tableConfig, formStrList, formSetList, tableName, softVersion, thisMDBPath, thisAccessPass, ctlFileList, startStr, endStr 
    dim configFolder 
    dim helpFolder                                                                  '帮助文件夹
    dim toolsName                                                                   '工具名称 给那个动态菜单加菜单20150218
    dim nTools2                                                                     '工具二排序
    dim nTools3                                                                     '网站制作排序
    dim tempContent 
    softVersion = "Version1.1" 
    rootFolder = "/VB工程/" 
    '复制文件夹 免得忘记保存而后悔20150313
    dim toFolder                                                                    '复制到文件夹
    if checkFolder(rootFolder) then
        for i = 0 to 90
            s = "/VB工程_" & i & "/" 
            if checkFolder(s) = false and toFolder = "" then
                toFolder = s 
            elseIf checkFolder(s) = true and toFolder <> "" then
                toFolder = "" 
            end if 
        next 
    end if 
    'Call EErr(RootFolder, ToFolder)
    call MoveFolder(rootFolder, toFolder) 
    call DeleteFolder(rootFolder) 
    call createFolder(rootFolder) 
    zFormFolder = rootFolder & "zForm/" 
    call createFolder(zFormFolder) 
    helpFolder = rootFolder & "Help/" 
    call createFolder(helpFolder) 
    call CopyFile("/DataDir/VB模块/Help/SystemLogin.ini", helpFolder & "SystemLogin.ini") 
    call CopyFile("/DataDir/VB模块/Help/Js处理.html", helpFolder & "Js处理.html") 

    call CopyFile("/DataDir/VB模块/右键下拉菜单代码.txt", rootFolder & "右键下拉菜单代码.txt") 
    call CopyFile("/DataDir/VB模块/快速登录.vbs", rootFolder & "快速登录.vbs") 



    'ConfigFolder = RootFolder & "Config/"
    'Call createFolder(ConfigFolder)

    c = "|Print.Asp|Access.Asp|Cai.Asp|EncDec.Asp|Function.Asp|GetWord.Asp|RegExp.Asp|StringNumber.Asp|XML.Asp|Time.Asp|URL.Asp|FSO.Asp|Check.Asp|Html.Asp|AdodbStream.Asp|Dictionary.Asp|PinYin.Asp|2014_Array.Asp|2014_Css.Asp|2014_GBUTF.Asp|2015_Formatting.Asp|2015_Param.Asp|2015_ASP.Asp|2015_Color.Asp|FunHTML.Asp|2015_PHP.Asp|2015_WebSql.Asp|2015_hexSwitch.Asp|ASPPHPAccess.Asp|2014_Author.Asp|2016_WebControl.asp|2014_News.Asp|2014_MainInfo.Asp|2014_SiteMap.Asp|2014_Template.Asp|" 
    c = c & "2014_Search.Asp|2014_Class.Asp|Incpage.Asp|2014_Banner.Asp|2014_Links.Asp|2014_Doc.Asp|2015_APGeneral.Asp|IE.asp|2014_Js.Asp|" 
    '2015_ToMyPHP.Asp|不要创建用不着
    dim searchStr, replaceStr 
    '添加模块列表
    splStr = split(c, "|")                                                          '|ZClassAspCode.Asp
    for each path in splStr
        if path <> "" and 1 = 1 then
            formName = mid(path, 1, inStrRev(path, ".") - 1) 
            '判断文件以数据打头，则加前缀 否则VB调用会出错
            if inStr("0123456789", left(formName, 1)) > 0 then
                formName = "y_" & formName                                                      '区部它是前面有数字
            end if 
            content = readFile(path, "") 

            searchStr = getStrCut(content, "'footer start", "'footer end", 2) 
            if searchStr <> "" then
                replaceStr = vbCrLf & getftext("\DataDir\VB模块\2015_Param.Asp.txt") 
                content = replace(content, searchStr, replaceStr) 
                call echo(path, "替换footer成功") 
            end if 

            for i = 1 to 10
                searchStr = getStrCut(content, "'vbdel start", "'vbdel end", 1) 
                if searchStr <> "" then
                    content = replace(content, searchStr, "") 
                end if 
            next 

            '组件
            if path = "ZClassAspCode.Asp" then
                toPath = zFormFolder & formName & ".cls" 
                content = strCut(content, "'VB开始", "'VB结束", 2) 
                c = "" 
                c = c & "VERSION 1.0 CLASS" & vbCrLf 
                c = c & "BEGIN" & vbCrLf 
                c = c & "  MultiUse = -1  'True" & vbCrLf 
                c = c & "  Persistable = 0  'NotPersistable" & vbCrLf 
                c = c & "  DataBindingBehavior = 0  'vbNone" & vbCrLf 
                c = c & "  DataSourceBehavior  = 0  'vbNone" & vbCrLf 
                c = c & "  MTSTransactionMode  = 0  'NotAnMTSObject" & vbCrLf 
                c = c & "END" & vbCrLf 
                c = c & "Attribute VB_Name = ""Z_" & formName & """" & vbCrLf                   '免得名称与里面函数重名
                c = c & "Attribute VB_GlobalNameSpace = False" & vbCrLf 
                c = c & "Attribute VB_Creatable = False" & vbCrLf 
                c = c & "Attribute VB_PredeclaredId = False" & vbCrLf 
                c = c & "Attribute VB_Exposed = False" & vbCrLf 
                content = c & content 

                basList = basList & "class=" & formName & "; zForm\" & formName & ".cls" & vbCrLf 
            '模块
            else
                toPath = zFormFolder & formName & ".bas" 
                content = strCut(content, "<" & "%", "%" & ">", 2) 
                content = "Attribute VB_Name = ""Z_" & formName & """" & vbCrLf & content 
                basList = basList & "Module=Z_" & formName & "; zForm\" & formName & ".bas" & vbCrLf 
            end if 
            tempContent = content 

            content = ASPToVB(content) 

            if content = "" then
                call eerr(formName, tempContent) 
            end if 

            startStr = "Function GetThisUrl()" 
            endStr = "End Function" 
            s = getStrCut(content, startStr, endStr, 1) 
            if s <> "" then
                content = replace(content, s, startStr & vbCrLf & endStr) 
            end if 


            content = replace(content, vbTab, "    ")                                       '替换Tab
            call echo(formName, toPath) 
            call CreateFile(toPath, content) 
        end if 
    next 
    path = "/DataDir/VB模块/Manage.vbp" 
    content = getFText(path) 
    '处理掉不需要的注释代码
    splStr = split(content, vbCrLf) : content = "" 
    for each s in splStr
        if left(trim(s), 1) <> "'" then
            content = content & s & vbCrLf 
        end if 
    next 
    content = replace(content, "{Bas列表}", basList) 
    '添加窗体列表
    frmListStr = "MDIForm1|Notepad|FrmHide|Login|IE|AutoAddTable|WebLog|WebIP|Eval|SortPlug|Test|Share|ShareEdition2|QQShareConfig|QQShareUrl|FrmEdit|QQShareSendUrl|YingZheng|" 
    frmListStr = "MDIForm1|Notepad|FrmHide|Login|IE|Share|ShareEdition2|QQShareConfig|QQShareUrl|QQShareSendUrl|FrmEdit" 
    '全二
    frmListStr = "MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|-|ShareEdition2|QQShareConfig|QQShareUrl|QQShareSendUrl|-|FrmCai|CaiManage|UpDateWebConfig|-|XiCiWebDialog|XiCiWebMDB|-|CaiBaiduUrl|BaiduCai|-|#SortPlug|WebCrack|Test|DownList|IEDebug|-|AutoAddTable|Ying|-|SortPlug|PseudoOriginal|-|Scan|WebShell|Eval|WebServer|-|QQ|Qzone|-|" 
    'QQ操作 'FrmListStr="MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|InputLocalFile|InputSqlTable|-|HTML5CSS3|FrmCrateNewWeb|TemplateCodeCreate|-|Test|IEDebug|-|WebLog|WebIP|WebSite|-|Eval|-|FrmCai|CaiManage|UpDateWebConfig|-|SinglePage|IEAutomatic|-|SearchEngine|SearchEngineInfoList|WebAdminScan|StatisticalData|-|WebManage|WebIeBox|-|MYProgram|-|SmallTool|ColorPicker|TxtEdit|-|ShieldWebSiteList|WebSiteList|SqlServerWebUrl|WebSiteInfo|WebShell|-|FrmSortPlug|jiayuan|"
    frmListStr = "MDIForm1|FrmHide|FrmEdit|Login|TableDialog|Notepad|IE|InputLocalFile|InputSqlTable|-|FrmCrateNewWeb|WebUrlScan|FrmCreateFunction|WebIeBox|-|Test|IEDebug|IEAutomatic|IEAutoHandle|-|Eval|-|FrmCai|CaiManage|-|WebAdminScan|ExcellentUrl|SmallTool|ColorPicker|TxtEdit|-|FrmSortPlug|AskDa|ZNRobot|JobWeb|JobCompany|FormJobCai|ComputerDir|analogBatchUP|" 
    frmListStr = frmListStr & "|FrmCrateNewHtmlPage|SinglePage|FrameLayout|FrmWebTemplateView|FrmCreateWebLable|UserMakeNav|--2|FileScan2015|HTML5CSS3|XyWebUp|MYProgram|--2|FrmAutoAddWebDefaultInfo|--2|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|IMitationWebInGoogleSEO|WebShell|FrmEvScan|FrmTestDialog1|FrmTestDialog2|FrmDown|FrmCrateNewHtmlPageOLD|Ying|YingZheng|" '

    nTools = 0                                                                      '工具菜单默认为零
    splStr = split(frmListStr, "|") 
    for each s in splStr
        '不为空  #不用
        if s <> "" and left(s, 1) <> "#" then
            '设置操作菜单
            if inStr("|-|IE|AutoAddTable|WebLog|WebIP|WebSite|Eval|SortPlug|Test|Share|QQShareConfig|QQShareUrl|QQShareSendUrl|ShareEdition2|FrmCai|CaiManage|UpDateWebConfig|XiCiWebDialog|XiCiWebMDB|IEDebug|IEAutomatic|IEAutoHandle|BaiduCai|CaiBaiduUrl|Ying|PseudoOriginal|Scan|WebShell|WebServer|WebCrack|DownList|QQ|QQName|QQVisitors|Qzone|BaiduZhidao|BaiduZhidaoWords|StatisticalData|BaiduZhidaoData|BaiduZhidaoWordsData|BaiduZhidaoMain|SearchEngine|SearchEngineInfoList|WebAdminScan|WebManage|WebIeBox|SinglePage|FrmSMS|SMSInfo|LocaDir|MYProgram|SmallTool|HTML5CSS3|FrmCrateNewWeb|FrmCrateNewHtmlPageOLD|TemplateCodeCreate|WebSiteList|SqlServerWebUrl|ShieldWebSiteList|AsynchDown|WebSiteInfo|FrmSortPlug|ExcellentUrl|WebUrlScan|FrmHangUp|XyWebUp|FrameLayout|FrmCreateFunction|QQJournal|AskDa|ZNRobot|JobWeb|FormJobCai|UserMakeNav|JobCompany|IMitationWebInGoogleSEO|FrmCrateNewHtmlPage|FileScan2015|FrmCreateWebLable||FrmWebTemplateView|ComputerDir|FrmAutoAddWebDefaultInfo|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|analogBatchUP|jiayuan|FrmEvScan|FrmTestDialog1|FrmTestDialog2|FrmDown|YingZheng|", "|" & s & "|") > 0 or left(s, 2) = "--" then
                menuTitle = "" : tableConfig = "" : tableName = "" : thisAccessPass = "" : thisMDBPath = "" 



                if s = "SortPlug" then
                    menuTitle = "模拟点击" 
                    yesSortPlug = "OK"                                                              '为作弊软件特殊使用
                    call CopyFile("/DataDir/VB模块/Action.ini", rootFolder & "Action.ini")          '拷贝配置文件


                elseIf s = "UserMakeNav" then
                    menuTitle = "网站导航生成" 
                elseIf s = "FileScan2015" then
                    menuTitle = "文件搜索2015" 
                elseIf s = "FrmControlDialog" then
                    menuTitle = "控制面板" 
                elseIf s = "FrmSqlIn" then
                    menuTitle = "Sql注入" 
                elseIf s = "FrmProgrameTransformation" then
                    menuTitle = "程序转换" 
                elseIf s = "analogBatchUP" then
                    menuTitle = "模似批量上传" 
                elseIf s = "FrmEvScan" then
                    menuTitle = "EvScan" 
                elseIf s = "FrmTestDialog1" then
                    menuTitle = "测试面板1" 
                elseIf s = "FrmTestDialog2" then
                    menuTitle = "测试面板2" 
                elseIf s = "FrmDown" then
                    menuTitle = "下载文件" 




                elseIf s = "IMitationWebInGoogleSEO" then
                    menuTitle = "仿站在Google优化" 

                    call CopyFolder("/DataDir/VB模块/IMitationWebInGoogleSEO", rootFolder & "/IMitationWebInGoogleSEO") '拷贝文件夹



                elseIf s = "XyWebUp" then
                    menuTitle = "批量上传网站" 
                elseIf s = "FrmCreateFunction" then
                    menuTitle = "生成函数面板" 
                elseIf s = "FrmSortPlug" then
                    menuTitle = "模拟操作软件" 
                elseIf s = "FrameLayout" then
                    menuTitle = "网页框架布局" 
                elseIf s = "FrmHangUp" then
                    menuTitle = "禁止窗口必需验证" 
                elseIf s = "MYProgram" then
                    menuTitle = "我的程序" 
                    call CopyFile("/DataDir/VB模块/AspFunction.ini", helpFolder & "/AspFunction.ini") '拷贝配置文件
                elseIf s = "SmallTool" then
                    menuTitle = "小工具" 
                elseIf s = "HTML5CSS3" then
                    menuTitle = "Html5+Css3快速生成工具" 
                    ctlFileList = ctlFileList & "UserControl=zForm\ControlDialog.ctl" & vbCrLf 
                    call CopyFile("\DataDir\VB模块\ControlDialog.ctl", zFormFolder & "ControlDialog.ctl") '复制文件
                    call CopyFile("\DataDir\VB模块\ControlDialog.ctx", zFormFolder & "ControlDialog.ctx") '复制文件
                elseIf s = "FrmCrateNewWeb" then
                    menuTitle = "创建新网页" 
                elseIf s = "FrmCrateNewHtmlPageOLD" then
                    menuTitle = "创建新网页(旧版)" 

                elseIf s = "FrmCreateWebLable" then
                    menuTitle = "生成网站嵌套标签" 
                elseIf s = "FrmCrateNewHtmlPage" then
                    menuTitle = "创建HTML页面" 
                elseIf s = "TemplateCodeCreate" then
                    menuTitle = "模板代码创建生成器" 
                elseIf s = "FrmWebTemplateView" then
                    menuTitle = "网站模板预览" 
                elseIf s = "FrmAutoAddWebDefaultInfo" then
                    menuTitle = "自动添加网站默认信息" 

                elseIf s = "IE" then
                    menuTitle = "浏览器" 
                elseIf s = "Test" then
                    menuTitle = "测试窗口" 
                elseIf s = "AutoAddTable" then
                    menuTitle = "自动填表" 
                elseIf s = "YingZheng" then
                    menuTitle = "破解验证码" 
                elseIf s = "Ying" then
                    menuTitle = "简易验证码识别系统" 
                    call CopyFolder("/DataDir/VB模块/图片破解", rootFolder & "图片破解")            '加载配置文件夹
                elseIf s = "XiCiWebDialog" then
                    menuTitle = "西祠论坛群发" 
                elseIf s = "FrmSMS" then
                    menuTitle = "短信" 
                elseIf s = "SMSInfo" then
                    menuTitle = "短信管理" 
                    tableName = "SMSInfo" 
                    thisAccessPass = "****" 
                    thisMDBPath = " ""I:\个人制作\个人\短信\SMS.mdb""" 
                    tableConfig = tableConfig & "发信人|DXName|1500|True[分]" 
                    tableConfig = tableConfig & "收信人|ToName|1500|True[分]" 
                    tableConfig = tableConfig & "发信时期|AddDate|1000|True[分]" 
                    tableConfig = tableConfig & "留言内容|Content|4000|True[分]" 
                elseIf s = "LocaDir" then
                    menuTitle = "本地目录" 
                    tableName = "LocaDir" 
                    thisMDBPath = "App.Path & ""/../../网站备份\数据库\LocaDir.Mdb""" 
                    thisMDBPath = """SqlServer数据库""" 
                    tableConfig = tableConfig & "文件夹|FolderPath|5000|True[分]" 
                    tableConfig = tableConfig & "文件名称|FileName|3500|True[分]" 
                    tableConfig = tableConfig & "文件大小|FileSize|1200|True[分]" 
                    tableConfig = tableConfig & "文件类型|FileType|1200|True[分]" 
                elseIf s = "ComputerDir" then
                    menuTitle = "电脑Scan" 
                    tableName = "ComputerDir" 
                    thisMDBPath = "App.Path & ""/../../网站备份\数据库\ComputerDir.mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|1600|True[分]" 
                    tableConfig = tableConfig & "文件路径|FilePath|3500|True[分]" 
                    tableConfig = tableConfig & "文件名称|FileName|1200|True[分]" 
                    tableConfig = tableConfig & "文件大小|FileSize|1200|True[分]" 
                    tableConfig = tableConfig & "文件类型|FileType|1200|True[分]" 



                elseIf s = "XiCiWebMDB" then
                    menuTitle = "西祠论坛数据管理" 
                    tableName = "XiCiWeb" 
                    tableConfig = tableConfig & "网址|HttpUrl|6000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|3000|True[分]" 
                    tableConfig = tableConfig & "是否使用|Through|1000|True[分]" 
                elseIf s = "WebLog" then
                    menuTitle = "网站日志" 
                    tableName = "WebLog" 
                    thisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb""" 
                    tableConfig = tableConfig & "来访网址|GoUrl|3500|True[分]" 
                    tableConfig = tableConfig & "当前浏览网址|HttpUrl|3500|True[分]" 
                    tableConfig = tableConfig & "访问时间|GoDateTime|2000|True[分]" 
                    tableConfig = tableConfig & "搜索引擎类型|SearchEngineType|1500|True[分]" 
                    tableConfig = tableConfig & "搜索关键词|SearchWord|2000|True[分]" 
                    tableConfig = tableConfig & "IP|IP|1600|True[分]" 
                    tableConfig = tableConfig & "IE|IE|2200|True[分]" 
                    tableConfig = tableConfig & "日期|AddDate|1200|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|2000|True[分]" 
                elseIf s = "WebIP" then
                    menuTitle = "网站流量" 
                    tableName = "WebIP" 
                    thisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb""" 
                    tableConfig = tableConfig & "添加日期|AddDate|3000|True[分]" 
                    tableConfig = tableConfig & "浏览次数|PVCount|1000|True[分]" 
                    tableConfig = tableConfig & "域名数|WebSiteCount|1000|True[分]" 
                    tableConfig = tableConfig & "IP|IPCount|1000|True[分]" 
                    tableConfig = tableConfig & "IE|IECount|1000|True[分]" 
                    tableConfig = tableConfig & "百度数|Baidu|1000|True[分]" 
                    tableConfig = tableConfig & "Google数|Google|1000|True[分]" 
                    tableConfig = tableConfig & "Yahoo数|Yahoo|1000|True[分]" 
                    tableConfig = tableConfig & "搜搜数|Soso|1000|True[分]" 
                    tableConfig = tableConfig & "360数|Cn360|1000|True[分]" 
                    tableConfig = tableConfig & "QQ空间数|Qzone|1000|True[分]" 
                    tableConfig = tableConfig & "微信数|WeiXin|1000|True[分]" 


                elseIf s = "jiayuan" then
                    menuTitle = "世纪佳缘" 
                    tableName = "jiayuan" 
                    thisMDBPath = """../../网站备份/数据库/jiayuan.mdb""" 
                    tableConfig = tableConfig & "会员ID|uid|1000|True[分]" 
                    tableConfig = tableConfig & "昵称|nickname|1500|True[分]" 
                    tableConfig = tableConfig & "性别|sex|800|True[分]" 
                    tableConfig = tableConfig & "年龄|Age|1000|True[分]" 
                    tableConfig = tableConfig & "身高|height|800|True[分]" 
                    tableConfig = tableConfig & "体重|Weight|1000|True[分]" 
                    tableConfig = tableConfig & "婚姻|marriage|1500|True[分]" 
                    tableConfig = tableConfig & "教育|education|1500|True[分]" 
                    tableConfig = tableConfig & "工资|minMonthlySalary|1000|True[分]" 
                    tableConfig = tableConfig & "兄弟姐妹|BrothersAndSisters|1500|True[分]" 
                    tableConfig = tableConfig & "个人介绍|shortnote|2000|True[分]" 



                elseIf s = "AskDa" then
                    menuTitle = "遇到问题和解决" 
                    tableName = "AskDa" 
                    thisMDBPath = """../../网站备份/数据库/AskDa.mdb""" 
                    tableConfig = tableConfig & "问|Ask|5000|True[分]" 
                    tableConfig = tableConfig & "答|Da|5000|True[分]" 
                    tableConfig = tableConfig & "添加日期|AddDate|3000|True[分]" 
                elseIf s = "ZNRobot" then
                    menuTitle = "智慧机器人" 
                    tableName = "ZNRobot" 
                    thisMDBPath = """../../网站备份/数据库/ZNRobot.mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|1600|True[分]" 
                    tableConfig = tableConfig & "标题|Title|4000|True[分]" 
                    tableConfig = tableConfig & "相关标签|RelatedTags|4000|True[分]" 
                elseIf s = "JobWeb" then
                    menuTitle = "招聘岗位管理" 
                    tableName = "JobWeb" 
                    thisMDBPath = """../../网站备份/数据库/JobAccess.mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|1600|True[分]" 
                    tableConfig = tableConfig & "搜索关键词|SearchWord|1300|True[分]" 
                    tableConfig = tableConfig & "标题|Title|3500|True[分]" 
                    tableConfig = tableConfig & "城市|City|800|True[分]" 
                    tableConfig = tableConfig & "区域|Area|1000|True[分]" 
                    tableConfig = tableConfig & "公司规模|CompanyScale|1200|True[分]" 
                    tableConfig = tableConfig & "经验|Experience|1200|True[分]" 
                    tableConfig = tableConfig & "学历|Education|1200|True[分]" 
                    tableConfig = tableConfig & "职位月薪|MonthlySalary|1200|True[分]" 
                    tableConfig = tableConfig & "招聘时间|JobTime|1200|True[分]" 
                    tableConfig = tableConfig & "公司名称|CompanyName|3800|True[分]" 
                elseIf s = "JobCompany" then
                    menuTitle = "招聘公司管理" 
                    tableName = "JobCompany" 
                    thisMDBPath = """../../网站备份/数据库/JobAccess.mdb""" 
                    tableConfig = tableConfig & "公司名称|CompanyName|3000|True[分]" 
                    tableConfig = tableConfig & "公司网址|CompanyUrl|3000|True[分]" 
                    tableConfig = tableConfig & "城市|City|800|True[分]" 
                    tableConfig = tableConfig & "区域|Area|1000|True[分]" 
                    tableConfig = tableConfig & "公司规模|CompanyScale|1200|True[分]" 
                    tableConfig = tableConfig & "公司性质|CompanyNature|1200|True[分]" 
                    tableConfig = tableConfig & "公司行业|CompanyIndustry|4000|True[分]" 
                    tableConfig = tableConfig & "最新发布|NewJobTime|1600|True[分]" 
                    tableConfig = tableConfig & "发布天数|ReleaseJobDay|1200|True[分]" 

                elseIf s = "FormJobCai" then
                    menuTitle = "招聘网站采集" 
                elseIf s = "WebSite" then
                    menuTitle = "域名统计" 
                    tableName = "WebSite" 
                    thisMDBPath = "GetIniS(""init"", ""DefaultQQAccessPath"", """") & ""\WebLog.Mdb""" 
                    tableConfig = tableConfig & "域名|WebSite|5000|True[分]" 
                    tableConfig = tableConfig & "数量|WebSiteCount|1000|True[分]" 

                elseIf s = "WebUrlScan" then
                    menuTitle = "网站链接Scan" 
                    tableName = "WebUrlScan" 
                    thisMDBPath = "App.Path & ""\..\..\网站备份\数据库\WebUrlScan.Mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|1600|True[分]" 
                    tableConfig = tableConfig & "标题|Title|4000|True[分]" 
                    tableConfig = tableConfig & "网页状态|WebState|1200|True[分]" 
                    tableConfig = tableConfig & "打开时间|OpenSpeed|1200|True[分]" 
                    tableConfig = tableConfig & "文件大小|WebFileSize|1200|True[分]" 
                    tableConfig = tableConfig & "链接类型|LinkType|1200|True[分]"              '分内链与外链
                elseIf s = "ExcellentUrl" then
                    menuTitle = "优秀网站" 
                    tableName = "ExcellentUrl" 
                    thisMDBPath = "App.Path & ""\..\..\网站备份\数据库\ExcellentUrl.Mdb""" 
                    'TableConfig = TableConfig & "大类|BigClassName|1600|True[分]"
                    'TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|8000|True[分]" 
                    tableConfig = tableConfig & "IP|IP|2200|True[分]" 
                    'TableConfig = TableConfig & "打开时间|OpenSpeed|1200|True[分]"
                    'TableConfig = TableConfig & "文件大小|WebFileSize|1200|True[分]"

                elseIf s = "SearchEngine" then
                    menuTitle = "搜索引擎" 
                    tableName = "SearchEngine" 
                    'ThisMDBPath = "App.Path & ""/../../网站备份\数据库\SearchEngine.Mdb"""    '本地Access数据库
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库
                    'TableConfig = TableConfig & "大类|BigClassName|1500|True[分]"
                    tableConfig = tableConfig & "网址|HttpUrl|4000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|6000|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "收录时间|IncludedTime|1500|True[分]" 

                    tableConfig = tableConfig & "Table数|TableNumb|1300|True[分]" 
                    tableConfig = tableConfig & "Div数|DivNumb|1000|True[分]" 
                    tableConfig = tableConfig & "Span数|SpanNumb|1000|True[分]" 
                    tableConfig = tableConfig & "Ul数|UlNumb|1000|True[分]" 
                    tableConfig = tableConfig & "Dl数|DlNumb|1000|True[分]" 
                elseIf s = "SqlServerWebUrl" then
                    menuTitle = "网站网址" 
                    tableName = "WebUrl" 
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库
                    tableConfig = tableConfig & "网址|HttpUrl|4000|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "域名ID|WebSiteID|1000|True[分]" 
                    tableConfig = tableConfig & "状态|WebStatus|1200|True[分]" 
                    tableConfig = tableConfig & "打开时间|OpenTime|1800|True[分]" 
                    tableConfig = tableConfig & "处理|Handle|1000|True[分]" 
                elseIf s = "WebSiteList" then
                    menuTitle = "域名列表" 
                    tableName = "WebSiteList" 
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库

                    tableConfig = tableConfig & "大类名称|BigClassName|1500|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "状态|WebStatus|1000|True[分]" 
                    tableConfig = tableConfig & "IP|IP|2200|True[分]" 
                    tableConfig = tableConfig & "IP地址|IPAddress|2000|True[分]" 
                    tableConfig = tableConfig & "打开时间|OpenTime|1400|True[分]" 
                    tableConfig = tableConfig & "处理|Handle|1000|True[分]" 
                elseIf s = "WebSiteInfo" then
                    menuTitle = "域名列表" 
                    tableName = "WebSiteInfo" 
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "总页数|WebCountPage|1000|True[分]" 
                    tableConfig = tableConfig & "首页内链数|HomeStationLink|1500|True[分]" 
                    tableConfig = tableConfig & "首页外链数|HomeInboundLink|1500|True[分]" 
                    tableConfig = tableConfig & "全部内链数|StationLink|1500|True[分]" 
                    tableConfig = tableConfig & "全部外链数|InboundLink|1500|True[分]" 

                elseIf s = "ShieldWebSiteList" then
                    menuTitle = "屏蔽域名列表" 
                    tableName = "ShieldWebSiteList" 
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "IP|IP|2000|True[分]" 
                    tableConfig = tableConfig & "状态|WebStatus|1200|True[分]" 
                    tableConfig = tableConfig & "打开时间|OpenTime|1800|True[分]" 
                    tableConfig = tableConfig & "处理|Handle|1000|True[分]" 
                elseIf s = "SearchEngineInfoList" then
                    menuTitle = "搜索引擎列表" 
                    tableName = "SearchEngineInfoList" 
                    thisMDBPath = "App.Path & ""/../../网站备份\数据库\SearchEngineInfoList.Mdb""" 
                    tableConfig = tableConfig & "大类|BigClassName|1500|True[分]" 
                    'TableConfig = TableConfig & "网址|HttpUrl|4000|True[分]"
                    tableConfig = tableConfig & "标题|Title|6000|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|3000|True[分]" 
                    tableConfig = tableConfig & "收录时间|IncludedTime|1300|True[分]" 
                    tableConfig = tableConfig & "页数|PageNumber|800|True[分]" 
                    tableConfig = tableConfig & "排名|Ranking|800|True[分]" 
                elseIf s = "WebAdminScan" then
                    menuTitle = "网站后台Scan" 
                    tableName = "WebAdminScan" 
                    thisMDBPath = "App.Path & ""/../../网站备份\数据库\WebAdminScan.Mdb""" 
                    'TableConfig = TableConfig & "大类|BigClassName|1500|True[分]"
                    tableConfig = tableConfig & "网址|HttpUrl|7000|True[分]" 
                    tableConfig = tableConfig & "响应状态|WebState|1200|True[分]" 
                    tableConfig = tableConfig & "打开时间|OpenSpeed|1200|True[分]" 
                    tableConfig = tableConfig & "文件大小|WebFileSize|1200|True[分]" 
                elseIf s = "Eval" then
                    menuTitle = "Eval" 
                elseIf s = "WebManage" then
                    menuTitle = "模板网页制作" 
                elseIf s = "WebIeBox" then
                    menuTitle = "模板样式管理" 
                    'ElseIf S = "Share" Then
                'MenuTitle = "自动分享"
                elseIf s = "ShareEdition2" then
                    menuTitle = "QQ点赞批量分享 改进版" 
                elseIf s = "QQShareConfig" then
                    menuTitle = "QQ分享配置" 
                    tableName = "QQShareConfig" 
                    tableConfig = tableConfig & "网站标题|WebTitle|3000|True[分]" 
                    tableConfig = tableConfig & "网站描述|WebDescription|3000|True[分]" 
                    tableConfig = tableConfig & "网站图片地址|ImgPath|3000|True[分]" 
                    tableConfig = tableConfig & "评语|PingYu|2000|True[分]" 
                    tableConfig = tableConfig & "是否使用|Through|2000|True[分]" 
                elseIf s = "QQShareUrl" then
                    menuTitle = "QQ分享网址" 
                    tableName = "QQShareUrl" 
                    tableConfig = tableConfig & "网址|HttpUrl|6000|True[分]" 
                    tableConfig = tableConfig & "审核|Through|800|True[分]" 

                elseIf s = "QQShareSendUrl" then
                    menuTitle = "已分享网址" 
                    tableName = "QQShareSendUrl" 
                    tableConfig = tableConfig & "已分享网址|HttpUrl|6000|True[分]" 
                    tableConfig = tableConfig & "通过|Through|800|True[分]" 

                elseIf s = "FrmCai" then
                    menuTitle = "采集" 
                elseIf s = "CaiManage" then
                    menuTitle = "采集管理" 
                    tableName = "Cai" 
                    thisMDBPath = "App.Path & ""/../../网站备份\数据库\Cai.Mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|2000|True[分]" 
                    tableConfig = tableConfig & "小类名称|SmallClassName|2000|True[分]" 
                    tableConfig = tableConfig & "子类名称|ThreeClassName|2000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|4000|True[分]" 
                    tableConfig = tableConfig & "网址|HttpUrl|4000|True[分]" 
                    tableConfig = tableConfig & "内容|Content|4000|True[分]" 
                elseIf s = "UpDateWebConfig" then
                    menuTitle = "批量上传到网站配置" 
                    tableName = "UpDateWebConfig" 
                    tableConfig = tableConfig & "批量上传到网站|WebSite|4000|True[分]" 
                    tableConfig = tableConfig & "网站标题|WebTitle|2000|True[分]" 
                    tableConfig = tableConfig & "网站关键词|WebKeywords|2000|True[分]" 
                    tableConfig = tableConfig & "网站描述|WebDescription|2000|True[分]" 
                    tableConfig = tableConfig & "文章追加后缀|SuffixCustom|2000|True[分]" 
                    tableConfig = tableConfig & "上传|Through|800|True[分]" 
                elseIf s = "BaiduCai" then
                    menuTitle = "百度网址管理" 
                    tableName = "BaiduCai" 
                    tableConfig = tableConfig & "大类名称|BigClassName|2000|True[分]" 
                    tableConfig = tableConfig & "小类名称|SmallClassName|2000|True[分]" 
                    tableConfig = tableConfig & "子类名称|ThreeClassName|2000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|2000|True[分]" 
                    tableConfig = tableConfig & "百度网址|BaiduUrl|2000|True[分]" 
                    tableConfig = tableConfig & "网址|HttpUrl|2000|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|2000|True[分]" 
                elseIf s = "BaiduZhidaoMain" then
                    menuTitle = "百度知道Main" 
                elseIf s = "BaiduZhidaoData" then
                    menuTitle = "百度知道数据" 
                    tableName = "BaiduZhidaoData" 
                    thisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoData.Mdb""" 
                    tableConfig = tableConfig & "百度ID|ZhidaoID|2200|True[分]" 
                    tableConfig = tableConfig & "标题|Title|6000|True[分]" 
                    tableConfig = tableConfig & "回答数|ReplyNumb|2000|True[分]" 
                    tableConfig = tableConfig & "支持顶数|SupportNumb|2000|True[分]" 
                elseIf s = "BaiduZhidaoWordsData" then
                    menuTitle = "百度知道关键词数据" 
                    tableName = "BaiduZhidaoWordsData" 
                    thisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoWordsData.Mdb""" 
                    tableConfig = tableConfig & "关键词|Word|2000|True[分]" 

                elseIf s = "WebShell" then
                    menuTitle = "WebShell管理" 
                    tableName = "WebShell" 
                    tableConfig = tableConfig & "大类名称|BigClassName|1800|True[分]" 
                    tableConfig = tableConfig & "小马|Ev1|6500|True[分]" 
                    tableConfig = tableConfig & "域名|WebSite|3200|True[分]" 
                    tableConfig = tableConfig & "通过|Through|800|True[分]" 
                elseIf s = "WebServer" then
                    menuTitle = "WebServer网站服务器" 
                    tableName = "WebServer" 
                    tableConfig = tableConfig & "IP|IP|2000|True[分]" 
                    tableConfig = tableConfig & "组件|Assembly|2000|True[分]" 
                    tableConfig = tableConfig & "安全级别|AnQuan|1200|True[分]" 
                elseIf s = "QQ" then
                    menuTitle = "QQ管理" 
                    tableName = "QQ" 
                    thisMDBPath = "DefaultQQAccessPath & ""\QQ.Mdb""" 
                    tableConfig = tableConfig & "QQ|QQ|1500|True[分]" 
                    tableConfig = tableConfig & "QQ名称|QQName|2000|True[分]" 
                    tableConfig = tableConfig & "性别|QQSex|1000|True[分]" 
                    tableConfig = tableConfig & "年龄|Age|1000|True[分]" 
                    'TableConfig = TableConfig & "QQ年龄|QQAge|1000|True[分]"
                    tableConfig = tableConfig & "空间极限|SpacePurview|1000|True[分]" 
                    tableConfig = tableConfig & "相册数|PhotoNumb|1000|True[分]" 
                    tableConfig = tableConfig & "说说数|ShuaShuaNumb|1000|True[分]" 
                    tableConfig = tableConfig & "日志数|BlogNumb|1000|True[分]" 
                    'TableConfig = TableConfig & "首页显访问|HomeVisitors|1000|True[分]"
                    'TableConfig = TableConfig & "访问权限|SpacePurview|1200|True[分]"
                    tableConfig = tableConfig & "赞|QzoneZai|800|True[分]" 
                    tableConfig = tableConfig & "等级|QzoneLevel|800|True[分]" 
                    tableConfig = tableConfig & "积分|QzoneIntegral|800|True[分]" 

                elseIf s = "QQName" then
                    menuTitle = "QQ名称管理" 
                    tableName = "QQName" 
                    thisMDBPath = "DefaultQQAccessPath & ""\QQName.Mdb""" 
                    tableConfig = tableConfig & "QQ|QQ|1500|True[分]" 
                    tableConfig = tableConfig & "QQ名称|QQName|2000|True[分]" 
                elseIf s = "QQVisitors" then
                    menuTitle = "QQ访客管理" 
                    tableName = "QQVisitors" 
                    thisMDBPath = "DefaultQQAccessPath & ""\QQVisitors.Mdb""" 
                    tableConfig = tableConfig & "QQ|QQ|1500|True[分]" 
                    tableConfig = tableConfig & "访客QQ|FQQ|15000|True[分]" 


                elseIf s = "QQJournal" then
                    menuTitle = "QQ日志" 
                    tableName = "QQJournal" 
                    thisMDBPath = """SqlServer"""                                                   '本地SqlServer数据库

                    tableConfig = tableConfig & "QQ|QQ|1500|True[分]" 

                    tableConfig = tableConfig & "类型|cate|1500|True[分]" 
                    tableConfig = tableConfig & "日志ID|blogId|1500|True[分]" 
                    tableConfig = tableConfig & "标题|title|3000|True[分]" 
                    tableConfig = tableConfig & "日志ID|blogId|1500|True[分]" 
                    tableConfig = tableConfig & "发布时间|pubTime|1500|True[分]" 



                elseIf s = "BaiduZhidao" then
                    menuTitle = "百度知道管理" 
                    tableName = "BaiduZhidao" 
                    thisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidao.Mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|2000|True[分]" 
                    tableConfig = tableConfig & "提问标题|Title|4000|True[分]" 
                    tableConfig = tableConfig & "提问网址|HttpUrl|2000|True[分]" 
                    tableConfig = tableConfig & "发布者|Author|1800|True[分]" 
                    tableConfig = tableConfig & "回答数|ReplyNumb|1000|True[分]" 
                    tableConfig = tableConfig & "回答者|ReplyName|1800|True[分]" 
                    tableConfig = tableConfig & "采纳率|AdoptionRate|1000|True[分]" 
                    tableConfig = tableConfig & "支持数|SupportNumb|1000|True[分]" 
                    tableConfig = tableConfig & "浏览次数|BrowseNumb|1200|True[分]" 
                    tableConfig = tableConfig & "通过|Through|800|True[分]" 
                elseIf s = "BaiduZhidaoWords" then
                    menuTitle = "百度知道相关词管理" 
                    tableName = "BaiduZhidaoWords" 
                    thisMDBPath = "DefaultQQAccessPath & ""\BaiduZhidaoWords.Mdb""" 
                    tableConfig = tableConfig & "搜索词|SearchWord|3000|True[分]" 
                    tableConfig = tableConfig & "相关词|RelatedWord|3000|True[分]" 
                    tableConfig = tableConfig & "排序|nSort|800|True[分]" 

                elseIf s = "StatisticalData" then
                    menuTitle = "数据统计管理" 
                    tableName = "StatisticalData" 
                    thisMDBPath = "DefaultQQAccessPath & ""\StatisticalData.Mdb""" 
                    tableConfig = tableConfig & "大类名称|BigClassName|2000|True[分]" 
                    tableConfig = tableConfig & "标题|Title|4000|True[分]" 
                    tableConfig = tableConfig & "数值|CountNumb|2000|True[分]" 

                elseIf s = "CaiBaiduUrl" then
                    menuTitle = "采集百度网址" 
                elseIf s = "IEDebug" then
                    menuTitle = "IE调试" 
                elseIf s = "WebCrack" then
                    menuTitle = "网站破解" 
                elseIf s = "DownList" then
                    menuTitle = "批量下载" 

                elseIf s = "PseudoOriginal" then
                    menuTitle = "仿原创工具" 
                    call createFolder(rootFolder & "/Data") 
                    call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\默认字典.txt", rootFolder & "/DataDir/默认字典.txt") '默认字典
                    call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\字符替换字典.txt", rootFolder & "/DataDir/字符替换字典.txt") '字符替换字典
                    call CopyFile("\DataDir\VB模块\服务器\Dict字典未用\简体繁体转换.txt", rootFolder & "/DataDir/简体繁体转换.txt") '简体繁体转换
                elseIf s = "Scan" then
                    menuTitle = "Scan扫描" 
                elseIf s = "Qzone" then
                    menuTitle = "QQ空间" 
                elseIf s = "SinglePage" then
                    menuTitle = "网站单页制作" 
                elseIf s = "IEAutomatic" then
                    menuTitle = "IE自动操作" 
                    call CopyFile("\DataDir\VB模块\任务动作代码.ini", rootFolder & "任务动作代码.ini") '复制文件
                    call CopyFile("\DataDir\VB模块\问问关键词列表.ini", rootFolder & "问问关键词列表.ini") '复制文件
                    call CopyFile("\DataDir\VB模块\自动运行动作.ini", rootFolder & "自动运行动作.ini") '复制文件
                    call CopyFile("\DataDir\VB模块\网站后台登录字典.ini", rootFolder & "网站后台登录字典.ini") '复制文件

                elseIf s = "IEAutoHandle" then
                    menuTitle = "IE自动处理2015" 
                elseIf s = "AsynchDown" then
                    menuTitle = "测试自动下载" 

                elseIf s = "-" or left(s, 2) = "--" then
                    menuTitle = "-" 
                end if 
                'call echo("s",s)
                '排除分割线 顶部显示菜单名称
                if s <> "-" and left(s, 2) <> "--" then
                    toolsMenu = toolsMenu & "    ElseIf S = """ & menuTitle & """ Then" & vbCrLf 
                    toolsMenu = toolsMenu & "        " & s & ".Show: " & s & ".SetFocus" & vbCrLf 
                    '为加载窗体则加个配置
                    if tableConfig <> "" then

                        toolsMenu = toolsMenu & "        If " & s & ".Caption <> ""$表面板$"" Then Exit Sub" & vbCrLf '判断不重复加载
                        toolsMenu = toolsMenu & "        " & s & ".ThisFormName=""" & s & """" & vbCrLf '窗体名称
                        toolsMenu = toolsMenu & "        " & s & ".Caption=""" & menuTitle & """" & vbCrLf '显示窗体名称
                        toolsMenu = toolsMenu & "        " & s & ".FrameManage.Caption=""" & menuTitle & """" & vbCrLf '显示窗体框架名称
                        toolsMenu = toolsMenu & "        " & s & ".TableName=""" & tableName & """" & vbCrLf '表名称


                        if thisAccessPass <> "" then toolsMenu = toolsMenu & "        " & s & ".ThisAccessPass=""" & thisAccessPass & """" & vbCrLf '数据库密码
                        'If ThisMDBPath <> "" Then ToolsMenu=ToolsMenu & "        " & S & ".ThisMDBPath=handlePath(" & ThisMDBPath & ")" & vbCrlf                    '数据库路径
                        if thisMDBPath <> "" then toolsMenu = toolsMenu & "        " & s & ".ThisMDBPath=" & thisMDBPath & vbCrLf '数据库路径


                        toolsMenu = toolsMenu & "        " & s & ".TableConfig=""" & tableConfig & """" & vbCrLf 
                        toolsMenu = toolsMenu & "        " & s & ".TableConfig=""" & tableConfig & """" & vbCrLf 

                        toolsMenu = toolsMenu & "        Call " & s & ".FormLoad()" & vbCrLf 

                    end if 
                end if 
                '加载数据窗体
                if tableConfig <> "" then
                    formStrList = formStrList & vbCrLf & "Public " & s & " As Form"                 '表格窗体定义
                    formSetList = formSetList & vbCrLf & "Set " & s & "=New TableDialog"            '加载窗体
                end if 
                '这里设置下拉按钮放在哪里一组20150610
                if inStr("|AskDa|ZNRobot|JobWeb|JobCompany|FormJobCai|FrmTestDialog1|FrmTestDialog2|FrmDown|Ying|YingZheng|", "|" & s & "|") > 0 then
                    '工具菜单
                    if nTools2 = 0 then
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Frm.Tools2Menu(0).Caption = """ & menuTitle & """" & vbCrLf 
                    else
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Load Frm.Tools2Menu(" & nTools2 & "): Frm.Tools2Menu(" & nTools2 & ").Caption = """ & menuTitle & """" & vbCrLf 
                    end if 
                    nTools2 = nTools2 + 1 
                '网站制作分组20150610
                elseIf inStr("|UserMakeNav|SinglePage|FrmCrateNewHtmlPage|FileScan2015|FrameLayout|FrmCreateWebLable|HTML5CSS3|XyWebUp|MYProgram|FrameLayout|FrmCreateWebLable|FrmWebTemplateView|--2|FrmAutoAddWebDefaultInfo|FrmControlDialog|FrmSqlIn|FrmProgrameTransformation|WebShell|FrmEvScan|", "|" & s & "|") > 0 then
                    '工具菜单
                    if nTools3 = 0 then
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Frm.Tools3Menu(0).Caption = """ & menuTitle & """" & vbCrLf 
                    else
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Load Frm.Tools3Menu(" & nTools3 & "): Frm.Tools3Menu(" & nTools3 & ").Caption = """ & menuTitle & """" & vbCrLf 
                    end if 
                    nTools3 = nTools3 + 1 
                else
                    '工具菜单
                    if nTools = 0 then
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Frm.ToolsMenu(0).Caption = """ & menuTitle & """" & vbCrLf 
                    else
                        toolsClick = toolsClick & "    if checkAdminQuanXian(""" & menuTitle & """) then Load Frm.ToolsMenu(" & nTools & "): Frm.ToolsMenu(" & nTools & ").Caption = """ & menuTitle & """" & vbCrLf 
                    end if 
                    nTools = nTools + 1 
                end if 
            end if 

            '20150328
            if inStr("|WebManage|HTML5CSS3|FrmCrateNewWeb|FrmCrateNewHtmlPageOLD|", "|" & s & "|") > 0 then
                '拷贝Template文件夹
                call CopyFolder("\DataDir\VB模块\Template", rootFolder & "/Template")           '复制文件夹
            end if 

            '默认打开窗体  排除分割线
            if s <> "MDIForm1" and s <> "-" and left(s, 2) <> "--" then
                openFrom = openFrom & "    ElseIf Tag = LCase(""" & s & """) Then" & vbCrLf & "        " & s & ".Show: " & s & ".SetFocus" & vbCrLf 
            end if 
            '排除分割线，并且数据配置表为空
            if s <> "-" and left(s, 2) <> "--" and tableConfig = "" then
                s = s & ".frm" 
                if checkFile("/DataDir/VB模块/" & s) = false then
                    call eerr("没有窗体", "/DataDir/VB模块/" & s) 
                end if 
                frmList = frmList & "Form=zForm\" & s & vbCrLf 
                call CopyFile("/DataDir/VB模块/" & s, zFormFolder & s) 
                '拷贝数数
                s = replace(s, ".frm", ".frx") 
                call CopyFile("/DataDir/VB模块/" & s, zFormFolder & s) 
            end if 
        end if 
    next 

    '异步下载模块
    ctlFileList = ctlFileList & "UserControl=zForm\Downloader.ctl" & vbCrLf 
    call CopyFile("\DataDir\VB模块\Downloader.ctl", zFormFolder & "Downloader.ctl") '复制文件
    call CopyFile("\DataDir\VB模块\Downloader.ctx", zFormFolder & "Downloader.ctx") '复制文件
    '自定义参数模块
    ctlFileList = ctlFileList & "UserControl=zForm\ControlParam.ctl" & vbCrLf 
    call CopyFile("\DataDir\VB模块\ControlParam.ctl", zFormFolder & "ControlParam.ctl") '复制文件
    call CopyFile("\DataDir\VB模块\ControlParam.ctx", zFormFolder & "ControlParam.ctx") '复制文件
    '自定义网址模块
    ctlFileList = ctlFileList & "UserControl=zForm\ControlUrl.ctl" & vbCrLf 
    call CopyFile("\DataDir\VB模块\ControlUrl.ctl", zFormFolder & "ControlUrl.ctl") '复制文件
    call CopyFile("\DataDir\VB模块\ControlUrl.ctx", zFormFolder & "ControlUrl.ctx") '复制文件



    content = replace(content, "Manage.exe", "网站制作软件.exe") 
    'Content = Replace(Content,"QQ313801120","www.maiside.net")        '替换成相应联系方式20150316
    'Call Eerr("FrmList",FrmList)
    content = replace(content, "{Frm列表}", frmList & ctlFileList) 
    call CreateFile(rootFolder & "Manage.vbp", content) 
    call echo("FormStrList", formStrList) 
    'MDIForm1.frm处理
    path = "/DataDir/VB模块/MDIForm1.frm" 
    content = getFText(path) 
    content = replace(content, "{工具}", toolsMenu) 
    content = replace(content, "{工具导航}", toolsClick) 
    content = replace(content, "{默认打开窗体}", openFrom) 
    content = replace(content, "{加载窗体}", formSetList) 
    content = replace(content, "Version1.0", "网站制作软件")                        'SoftVersion版本号
    '伪原创按钮处理
    if inStr(frmListStr, "PseudoOriginal") = 0 then
        'Call CreateFile("测试文件一会删除.txt",Content)            '这是因为MDIForm1.frm文件被替换的原因
        'Call Rw(Content):Response.End()
        tempStr = strCut(content, vbCrLf & "'伪原创", "End Sub", 1) 
        content = replace(content, tempStr, vbCrLf) 
    end if 


    if yesSortPlug = "OK" then
        content = replace(replace(content, "[SortPlug开始]", ""), "[SortPlug结束]", "") 
    else
        tempStr = strCut(content, "[SortPlug开始]", "[SortPlug结束]", 1) 
        content = replace(content, tempStr, "") 
    end if 
    call CreateFile(zFormFolder & "MDIForm1.frm", content) 

    '复制基础模块
    splStr = split("VB_Img.bas|CloseEXE.bas|Ying1.bas|Ying2.bas|VB_Config.bas|API.bas|VB_Conn.bas|WinMain.bas|Md5.bas|VB_Common.bas|Tp.bas|FindEXE.bas|Input.bas|Rwini.bas|clsWaitableTimer.cls|GPS.cls|Request.cls|AspCode.cls|NetWorkClass.cls|modExcel.bas|VB_ImportAndExportWeb.bas|WebSiteGetIP.bas|cx_extractor.bas|PingIP.bas|RT.bas|DomainNameToIP.bas|VB_Computer.bas|VB_CaiClass.bas|VB_VBControl.bas|ClassRichTextBox.cls|VBJSScript.bas|AspToPhpClass.cls|AspActionClass.cls|ControlTemplateStyle.ctl|VB_ModuleFunction.bas|VB_CallCommand.bas", "|") 
    for each s in splStr
        if s <> "" then
            call CopyFile("/DataDir/VB模块/" & s, zFormFolder & s) 
        'BasList=BasList & "Module="& FormName &"; zForm\"& FormName &".bas" & vbCrlf
        end if 
    next 
    call CopyFolder("/DataDir/VB模块/Images", rootFolder & "Images") 
    call CopyFolder("/DataDir/VB模块/Edit", rootFolder & "Edit") 
    call CopyFolder("/DataDir/VB模块/Config", rootFolder & "Config") 
    'Call CopyFolder("/DataDir/VB模块/dll", RootFolder & "dll")


    call CopyFile("/DataDir/VB模块/SysInfo.ini", rootFolder & "SysInfo.ini") 


    '定义全局窗体
    path = zFormFolder & "/VB_Config.bas" 
    content = getFText(path) 
    s = strCut(content, "'{$窗体开始$}", "'{$窗体结束$}", 1) 
    content = replace(content, s, "'{$窗体开始$}" & formStrList & vbCrLf & "'{$窗体结束$}") 
    '替换版本号
    s = strCut(content, "Public Const SoftVersion = """, """ '软件版本", 1) 
    content = replace(content, s, "Public Const SoftVersion = """ & softVersion & """ '软件版本") 
    call CreateFile(path, content) 
    if 1 = 2 then
        dataFolder = rootFolder & "DataDir/" 
        call createFolder(dataFolder) 
        call createFolder(dataFolder & "Log")                                           '日志文件夹

        call CopyFile("/DataDir/VB模块/最近打开文档.ini", dataFolder & "最近打开文档.ini") 
        'Call CopyFolder("/DataDir/VB模块/Eval", RootFolder & "Eval")
        call CopyFile("/DataDir/VB模块/sql.txt", dataFolder & "sql.txt") 
        call CopyFile("/DataDir/VB模块/Desk.jpg", dataFolder & "Desk.jpg") 
    end if 
end sub 
%> 


