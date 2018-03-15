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
<% 
'======================================
'数据库表操作
'======================================


'给表配置
function showClassName(table, content, ascDescStr, printStr, TDWidth)
    dim splStr, splxx, s, c, n, i 
    dim IDStr, QQStr, QQNameStr, handleStr, visitorsListStr, spacePurviewStr, onLineTimeStr, QQSexStr, QQAgeStr, ageStr 
    dim bigClassNameStr, smallClassNameStr, countNumbStr, QQNumbStr 
    dim categoryStr, titleStr, blogidStr, messageContentStr, fileNameStr, folderNameStr, dirPathStr 
    dim photoPreStr, photoNameStr, bigPhotoUrlStr, photoIDStr, upDateTimeStr, photoUrlStr, contentStr, messageQQStr, messageUpDateTimeStr 
    dim addTimeStr, searchValueStr, searchTitleStr, searchContentStr, searchTimeStr, searchSortStr, webSiteWebHttpUrlStr 
    dim userNameStr, priceStr, farePriceStr, provinceStr, cityStr, descriptionStr, replyNumbStr, replyNameStr, QQTouXiangStr 
    dim zipSizeStr, windowsStr, sortStr, photoShowTypeStr, workTitleStr, companyStr, diQuStr, workEXPStr, QQBeiZhuNameStr 
    dim authorStr, renQiStr, webTitleStr, webUrlStr, dayNumbStr, ciStr, shiStr, poJieStr, poJiePhotoStr, downNumbStr, ENStr, CNStr 
    dim httpUrlStr, playTimeStr, playNumbStr, postingsTimeStr, navNameStr, threeClassNameStr, clickNumbStr, mutualFriendStr, visitorsNumbStr 
    dim supportNumbStr, createDateStr, throughStr, searchWordStr, xiangGuanWordStr, adoptionRateStr, browseNumbStr, emailStr, IPStr 
    dim addDateTimeStr, filePathStr, myMoneyStr, replyTimeStr, askTypeStr, askIdStr, webSiteStr, pageNumberStr, rankingStr, includedTimeStr 
    dim searchUrlStr, cateStr, pubTimeStr 
    IDStr = "选择|42|," 
    cateStr = "分类|42|Cate|分类重复," 
    pubTimeStr = "发布时间|100|PubTime|日志发布时间重复," 
    QQStr = "QQ号码|70|QQ|QQ号码重复|QQ号码搜索," 
    QQNameStr = "QQ名称|200|QQName|QQ名称重复|QQ名称搜索," 
    handleStr = "操作||详细|||修改," 
    visitorsListStr = "访客列表|80|VisitorsList||访客列表搜索," 
    visitorsNumbStr = "访客数量|80|VisitorsNumb||访客数量搜索," 

    spacePurviewStr = "空间权限|80|SpacePurview|空间权限重复|空间权限搜索," 
    QQSexStr = "QQ性别|60|QQSex|QQ性别重复|QQ性别搜索," 
    QQAgeStr = "QQ年龄|60|QQAge|QQ年龄重复|QQ年龄搜索," 
    ageStr = "年龄|60|Age|年龄重复|年龄搜索," 
    '关系
    bigClassNameStr = "大类名称|100|BigClassName|大类名称重复|大类名称搜索," 
    smallClassNameStr = "小类名称|100|SmallClassName|小类名称重复|小类名称搜索," 
    threeClassNameStr = "子类名称|100|ThreeClassName|子类名称重复|子类名称搜索," 

    searchWordStr = "搜索词|200|BigClassName|搜索词重复|搜索词搜索," 
    xiangGuanWordStr = "相关词|200|SmallClassName|相关词重复|相关词搜索," 

    countNumbStr = "好友度|60|CountNumb,"                                           '注意这里 下面 CountNumbStr总数给代替了
    QQNumbStr = "存在数|60|QQNumb," 
    '日志
    onLineTimeStr = "登录时间|116|OnLineTime|登录时间重复|登录时间搜索," 
    categoryStr = "日志类型|80|Category|日志类型重复|日志类型搜索," 
    titleStr = "标题|204|Title|标题重复|标题搜索," 
    blogidStr = "日志ID|70|Blogid|日志ID重复," 
    '相册
    photoPreStr = "缩览图|100|PhotoPre," 
    photoNameStr = "图片名称|100|PhotoName|图片名称重复|图片名称搜索," 
    bigPhotoUrlStr = "图片缩览图|100|BigPhotoUrl||图片缩览图搜索," 
    photoIDStr = "图片ID|100|PhotoID|图片ID重复|图片ID搜索," 
    messageQQStr = "留言者QQ|70|MessageQQ|留言者QQ重复|留言者QQ搜索," 
    messageUpDateTimeStr = "留言时间|120|UpDateTime|留言时间重复|留言时间搜索," 
    photoUrlStr = "留言缩览图|100|PhotoUrl," 
    messageContentStr = "留言内容||Content||留言内容搜索," 
    photoShowTypeStr = "存在大小图|80|PhotoShowType|存在大小图重复|存在大小图搜索," 
    '相同
    countNumbStr = "总数|60|CountNumb," 
    addTimeStr = "添加时间|120|AddTime|添加时间重复|添加时间搜索," 
    '搜索引擎
    searchValueStr = "搜索值|100|SearchValue|搜索值重复|搜索值搜索," 
    searchTitleStr = "搜索标题|200|SearchTitle|搜索标题重复|搜索标题搜索," 
    searchContentStr = "搜索内容|200|SearchContent||搜索内容搜索," 
    searchTimeStr = "搜索时间|80|SearchTime|搜索时间重复|搜索时间搜索," 
    searchSortStr = "搜索排序|80|SearchSort|搜索排序重复|搜索排序搜索," 
    webSiteWebHttpUrlStr = "域名/网址|200|域名/网址|域名重复|网址搜索/域名搜索," 
    '淘宝
    userNameStr = "用户|100|UserName|用户重复|用户搜索," 
    priceStr = "价格|60|Price|价格重复|价格搜索," 
    farePriceStr = "运费|60|FarePrice|运费重复|运费搜索," 
    provinceStr = "省份|40|Province|省份重复|省份搜索," 
    cityStr = "城市|40|City|城市重复|城市搜索," 
    descriptionStr = "描述值|60|Description|描述值重复|描述值搜索," 
    replyNumbStr = "评论数|60|ReplyNumb|评论数重复|评论数搜索," 
    replyNameStr = "评论者|60|ReplyName|评论者重复|评论者搜索," 
    clickNumbStr = "点击数|60|ClickNumb|点击数重复|点击数搜索," 
    '产品
    contentStr = "备注|200|Content||备注搜索,"                                      '这样注意200可去掉，让它自动
    upDateTimeStr = "上传时间|80|UpDateTime|上传时间重复|上传时间搜索," 
    zipSizeStr = "文件大小|60|ZipSize|文件大小重复," 
    windowsStr = "运行平台|80|Windows|运行平台重复|运行平台搜索," 
    downNumbStr = "下载数|80|DownNumb|下载数重复|下载数搜索," 
    'Words
    sortStr = "排序|60|Sort|排序重复|排序搜索," 
    '招聘
    workTitleStr = "招聘职位|100|WorkTitle|招聘职位重复|招聘职位搜索," 
    companyStr = "公司名称|100|Company|公司名称重复|公司名称搜索," 
    diQuStr = "地区|40|DiQu|地区重复|地区搜索," 
    workEXPStr = "工作经验|90|WorkEXP|工作经验重复|工作经验搜索," 
    'QQ备注
    QQBeiZhuNameStr = "QQ备注名|90|QQBeiZhuName|QQ备注名重复|QQ备注名搜索," 
    QQTouXiangStr = "QQ头像|100|QQ头像||," 
    'Flash163SWF
    authorStr = "作者|90|Author|作者重复|作者搜索," 
    renQiStr = "人气|90|RenQi|人气重复|人气搜索," 
    'Web
    webTitleStr = "网站标题|200|WebTitle|网站标题重复|网站标题搜索," 
    webUrlStr = "网站地址|200|WebUrl|网站地址重复|网站地址搜索," 
    dayNumbStr = "星期数|200|DayNumb|星期数重复|星期数搜索," 
    'ce
    ciStr = "单词|100|Ci|单词重复|单词搜索," 
    shiStr = "英文解释||Shi|英文解释重复|英文解释搜索," 
    'QQ图片破解
    poJieStr = "破解值|80|PoJie|破解值重复|破解值搜索," 
    poJiePhotoStr = "破解图片|80|PoJiePhoto," 
    'CnEnWords
    ENStr = "英文|120|EN|英文重复|英文搜索," 
    CNStr = "中文|200|CN|中文重复|中文搜索," 
    '文件（夹）目录
    fileNameStr = "文件|150|FileName|文件重复|文件搜索," 
    folderNameStr = "文件夹|120|FolderName|文件夹重复|文件夹搜索," 
    dirPathStr = "目录路径||DirPath|目录路径重复|目录路径搜索," 
    '视频
    httpUrlStr = "网址|200|HttpUrl|网址重复|网址搜索," 
    playTimeStr = "播放时间|80|PlayTime|播放时间重复|播放时间搜索," 
    playNumbStr = "播放次数|80|PlayNumb|播放次数重复|播放次数搜索," 
    postingsTimeStr = "发布时间|100|PostingsTime|发布时间重复|发布时间搜索," 
    '其它
    navNameStr = "导航名称|100|cname|导航名称重复|导航名称搜索," 
    mutualFriendStr = "共同好友|60|MutualFriend|共同好友重复|共同好友搜索," 
    supportNumbStr = "支持数|80|SupportNumb|支持数重复|支持数搜索," 
    createDateStr = "提交日期|80|CreateDate|提交日期重复|提交日期搜索," 
    throughStr = "是否通过|80|Through|是否通过重复|是否通过搜索," 
    adoptionRateStr = "采纳率|80|AdoptionRate|采纳率重复|采纳率搜索," 
    browseNumbStr = "浏览次数|80|BrowseNumb|浏览次数重复|浏览次数搜索," 
    emailStr = "邮箱|120|Email|邮箱重复|邮箱搜索," 
    IPStr = "IP|120|IP|IP重复|IP搜索," 
    addDateTimeStr = "添加时间|120|AddDateTime|添加时间重复|添加时间搜索," 
    filePathStr = "文件路径|120|FilePath|文件路径重复|文件路径搜索," 
    myMoneyStr = "用户积分|120|MyMoney|用户积分重复|用户积分搜索," 
    replyTimeStr = "提问时间|120|ReplyTime|提问时间重复|提问时间搜索," 
    askTypeStr = "问问类型|120|AskType|问问类型重复|问问类型搜索," 
    askIdStr = "问问ID|120|AskId|问问ID重复|问问ID搜索," 
    webSiteStr = "域名|120|WebSite|域名重复|域名搜索," 
    pageNumberStr = "页数|120|PageNumber|页数重|复页数搜索," 
    rankingStr = "排名|120|Ranking|排名重复|排名搜索," 
    includedTimeStr = "收录时间|120|IncludedTime|收录时间重复|收录时间搜索," 
    searchUrlStr = "搜索引擎网址|120|SearchUrl|搜索引擎网址重复|搜索引擎网址搜索," 

    '【产品表】
    if checkStr(table, "^Product$") then
        content = IDStr & bigClassNameStr & smallClassNameStr & threeClassNameStr & titleStr & contentStr & upDateTimeStr & handleStr 
        ascDescStr = "点击数" 
        printStr = "标题|内容" 
    '【QQ名称表】
    elseIf checkStr(table, "^QQNameTable$") then
        content = IDStr & QQStr & QQNameStr & contentStr & handleStr 
        ascDescStr = "QQ号码|QQ名称" 
        printStr = "QQ号码|QQ名称|QQ名称和号码" 
    '【QQ备注表】
    elseIf checkStr(table, "^QQBeiZhu$") then
        content = IDStr & QQStr & QQBeiZhuNameStr & handleStr 
        ascDescStr = "QQ号码|QQ名称" 
        printStr = "QQ号码|QQ备注名|QQ名称和号码" 
    '【QQ表】
    elseIf checkStr(table, "^QQ$") then
        if rq("Style") = "2" then
            content = IDStr & QQStr & QQNameStr & QQSexStr & ageStr & upDateTimeStr & QQAgeStr & handleStr 
            ascDescStr = "QQ号码|QQ名称|QQ性别|年龄|上传时间|QQ年龄" 
            printStr = "QQ号码|QQ访客|QQ和访客" 
        else
            content = IDStr & bigClassNameStr & QQStr & QQNameStr & visitorsListStr & spacePurviewStr & onLineTimeStr & mutualFriendStr & visitorsNumbStr & handleStr 
            ascDescStr = "QQ号码|QQ名称|空间权限|登录时间|访客数量" 
            printStr = "QQ号码|QQ访客|QQ和访客" 
        end if 
    '【关系QQ表】
    elseIf checkStr(table, "^GuanXiQQ$") then
        content = IDStr & bigClassNameStr & smallClassNameStr & QQStr & QQNameStr & countNumbStr & QQNumbStr & handleStr 
        ascDescStr = "好友度|存在数|好友度＋存在数|[联合查询]" 
        printStr = "QQ号码|QQ名称|QQ名称和号码" 
    '【日志表】
    elseIf checkStr(table, "^DailyRcord$") then
        content = IDStr & QQStr & categoryStr & titleStr & replyNumbStr & blogidStr & handleStr 
        ascDescStr = "QQ号码|评论数" 
        printStr = "QQ号码|日志标题|日志地址" 
    '【QQ相册类型表】
    elseIf checkStr(table, "^QQPhotoClass$") then
        content = IDStr & bigClassNameStr & QQStr & photoPreStr & photoShowTypeStr & handleStr 
        ascDescStr = "QQ号码" 
        printStr = "大类名称|QQ号码|缩览图" 
    '【QQ相册信息表】
    elseIf checkStr(table, "^QQPhotoProduct$") then
        content = IDStr & QQStr & photoNameStr & bigPhotoUrlStr & photoIDStr & photoShowTypeStr & handleStr 
        ascDescStr = "QQ号码" 
        printStr = "QQ号码|相册小图|相册大图" 
    '【QQ相册留言表】
    elseIf checkStr(table, "^QQPhotoMessage$") then
        content = IDStr & messageQQStr & QQStr & QQNameStr & messageUpDateTimeStr & photoUrlStr & photoShowTypeStr & messageContentStr 
        ascDescStr = "QQ号码|留言时间" 
        printStr = "QQ号码|留言图片|留言内容" 
    '【QQ在线表】
    elseIf checkStr(table, "^QQOnLine") then
        content = IDStr & QQStr & onLineTimeStr & handleStr 
        ascDescStr = "QQ号码|登录时间" 
        printStr = "QQ号码" 
    '【QQ日志】
    elseIf checkStr(table, "^QQJournal") then
        content = IDStr & QQStr & cateStr & blogidStr & titleStr & pubTimeStr & handleStr 
        ascDescStr = "QQ号码|登录时间" 
        printStr = "QQ号码" 
    '【相同表】
    elseIf checkStr(table, "^WithTable$") then
        content = IDStr & bigClassNameStr & smallClassNameStr & titleStr & countNumbStr & handleStr 
        ascDescStr = "总数" 
        printStr = "标题|分析出图" 
    '【QQ暂存表】
    elseIf checkStr(table, "^QQTemp$") then
        content = IDStr & QQStr & addTimeStr & handleStr 
        ascDescStr = "QQ号码|添加时间" 
        printStr = "QQ号码" 
    '【搜索引擎表】
    elseIf checkStr(table, "^SearchEngine$") then
        content = IDStr & searchValueStr & searchTitleStr & searchTimeStr & searchSortStr & searchContentStr & webSiteWebHttpUrlStr & handleStr 
        ascDescStr = "QQ号码|添加时间" 
        printStr = "搜索标题|搜索网址|域名" 
    '【淘宝表】
    elseIf checkStr(table, "^TaoBaoProduct$") then
        content = IDStr & searchValueStr & userNameStr & priceStr & farePriceStr & provinceStr & cityStr & descriptionStr & replyNumbStr 
        ascDescStr = "价格|运费|描述值|评论" 
        printStr = "用户|标题|价格" 
    '【文档表】
    elseIf checkStr(table, "^Word$") then
        content = IDStr & titleStr & addTimeStr & sortStr & handleStr 
        ascDescStr = "添加时间|排序" 
        printStr = "标题" 
    '【51招聘表】
    elseIf checkStr(table, "^JOB_51Job$") then
        content = IDStr & searchValueStr & workTitleStr & companyStr & cityStr & diQuStr & upDateTimeStr & workEXPStr & handleStr 
        ascDescStr = "上传时间|排序" 
        printStr = "招聘职位|公司名称|职位简介|招聘职位和职位简介" 
    '【？？？表】
    elseIf checkStr(table, "^QQTouXiang$") then
        content = IDStr & QQTouXiangStr & QQStr & photoShowTypeStr & handleStr 
        ascDescStr = "添加时间|排序" 
        printStr = "标题" 
    '【大活动漫表】
    elseIf checkStr(table, "^Flash163SWF$") then
        content = IDStr & titleStr & authorStr & renQiStr & photoShowTypeStr & handleStr 
        ascDescStr = "标题|人气|作者" 
        printStr = "标题|图片地址|SWF地址|下载地址" 
    '【QQ访客表】
    elseIf checkStr(table, "^QQVisitorsStat$") then
        content = IDStr & QQStr & QQNameStr & countNumbStr & handleStr 
        ascDescStr = "QQ号码|QQ名称|总数" 
        printStr = "QQ号码|QQ名称|QQ名称和号码" 
    '【暂存网址表】
    elseIf checkStr(table, "^TempUrl$") then
        content = IDStr & folderNameStr & fileNameStr & dirPathStr & handleStr 
        ascDescStr = "文件|文件夹|目录路径" 
        printStr = "文件|文件夹|目录路径" 
    '【视频表】
    elseIf checkStr(table, "^Video$") then
        content = IDStr & titleStr & httpUrlStr & playTimeStr & playNumbStr & postingsTimeStr & authorStr & handleStr 
        ascDescStr = "播放时间|播放次数|发布时间" 
        printStr = "视频图片|视频标题|视频网址|视频标题加网址" 
    '【Web表】外部MDB表需要
    elseIf checkStr(table, "^Web$") then
        content = IDStr & webTitleStr & webUrlStr & dayNumbStr & handleStr 
        ascDescStr = "星期数" 
        printStr = "网站标题|网站地址|网站标题加地址" 
    '【单词表】外部MDB表需要
    elseIf checkStr(table, "^ce$") then
        content = IDStr & ciStr & shiStr & handleStr 
        ascDescStr = "单词" 
        printStr = "单词|英文解释|单词加英文解释" 
    '【验证码表】外部MDB表需要
    elseIf checkStr(table, "^QQYingZheng$") then
        content = IDStr & titleStr & poJieStr & poJiePhotoStr & handleStr 
        ascDescStr = "破解值" 
        printStr = "破解值" 
    '【中英表】外部MDB表需要
    elseIf checkStr(table, "^CnEnWords$") then
        content = IDStr & ENStr & CNStr & sortStr & handleStr 
        ascDescStr = "排序" 
        printStr = "中英对照" 
    '【类表】外部MDB表需要
    elseIf checkStr(table, "^class$") then
        content = IDStr & navNameStr & httpUrlStr & handleStr 
        ascDescStr = "导航名称" 
        printStr = "导航名称" 
    '【类表】QQ名称表
    elseIf checkStr(table, "^qqname$") then
        content = IDStr & QQStr & QQNameStr & handleStr 
        ascDescStr = "QQ号码|QQ名称" 
        printStr = "QQ号码|QQ名称" 
    '【类表】百度知道表
    elseIf checkStr(table, "^baiduzhidao$") then
        content = IDStr & bigClassNameStr & smallClassNameStr & titleStr & httpUrlStr & authorStr & replyNumbStr & replyNameStr & adoptionRateStr & supportNumbStr & createDateStr & browseNumbStr & throughStr & handleStr 
        ascDescStr = "提交日期|评论数|浏览次数|是否通过" 
        printStr = "|" 
    '【类表】百度知道相关词表
    elseIf checkStr(table, "^BaiduZhidaoWords$") then
        content = IDStr & searchWordStr & xiangGuanWordStr & handleStr 
        ascDescStr = "|" 
        printStr = "|" 
    '【类表】数据统计表
    elseIf checkStr(table, "^StatisticalData$") then
        content = IDStr & bigClassNameStr & titleStr & countNumbStr & handleStr 
        ascDescStr = "总数" 
        printStr = "|" 
    '【类表】软件任务表
    elseIf checkStr(table, "^SoftTask$") then
        content = IDStr & bigClassNameStr & emailStr & contentStr & upDateTimeStr & filePathStr & throughStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 
    '【类表】会员表
    elseIf checkStr(table, "^Member$") then
        content = IDStr & userNameStr & myMoneyStr & emailStr & IPStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 
    '【类表】会员日志表
    elseIf checkStr(table, "^UserLog$") then
        content = IDStr & bigClassNameStr & contentStr & IPStr & addDateTimeStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 
    '【类表】反馈表
    elseIf checkStr(table, "^Feedback$") then
        content = IDStr & emailStr & contentStr & IPStr & addDateTimeStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 
    '【类表】关键词表
    elseIf checkStr(table, "^Words$") then
        content = IDStr & bigClassNameStr & emailStr & titleStr & IPStr & addDateTimeStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 
    '【类表】问问表
    elseIf checkStr(table, "^Ask$") then
        content = IDStr & askTypeStr & bigClassNameStr & askIdStr & titleStr & replyNumbStr & replyTimeStr & handleStr 
        ascDescStr = "提问时间" 
        printStr = "|" 
    '【类表】搜索引擎收录词
    elseIf checkStr(table, "^SearchEngineInfoList$") then
        content = IDStr & bigClassNameStr & webSiteStr & titleStr & pageNumberStr & rankingStr & includedTimeStr & searchUrlStr & handleStr 
        ascDescStr = "" 
        printStr = "|" 

    '未配置
    else
        call eerr("【" & table & "】表", "未配置，已停止。注意：【请到 Fun_Table.Asp 页找到 ShowClassName 函数修改】") 

    end if 
    '循环获得标题列表
    splStr = split(content, ",") 
    n = 1 
    for each s in splStr
        if s <> "" then
            splxx = split(s & "|||", "|") 
            TDWidth(n) = splxx(1) 
            c = c & "<td width=""" & splxx(1) & """ align=""center"" bgcolor=""#FFFFFF"" class=""B14"">" & splxx(0) & "</td>" & vbCrLf 
            n = n + 1 
        end if 
    next 
    showClassName = c 
end function 


'搜索类型 加 多个查询 Or 与 And
function handleSearchSQL(searchType, searchValue)
    dim addSql, seectField, svalueType 
    '判断是否为等于搜索
    if left(searchValue, 1) = """" and right(searchValue, 1) = """" then
        searchValue = left(searchValue, len(searchValue) - 1) 
        searchValue = right(searchValue, len(searchValue) - 1) 
        svalueType = "真" 
    end if 
    if searchValue <> "" then
        if searchType = "QQ号码搜索" then
            seectField = "QQ" 
        elseIf searchType = "QQ名称搜索" then
            seectField = "QQName" 
        elseIf searchType = "访客列表搜索" then
            seectField = "VisitorsList" 
        elseIf searchType = "大类名称搜索" then
            seectField = "BigClassName" 
        elseIf searchType = "小类名称搜索" then
            seectField = "SmallClassName" 
        elseIf searchType = "图片名称搜索" then
            seectField = "PhotoName" 
        elseIf searchType = "留言者QQ搜索" then
            seectField = "MessageQQ" 
        elseIf searchType = "登录时间搜索" then
            seectField = "OnLineTime" 
        elseIf searchType = "标题搜索" then
            seectField = "Title" 
        elseIf searchType = "添加时间搜索" then
            seectField = "AddTime" 
        elseIf searchType = "搜索值搜索" then
            seectField = "SearchValue" 
        elseIf searchType = "搜索标题搜索" then
            seectField = "SearchTitle" 
        elseIf searchType = "搜索时间搜索" then
            seectField = "SearchTime" 
        elseIf searchType = "搜索排序搜索" then
            seectField = "SearchSort" 
        elseIf searchType = "搜索内容搜索" then
            seectField = "SearchContent" 
        elseIf searchType = "搜索网址搜索" then
            seectField = "SearchHttpUrl" 
        elseIf searchType = "网址搜索" then
            seectField = "HttpUrl" 
        elseIf searchType = "域名搜索" then
            seectField = "WebSite" 
        elseIf searchType = "备注搜索" then
            seectField = "Content" 
        elseIf searchType = "上传时间搜索" then
            seectField = "UpDateTime" 
        elseIf searchType = "运行平台搜索" then
            seectField = "Windows" 
        elseIf searchType = "存在大小图搜索" then
            seectField = "PhotoShowType" 
        elseIf searchType = "图片ID搜索" then
            seectField = "PhotoID" 
        elseIf searchType = "公司名称搜索" then
            seectField = "Company" 
        elseIf searchType = "招聘职位搜索" then
            seectField = "WorkTitle" 
        elseIf searchType = "网站标题搜索" then
            seectField = "WebTitle" 
        elseIf searchType = "网站地址搜索" then
            seectField = "WebUrl" 
        elseIf searchType = "星期数搜索" then
            seectField = "DayNumb" 
        elseIf searchType = "单词搜索" then
            seectField = "Ci" 
        elseIf searchType = "英文解释搜索" then
            seectField = "Shi" 
        elseIf searchType = "图片缩览图搜索" then
            seectField = "BigPhotoUrl" 
        elseIf searchType = "QQ性别搜索" then
            seectField = "QQSex" 
        elseIf searchType = "QQ年龄搜索" then
            seectField = "QQAge" 
        elseIf searchType = "破解值搜索" then
            seectField = "PoJie" 
        elseIf searchType = "下载数搜索" then
            seectField = "DownNumb" 
        elseIf searchType = "评论数搜索" then
            seectField = "ReplyNumb" 
        elseIf searchType = "日志类型搜索" then
            seectField = "Category" 
        elseIf searchType = "英文搜索" then
            seectField = "EN" 
        elseIf searchType = "中文搜索" then
            seectField = "CN" 
        elseIf searchType = "排序搜索" then
            seectField = "Sort" 
        elseIf searchType = "文件搜索" then
            seectField = "FileName" 
        elseIf searchType = "文件夹搜索" then
            seectField = "FolderName" 
        elseIf searchType = "目录路径搜索" then
            seectField = "DirPath" 
        elseIf searchType = "播放时间搜索" then
            seectField = "PlayTime" 
        elseIf searchType = "播放次数搜索" then
            seectField = "PlayNumb" 
        elseIf searchType = "发布时间搜索" then
            seectField = "PostingsTime" 
        elseIf searchType = "作者搜索" then
            seectField = "Author" 
        elseIf searchType = "搜索词搜索" then
            seectField = "BigClassName" 
        elseIf searchType = "相关词搜索" then
            seectField = "SmallClassName" 
        elseIf searchType = "采纳率搜索" then
            seectField = "AdoptionRate" 
        elseIf searchType = "支持数搜索" then
            seectField = "SupportNumb" 
        elseIf searchType = "提交日期搜索" then
            seectField = "CreateDate" 
        elseIf searchType = "浏览次数搜索" then
            seectField = "BrowseNumb" 
        elseIf searchType = "是否通过搜索" then
            seectField = "Through" 
        elseIf searchType = "提问时间搜索" then
            seectField = "ReplyTime" 
        elseIf searchType = "问问ID搜索" then
            seectField = "AskId" 



        elseIf searchType <> "" then
            call echo(searchType, "搜索不存在") 
        end if 


        '获得SQL
        if seectField <> "" then
            '为SqlServer数据库 字段二有需要转换为字符类型
            if MDBPath = "" then
                '添加于2013,10,24  让字段为数字类型转为字符类型，要不然搜索有问题，但不是数字类型不要用，否则就搜索不到全部内容
                call OpenConn() 
                if conn.execute("Select " & seectField & " From [" & table & "]")(0).type = 3 then
                    seectField = "Convert(VarChar, " & seectField & ")" 
                end if 
            end if 
            addSql = "Where " & seectField & " Like '%" & searchValue & "%'" 
            '高级搜索
            if svalueType = "真" then
                addSql = "Where " & seectField & " = '" & searchValue & "'" 
            end if 
        end if 
        '多个查询 Or 与 And
        addSql = orAndSearch(addSql, seectField, searchValue) 
    end if 
    handleSearchSQL = addSql 
end function 

'处理排序
function handleAscDesc(addSql, ascDescType)
    if ascDescType = "" then
        handleAscDesc = addSql 
        exit function 
    end if 

    if ascDescType = "QQ号码正排序" then
        addSql = addSql & "   Order By Convert(Money, QQ) Asc" 
    elseIf ascDescType = "QQ号码倒排序" then
        addSql = addSql & "   Order By Convert(Money, QQ) Desc" 

    elseIf ascDescType = "QQ名称正排序" then
        addSql = addSql & "   Order By QQName Asc" 
    elseIf ascDescType = "QQ名称倒排序" then
        addSql = addSql & "   Order By QQName Desc" 

    elseIf ascDescType = "备注为空" then
        if inStr(addSql, "Where ") > 0 then addSql = addSql & "   And" else addSql = addSql & "   Where" 
        addSql = addSql & " Content Is Null" 
    elseIf ascDescType = "备注不为空" then
        if inStr(addSql, "Where ") > 0 then addSql = addSql & "   And" else addSql = addSql & "   Where" 
        addSql = addSql & " Content Is Not Null" 

    elseIf ascDescType = "空间权限正排序" then

        addSql = addSql & "   Order By SpacePurview Asc" 
    elseIf ascDescType = "空间权限倒排序" then
        addSql = addSql & "   Order By SpacePurview Desc" 

    elseIf ascDescType = "登录时间正排序" then
        addSql = addSql & "   Order By OnLineTime Asc" 
    elseIf ascDescType = "登录时间倒排序" then
        addSql = addSql & "   Order By OnLineTime Desc" 

    elseIf ascDescType = "好友度正排序" then
        addSql = addSql & "   Order By CountNumb" 
    elseIf ascDescType = "好友度倒排序" then
        addSql = addSql & "   Order By CountNumb Desc" 

    elseIf ascDescType = "存在数正排序" then
        addSql = addSql & "   Order By QQNumb" 
    elseIf ascDescType = "存在数倒排序" then
        addSql = addSql & "   Order By QQNumb Desc" 

    elseIf ascDescType = "好友度＋存在数正排序" then
        addSql = addSql & "   Order By CountNumb+QQNumb" 
    elseIf ascDescType = "好友度＋存在数倒排序" then
        addSql = addSql & "   Order By CountNumb+QQNumb Desc" 

    elseIf ascDescType = "评论数正排序" then
        addSql = addSql & "   Order By ReplyNumb Asc" 
    elseIf ascDescType = "评论数倒排序" then
        addSql = addSql & "   Order By ReplyNumb Desc" 

    elseIf ascDescType = "留言时间正排序" then
        addSql = addSql & "   Order By UpDateTime Asc" 
    elseIf ascDescType = "留言时间倒排序" then
        addSql = addSql & "   Order By UpDateTime Desc" 

    elseIf ascDescType = "总数正排序" then
        addSql = addSql & "   Order By CountNumb Asc" 
    elseIf ascDescType = "总数倒排序" then
        addSql = addSql & "   Order By CountNumb Desc" 

    elseIf ascDescType = "添加时间正排序" then
        addSql = addSql & "   Order By AddTime Asc" 
    elseIf ascDescType = "添加时间倒排序" then
        addSql = addSql & "   Order By AddTime Desc" 

    elseIf ascDescType = "价格正排序" then
        addSql = addSql & "   Order By Price Asc" 
    elseIf ascDescType = "价格倒排序" then
        addSql = addSql & "   Order By Price Desc" 

    elseIf ascDescType = "运费正排序" then
        addSql = addSql & "   Order By FarePrice Asc" 
    elseIf ascDescType = "运费倒排序" then
        addSql = addSql & "   Order By FarePrice Desc" 

    elseIf ascDescType = "运费正排序" then
        addSql = addSql & "   Order By FarePrice Asc" 
    elseIf ascDescType = "运费倒排序" then
        addSql = addSql & "   Order By FarePrice Desc" 

    elseIf ascDescType = "描述值正排序" then
        addSql = addSql & "   Order By Description Asc" 
    elseIf ascDescType = "描述值倒排序" then
        addSql = addSql & "   Order By Description Desc" 

    elseIf ascDescType = "评论正排序" then
        addSql = addSql & "   Order By ReplyNum Asc" 
    elseIf ascDescType = "评论倒排序" then
        addSql = addSql & "   Order By ReplyNum Desc" 

    elseIf ascDescType = "上传时间正排序" then
        addSql = addSql & "   Order By UpDateTime Asc" 
    elseIf ascDescType = "上传时间倒排序" then
        addSql = addSql & "   Order By UpDateTime Desc" 

    elseIf ascDescType = "文件大小正排序" then
        addSql = addSql & "   Order By ZipSize Asc" 
    elseIf ascDescType = "文件大小倒排序" then
        addSql = addSql & "   Order By ZipSize Desc" 

    elseIf ascDescType = "运行平台正排序" then
        addSql = addSql & "   Order By Windows Asc" 
    elseIf ascDescType = "运行平台倒排序" then
        addSql = addSql & "   Order By Windows Desc" 

    elseIf ascDescType = "排序正排序" then
        addSql = addSql & "   Order By Sort Asc" 
    elseIf ascDescType = "排序倒排序" then
        addSql = addSql & "   Order By Sort Desc" 

    elseIf ascDescType = "标题正排序" then
        addSql = addSql & "   Order By Title Asc" 
    elseIf ascDescType = "标题倒排序" then
        addSql = addSql & "   Order By Title Desc" 

    elseIf ascDescType = "人气正排序" then
        addSql = addSql & "   Order By RenQi Asc" 
    elseIf ascDescType = "人气倒排序" then
        addSql = addSql & "   Order By RenQi Desc" 

    elseIf ascDescType = "作者正排序" then
        addSql = addSql & "   Order By Author Asc" 
    elseIf ascDescType = "作者倒排序" then
        addSql = addSql & "   Order By Author Desc" 


    elseIf ascDescType = "单词正排序" then
        addSql = addSql & "   Order By Ci Asc" 
    elseIf ascDescType = "单词倒排序" then
        addSql = addSql & "   Order By Ci Desc" 

    elseIf ascDescType = "破解值正排序" then
        addSql = addSql & "   Order By PoJie Asc" 
    elseIf ascDescType = "破解值倒排序" then
        addSql = addSql & "   Order By PoJie Desc" 

    elseIf ascDescType = "QQ性别正排序" then
        addSql = addSql & "   Order By QQSex Asc" 
    elseIf ascDescType = "QQ性别倒排序" then
        addSql = addSql & "   Order By QQSex Desc" 

    elseIf ascDescType = "QQ年龄正排序" then
        addSql = addSql & "   Order By QQAge Asc" 
    elseIf ascDescType = "QQ年龄倒排序" then
        addSql = addSql & "   Order By QQAge Desc" 

    elseIf ascDescType = "年龄正排序" then
        addSql = addSql & "   Order By Age Asc" 
    elseIf ascDescType = "年龄倒排序" then
        addSql = addSql & "   Order By Age Desc" 


    elseIf ascDescType = "下载数正排序" then
        addSql = addSql & "   Order By DownNumb Asc" 
    elseIf ascDescType = "下载数倒排序" then
        addSql = addSql & "   Order By DownNumb Desc" 

    elseIf ascDescType = "文件正排序" then
        addSql = addSql & "   Order By FileName Asc" 
    elseIf ascDescType = "文件倒排序" then
        addSql = addSql & "   Order By FileName Desc" 

    elseIf ascDescType = "文件夹正排序" then
        addSql = addSql & "   Order By FolderName Asc" 
    elseIf ascDescType = "文件夹倒排序" then
        addSql = addSql & "   Order By FolderName Desc" 

    elseIf ascDescType = "目录路径正排序" then
        addSql = addSql & "   Order By DirPath Asc" 
    elseIf ascDescType = "目录路径倒排序" then
        addSql = addSql & "   Order By DirPath Desc" 

    elseIf ascDescType = "播放时间正排序" then
        addSql = addSql & "   Order By PlayTime Asc" 
    elseIf ascDescType = "播放时间倒排序" then
        addSql = addSql & "   Order By PlayTime Desc" 

    elseIf ascDescType = "播放次数正排序" then
        addSql = addSql & "   Order By PlayNumb Asc" 
    elseIf ascDescType = "播放次数倒排序" then
        addSql = addSql & "   Order By PlayNumb Desc" 

    elseIf ascDescType = "发布时间正排序" then
        addSql = addSql & "   Order By PostingsTime Asc" 
    elseIf ascDescType = "发布时间倒排序" then
        addSql = addSql & "   Order By PostingsTime Desc" 

    elseIf ascDescType = "点击数正排序" then
        addSql = addSql & "   Order By ClickNumb Asc" 
    elseIf ascDescType = "点击数倒排序" then
        addSql = addSql & "   Order By ClickNumb Desc" 

    elseIf ascDescType = "访客数量正排序" then
        addSql = addSql & "   Order By VisitorsNumb Asc" 
    elseIf ascDescType = "访客数量倒排序" then
        addSql = addSql & "   Order By VisitorsNumb Desc" 

    elseIf ascDescType = "提交日期正排序" then
        addSql = addSql & "   Order By CreateDate Asc" 
    elseIf ascDescType = "提交日期倒排序" then
        addSql = addSql & "   Order By CreateDate Desc" 

    elseIf ascDescType = "浏览次数正排序" then
        addSql = addSql & "   Order By BrowseNumb Asc" 
    elseIf ascDescType = "浏览次数倒排序" then
        addSql = addSql & "   Order By BrowseNumb Desc" 

    elseIf ascDescType = "是否通过正排序" then
        addSql = addSql & "   Order By Through Asc" 
    elseIf ascDescType = "是否通过倒排序" then
        addSql = addSql & "   Order By Through Desc" 

    elseIf ascDescType = "提问时间正排序" then
        addSql = addSql & "   Order By ReplyTime Asc" 
    elseIf ascDescType = "提问时间倒排序" then
        addSql = addSql & "   Order By ReplyTime Desc" 






    elseIf ascDescType = "联合查询" then
        call LianHeSearch() 
    else
        call echo(ascDescType, "排序不存在") 
    end if 


    handleAscDesc = addSql 
end function 
'联合查询 【GuanXiQQ.Asp表】
function lianHeSearch()
    dim searchType, searchValue 
    searchType = request("SearchType") 
    searchValue = request("SearchValue") 
    '联合查询，某个QQ不存在，提示操作
    if searchType = "" and inStr(searchValue, "|") > 0 then
        CAll echo("", "联合") 
        dim splStr, QQIN, qQ, yesNo, c, n 
        yesNo = true 
        splStr = split(searchValue, "|") 
        for each qQ in splStr
            if qQ <> "" then
                if QQIN = "" then
                    QQIN = "'" & qQ & "'" 
                else
                    QQIN = QQIN & ",'" & qQ & "'" 
                end if 
                rs.open "Select BigClassName From [GuanXiQQ] Where BigClassName = '" & qQ & "'", conn, 1, 1 
                if rs.eof then
                    c = c & "<div align=center><iframe width='800' height=100 frameborder=0 scrolling=no src='Print.Asp?act=处理分析QQ&QQ=" & qQ & "'></iframe></div>" 
                    n = n + 1 
                    yesNo = false 
                end if : rs.close 
            end if 
        next 

        if yesNo = false then
            c = c & "<div align=center style='color:red;'>提示：有未处理的关系QQ【<b>" & n & "</b>】个</div>" 
            call rw(c) 
            response.end() 
        end if 
        conn.execute("Delete From [GuanXiQQ] Where BigClassName='联合查询'") 
        rs.open "Select * From [GuanXiQQ] Where BigClassName In(" & QQIN & ")", conn, 1, 1 
        while not rs.eof
            '合
            rsx.open "Select * From [GuanXiQQ] Where BigClassName='联合查询' And QQ='" & rs("QQ") & "'", conn, 1, 3 
            if rsx.eof then
                rsx.addNew 
                rsx("BigClassName") = "联合查询" 
                rsx("SearchList") = searchValue 
                rsx("SmallClassName") = rs("BigClassName") 
                rsx("QQ") = rs("QQ") 
                rsx("QQName") = rs("QQName") 
                rsx("CountNumb") = rs("CountNumb") + rs("QQNumb") 
            else
                rsx("CountNumb") = rsx("CountNumb") + rs("CountNumb") + rs("QQNumb") 
                rsx("QQNumb") = rsx("QQNumb") + 1 
                '累加
                if inStr("|" & rsx("SmallClassName") & "|", "|" & rs("QQ") & "|") = false then
                    rsx("SmallClassName") = rsx("SmallClassName") & "|" & rs("BigClassName") 
                end if 
            end if 
            rsx.update : rsx.close 
        rs.moveNext : wend : rs.close 
        '跳转到 联合查询
        CAll eerr("哈哈", "哈哈") 
        'Response.Redirect("GuanXiQQ.Asp?act=ShowGuanXiQQList&BigClassName=联合查询&AscDescType=两者相加倒排序&CountNumb=1")
        response.end() 
    end if 
'LianHeSearch = AddSql
end function 
'处理重复
function handleReset(addSql, AType, table)
    if AType = "" then
        handleReset = addSql 
        exit function 
    end if 

    if AType = "QQ号码重复" then
        addSql = "Select QQ From [" & table & "] Group By QQ Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "QQ名称重复" then
        addSql = "Select QQName  From [" & table & "] Group By QQName Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "空间权限重复" then
        addSql = "Select SpacePurview From [" & table & "] Group By SpacePurview Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "登录时间重复" then
        addSql = "Select OnLineTime  From [" & table & "] Group By OnLineTime Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "大类名称重复" then
        addSql = "Select BigClassName From [" & table & "] Group By BigClassName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "小类名称重复" then
        addSql = "Select SmallClassName  From [" & table & "] Group By SmallClassName Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "日志类型重复" then
        addSql = "Select QQcategory From [" & table & "] Group By QQcategory Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "标题重复" then
        addSql = "Select Title  From [" & table & "] Group By Title Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "评论数重复" then
        addSql = "Select ReplyNumb  From [" & table & "] Group By ReplyNumb Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "日志ID重复" then
        addSql = "Select Blogid  From [" & table & "] Group By Blogid Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "添加时间重复" then
        addSql = "Select AddTime  From [" & table & "] Group By AddTime Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "域名重复" then
        addSql = "Select WebSite  From [" & table & "] Group By WebSite Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "搜索值重复" then
        addSql = "Select SearchValue  From [" & table & "] Group By SearchValue Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "用户重复" then
        addSql = "Select UserName  From [" & table & "] Group By UserName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "价格重复" then
        addSql = "Select Price  From [" & table & "] Group By Price Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "运费重复" then
        addSql = "Select FarePrice  From [" & table & "] Group By FarePrice Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "省份重复" then
        addSql = "Select Province  From [" & table & "] Group By Province Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "城市重复" then
        addSql = "Select City  From [" & table & "] Group By City Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "描述值重复" then
        addSql = "Select Description  From [" & table & "] Group By Description Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "评论重复" then
        addSql = "Select ReplyNum  From [" & table & "] Group By ReplyNum Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "QQ备注名重复" then
        addSql = "Select QQBeiZhuName  From [" & table & "] Group By QQBeiZhuName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "公司名称重复" then
        addSql = "Select Company  From [" & table & "] Group By Company Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "招聘职位重复" then
        addSql = "Select WorkTitle  From [" & table & "] Group By WorkTitle Having(Count(*))>" & cint(rq("N")) 


    elseIf AType = "作者重复" then
        addSql = "Select Author  From [" & table & "] Group By Author Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "人气重复" then
        addSql = "Select RenQi  From [" & table & "] Group By RenQi Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "存在大小图重复" then
        addSql = "Select PhotoShowType  From [" & table & "] Group By PhotoShowType Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "网站标题重复" then
        addSql = "Select WebTitle  From [" & table & "] Group By WebTitle Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "网站地址重复" then
        addSql = "Select WebUrl  From [" & table & "] Group By WebUrl Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "星期数重复" then
        addSql = "Select DayNumb  From [" & table & "] Group By DayNumb Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "上传时间重复" then
        addSql = "Select UpDateTime  From [" & table & "] Group By UpDateTime Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "QQ年龄重复" then
        addSql = "Select QQAge  From [" & table & "] Group By QQAge Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "QQ性别重复" then
        addSql = "Select QQSex  From [" & table & "] Group By QQSex Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "年龄重复" then
        addSql = "Select Age  From [" & table & "] Group By Age Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "运行平台重复" then
        addSql = "Select Windows  From [" & table & "] Group By Windows Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "下载数重复" then
        addSql = "Select DownNumb  From [" & table & "] Group By DownNumb Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "文件大小重复" then
        addSql = "Select ZipSize  From [" & table & "] Group By ZipSize Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "文件重复" then
        addSql = "Select FileName  From [" & table & "] Group By FileName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "文件夹重复" then
        addSql = "Select FolderName  From [" & table & "] Group By FolderName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "目录路径重复" then
        addSql = "Select DirPath  From [" & table & "] Group By DirPath Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "网址重复" then
        addSql = "Select HttpUrl  From [" & table & "] Group By HttpUrl Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "播放时间重复" then
        addSql = "Select PlayTime  From [" & table & "] Group By PlayTime Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "播放次数重复" then
        addSql = "Select PlayNumb  From [" & table & "] Group By PlayNumb Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "发布时间重复" then
        addSql = "Select PostingsTime  From [" & table & "] Group By PostingsTime Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "子类名称重复" then
        addSql = "Select ThreeClassName  From [" & table & "] Group By ThreeClassName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "点击数重复" then
        addSql = "Select ClickNumb  From [" & table & "] Group By ClickNumb Having(Count(*))>" & cint(rq("N")) 

    elseIf AType = "搜索词重复" then
        addSql = "Select BigClassName  From [" & table & "] Group By BigClassName Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "分类重复" then
        addSql = "Select Cate  From [" & table & "] Group By Cate Having(Count(*))>" & cint(rq("N")) 
    elseIf AType = "日志发布时间重复" then
        addSql = "Select pubTime  From [" & table & "] Group By pubTime Having(Count(*))>" & cint(rq("N")) 



    '外部MDB
    elseIf AType = "单词重复" then
        addSql = "Select Ci  From [" & table & "] Group By Ci Having(Count(*))>" & cint(rq("N")) 
    '外部MDB
    elseIf AType = "英文解释重复" then
        addSql = "Select Shi  From [" & table & "] Group By Shi Having(Count(*))>" & cint(rq("N")) 
    '外部MDB
    elseIf AType = "英文解释重复" then
        addSql = "Select Shi  From [" & table & "] Group By Shi Having(Count(*))>" & cint(rq("N")) 



    else
        call echo(AType, "重复不存在") 
    end if 

    handleReset = addSql 
end function 

'自定义SQL
function selfSql(addSql, searchValue, table)
    dim sql 

    searchValue = regExp_Replace(searchValue, "SQL *=", "Sql ") 
    if checkStr(searchValue, "SQL ") = true then
        sql = trim(searchValue) 
        sql = trim(right(sql, len(sql) - 4)) 
        if checkStr(sql, "Delete ") = true then
            sql = sql & " &nbsp; <a href=""?act=ClearTable&MDBPath=" & request("MDBPath") & "&SQL=" & sql & """ target=""_blank"" class=""ARed"" onClick=""return Confirm();"">删除记录</a>" 
            CAll eerr("删除很危险，请慎重", sql) 
        elseIf checkStr(sql, "Update ") = true then
            CAll eerr("更新表操作", sql) 
        elseIf checkStr(sql, "Select ") = false then
            sql = "Select * From [" & table & "] " & sql 
        end if 
        addSql = sql 
    end if 
    selfSql = addSql 
end function 

'替换多个或单个搜索值
function replaceSearchValue(content, searchValue)
    dim splStr, s 
    if inStr(searchValue, " Or ") > 0 then
        splStr = split(searchValue, " Or ") 
        for each s in splStr
            if s <> "" then
                content = replace(content, s, "<font color=red><b>" & s & "</b></font>") 
            end if 
        next 
    elseIf inStr(searchValue, " And ") > 0 then
        splStr = split(searchValue, " And ") 
        for each s in splStr
            if s <> "" then
                content = replace(content, s, "<font color=red><b>" & s & "</b></font>") 
            end if 
        next 
    elseIf content <> "" then
        content = replace(content, searchValue, "<font color=red><b>" & searchValue & "</b></font>") 
    end if 
    replaceSearchValue = content 
end function 

'显示复制内容列表
sub showCopyList(title, nOK, strList, startTime)
    if startTime <> "" then startTime = "用时：" & printTimeValue(dateDiff("s", startTime, now())) & "，&nbsp;" 
    call echo("【" & title & "】", "共有【" & nOK & "】条，&nbsp;" & startTime & "<a href=javascript:CopyTEXT('CopyContent');>复制</a>") 
    CAll rw("<div style='margin-left:220px;'><textarea name='CopyContent' id='CopyContent' style='width:800px;height:400px;'>" & strList & "</textarea></div>") 
end sub 
%> 


