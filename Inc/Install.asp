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
<!--#Include file = "Config.Asp"-->
<!DOCTYPE html> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /> 
<title>安装SQLServer/Access数据库,QQ:313801120:::</title>  
<style type="text/css"> 
a img{border:none}  
.imga{vertical-align:bottom;} 
.imgb{vertical-align:bottom; display:block;}/*解决图片下面有空隙的简单方法  以前是img{这样图片就会换行*/ 
body,div,p,img,dl,dt,dd,ul,ol,li,h1,h2,h3,h4,h5,h6,pre,form,fieldset,input,textarea,blockquote{padding:0px;margin:0px} 
li{list-style-type:none} 
a{font-size:12px;line-height:18px;color:#000000;text-decoration:none} 
a:hover{text-decoration:none;color:#000099} 
/*PHPCMS表单样式*/ 
.input-text, .measure-input, textarea, input.date, input.endDate, .input-focus { 
  border: 1px solid #A7A6AA; 
  height: 22px; 
  line-height:22px; 
  margin: 0 5px 0 0; 
  padding: 2px 0 2px 5px; 
  border: 1px solid #d0d0d0; 
  background: #FFF url(../images/input.png) repeat-x; 
  font-family: Verdana, Geneva, sans-serif,"宋体"; 
  font-size: 12px; 
} 
input.date, input.endDate { 
  background: #fff url(../images/input_date.png) no-repeat right 3px; 
  padding-right: 18px; 
  font-size: 12px; 
} 
select { 
  vertical-align: middle; 
  background: none repeat scroll 0 0 #F9F9F9; 
  border-color: #666666 #CCCCCC #CCCCCC #666666; 
  border-style: solid; 
  border-width: 1px; 
  color: #333; 
  padding: 2px; 
} 
/*自定义按钮 待改进*/ 
.btnclick1{ 
    color: #000; 
    font-size: 14px; 
    padding:0 20px; 
    background-color:#fff; 
    border:0px; 
    border:1px solid #666666; 
    text-decoration:none; 
    cursor:pointer; 
    line-height: 26px; 
    font-weight:bold; 
    border-radius:5px;-moz-border-radius:5px;-webkit-border-radius:5px;-o-border-radius:5px; 
} 
.btnclick1:hover{ 
    background-color:#E6E6E6; 
    background-position: 0px -50px;     
} 
.btnclick1:active{ 
    background-color:#fff; 
} 



.pright { 
width: 720px; 
margin:0 auto; 
} 

.pr-title { 
width: 720px; 
height: 22px; 
margin: 8px auto 0px;  
overflow: hidden; 
} 
.pr-title h3 { 
width: 158px; 
height: 22px; 
line-height: 22px; 
overflow: hidden; 
display: block; 
font-size: 12px; 
padding-top: 1px; 
text-indent: 10px;  
letter-spacing: 2px; 
color: #6D8A4F; 
font-weight: bold; 
} 



input{ 
vertical-align:middle; 
margin-right:3px; 
font-size:12px; 
} 
textarea{ 
vertical-align:top; 
font-size:12px; 
line-height:156%; 
border:1px solid #AAA; 
padding:3px; 
letter-spacing:1px; 
word-break:break-all; 
overflow-y:auto; 
} 
.input-txt{ 
padding:4px 8px 4px 6px; 
border:1px solid #AAA; 
font-size:12px; 
color:#000; 
width:200px; 
} 

.textipt_on{ 
border:1px solid #F90; 
} 


hr{ 
height:1px; 
line-height:1px; 
overflow:hidden; 
border-width:1px 0px 0px 0px ; 
border-top:1px solid #E6E6E6;/*兼容Opera*/ 
} 
hr:empty { 
margin:8px 0px 7px 0px !important; 
margin:0px; 
} 

small{ 
font-size:12px; 
} 

.moncolor td{ 
background:#FFC; 
} 

.twbox{ 
width:706px; 
border:1px solid #CFDCC9; 
font-size:12px; 
overflow:hidden; 
margin:8px auto; 

} 


.twbox thead tr td{  
height:31px; 
line-height:31px; 
text-indent:10px; 
} 

.twbox thead tr td strong{ 
letter-spacing:2px; 
margin-right:14px; 
color:#FFF; 
font-size:14px; 
} 


.twbox thead tr td span{ 
color:#CDA; 
} 


.twbox thead tr td  p{ 
height:31px; 
display:inline; 
float:right; 
margin:-31px 10px 0 0; 
overflow:hidden; 
} 
.twbox thead tr td  p *{ 
float:right; 
} 
.twbox thead tr td a.thlink{ 
color:#FFF; 
} 
.twbox thead tr td a.thlink:hover{ 
color:#FFFF00; 
text-decoration:none; 
} 


.twbox tbody { 
overflow:hidden; 
text-align:left; 
} 

.twbox tbody tr th{ 
background:#F9FCEF; 
color:#6D8A4F; 
line-height:21px; 
height:21px; 
text-indent:30px; 
font-weight:normal; 
border-bottom:1px solid #EFF7D8; 
letter-spacing:2px; 
} 

.twbox tbody tr td{ 
padding:7px; 
border-bottom:1px solid #F2F2F2; 
color:#333; 
vertical-align:top; 
} 

.twbox tbody tr td p{ 
line-height:21px; 
} 
.twbox tbody tr td p strong img{ 
vertical-align:middle; 
} 
.twbox tbody tr td img{ 
vertical-align:top; 
margin:0px 10px 5px 0px; 

} 
.twbox tbody tr td small{ 
color:#888; 
} 

.twbox tfoot tr td{ 
padding:10px; 
line-height:25px; 
text-align:center; 
} 
.twbox tfoot tr td p{ 
line-height:21px; 
margin-bottom:10px; 
} 

input.but{ 
height:26px; 
padding-left:6px; 
padding-right:6px; 
line-height:26px; 
font-weight:bold; 
letter-spacing:1px; 
color:#FFF; 
background-color:#FC3; 
} 

.onetd{ 
width:120px; 
text-align:right; 
line-height:25px; 
} 

.mytipwrap{ 
    line-height:30px; 
    color:#999999; 
} 
a.mytip { 
    line-height: 14px; 
    padding: 6px 20px; 
    border-style: solid; 
    border-width: 1px; 
    border-color: #EEE #CCC #CCC #EEE; 
    background: #FAFAFA; 
    color: #333;  
    margin-right:10px; 
    text-decoration: none; 
} 
</style> 
</head> 
<body> 
<% 


dim selectDatabase, dbhost, dbuser, dbpwd, dbname, accessDir, accessPath, configFilePath, content, startStr, endStr, findStr, replaceStr 
dim tempContent, msgStr, isYes, connStr 
dim showLayout				'显示布局

dim configTableList,tableListC
 
	configTableList="|webcolumn|articledetail|onepage|tablecomment|website|admin|searchstat|friendlink|job|guestbook|feedback|listmenu|lineqq|linevote|member|memberlog|payment|previeworder|productcomment|websitestat|systemlog|weblayout|webmodule|caiweb|caiconfig|caidata|bidding|weburlscan|webdomain|CaiColumn|CaiDetail|webStyle|"
  
  configTableList=lcase(configTableList)
  
  
                  dim webdataDir
				  webdataDir=request("webdataDir")
				  if request("webdataDir")="" then
                  	webdataDir="/Data/WebData"
				  end if

call loadRun()		'【@是.netc屏蔽@】'【@是jsp屏蔽@】
sub loadRun()
	if EDITORTYPE="asp" then
    	configFilePath = "/inc/Config.Asp" 
	elseif EDITORTYPE="aspx" then
    	configFilePath = "/inc_csharp/Config.Aspx" 
	elseif EDITORTYPE="jsp" then
    	configFilePath = "/inc_jsp/config.jsp" 
	end if
	if request("act") = "createAccess" then
		call createAccess()                                       '创建数据库
	elseif request("act") = "displayImportTableLayout" then
		call displayImportTableLayout()               '显示导入数据页面
	elseif request("act") = "displayImportData" then
		call displayImportData()                             '显示导入数据
	else
		call displayDefault()
	end if
	'call echo("tableListC",tableListC)
end sub


'判断创建表是否可以生成
function checkCreateTable(byval tableName)
	dim sTableName
	sTableName=lcase(tableName)
	if db_PREFIX<>"" then
		sTableName=mid(sTableName, len(db_PREFIX)+1 )
	end if
	'call echo("db_PREFIX",db_PREFIX)
	'call echo("sTableName",sTableName)
	'call eerr("tableName",tableName)
	tableListC = tableListC & sTableName & "|"
 	'call echo(configTableList,sTableName)
	if instr("|"& configTableList &"|","|"& sTableName &"|")=false and instr("|"& configTableList &"|","|*|")=false then
		checkCreateTable=true
	else
		checkCreateTable=checkTable_show(tableName)
	end if
end function


'注意，创建数据时 Default """"  默认值用这个   不要用这种Default ''    20150506
'网站数据库
sub webData(db_PREFIX, loginname, loginpwd)
    dim tableName, sql, isAddData, splStr, splxx, i, s, splTrue, bigClassName 
    dim splUrl, splNavName, sNavClassName, splImg, templateIndex, templateHome, templateMain 
    'call echo("db_PREFIX",db_PREFIX)

    '采集栏目表-------------------------------
    tableName = db_PREFIX & "CaiColumn"  
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 

        sql = sql & "columnName VarChar Default """","                                  '导航名称 
        sql = sql & "columnType VarChar Default """","                                  '导航类型
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "httpurl VarChar Default """","                                     '网址
        sql = sql & "morePageUrl VarChar Default """","                                 '更多页配置  {*}        为i页
        sql = sql & "charset VarChar Default """","                                     '字符编码
        sql = sql & "startPage Int Default 0,"                                           '当前页数
        sql = sql & "endPage Int Default 0,"                                          '总页数 
		
        sql = sql & "sType VarChar Default """","                                       '类型 
        sql = sql & "sAction VarChar Default """","                                     '动作 
        sql = sql & "labelNameList VarChar Default """","                               '标签名称列表 
        sql = sql & "htmlValue VarChar Default """","                                   'html内容值 
		
        sql = sql & "fieldName VarChar Default """","                                   '字段名称
        sql = sql & "fieldCheck Int Default 0,"                                         '字段检测
		
        sql = sql & "isCaiPage YesNo Default No,"                                       '是否为采集页 

        sql = sql & "isThrough YesNo Default Yes,"                                       '通过审核 
        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "bigImage VarChar Default """","                                    '大图片地址
        sql = sql & "bannerImage VarChar Default """","                                 '当前Banner图片 
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "fontB YesNo Default No,"                                           '字粗
        sql = sql & "sortRank Int Default 0,"                                           '排序编号 
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间 
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 

    '采集信息表-------------------------------
    tableName = db_PREFIX & "CaiDetail" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "sortRank Int Default 0,"                                           '排序编号 

        sql = sql & "value1 Text Default """","                                         '内容1
        sql = sql & "value2 Text Default """","                                         '内容2
        sql = sql & "value3 Text Default """","                                         '内容3
        sql = sql & "value4 Text Default """","                                         '内容4
        sql = sql & "value5 Text Default """","                                         '内容5
        sql = sql & "value6 Text Default """","                                         '内容6
        sql = sql & "value7 Text Default """","                                         '内容7
        sql = sql & "value8 Text Default """","                                         '内容8
        sql = sql & "value9 Text Default """","                                         '内容9

        sql = sql & "fieldName1 VarChar Default """","                                  '字段名称1
        sql = sql & "fieldName2 VarChar Default """","                                  '字段名称2
        sql = sql & "fieldName3 VarChar Default """","                                  '字段名称3
        sql = sql & "fieldName4 VarChar Default """","                                  '字段名称4
        sql = sql & "fieldName5 VarChar Default """","                                  '字段名称5
        sql = sql & "fieldName6 VarChar Default """","                                  '字段名称6
        sql = sql & "fieldName7 VarChar Default """","                                  '字段名称7
        sql = sql & "fieldName8 VarChar Default """","                                  '字段名称8
        sql = sql & "fieldName9 VarChar Default """","                                  '字段名称9

        sql = sql & "fieldCheck1 VarChar Default """","                                 '字段检测1
        sql = sql & "fieldCheck2 VarChar Default """","                                 '字段检测2
        sql = sql & "fieldCheck3 VarChar Default """","                                 '字段检测3
        sql = sql & "fieldCheck4 VarChar Default """","                                 '字段检测4
        sql = sql & "fieldCheck5 VarChar Default """","                                 '字段检测5
        sql = sql & "fieldCheck6 VarChar Default """","                                 '字段检测6
        sql = sql & "fieldCheck7 VarChar Default """","                                 '字段检测7
        sql = sql & "fieldCheck8 VarChar Default """","                                 '字段检测8
        sql = sql & "fieldCheck9 VarChar Default """","                                 '字段检测9

        sql = sql & "title VarChar Default """","                                       '标题 
        sql = sql & "isThrough YesNo Default No,"                                       '通过审核 
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间 
        sql = sql & "flags VarChar Default """","                                       '旗 
        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍 
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
	
		
    '网站栏目表-------------------------------
    tableName = db_PREFIX & "WebColumn"  
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 

        sql = sql & "columnName VarChar Default """","                                  '导航名称
        sql = sql & "columnEnName VarChar Default """","                                '导航名称(英文)
        sql = sql & "columnType VarChar Default """","                                  '导航类型
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "adminId Int Default 0,"                                            '管理员id

        sql = sql & "isDisplay YesNo Default No,"                                                  '是否显示
        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "bigImage VarChar Default """","                                    '大图片地址
        sql = sql & "bannerImage VarChar Default """","                                 '当前Banner图片

        sql = sql & "flags VarChar Default """","                                       '旗
        sql = sql & "displayTitle VarChar Default """","                                '显示标题
        sql = sql & "lableTitle VarChar Default """","                                  '标签描述


        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "fontB YesNo Default No,"                                           '字粗
        sql = sql & "isMakeHtml YesNo Default Yes,"                                     '开启生成Html




        sql = sql & "nPageSize Int Default 10,"                                         '每页显示条数
        sql = sql & "sortSql VarChar Default """","                                     '排序SQL
        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '是否开启生成Html

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间


        sql = sql & "noteBody Text Default """","                                       '笔记信息

        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "tempTXT1 Text Default """","                                       '暂存文本1
        sql = sql & "tempTXT2 Text Default """","                                       '暂存文件2
		
		
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 

    '文章信息表-------------------------------
    tableName = db_PREFIX & "ArticleDetail" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "adminId Int Default 0,"                                            '管理员id

        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "bigImage VarChar Default """","                                    '大图片地址
        sql = sql & "bannerImage VarChar Default """","                                 '当前Banner图片
        sql = sql & "downloadFile VarChar Default """","                                '下载文件
        sql = sql & "smallimageAlt VarChar Default """","                               '小图alt
        sql = sql & "bigimageAlt VarChar Default """","                                 '大图alt
        sql = sql & "bannerimageAlt VarChar Default """","                              'Banneralt
		
		
        sql = sql & "image960px VarChar Default """","                                  '960像素图片
        sql = sql & "image1280px VarChar Default """","                                  '1280像素图片


        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "titleColor VarChar Default """","                                  '标题颜色
        sql = sql & "titleAlt VarChar Default """","                                    '标题alt
        sql = sql & "lableTitle VarChar Default """","                                  '标签描述

        sql = sql & "isThrough YesNo Default No,"                                       '通过审核

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "occasions VarChar Default """","                                   '场合
        sql = sql & "hotline VarChar Default """","                                     '热线
        sql = sql & "model VarChar Default """","                                       '型号
        sql = sql & "author VarChar Default """","                                      '作者
        sql = sql & "articleSource VarChar Default """","                               '文章来源

        sql = sql & "price Float Default 0,"                                            '产品价格
        sql = sql & "newPrice Float Default 0,"                                         '产品新价格20150729
        sql = sql & "memberPrice Float Default 0,"                                      '会员价格
        sql = sql & "sold Int Default 0,"                                               '已售出

        sql = sql & "memberType VarChar Default """","                                  '会员类型
        sql = sql & "memberUser VarChar Default """","                                  '会员名
        sql = sql & "hits Int Default 0,"                                               '点击数



        sql = sql & "productAbout Text Default """","                                   '产品介绍
        sql = sql & "articleDescription Text Default """","                             '摘要(文章简要描述)20141225

        sql = sql & "httpUrl VarChar Default """","                                     '网址
        sql = sql & "recordUrl VarChar Default """","                                   '远程网址
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接

        sql = sql & "flags VarChar Default """","                                       '旗

        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '开启生成Html



        sql = sql & "articleInfoStyle VarChar Default """","                            '列表内容样式
        sql = sql & "articleInfoPhotoWidth VarChar Default """","                       '内容页图片宽
        sql = sql & "articleInfoPhotoHeight VarChar Default """","                      '内容页图片高 0为引用上面的 -1为空


        sql = sql & "relatedTags VarChar Default """","                                 '相关标签20141219
        sql = sql & "weight Int Default 0,"                                             '权重：     (越小越靠前)



        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍

        sql = sql & "tempTXT1 Text Default """","                                       '暂存文本1
        sql = sql & "tempTXT2 Text Default """","                                       '暂存文件2
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 

    '单页表-------------------------------
    tableName = db_PREFIX & "OnePage" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "displayTitle VarChar Default """","                                '显示标题
        sql = sql & "adminId Int Default 0,"                                            '管理员id

        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "fontB YesNo Default No,"                                           '字粗
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接
        sql = sql & "sortRank Int Default 0,"                                           '顺序
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "isRecommend YesNo Default Yes,"                                    '推荐
        sql = sql & "lableTitle VarChar Default """","                                  '标签描述
        sql = sql & "banner VarChar Default """","                                      '当前Banner图片
		
		

        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "bigImage VarChar Default """","                                    '大图片地址
        sql = sql & "bannerImage VarChar Default """","                                 '当前Banner图片
        sql = sql & "downloadFile VarChar Default """","                                '下载文件
        sql = sql & "smallimageAlt VarChar Default """","                               '小图alt
        sql = sql & "bigimageAlt VarChar Default """","                                 '大图alt
        sql = sql & "bannerimageAlt VarChar Default """","                              'Banneralt

        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '开启生成Html

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 

    '文章评论表 20160129-------------------------------
    tableName = db_PREFIX & "TableComment" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "UserID Int Default 0,"                                             '用户ID
        sql = sql & "itemID Int Default 0,"                                             '项目ID 产品ID
        sql = sql & "tableName VarChar Default """","                                   '表名称
        sql = sql & "userName VarChar Default """","                                    '用户名称
        sql = sql & "title VarChar Default """","                                       '评论标题
        sql = sql & "EMail VarChar Default """","                                       '邮箱
        sql = sql & "tel VarChar Default """","                                         '电话ID
        sql = sql & "ip VarChar Default """","                                          '评论者IP

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "reply Text Default """","                                          '回复内容
        sql = sql & "noteBody Text Default """","                                       '笔记信息

        sql = sql & "isThrough Int Default 0,"                                          '是否通过
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 

    '网站信息表-------------------------------
    tableName = db_PREFIX & "WebSite" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "WebSiteUrl VarChar Default """","                                  '网站地址
        sql = sql & "webSiteBottom Text Default """","                                  '底部信息 WebSiteBottom
        sql = sql & "WebSiteFlow Int Default 999,"                                      '网站统计数
        sql = sql & "WebSiteFlowStyle Int Default 1,"                                   '网站统计样式
        sql = sql & "WebSiteFlowMedian Int Default 6,"                                  '网站统计几位数Median中位数
        sql = sql & "ProductList VarChar Default """","                                 '产品列表
        sql = sql & "NewsList VarChar Default """","                                    '新闻列表
        sql = sql & "NewsDid VarChar Default """","                                     '新闻大类
        '201109追加
        sql = sql & "TZ51La VarChar Default """","                                      '统计网站
        sql = sql & "UserEmail VarChar Default """","                                   '用户Email
        sql = sql & "ProductDid VarChar Default """","                                  '产品大类
        sql = sql & "TemplateIndex Text Default """","                                '模板首页
        sql = sql & "TemplateHome Text Default """","                                 '模板内页
        sql = sql & "TemplateMain Text Default """","                               '模板详细页
        sql = sql & "TemplateMain2 Text Default """","                              '模板详细页2
        sql = sql & "TemplateMain3 Text Default """","                              '模板详细页3
        sql = sql & "UseNumb Text Default """","                                        '网站使用次数
        '20111118追加
        sql = sql & "WebRecord VarChar Default """","                                   '网站记录
        sql = sql & "ContentWebRecord Text Default """","                               '全部网站记录
        '20111128追加
        sql = sql & "UseHttpUrl VarChar Default """","                                  '当前使用网址
        sql = sql & "TempUseHttpUrl Text Default """","                                 '存储使用过的网址
        sql = sql & "WebDate DateTime Default Date(),"                                  '网站时间
        '2013,12,18
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述

        sql = sql & "WebTemplate VarChar Default """","                         '网站模板路径
        sql = sql & "WebSkins VarChar Default """","                       '网站样式路径'暂时不要用

        sql = sql & "WebFolderName VarChar Default """","                            '网站文件夹名称 new20150721


        sql = sql & "WebImages VarChar Default """","               '网站Images路径
        sql = sql & "WebCss VarChar Default """","                     '网站Css路径
        sql = sql & "WebJs VarChar Default """","                       '网站Js路径

        sql = sql & "AddWebSite YesNo Default No,"                                      'URL前添加域名
        sql = sql & "UpdateHtml YesNo Default No,"                                      '修改后更新HTML
        sql = sql & "isHtmlFormatting YesNo Default No,"                                '是否为HTML格式化
        sql = sql & "isWebLabelClose YesNo Default No,"                                 '是否为闭合网页标签
        sql = sql & "isCnToEn YesNo Default No,"                                        '是否为文本名汉字转拼音
        sql = sql & "flags VarChar Default """","                                       '旗


        '20150114
        sql = sql & "ModuleSkins VarChar Default """","                                 '模块皮肤
        sql = sql & "FindTpl VarChar Default """","                                     '查找要替换模板内容
        sql = sql & "ReplaceTpl VarChar Default """","                                  '替换模板指定内容，批量

        sql = sql & "WebCodeFindTpl VarChar Default """","                              '查找要替换网页内容
        sql = sql & "WebCodeReplaceTpl VarChar Default """","                           '替换网页指定内容，批量

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
		
        sql = sql & "templateSetCode VarChar Default """","                             '读取模板时默认编码
		
        sql = sql & "isMemberVerification VarChar Default 0,"                           '开启会员验证


        '最后垫底
        sql = sql & "WebHtml VarChar Default 0)"                                        '网站运行状态，如ASP，HTML
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
		
        conn.execute(sql) 
    end if 

    '管理员表-------------------------------
    tableName = db_PREFIX & "Admin" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "userName VarChar Default """","                                    '账号
        sql = sql & "pwd VarChar Default """","                                         '密码
        sql = sql & "pseudonym VarChar Default """","                                   '笔名


        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "regIP VarChar Default """","                                       '注册IP
        sql = sql & "upIP VarChar Default """","                                        '修改IP


        sql = sql & "quanXian Text Default """","                                       '管理权限
        sql = sql & "verificationMode Int Default 1,"                                   '验证方式，1为Session
        sql = sql & "adminLevel VarChar Default """","                                  '级别
        sql = sql & "channel VarChar Default """","                                     '频道
        sql = sql & "mtest YesNo Default No,"                                           '修改后更新HTML
        sql = sql & "flags Text Default """","                                          '旗        只有这个后台做了处理 因为权限特别多


        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '搜索统计表-------------------------------
    tableName = db_PREFIX & "SearchStat" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "userid Int Default -1,"                                          	'会员id		
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "title VarChar Default """","                                       '搜索关键词
        sql = sql & "isThrough YesNo Default No,"                                       '通过
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '是否开启生成Html
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "author VarChar Default """","                                      '作者

        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接
        sql = sql & "flags VarChar Default """","                                       '旗

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '友情链接表-------------------------------
    tableName = db_PREFIX & "FriendLink" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 

        sql = sql & "adminId Int Default 0,"                                            '管理员id
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "titleColor VarChar Default """","                                  '标题颜色
        sql = sql & "lableTitle VarChar Default """","                                  '标签描述
        sql = sql & "httpUrl VarChar Default """","                                     '网址
        sql = sql & "sortRank Int Default 0,"                                           '排序编号

        sql = sql & "titleAlt VarChar Default """","                                    '标题alt
        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "smallimageAlt VarChar Default """","                               '小图alt


        sql = sql & "isThrough YesNo Default No,"                                       '通过审核
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间


        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接
        sql = sql & "flags VarChar Default """","                                       '旗

        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '开启生成Html

        sql = sql & "weight Int Default 0,"                                             '权重：     (越小越靠前)

        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '招聘表-------------------------------
    tableName = db_PREFIX & "Job" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "title VarChar Default """","                                       '职位名称
        sql = sql & "sex VarChar Default """","                                         '性别
        sql = sql & "age VarChar Default """","                                         '年龄
        sql = sql & "education VarChar Default """","                                   '教育
        sql = sql & "workArea VarChar Default """","                                    '工作区
        sql = sql & "monthly VarChar Default """","                                     '月薪待遇
        sql = sql & "startDateTime VarChar Default """","                               '发布日期
        sql = sql & "endDateTime VarChar Default """","                                 '截止日期
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式

        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "noFollow Int Default 0,"                                           '不要追踪此网页上的链接
        sql = sql & "flags VarChar Default """","                                       '旗
        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "titleAlt VarChar Default """","                                    '标题alt
        sql = sql & "smallImage VarChar Default """","                                  '小图片地址
        sql = sql & "smallimageAlt VarChar Default """","                               '小图alt
        sql = sql & "isThrough Int Default 0,"                                          '是否通过

        sql = sql & "isHtml YesNo Default No,"                                          '是否为html
        sql = sql & "isOnHtml YesNo Default Yes,"                                       '开启生成Html
		
		
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '留言表-------------------------------
    tableName = db_PREFIX & "GuestBook" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "columnid VarChar Default """","                                    '栏目ID
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "guestName VarChar Default """","                                   '姓名
        sql = sql & "tel VarChar Default """","                                         '电话
        sql = sql & "fax VarChar Default """","                                         '传真
        sql = sql & "EMail VarChar Default """","                                       '邮箱
        sql = sql & "mobile VarChar Default """","                                      '手机
        sql = sql & "QQ VarChar Default """","                                          'QQ号码
        sql = sql & "MSN VarChar Default """","                                         'MSN号码
        sql = sql & "company VarChar Default """","                                     '公司
        sql = sql & "address VarChar Default """","                                     '地址
        sql = sql & "postcode VarChar Default """","                                    '邮编
        sql = sql & "ip VarChar Default """","                                          'IP地址
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "fontColor VarChar Default """","                                   '字颜色

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "isThrough Int Default 0,"                                          '是否通过


        sql = sql & "reply Text Default """","                                          '回复内容
        sql = sql & "replyip Text Default """","                                        '回复IP
        sql = sql & "replydatetime Text Default """","                                  '回复时间

        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '反馈表-------------------------------
    tableName = db_PREFIX & "Feedback" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "columnid VarChar Default """","                                    '栏目ID
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "feedbacktype VarChar Default """","                                '反馈类型
        sql = sql & "guestName VarChar Default """","                                   '姓名
        sql = sql & "tel VarChar Default """","                                         '电话
        sql = sql & "fax VarChar Default """","                                         '传真
        sql = sql & "EMail VarChar Default """","                                       '邮箱
        sql = sql & "mobile VarChar Default """","                                      '手机
        sql = sql & "QQ VarChar Default """","                                          'QQ号码
        sql = sql & "MSN VarChar Default """","                                         'MSN号码
        sql = sql & "company VarChar Default """","                                     '公司
        sql = sql & "address VarChar Default """","                                     '地址
        sql = sql & "postcode VarChar Default """","                                    '邮编
        sql = sql & "ip VarChar Default """","                                          'IP地址
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "fontColor VarChar Default """","                                   '字颜色

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "isThrough Int Default 0,"                                          '是否通过


        sql = sql & "reply Text Default """","                                          '回复内容
        sql = sql & "replyip Text Default """","                                        '回复IP
        sql = sql & "replydatetime Text Default """","                                  '回复时间

        sql = sql & "noteBody Text Default """","                                       '笔记信息
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '后台菜单表-------------------------------
    tableName = db_PREFIX & "ListMenu" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "lablename VarChar Default """","                                   '标签名称
        sql = sql & "customAUrl VarChar Default """","                                  '自定义链接网址

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "isDisplay YesNo Default No,"                                                  '显示与隐藏
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '在线QQ表------------------------------
    tableName = db_PREFIX & "LineQQ" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "QQ VarChar Default """","                                          'QQ号码
        sql = sql & "IsOnlineChat YesNo Default No,"                                    '是否为在线聊天
        sql = sql & "IsAddFriend YesNo Default Yes,"                                    '是否为加为好友

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "bodycontent Text)"                                                 '更多详细内容
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '投票表-------------------------------
    tableName = db_PREFIX & "LineVote" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "title VarChar Default """","                                       '投票标题
        sql = sql & "Option1 VarChar Default """","                                     '投票选项一
        sql = sql & "Option2 VarChar Default """","                                     '投票选项二
        sql = sql & "Option3 VarChar Default """","                                     '投票选项三
        sql = sql & "Option4 VarChar Default """","                                     '投票选项四
        sql = sql & "Option5 VarChar Default """","                                     '投票选项五
        sql = sql & "Option6 VarChar Default """","                                     '投票选项六
        sql = sql & "Num1 Int,"                                                         '投票选项数一
        sql = sql & "Num2 Int,"                                                         '投票选项数二
        sql = sql & "Num3 Int,"                                                         '投票选项数三
        sql = sql & "Num4 Int,"                                                         '投票选项数四
        sql = sql & "Num5 Int,"                                                         '投票选项数五
        sql = sql & "Num6 Int,"                                                         '投票选项数六
        sql = sql & "isDisplay YesNo Default Yes,"                                      '投票显示关闭
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述
        sql = sql & "folderName VarChar Default """","                                  'HTMl文件夹名称
        sql = sql & "fileName VarChar Default """","                                    'HTML文件名称
        sql = sql & "banner VarChar Default """","                                      '当前Banner图片
        sql = sql & "templatePath VarChar Default """","                                '模板路径
        sql = sql & "target VarChar Default """","                                      '链接打开方式
        sql = sql & "fontColor VarChar Default """","                                   '字颜色
        sql = sql & "fontB YesNo Default No,"                                           '字粗
        sql = sql & "OnHtml YesNo Default Yes,"                                         '开启生成Html

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "VoteType YesNo Default Yes)"                                       '投票类型
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '会员表-------------------------------
    tableName = db_PREFIX & "Member" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "userid Int Default -1,"                                          	'会员id	
        sql = sql & "parentId Int Default -1,"                                          '父栏目ID
        sql = sql & "userType VarChar Default """","                                    '会员类型
        sql = sql & "userName VarChar Default """","                                    '会员账号
        sql = sql & "pwd VarChar Default """","                                         '会员密码
        sql = sql & "yunPwd VarChar Default """","                                      '原密码
        sql = sql & "sex VarChar Default """","                                         '性别  男
        sql = sql & "age Int Default 18,"                                               '年龄
        sql = sql & "tel VarChar Default """","                                         '电话
        sql = sql & "phone VarChar Default """","                                       '手机
        sql = sql & "fax VarChar Default """","                                         '传真
        sql = sql & "email VarChar Default """","                                       '邮箱
        sql = sql & "postcode VarChar Default """","                                    '邮编
        sql = sql & "address VarChar Default """","                                     '地址
        sql = sql & "company VarChar Default """","                                     '公司

        sql = sql & "regIP VarChar Default """","                                       '注册IP
        sql = sql & "regTime VarChar Default """","                                     '注册时间
		
        sql = sql & "lastLoginIP VarChar Default """","                            	'上次登录IP
        sql = sql & "lastLoginTime DateTime Default Now(),"                            '上次登录时间
        sql = sql & "loginIp VarChar Default """","                                     '登录IP
        sql = sql & "loginTime DateTime Default Now(),"                                '登录时间
        sql = sql & "loginCount Int Default 0,"                                         '登录次数


        sql = sql & "actionList Text Default """","                                      '动作列表
        sql = sql & "actionTime DateTime Default Now(),"                            '动作执行时间


        'QQ信息
        sql = sql & "openId VarChar Default """","                                      'openId
        sql = sql & "accessToken VarChar Default """","                                 'accessToken
        sql = sql & "nickname VarChar Default """","                                    'nickname
        sql = sql & "qqphoto VarChar Default """","                                     'QQ头像
        sql = sql & "useryear Int Default 18,"                                          '那一年出生


        sql = sql & "Province VarChar Default """","                                    '省份
        sql = sql & "City VarChar Default """","                                        '城市
        sql = sql & "Area VarChar Default """","                                        '区域20150305



		'特殊
        sql = sql & "machineName VarChar Default """","                                       '机器名
        sql = sql & "OSVersion VarChar Default """","                                       '操作系统
        sql = sql & "softPath VarChar Default """","                                       '软件路径
        sql = sql & "srceenWidth VarChar Default """","                                       '屏幕宽
        sql = sql & "srceenHeight VarChar Default """","                                       '屏幕高

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
		
        sql = sql & "expireDateTime DateTime Default Date(),"                               '到期时间
        sql = sql & "followIdList VarChar Default """","                                        '关注文章ID列表
		
        sql = sql & "isThrough Int Default 0,"                                          '是否通过
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '会员日志表-------------------------------
    tableName = db_PREFIX & "MemberLog" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "userid Int Default -1,"                                            '会员id
        sql = sql & "title VarChar Default """","                                       '标题
		'特殊
        sql = sql & "machineName VarChar Default """","                                       '机器名
        sql = sql & "OSVersion VarChar Default """","                                       '操作系统
        sql = sql & "softPath VarChar Default """","                                       '软件路径
        sql = sql & "srceenWidth VarChar Default """","                                       '屏幕宽
        sql = sql & "srceenHeight VarChar Default """","                                       '屏幕高
		
		
        sql = sql & "txt1 Text Default """","                                      '存储内容一
        sql = sql & "txt2 Text Default """","                                      '存储内容二

        sql = sql & "ip VarChar Default """","                                          'IP
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "isThrough Int Default 0,"                                          '是否通过
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '付款信息表-------------------------------
    tableName = db_PREFIX & "Payment" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "UserName VarChar Default """","                                    '会员账号
        sql = sql & "MemberId VarChar Default """","                                    '会员ID编号
        sql = sql & "Sex VarChar Default """","                                         '性别男
        sql = sql & "Age Int Default 18,"                                               '年龄

        sql = sql & "Tel VarChar Default """","                                         '电话
        sql = sql & "Phone VarChar Default """","                                       '手机
        sql = sql & "Fax VarChar Default """","                                         '传真
        sql = sql & "Email VarChar Default """","                                       '邮箱
        sql = sql & "postcode VarChar Default """","                                    '邮编
        sql = sql & "Address VarChar Default """","                                     '地址
        sql = sql & "Company VarChar Default """","                                     '公司
        sql = sql & "QQMSN VarChar Default """","                                       '公司
        sql = sql & "Ip VarChar Default """","                                          'IP

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "isThrough Int Default 0)"                                          '是否通过
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '购物信息表-------------------------------
    tableName = db_PREFIX & "Previeworder" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "MemberId VarChar Default """","                                    '会员编号账号
        sql = sql & "OrderId VarChar Default """","                                     '订单ID编号
        sql = sql & "ProductId VarChar Default """","                                   '产品ID编号
        sql = sql & "title VarChar Default """","                                       '产品名称
        sql = sql & "Total VarChar Default """","                                       '些产品总数
        sql = sql & "Price Int Default 18,"                                             '此产品单价
        sql = sql & "ProductSum VarChar Default """","                                  '此产品总

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "isThrough YesNo Default No)"                                       '是否发货
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '产品评论表-------------------------------
    tableName = db_PREFIX & "ProductComment" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "UserName VarChar Default """","                                    '评论者
        sql = sql & "title VarChar Default """","                                       '评论标题
        sql = sql & "Pid Int,"                                                          '信息ID
        sql = sql & "PTitle VarChar Default """","                                      '信息标题
        sql = sql & "bodycontent Text Default """","                                    '评论内容
        sql = sql & "Sort Int Default 1,"                                               '评论排序
        sql = sql & "Ip VarChar Default """","                                          '评论者IP

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "isThrough YesNo Default No)"                                       '是否显示评论
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '网站统计表 20160203-------------------------------
    tableName = db_PREFIX & "WebSiteStat" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "visiturl Text Default """","                                       '来访网址
        sql = sql & "viewurl Text Default """","                                        '浏览网址
        sql = sql & "browser VarChar Default """","                                     '浏览器版本
        sql = sql & "operatingSystem VarChar Default """","                             '操作系统
        sql = sql & "screenWH VarChar Default """","                                    '屏幕宽高
        sql = sql & "moreInfo Text Default """","                                       '更多信息
        sql = sql & "viewDateTime DateTime Default Now(),"                              '访问时间
        sql = sql & "ip VarChar Default """","                                          '访问IP

        sql = sql & "dateClass VarChar Default """","                                   '日期分类

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间



        sql = sql & "isThrough Int Default 0,"                                          '是否通过
        sql = sql & "noteInfo Text Default """","                                       '笔记信息
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '系统日志表-------------------------------
    tableName = db_PREFIX & "SystemLog" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "msgStr Text Default """","                                         '信息内容
        sql = sql & "tableName VarChar Default """","                                   '表名 也可以说module模块
        sql = sql & "url Text Default """","                                            '动作
        sql = sql & "adminId Int Default 0,"                                            '管理员id
        sql = sql & "adminName VarChar Default 0,"                                      '管理员名称
        sql = sql & "IP VarChar Default """","                                          'IP地址


        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if
	'模板样式表 20151228-------------------------------
    tableName = db_PREFIX & "webStyle" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "title VarChar Default """","                                  '标题
        sql = sql & "topStyle VarChar Default """","                               '顶部样式
        sql = sql & "defaultStyle VarChar Default """","                               '默认样式
        sql = sql & "leftStyle VarChar Default """","                                  '左边样式
        sql = sql & "rightStyle VarChar Default """","                                 '右边样式
        sql = sql & "navStyle VarChar Default """","                               'nav样式
        sql = sql & "footStyle VarChar Default """","                               'foot样式
        sql = sql & "mainStyle VarChar Default """","                                  'main页样式
        sql = sql & "productStyle VarChar Default """","                               'product页样式
        sql = sql & "newsStyle VarChar Default """","                               'news页样式
        sql = sql & "downStyle VarChar Default """","                               'down页样式
        sql = sql & "guestBookStyle VarChar Default """","                               'guestBook页样式
        sql = sql & "isThrough YesNo Default No,"                                       '通过审核
  
        sql = sql & "sortRank Int Default 0,"                                           '排序编号 
        sql = sql & "author VarChar Default """","                                      '作者
        sql = sql & "views Int Default 0,"                                              '点击次数  
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间 
        sql = sql & "aboutcontent Text Default """","                                   '简单介绍 
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if   
    '布局表 20151228-------------------------------
    tableName = db_PREFIX & "webLayout" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "layoutName VarChar Default """","                                  '布局名称
        sql = sql & "layoutList VarChar Default """","                                  '布局列表

        sql = sql & "sourceStr VarChar Default """","                                   '被替换内容
        sql = sql & "replaceStr VarChar Default """","                                  '统一替换内容

        sql = sql & "actionContent Text Default """","                                  '动作内容

        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "isDisplay YesNo Default No,"                                                  '是否显示
        sql = sql & "author VarChar Default """","                                      '作者
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "adminId Int Default 0,"                                            '管理员id


        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '模块表 20151228-------------------------------
    tableName = db_PREFIX & "webModule" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "moduleType VarChar Default """","                                  '模块类型
        sql = sql & "moduleName VarChar Default """","                                  '模块名称
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "isDisplay YesNo Default No,"                                                  '是否显示
        sql = sql & "author VarChar Default """","                                      '作者
        sql = sql & "views Int Default 0,"                                              '点击次数
        sql = sql & "adminId Int Default 0,"                                            '管理员id


        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间

        sql = sql & "aboutcontent Text Default """","                                   '简单介绍
        sql = sql & "memberUserCheck Int Default 0,"                                    '会员检测 1为需要登录，2不需要登录也显示
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '采集网站表-------------------------------
    tableName = db_PREFIX & "CaiWeb" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "columnname VarChar Default """","                                  '栏目名称
        sql = sql & "httpurl VarChar Default """","                                     '网址
        sql = sql & "morePageUrl VarChar Default """","                                 '更多页配置  {*}        为i页
        sql = sql & "charset VarChar Default """","                                     '字符编码
        sql = sql & "thisPage Int Default 0,"                                           '当前页数
        sql = sql & "countPage Int Default 0,"                                          '总页数
        sql = sql & "sType VarChar Default """","                                       '类型
        sql = sql & "sortRank Int Default 0,"                                           '排序编号

        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "IP VarChar Default """","                                          'IP地址
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '采集配置表-------------------------------
    tableName = db_PREFIX & "CaiConfig" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "sType VarChar Default """","                                       '类型
        sql = sql & "sAction VarChar Default """","                                     '动作
        sql = sql & "startStr VarChar Default """","                                    '开始截取字符
        sql = sql & "endStr VarChar Default """","                                      '结束截取字符
        sql = sql & "startAddStr VarChar Default """","                                 '开始截取字符前面追加字符
        sql = sql & "endAddStr VarChar Default """","                                   '结束截取字符后台追加字符
        sql = sql & "fieldName VarChar Default """","                                   '字段名称
        sql = sql & "fieldCheck Int Default 0,"                                         '字段检测

        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "IP VarChar Default """","                                          'IP地址

        sql = sql & "isThrough YesNo Default No,"                                       '通过审核
        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '采集数据表-------------------------------
    tableName = db_PREFIX & "CaiData" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "BigClassName VarChar Default """","                                '大类名称
        sql = sql & "columnname VarChar Default """","                                  '栏目名称

        sql = sql & "sType VarChar Default """","                                       '类型

        sql = sql & "value1 Text Default """","                                         '内容1
        sql = sql & "value2 Text Default """","                                         '内容2
        sql = sql & "value3 Text Default """","                                         '内容3
        sql = sql & "value4 Text Default """","                                         '内容4
        sql = sql & "value5 Text Default """","                                         '内容5
        sql = sql & "value6 Text Default """","                                         '内容6

        sql = sql & "fieldName1 VarChar Default """","                                  '字段名称1
        sql = sql & "fieldName2 VarChar Default """","                                  '字段名称2
        sql = sql & "fieldName3 VarChar Default """","                                  '字段名称3
        sql = sql & "fieldName4 VarChar Default """","                                  '字段名称4
        sql = sql & "fieldName5 VarChar Default """","                                  '字段名称5
        sql = sql & "fieldName6 VarChar Default """","                                  '字段名称6

        sql = sql & "fieldCheck1 VarChar Default """","                                 '字段检测1
        sql = sql & "fieldCheck2 VarChar Default """","                                 '字段检测2
        sql = sql & "fieldCheck3 VarChar Default """","                                 '字段检测3
        sql = sql & "fieldCheck4 VarChar Default """","                                 '字段检测4
        sql = sql & "fieldCheck5 VarChar Default """","                                 '字段检测5
        sql = sql & "fieldCheck6 VarChar Default """","                                 '字段检测6

        sql = sql & "sortRank Int Default 0,"                                           '排序编号
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "IP VarChar Default """","                                          'IP地址
        sql = sql & "isThrough YesNo Default No,"                                       '通过审核

        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '竞价表(百度竞价) 20151228-------------------------------
    tableName = db_PREFIX & "Bidding" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "searchWords VarChar Default """","                                 '搜索词
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "showReason VarChar Default """","                                  '展现理由
        sql = sql & "nComputerSearch Int Default 0,"                                    '电脑搜索量
        sql = sql & "nMoblieSearch Int Default 0,"                                      '移动搜索量
        sql = sql & "nCountSearch Int Default 0,"                                       '总体搜索量

        sql = sql & "nWordPrice Int Default 0,"                                         '新词左侧（上方）指导价
        sql = sql & "nDegree Int Default 0,"                                            '竞争激烈程度

        sql = sql & "sortRank Int Default 0,"                                           '顺序
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间


        sql = sql & "bodycontent Text Default """")"                                    '备注
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 
    '网站URL扫描表 20160428-------------------------------
    tableName = db_PREFIX & "WebUrlScan" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "bigClassName VarChar Default """","                                '大类名称
        sql = sql & "linkType VarChar Default """","                                    '链接类型
        sql = sql & "website VarChar Default """","                                     '域名
        sql = sql & "title VarChar Default """","                                       '标题
        sql = sql & "httpurl VarChar Default """","                                     '网址
        sql = sql & "totitle VarChar Default """","                                     '来访标题
        sql = sql & "tohttpurl VarChar Default """","                                   '来访网址
        sql = sql & "webState Int Default 0,"                                           '状态
        sql = sql & "openSpeed Int Default 0,"                                          '打开时间
        sql = sql & "charset VarChar Default """","                                     '字符编码
        sql = sql & "content_type VarChar Default """","                                '内容类型
        sql = sql & "link_count Int Default 0,"                                         '链接总数

        sql = sql & "webSize Int Default 0,"                                            '网页大小
        sql = sql & "sortRank Int Default 0,"                                           '顺序
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                               '修改时间
        sql = sql & "isThrough YesNo Default No)"                                       '通过审核
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if
	
    '网站域名表 20160511-------------------------------
    tableName = db_PREFIX & "WebDomain" 
    if checkCreateTable(tableName) = false then
        sql = "Create Table " & tableName & " (Id Int Identity(0,1) Primary Key," 
        sql = sql & "bigClassName VarChar Default """","                                '大类名称
        sql = sql & "website VarChar Default """","                                     '域名
        sql = sql & "webTitle VarChar Default """","                                    '网站自定标题
        sql = sql & "webKeywords VarChar Default """","                                 '网站关键词
        sql = sql & "webDescription Text Default """","                                 '网站描述

        sql = sql & "webState Int Default 0,"                                           '状态
        sql = sql & "openSpeed Int Default 0,"                                          '打开时间
        sql = sql & "charset VarChar Default """","                                     '字符编码
        sql = sql & "content_type VarChar Default """","                                '内容类型
        sql = sql & "server_name VarChar Default """","                                 '服务器名称


        sql = sql & "isAsp YesNo Default No,"                                           '可运行Asp程序
        sql = sql & "isAspx YesNo Default No,"                                          '可运行Aspx程序
        sql = sql & "isPhp YesNo Default No,"                                           '可运行Php程序
        sql = sql & "isJsp YesNo Default No,"                                           '可运行Jsp程序
        sql = sql & "isHtm YesNo Default No,"                                           '可运行Htm程序
        sql = sql & "isHtml YesNo Default No,"                                          '可运行Html程序
        sql = sql & "nlinks Int Default 0,"                                             '外链数
        sql = sql & "links Text Default """","                                          '外链列表
		
        sql = sql & "nmails Int Default 0,"                                             '邮箱数
        sql = sql & "mails Text Default """","                               			'邮箱列表
 
        sql = sql & "homepagelist VarChar Default """","                                '首页列表

        sql = sql & "flags VarChar Default """","                                       '旗
        sql = sql & "webSize Int Default 0,"                                            '网页大小
        sql = sql & "sortRank Int Default 0,"                                           '顺序
        sql = sql & "addDateTime DateTime Default Now(),"                               '创建时间
        sql = sql & "upDateTime DateTime Default Now(),"                                '修改时间

        sql = sql & "isDomain YesNo Default No,"                                        '是否为域名
        sql = sql & "isThrough YesNo Default No)"                                       '通过审核
        if MDBPath = "" then sql = handleSqlServer(sql)                                 '把Access数据库类型转成SqlServer数据库类型
        conn.execute(sql) 
    end if 


    isAddData = true 
    if isAddData = true then
        'Add Admin
        conn.execute("insert into " & db_PREFIX & "admin (username,pwd,flags) values('" & loginname & "','" & myMD5(loginpwd) & "','|*|')") 

        'Add memeber
        'conn.execute("insert into " & db_PREFIX & "member (username,pwd) values('1','" & myMD5("1") & "' )") 
        'conn.execute("insert into " & db_PREFIX & "member (username,pwd) values('test','" & myMD5("test") & "' )")  
        'Add WebSite
        dim fieldNameList, fieldValueList 
        fieldNameList = "webTitle,webKeywords,webDescription,WebSiteUrl,WebSiteBottom" 
        fieldValueList = "'网站标题','网站关键词','网站描述','http://www.baidu.com','版权所有 小云所有 电话：021-66666666 传真：021-8888888<br>联系QQ：313801120 邮箱：313801120@qq.com'" 
        conn.execute("insert into " & db_PREFIX & "website (" & fieldNameList & ") values(" & fieldValueList & ")") 

		'delete menu
 	 	conn.execute("delete from " & db_PREFIX & "ListMenu") 

        'add menu
        dim parentid, title, lableName, content, tempS, url, sIsdisplay, nCount 
        content = getftext(adminDir & "/admin_menu_config.ini") 
        content =  replace(content, vbTab, "    ") 
		if instr(content,vbcrlf)=0 then
			content = replace(content,chr(10), vbCrLf) 		'从github上下载不处理会出有问题
		end if 
        splStr = split(content, vbCrLf) 
        nCount = 0 
        for each s in splStr
            tempS = s 
            s = trim(s) 
            if phptrim(s) <> "" and left(phptrim(s) & " ", 1) <> "#" then
                nCount = nCount + 1                                                             '总数
                if left(tempS, 4) = "    " then
                else
                    parentid = "-1" 
                end if 

                if phptrim(LCase(tempS)) = "end" then
                    exit for 
                elseIf s <> "" then
                    title = mid(s & " ", 1, inStr(s & " ", " ") - 1) 
                    lableName = getStrCut(s, "lablename='", "'", 2) 
                    url = getStrCut(s, "url='", "'", 2) 
                    sIsdisplay = getStrCut(s, "isdisplay='", "'", 2) 
                    if sIsdisplay = "" then
                        sIsdisplay = "1" 
                    end if 
                    if title <> "" then
                        call echo("lablename", lableName) 

                        conn.execute("insert into " & db_PREFIX & "ListMenu (title,parentid,sortrank,lablename,isdisplay,customaurl) values('" & title & "'," & parentid & "," & nCount & ",'" & lableName & "'," & sIsdisplay & ",'" & url & "')") 
                        if parentid = "-1" then
							'【@是jsp显示@】try{
                            rs.open "select * from " & db_PREFIX & "ListMenu where title='" & title & "'", conn, 1, 1 
                            if not rs.EOF then
                                parentid = rs("id") 
                            end if : rs.close 
							'【@是jsp显示@】}catch(Exception e){} 
                        end if 

                        call echo(title, s) 
                        call echo(url, lableName) 
                    end if 
                end if 
            end if 
        next 


    end if 

end sub 
 
'创建数据库
sub createAccess()
    dim loginname, loginpwd 
    loginname = request("loginname") 
    loginpwd = request("loginpwd") 
    call openconn() 

	
    call webData(db_PREFIX, loginname, loginpwd)                                    '导入网站数据
	
	showLayout="step5"			'其它是没有内容的
end sub 
'显示默认
sub displayDefault() 

	showLayout="step1"
    '默认配置
    selectDatabase = "access" 
    dbhost = "localhost" 
    dbuser = "sa" 
    dbpwd = "sa" 
    dbname = "webdata" 
    accessDir = "/data/" 
    msgStr = "创建数据库" 
    isYes = true 
    if request.form("selectDatabase") <> "" then
        selectDatabase = request.form("selectDatabase") 
        dbhost = request.form("dbhost") 
        dbuser = request.form("dbuser") 
        dbpwd = request.form("dbpwd") 
        dbname = request.form("dbname") 
        accessDir = phptrim(request.form("accessDir")) 
        accessPath = accessDir & "/data.mdb" 

        if selectDatabase = "access" then
            if accessDir = "" then
                msgStr = "access数据库安装目录不能为空" 
                isYes = false 
            elseif checkFolder(accessDir) = false then
                msgStr = "access数据库安装目录不存在" 
                isYes = false 
            elseif checkFile(accessPath) = true then
                msgStr = "access数据库存在，使用（" & accessPath & "）" 
                'isYes = false 
            else
                call createMdb(accessPath)       ' 创建数据库
            end if 

        'sqlserver 测试
        else
            accessPath = "" 
			if EDITORTYPE="aspx" then
				connStr = "server='"& dbhost &",1433';database='"& dbname &"';uid='"& dbuser &"';pwd='"& dbpwd &"';" 
			else
            	connStr = " Password = " & dbpwd & "; user id =" & dbuser & "; Initial Catalog =" & dbname & "; data source =" & dbhost & ",1433;Provider = sqloledb;" 
			end if
			 
	
            if checkSqlServer(connStr) = false then
                call echo("connStr", connStr) 
                msgStr = "链接SqlServer数据库失败，检测账号密码" 
                isYes = false 
            end if 
 
        end if 

        '为真则创建数据库
        if isYes = true then 
			if EDITORTYPE="jsp" then
            	content = readFile(configFilePath,"utf-8") 
			else
            	content = getFtext(configFilePath) 
			end if
			tempContent = content 

            '数据库
            startStr = "MDBPath = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1)
				if EDITORTYPE="aspx" then
					if accessPath<>"" then
						accessPath=handlePath(accessPath)
					end if
					accessPath=replace(accessPath,"\","\\")
				else
					accessPath=replace(accessPath,"\\","\")
					accessPath=replace(accessPath,handlePath("/"),"/")	'20170517
				end if 
                replaceStr = startStr & accessPath & endStr 
                content = replace(content, findStr, replaceStr) 
            end if 
            '数据库
            startStr = "databaseType = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1) 
                replaceStr = startStr & selectDatabase & endStr 
                content = replace(content, findStr, replaceStr) 
            end if 



            startStr = "sqlServerHostIP = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1) 
				if EDITORTYPE="aspx" then 
					dbhost=replace(dbhost,"\","\\") 
				end if 
                replaceStr = startStr & dbhost & endStr
                content = replace(content, findStr, replaceStr) 
            end if 
            startStr = "sqlServerUsername = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1) 
                replaceStr = startStr & dbuser & endStr 
                content = replace(content, findStr, replaceStr) 
            end if 
            startStr = "sqlServerPassword = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1) 
                replaceStr = startStr & dbpwd & endStr 
                content = replace(content, findStr, replaceStr) 
            end if 
            startStr = "sqlServerDatabaseName = """ : endStr = """" 
            if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
                findStr = getStrCut(content, startStr, endStr, 1) 
                replaceStr = startStr & dbname & endStr 
                content = replace(content, findStr, replaceStr) 
            end if 

            '更新配置文件
            if tempContent <> content then
				if EDITORTYPE="jsp" then
					call writeToFile(configFilePath, content,"utf-8")
				else
                	call createfile(configFilePath, content) 
            	end if
			end if 

			showLayout="step2" 
        end if 
    end if 
 end sub
'显示导入表面板
sub displayImportTableLayout()
    dim content, startStr, endStr, findStr, replaceStr  
	
	if EDITORTYPE="jsp" then
		content = readFile(configFilePath,"utf-8")
	else
    	content = getFtext(configFilePath) 
    end if
	startStr = "db_PREFIX = """ 
    endStr = """" 
    if instr(content, startStr) > 0 and instr(content, startStr) > 0 then
        findStr = getStrCut(content, startStr, endStr, 1) 
        replaceStr = startStr & request("db_PREFIX") & endStr 
        content = replace(content, findStr, replaceStr) 
		if EDITORTYPE="jsp" then
			call writeToFile(configFilePath, content,"utf-8")
		else
        	call createfile(configFilePath, content) 
    	end if
	end if 

    '删除缓冲文件
    WEB_CACHEFile = replace(replace(WEB_CACHEFile, "[adminDir]", adminDir), "[EDITORTYPE]", EDITORTYPE) 
    'call echo("WEB_CACHEFile",WEB_CACHEFile)
    call deleteFile(WEB_CACHEFile) 
	showLayout="step3" 
end sub
'显示导入数据
sub displayImportData()
    call setsession("adminusername", "PAAJCMS") 
    call setsession("adminflags", "|*|")
    call setsession("adminId", 0) 	'20170517  默认ID为0
	showLayout="step4"
end sub 
%>
<%if showLayout="step1" then%>
<form id="form1" name="form1" method="post" action="?act=displayDefault&webdataDir=<%=webdataDir%>"> 
<div class="pright"> 
    <div class="pr-title"><h3>第一步 选择数据库</h3></div> 
    <label for="radio"><input name="selectDatabase" type="radio" id="radio" value="access" <%=IIF(selectDatabase = "access","checked","")%> onClick="selectInsrtDatabase(this.value);" /> 
    Access数据库</label> 
    <label for="radio2"><input name="selectDatabase" type="radio" id="radio2" value="sqlserver" <%=IIF(selectDatabase = "sqlserver","checked","")%> onClick="selectInsrtDatabase(this.value);" /> 
  SqlServer数据库</label> 
  <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox"> 

        <tbody> 
            <tr> 
                <td class="onetd"><strong>提示：</strong></td> 
                <td> <div class="mytipwrap"><%=msgStr%> &nbsp;<a href="<%=WEB_VIEWURL%>" class="mytip" target="_blank">访问网站首页</a><a href="<%=WEB_ADMINURL%>" class="mytip" target="_blank">登录网站后台</a> </div> </td> 
            </tr>  
        </tbody> 
        <tbody id="access_layout" <%=IIF(selectDatabase = "sqlserver","style='display:none'","")%>>            
            <tr> 
                <td class="onetd"><strong>Access目录：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="accessDir" id="accessDir" type="text" value="<%=accessDir%>" class="input-txt"  > 
                  </div> 
                    <div style="float:left" id="havedbsta"><font color="red"></font></div>                </td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>所需数组：</strong></td> 
                <td> 
               	<%
				call echo("数据库创建组件" , checkObject("ADOX.Catalog"))
				call echo("数据库连接组件" , checkObject("Adodb.connection"))
				%> 
                 
                 </td> 
            </tr> 
        </tbody> 
         
        <tbody id="sqlserver_layout" <%=IIF(selectDatabase = "access","style='display:none'","")%> > 
            <tr> 
                <td class="onetd"><strong>数据库主机：</strong></td> 
                <td><input name="dbhost" id="dbhost" type="text" value="<%=dbhost%>" class="input-txt"> 
                    <small>一般为localhost 或 .\SQLEXPRESS</small>                </td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>数据库用户：</strong></td> 
                <td><input name="dbuser" id="dbuser" type="text" value="<%=dbuser%>" class="input-txt"></td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>数据库密码：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="dbpwd" type="text" value="<%=dbpwd%>" class="input-txt" id="dbpwd" onChange="TestDb()" > 
                  </div> 
                    <div style="float:left" id="dbpwdsta"><font color="red"></font></div>                </td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>数据库名称：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="dbname" id="dbname" type="text" value="<% =dbname%>" class="input-txt" onChange="HaveDB()">
                  请确定数据库名称存在
                    </div> 
                    <div style="float:left" id="havedbsta"><font color="red"></font></div>                </td> 
            </tr> 
        </tbody> 
         
    </table>   
<div class="btn-box">  
        <input name="提交" type="submit" class="btnclick1" onClick="window.location.href='index.php?step=3&accessDir=<%=request("accessDir")%>';" value="继续"> 
    </div> 
</div> 
<script language="javascript"> 
function selectInsrtDatabase(sValue){ 
    //alert(document.getElementById("sqlserver_layout").innerHTML) 
    if(sValue=="access"){  
        document.getElementById("access_layout").style.display = "" 
        document.getElementById("sqlserver_layout").style.display="none"; 
    }else{ 
        document.getElementById("access_layout").style.display="none"; 
        document.getElementById("sqlserver_layout").style.display=""; 
    } 
} 
</script> 
</form> 
<%elseif showLayout="step2" then%> 
<form id="form1" name="form1" method="post" action="?act=displayImportTableLayout&webdataDir=<%=webdataDir%>"> 
<div class="pright"> 
    <div class="pr-title"><h3>第二步 设置登录密码</h3></div> 
    <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox"> 
        <tbody> 
            <tr> 
                <td class="onetd"><strong>提示：</strong></td> 
                <td> <div class="mytipwrap">创建数据库成功。(如果不恢复数据库将退出) &nbsp;<a href="<% =WEB_VIEWURL%>" class="mytip" target="_blank">访问网站首页</a><a href="<% =WEB_ADMINURL%>" class="mytip" target="_blank">登录网站后台</a> </div> </td> 
            </tr>  
            <tr> 
                <td class="onetd"><strong>表前缀：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="db_PREFIX" id="db_PREFIX" type="text" value="xy_" class="input-txt"> 
                  </div> 
                </td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>登录账号：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="loginname"   type="text" value="admin" class="input-txt"> 
                  </div> 
                </td> 
            </tr> 
            <tr> 
                <td class="onetd"><strong>登录密码：</strong></td> 
                <td> 
                  <div style="float:left;margin-right:3px;"><input name="loginpwd"  type="text" value="admin" class="input-txt"> 
                  </div> 
                </td> 
            </tr> 
      </tbody> 
    </table> 
    <div class="btn-box">  
        <input name="提交" type="submit" class="btnclick1"  value="继续"> 
    </div> 
</div> 
</form> 
<%elseif showLayout="step3" then%> 
<form id="form1" name="form1" method="post" action="?act=displayImportData&webdataDir=<%=webdataDir%>"> 
<div class="pright"> 
  <div class="pr-title"><h3>第三步 导入数据</h3></div>     
<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox"> 
        <tbody> 
            <tr> 
                <td class="onetd"><strong>提示：</strong></td> 
                <td> <div class="mytipwrap">创建表完成。(如果不恢复数据库将退出) &nbsp;<a href="<% =WEB_VIEWURL%>" class="mytip" target="_blank">访问网站首页</a><a href="<% =WEB_ADMINURL%>" class="mytip" target="_blank">登录网站后台</a> </div> </td> 
            </tr> 
            <tr> 
              <td colspan="2"> 
              <iframe src="?act=createAccess&loginname=<% =request("loginname")%>&loginpwd=<% =request("loginpwd")%>" width="100%" height="350"></iframe> 
              </td> 
            </tr>  
      </tbody> 
    </table> 
    <div class="btn-box">  
        <input name="提交" type="submit" class="btnclick1"  value="继续"> 
    </div> 
</div> 
</form> 
<%elseif showLayout="step4" then%>  
    <div class="pright"> 
      <div class="pr-title"><h3>第四步 完成</h3></div>     
    <table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox"> 
            <tbody> 
                <tr> 
                    <td class="onetd"><strong>提示：</strong></td> 
                    <td> <div class="mytipwrap">导入数据完成 &nbsp;<a href="<% =WEB_VIEWURL%>" class="mytip" target="_blank">访问网站首页</a><a href="<% =WEB_ADMINURL%>" class="mytip" target="_blank">登录网站后台</a> </div> </td> 
                </tr> 
                <tr> 
                  <td colspan="2"> 
                  <iframe src="<% =WEB_ADMINURL%>?act=setAccess&webdataDir=<%=webdataDir%>&login=out&n=<%=getRnd(9)%>" width="100%" height="350"></iframe> 
                  </td> 
                </tr>  
          </tbody> 
        </table>  
    </div>  
<%end if%>
</body> 
</html> 


