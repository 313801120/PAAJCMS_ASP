<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-02-27
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 


'http://127.0.0.1/函数/ClassAspCode.Asp?act=GetFileFunctionStrList    从这里面获得最新设置


 
'**************************************** 给php用 通用 ****************************************
  
'替换参数值 2014  12 01
function newReplaceValueParam(content, paramName, replaceStr)
    dim startStr, endStr, labelStr, sLen, sNTimeFormat, delHtmlYes, trimYes 
    'ReplaceStr = ReplaceStr & "这里面放上内容在这时碳呀。"
    'ReplaceStr = CStr(ReplaceStr)            '转成字符类型
    if isNul(replaceStr) = true then replaceStr = "" 

    startStr = "[$" & paramName : endStr = "$]" 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        labelStr = strCut(content, startStr, endStr, 1) 
        '删除Html
        delHtmlYes = rParam(labelStr, "DelHtml")                                        '是否删除Html
        if delHtmlYes = "true" then replaceStr = replace(delHtml(replaceStr), "<", "&lt;") 'HTML处理
        '删除两边空格
        trimYes = rParam(labelStr, "Trim")                                              '是否删除两边空格
        if trimYes = "true" then replaceStr = trimVbCrlf(replaceStr) 

        '截取字符处理
        sLen = handleNumber(rParam(labelStr, "Len")) 
        'If nLen<>"" Then ReplaceStr = CutStr(ReplaceStr,nLen,"null")' Left(ReplaceStr,nLen)
        if sLen <> "" then replaceStr = cutStr(replaceStr, cint(sLen), "...")           'Left(ReplaceStr,nLen)


        '时间处理
        sNTimeFormat = rParam(labelStr, "Format_Time")                                  '时间处理值
        if sNTimeFormat <> "" then
            replaceStr = format_Time(replaceStr, cint(sNTimeFormat)) 
        end if 
        content = replace(content, labelStr, replaceStr) 

    end if 
    newReplaceValueParam = content 
end function 

'根据标签找到对应内容
function newRParam(dataCode, action, moduleName)
    dim defaultStr, startStr, endStr 
    defaultStr = rParam(action, moduleName) 
    startStr = "<!--#" & defaultStr & " start#-->" 
    endStr = "<!--#" & defaultStr & " end#-->" 

    if defaultStr <> "" then
        '判断是否存在
        if inStr(dataCode, startStr) > 0 and inStr(dataCode, endStr) > 0 then
            defaultStr = strCut(dataCode, startStr, endStr, 2) 
        else
            startStr = "<!--#" & defaultStr 
            endStr = "#-->" 
            if inStr(dataCode, startStr) > 0 and inStr(dataCode, endStr) > 0 then
                defaultStr = strCut(dataCode, startStr, endStr, 2) 

            'Call Echo("有","StartStr=" & StartStr & ",EndStr=" & EndStr  & ",Default=" & Default)
            end if 
        end if 
    end if 
    newRParam = defaultStr 
end function 




'**************************************** 给php用 通用 **************************************** end


'运行全部动作(20150827)
function getContentAllRunStr(content)
    dim splStr, s, c, tempS 
    splStr = split(content, vbCrLf) 
    for each s in splStr
        if s <> "" then
            tempS = s 
            s = replace(replace(s, chr(10), ""), chr(13), "") '奇怪为什么 s里会有 chr(10)与chr(13) 呢？
            c = c & tempS & "=" & handleContentCode(s, "") & vbCrLf 
        end if 
    next 
    getContentAllRunStr = c 
end function 

'获得内容运行后字符  用法 len(aaaa)  没双引单引
function getContentRunStr(byVal content)
    getContentRunStr = handleContentCode(content, "") 
end function 
'处理内容里有""，则给它删除掉20150329
'检测内容运行后字符
function checkContentRunStr(byVal content)
    checkContentRunStr = handleContentCode(content, "check") 
end function 
'处理双引号
function handleDoubleQuotation(byVal s)
    dim sNew 
    sNew = phpTrim(s) 
    if left(sNew, 1) = """" and right(sNew, 1) = """" then
        s = mid(sNew, 2, len(sNew) - 2) 
    end if 
    handleDoubleQuotation = s 
end function 
'辅助上面
function strDQ(byVal s)
    strDQ = handleDoubleQuotation(s) 
end function 
'处理成数据 20150330
function handleToArray(content)
    dim splStr, i 
    content = strCut(content, "(", ")", 2) 
    'Call Rw(Content)
    splStr = split(content, ",") 
    'Call Rw("<hr>")
    for i = 0 to uBound(splStr)
        splStr(i) = strDQ(splStr(i)) 
    'Call Echo(I,SplStr(I)) 
    next 
    handleToArray = splStr 
end function 
'处理内容里代码 引擎20150324   http://127.0.0.1/函数/ClassAspCode.Asp?act=GetFileFunctionStrList  获得最新    version 1.0
function handleContentCode(byVal content, sType)
    dim startStr, endStr, s, yunStr, splStr, splxx, splContent 
    dim labelName                                                                   '标签名称
    dim nStartStr                                                                   '开始字符

    dim contentLCase                                                                '定义内容小写
    contentLCase = LCase(content)                                                   '内容小写

    if sType = "check" or sType = "检测" then
        handleContentCode = false 
    end if 
	'函数列表 start
	
    splContent = "|getblankhtmlbody|errortext|printpre|foldingmenu|returnren|returnhr|returnrenhr|showerr|printforminfo|dedecmsmsg|handleaccessdata|adsql|checkaccesspass|checksql|createtable|showistable|getrecordcount|tablecount|" 
    splContent = splContent & "|opentable|inserttable|gettabletext|createmdb|deletetable|deletealltable|checktable|existstable|edittable|checkfield|existsfield|editfield|getfieldalt|gettablelist_conn|gettablelist|getfieldlist|getfieldconfiglist|" 
    splContent = splContent & "|getdifferenttablefieldlist|compactdb|editaccesspassword|handlesqlserver|exporttoexcel|importtodatabase|getdatatablelist|settablefield|gettypestring|getproductrndidlist|getfieldvalue|getconnfieldvalue|setfieldvalue|" 
    splContent = splContent & "|filebinaryaddaccess|showaccessstream|gethttppage|handlegethttppage|geturlcontent|gethttpurl|bytestobstr|strcut|getstrcut|cutstr|cutstrnolu|setcuttdstr|cuttdstr|getarraylist|getarray|getarray1|getsplit|" 
    splContent = splContent & "|getsplitcount|agent|specialhtmluploadencryptiondecrypt|encaspcontent|uppercaseorlowercase|enccode|deccode|tounicode|japan|japan26|unjapan26|handlehtml|unhandlehtml|lcaseenc|lcasedec|htmltojs|jstohtml|htmltoasp|" 
    splContent = splContent & "|htmltoaspdim|asptohtml|setfilename|unsetfilename|htmltoascchr|unhtmltoascchr|variabledisplacement|xorenc|xordec|handleevaladddec|evaladddec|vbevaladddec|xdec|tohexcolor|colorrgb|simpleenc|simpledec|encaspevalrun|" 
    splContent = splContent & "|send_email|serversend_email|findaspcode|findaspfunlist|findaspfunctioncode|handlereplacelist|getfunnamedim|leftspace|rightspace|infunadddim|getqq|gettel|getmail|getimgstr|getcssstr|getjsstr|handleaspexecute|" 
    splContent = splContent & "|geturlstr|getpageconfig|getsqlinurl|handlsqlinurl|getscanfunctionnamelist|getfunname|getfundimname|getdimname|getsetname|replacedim|replacedim2|findfolderrepeatfunction|findfilerepeatfunction|findcontentrepeatfunction|" 
    splContent = splContent & "|handlecontentrepeatfunction|replace2|allreplace|replaceonenolu|optimizeaspcode|aspcodeconfusion|aspvariablegarbled|lucasegarbled|aspfunctionconfusion|getaspfunnamelist|aspcodesafe|cutdakuohao|cutxiankuohao|" 
    splContent = splContent & "|cutfunctionvvalue|getstrintcontentnumb|getkeyword|gethandlekeyword|getchinastr|checkstr|regexp_checkstr|existsstr|regexp_instr|regexp_replace|replacecolor|delhtml|regexp_delhtml|getaurltitlelist|getaurl|" 
    splContent = splContent & "|getatitle|getaurltitle|getcontentaurllist|regexp_getimgurllist|getimg|getimgjpg|getimggif|getimgbmp|getaspfunctionlist|getaspdimfunction|strlength|regexpgetlenght|stringlength|getcsshref|getjssrc|delhtmlscript|" 
    splContent = splContent & "|getmeta|delhtmlnote|delhtmlmynote|getstrimglist|getemaillist|oflink_lcasetag|scripthtml|getregexp|getimgjsurl|regexpgetstr|newgetpingip|getpingip|getpinginfo|regexptest|getregexpstr|regexpreplacecssdir|" 
    splContent = splContent & "|deletenullrow|isstrtransferred|getbl|getcountstep|getchina|ischina|checkchina|phprand|phprnd|deleterepeatstr|replacen|fillzero|caseinsensitivereplace|array_sort|handlezipsize|getrnd|mathrandom|getrndaz|" 
    splContent = splContent & "|getrand|copystrnumb|str_repeat|copystr|contentaddtab|deleteendstr|tonumber|handlenumber|strdrawint|getfirstnegativenumber|checknumbertype|checkstrisnumbertype|handlenumbertype|getnumber|checknumb|getdiannumb|" 
    splContent = splContent & "|getcountpage|getpagenumb|getcaihandlecountpage|getcaisortcountpage|minmaxbetween|getstrfiletype|val|strlen|addindent|getnumberbeforestr|makepassword|rndcode|getrandomphonenumber|lenstr|tostring|remotenumber|" 
    splContent = splContent & "|handlehavestr|havelowercase|haveuppercase|havenumber|havechina|saveremotefile|downfile|popupdownfile|xmlgetstatus|xmlpost|utfxmlpost|handlexmlpost|xmlget|handlexmlget|cookiespost|handlecookiespost|getremotefilesize|" 
    splContent = splContent & "|downremotefile|imagestobase64|base64toimages|getsqlincontent|handlesqlin|sqlinurl|handlurlcanshu|handlestrcharset|urlencoding|getwinhttp|format_time|gethandledate|getdaystatus|printtimevalue|printasktime|" 
    splContent = splContent & "|gettimerset|calculationtimer|gettimer|vbgettimer|vbechotimer|vbruntimer|getthisurlfileparam|getthisurlnoparam|getgotourl|getgotourlnoparam|getgotourlnofilename|getip2|getip|getservicerip|getremoteip|handlehttpurl|" 
    splContent = splContent & "|handlefileurl|handleurlcomplete|urladdhttpurl|getport|webdomain|host|httphost|gethost|geturl|getthisurl|getthisurlnofilename|getwebsitename|getwebsite|checkurl|checkhttpurl|isurl|ishttpurl|fullhttpurl|" 
    splContent = splContent & "|batchfullhttpurl|urljianjiehandle|urltoasc|getwebtitle|getcontentahref|getcontentimgsrc|handleconenturl|replacecontentjsdir|handlelink|getendurlhandlename|geturllistinwebsitelist|getthisurlfilename|handlewebhtmlimg|" 
    splContent = splContent & "|handlewebcssimg|batchhandleurlintegrity|replacecontentimagepath|getcsslinklist|gethtmlbackgroundurllist|getimgurllist|getlinkurl|localurlorremoteurl|checkremoteurl|asailinkadd|asailinkdel|iswww|geturladdtoparam|" 
    splContent = splContent & "|groupurl|handlepostcookiesparame|remotehttpurlparameter|checkurlname|handlehttpurlarray|remotejscssparam|handleremotejscssparam|batchhandlehttpurlcomplete|iswebsite|issonwebsite|handleiswebsite|getcontenturllist|" 
    splContent = splContent & "|handlegetcontenturllist|getinchain|getfiletext|getftext|getfiletype|checkfile|existsfile|createfile|savefile|createaddfile|createaddupfile|addtofile|savebytefile|createbytefile|movefile|editfile|copyfile|" 
    splContent = splContent & "|deletefile|delselfile|delfilelist|batchdelfile|deleteappointfilelist|getfilesize|filesize|getfsize|getfilecreatedate|filecreatedate|getfileeditdate|fileeditdate|getfileedittime|getdatelastaccessed|modifyfilemodifydate|" 
    splContent = splContent & "|editfileeditdate|existszhidufile|getfileattribute|editfileattribute|existsfileday|checkfileday|createfolder|createdirfolder|createdomain|createdomainfolder|movefolder|editfolder|copyfolder|deletefolder|" 
    splContent = splContent & "|getfoldersize|checkfolder|existsfolder|getfoldercreatedate|getfoldereditdate|getfoldertime|getfolderattribute|editfolderattribute|existsfolderday|checkfolderday|getfilefolderlist|getdirjpglist|getdirpnglist|" 
    splContent = splContent & "|getdirinilist|getdirtxtlist|getdirjslist|getdircsslist|getdirhtmllist|getdirasplist|getdirphplist|getdirfilelist|getdirjpgnamelist|getdirpngnamelist|getdirininamelist|getdirtxtnamelist|getdirjsnamelist|" 
    splContent = splContent & "|getdircssnamelist|getdirhtmlnamelist|getdiraspnamelist|getdirphpnamelist|getdirfilenamelist|getdirallfilelist|getdirallfilenamelist|getdirfolderlist|getdirfoldernamelist|getdirallfolderlist|getdirallfoldernamelist|" 
    splContent = splContent & "|getfilefoldernumb|getdirfoldernumb|getdirfoldersize|getdirfilenumb|getdirallfilenumb|getdirallfoldernumb|getfolderlist|getdirfilesort|getdirfoldersort|deldirallfile|deleteallfile|getupdatefile|handlepath|" 
    splContent = splContent & "|realpath|fullpath|handlerelativepath|dirname|isobjinstalled|existsobject|getfileobject|getdictionary|getstream|showattributetip|printspacevalue|printspacesize|printfilesize|getdeskpath|createurlshortcut|" 
    splContent = splContent & "|createfoldershortcut|finddirfilename|findsubdirfilename|getarticlefillcontent|getfilepathname|getextension|removefiledir|basename|getfileextname|handlefilepatharray|getstrfilename|getfilename|getfileattr|" 
    splContent = splContent & "|getfileext|getparentpath|checkurlfilenameparam|isucase|islcase|checkerr|strtotruefalse|strtruefalsetoint|checkvbcrlf|checkbr|isparity|checkeval|replace_sqltext|chkpost|ismail|isvalidemail|iscode|checknumber|" 
    splContent = splContent & "|isnumber|checkabc|getlen|checktime|isnul|foundinarr|displayhtml|undisplayhtml|handleclosehtml|handlehtmlahref|addtooraddhttpurl|gethtmllablename|removeblanklines|checkcode|getfilecharset|seniorcheckcode|" 
    splContent = splContent & "|readfile|getstext|getftextutf|writetofile|createfilegbk|createfileutf|getugtext|checkcontentcharset|checkfilecode|getfilebinarycontent|binaryreadfile|readbinary|decsavebinary|writebinary|getfilecharset2|" 
    splContent = splContent & "|bintohex|base64encode|base64decode|asphelptip|handlewrite|jiantitabfanti|jiyicireplace|cireplace|testgetpy|pinyin|pinyin2|entocompany|remotecontentjinghao|remotearrayjinghao|geteachstraddvalue|getvalueinarrayid|" 
    splContent = splContent & "|checkvalueinarray|deleterepeatarray|deleterepeatcontent|getarraycount|randomshow|arrayrandomshow|printarray|echoarray|str_shuffle|characterupset|str_split|removenullvaluearray|removerepeatvaluearray|handlearray|" 
    splContent = splContent & "|arraytostring|testarraydata|selectsort|quicksort|bubblesort|insertsort|requicksort|arrayreverse|responsearray|max|min|testtwodimensionalarray|handlesplitarray|twodimensionalarrayasc|twodimensionalarraydesc|" 
    splContent = splContent & "|csscompression|deletecssnote|uncsscompression|trimvbcrlf|phptrim|phpltrim|phprtrim|trimvbtab|ltrimvbtab|rtrimvbtab|gethtmlcssstyle|handlecleancss|removeexcessrow|cssaddtostyle|checkcssstyle|cssstyleaddtoparam|" 
    splContent = splContent & "|findcssstyle|handlecutcsscode|handlecutdivcode|gethtmlbackgroundimglist|gethandlewebhtmllink|getcsslisturllist|handlereadcsscontent|handlecsspx|handlehtmlstyle|utf2gb|checkutftogb2312|convchinese|utf8_urlencoding|" 
    splContent = splContent & "|toutf8|chinese2unicode|isvalidhex|escape|unescape|chinesetounicode|unicodetochinese|urlencode|urldecode|toUTFChar|encdecurl|encurl|decurl|infocolor|imgerror|targetstr|atarget|atitle|imgalt|imgtitlealt|arel|" 
    splContent = splContent & "|styleclass|textfontb|textfontcolor|fontcolorfontb|getdefaultfilename|ahref|imgsrc|selecttarget|selectfontcolor|selectsex|selectsessioncookies|showselectlist|showarticleliststyle|showarticleinfostyle|handlearticleliststyleorinfostyle|" 
    splContent = splContent & "|showwebmoduleskins|showradiolist|inputcheckbox|inputcheckbox3|handleinputcheckbox|inputtext|inputtext2|inputlefttext|inputlefttexthelptextright|inputlefttextcontent|inputrighttextcontent|handleinputleftrighttextcontent|" 
    splContent = splContent & "|inputleftpasstext|inputleftpasstextcontent|inputhiddentext|inputtextarea|inputhiddentextarea|handleinputhiddentextarea|showselectdirlist|inputdisabled|inputaddalt|webtitle_inputtextarea|webkeywords_inputtext|" 
    splContent = splContent & "|webdescription_inputtextarea|foldername_inputtext|filename_inputtext|templatepath_inputtext|clickpinyinhtmlstr|showfontcolorfontb|showsort|showwebnavtype|showonhtml|showthrough|showrecommend|showonoffimg|" 
    splContent = splContent & "|newshowonoffimg|controldialogcss|batchdeletetempstr|htmlformatting|handlehtmlformatting|formatting|ziphtml|checkhtmlformatting|newreplacevalueparam|newrparam|getcontentallrunstr|getcontentrunstr|checkcontentrunstr|" 
    splContent = splContent & "|handledoublequotation|strdq|handletoarray|handlecontentcode|handleinmodule|clearrparam|atrparam|rparam|handlerparam|getconfigblock|getconfigblock2|getconfigblock3|getcutconfigblock|getconfigcontentblock|" 
    splContent = splContent & "|getconfigfileblock|setconfigfileblock|delconfigblock|delconfigblock2|delconfigblock3|delcutconfigblock|getfileparamvalue|setfileparamvalue|handlegetsetfileparamevalue|setrparam|addreplacerparam|deletestrcut|" 
    splContent = splContent & "|replacecontentmodule|replacecontentrowmodule|handleconfigfile|getrparam|getdefaultvalue|addmodulereplacearray|modulefindcontent|batchfindfunction|getfindfilefunction|getaspcode|handlegetaspcode|delaspnote|" 
'函数列表 end
	

    '系统自带 不可删除
    splContent = splContent & "|now|date|len|left|right|mid|lcase|ucase|ltrim|rtrim|trim|" 


    splStr = split(splContent, "|") 
    for each labelName in splStr
        if labelName <> "" then
            startStr = LCase(labelName & "(") : endStr = ")" 
            nStartStr = inStr(" " & contentLCase, " " & startStr) 
            if nStartStr > 0 and inStr(contentLCase, endStr) > 0 then
                if sType = "check" or sType = "检测" then
                    handleContentCode = true                 '为检测 有函数  退出
                    exit function 
                end if 
                exit for 
            end if 
        end if 
    next 
    if labelName = "" then
        if sType <> "check" and sType <> "检测" then
            handleContentCode = content 
        end if 
        exit function 
    end if 

    labelName = LCase(labelName)                                                    '小写

    startStr = labelName & "(" : endStr = ")" 
    nStartStr = inStr(contentLCase, startStr) 
    '记录查的原值，不分大小写
    yunStr = mid(content, inStr(contentLCase, startStr)) 
    yunStr = mid(yunStr, 1, inStrRev(yunStr, endStr)) 

    s = mid(content, inStr(contentLCase, startStr) + len(startStr)) 

    s = mid(s, 1, inStrRev(s, endStr) - 1) 
    '多次处理 注意和下面的判断
    if checkContentRunStr(content) = true then
        s = getContentRunStr(s) 
    end if 
    '处理内容里有""，则给它删除掉20150329
    s = handleDoubleQuotation(s) 

    splxx = split(s & ",1", ",") 

    '-----------------------------------------------------------

    '函数判断 start
    if labelName = "getblankhtmlbody" then
        s = getBlankHtmlBody(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "errortext" then
        s = errorText(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "printpre" then
        s = printPre(strDQ(s)) 
    elseIf labelName = "foldingmenu" then
        s = foldingMenu(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "returnren" then
        s = returnRen(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "returnhr" then
        s = returnHr() 
    elseIf labelName = "returnrenhr" then
        s = returnRenHr() 
    elseIf labelName = "showerr" then
        call showErr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "printforminfo" then
        call printFormInfo() 
    elseIf labelName = "dedecmsmsg" then
        s = dedeCMSMsg() 
    elseIf labelName = "handleaccessdata" then
        s = handleAccessData(strDQ(s)) 
    elseIf labelName = "adsql" then
        s = aDSql(strDQ(s)) 
    elseIf labelName = "checkaccesspass" then
        s = checkAccessPass(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checksql" then
        s = checkSql(strDQ(s)) 
    elseIf labelName = "createtable" then
        s = createTable(strDQ(s)) 
    elseIf labelName = "showistable" then
        s = showIsTable(strDQ(s)) 
    elseIf labelName = "getrecordcount" then
        s = getRecordCount(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "tablecount" then
        s = tableCount(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "opentable" then
        s = openTable(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "inserttable" then
        s = insertTable(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "gettabletext" then
        s = getTableText(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "createmdb" then
        s = createMdb(strDQ(s)) 
    elseIf labelName = "deletetable" then
        s = deleteTable(strDQ(s)) 
    elseIf labelName = "deletealltable" then
        s = deleteAllTable(strDQ(s)) 
    elseIf labelName = "checktable" then
        s = checkTable(strDQ(s)) 
    elseIf labelName = "existstable" then
        s = existsTable(strDQ(s)) 
    elseIf labelName = "edittable" then
        s = editTable(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "checkfield" then
        s = checkField(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "existsfield" then
        s = existsField(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "editfield" then
        s = editField(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getfieldalt" then
        s = getFieldAlt(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gettablelist_conn" then
        s = getTableList_Conn(strDQ(s)) 
    elseIf labelName = "gettablelist" then
        s = getTableList() 
    elseIf labelName = "getfieldlist" then
        s = getFieldList(strDQ(s)) 
    elseIf labelName = "getfieldconfiglist" then
        s = getFieldConfigList(strDQ(s)) 
    elseIf labelName = "getdifferenttablefieldlist" then
        s = getDifferentTableFieldList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "compactdb" then
        s = compactDB(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "editaccesspassword" then
        s = editAccessPassWord(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handlesqlserver" then
        s = handleSqlServer(strDQ(s)) 
    elseIf labelName = "exporttoexcel" then
        call exportToExcel(strDQ(s)) 
    elseIf labelName = "importtodatabase" then
        call importToDatabase() 
    elseIf labelName = "getdatatablelist" then
        s = getDataTableList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "settablefield" then
        s = setTableField(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gettypestring" then
        s = getTypeString(strDQ(s)) 
    elseIf labelName = "getproductrndidlist" then
        s = getProductRndIdList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfieldvalue" then
        s = getFieldValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getconnfieldvalue" then
        s = getConnFieldValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "setfieldvalue" then
        s = setFieldValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "filebinaryaddaccess" then
        s = fileBinaryAddAccess(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5))) 
    elseIf labelName = "showaccessstream" then
        s = showAccessStream(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "gethttppage" then
        s = getHttpPage(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlegethttppage" then
        s = handleGetHttpPage(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "geturlcontent" then
        s = getURLContent(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gethttpurl" then
        s = getHttpUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "bytestobstr" then
        s = bytesToBstr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "strcut" then
        s = strCut(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "getstrcut" then
        s = getStrCut(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "cutstr" then
        s = cutStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "cutstrnolu" then
        s = cutStrNOLU(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "setcuttdstr" then
        s = setCutTDStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "cuttdstr" then
        s = cutTDStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getarraylist" then
        s = getArrayList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5))) 
    elseIf labelName = "getarray" then
        s = getArray(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "getarray1" then
        s = getArray1(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "getsplit" then
        s = getSplit(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getsplitcount" then
        s = getSplitCount(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "agent" then
        s = agent(strDQ(s)) 
    elseIf labelName = "specialhtmluploadencryptiondecrypt" then
        s = specialHtmlUploadEncryptionDecrypt(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "encaspcontent" then
        s = encAspContent(strDQ(s)) 
    elseIf labelName = "uppercaseorlowercase" then
        s = upperCaseORLowerCase(strDQ(s)) 
    elseIf labelName = "enccode" then
        s = encCode(strDQ(s)) 
    elseIf labelName = "deccode" then
        s = decCode(strDQ(s)) 
    elseIf labelName = "tounicode" then
        s = toUnicode(strDQ(s)) 
    elseIf labelName = "japan" then
        s = japan(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "japan26" then
        s = japan26(strDQ(s)) 
    elseIf labelName = "unjapan26" then
        s = unJapan26(strDQ(s)) 
    elseIf labelName = "handlehtml" then
        s = handleHTML(strDQ(s)) 
    elseIf labelName = "unhandlehtml" then
        s = unHandleHTML(strDQ(s)) 
    elseIf labelName = "lcaseenc" then
        s = lcaseEnc(strDQ(s)) 
    elseIf labelName = "lcasedec" then
        s = lcaseDec(strDQ(s)) 
    elseIf labelName = "htmltojs" then
        s = htmlToJs(strDQ(s)) 
    elseIf labelName = "jstohtml" then
        s = jsToHtml(strDQ(s)) 
    elseIf labelName = "htmltoasp" then
        s = htmlToAsp(strDQ(s)) 
    elseIf labelName = "htmltoaspdim" then
        s = htmlToAspDim(strDQ(s)) 
    elseIf labelName = "asptohtml" then
        s = aspToHtml(strDQ(s)) 
    elseIf labelName = "setfilename" then
        s = setFileName(strDQ(s)) 
    elseIf labelName = "unsetfilename" then
        s = unSetFileName(strDQ(s)) 
    elseIf labelName = "htmltoascchr" then
        s = HTMLToAscChr(strDQ(s)) 
    elseIf labelName = "unhtmltoascchr" then
        s = unHTMLToAscChr(strDQ(s)) 
    elseIf labelName = "variabledisplacement" then
        s = variableDisplacement(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "xorenc" then
        s = xorEnc(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "xordec" then
        s = xorDec(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handleevaladddec" then
        s = handleEvalAddDec(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "evaladddec" then
        s = evalAddDec(strDQ(s)) 
    elseIf labelName = "vbevaladddec" then
        s = VBEvalAddDec(strDQ(s)) 
    elseIf labelName = "xdec" then
        s = XDec(strDQ(s)) 
    elseIf labelName = "tohexcolor" then
        s = toHexColor(strDQ(s)) 
    elseIf labelName = "colorrgb" then
        s = colorRGB(strDQ(s)) 
    elseIf labelName = "simpleenc" then
        s = simpleEnc(strDQ(s)) 
    elseIf labelName = "simpledec" then
        s = simpleDec(strDQ(s)) 
    elseIf labelName = "encaspevalrun" then
        s = encAspEvalRun(strDQ(s)) 
    elseIf labelName = "send_email" then
        s = send_Email(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "serversend_email" then
        s = serverSend_Email(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5))) 
    elseIf labelName = "findaspcode" then
        s = findAspCode(strDQ(s)) 
    elseIf labelName = "findaspfunlist" then
        s = findAspFunList(strDQ(s)) 
    elseIf labelName = "findaspfunctioncode" then
        s = findAspFunctionCode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlereplacelist" then
        s = handleReplaceList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getfunnamedim" then
        s = getFunNameDim(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "leftspace" then
        s = leftSpace(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "rightspace" then
        s = rightSpace(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "infunadddim" then
        s = inFunAddDim(strDQ(s)) 
    elseIf labelName = "getqq" then
        s = getQQ(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gettel" then
        s = getTel(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getmail" then
        s = getMail(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getimgstr" then
        s = getImgStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getcssstr" then
        s = getCssStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getjsstr" then
        s = getJsStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handleaspexecute" then
        s = handleAspExecute(strDQ(s)) 
    elseIf labelName = "geturlstr" then
        s = getUrlStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "getpageconfig" then
        s = getPageConfig(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getsqlinurl" then
        s = getSqlInUrl(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handlsqlinurl" then
        s = handlSqlInUrl(strDQ(s)) 
    elseIf labelName = "getscanfunctionnamelist" then
        s = getScanFunctionNameList(strDQ(s)) 
    elseIf labelName = "getfunname" then
        s = getFunName(strDQ(s)) 
    elseIf labelName = "getfundimname" then
        s = getFunDimName(strDQ(s)) 
    elseIf labelName = "getdimname" then
        s = getDimName(strDQ(s)) 
    elseIf labelName = "getsetname" then
        s = getSetName(strDQ(s)) 
    elseIf labelName = "replacedim" then
        s = replaceDim(strDQ(s)) 
    elseIf labelName = "replacedim2" then
        s = replaceDim2(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "findfolderrepeatfunction" then
        s = findFolderRepeatFunction(strDQ(s)) 
    elseIf labelName = "findfilerepeatfunction" then
        s = findFileRepeatFunction(strDQ(s)) 
    elseIf labelName = "findcontentrepeatfunction" then
        s = findContentRepeatFunction(strDQ(s)) 
    elseIf labelName = "handlecontentrepeatfunction" then
        s = handleContentRepeatFunction(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "replace2" then
        s = replace2(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "allreplace" then
        s = allReplace(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "replaceonenolu" then
        s = replaceOneNOLU(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "optimizeaspcode" then
        s = optimizeAspCode(strDQ(s)) 
    elseIf labelName = "aspcodeconfusion" then
        s = aspCodeConfusion(strDQ(s)) 
    elseIf labelName = "aspvariablegarbled" then
        s = ASPVariableGarbled(strDQ(s)) 
    elseIf labelName = "lucasegarbled" then
        s = LUCaseGarbled(strDQ(s)) 
    elseIf labelName = "aspfunctionconfusion" then
        s = aspFunctionConfusion(strDQ(s)) 
    elseIf labelName = "getaspfunnamelist" then
        s = getAspFunNameList(strDQ(s)) 
    elseIf labelName = "aspcodesafe" then
        s = aspCodeSafe(strDQ(s)) 
    elseIf labelName = "cutdakuohao" then
        s = cutDaKuoHao(strDQ(s)) 
    elseIf labelName = "cutxiankuohao" then
        s = cutXianKuoHao(strDQ(s)) 
    elseIf labelName = "cutfunctionvvalue" then
        s = cutFunctionvValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getstrintcontentnumb" then
        s = getStrIntContentNumb(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getkeyword" then
        s = getKeyWord(strDQ(s)) 
    elseIf labelName = "gethandlekeyword" then
        s = getHandleKeyWord(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getchinastr" then
        s = getChinaStr(strDQ(s)) 
    elseIf labelName = "checkstr" then
        s = checkStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "regexp_checkstr" then
        s = regExp_CheckStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "existsstr" then
        s = existsStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "regexp_instr" then
        s = regExp_Instr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "regexp_replace" then
        s = regExp_Replace(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "replacecolor" then
        s = replaceColor(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "delhtml" then
        s = delHtml(strDQ(s)) 
    elseIf labelName = "regexp_delhtml" then
        s = regExp_DelHtml(strDQ(s)) 
    elseIf labelName = "getaurltitlelist" then
        s = getAUrlTitleList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getaurl" then
        s = getAURL(strDQ(s)) 
    elseIf labelName = "getatitle" then
        s = getATitle(strDQ(s)) 
    elseIf labelName = "getaurltitle" then
        s = getAURLTitle(strDQ(s)) 
    elseIf labelName = "getcontentaurllist" then
        s = getContentAUrlList(strDQ(s)) 
    elseIf labelName = "regexp_getimgurllist" then
        s = regExp_GetImgUrlList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getimg" then
        s = getIMG(strDQ(s)) 
    elseIf labelName = "getimgjpg" then
        s = getIMGJpg(strDQ(s)) 
    elseIf labelName = "getimggif" then
        s = getIMGGif(strDQ(s)) 
    elseIf labelName = "getimgbmp" then
        s = getIMGBmp(strDQ(s)) 
    elseIf labelName = "getaspfunctionlist" then
        s = getASPFunctionList(strDQ(s)) 
    elseIf labelName = "getaspdimfunction" then
        s = getASPDimFunction(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "strlength" then
        s = strLength(strDQ(s)) 
    elseIf labelName = "regexpgetlenght" then
        s = regExpGetLenght(strDQ(s)) 
    elseIf labelName = "stringlength" then
        s = stringLength(strDQ(s)) 
    elseIf labelName = "getcsshref" then
        s = getCssHref(strDQ(s)) 
    elseIf labelName = "getjssrc" then
        s = getJsSrc(strDQ(s)) 
    elseIf labelName = "delhtmlscript" then
        s = delHTMLScript(strDQ(s)) 
    elseIf labelName = "getmeta" then
        s = getMeta(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "delhtmlnote" then
        s = delHtmlNote(strDQ(s)) 
    elseIf labelName = "delhtmlmynote" then
        s = delHtmlMyNote(strDQ(s)) 
    elseIf labelName = "getstrimglist" then
        s = getStrImgList(strDQ(s)) 
    elseIf labelName = "getemaillist" then
        s = getEMailList(strDQ(s)) 
    elseIf labelName = "oflink_lcasetag" then
        s = oflink_lcasetag(strDQ(s)) 
    elseIf labelName = "scripthtml" then
        s = scriptHtml(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getregexp" then
        s = getRegExp(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "getimgjsurl" then
        s = getImgJsUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "regexpgetstr" then
        s = regExpGetStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "newgetpingip" then
        s = newGetPingIP(strDQ(s)) 
    elseIf labelName = "getpingip" then
        s = getPingIP(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getpinginfo" then
        s = getPingInfo(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "regexptest" then
        s = regExpTest(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getregexpstr" then
        s = getRegExpStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "regexpreplacecssdir" then
        s = regExpReplaceCssDir(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deletenullrow" then
        s = deleteNullRow(strDQ(s)) 
    elseIf labelName = "isstrtransferred" then
        s = isStrTransferred(strDQ(s)) 
    elseIf labelName = "getbl" then
        s = getBL(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getcountstep" then
        s = getCountStep(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "getchina" then
        s = getChina(strDQ(s)) 
    elseIf labelName = "ischina" then
        s = isChina(strDQ(s)) 
    elseIf labelName = "checkchina" then
        s = checkChina(strDQ(s)) 
    elseIf labelName = "phprand" then
        s = pHPRand(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "phprnd" then
        s = pHPRnd(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deleterepeatstr" then
        s = deleteRepeatStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "replacen" then
        s = replaceN(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "fillzero" then
        s = fillZero(strDQ(s)) 
    elseIf labelName = "caseinsensitivereplace" then
        s = caseInsensitiveReplace(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "array_sort" then
        s = array_Sort(strDQ(s)) 
    elseIf labelName = "handlezipsize" then
        s = handleZipSize(strDQ(s)) 
    elseIf labelName = "getrnd" then
        s = getRnd(strDQ(s)) 
    elseIf labelName = "mathrandom" then
        s = mathRandom() 
    elseIf labelName = "getrndaz" then
        s = getRndAZ(strDQ(s)) 
    elseIf labelName = "getrand" then
        s = getRand(strDQ(s)) 
    elseIf labelName = "copystrnumb" then
        s = copyStrNumb(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "str_repeat" then
        s = str_Repeat(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "copystr" then
        s = copyStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "contentaddtab" then
        s = contentAddTab(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deleteendstr" then
        s = deleteEndStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "tonumber" then
        s = toNumber(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlenumber" then
        s = handleNumber(strDQ(s)) 
    elseIf labelName = "strdrawint" then
        s = strDrawInt(strDQ(s)) 
    elseIf labelName = "getfirstnegativenumber" then
        s = getFirstNegativeNumber(strDQ(s)) 
    elseIf labelName = "checknumbertype" then
        s = checkNumberType(strDQ(s)) 
    elseIf labelName = "checkstrisnumbertype" then
        s = checkStrIsNumberType(strDQ(s)) 
    elseIf labelName = "handlenumbertype" then
        s = handleNumberType(strDQ(s)) 
    elseIf labelName = "getnumber" then
        s = getNumber(strDQ(s)) 
    elseIf labelName = "checknumb" then
        s = checkNumb(strDQ(s)) 
    elseIf labelName = "getdiannumb" then
        s = getDianNumb(strDQ(s)) 
    elseIf labelName = "getcountpage" then
        s = getCountPage(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getpagenumb" then
        s = getPageNumb(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getcaihandlecountpage" then
        s = getCaiHandleCountPage(strDQ(s)) 
    elseIf labelName = "getcaisortcountpage" then
        s = getCaiSortCountPage(strDQ(s)) 
    elseIf labelName = "minmaxbetween" then
        s = minMaxBetween(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getstrfiletype" then
        s = getStrFileType(strDQ(s)) 
    elseIf labelName = "val" then
        s = val(strDQ(s)) 
    elseIf labelName = "strlen" then
        s = strLen(strDQ(s)) 
    elseIf labelName = "addindent" then
        s = addIndent(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getnumberbeforestr" then
        s = getNumberBeforeStr(strDQ(s)) 
    elseIf labelName = "makepassword" then
        s = makePassword(strDQ(s)) 
    elseIf labelName = "rndcode" then
        s = rndcode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getrandomphonenumber" then
        s = getRandomPhoneNumber(strDQ(s)) 
    elseIf labelName = "lenstr" then
        s = lenStr(strDQ(s)) 
    elseIf labelName = "tostring" then
        s = toString(strDQ(s)) 
    elseIf labelName = "remotenumber" then
        s = remoteNumber(strDQ(s)) 
    elseIf labelName = "handlehavestr" then
        s = handleHaveStr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "havelowercase" then
        s = haveLowerCase(strDQ(s)) 
    elseIf labelName = "haveuppercase" then
        s = haveUpperCase(strDQ(s)) 
    elseIf labelName = "havenumber" then
        s = haveNumber(strDQ(s)) 
    elseIf labelName = "havechina" then
        s = haveChina(strDQ(s)) 
    elseIf labelName = "saveremotefile" then
        s = saveRemoteFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "downfile" then
        s = downFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "popupdownfile" then
        s = popupDownFile(strDQ(s)) 
    elseIf labelName = "xmlgetstatus" then
        s = XMLGetStatus(strDQ(s)) 
    elseIf labelName = "xmlpost" then
        s = xmlPost(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "utfxmlpost" then
        s = UTFXMLPost(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlexmlpost" then
        s = handleXMLPost(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "xmlget" then
        s = xmlGet(strDQ(s)) 
    elseIf labelName = "handlexmlget" then
        s = handleXmlGet(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "cookiespost" then
        s = cookiesPost(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlecookiespost" then
        s = handleCookiesPost(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getremotefilesize" then
        s = getRemoteFileSize(strDQ(s)) 
    elseIf labelName = "downremotefile" then
        s = downRemoteFile(strDQ(s)) 
    elseIf labelName = "imagestobase64" then
        s = imagesToBase64(strDQ(s)) 
    elseIf labelName = "base64toimages" then
        s = base64ToImages(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getsqlincontent" then
        s = getSqlInContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handlesqlin" then
        s = handleSqlIn(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "sqlinurl" then
        s = sqlInUrl(strDQ(s)) 
    elseIf labelName = "handlurlcanshu" then
        s = handlUrlCanShu(strDQ(s)) 
    elseIf labelName = "handlestrcharset" then
        s = handleStrCharSet(strDQ(s)) 
    elseIf labelName = "urlencoding" then
        s = URLEncoding(strDQ(s)) 
    elseIf labelName = "getwinhttp" then
        s = getWinHttp(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "format_time" then
        s = format_Time(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gethandledate" then
        s = getHandleDate(strDQ(s)) 
    elseIf labelName = "getdaystatus" then
        s = getDayStatus(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "printtimevalue" then
        s = printTimeValue(strDQ(s)) 
    elseIf labelName = "printasktime" then
        s = printAskTime(strDQ(s)) 
    elseIf labelName = "gettimerset" then
        s = getTimerSet() 
    elseIf labelName = "calculationtimer" then
        s = calculationTimer(strDQ(s)) 
    elseIf labelName = "gettimer" then
        s = getTimer() 
    elseIf labelName = "vbgettimer" then
        s = vBGetTimer() 
    elseIf labelName = "vbechotimer" then
        s = vBEchoTimer() 
    elseIf labelName = "vbruntimer" then
        s = vBRunTimer(strDQ(s)) 
    elseIf labelName = "getthisurlfileparam" then
        s = getThisUrlFileParam() 
    elseIf labelName = "getthisurlnoparam" then
        s = getThisUrlNoParam() 
    elseIf labelName = "getgotourl" then
        s = getGoToUrl() 
    elseIf labelName = "getgotourlnoparam" then
        s = getGoToUrlNoParam() 
    elseIf labelName = "getgotourlnofilename" then
        s = getGoToUrlNoFileName() 
    elseIf labelName = "getip2" then
        s = getIP2() 
    elseIf labelName = "getip" then
        s = getIP() 
    elseIf labelName = "getservicerip" then
        s = getServicerIP() 
    elseIf labelName = "getremoteip" then
        s = getRemoteIP() 
    elseIf labelName = "handlehttpurl" then
        s = handleHttpUrl(strDQ(s)) 
    elseIf labelName = "handlefileurl" then
        s = handleFileUrl(strDQ(s)) 
    elseIf labelName = "handleurlcomplete" then
        s = handleUrlComplete(strDQ(s)) 
    elseIf labelName = "urladdhttpurl" then
        s = urlAddHttpUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getport" then
        s = getPort() 
    elseIf labelName = "webdomain" then
        s = webDoMain() 
    elseIf labelName = "host" then
        s = host() 
    elseIf labelName = "httphost" then
        s = httpHost() 
    elseIf labelName = "gethost" then
        s = getHost() 
    elseIf labelName = "geturl" then
        s = getUrl() 
    elseIf labelName = "getthisurl" then
        s = getThisUrl() 
    elseIf labelName = "getthisurlnofilename" then
        s = getThisUrlNoFileName() 
    elseIf labelName = "getwebsitename" then
        s = getWebSiteName(strDQ(s)) 
    elseIf labelName = "getwebsite" then
        s = getWebSite(strDQ(s)) 
    elseIf labelName = "checkurl" then
        s = checkUrl(strDQ(s)) 
    elseIf labelName = "checkhttpurl" then
        s = checkHttpUrl(strDQ(s)) 
    elseIf labelName = "isurl" then
        s = isUrl(strDQ(s)) 
    elseIf labelName = "ishttpurl" then
        s = isHttpUrl(strDQ(s)) 
    elseIf labelName = "fullhttpurl" then
        s = fullHttpUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "batchfullhttpurl" then
        s = batchFullHttpUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "urljianjiehandle" then
        s = uRLJianJieHandle(strDQ(s)) 
    elseIf labelName = "urltoasc" then
        s = urlToAsc(strDQ(s)) 
    elseIf labelName = "getwebtitle" then
        s = getWebTitle(strDQ(s)) 
    elseIf labelName = "getcontentahref" then
        s = getContentAHref(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "getcontentimgsrc" then
        s = getContentImgSrc(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handleconenturl" then
        s = handleConentUrl(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "replacecontentjsdir" then
        s = replaceContentJsDir(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handlelink" then
        s = handleLink(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5)), strDQ(splxx(6))) 
    elseIf labelName = "getendurlhandlename" then
        s = getEndUrlHandleName(strDQ(s)) 
    elseIf labelName = "geturllistinwebsitelist" then
        s = getUrlListInWebSiteList(strDQ(s)) 
    elseIf labelName = "getthisurlfilename" then
        s = getThisUrlFileName() 
    elseIf labelName = "handlewebhtmlimg" then
        s = handleWebHtmlImg(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handlewebcssimg" then
        s = handleWebCssImg(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "batchhandleurlintegrity" then
        s = batchHandleUrlIntegrity(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "replacecontentimagepath" then
        s = replaceContentImagePath(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getcsslinklist" then
        s = getCssLinkList(strDQ(s)) 
    elseIf labelName = "gethtmlbackgroundurllist" then
        s = getHtmlBackGroundUrlList(strDQ(s)) 
    elseIf labelName = "getimgurllist" then
        s = getImgUrlList(strDQ(s)) 
    elseIf labelName = "getlinkurl" then
        s = getLinkUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "localurlorremoteurl" then
        s = localUrlOrRemoteUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkremoteurl" then
        s = checkRemoteUrl(strDQ(s)) 
    elseIf labelName = "asailinkadd" then
        s = asaiLinkAdd(strDQ(s)) 
    elseIf labelName = "asailinkdel" then
        s = asaiLinkDel(strDQ(s)) 
    elseIf labelName = "iswww" then
        s = iswww(strDQ(s)) 
    elseIf labelName = "geturladdtoparam" then
        s = getUrlAddToParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "groupurl" then
        s = groupUrl(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlepostcookiesparame" then
        s = handlePostCookiesParame(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "remotehttpurlparameter" then
        s = remoteHttpUrlParameter(strDQ(s)) 
    elseIf labelName = "checkurlname" then
        s = checkUrlName(strDQ(s)) 
    elseIf labelName = "handlehttpurlarray" then
        s = handleHttpUrlArray(strDQ(s)) 
    elseIf labelName = "remotejscssparam" then
        s = remoteJsCssParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handleremotejscssparam" then
        s = handleRemoteJsCssParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "batchhandlehttpurlcomplete" then
        s = batchHandleHttpUrlComplete(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "iswebsite" then
        s = isWebSite(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "issonwebsite" then
        s = isSonWebSite(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handleiswebsite" then
        s = handleIsWebSite(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getcontenturllist" then
        s = getContentUrlList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlegetcontenturllist" then
        s = handleGetContentUrlList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getinchain" then
        s = getInChain(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfiletext" then
        s = getFileText(strDQ(s)) 
    elseIf labelName = "getftext" then
        s = getFText(strDQ(s)) 
    elseIf labelName = "getfiletype" then
        s = getFileType(strDQ(s)) 
    elseIf labelName = "checkfile" then
        s = checkFile(strDQ(s)) 
    elseIf labelName = "existsfile" then
        s = existsFile(strDQ(s)) 
    elseIf labelName = "createfile" then
        s = createFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "savefile" then
        s = saveFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createaddfile" then
        s = createAddFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createaddupfile" then
        s = createAddUpFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "addtofile" then
        s = addToFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "savebytefile" then
        s = saveByteFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createbytefile" then
        s = createByteFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "movefile" then
        s = moveFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "editfile" then
        s = editFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "copyfile" then
        s = copyFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deletefile" then
        s = deleteFile(strDQ(s)) 
    elseIf labelName = "delselfile" then
        s = delSelFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "delfilelist" then
        s = delFileList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "batchdelfile" then
        s = batchDelFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deleteappointfilelist" then
        s = deleteAppointFileList(strDQ(s)) 
    elseIf labelName = "getfilesize" then
        s = getFileSize(strDQ(s)) 
    elseIf labelName = "filesize" then
        s = fileSize(strDQ(s)) 
    elseIf labelName = "getfsize" then
        s = getFSize(strDQ(s)) 
    elseIf labelName = "getfilecreatedate" then
        s = getFileCreateDate(strDQ(s)) 
    elseIf labelName = "filecreatedate" then
        s = fileCreateDate(strDQ(s)) 
    elseIf labelName = "getfileeditdate" then
        s = getFileEditDate(strDQ(s)) 
    elseIf labelName = "fileeditdate" then
        s = fileEditDate(strDQ(s)) 
    elseIf labelName = "getfileedittime" then
        s = getFileEditTime(strDQ(s)) 
    elseIf labelName = "getdatelastaccessed" then
        s = getDateLastAccessed(strDQ(s)) 
    elseIf labelName = "modifyfilemodifydate" then
        s = modifyFileModifydate(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "editfileeditdate" then
        s = editFileEditDate(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "existszhidufile" then
        s = existsZhiDuFile(strDQ(s)) 
    elseIf labelName = "getfileattribute" then
        s = getFileAttribute(strDQ(s)) 
    elseIf labelName = "editfileattribute" then
        s = editFileAttribute(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "existsfileday" then
        s = existsFileDay(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkfileday" then
        s = checkFileDay(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createfolder" then
        s = createFolder(strDQ(s)) 
    elseIf labelName = "createdirfolder" then
        s = createDirFolder(strDQ(s)) 
    elseIf labelName = "createdomain" then
        s = createDomain(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createdomainfolder" then
        s = createDomainFolder(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "movefolder" then
        s = moveFolder(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "editfolder" then
        s = editFolder(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "copyfolder" then
        s = copyFolder(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deletefolder" then
        s = deleteFolder(strDQ(s)) 
    elseIf labelName = "getfoldersize" then
        s = getFolderSize(strDQ(s)) 
    elseIf labelName = "checkfolder" then
        s = checkFolder(strDQ(s)) 
    elseIf labelName = "existsfolder" then
        s = existsFolder(strDQ(s)) 
    elseIf labelName = "getfoldercreatedate" then
        s = getFolderCreateDate(strDQ(s)) 
    elseIf labelName = "getfoldereditdate" then
        s = getFolderEditDate(strDQ(s)) 
    elseIf labelName = "getfoldertime" then
        s = getFolderTime(strDQ(s)) 
    elseIf labelName = "getfolderattribute" then
        s = getFolderAttribute(strDQ(s)) 
    elseIf labelName = "editfolderattribute" then
        s = editFolderAttribute(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "existsfolderday" then
        s = existsFolderDay(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkfolderday" then
        s = checkFolderDay(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfilefolderlist" then
        s = getFileFolderList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5)), strDQ(splxx(6))) 
    elseIf labelName = "getdirjpglist" then
        s = getDirJpgList(strDQ(s)) 
    elseIf labelName = "getdirpnglist" then
        s = getDirPngList(strDQ(s)) 
    elseIf labelName = "getdirinilist" then
        s = getDirIniList(strDQ(s)) 
    elseIf labelName = "getdirtxtlist" then
        s = getDirTxtList(strDQ(s)) 
    elseIf labelName = "getdirjslist" then
        s = getDirJsList(strDQ(s)) 
    elseIf labelName = "getdircsslist" then
        s = getDirCssList(strDQ(s)) 
    elseIf labelName = "getdirhtmllist" then
        s = getDirHtmlList(strDQ(s)) 
    elseIf labelName = "getdirasplist" then
        s = getDirAspList(strDQ(s)) 
    elseIf labelName = "getdirphplist" then
        s = getDirPhpList(strDQ(s)) 
    elseIf labelName = "getdirfilelist" then
        s = getDirFileList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdirjpgnamelist" then
        s = getDirJpgNameList(strDQ(s)) 
    elseIf labelName = "getdirpngnamelist" then
        s = getDirPngNameList(strDQ(s)) 
    elseIf labelName = "getdirininamelist" then
        s = getDirIniNameList(strDQ(s)) 
    elseIf labelName = "getdirtxtnamelist" then
        s = getDirTxtNameList(strDQ(s)) 
    elseIf labelName = "getdirjsnamelist" then
        s = getDirJsNameList(strDQ(s)) 
    elseIf labelName = "getdircssnamelist" then
        s = getDirCssNameList(strDQ(s)) 
    elseIf labelName = "getdirhtmlnamelist" then
        s = getDirHtmlNameList(strDQ(s)) 
    elseIf labelName = "getdiraspnamelist" then
        s = getDirAspNameList(strDQ(s)) 
    elseIf labelName = "getdirphpnamelist" then
        s = getDirPhpNameList(strDQ(s)) 
    elseIf labelName = "getdirfilenamelist" then
        s = getDirFileNameList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdirallfilelist" then
        s = getDirAllFileList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdirallfilenamelist" then
        s = getDirAllFileNameList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdirfolderlist" then
        s = getDirFolderList(strDQ(s)) 
    elseIf labelName = "getdirfoldernamelist" then
        s = getDirFolderNameList(strDQ(s)) 
    elseIf labelName = "getdirallfolderlist" then
        s = getDirAllFolderList(strDQ(s)) 
    elseIf labelName = "getdirallfoldernamelist" then
        s = getDirAllFolderNameList(strDQ(s)) 
    elseIf labelName = "getfilefoldernumb" then
        s = getFileFolderNumb(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getdirfoldernumb" then
        s = getDirFolderNumb(strDQ(s)) 
    elseIf labelName = "getdirfoldersize" then
        s = getDirFolderSize(strDQ(s)) 
    elseIf labelName = "getdirfilenumb" then
        s = getDirFileNumb(strDQ(s)) 
    elseIf labelName = "getdirallfilenumb" then
        s = getDirAllFileNumb(strDQ(s)) 
    elseIf labelName = "getdirallfoldernumb" then
        s = getDirAllFolderNumb(strDQ(s)) 
    elseIf labelName = "getfolderlist" then
        s = getFolderList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdirfilesort" then
        s = getDirFileSort(strDQ(s)) 
    elseIf labelName = "getdirfoldersort" then
        s = getDirFolderSort(strDQ(s)) 
    elseIf labelName = "deldirallfile" then
        call delDirAllFile(strDQ(s)) 
    elseIf labelName = "deleteallfile" then
        call deleteAllFile(strDQ(s)) 
    elseIf labelName = "getupdatefile" then
        s = getUpdateFile(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "handlepath" then
        s = handlePath(strDQ(s)) 
    elseIf labelName = "realpath" then
        s = realPath(strDQ(s)) 
    elseIf labelName = "fullpath" then
        s = fullPath(strDQ(s)) 
    elseIf labelName = "handlerelativepath" then
        s = handleRelativePath(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "dirname" then
        s = dirName(strDQ(s)) 
    elseIf labelName = "isobjinstalled" then
        s = isObjInstalled(strDQ(s)) 
    elseIf labelName = "existsobject" then
        s = existsObject(strDQ(s)) 
    elseIf labelName = "getfileobject" then
        s = getFileObject() 
    elseIf labelName = "getdictionary" then
        s = getDictionary() 
    elseIf labelName = "getstream" then
        s = getStream() 
    elseIf labelName = "showattributetip" then
        s = showAttributeTip(strDQ(s)) 
    elseIf labelName = "printspacevalue" then
        s = printSpaceValue(strDQ(s)) 
    elseIf labelName = "printspacesize" then
        s = printSpaceSize(strDQ(s)) 
    elseIf labelName = "printfilesize" then
        s = printFileSize(strDQ(s)) 
    elseIf labelName = "getdeskpath" then
        s = getDeskPath() 
    elseIf labelName = "createurlshortcut" then
        s = createUrlShortcut(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createfoldershortcut" then
        s = createFolderShortcut(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "finddirfilename" then
        s = findDirFileName(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "findsubdirfilename" then
        s = findSubDirFileName(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getarticlefillcontent" then
        s = getArticleFillContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getfilepathname" then
        s = getFilePathName(strDQ(s)) 
    elseIf labelName = "getextension" then
        s = getExtension(strDQ(s)) 
    elseIf labelName = "removefiledir" then
        s = removeFileDir(strDQ(s)) 
    elseIf labelName = "basename" then
        s = baseName(strDQ(s)) 
    elseIf labelName = "getfileextname" then
        s = getFileExtName(strDQ(s)) 
    elseIf labelName = "handlefilepatharray" then
        s = handleFilePathArray(strDQ(s)) 
    elseIf labelName = "getstrfilename" then
        s = getStrFileName(strDQ(s)) 
    elseIf labelName = "getfilename" then
        s = getFileName(strDQ(s)) 
    elseIf labelName = "getfileattr" then
        s = getFileAttr(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfileext" then
        s = getFileExt(strDQ(s)) 
    elseIf labelName = "getparentpath" then
        s = getParentPath(strDQ(s)) 
    elseIf labelName = "checkurlfilenameparam" then
        s = checkUrlFileNameParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "isucase" then
        s = isUCase(strDQ(s)) 
    elseIf labelName = "islcase" then
        s = isLCase(strDQ(s)) 
    elseIf labelName = "checkerr" then
        call checkErr(strDQ(s)) 
    elseIf labelName = "strtotruefalse" then
        s = strToTrueFalse(strDQ(s)) 
    elseIf labelName = "strtruefalsetoint" then
        s = strTrueFalseToInt(strDQ(s)) 
    elseIf labelName = "checkvbcrlf" then
        s = checkVbCrlf(strDQ(s)) 
    elseIf labelName = "checkbr" then
        s = checkBr(strDQ(s)) 
    elseIf labelName = "isparity" then
        s = isParity(strDQ(s)) 
    elseIf labelName = "checkeval" then
        s = checkEval(strDQ(s)) 
    elseIf labelName = "replace_sqltext" then
        s = replace_SQLText(strDQ(s)) 
    elseIf labelName = "chkpost" then
        s = chkPost() 
    elseIf labelName = "ismail" then
        s = isMail(strDQ(s)) 
    elseIf labelName = "isvalidemail" then
        s = isValidEmail(strDQ(s)) 
    elseIf labelName = "iscode" then
        s = isCode(strDQ(s)) 
    elseIf labelName = "checknumber" then
        s = checkNumber(strDQ(s)) 
    elseIf labelName = "isnumber" then
        s = isNumber(strDQ(s)) 
    elseIf labelName = "checkabc" then
        s = checkABC(strDQ(s)) 
    elseIf labelName = "getlen" then
        s = getLen(strDQ(s)) 
    elseIf labelName = "checktime" then
        s = checkTime(strDQ(s)) 
    elseIf labelName = "isnul" then
        s = isNul(strDQ(s)) 
    elseIf labelName = "foundinarr" then
        s = foundInArr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "displayhtml" then
        s = displayHtml(strDQ(s)) 
    elseIf labelName = "undisplayhtml" then
        s = unDisplayHtml(strDQ(s)) 
    elseIf labelName = "handleclosehtml" then
        s = handleCloseHtml(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handlehtmlahref" then
        s = handleHtmlAHref(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "addtooraddhttpurl" then
        s = addToOrAddHttpUrl(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "gethtmllablename" then
        s = getHtmlLableName(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "removeblanklines" then
        s = removeBlankLines(strDQ(s)) 
    elseIf labelName = "checkcode" then
        s = checkCode(strDQ(s)) 
    elseIf labelName = "getfilecharset" then
        s = getFileCharset(strDQ(s)) 
    elseIf labelName = "seniorcheckcode" then
        s = seniorCheckCode(strDQ(s)) 
    elseIf labelName = "readfile" then
        s = readFile(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getstext" then
        s = getStext(strDQ(s)) 
    elseIf labelName = "getftextutf" then
        s = getFTextUTF(strDQ(s)) 
    elseIf labelName = "writetofile" then
        s = writeToFile(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "createfilegbk" then
        s = createFileGBK(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "createfileutf" then
        s = createFileUTF(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getugtext" then
        s = getUGText(strDQ(s)) 
    elseIf labelName = "checkcontentcharset" then
        s = checkContentCharset(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkfilecode" then
        s = checkFileCode(strDQ(s)) 
    elseIf labelName = "getfilebinarycontent" then
        s = getFileBinaryContent(strDQ(s)) 
    elseIf labelName = "binaryreadfile" then
        s = binaryReadFile(strDQ(s)) 
    elseIf labelName = "readbinary" then
        s = readBinary(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "decsavebinary" then
        s = decSaveBinary(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "writebinary" then
        call writeBinary(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfilecharset2" then
        s = getFileCharset2(strDQ(s)) 
    elseIf labelName = "bintohex" then
        s = binToHex(strDQ(s)) 
    elseIf labelName = "base64encode" then
        s = base64Encode(strDQ(s)) 
    elseIf labelName = "base64decode" then
        s = base64Decode(strDQ(s)) 
    elseIf labelName = "asphelptip" then
        s = ASPHelpTip(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlewrite" then
        s = handleWrite(strDQ(s)) 
    elseIf labelName = "jiantitabfanti" then
        s = jianTiTabFanTi(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "jiyicireplace" then
        s = jiYiCiReplace(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "cireplace" then
        s = ciReplace(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "testgetpy" then
        call testGetPY() 
    elseIf labelName = "pinyin" then
        s = pinYin(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "pinyin2" then
        s = pinYin2(strDQ(s)) 
    elseIf labelName = "entocompany" then
        s = enToCompany(strDQ(s)) 
    elseIf labelName = "remotecontentjinghao" then
        s = remoteContentJingHao(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "remotearrayjinghao" then
        s = remoteArrayJingHao(strDQ(s)) 
    elseIf labelName = "geteachstraddvalue" then
        s = getEachStrAddValue(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getvalueinarrayid" then
        s = getValueInArrayID(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkvalueinarray" then
        s = checkValueInArray(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deleterepeatarray" then
        s = deleteRepeatArray(strDQ(s)) 
    elseIf labelName = "deleterepeatcontent" then
        s = deleteRepeatContent(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getarraycount" then
        s = getArrayCount(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "randomshow" then

        s = randomShow(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "arrayrandomshow" then
        s = arrayRandomShow(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "printarray" then
        call printArray(strDQ(s)) 
    elseIf labelName = "echoarray" then
        call echoArray(strDQ(s)) 
    elseIf labelName = "str_shuffle" then
        s = str_Shuffle(strDQ(s)) 
    elseIf labelName = "characterupset" then
        s = characterUpset(strDQ(s)) 
    elseIf labelName = "str_split" then
        s = str_Split(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "removenullvaluearray" then
        s = removeNullValueArray(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "removerepeatvaluearray" then
        s = removeRepeatValueArray(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlearray" then
        s = handleArray(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "arraytostring" then
        s = arrayToString(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "testarraydata" then
        s = testArrayData() 
    elseIf labelName = "selectsort" then
        s = selectSort(strDQ(s)) 
    elseIf labelName = "quicksort" then
        s = quickSort(strDQ(s)) 
    elseIf labelName = "bubblesort" then
        s = bubbleSort(strDQ(s)) 
    elseIf labelName = "insertsort" then
        s = insertSort(strDQ(s)) 
    elseIf labelName = "requicksort" then
        s = reQuickSort(strDQ(s)) 
    elseIf labelName = "arrayreverse" then
        s = arrayReverse() 
    elseIf labelName = "responsearray" then
        call responseArray(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "max" then
        s = max(strDQ(s)) 
    elseIf labelName = "min" then
        s = min(strDQ(s)) 
    elseIf labelName = "testtwodimensionalarray" then
        s = testTwoDimensionalArray() 
    elseIf labelName = "handlesplitarray" then
        s = handleSplitArray(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "twodimensionalarrayasc" then
        s = twoDimensionalArrayAsc(strDQ(s)) 
    elseIf labelName = "twodimensionalarraydesc" then
        s = twoDimensionalArrayDesc(strDQ(s)) 
    elseIf labelName = "csscompression" then
        s = cssCompression(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "deletecssnote" then
        s = deleteCssNote(strDQ(s)) 
    elseIf labelName = "uncsscompression" then
        s = unCssCompression(strDQ(s)) 
    elseIf labelName = "trimvbcrlf" then
        s = trimVbCrlf(strDQ(s)) 
    elseIf labelName = "phptrim" then
        s = phpTrim(strDQ(s)) 
    elseIf labelName = "phpltrim" then
        s = phpLTrim(strDQ(s)) 
    elseIf labelName = "phprtrim" then
        s = phpRTrim(strDQ(s)) 
    elseIf labelName = "trimvbtab" then
        s = trimVbTab(strDQ(s)) 
    elseIf labelName = "ltrimvbtab" then
        s = lTrimVbTab(strDQ(s)) 
    elseIf labelName = "rtrimvbtab" then
        s = rTrimVBTab(strDQ(s)) 
    elseIf labelName = "gethtmlcssstyle" then
        s = getHtmlCssStyle(strDQ(s)) 
    elseIf labelName = "handlecleancss" then
        s = handleCleanCss(strDQ(s)) 
    elseIf labelName = "removeexcessrow" then
        s = removeExcessRow(strDQ(s)) 
    elseIf labelName = "cssaddtostyle" then
        s = cssAddToStyle(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "checkcssstyle" then
        s = checkCssStyle(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "cssstyleaddtoparam" then
        s = cssStyleAddToParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "findcssstyle" then
        s = findCssStyle(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlecutcsscode" then
        s = handleCutCssCode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlecutdivcode" then
        s = handleCutDivCode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "gethtmlbackgroundimglist" then
        s = getHtmlBackGroundImgList(strDQ(s)) 
    elseIf labelName = "gethandlewebhtmllink" then
        s = getHandleWebHtmlLink(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getcsslisturllist" then
        s = getCssListUrlList(strDQ(s)) 
    elseIf labelName = "handlereadcsscontent" then
        s = handleReadCssContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handlecsspx" then
        s = handleCssPX(strDQ(s)) 
    elseIf labelName = "handlehtmlstyle" then
        s = handleHtmlStyle(strDQ(s)) 
    elseIf labelName = "utf2gb" then
        s = UTF2GB(strDQ(s)) 
    elseIf labelName = "checkutftogb2312" then
        s = checkUTFToGB2312(strDQ(s)) 
    elseIf labelName = "convchinese" then
        s = convChinese(strDQ(s)) 
    elseIf labelName = "utf8_urlencoding" then
        s = UTF8_URLEncoding(strDQ(s)) 
    elseIf labelName = "toutf8" then
        s = toUTF8(strDQ(s)) 
    elseIf labelName = "chinese2unicode" then
        s = chinese2unicode(strDQ(s)) 
    elseIf labelName = "isvalidhex" then
        s = isvalidhex(strDQ(s)) 
    elseIf labelName = "escape" then
        s = escape(strDQ(s)) 
    elseIf labelName = "unescape" then
        s = unescape(strDQ(s)) 
    elseIf labelName = "chinesetounicode" then
        s = chineseToUnicode(strDQ(s)) 
    elseIf labelName = "unicodetochinese" then
        s = unicodeToChinese(strDQ(s)) 
    elseIf labelName = "urlencode" then
        s = URLEncode(strDQ(s)) 
    elseIf labelName = "urldecode" then
        s = URLDecode(strDQ(s)) 
    elseIf labelName = "toUTFChar" then
        s = toUTFChar(strDQ(s)) 
    elseIf labelName = "encdecurl" then
        s = encDecURL(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "encurl" then
        s = encURL(strDQ(s)) 
    elseIf labelName = "decurl" then
        s = decURL(strDQ(s)) 
    elseIf labelName = "infocolor" then
        s = infoColor(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "imgerror" then
        s = imgError() 
    elseIf labelName = "targetstr" then
        s = handleTargetStr(strDQ(s)) 
    elseIf labelName = "atarget" then
        s = aTarget(strDQ(s)) 
    elseIf labelName = "atitle" then
        s = aTitle(strDQ(s)) 
    elseIf labelName = "imgalt" then
        s = imgAlt(strDQ(s)) 
    elseIf labelName = "imgtitlealt" then
        s = imgTitleAlt(strDQ(s)) 
    elseIf labelName = "arel" then
        s = aRel(strDQ(s)) 
    elseIf labelName = "styleclass" then
        s = styleClass(strDQ(s)) 
    elseIf labelName = "textfontb" then
        s = textFontB(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "textfontcolor" then
        s = textFontColor(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "fontcolorfontb" then
        s = fontColorFontB(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getdefaultfilename" then
        s = getDefaultFileName() 
    elseIf labelName = "ahref" then
        s = aHref(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "imgsrc" then
        s = imgSrc(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "selecttarget" then
        s = selectTarget(strDQ(s)) 
    elseIf labelName = "selectfontcolor" then
        s = selectFontColor(strDQ(s)) 
    elseIf labelName = "selectsex" then
        s = selectSex(strDQ(s)) 
    elseIf labelName = "selectsessioncookies" then
        s = selectSessionCookies(strDQ(s)) 
    elseIf labelName = "showselectlist" then
        s = showSelectList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "showarticleliststyle" then
        s = showArticleListStyle(strDQ(s)) 
    elseIf labelName = "showarticleinfostyle" then
        s = showArticleInfoStyle(strDQ(s)) 
    elseIf labelName = "handlearticleliststyleorinfostyle" then
        s = handleArticleListStyleOrInfoStyle(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "showwebmoduleskins" then
        s = showWebModuleSkins(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "showradiolist" then
        s = showRadioList(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputcheckbox" then
        s = inputCheckBox(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "inputcheckbox3" then
        s = inputCheckBox3(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handleinputcheckbox" then
        s = handleInputCheckBox(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "inputtext" then
        s = inputText(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputtext2" then
        s = inputText2(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "inputlefttext" then
        s = inputLeftText(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputlefttexthelptextright" then
        s = inputLeftTextHelpTextRight(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputlefttextcontent" then
        s = inputLeftTextContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputrighttextcontent" then
        s = inputRightTextContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handleinputleftrighttextcontent" then
        s = handleInputLeftRightTextContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "inputleftpasstext" then
        s = inputLeftPassText(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputleftpasstextcontent" then
        s = inputLeftPassTextContent(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "inputhiddentext" then
        s = inputHiddenText(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "inputtextarea" then
        s = inputTextArea(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "inputhiddentextarea" then
        s = inputHiddenTextArea(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "handleinputhiddentextarea" then
        s = handleInputHiddenTextArea(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5))) 
    elseIf labelName = "showselectdirlist" then
        s = showSelectDirList(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "inputdisabled" then
        s = inputDisabled(strDQ(s)) 
    elseIf labelName = "inputaddalt" then
        s = inputAddAlt(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "webtitle_inputtextarea" then
        s = webTitle_InputTextArea(strDQ(s)) 
    elseIf labelName = "webkeywords_inputtext" then
        s = webKeywords_InputText(strDQ(s)) 
    elseIf labelName = "webdescription_inputtextarea" then
        s = webDescription_InputTextArea(strDQ(s)) 
    elseIf labelName = "foldername_inputtext" then
        s = folderName_InputText(strDQ(s)) 
    elseIf labelName = "filename_inputtext" then
        s = fileName_InputText(strDQ(s)) 
    elseIf labelName = "templatepath_inputtext" then
        s = templatePath_InputText(strDQ(s)) 
    elseIf labelName = "clickpinyinhtmlstr" then
        s = clickPinYinHTMLStr(strDQ(s)) 
    elseIf labelName = "showfontcolorfontb" then
        s = showFontColorFontB(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "showsort" then
        s = showSort(strDQ(s)) 
    elseIf labelName = "showwebnavtype" then
        s = showWebNavType(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4)), strDQ(splxx(5))) 
    elseIf labelName = "showonhtml" then
        s = showOnHtml(strDQ(s)) 
    elseIf labelName = "showthrough" then
        s = showThrough(strDQ(s)) 
    elseIf labelName = "showrecommend" then
        s = showRecommend(strDQ(s)) 
    elseIf labelName = "showonoffimg" then
        s = showOnOffImg(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "newshowonoffimg" then
        s = newShowOnOffImg(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "controldialogcss" then
        s = controlDialogCss() 
    elseIf labelName = "batchdeletetempstr" then
        s = batchDeleteTempStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "htmlformatting" then
        s = htmlFormatting(strDQ(s)) 
    elseIf labelName = "handlehtmlformatting" then
        s = handleHtmlFormatting(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "formatting" then
        s = formatting(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "ziphtml" then
        s = zipHTML(strDQ(s)) 
    elseIf labelName = "checkhtmlformatting" then
        s = checkHtmlFormatting(strDQ(s)) 
    elseIf labelName = "newreplacevalueparam" then
        s = newReplaceValueParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "newrparam" then
        s = newRParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getcontentallrunstr" then
        s = getContentAllRunStr(strDQ(s)) 
    elseIf labelName = "getcontentrunstr" then
        s = getContentRunStr(strDQ(s)) 
    elseIf labelName = "checkcontentrunstr" then
        s = checkContentRunStr(strDQ(s)) 
    elseIf labelName = "handledoublequotation" then
        s = handleDoubleQuotation(strDQ(s)) 
    elseIf labelName = "strdq" then
        s = strDQ(strDQ(s)) 
    elseIf labelName = "handletoarray" then
        s = handleToArray(strDQ(s)) 
    elseIf labelName = "handlecontentcode" then
        s = handleContentCode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handleinmodule" then
        s = handleInModule(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "clearrparam" then
        s = clearRParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "atrparam" then
        s = atRParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "rparam" then
        s = rParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "handlerparam" then
        s = handleRParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getconfigblock" then
        s = getConfigBlock(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getconfigblock2" then
        s = getConfigBlock2(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getconfigblock3" then
        s = getConfigBlock3(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getcutconfigblock" then
        s = getCutConfigBlock(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "getconfigcontentblock" then
        s = getConfigContentBlock(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getconfigfileblock" then
        s = getConfigFileBlock(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "setconfigfileblock" then
        s = setConfigFileBlock(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "delconfigblock" then
        s = delConfigBlock(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "delconfigblock2" then
        s = delConfigBlock2(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "delconfigblock3" then
        s = delConfigBlock3(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "delcutconfigblock" then
        s = delCutConfigBlock(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "getfileparamvalue" then
        s = getFileParamValue(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "setfileparamvalue" then
        s = setFileParamValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "handlegetsetfileparamevalue" then
        s = handleGetSetFileParameValue(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "setrparam" then
        s = setRParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "addreplacerparam" then
        s = addReplaceRParam(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "deletestrcut" then
        s = deleteStrCut(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "replacecontentmodule" then
        s = replaceContentModule(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3)), strDQ(splxx(4))) 
    elseIf labelName = "replacecontentrowmodule" then
        s = replaceContentRowModule(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "handleconfigfile" then
        call handleConfigFile(strDQ(s)) 
    elseIf labelName = "getrparam" then
        s = getRParam(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getdefaultvalue" then
        s = getDefaultValue(strDQ(s)) 
    elseIf labelName = "addmodulereplacearray" then
        call addModuleReplaceArray(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "modulefindcontent" then
        s = moduleFindContent(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "batchfindfunction" then
        s = batchFindFunction(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getfindfilefunction" then
        s = getFindFileFunction(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "getaspcode" then
        s = getASPCode(strDQ(s)) 
    elseIf labelName = "handlegetaspcode" then
        s = handleGetAspCode(strDQ(splxx(0)), strDQ(splxx(1))) 
    elseIf labelName = "delaspnote" then
        s = delAspNote(strDQ(s)) 
    elseIf labelName = "findaspfunction" then
        s = findAspFunction(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2)), strDQ(splxx(3))) 
    elseIf labelName = "firstletteruppercase" then
        s = firstLetterUppercase(strDQ(s)) 
    elseIf labelName = "contentlushow" then
        s = contentLUShow(strDQ(s)) 
    elseIf labelName = "asptovb" then
        s = ASPToVB(strDQ(s)) 
    elseIf labelName = "asptovbnet" then
        s = ASPToVBNET(strDQ(s)) 
    elseIf labelName = "dimadddefaultvalue" then
        s = dimAddDefaultValue(strDQ(s)) 
    elseIf labelName = "functionparamtonet" then
        s = functionParamtoNet(strDQ(s)) 
    elseIf labelName = "getbeforenstr" then
        s = getBeforeNStr(strDQ(splxx(0)), strDQ(splxx(1)), strDQ(splxx(2))) 
    elseIf labelName = "getstrdim" then
        s = getStrDim(strDQ(s)) 
    elseIf labelName = "getfunctionstrdim" then
        s = getFunctionStrDim(strDQ(s)) 
    elseIf labelName = "getrandcolor" then
        s = getRandColor() 
	'函数判断 end
        '----------------------- ASP自身内部函数处理 -----------------------------
    '系统函数 Now
    elseIf labelName = "now" then
        s = now() 
    '系统函数 Date
    elseIf labelName = "date" then
        s = date()
    '系统函数 Len
    elseIf labelName = "len" then
        s = len(s) 
    '系统函数 LCase
    elseIf labelName = "lcase" then
        s = LCase(s) 
    '系统函数 UCase
    elseIf labelName = "ucase" then
        s = uCase(s) 
    '系统函数 ltrim
    elseIf labelName = "ltrim" then
        s = LTrim(s) 
    '系统函数 rtrim
    elseIf labelName = "rtrim" then
        s = rTrim(s) 
    '系统函数 trim
    elseIf labelName = "trim" then
        s = trim(s) 
    '系统函数 Left
    elseIf labelName = "left" then
        splxx = split(s & ",1", ",") 
        s = left(trim(splxx(0)), trim(splxx(1))) 
    '系统函数 Right
    elseIf labelName = "right" then
        s = right(trim(splxx(0)), trim(splxx(1))) 
    '系统函数 Mid
    elseIf labelName = "mid" then
        splxx = split(s & ",99999", ",") 
        if uBound(splxx) >= 2 then
            s = mid(trim(splxx(0)), trim(splxx(1)), trim(splxx(2))) 
        else
            s = mid(trim(splxx(0)), trim(splxx(1))) 
        end if 
    end if 
    content = replace(content, yunStr, s) 
    handleContentCode = content 
end function 


'内部模块处理 HandleInModule(Content,"start") HandleInModule(Content,"end")
function handleInModule(content, sType)
    sType = LCase(CStr(sType)) 
    if sType = "1" or sType = "start" then
        content = replace(content, "\'", "\|*|\") 
        content = replace(content, "\=", "\|&|\")                                   '后加20141024
    elseIf sType = "2" or sType = "end" then
        content = replace(content, "\|*|\", "'") 
        content = replace(content, "\$", "$") 
        content = replace(content, "\}", "}") 

        content = replace(content, "\|&|\", "=")                                    '后加20141024
    end if 
    handleInModule = content 
end function 
'清除特殊样式后获得标签值
function clearRParam(byVal action, byVal lableStr)
    dim s 
    'Action=Replace(Action,"\'","【|\‘|】")
    action = replace(action, "\'", "")                                              '把这种清掉
    s = rParam(action, lableStr) 
    's=replace(s,"【|\‘|】", "\'")
    clearRParam = s 
end function 
'获得参数内容后 放到动作里处理一下（20151023）
function atRParam(byVal action, byVal lableStr)
    atRParam = rParam(action, lableStr) 
    if inStr(atRParam, "{$") > 0 and inStr(atRParam, "$}") > 0 then
        atRParam = handleTemplateAction(atRParam, false)                                '处理动作  注意这是之前旧版本
        atRParam = handleAction(atRParam)                                               '这是新版本
    end if 
end function 
'读单个参数值  Title = RParam(Action,"Title")     起强版获取参数值20150723
function rParam(byVal action, byVal lableStr)
    dim s 

    '原始 单引号
    s = handleRParam(action, lableStr, "'") 
    '原始 双引号
    if s = "" then
        s = handleRParam(action, lableStr, """") 
    end if 
    '原始 空
    if s = "" then
        s = handleRParam(action, lableStr, "") 
    end if 

    '小写 单引号
    if s = "" then
        s = handleRParam(action, LCase(lableStr), "'") 
    end if 
    '小写 双引号
    if s = "" then
        s = handleRParam(action, LCase(lableStr), """") 
    end if 
    '小写 空
    if s = "" then
        s = handleRParam(action, LCase(lableStr), "") 
    end if 

    '大写 单引号
    if s = "" then
        s = handleRParam(action, uCase(lableStr), "'") 
    end if 
    '大写 双引号
    if s = "" then
        s = handleRParam(action, uCase(lableStr), """") 
    end if 
    '大写 空
    if s = "" then
        s = handleRParam(action, uCase(lableStr), "") 
    end if  
    '不要这个，要不不稳定(20151022)
    'if s=false then s=""
    if s = "[#空*值_#]" then s = ""  
    rParam = s 
end function 
'处理 读单个参数值
function handleRParam(byVal action, byVal lableStr, byVal typeStr)
    dim lalbeName, endTypeStr, isTrue, s 
    isTrue = false                                                                  '是否为真
    endTypeStr = IIF(typeStr <> "", typeStr, " ") 
    action = vbCrLf & " " & action                                                  '给它也加个空格，要不然在没有函数，前面就没有空格
    '默认前面加空格
    lalbeName = " " & lableStr                                                      '加个空格是为了精准
    '不存在  前面加点
    if inStr(action, lalbeName & "=" & typeStr) = 0 and isTrue = false then
        lalbeName = "'" & lableStr 
    else
        isTrue = true 
    end if 
    '不存在 前面加双引号
    if inStr(action, lalbeName & "=" & typeStr) = 0 and isTrue = false then
        lalbeName = """" & lableStr 
    else
        isTrue = true 
    end if 
    '不存在    前面加TAB
    if inStr(action, lalbeName & "=" & typeStr) = 0 and isTrue = false then
        lalbeName = vbTab & lableStr 
    else
        isTrue = true 
    end if 
    '不存在    前面加换行
    if inStr(action, lalbeName & "=" & typeStr) = 0 and isTrue = false then
        lalbeName = vbCrLf & lableStr 
    else
        isTrue = true 
    end if 
    if inStr(action, lalbeName & "=" & typeStr) > 0 and inStr(action, endTypeStr) > 0 then
        s = strCut(action, lalbeName & "=" & typeStr, endTypeStr, 2) 
        s = handleInModule(s, "end")                                                    '处理里面参数 追加于20141031            还原内容值

        if s = "" then
            s = "[#空*值_#]" 
        end if 

        '判断是否对参数进行动作制作
        if inStr(s, "{$") > 0 and inStr(s, "$}") > 0 then

            'handleRParam=HandleTemplateAction(handleRParam,true)        '处理动作
            'handleRParam = handleModuleReplaceArray(handleRParam)'给AddSqL处理一下动作 这是处理替换，不需要，因为在 HandleTemplateAction有替换了(20151021)

        end if 
        '不要这个，要不不稳定(20151022)
        'if handleRParam="" then
        'handleRParam=false
    'end if
    end if 
    handleRParam = s 
end function 


'获得配置块 20150105 GetConfigBlock(ConfigContent, BlockName)
function getConfigBlock(configContent, blockName)
    getConfigBlock = getCutConfigBlock(configContent, blockName, "#", "#") 
end function 
'获得配置块 20150105
function getConfigBlock2(configContent, blockName)
    getConfigBlock2 = getCutConfigBlock(configContent, blockName, "[#", "#]") 
end function 
'获得配置块 20150105
function getConfigBlock3(configContent, blockName)
    getConfigBlock3 = getCutConfigBlock(configContent, blockName, "[$", "$]") 
end function 
'截取配置内容中块 20150105
function getCutConfigBlock(configContent, blockName, startLable, endLable)
    dim startStr, endStr 
    startStr = startLable & blockName & endLable 
    endStr = startLable & blockName & endLable 
    getCutConfigBlock = "" 
    '开始标签处理
    if inStr(configContent, startStr & " start") > 0 then
        startStr = startStr & " start" 
    else
        startStr = startStr & " Start" 
    end if 
    '结束标签处理
    if inStr(configContent, endStr & " end") > 0 then
        endStr = endStr & " end" 
    else
        endStr = endStr & " End" 
    end if 

    if inStr(configContent, startStr) > 0 and inStr(configContent, endStr) > 0 then
        getCutConfigBlock = strCut(configContent, startStr, endStr, 2) 
    end if 
end function 
'获得配置内容块20150401
function getConfigContentBlock(byVal configContent, blockName)
    getConfigContentBlock = getCutConfigBlock(configContent, blockName, "", "") 
end function 
'获得配置文件里块20150401  getConfigFileBlock(ConfigPath, "#txtRunCode#")  测试标签块时则自动创建
function getConfigFileBlock(byVal configFile, blockName) 
    dim content, findStr, replaceStr, startStr, endStr 
    content = getFText(configFile) 
    getConfigFileBlock = "" 
    'MsgBox ("ConfigFile=" & ConfigFile & "(" & checkFile(ConfigFile) & "，" & GetFSize(ConfigFile) & ")" & vbCrLf & "Content=" & Content)
    startStr = blockName & " start" 
    endStr = blockName & " end" 
    replaceStr = startStr & "" & endStr 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        findStr = strCut(content, startStr, endStr, 2) 
        getConfigFileBlock = findStr 
    else
        call CreateFile(configFile, content & replaceStr) 
    end if 
end function 
'设置配置文件里块 20150401 call setConfigFileBlock(ConfigFile, "aaabbc", "#上传目录列表#")  存在则更新
function setConfigFileBlock(byVal configFile, byVal writeContent, blockName)
    dim content, findStr, replaceStr, startStr, endStr 
    content = getFText(configFile) 
    startStr = blockName & " start" 
    endStr = blockName & " end" 
    replaceStr = startStr & writeContent & endStr 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        findStr = strCut(content, startStr, endStr, 1) 
        content = replace(content, findStr, replaceStr) 
    else
        content = content & replaceStr 
    end if 
    call CreateFile(configFile, content) 
    setConfigFileBlock = content 
end function 

'删除配置块 20150322
function delConfigBlock(config, blockName)
    delConfigBlock = delCutConfigBlock(config, blockName, "#", "#") 
end function 
'删除配置块 20150322
function delConfigBlock2(config, blockName)
    delConfigBlock2 = delCutConfigBlock(config, blockName, "[#", "#]") 
end function 

'删除配置块 20150322
function delConfigBlock3(config, blockName)
    delConfigBlock3 = delCutConfigBlock(config, blockName, "[$", "$]") 
end function 
'删除配置内容 20150322
function delCutConfigBlock(config, blockName, startLable, endLable)
    dim startStr, endStr, s 
    startStr = startLable & blockName & endLable & " start" 
    endStr = startLable & blockName & endLable & " end" 
    if inStr(config, startStr) > 0 and inStr(config, endStr) > 0 then
        s = strCut(config, startStr, endStr, 1) 
        config = replace(config, s, "") 
    end if 
    delCutConfigBlock = config 
end function 




'获得文件里设置参数　20150315
function getFileParamValue(configPath, paramName)
    getFileParamValue = handleGetSetFileParameValue(configPath, paramName, "", "获得") 
end function 
'设置文件里设置参数　20150315
function setFileParamValue(configPath, paramName, valueStr)
    setFileParamValue = handleGetSetFileParameValue(configPath, paramName, valueStr, "设置") 
end function 
'处理获得设置文件参数值　20150315
function handleGetSetFileParameValue(configPath, paramName, valueStr, sType)
    dim content, startStr, endStr, yunStr, replaceStr 
    '文件为假时，创建一个空文件看看，如果不能创建这个文件则说明这个文件地址有问题，则退出20150324
    if checkFile(configPath) = false then
        call CreateFile(configPath, "") 
    end if 
    handleGetSetFileParameValue = "" 
    if checkFile(configPath) = false then exit function                             '文件不存在则退出

    content = trimVbCrlf(getFText(configPath)) 
    startStr = vbCrLf & paramName & "=" : endStr = vbCrLf 
    replaceStr = vbCrLf & paramName & "=" & valueStr & vbCrLf 
    if inStr(vbCrLf & content, startStr) > 0 and inStr(content & vbCrLf, endStr) > 0 then
        yunStr = strCut(vbCrLf & content & vbCrLf, startStr, endStr, 2) 
        if sType = "获得" then
            handleGetSetFileParameValue = yunStr 
            exit function 
        end if 
        yunStr = startStr & yunStr & endStr 
        content = replace(vbCrLf & content & vbCrLf, yunStr, replaceStr) 
        call CreateFile(configPath, content) 
    else
        call CreateFile(configPath, content & vbCrLf & trimVbCrlf(replaceStr)) 
    end if 
end function 

'设置内容里参数 20150611
function setRParam(configPath, paramName, paramValue, sIsNoAdd)
    dim content, startStr, endStr, s 
    content = phpTrim(getFText(configPath)) 
    startStr = paramName & "='" : endStr = "'" 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        s = strCut(content, startStr, endStr, 2) 
        content = replace(content, startStr & s & endStr, startStr & paramValue & endStr) 
        call CreateFile(configPath, content) 

    elseIf trim(sIsNoAdd) = "1" then
        call createAddFile(configPath, startStr & paramValue & endStr) 
    end if 
end function 

'追加或替换参数值 20150615
function addReplaceRParam(byVal content, startStr, endStr, byVal valueStr)
    dim s 
    valueStr = startStr & valueStr & endStr 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        s = strCut(content, startStr, endStr, 1) 
        content = replace(content, s, valueStr) 
    else
        content = content & vbCrLf & valueStr 
    end if 
    addReplaceRParam = content 
end function 
'删除指定字符N次
function deleteStrCut(byVal content, byVal startStr, byVal endStr, byVal cutType, nDelCount)
    dim i, s 
    if nDelCount = 0 then
        nDelCount = 99 
    end if 
    for i = 0 to nDelCount
        s = getStrCut(content, startStr, endStr, 1) 
        if s <> "" then
            content = replace(content, s, "") 
        else
            exit for 
        end if 
    next 
    deleteStrCut = content 
end function 



'call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容",""))
'call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加在前"))
'call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加"))
'call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","追加在前"))
'call Rwend(replaceContentModule(getftext("1.html"),"<div>","</div>","替换内容","外部追加在前"))
'call Rwend(replaceContentModule(getftext("1.html"), "<div>", "</div>", "替换内容", "外部追加"))
'替换内容里模块   ReplaceType(空为替换，追加在前，追加在后(追加)，外部追加在前，外部追加在后(外部追加))
function replaceContentModule(byVal content, startStr, endStr, replaceValue, replaceType)
    dim splStr, splxx, s, i, sSplType, valueList, sNewStartStr, sNewEndStr, sourceValueList, sourceValue, tempS, sNewReplaceValue 
    if inStr(content, startStr) = 0 and inStr(content, endStr) = 0 then
        replaceContentModule = content 
        exit function 
    end if 
    sSplType = "$Array$" 
    for i = 1 to 99
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            s = strCut(content, startStr, endStr, 1) : tempS = s 
            s = mid(s, len(startStr) + 1, len(s) - len(startStr) - len(endStr)) 
            sNewStartStr = getEachStrAddValue(startStr, "|*|") 
            if inStr(sSplType & valueList & sSplType, sSplType & sNewStartStr & sSplType) = 0 then
                if valueList <> "" then valueList = valueList & sSplType 
                valueList = valueList & sNewStartStr 
                if sourceValueList <> "" then sourceValueList = sourceValueList & sSplType 
                sourceValueList = sourceValueList & startStr 
            end if 
            sNewEndStr = getEachStrAddValue(endStr, "|*|") 
            if inStr(sSplType & valueList & sSplType, sSplType & sNewEndStr & sSplType) = 0 then
                if valueList <> "" then valueList = valueList & sSplType 
                valueList = valueList & sNewEndStr 
                if sourceValueList <> "" then sourceValueList = sourceValueList & sSplType 
                sourceValueList = sourceValueList & endStr 
            end if 

            if replaceType = "追加在前" then
                sNewReplaceValue = sNewStartStr & replaceValue & s & sNewEndStr 
            elseIf replaceType = "追加在后" or replaceType = "追加" then
                sNewReplaceValue = sNewStartStr & s & replaceValue & sNewEndStr 
            elseIf replaceType = "外部追加在前" then
                sNewReplaceValue = replaceValue & sNewStartStr & s & sNewEndStr 

            elseIf replaceType = "外部追加在后" or replaceType = "外部追加" then
                sNewReplaceValue = sNewStartStr & s & sNewEndStr & replaceValue 
            else
                sNewReplaceValue = replaceValue 
            end if 

            content = replace(content, tempS, sNewReplaceValue) 
        else
            exit for 
        end if 
    next 
    'call rwend(content)
    splStr = split(valueList, sSplType) 
    splxx = split(sourceValueList, sSplType) 
    for i = 0 to uBound(splStr)
        sourceValue = splStr(i) 
        replaceValue = splxx(i) 
        content = replace(content, sourceValue, replaceValue) 
    next 
    replaceContentModule = content 
end function 

'call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", "追加在前"))
'call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", "追加"))
'call rwend(replaceContentRowModule(getftext("1.html"),"<div>11</div>", "替换内容", ""))
'替换内容里一行模块   ReplaceType(空为替换，追加在前，追加在后(追加))
function replaceContentRowModule(content, searchValue, replaceValue, replaceType)
    dim splStr, splxx, i, sSplType, valueList, sourceValueList, sourceValue, sNewReplaceValue, sNewSearchValue 
    sSplType = "$Array$" 
    for i = 1 to 99
        if inStr(content, searchValue) > 0 then
            sNewSearchValue = getEachStrAddValue(searchValue, "|*|") 
            if inStr(sSplType & valueList & sSplType, sSplType & sNewSearchValue & sSplType) = 0 then
                if valueList <> "" then valueList = valueList & sSplType 
                valueList = valueList & sNewSearchValue 
                if sourceValueList <> "" then sourceValueList = sourceValueList & sSplType 
                sourceValueList = sourceValueList & searchValue 
            end if 
            if replaceType = "追加在前" then
                sNewReplaceValue = replaceValue & sNewSearchValue 
            elseIf replaceType = "追加在后" or replaceType = "追加" then
                sNewReplaceValue = sNewSearchValue & replaceValue 
            else
                sNewReplaceValue = replaceValue 
            end if 
            content = replace(content, searchValue, sNewReplaceValue) 
        else
            exit for 
        end if 
    next 

    'call rwend(content)
    splStr = split(valueList, sSplType) 
    splxx = split(sourceValueList, sSplType) 
    for i = 0 to uBound(splStr)
        sourceValue = splStr(i) 
        replaceValue = splxx(i) 
        content = replace(content, sourceValue, replaceValue) 
    next 
    replaceContentRowModule = content 

end function 
'处理配置文档(20150804)
sub handleConfigFile(configPath)
    dim c 
    if checkFile(configPath) = false then
        c = "#Help帮助# start" & vbCrLf & "默认帮助内容" & vbCrLf & "#Help帮助# end" 
        call CreateFile(configPath, c) 
    end if 
end sub 

'获得内容里指定类型值   RParam加强版(20161025)
function getRParam(byVal content, byVal lableStr)
    dim contentLCase, endS, i, s, c, isStart, startStr, isValue 
    content = " " & content & " "                                                   '避免更精准获得值
    contentLCase = LCase(content) 
    lableStr = LCase(lableStr) 
    endS = mid(content, inStr(contentLCase, lableStr) + len(lableStr)) 
    'call echo("ends",ends)
    isStart = false                                                                 '是否有开始类型值
    isValue = false                                                                 '是否有值
    for i = 1 to len(endS)
        s = mid(endS, i, 1) 
        if isStart = true then
            if s <> "" then
                if startStr = "" then
                    startStr = s 
                else
                    if startStr = """" or startStr = "'" then
                        if s = startStr then
                            isValue = true 
                            exit for 
                        end if 
                    elseIf s = " " and c = "" then

                    elseIf s = " " or s = "/" or s = ">" then
                        isValue = true 
                        exit for 
                    end if 
                    if s <> " " then
                        c = c & s 
                    end if 
                end if 
            end if 
        end if 

        if s = "=" then
            isStart = true 
        end if 
    next 
    if isValue = false then
        c = "" 
    end if 
    getRParam = c 
'call echo("c",c)
end function 




'获得模板某标签默认内容 代码进行了二次查找 会在HTML模板里二次查找默认值
function getDefaultValue(action)
    getDefaultValue = moduleFindContent(action, "default") 
end function 

'添加模块替换数组
sub addModuleReplaceArray(title, content)
    dim i 
    for i = 1 to uBound(moduleReplaceArray) - 1
        if moduleReplaceArray(i, 0) = "" then
            moduleReplaceArray(i, 0) = title 
            moduleReplaceArray(0, i) = content 
            exit sub 
        end if 
    next 
end sub 

'根据标签找到对应内容
function moduleFindContent(action, moduleName)
    dim defaultStr, startStr, endStr 
    defaultStr = rParam(action, moduleName)                                         '把转小写LCase去掉 （20151008）

    startStr = "<!--#" & defaultStr & " start#-->" 
    endStr = "<!--#" & defaultStr & " end#-->" 
    '[_18年独家一次性祛斑第一品牌2014年10月21日 10时59分]
    'Call Echo("Default",Default)
    '判断是否存在
    if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
        defaultStr = getStrCut(code, startStr, endStr, 2) 
    elseIf defaultStr <> "" then
        startStr = "<!--#" & defaultStr 
        endStr = "#-->" 
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            defaultStr = getStrCut(code, startStr, endStr, 2) 
        end if 
    end if 

    '删除默认值20150712
    dim deletedefault 
    deletedefault = rParam(action, "deletedefault") 
    if deletedefault = "true" then
        call addModuleReplaceArray("【删除】", startStr & defaultStr & endStr) 
    end if 
    moduleFindContent = defaultStr 
end function 




'配置内容设置
function getNewConfig(byval content,labelName)
	getNewConfig=handleNewConfig(content,labelName,"","get")
end function
function setNewConfig(content,labelName,labelValue)
	content=handleNewConfig(content,labelName,labelValue,"set")
	setNewConfig=content
end function
function delNewConfig(content,labelName)
	delNewConfig=handleNewConfig(content,labelName,"","del")
end function
'解密 
function decrypt_GetNewConfig(byval content,labelName)
	decrypt_GetNewConfig=decCode(handleNewConfig(content,labelName,"","get"))
end function
'加密
function encryption_SetNewConfig(content,labelName,byval labelValue) 
	content=handleNewConfig(content,labelName,encCode(labelValue),"set")
	encryption_SetNewConfig=content
end function



'call echo("aa=",getNewConfig(c, "aa"))
'call echo("aa=",setNewConfig(c, "aa","中国人"))
'call echo("bb=",setNewConfig(c, "bb","中国人"))
'call echo("aa=",delNewConfig(c, "aa"))
'配置内容设置 20161007
function handleNewConfig(byval c,labelName,byval labelValue,sType)
	dim startStr,endStr,s,nLen,leftS,contentS,rightS,tempC
	tempC=c
	c=vbcrlf & c & vbcrlf
	
	labelValue=replace(labelValue,"\\","\\")			'设置值特殊字符转义
	labelValue=replace(labelValue,"'","\'")			'设置值特殊字符转义
	labelValue=replace(labelValue,"=","\=")			'设置值特殊字符转义
	labelValue=replace(labelValue,vbcrlf,"\" & vbcrlf)			'设置值特殊字符转义
	
	c=replace(c, "'" & vbcrlf, "[|dyh vbcrlf+pie|]")	'处理'换行
	
	c=replace(c, "\\", "[|dyh shuanpie|]")				'替换特殊转义的'撇
	c=replace(c, "\'", "[|dyh danpie|]")				'替换特殊转义的'撇
	c=replace(c, "\=", "[|dyh dengpie|]")				'替换特殊转义的'撇
	c=replace(c, "\" & vbcrlf, "[|dyh pie+br|]")				'替换特殊转义的'撇 
	
	c=replace(c,  "[|dyh vbcrlf+pie|]", "'" & vbcrlf)  	'恢复'换行
	startStr=vbcrlf & labelName & "='"
	endStr="'" & vbcrlf
	nLen=instr(c,startStr)				'截取开始字符位置
	 
	if nLen>0 then 
		leftS=mid(c,1,nLen-1)			'内容左边字符
		c=mid(c,nLen+len(startStr) )	'左边往后内容
		nLen=instr(c,endStr)			'截取结束字符位置
		if nLen>0 then
			contentS=mid(c,1,nLen-1)		'内容中间字符
			rightS=mid(c,nLen+len(endStr))		'内容右边字符
		end if 
		'c=leftS & startStr & contentS & endStr & rightS		'组合后完整内容
		
		'获得值
		if sType="get" then
			s=contentS 
			s=replace(s, "[|dyh shuanpie|]", "\")				'替换特殊转义的'撇
			s=replace(s, "[|dyh danpie|]", "'")				'替换特殊转义的'撇
			s=replace(s, "[|dyh dengpie|]", "=")				'替换特殊转义的'撇
			s=replace(s, "[|dyh pie+br|]", vbcrlf)				'替换特殊转义的'撇 
		elseif sType="set" then
			s=leftS & startStr & labelValue & endStr & rightS
			s=mid(s,3,len(s)-4)
		elseif sType="del" then
			s=leftS & rightS
		end if 
		
		s=replace(s, "[|dyh shuanpie|]", "\\")				'替换特殊转义的'撇
		s=replace(s, "[|dyh danpie|]", "\'")				'替换特殊转义的'撇
		s=replace(s, "[|dyh dengpie|]", "\=")				'替换特殊转义的'撇
		s=replace(s, "[|dyh pie+br|]", "\" & vbcrlf)				'替换特殊转义的'撇 
		
		
	elseif sType="set" then
		s=tempC & startStr & labelValue & "'"
	end if
	
	handleNewConfig=s
end function 
%>      

