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
<!--#Include file = "Config.Asp"-->  

<% 
'2015_ASP.Asp   dimAddDefaultValue里设置默认变量为空

call CreateVBDialog() 
sub createVBDialog()
    dim splStr, i, s, c, content, rootFolder, zFormFolder, dataFolder, path, toPath, formName, basList, frmList, toolsMenu, frmListStr, openFrom, nTools 
    dim tempContent, startStr, endStr, splxx, functionName, centerStr 
    c = "|Print.Asp|Access.Asp|Cai.Asp|EncDec.Asp|Function.Asp|GetWord.Asp|RegExp.Asp|StringNumber.Asp|XML.Asp|Time.Asp|URL.Asp|FSO.Asp|Check.Asp|Html.Asp|AdodbStream.Asp|Dictionary.Asp|PinYin.Asp|2014_Array.Asp|2014_Css.Asp|2014_GBUTF.Asp|2015_Formatting.Asp|2015_Param.Asp|2015_ASP.Asp|2015_Color.Asp|FunHTML.Asp|2015_PHP.Asp|2015_WebSql.Asp|2015_hexSwitch.Asp" 

    zFormFolder = "D:\WindowsApplication1\zForm\Z_" 


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
            content = readFile("/inc/" & path, "") 

            searchStr = getStrCut(content, "'footer start", "'footer end", 2) 
            if searchStr <> "" then
                replaceStr = vbCrLf & getftext("\DataDir\VB模块\2015_Param.Asp.txt") 
                content = replace(content, searchStr, replaceStr) 
                call echo(path, "替换footer成功") 
            end if 
            '删除不处理的内容   'NOVBNet start                'NOVBNet end
            for i = 1 to 3
                searchStr = getStrCut(content, "'NOVBNet start", "'NOVBNet end", 2) 
                if searchStr <> "" then
                    content = replace(content, searchStr, "") 
                else
                    exit for 
                end if 
            next 
            '组件
            if path = "ZClassAspCode.Asp" then

            '模块
            else
                toPath = zFormFolder & formName & ".vb" 
                content = strCut(content, "<" & "%", "%" & ">", 2) 
                content = "Module " & formName & vbCrLf & content & vbCrLf & "End Module" 
            end if 
            tempContent = content 
            content = replace(content, "撇揦：星？“左右横。，", "撇揦：星？"" & Chr(-24144) & ""左右横。，") 

            if path = "Time.Asp" then
                content = replace(content, "d = CStr(Day(timeStr))", "d = Format(timeStr, ""dd"")") 
            end if 

            content = replace(content, vbTab, "    ") 
            content = ASPToVB(content) 
            content = ASPToVBNET(content) 


            if content = "" then
                call eerr(formName, tempContent) 
            end if 



            '函数为空
            splxx = split("getThisUrl|chkPost|getUpdateFile|showErr|printFormInfo|popupDownFile|downRemoteFile|c10to2|agent|getWebHome|webDebug|getThisUrlFileName|webDoMain|checkUrlName|checkUrlName|handleContentCode|replaceContentModule|replaceContentRowModule|addModuleReplaceArray|moduleFindContent|vBGetTimer|vBEchoTimer|||||", "|") 
            for each functionName in splxx
                if functionName <> "" then
                    startStr = "Function " & functionName 
                    endStr = "End Function" 
                    s = getStrCut(content, startStr, endStr, 1) 
                    if s <> "" then
                        centerStr = "" 

                        startStr = mid(s, 1, instr(s, ")")) 
                        if functionName = "GetFileText" then
                            centerStr = vbCrLf & "Call handlePath(filePath)    '获得完整路径 " & vbCrLf & "If IO.File.Exists(filePath) = True Then" & vbCrLf 
                            centerStr = centerStr & "    Dim TxtReader As IO.StreamReader = New IO.StreamReader(filePath, System.Text.Encoding.Default)" & vbCrLf 
                            centerStr = centerStr & "    MyTxtReader = TxtReader.ReadToEnd" & vbCrLf 
                            centerStr = centerStr & "    TxtReader.Close()" & vbCrLf 
                            centerStr = centerStr & "End If" & vbCrLf 
                        end if 
                        content = replace(content, s, startStr & vbCrLf & functionName & "=""" & centerStr & vbCrLf & endStr) 
                    end if 

                    startStr = "Sub " & functionName 
                    endStr = "End Sub" 
                    s = getStrCut(content, startStr, endStr, 1) 
                    if s <> "" then
                        startStr = mid(s, 1, instr(s, ")")) 
                        content = replace(content, s, startStr & vbCrLf & endStr) 
                    end if 
                end if 
            next 

            '删除函数
            splxx = split("getTimerSet|calculationTimer|getTimer|VBGetTimer|VBEchoTimer|errorText|getPort|||||||||||", "|") 
            for each functionName in splxx
                if functionName <> "" then
                    startStr = "Sub " & functionName 
                    endStr = "End Sub" 
                    s = getStrCut(content, startStr, endStr, 1) 
                    if s <> "" then
                        content = replace(content, s, "") 
                    end if 

                    startStr = "Function " & functionName 
                    endStr = "End Function" 
                    s = getStrCut(content, startStr, endStr, 1) 
                    if s <> "" then
                        content = replace(content, s, "") 
                    end if 
                end if 
            next 

            '替换
            content = regExp_Replace(content, "If Err Then", "If Err.Number Then") 
            content = regExp_Replace(content, " Err <> 0 ", " Err.Number <> 0 ") 
            content = regExp_Replace(content, "If Err Then", "If Err.Numbe Then") 
            content = regExp_Replace(content, "Set xml ", "xml") 
            content = regExp_Replace(content, "Set Stm ", "Stm") 





            content = replace(content, "App.Path ", "Application.StartupPath() ")           '系统路径
            content = replace(content, "C2 = ""撇揦：星？“左右横。，""", "C2 = ""撇揦：星？"" &  chr(""-24144"") &  ""左右横。，""") 

            content = replace(content, "Cstr(Day(s_Time))", "Cstr(Format(Now(), ""dd""))")  '删除set



            content = replace(content, " LenB(", " Len(") 
            content = replace(content, " AscB(", " Asc(") 
            content = replace(content, " MidB(", " Mid(") 
            content = replace(content, " ChrB(", " Chr(") 

            content = replace(content, "(LenB(", "(Len(") 
            content = replace(content, "(AscB(", "(Asc(") 
            content = replace(content, "(MidB(", "(Mid(") 
            content = replace(content, "(ChrB(", "(Chr(") 

            content = replace(content, """”""", "chr(-24143)") 
            content = replace(content, """“""", "chr(-24144)") 
            content = replace(content, "D = Cstr(Day(timeStr))", "D = CStr(CDate(timeStr).Day)") '在vb.net获得日期处理



            content = replace(content, vbTab, "    ")                                       '替换Tab
            call echo(formName, toPath) 
            call CreateFile(toPath, content) 
        end if 
    next 
end sub 
%> 


