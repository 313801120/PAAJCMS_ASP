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
'保存数据   ?act=saveData
Function saveData(sType)
	
	if getsession("yzm")="" then
		call eerr("提示","验证码失效")
	end if
	
	'if instr("|"& getFormFieldList() &"|","|yzm|") then		 
	if getsession("yzm")<>request.Form("yzm") then
		call eerr("提示","验证码错误")
	end if
	call setsession("yzm","")			'清空验证码
	  

    '保存文章评论
    If sType = "articlecomment" Then 		
		call autoSavePostData("", "tablecomment", "tablename||ArticleDetail,adddatetime|now,itemid||"& request("itemid") &",adddatetime,ip||"& getip())
    	Call echo("提示", "评论提交成功，等待管理员审核") 
		
    ElseIf sType = "feedback" Then
		if request.Form("guestname")="" then
			call eerr("提示","姓名为空")
		end if
		call autoSavePostData("", "feedback", "isthrough|numb|0,adddatetime|now,ip||"& getip() &",columnid||" & request.QueryString("columnid"))
		Call echo("提示", "反馈提交成功，等待管理员审核") 
    ElseIf sType = "guestbook" Then		
		if request.Form("guestname")="" then
			call eerr("提示","姓名为空")
		end if	
		call autoSavePostData("", "guestbook", "isthrough|numb|0,adddatetime|now,ip||"& getip() &",columnid||" & request.QueryString("columnid"))
		Call echo("提示", "留言提交成功，等待管理员审核")
	
	 ElseIf sType = "articledetail" Then	
		call autoSavePostData("", "articledetail", "title|bodycontent,adddatetime|now,ip||"& getip())
		Call echo("提示", "文章提交成功")		
    End If 
	response.End()
End Function  
'自动保存POST数据到表
function autoSavePostData(sID, tableName, fieldNameList)
	dim sql 
	sql=getPostSql(sID, tableName, fieldNameList)
    '检测SQL
    If checksql(sql) = False Then
        Call errorLog("出错提示4：<hr>sql=" & sql & "<br>") 
        Exit Function 
    End If
	'conn.execute(sql)			'checksql这一步就已经执行了不需要再执行了20160410
end function
'获得Post发送表单处理SQL语句 20160309
Function getPostSql(sID, tableName, fieldNameList)
    Dim valueStr, editValueStr, sql
    Dim splStr, splxx, s, fieldList 
    Dim fieldName, defaultFieldValue                                                '字段名称
    Dim fieldSetType                                                                '字段设置类型
    Dim fieldValue                                                                  '字段值

    Dim systemFieldList                                                             '表字段列表
    systemFieldList = getHandleFieldList(db_PREFIX & tableName, "字段配置列表") 

    Dim postFieldList                                                               'post字段列表
    Dim splPost, fieldContent, fieldConfig 
    postFieldList = getFormFieldList()  
	'以后再把下面与上面这两种处理方法事成一种看看行不行
    splPost = Split(postFieldList, "|") 
    For Each fieldName In splPost
        fieldContent = Request.Form(fieldName)  
        If InStr(systemFieldList, "," & fieldName & "|") > 0 And InStr("," & fieldList & ",", "," & fieldName & ",") = 0 Then
            '为自定义的
            If InStr(fieldNameList, "," & fieldName & "|") > 0 Then
                fieldConfig = Mid(fieldNameList, InStr(fieldNameList, "," & fieldName & "|") + 1) 
            Else
                fieldConfig = Mid(systemFieldList, InStr(systemFieldList, "," & fieldName & "|") + 1) 
            End If 
            fieldConfig = Mid(fieldConfig, 1, InStr(fieldConfig, ",") - 1) 
            'call echo("config",fieldConfig)
            'call echo(fieldName,fieldContent)
			'call echo("fieldConfig",fieldConfig)
            splxx = Split(fieldConfig & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
			fieldSetType="":defaultFieldValue=""
			'【@是jsp显示@】try{
            fieldSetType = splxx(1)                                                         '字段设置类型
            defaultFieldValue = splxx(2)                                                    '默认字段值			
			'【@是jsp显示@】}catch(Exception e){}
            fieldValue = ADSqlRf(fieldName)                                                 '代替上面，因为它处理了'符号
            'call echo("fieldValue",fieldValue)
			'排序密码不处理
			if fieldValue<>"#NO******NO#" then
				'md5加密
				If fieldSetType = "md5" Then
					fieldValue = myMD5(fieldValue) 
				End If 
	
				If fieldSetType = "yesno" Then
					If fieldValue = "" Then
						fieldValue = defaultFieldValue 
					End If 
				'不为数字类型加单引号
				ElseIf fieldSetType = "numb" Then
					If fieldValue = "" Then
						fieldValue = defaultFieldValue 
					End If 
	
				ElseIf fieldName = "flags" Then
					'PHP里用法
					If EDITORTYPE = "php" Then
						If fieldValue <> "" Then									'【@不是php屏蔽@】
							fieldValue = "|" & arrayToString(fieldValue, "|") '【@不是php屏蔽@】
						End If '【@不是php屏蔽@】
					Else
						fieldValue = "|" & arrayToString(Split(fieldValue, ", "), "|") 
					End If 
	
	
					fieldValue = "'" & fieldValue & "'" 
	
				'为时间
				ElseIf fieldSetType = "time" or fieldSetType = "now" Then
					If fieldValue = "" Then
						fieldValue = cstr(now())	'给.net用 
					End If 
					fieldValue = "'" & fieldValue & "'" 
				'为时期
				ElseIf fieldSetType = "date" Then
					If fieldValue = "" Then
						fieldValue = Date() 
					End If 
					fieldValue = "'" & fieldValue & "'" 
	
				Else
					fieldValue = "'" & fieldValue & "'" 
				End If 
				if EDITORTYPE <> "asp" then
					fieldValue=unescape(fieldValue)			'解码20160418   把+处理掉了为为什么要解码？可能在php里get方式发送时(20160810) 
				end if
				If valueStr <> "" Then 
					valueStr = valueStr & "," 
					editValueStr = editValueStr & "," 
				End If  
				valueStr = valueStr & fieldValue 
				editValueStr = editValueStr & fieldName & "=" & fieldValue 
			end if
			If fieldList <> "" Then
				fieldList = fieldList & ","  
			End If 
			fieldList = fieldList & fieldName 
			
			
        End If 
    Next  
    '自定义字段是否需要写入默认值  有的
    splStr = Split(fieldNameList, ",") 
    For Each s In splStr
        If InStr(s, "|") > 0 Then
            splxx = Split(s & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
			fieldSetType="":fieldValue="" 
			'【@是jsp显示@】try{
            fieldSetType = splxx(1)                                                         '字段设置类型
            fieldValue = splxx(2)                                                           '默认字段值
			'【@是jsp显示@】}catch(Exception e){}

            If InStr(systemFieldList, "," & fieldName & "|") > 0 And InStr("," & fieldList & ",", "," & fieldName & ",") = False Then
			  
                if fieldSetType = "date" and fieldValue="" then
                    fieldValue = Date() 
                elseif (fieldSetType = "time" or fieldSetType = "now") and fieldValue="" then
                    fieldValue =cstr(Now())			'给.net用 
				end if
				If fieldSetType <> "yesno" And fieldSetType <> "numb" Then
                    fieldValue = "'" & fieldValue & "'" 
                End If 
                If fieldList <> "" Then
                    fieldList = fieldList & "," 
                    valueStr = valueStr & "," 
                    editValueStr = editValueStr & "," 
                End If 
                fieldList = fieldList & fieldName 
                valueStr = valueStr & fieldValue 
                editValueStr = editValueStr & fieldName & "=" & fieldValue 
            'call echo(fieldName,fieldSetType)
            End If 
        End If 
    Next
	
	dim idname:idname=request("idname")
	if idname="" then
		idname="id"
	end if

    If sID = "" Then
        sql = "insert into " & db_PREFIX & "" & tableName & " (" & fieldList & ",updatetime) values(" & valueStr & ",'" & Now() & "')" 
    Else
		'判断是否自动添加修改字段
		if instr(systemFieldList,",updatetime|")>0 then
			editValueStr=editValueStr & ",updatetime='" & Now() & "'"
		end if
        sql = "update " & db_PREFIX & "" & tableName & " set " & editValueStr & " where "& idname &"=" & sID 
    End If 
    getPostSql = sql 
End Function 
%> 

