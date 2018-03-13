<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-13
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'Form

'表单提交
sub formSubmit()
    dim tableName, splStr, s, fieldName, fieldContent, fieldList, YZM  
    call OpenConn() 
    splStr = split(request.form(), "&") 
    tableName = rf("TableName") 
    YZM = trim(rf("YZM")) 

    if YZM <> "" then
        if getsession("YZM") <> YZM then
            call Javascript("返回", "验证码不正确", "") 
            response.end() 
        end if 
    end if 

    fieldList = lCase(getFieldList(tableName)) 
	
	
	'【@是jsp显示@】try{
    rs.open "Select * From [" & tableName & "]", conn, 1, 3 
    rs.addNew 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
        'FieldContent = Mid(S,InStr(S,"=")+1)
        fieldContent = rf(fieldName) 
        if inStr("," & fieldList & ",", "," & fieldName & ",") > 0 then
            rs(fieldName) = fieldContent '【@不是asp屏蔽@】  
        end if 
    'Call Echo(FieldName,FieldContent & "," & unescape(FieldContent)) 
    next 
    rs.update : rs.close 
	'【@是jsp显示@】}catch(Exception e){} 

    call Javascript("返回", "提交" & rf("DialogTitle") & "成功", "") 
end sub 

'获得POST字段名称列表 20160226
function getFormFieldList()
    dim s, c, splStr, fieldName 
    splStr = split(request.form(), "&") 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
        if c <> "" then c = c & "|" 
        c = c & fieldName 
    next 
    getFormFieldList = c 
end function 


'获得POST字段转JSON 20170311
function getPostToJson() 
	dim aspD, aspHelpTitle, aspHelpContent 
	set aspD = createObject("Scripting.Dictionary")  
    dim s, c, splStr, fieldName,fieldValue 
    splStr = split(request.form(), "&") 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
		fieldValue=request.Form(fieldName)
		aspD.add fieldName, fieldValue
    next  
	getPostToJson=json_encode(aspD)
end function 
'获得GET字段转JSON 20170311
function getGetToJson() 
	dim aspD, aspHelpTitle, aspHelpContent 
	set aspD = createObject("Scripting.Dictionary")  
    dim s, c, splStr, fieldName,fieldValue 
    splStr = split(request.QueryString(), "&") 
    for each s in splStr
        fieldName = lCase(mid(s, 1, inStr(s, "=") - 1)) 
		fieldValue=request.QueryString(fieldName)
		aspD.add fieldName, fieldValue
    next  
	getGetToJson=json_encode(aspD)
end function 
%> 










