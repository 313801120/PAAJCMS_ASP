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
'Dictionary 字典 (2013,9,27)

'开始
dim aspD, aspHelpTitle, aspHelpContent 
set aspD = createObject("Scripting.Dictionary")
aspD.add "Abs", "返回数字的绝对值" 
aspD.add "Sqr", "返回数值表达式的平方根" 
aspD.add "Sgn", "返回表示数字符号的整数" 
aspD.add "Rnd", "返回一个随机生成的数字" 
aspD.add "Log", "返回指定数值的自然对数" 
aspD.add "Exp", "返回e的指定次幂" 
aspD.add "Sin", "返回指定弧度的正弦值" 
aspD.add "Cos", "返回指定弧度的余弦值" 
aspD.add "Tan", "返回指定弧度的正切值" 
aspD.add "Atn", "返回指定数值的反正切值" 
aspD.add "Int", "返回指定数字的整数部分" 
aspD.add "Round", "对参数按指定的位数进行四舍五入" 
aspD.add "Fix", "返回指定数值的整数部分" 
aspD.add "Now", "返回当前的日期时间" 
aspD.add "Date", "返回当前的系统日期" 
aspD.add "Time", "返回当前的系统时间" 
aspD.add "Year", "从指定的表达式中提取年份" 
aspD.add "Month", "从指定的表达式中提取月份" 
aspD.add "Day", "从指定表达式中提取日期" 
aspD.add "Hour", "从指定的表达式中提取整点值" 
aspD.add "Minute", "从指定的表达式中提取分钟数" 
aspD.add "Second", "从指定的表达式中提取秒数" 
aspD.add "DateAdd", "将指定日期添加上指定时间间隔并返回新值" 
aspD.add "DateDiff", "计算两个日期之间的时间间隔" 
aspD.add "DatePart", "返回给定日期中的某一部分" 
aspD.add "DateSerial", "根据提供的年月日的数值返回一个Date子类型的数据" 
aspD.add "DateValue", "将指定的字符串转化为Date子类型的数据" 
aspD.add "MonthName", "根据指定的数字返回代表该月份的字符串" 
aspD.add "Weekday", "返回代表一周中某一天的整数" 
aspD.add "WeekdayName", "根据周次返回代表该周次的字符串" 
aspD.add "TimeSerial", "根据提供的时分秒的数值返回一个Date子类型的数据" 
aspD.add "TimeValue", "将指定的字符串转化为Date子类型的数据" 
aspD.add "CBool", "将表达式转化为Boolean子类型" 
aspD.add "CByte", "将表达式转化为Byte子类型" 
aspD.add "CDbl", "将表达式转化为Double子类型" 
aspD.add "Chr", "将ANSI字符代码（～）转化为对应的字符" 
aspD.add "CInt", "将表达式转化为Integer数值子类型" 
aspD.add "CLng", "将表达式转化为Long数值子类型" 
aspD.add "Cstr", "将表达式转化为String子类型" 
aspD.add "CCur", "将表达式转化为Currency子类型" 
aspD.add "CDate", "将日期表达式转化为Date子类型" 
aspD.add "CSng", "将表达式转化为Single数值子类型" 
aspD.add "Asc", "返回与字符串首字母相对应的ANSI字符代码（～）" 
aspD.add "Oct", "返回八进制的字符串" 
aspD.add "Hex", "返回十六进制的字符串" 
aspD.add "FormatNumber", "格式化指定的数值表达式" 
aspD.add "FormatDateTime", "格式化指定的日期时间表达式" 
aspD.add "FormatCurrency", "格式化指定的货币表达式" 
aspD.add "FormatPercent", "格式化某一数值为百分数" 
aspD.add "IsArray", "判断某一变量是否为数组变量" 
aspD.add "IsDate", "判断某一变量是否为或者可以转化为日期表达式" 
aspD.add "IsEmpty", "判断变量是否为空" 
aspD.add "IsNull", "判断表达式中是否包含有效数据" 
aspD.add "IsNumeric", "判断表达式的值是否为数字" 
aspD.add "Is_object_", "判断表达式的值是否为某一对象" 
aspD.add "VarType", "返回指定变量的子类型" 
aspD.add "TypeName", "返回指定变量子类型的信息" 
aspD.add "MsgBox", "弹出含有提示信息的对话框" 
aspD.add "InputBox", "弹出可输入数据的对话框" 
aspD.add "ScriptEngine", "返回一个代表当前所使用的脚本语言的字符串" 
aspD.add "ScriptEngineBuildVersion", "返回所使用的脚本引擎的编译版本号" 
aspD.add "ScriptEngineMajorVersion", "返回所使用的脚本引擎的主版本号" 
aspD.add "ScriptEngineMinorVersion", "返回所使用脚本引擎的次版本号" 
aspD.add "Array", "利用所给的参数返回一个数组" 
aspD.add "UBound", "返回数组的最大可用下标" 
aspD.add "LBound", "返回数组的最小可用下标" 
aspD.add "Filter", "返回一个数组其中包含了满足条件的字符串" 
aspD.add "InStr", "返回某字符串在指定字符串中第一次出现的位置" 
aspD.add "InStrRev", "返回某字符串在指定字符串中从结尾算起的出现位置" 
aspD.add "Join", "将数组中的字符串连接为一个字符串" 
aspD.add "LCase", "将字符串中的字母全部转换为小写形式" 
aspD.add "Left", "从字符串的左端取指定数目的字符" 
aspD.add "Len", "返回字符串的长度或者存储某一变量所需要的字节数" 
aspD.add "Ltrim", "返回不带前导空格的字符串" 
aspD.add "Replace", "将字符串中指定字符串替换为其他内容" 
aspD.add "Right", "从字符串的右侧获取指定数目的字符" 
aspD.add "Rtrim", "删除字符串的后续空格" 
aspD.add "Mid", "从字符串中指定位置获取指定长度的字符串" 
aspD.add "Split", "返回包含指定数目子字符串的一维数组" 
aspD.add "StrComp", "对两个指定的字符串进行比较" 
aspD.add "StrReverse", "将指定字符串倒序排列并返回重新排列后的字符串" 
aspD.add "String", "返回由指定数目的某一字符组成的字符串" 
aspD.add "Trim", "删除字符串的前导和后续空格" 
aspD.add "Space", "返回由指定数目的空格组成的字符串" 
aspD.add "UCase", "将字符串中的字母全部转换为大写形式" 
aspD.add "LoadPicture", "返回一个图片对象" 
aspD.add "Create_object_", "创建并返回一个对象" 
aspD.add "Get_object_", "从指定文件中获得一个Automation对象" 
aspD.add "RGB", "返回代表某一颜色的整数" 
aspHelpTitle = aspD.keys 
aspHelpContent = aspD.items 
'获得ASP内置函数帮助说明
function ASPHelpTip(byVal funName, byVal content)
	dim i, s, c 
	for i = 0 to aspD.count - 1
		if LCase(aspHelpTitle(i)) = LCase(funName) then
			if LCase(funName) = "chr" then
				s = left(content, inStr(content, ")") - 1) 
				s = replace(s, "(", "") 
				c = vbCrLf & "Chr(" & s & ")=" 
				if checkNumber(s) = true then                  '判断Chr 里面是否为数字
					if s = "10" or s = "13" then
						c = c & "(回车键)" 
					elseIf s = "32" then
						c = c & "(空格)" 
					elseIf s = "0" then
						c = c & "(中断程序,停止执行)" 
					else
						c = c & "(" & chr(s) & ")" 
					end if 
					c = replace(c, "'", "&#x27;") 
				else
					c = c & s 
				end if 
			end if 
			ASPHelpTip = aspHelpContent(i) & c 
			exit for 
		end if 
	next 
end function 
'处理用户输入的值
function handleWrite(jc)
	dim s 
	s = left(jc, len(jc) - 1) 
	s = right(s, len(s) - 1) 
	s = LCase(trim(s)) 
	if s = "adodb.connection" then
		jc = """Adodb.Connection""" 
	elseIf s = "adodb.recordset" then
		jc = """Adodb.RecordSet""" 
	end if 
	handleWrite = jc 
end function 
'结束
'简体与繁体转换 用法 JianTiTabFanTi("罢摆败颁办绊帮绑镑谤",2)2为繁体
function jianTiTabFanTi(path, content, sType)
	dim jianStr, fanStr, jianS, fanS, i, s, c 
	c = getFText(path) & vbCrLf 
	sType = cstr(sType)                             '字符类型
	jianStr = strCut(c, "简体" & vbCrLf, vbCrLf, 2) 
	fanStr = strCut(c, "繁体" & vbCrLf, vbCrLf, 2) 
	if 1 = 2 then
		if len(content) * 2 < len(jianStr) then
			for i = 1 to len(content)
				s = mid(content, i, 1) 
				'繁体转简体
				if sType = "2" and inStr(jianStr, s) > 0 then
					jianS = mid(jianStr, inStr(jianStr, s), 1) 
					fanS = mid(fanStr, inStr(jianStr, s), 1) 
					content = replace(content, jianS, fanS) 
					call echo("", "") 
				elseIf inStr(fanStr, s) > 0 then
					jianS = mid(jianStr, inStr(fanStr, s), 1) 
					fanS = mid(fanStr, inStr(fanStr, s), 1) 
					content = replace(content, fanS, jianS) 
				end if 
			next 
			jianTiTabFanTi = content 
			exit function 
		end if 
	end if 
	for i = 1 to len(jianStr)
		jianS = mid(jianStr, i, 1) 
		fanS = mid(fanStr, i, 1) 
		'繁体转简体
		if sType = "2" then
			content = replace(content, jianS, fanS) 
		else
			content = replace(content, fanS, jianS) 
		end if 
	next 
	jianTiTabFanTi = content 
end function 
'近意词替换
function jiYiCiReplace(path, content)
	dim splStr, splxx, i, s, c 
	splStr = split(getFText(path), vbCrLf) 
	for each s in splStr
		if inStr(s, ",") and left(trim(s), 1) <> "#" then
			splxx = split(s, ",") 
			if uBound(splxx) >= 1 then
				content = replace(content, splxx(0), splxx(1)) 
			end if 
		end if 
	next 
	jiYiCiReplace = content 
end function 
'替换字符
function ciReplace(THContent, content)
	dim splStr, splxx, i, s, c 
	splStr = split(THContent, vbCrLf) 
	for each s in splStr
		if inStr(s, ",") and left(trim(s), 1) <> "#" then
			splxx = split(s, ",") 
			if uBound(splxx) >= 1 then
				content = replace(content, splxx(0), splxx(1)) 
			end if 
		end if 
	next 
	ciReplace = content 
end function 

%>  


