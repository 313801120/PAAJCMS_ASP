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
'数据处理函数库


'dim rsDataArray(99,2),s,c,i,j,k
'call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
'call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
'call handleDataArray(rsDataArray,"add","asp","asp是active server pages")
'call handleDataArray(rsDataArray,"del","asp","asp是active server pages")
'call handleDataArray(rsDataArray,"edit","asp","eeeeeeeeeee")
'call handleDataArray(rsDataArray,"打印","","")
'操作二维数组 当数据库来用 20160727 
function handleDataArray(array2Data, byval sAction, fieldName, fieldValue) 
    dim title, content, i 
    sAction = lcase(sAction) 
    for i = 0 to uBound(array2Data)
        title = array2Data(i, 0) 
        content = array2Data(i, 1) 
        if sAction = "打印" or sAction = "print" then
            call echo(i & "、" & title, content) 
        elseif title <> "" then
            if(sAction = "修改" or sAction = "edit") and title = fieldName then
                array2Data(i, 0) = fieldName 
                array2Data(i, 1) = fieldValue 
                exit for 
            elseif(sAction = "删除" or sAction = "del") and title = fieldName then
                array2Data(i, 0) = "" 
                array2Data(i, 1) = "" 
                exit for 
            elseif(sAction = "获得" or sAction = "get") and title = fieldName then
                handleDataArray = content 
                exit function 
            end if 
        elseif sAction = "添加" or sAction = "add" then
            array2Data(i, 0) = fieldName 
            array2Data(i, 1) = fieldValue 
            exit for 
        end if 
    next 
end function 

'内容名称排序
function contentNameSort(content, sType)
    dim splStr, arrayStr(99), fileName, isOther, otherStr, sid, nIndex, c, s, i, left1 
    splStr = split(content, vbCrLf) 
    for each s in splStr
        if s <> "" then
            fileName = getStrFileName(s) 
            isOther = true 
            left1 = left(fileName, 1) 
            if inStr(fileName, "、") > 0 then
                sid = replace(left(fileName, 2), "、", "") 
                if isNumber(sid) then
                    nIndex = cint(sid) 
                    arrayStr(nIndex) = arrayStr(nIndex) & s & vbCrLf 
                    isOther = false 
                end if 
            end if 

            if inStr(sType, left1) = 0 and isOther = true then
                otherStr = otherStr & s & vbCrLf 
            end if 
        end if 
    next 
    for i = 0 to uBound(arrayStr)
        c = c & arrayStr(i) 
    next 
    contentNameSort = c & otherStr 
end function 


'删除内容有#号列表(20150818)
function remoteContentJingHao(content, sSplType)
    dim splStr, s, c 
    splStr = split(content, sSplType) 
    for each s in splStr
        if left(phpTrim(s), 1) <> "#" and left(phpTrim(s), 1) <> "_" then
            if c <> "" then c = c & sSplType 
            c = c & s 
        end if 
    next 
    remoteContentJingHao = c 
end function 
'删除数组有#号或这个_号列表(20150818)
function remoteArrayJingHao(splStr)
    dim s, c, sSplType 
    sSplType = "[|-|_]" 
    for each s in splStr
        if left(phpTrim(s), 1) <> "#" and left(phpTrim(s), 1) <> "_" then
            if c <> "" then c = c & sSplType 
            c = c & s 
        end if 
    next 
    splStr = split(c, sSplType) 
    remoteArrayJingHao = splStr 
end function 

'每个字符加指定值
function getEachStrAddValue(content, valueStr)
    dim i, s, c 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        c = c & s & valueStr 
    next 
    getEachStrAddValue = c 
end function 
'获得值在数组里位置 20150708
function getValueInArrayID(splStr, valueStr)
    dim i 
    getValueInArrayID = -1 
    for i = 0 to uBound(splStr)
        if splStr(i) = valueStr then
            getValueInArrayID = i 
            exit for 
        end if 
    next 
end function 
'判断值是否在数组里
function checkValueInArray(splStr, valueStr)
    dim i 
    checkValueInArray = false 
    for i = 0 to uBound(splStr)
        if splStr(i) = valueStr then
            checkValueInArray = true 
            exit for 
        end if 
    next 
end function 

'删除重复数组  20141220
function deleteRepeatArray(splStr)
    dim sSplType, s, c 
    sSplType = "[|-|_]" 
    for each s in splStr
        if s <> "" then
            if inStr(sSplType & c & sSplType, sSplType & s & sSplType) = 0 then
                c = c & s & sSplType 
            end if 
        end if 
    next 
    if c <> "" then c = left(c, len(c) - len(sSplType)) 
    splStr = split(c, sSplType) 
    deleteRepeatArray = splStr 
end function 
'删除重复内容 自定分割类型 20150724
function deleteRepeatContent(content, sSplType)
    dim splStr, s, c 
    splStr = split(content, sSplType) 
    for each s in splStr
        if s <> "" then
            if inStr(sSplType & c & sSplType, sSplType & s & sSplType) = 0 then
                if c <> "" then c = c & sSplType 
                c = c & s 
            end if 
        end if 
    next 
    deleteRepeatContent = c 
end function 

'获得数据总数
function getArrayCount(content, sSpl)
    dim splStr 
    getArrayCount = 0 
    if inStr(content, sSpl) > 0 or trim(content)<>"" then
        splStr = split(content, sSpl) 
        getArrayCount = uBound(splStr) + 1 
    end if 
end function 
'随机显示内容 randomShow("1,2,3,4,5,6,7,8,9", ",", 2)
function randomShow(content, sSplType, nSwitch)
    dim splStr, s, c, n, i 
    randomize 
    for i = 1 to nSwitch
        splStr = split(content, sSplType) 
        for each s in splStr
            n = cInt(rnd * 100) 
            if n > 50 then
                c = c & s & sSplType 
            else
                c = s & sSplType & c 
            end if 
        next 
        if c <> "" then c = left(c, len(c) - len(sSplType)) 
        content = c 
        c = "" 
    next 
    'Call Echo("C",C)
    randomShow = content 
end function 
'数组随机显示 ArrayRandomShow("1,2,3,4,5,6,7,8,9", 2)
function arrayRandomShow(byVal splStr, nSwitch)
    dim s, c, n, i, sSplType 
    sSplType = "[|-|_]" 
    randomize 
    for i = 1 to nSwitch
        for each s in splStr
            n = cInt(rnd * 100) 
            if n > 50 then
                c = c & s & sSplType 
            else
                c = s & sSplType & c 
            end if 
        'Call Echo(S,N)
        next 
        if c <> "" then c = left(c, len(c) - len(sSplType)) 
        splStr = split(c, sSplType) : c = "" 
    next 
    arrayRandomShow = splStr 
end function 
'打印数组内容
sub printArray(splStr)
    dim i, s 
    for i = 0 to uBound(splStr)
        s = splStr(i) 
        call echo(i, s) 
    next 
end sub 
'显示数组内容  (辅助上面)
sub echoArray(splStr)
    call PrintArray(splStr) 
end sub 
'返回打乱后的字符串。Shuffle=洗牌  2014 12 02
function str_Shuffle(content)
    dim i, s, c, n 
    randomize 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        if c = "" then
            c = s 
        elseIf len(c) = 1 then
            n = cInt(rnd * 100) 
            if n > 50 then
                c = c & s 
            else
                c = s & c 
            end if 
        else
            n = cInt(rnd * len(c)) + 1 
            c = mid(c, 1, n) & s & mid(c, n + 1) 
        end if 
    next 
    str_Shuffle = c 
end function 
'字符打乱
function characterUpset(content)
    characterUpset = str_Shuffle(content) 
end function 
'将字符串转换为数组   PHP里用到，暂时留着   把内容与几个字符分割成数组  如  abcefg,2     0=ab 1=ce 2=fg
function str_Split(content, nSplitLength)
    dim i, s, c, n, arrStr(99), nArr 
    if nSplitLength <= 0 then nSplitLength = 1 
    n = 0 : nArr = 0 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        c = c & s 
        n = n + 1 
        if n = nSplitLength then
            arrStr(nArr) = c 
            c = "" : n = 0 
            nArr = nArr + 1 
        end if 
    next 
    str_Split = arrStr 
end function 
'移除空值数组
function removeNullValueArray(content, sSplType)
    removeNullValueArray = handleArray(content, sSplType, "nonull") 
end function 
'移除重复数组
function removeRepeatValueArray(content, sSplType)
    removeRepeatValueArray = handleArray(content, sSplType, "norepeat") 
end function 
'处理数组
function handleArray(content, sSplType, sType)
    dim splStr, s, c, isTrue 
    sType = "|" & lCase(sType) & "|" 
    splStr = split(content, sSplType) 
    for each s in splStr
        isTrue = true 
        if inStr(sType, "|nonull|") > 0 and isTrue = true then
            if s = "" then isTrue = false 
        end if 
        if inStr(sType, "|norepeat|") > 0 and isTrue = true then
            if inStr(sSplType & c & sSplType, sSplType & s & sSplType) > 0 then isTrue = false 
        end if 
        if isTrue = true then c = c & s & sSplType 
    next 
    if c <> "" then c = left(c, len(c) - len(sSplType)) 
    handleArray = c 
end function  
'处理转字符(20151124)
function arrayToString(splStr, addtoStr)
    dim s, c  
    for each s in splStr
        if s <> "" then    
            c = c & s & addtoStr 
        end if 
    next  
    arrayToString = c  
end function 
'测试数组数据 20141217
sub testArrayData()
    dim nArrayData 
    nArrayData = array(3, 2, 4, 1, 6, 0) 
    call responseArray(nArrayData, "原来顺序") 
    call responseArray(nArraySelectSort(nArrayData), "选择排序") 
    call responseArray(nArrayQuickSort(nArrayData), "快速排序") 
    call responseArray(nArrayInsertSort(nArrayData), "插入排序") 
    call responseArray(nArrayBubbleSort(nArrayData), "冒泡排序") 
    call responseArray(nArrayReQuickSort(nArrayData), "反序排序") 
    response.write("<b>最 大 值：</b>" & nGetMaxArray(nArrayData) & "<hr>") 
    response.write("<b>最 小 值：</b>" & nGetMinArray(nArrayData) & "<hr>") 
end sub 
'===================================
'输出数组
'===================================
sub responseArray(nArrayData, str)
    dim s, i 
    s = "" 
    response.write("<b>" & str & "：</b>") 
    for i = 0 to uBound(nArrayData)
        s = s & nArrayData(i) & "," 
    next 
    s = left(s, len(s) - 1) 
    response.write(s) 
    response.write("<hr>") 
end sub
'===================================
'选择排序
'===================================
function nArraySelectSort(nArrayData)
    dim i, j, nK 
    dim nBound, nTemp 
    nBound = uBound(nArrayData) 

    for i = 0 to nBound - 1
        nK = i 
        for j = i + 1 to nBound
            if nArrayData(nK) > nArrayData(j) then
                nK = j 
            end if 
        next 
        nTemp = nArrayData(i) 
        nArrayData(i) = nArrayData(nK) 
        nArrayData(nK) = nTemp
    next 
    nArraySelectSort = nArrayData 
end function 
'===================================
'快速排序
'===================================
function nArrayQuickSort(nArrayData)
    dim i, j 
    dim nBound, nTemp 
    nBound = uBound(nArrayData) 
    for i = 0 to nBound - 1
        for j = i + 1 to nBound
            if nArrayData(i) > nArrayData(j) then
                nTemp = nArrayData(i) 
                nArrayData(i) = nArrayData(j) 
                nArrayData(j) = nTemp
            end if 
        next 
    next 
    nArrayQuickSort = nArrayData 
end function 
'===================================
'冒泡排序
'===================================
function nArrayBubbleSort(nArrayData)
    dim nBound 
    nBound = uBound(nArrayData) 
    dim isSorted, i,j, nTemp
    isSorted = false 
'    do while nBound > 0 and isSorted = false		不用这种
	for j=1 to 9999	
        isSorted = true 
        for i = 0 to nBound - 1
            if nArrayData(i) > nArrayData(i + 1) then
                nTemp = nArrayData(i) 
                nArrayData(i) = nArrayData(i + 1) 
                nArrayData(i + 1) = nTemp
                isSorted = false 
            end if 
        next 
        nBound = nBound - 1 
		if isSorted=false then 
			exit for
		end if
    next

    nArrayBubbleSort = nArrayData 
end function 
'===================================
'插入排序
'===================================
function nArrayInsertSort(nArrayData)
    dim nBound 
    nBound = uBound(nArrayData) 
    dim i, j, nT,nTemp

    for i = 1 to nBound
        nTemp = nArrayData(i) 
        j = i 
        do while nT < nArrayData(j - 1) and j > 0
            nArrayData(j) = nArrayData(j - 1) 
            j = j - 1  
        loop 
        nArrayData(j) = nTemp
    next 
    nArrayInsertSort = nArrayData 
end function 
'===================================
'快速排序-反序排列
'===================================
function nArrayReQuickSort(nArrayData)
    dim i, nBound, nArrayTemp 
    nArrayData = nArrayQuickSort(nArrayData) 
    nArrayTemp = nArrayQuickSort(nArrayData) 
    nBound = uBound(nArrayData) 
    for i = 0 to nBound
        nArrayData(i) = nArrayTemp(nBound - i) 
    next 
    nArrayReQuickSort = nArrayData 
end function 
'数组反向
function nArrayReverse(nArrayData)
    nArrayReverse = nArrayReQuickSort(nArrayData) 
end function 
 
'===================================
'求数组最大值
'=================================== 
function nGetMaxArray(nArrayData)
    dim i, j, nBound, nTemp 
    nArrayData = nArrayQuickSort(nArrayData) 
    nBound = uBound(nArrayData) 
    for i = 0 to nBound 
        for j = i + 1 to nBound
            if nArrayData(j) > nArrayData(i) then
                nTemp = nArrayData(i) 
                nArrayData(i) = nArrayData(j) 
                nArrayData(j) = nTemp 
            end if 
        next 
    next 
    nGetMaxArray = nArrayData(0) 
end function 
'===================================
'求数组最小值
'===================================
function nGetMinArray(nArrayData)
    dim i, j, nBound, nTemp  
    nArrayData = nArrayQuickSort(nArrayData) 
    nBound = uBound(nArrayData) 
    for i = 0 to nBound
        for j = i + 1 to nBound
            if nArrayData(j) > nArrayData(i) then
                nTemp  = nArrayData(i) 
                nArrayData(i) = nArrayData(j) 
                nArrayData(j) = nTemp  
            end if 
        next 
    next 
    nGetMinArray = nArrayData(nBound) 
end function 


 
'【删除.netc start】

'测试二维数组显示
function testTwoDimensionalArray()
    dim splStr, i 
    splStr = handleSplitArray("9-g|2-b|3-a|1-中", "|", "-") 
    for i = 0 to uBound(splStr)
        call echo(splStr(i, 1), splStr(i, 0)) : doevents 
    next 
    call echo("", "Ａｄｃ") 
    call TwoDimensionalArrayAsc(splStr)                                            '二维数组正排序
    for i = 0 to uBound(splStr)
        call echo(splStr(i, 1), splStr(i, 0)) : doevents 
    next 
    call echo("", "Desc") 
    call TwoDimensionalArrayDesc(splStr)                                           '二维数组倒排序
    for i = 0 to uBound(splStr)
        call echo(splStr(i, 1), splStr(i, 0)) : doevents 
    next 
end function 

'处理分割二维数组 20150313
function handleSplitArray(content, splOneType, splTowType)
    dim splA, splB, splStr, splxx, i, s, c, j, t, sSplType 
    sSplType = "[|Array|]" 
    splStr = split(content, splOneType) 
    for i = 0 to uBound(splStr)
        if splStr(i) <> "" then
            splxx = split(splStr(i), splTowType) 
            splA = splA & splxx(0) & sSplType 
            splB = splB & splxx(1) & sSplType 
        end if 
    next 
    if splA <> "" then splA = left(splA, len(splA) - len(sSplType)) 
    if splB <> "" then splB = left(splB, len(splB) - len(sSplType)) 
    splA = split(splA, sSplType) 
    splB = split(splB, sSplType) 

    reDim splStr(uBound(splA), uBound(splB)) 
    for i = 0 to uBound(splA)
        splStr(i, 0) = splA(i) 
        splStr(i, 1) = splB(i) 
    next 
    handleSplitArray = splStr 
end function 
'二维数组正排序 20150313
function twoDimensionalArrayAsc(arr2Date)
    dim i, j, t 
    for i = 0 to uBound(arr2Date)
        if isNul(arr2Date(i, 0)) = false then
            for j = i to uBound(arr2Date)
                if arr2Date(i, 0) > arr2Date(j, 0) then
                    t = arr2Date(i, 0) 
                    arr2Date(i, 0) = arr2Date(j, 0) 
                    arr2Date(j, 0) = t 
                    t = arr2Date(i, 1) 
                    arr2Date(i, 1) = arr2Date(j, 1) 
                    arr2Date(j, 1) = t 
                end if 
            next 
        end if 
    next 
    twoDimensionalArrayAsc = arr2Date 
end function 
'二维数组倒排序 20150313
function twoDimensionalArrayDesc(arr2Date)
    dim i, j, t 
    for i = 0 to uBound(arr2Date)
        if isNul(arr2Date(i, 0)) = false then
            for j = i to uBound(arr2Date)
                if arr2Date(i, 0) < arr2Date(j, 0) then
                    t = arr2Date(i, 0) 
                    arr2Date(i, 0) = arr2Date(j, 0) 
                    arr2Date(j, 0) = t 
                    t = arr2Date(i, 1) 
                    arr2Date(i, 1) = arr2Date(j, 1) 
                    arr2Date(j, 1) = t 
                end if 
            next 
        end if 
    next 
    twoDimensionalArrayDesc = arr2Date 
end function 
'【删除.netc end】
%>   


