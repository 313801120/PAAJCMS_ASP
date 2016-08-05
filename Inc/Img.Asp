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


'技巧：计算图片等比率缩放方法 宽=213,高=160  缩放后宽=100,高=75   100除213得到小数后两位，四舍五入
'document.write("\n<br>显示" + 100/213)
'document.write("\n<br>计算" + 160*0.47)


'用法
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,""))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"高"))
'Call Rw(HandleShowImg("UploadFiles/Soso.gif",100,60,"宽"))
'获得图片类型    getFileContentType("/aa.jpg")

'VB不引用 Start
'psd不行，等加
class GPS
    dim aso 
    private sub class_Initialize()
        set aso = createObject("ADODB.Stream")
            aso.mode = 3 
            aso.type = 1 
            aso.open 
    end sub
    private sub class_Terminate()
        set aso = nothing 
    end sub 
    private function bin2Str(bin)
        dim g, str, clow 
        for g = 1 to lenB(bin)
            clow = midB(bin, g, 1) 
            if ascB(clow) < 128 then
                str = str & chr(ascB(clow)) 
            else
                g = g + 1 
                if g <= lenB(bin) then str = str & chr(ascW(midB(bin, g, 1) & clow)) 
            end if 
        next 
        bin2Str = str 
    end function 
    private function binVal(bin)
        dim ret 
        dim h 
        ret = 0 
        for h = lenB(bin) to 1 step - 1
            ret = ret * 256 + ascB(midB(bin, h, 1)) 
        next 
        binVal = ret 
    end function 
    private function binVal2(bin)
        dim ret, g 
        ret = 0 
        for g = 1 to lenB(bin)
            ret = ret * 256 + ascB(midB(bin, g, 1)) 
        next 
        binVal2 = ret 
    end function 
    '///以下是调用代码/// 获得图片文件大小
    function getImageSize(filespec)
        dim bFlag, p1 
        if checkFile(filespec) = false then exit function                               '图片文件不存 则退出

        aso.loadFromFile(filespec) 
        'On Error Resume Next
        dim ret(3) 
        'call echo("filespec",filespec)
        bFlag = aso.read(3) 
        'Call Echo("bFlag1",Hex(binVal(bFlag)))
        select case hex(binVal(bFlag))
            case "4E5089"
                aso.read(15) 
                ret(0) = "PNG" 
                ret(1) = binVal2(aso.read(2)) 
                aso.read(2) 
                ret(2) = binVal2(aso.read(2)) 
            case "464947"
                aso.read(3) 
                ret(0) = "GIF" 
                ret(1) = binVal(aso.read(2)) 
                ret(2) = binVal(aso.read(2)) 
            case "FFD8FF"
                do
                do : p1 = binVal(aso.read(1)) : loop while p1 = 255 and not aso.EOS
                if p1 > 191 and p1 < 196 then exit do else aso.read(binval2(aso.read(2)) - 2) 
                do : p1 = binVal(aso.read(1)) : loop while p1 < 255 and not aso.EOS 
                loop while true 
                aso.read(3) 
                ret(0) = "JPG" 
                ret(2) = binval2(aso.read(2)) 
                ret(1) = binval2(aso.read(2)) 
            case else
                if left(bin2Str(bFlag), 2) = "BM" then
                    aso.read(15) 
                    ret(0) = "BMP" 
                    ret(1) = binval(aso.read(4)) 
                    ret(2) = binval(aso.read(4)) 
                else
                    ret(0) = "" 
                end if
        end select
        ret(3) = "width=""" & ret(1) & """ height=""" & ret(2) & """" 
        getImageSize = ret 
    end function 
end class 
'VB不引用 End

'处理显示图片
function handleShowImg(imagename, setWidth, setHeight, sType)
    dim filePath 
    dim fs 

    filePath = handlePath(imagename) 
    '/////建立GPS类实体/////

    dim getPicSize 
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            '/////获取图片类型/////
            dim picSuffixName 
            picSuffixName = fs.getExtensionName(filePath) 
            dim PD                                          '//Picture Dimension
            dim PWidth, PHeight 
            select case picSuffixName
                case "gif", "bmp", "jpg", "png"
                    '/////调用GPS通用类中的GetImageSize函数获取图片尺寸/////

                    PD = getPicSize.getImageSize(filePath) 
                    PWidth = PD(1)                         '//获取图片宽度
                    PHeight = PD(2)                        '//获取图片高度
            end select
        set fs = nothing 
    set getPicSize = nothing 
    dim PXWidth, PXHeight 
    dim pp                                                                          '//长宽比

    call echo("PWidth", PWidth) 
    call echo("PHeight", PHeight) 

    call echo("setWidth", setWidth) 
    call echo("setHeight", setHeight) 

    '获得图片宽为零或为空则
    if PWidth = 0 or PWidth = "" then
        PXWidth = 0 
        PXHeight = 0 
    else
        if setHeight < setWidth and setHeight > 0 then
            pp = formatNumber(PHeight / PWidth, 2)          '//高宽 百分比

            PXHeight = setHeight 
            PXWidth = formatNumber(setHeight * pp, 0) 
        else
            pp = formatNumber(PWidth / PHeight, 2)          '//长宽 百分比

            PXWidth = setWidth 
            PXHeight = formatNumber(setWidth / pp, 0) 
            call echo(pp, setWidth & "   " & setWidth / pp) 
        end if 
    end if 

    call echo("宽百分比", formatNumber(PWidth / PHeight, 2)) 
    call echo("高百分比", formatNumber(PHeight / PWidth, 2)) 

    call echo("PXWidth", PXWidth) 
    call echo("PXHeight", PXHeight) 

    if sType = "宽" then
        if setWidth < PXWidth then PXWidth = setWidth 
        handleShowImg = PXWidth 
    elseIf sType = "高" then
        if setHeight < PXWidth then PXWidth = setHeight 
        handleShowImg = PXWidth 
    else
        handleShowImg = "<img onerror=javascript:this.src='product_pic/wutu.jpg';this.width='92';this.height='130' src='" & imagename & "' border='0' width='" & PXWidth & "' height='" & PXHeight & "'>显示图片" 
    end if 
    'response.Write "<br>"
    'response.write PXWidth&"X"&PXHeight  '当前图片显示的尺寸
    'response.Write "<br>"
    'response.write PWidth&"X"&PHeight  '图片原始尺寸
'onload='javascript:if(this.width>940) this.width=940;'                加载完后处理
end function 
'获得图片宽高
function getImageWidthHeight(filePath, width, height)
    dim getPicSize, fs 
    call handlePath(filePath)                                                       '获得完整路径
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            dim picSuffixName 
            picSuffixName = fs.getExtensionName(filePath) 
            dim PD 
            select case picSuffixName
                case "gif", "bmp", "jpg", "png"
                    PD = getPicSize.getImageSize(filePath) 
                    width = PD(1)                          '//获取图片宽度
                    height = PD(2)                         '//获取图片高度
            end select
        set fs = nothing 
    set getPicSize = nothing 
    getImageWidthHeight = split(width & "|" & height, "|") 
end function 
'获得图片宽值
function getImageWidthValue(filePath)
    dim width 
    call GetImageWidthHeight(filePath, width, "") 
    getImageWidthValue = width 
end function 
'获得文件内容类型
function getFileContentType(filePath)
    dim getPicSize, fs 
    call handlePath(filePath)                                                       '获得完整路径
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            dim PD 
            PD = getPicSize.getImageSize(filePath) 
            getFileContentType = PD(0)                      '//获取图片类型
        set fs = nothing 
    set getPicSize = nothing 
end function 
'获得图片类型
function getImageType(filePath)
    getImageType = getFileContentType(filePath) 
end function 
'瀑布流方式显示图片列表
function imageWaterfall(folderPath, nRow, setWidth)
    dim content, splStr, i, imgPath, c, s, nProportion, width, height, thisI, nNumb 
    dim splC(99), splN(99) 
    nRow = nRow - 1                                                                 '减一，循环从1开始就不显示内容，晕
    '初始化
    for i = 0 to nRow
        splC(i) = "" 
        splN(i) = 0 
    next 
    'Content = GetDirFileNameList(FolderPath,"")
    content = getFileFolderList(folderPath, true, "全部", "", "全部文件夹", "NO", "") 
    content = replace(LCase(content), LCase(server.mapPath("/")), "")               '强力替换路径

    splStr = split(content, vbCrLf) : folderPath = replace(folderPath, "\", "/") 
    for each imgPath in splStr
        s = "<span><img src=" & host & imgPath & " width='" & setWidth & "'></span>" 
        call GetImageWidthHeight(imgPath, width, height) 
        if width > setWidth then
            nProportion = formatNumber(setWidth / width, 2) 
            width = setWidth 
            height = formatNumber(height * nProportion, 0) 
        end if 
        thisI = 0 : nNumb = "" 
        for i = 0 to nRow
            if nNumb = "" then
                nNumb = splN(i) 
            elseIf splN(i) < nNumb then
                thisI = i 
                nNumb = splN(i) 
            end if 
        next 
        splN(thisI) = splN(thisI) + height 
        splC(thisI) = splC(thisI) & s 
    next 
    for i = 0 to nRow
        c = c & "<div class=""ImgDivBox"">" & splC(i) & "</div>" & vbCrLf 
    next 
    c = c & "<div class=clear></div>" 
    imageWaterfall = c 
end function 

'图片等比例缩放(20151218)  自动加宽 加高  加padding-left 加padding-top    例 Call rw("<img src='"& tempFilePath &"'" & zoomImage(imgWidth, imgHeight, 35, 35, 0, 0, "") & " />")
function zoomImage(nWrapWidth, nWrapHeight, imgWidth, imgHeight, nSetWidth, nSetHeight, nLeft, nTop, sType)
    dim htmlStr, css, nNewWidth, nNewHeight 


    if nWrapWidth = "" then nWrapWidth = 0 
    if nWrapHeight = "" then nWrapHeight = 0 
    nWrapWidth = cint(nWrapWidth) : nWrapHeight = cint(nWrapHeight) 

    if nLeft = "" then nLeft = 0 
    if nTop = "" then nTop = 0 
    imgWidth = cint(imgWidth) : imgHeight = cint(imgHeight) : nSetWidth = cint(nSetWidth) : nSetHeight = cint(nSetHeight) : nLeft = cint(nLeft) : nTop = cint(nTop) 

    dim temp_nLeft, temp_nTop                                                       '暂存
    temp_nLeft = nLeft : temp_nTop = nTop 

    if imgWidth / imgHeight >= nSetWidth / nSetHeight then
        if imgWidth > nSetWidth then
            nNewWidth = nSetWidth 
            nNewHeight =(imgHeight * nSetWidth) / imgWidth 
        else
            nNewWidth = imgWidth 
            nNewHeight = imgHeight 
        end if 
    else
        if imgHeight > nSetHeight then
            nNewHeight = nSetHeight 
            nNewWidth =(imgWidth * nSetHeight) / imgHeight 
        else
            nNewWidth = imgWidth 
            nNewHeight = imgHeight 
        end if 
    end if 

    nNewWidth = int(nNewWidth) 
    nNewHeight = int(nNewHeight) 
    nLeft = nLeft + nSetWidth - nNewWidth 
    nTop = nTop + nSetHeight - nNewHeight 
    htmlStr = " width=""" & nNewWidth & """ height=""" & nNewHeight & """" 

    if nWrapHeight > 0 then
        'call echo("nWrapHeight",nWrapHeight)
        if instr(sType, "|imgulmiddel|") > 0 then
            nTop =(nWrapHeight - nNewHeight) / 2 
        end if 
    end if 







    if nLeft > 0 and nLeft <> temp_nLeft then
        css = css & "padding-left:" & nLeft & "px;" 
    end if 
    if nTop > 0 and nTop <> temp_nLeft then
        css = css & "padding-top:" & nTop & "px;" 
    end if 
    'Call echo("nLeft", nLeft)
    'Call echo("nTop", nTop)
    if css <> "" then
        css = " style=""" & css & """" 
    end if 
    zoomImage = htmlStr & css 
end function 

%>   

