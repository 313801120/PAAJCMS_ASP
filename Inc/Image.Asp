<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-02-27
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%  

'������ʾͼƬ
function handleShowImg(imagename, setWidth, setHeight, sType)
    dim filePath 
    dim fs 

    filePath = handlePath(imagename) 
    '/////����GPS��ʵ��/////

    dim getPicSize 
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            '/////��ȡͼƬ����/////
            dim picSuffixName 
            picSuffixName = fs.getExtensionName(filePath) 
            dim PD                                          '//Picture Dimension
            dim PWidth, PHeight 
            select case picSuffixName
                case "gif", "bmp", "jpg", "png"
                    '/////����GPSͨ�����е�GetImageSize������ȡͼƬ�ߴ�/////

                    PD = getPicSize.getImageSize(filePath) 
                    PWidth = PD(1)                         '//��ȡͼƬ����
                    PHeight = PD(2)                        '//��ȡͼƬ�߶�
            end select
        set fs = nothing 
    set getPicSize = nothing 
    dim PXWidth, PXHeight 
    dim pp                                                                          '//������

    call echo("PWidth", PWidth) 
    call echo("PHeight", PHeight) 

    call echo("setWidth", setWidth) 
    call echo("setHeight", setHeight) 

    '���ͼƬ��Ϊ���Ϊ����
    if PWidth = 0 or PWidth = "" then
        PXWidth = 0 
        PXHeight = 0 
    else
        if setHeight < setWidth and setHeight > 0 then
            pp = formatNumber(PHeight / PWidth, 2)          '//�߿� �ٷֱ�

            PXHeight = setHeight 
            PXWidth = formatNumber(setHeight * pp, 0) 
        else
            pp = formatNumber(PWidth / PHeight, 2)          '//���� �ٷֱ�

            PXWidth = setWidth 
            PXHeight = formatNumber(setWidth / pp, 0) 
            call echo(pp, setWidth & "   " & setWidth / pp) 
        end if 
    end if 

    call echo("���ٷֱ�", formatNumber(PWidth / PHeight, 2)) 
    call echo("�߰ٷֱ�", formatNumber(PHeight / PWidth, 2)) 

    call echo("PXWidth", PXWidth) 
    call echo("PXHeight", PXHeight) 

    if sType = "��" then
        if setWidth < PXWidth then PXWidth = setWidth 
        handleShowImg = PXWidth 
    elseIf sType = "��" then
        if setHeight < PXWidth then PXWidth = setHeight 
        handleShowImg = PXWidth 
    else
        handleShowImg = "<img onerror=javascript:this.src='product_pic/wutu.jpg';this.width='92';this.height='130' src='" & imagename & "' border='0' width='" & PXWidth & "' height='" & PXHeight & "'>��ʾͼƬ" 
    end if 
    'response.Write "<br>"
    'response.write PXWidth&"X"&PXHeight  '��ǰͼƬ��ʾ�ĳߴ�
    'response.Write "<br>"
    'response.write PWidth&"X"&PHeight  'ͼƬԭʼ�ߴ�
'onload='javascript:if(this.width>940) this.width=940;'                ���������
end function 
'���ͼƬ����
function getImageWidthHeight(filePath, width, height)
    dim getPicSize, fs 
    call handlePath(filePath)                                                       '�������·��

	getImageWidthHeight = split("0|0", "|") 
	If checkFile(filePath) = False Then
		Exit Function
	End If
	If getFSize(filePath) = 0 Then
		Exit Function
	End If

	
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            dim picSuffixName 
            picSuffixName = fs.getExtensionName(filePath) 
            dim PD 
            select case picSuffixName
                case "gif", "bmp", "jpg", "png"
                    PD = getPicSize.getImageSize(filePath) 
                    width = PD(1)                          '//��ȡͼƬ����
                    height = PD(2)                         '//��ȡͼƬ�߶�
            end select
        set fs = nothing 
    set getPicSize = nothing 
    getImageWidthHeight = split(width & "|" & height, "|") 
end function 
'���ͼƬ��ֵ
function getImageWidthValue(filePath)
    dim width 
    call GetImageWidthHeight(filePath, width, "") 
    getImageWidthValue = width 
end function 
'����ļ���������
function getFileContentType(filePath)
    dim getPicSize, fs 
    call handlePath(filePath)                                                       '�������·��

	getFileContentType = ""
	If checkFile(filePath) = False Then
		Exit Function
	End If
	If getFSize(filePath) = 0 Then
		Exit Function
	End If
	
    set getPicSize = new GPS
        set fs = createObject("Scripting.FileSystemObject")
            dim PD 
			call echo("filePath",filePath):doevents
            PD = getPicSize.getImageSize(filePath) 
            getFileContentType = PD(0)                      '//��ȡͼƬ����
        set fs = nothing 
    set getPicSize = nothing 
end function 
'���ͼƬ����
function getImageType(filePath)
    getImageType = getFileContentType(filePath) 
end function 
'�ٲ�����ʽ��ʾͼƬ�б�
function imageWaterfall(folderPath, nRow, setWidth)
    dim content, splStr, i, imgPath, c, s, nProportion, width, height, thisI, nNumb 
    dim splC(99), splN(99) 
    nRow = nRow - 1                                                                 '��һ��ѭ����1��ʼ�Ͳ���ʾ���ݣ���
    '��ʼ��
    for i = 0 to nRow
        splC(i) = "" 
        splN(i) = 0 
    next 
    'Content = GetDirFileNameList(FolderPath,"")
    content = getFileFolderList(folderPath, true, "ȫ��", "", "ȫ���ļ���", "NO", "") 
    content = replace(LCase(content), LCase(server.mapPath("/")), "")               'ǿ���滻·��

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

'ͼƬ�ȱ�������(20151218)  �Զ��ӿ� �Ӹ�  ��padding-left ��padding-top    �� Call rw("<img src='"& tempFilePath &"'" & zoomImage(imgWidth, imgHeight, 35, 35, 0, 0, "") & " />")
function zoomImage(nWrapWidth, nWrapHeight, imgWidth, imgHeight, nSetWidth, nSetHeight, nLeft, nTop, sType)
    dim htmlStr, css, nNewWidth, nNewHeight 


    if nWrapWidth = "" then nWrapWidth = 0 
    if nWrapHeight = "" then nWrapHeight = 0 
    nWrapWidth = cint(nWrapWidth) : nWrapHeight = cint(nWrapHeight) 

    if nLeft = "" then nLeft = 0 
    if nTop = "" then nTop = 0 
    imgWidth = cint(imgWidth) : imgHeight = cint(imgHeight) : nSetWidth = cint(nSetWidth) : nSetHeight = cint(nSetHeight) : nLeft = cint(nLeft) : nTop = cint(nTop) 

    dim temp_nLeft, temp_nTop                                                       '�ݴ�
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