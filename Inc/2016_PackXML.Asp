<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-04
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 

class xmlZIP
    dim zipPathDir, zipPathFile, zipFileExt 
    dim starTime, endTime 
    dim pubXmlZipDir                                                  'xml���Ŀ¼
    dim pubIsEchoMsg                                                  '�Ƿ���Դ�ӡ��Ϣ
    dim webRootPathDir                                                '��վ��Ŀ¼·��

    sub callRun(thisWebRootPathDir, xmlZipDir, isEchoMsg, xmlFileName)
        pubXmlZipDir = xmlZipDir & "/"                                                  'xmlĿ¼
        pubIsEchoMsg = isEchoMsg                                                        '�Ƿ����
        '��ΪĬ�ϵ�ǰ�ļ���
        webRootPathDir = thisWebRootPathDir                                             '������
        'call echo("webRootPathDir",webRootPathDir)

        zipPathDir = server.mapPath("./") & "\" 

        '�ڴ˸���Ҫ����ļ��е�·��
        'ZipPathDir="D:\MYWEB\WEBINFO"
        '���ɵ�xml�ļ�
        zipPathFile = xmlFileName 
        if zipPathFile = "" then
            zipPathFile = "update.xml" 
        end if 
        '�����д�����ļ���չ��
        zipFileExt = "db;bak" 
        if right(zipPathDir, 1) <> "\" then zipPathDir = zipPathDir & "\" 

        response.write("����·����" & zipPathDir & zipPathFile & "<hr>") 
        '��ʼ���
        createXml(zipPathDir & zipPathFile) 
    end sub 
    '����Ŀ¼�ڵ������ļ��Լ��ļ���
    sub loadData(dirPath)
        dim xmlDoc 
        dim fso                                                                         'fso����
        dim objFolder                                                                   '�ļ��ж���
        dim objSubFolders                                                               '���ļ��м���
        dim objSubFolder                                                                '���ļ��ж���
        dim objFiles                                                                    '�ļ�����
        dim objFile                                                                     '�ļ�����
        dim objStream 
        dim pathname, xFolder, xFPath, xFile, xPath, xStream 
        dim pathNameStr 

        if pubIsEchoMsg = true then
            response.write("==========" & dirPath & "==========<br>") 
        end if 
        set fso = createObject("Scripting.FileSystemObject")
            set objFolder = fso.getFolder(dirPath)                                          '�����ļ��ж���

                if pubIsEchoMsg = true then
                    response.write(dirPath) 
                    response.flush 
                end if 
                set xmlDoc = createObject("Microsoft.XMLDOM")
                    xmlDoc.load(server.mapPath(zipPathFile)) 
                    xmlDoc.async = false 
                    'д��ÿ���ļ���·��
                    set xFolder = xmlDoc.selectSingleNode("//root").appendChild(xmlDoc.createElement("folder"))
                        set xFPath = xFolder.appendChild(xmlDoc.createElement("path"))
                            xFPath.text = replace(dirPath, webRootPathDir, "") 'zipPathDir  ע��
                            set objFiles = objFolder.files
                                for each objFile in objFiles
                                    if lCase(dirPath & objFile.name) <> lCase(request.serverVariables("PATH_TRANSLATED")) and lCase(dirPath & objFile.name) <> lCase(dirPath & zipPathFile) then
                                        if ext(objFile.name) then
                                            pathNameStr = dirPath & "" & objFile.name 
                                            if pubIsEchoMsg = true then
                                                response.write("---<br/>") 
                                                response.write(pathNameStr & "") 
                                                response.flush 
                                            end if 
                                            '================================================

                                            'д���ļ���·�����ļ�����
                                            set xFile = xmlDoc.selectSingleNode("//root").appendChild(xmlDoc.createElement("file"))
                                                set xPath = xFile.appendChild(xmlDoc.createElement("path"))
                                                    'call echo(zipPathDir,webRootPathDir)
                                                    xPath.text = replace(pathNameStr, webRootPathDir, "") 'zipPathDir  ע��
                                                    '�����ļ��������ļ����ݣ���д��XML�ļ���
                                                    set objStream = createObject("ADODB.Stream")
                                                        objStream.type = 1 
                                                        objStream.open 
                                                        objStream.loadFromFile(pathNameStr) 
                                                        objStream.position = 0 
                                                        set xStream = xFile.appendChild(xmlDoc.createElement("stream"))
                                                            xStream.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes" 
                                                            '�ļ����ݲ��ö��Ʒ�ʽ���
                                                            xStream.dataType = "bin.base64" 
                                                            xStream.nodeTypedValue = objStream.read() 
                                                        set objStream = nothing 
                                                    set xPath = nothing 
                                                set xStream = nothing 
                                            set xFile = nothing 
                                            '================================================

                                        end if 
                                    end if 
                                next 
                                if pubIsEchoMsg = true then
                                    response.write("<p>") 
                                end if 
                                xmlDoc.save(server.mapPath(zipPathFile)) 
                            set xFPath = nothing 
                        set xFolder = nothing 
                    set xmlDoc = nothing 
                    '���������ļ��ж���
                    set objSubFolders = objFolder.subFolders
                        '���õݹ�������ļ���
                        for each objSubFolder in objSubFolders
                            pathname = dirPath & objSubFolder.name & "\" 
                            loadData(pathname) 
                        next 
                    set objFolder = nothing 
                set objSubFolders = nothing 
            set fso = nothing 
    end sub
    '����һ���յ�XML�ļ���Ϊд���ļ���׼��
    sub createXml(filePath)
        '����ʼִ��ʱ��
        starTime = timer() 
        dim xmlDoc, root 
        set xmlDoc = createObject("Microsoft.XMLDOM")
            xmlDoc.async = false 
            set root = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
                xmlDoc.appendChild(root) 
                xmlDoc.appendChild(xmlDoc.createElement("root")) 
                if inStr(filePath, ":") = false then filePath = server.mapPath(filePath) 
                xmlDoc.save(filePath) 
            set root = nothing 
        set xmlDoc = nothing 
        'call eerr(ZipPathDir & "inc/",pubXmlZipDir)
        'call echo(ZipPathDir & "newweb/http���X�Xwww��thinkphp��cn�X/web/",pubXmlZipDir)
        loadData(pubXmlZipDir) 
        '�������ʱ��
        endTime = timer() 
        response.write("ҳ��ִ��ʱ�䣺" & formatNumber((endTime - starTime), 3) & "��") 
    end sub 
    '�ж��ļ������Ƿ�Ϸ�
    function ext(fileName)
        ext = true 
        dim temp_ext, e 
        temp_ext = split(zipFileExt, ";") 
        for e = 0 to uBound(temp_ext)
            if mid(fileName, inStrRev(fileName, ".") + 1) = temp_ext(e) then ext = false 
        next 
    end function 
end class 
%> 

