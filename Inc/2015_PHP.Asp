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
'Version1.0 20150902

'php��������ݿ�ʱ�õ�
'call rw(implode(",",array("aa","bb","cc")))    ��ʾaa,bb,cc
function implode(sSplType,arrayList)
	dim s,c
	for each s in arrayList
		if c<>"" then
			c=c & sSplType
		end if
		c=c & s
	next
	implode=c
end function


'��ʽ��JS��PHP����
Function formattingJSPHP(ByVal content)
    Dim splStr, sx, tempSx, s, clearS, c, i, j, yesASP, nSYHCount, s1, s2, tempZc, rowC, tempRow, noteStr, downRow, downWord, nc, isLoop, beforeStr, afterStr, endCode,tempEndCode 
    Dim UpWord, UpWord2, yesOK 
    Dim wc                                                                          '�����ı��洢����
    Dim wcType                                                                      '�����ı����ͣ��� " �� '
    Dim zc                                                                          '��ĸ�ļ��洢����
    Dim tempS 
    Dim RSpaceStr, RSpaceNumb, addToRSpaceNumb                                      '�пո���ո���,׷������
    Dim MultiLineComment                                                            '����ע��
    MultiLineComment = ""                                                           '����ע��Ĭ��Ϊ��
    RSpaceStr = "    " : RSpaceNumb = 0                                             '�пո�ֵ �� �ո��� ���ó�ֵ
    nSYHCount = 0                                                                   '˫����Ĭ��Ϊ0
    Dim isAddToSYH                                                                  '�Ƿ��ۼ�˫����

    splStr = Split(content, vbCrLf)                                                 '�ָ���
    'ѭ������
    For j = 0 To UBound(splStr)
        s = splStr(j) 
        s = phptrim(s) 
        s = Replace(Replace(s, Chr(10), ""), Chr(13), "") '���Ϊʲô s����� chr(10)��chr(13) �أ�
        clearS = phpTrim(s)                                                             '������߿ո����˸�
        tempS = s                                                                       '�ݴ�S
        rowC = "" : tempRow = ""                                                        '���ÿ��ASP������ݴ������д���
        noteStr = ""                                                                    'Ĭ��ע�ʹ���Ϊ��
        downRow = ""                                                                    '��һ�д���
        downWord = ""                                                                   '��һ�е���
        addToRSpaceNumb = 0                                                             '���׷������
        If(j + 1) <= UBound(splStr) Then
            downRow = splStr(j + 1) 
            downWord = GetBeforeNStr(downRow, "ȫ��", 1) 
        End If 
        nc = ""                                                                         '����Ϊ��
        isLoop = True                                                                   'ѭ���ַ�Ϊ��
        'ѭ��ÿ���ַ�
        For i = 1 To Len(s)
            sx = Mid(s, i, 1) : tempSx = sx 
            beforeStr = Right(Replace(Mid(s, 1, i - 1), " ", ""), 1)                        '��һ���ַ�
            afterStr = Left(Replace(Mid(s, i + 1), " ", ""), 1)                             '��һ���ַ�
            endCode = Mid(s, i + 1)                                                         '��ǰ�ַ���������� һ��
            '�����ı�
            If(sx = """" Or sx = "'" And wcType = "") Or sx = wcType Or wc <> "" Then
                isAddToSYH = True 
                '����һ�ּ򵥵ķ�����������(20150914)
                If isAddToSYH = True And beforeStr = "\" Then
					
                    If Len(wc) >=1 Then
                        If isStrTransferred(wc)=true Then		'Ϊת���ַ�Ϊ��
							'call echo(wc,isStrTransferred(wc))
                            isAddToSYH = False 
                        End If 
                    Else
                        isAddToSYH = False 
                    End If 
					'call echo(wc,isAddToSYH)
                End If 
                If wc = "" Then
                    wcType = sx 
                End If 

                '˫�����ۼ�
                If sx = wcType And isAddToSYH = True Then nSYHCount = nSYHCount + 1             '�ų���һ���ַ�Ϊ\���ת���ַ�(20150914)


                '�ж��Ƿ�"�����
                If nSYHCount Mod 2 = 0 And beforeStr <> "\" Then
                    If Mid(s, i + 1, 1) <> wcType Then
                        wc = wc & sx 
						'call echo("wc",wc)
                        rowC = rowC & wc                      '�д����ۼ�
						'call echo("wc",wc)
                        nSYHCount = 0 : wc = ""               '���
                        wcType = "" 
                    Else
                        wc = wc & sx 
                    End If 
                Else
                    wc = wc & sx 
                End If

            '����ע��
            ElseIf(sx = "/" And afterStr = "*") Or MultiLineComment <> "" Then
                noteStr = Mid(s, i) 
                'call echo("����ע��",noteStr)
                If MultiLineComment <> "" Then MultiLineComment = MultiLineComment & vbCrLf 
                MultiLineComment = MultiLineComment & noteStr 
                If InStr(noteStr, "*/") > 0 Then
                    rowC = rowC & MultiLineComment 
                    MultiLineComment = "" 
                End If 
                Exit For 
            'ע�����˳� ��ѡע��
            ElseIf sx = "/" And afterStr = "/" Then
                noteStr = Mid(s, i) 
                rowC = rowC & noteStr 
                Exit For 

            '+-*\=&   �ַ����⴦��
            ElseIf InStr(".&=,+-*/:()><;", sx) > 0 And nc = "" Then
                If zc <> "" Then
                    tempZc = zc 
                    UpWord2 = UpWord                                                         '��¼����һ������
                    UpWord = LCase(tempZc)                                                   '��¼��һ����������   ΪСд
                    rowC = rowC & zc 
                    zc = ""                                                                  '�����ĸ����
                End If
				'�����if(1=1){;;;;;;;;}   20151224
				if sx=";" and instr("{;",right(trim(rowC),1))>0 then
					'call eerr(endcode,rowc)
					sx=""
				end if
                rowC = rowC & sx 
                UpWord2 = UpWord                                                                '��¼����һ������
                UpWord = sx 
            '��ĸ
            ElseIf checkABC(sx) = True Or sx = "_" Or zc <> "" Then
                zc = zc & sx 
                yesOK = True 
                s1 = Mid(s & " ", i + 1, 1) 
                s2 = Mid(zc, 1, 1) 
                If checkABC(s1) <> True And s1 <> "_" Then
                    yesOK = False 
                End If 
                '������������������
                If checkNumber(s1) = True And checkABC(s2) = True Then
                    yesOK = True 
                End If 
                If yesOK = False Then
                    'Remע��
                    If LCase(zc) = "rem" Then
                        sx = Mid(s, i + 1) : i = i + Len(sx) 
                        zc = zc & sx 
                    End If 

                    tempZc = zc 

                    UpWord2 = UpWord                                                         '��¼����һ������
                    UpWord = LCase(tempZc)                                                   '��¼��һ����������   ΪСд
                    rowC = rowC & zc 
                    zc = ""                                                                  '�����ĸ����
                End If 
            'Ϊ����
            ElseIf checkNumber(sx) = True Or nc <> "" Then
                nc = nc & sx 
                If afterStr <> "." And checkNumber(afterStr) = False Then
                    rowC = rowC & nc 
                    nc = "" 
                End If 
            Else
                'JS��PHP����{}����
                If sx = "{" Then 		    	
					If phptrim(rowC) <> "" Then
						addToRSpaceNumb = 1
					Else
						RSpaceNumb = RSpaceNumb + 1 
					End If  
                ElseIf sx = "}" Then
                    If RSpaceNumb > 0 Then
                        If phptrim(rowC) <> "" Then							
							'ɾ����̨ע��20151224
							tempEndCode =endCode
							if instr(tempEndCode,"//")>0 then
								tempEndCode=mid(tempEndCode,1,instr(tempEndCode,"//")-1) 
							end if					
                            '�޸�if(a==b){}  ����		
                            If phptrim(tempEndCode) <> "" Then
                                addToRSpaceNumb = -1 
                            Else
                                addToRSpaceNumb = 0 
                            End If 
                        Else
                            RSpaceNumb = RSpaceNumb - 1 
                        End If 
                    End If  
                End If 

                yesOK = True 
                If sx = " " And i > 1 Then
                    If Mid(s, i - 1, 1) = " " Then
                        yesOK = False 
                    End If 
                End If 
                If yesOK = True Then
                    rowC = rowC & sx 
                'call echo("sx",sx)
                End If 
            End If 
            '�ݴ�ÿ������
            tempRow = tempRow & sx 
            If isLoop = False Then Exit For 
        Next 
        If rowC = ";}" Then
            rowC = "}" 
        End If 
        If rowC <> "" And RSpaceNumb > 0 Then
            rowC = CopyStrNumb(RSpaceStr, RSpaceNumb) & Replace(rowC, vbCrLf, vbCrLf & CopyStrNumb(RSpaceStr, RSpaceNumb)) '�޽�20150902
        End If 
        RSpaceNumb = RSpaceNumb + addToRSpaceNumb 
        If MultiLineComment = "" Then                                                   '�޽�20150902
            c = c & rowC & vbCrLf 
        End If 
    Next 
    formattingJSPHP = c 
End Function 

 
'���� Install.php �ļ�
function printAccessToPHPInstallFile()
    Dim content, splTableName, tableName, splFieldName, fieldName, fieldType, c ,tempTableName
	dim tableNameList			'�������б�
    Call openconn() 
    content = getTableList() 
    'call echo("content",content)
    splTableName = Split(content, vbCrLf) 
    For Each tableName In splTableName
		tempTableName=tableName
        tableName = LCase(tableName)                                                    'ת��Сд
		if instr(tableName,"_")>0 then
			tableName=mid(tableName,instr(tableName,"_")+1) 
		end if
		'�ų��ظ�������
		if instr("|"& tableNameList &"|","|"& tableName &"|")=false then
			content = getFieldList(tempTableName) 
			Call echo(tableName, content) 
			splFieldName = Split(content, ",") 
			c = c & """{$DB_PREFIX}" & tableName & """ => ""CREATE TABLE `{$DB_PREFIX}" & tableName & "` (" & vbCrLf 
			c = c & "`id` int(5) unsigned NOT NULL auto_increment," & vbCrLf 
			For Each fieldName In splFieldName
				If fieldName <> "" Then 
					fieldName = LCase(fieldName)                                                    'ת��Сд
					fieldType = lcase(getFieldAlt(tempTableName, fieldName))
					'call echo(copyStr("&nbsp;",8) & "fieldName",fieldName   & "("& fieldType &")")        'mediumtext
					'call echo(fieldname,fieldType)
					If fieldType = "varchar" Then
						c = c & "`" & fieldName & "` varchar(255) NOT NULL default ''," & vbCrLf 
					ElseIf fieldType = "int" Then
						c = c & "`" & fieldName & "` int(8) NOT NULL default '0'," & vbCrLf 
					ElseIf fieldType = "float" Then
						c = c & "`" & fieldName & "` float(8) NOT NULL default '0'," & vbCrLf  
					ElseIf fieldType = "yesno" Then
						c = c & "`" & fieldName & "` tinyint(1) NOT NULL default '0'," & vbCrLf  						
						
					'  InStr("|aboutcontent|bodycontent|content|template|", "|" & fieldName & "|") > 0 Or InStr(fieldName, "module") > 0
					ElseIf fieldType="text" Then
						c = c & "`" & fieldName & "` mediumtext," & vbCrLf 
	
					Else
						c = c & "`" & fieldName & "` varchar(250) NOT NULL default ''," & vbCrLf 
					End If 
				End If 
			Next 
			c = c & "PRIMARY KEY  (`Id`)" & vbCrLf 
			c = c & "){$char};""," & vbCrLf 
		end if
    Next 
    Dim installFilePath, startStr, endStr, searchStr, replaceStr 
    startStr = "//access start" 
    endStr = "//access end" 
    installFilePath = "\PHP2\ImageWaterMark\Include\Install.php" 
    content = getftext(installFilePath) 
    searchStr = getStrCut(content, startStr, endStr, 1) 
    replaceStr = startStr & vbCrLf & c & endStr 
    content = Replace(content, searchStr, replaceStr) 
    'call eerr("",content)
    Call createfile(installFilePath, content) 
    Call echo("������ɣ���ӡ���ݿ⵽php�������ļ�", "<a href=""http://127.0.0.1/PHP2/ImageWaterMark/Include/Install.php?db_PREFIX=xy_"">http://127.0.0.1/PHP2/ImageWaterMark/Include/Install.php?db_PREFIX=xy_</a>") 
End function 
 
 'php����������� 20160624 ֻ�ܶԺ�������ı�������  �ӻ���js�����︴�ƹ����ĵģ�������
Function phpCodeConfusion(content)
    Dim splStr, i, isCode, YesWord, Sx, s, Wc, Zc, s1, AspCode, SYHCount, UpWord 
    Dim UpWordn, tempS, DimList 
    Dim YesFunction                                                                 '�����Ƿ�Ϊ��
    Dim StartFunction                                                               '��ʼ���� Ŀ����Ϊ����function default ������������û��()   20150218
    StartFunction = False                                                           'Ĭ�Ͽ�ʼ����Ϊ��
    'If nType="" Then  nType  = 0
    isCode = False                                                                  '��ASP Ĭ��Ϊ��
    YesFunction = False                                                             '�Ǻ��� Ĭ��Ϊ��
    YesWord = False                                                                 '�ǵ��� Ĭ��Ϊ��
    SYHCount = 0                                                                    '˫����Ĭ��Ϊ0
    splStr = Split(content, vbCrLf)                                                 '�ָ���
    Dim wcType                                                                      '�����ı����ͣ��� " �� '
    Dim isAddToSYH                                                                  '�Ƿ��ۼ�˫����
	dim beforeStr, afterStr, endCode,nSYHCount
    'ѭ������
    For Each s In splStr
        'ѭ��ÿ���ַ�
        For i = 1 To Len(s)
            Sx = Mid(s, i, 1) 
            'Asp��ʼ
            If Sx = "<" And Wc = "" Then                                                 '����ı�����Ϊ�� WcΪ������� ��"<%" �ų� �޸���20140412
                If Mid(s, i + 1, 1) = "?" Then
                    isCode = True                                                                   'ASPΪ��
                    i = i + 1                                                                       '��1�����ܼ�2��Ҫ��Ȼ<%function Test() �ͽ�ȡ����
                    Sx = Mid(s, i, 1) 
                    AspCode = AspCode & "<"
                End If 
            'ASP����
            ElseIf Sx = "?" And Mid(s, i + 1, 1) = ">" And Wc = "" Then                     'WcΪ�������
                isCode = False                                                                  'ASPΪ��
                i = i + 1                                                                       '���ܼ�2��ֻ�ܼ�1����Ϊ���ﶨ��ASPΪ�٣���������һ����ʾ����� 'ASP����Ϊ��
                Sx = Mid(s, i, 8) 
                AspCode = AspCode & "?" 
            End If 
            If isCode = True Then
             
				beforeStr = Right(Replace(Mid(s, 1, i - 1), " ", ""), 1)                        '��һ���ַ�
				afterStr = Left(Replace(Mid(s, i + 1), " ", ""), 1)                             '��һ���ַ�
				endCode = Mid(s, i + 1)                                                         '��ǰ�ַ���������� һ��
				'�����ı�
				If(sx = """" Or sx = "'" And wcType = "") Or sx = wcType Or wc <> "" Then
					isAddToSYH = True 
					'����һ�ּ򵥵ķ�����������(20150914)
					If isAddToSYH = True And beforeStr = "\" Then
						
						If Len(wc) >=1 Then
							If isStrTransferred(wc)=true Then		'Ϊת���ַ�Ϊ��
								'call echo(wc,isStrTransferred(wc))
								isAddToSYH = False 
							End If 
						Else
							isAddToSYH = False 
						End If 
						'call echo(wc,isAddToSYH)
					End If 
					If wc = "" Then
						wcType = sx 
					End If 
	
					'˫�����ۼ�
					If sx = wcType And isAddToSYH = True Then nSYHCount = nSYHCount + 1             '�ų���һ���ַ�Ϊ\���ת���ַ�(20150914)
	
	
					'�ж��Ƿ�"�����
					If nSYHCount Mod 2 = 0 And beforeStr <> "\" Then
						If Mid(s, i + 1, 1) <> wcType Then
							wc = wc & sx 
							AspCode = AspCode & wc                      '�д����ۼ�
							'call echo("wc",wc)
							nSYHCount = 0 : wc = ""               '���
							wcType = "" 
						Else
							wc = wc & sx 
						End If 
					Else
						wc = wc & sx 
					End If
					
                ElseIf Sx = "'" Then                                                            'ע�����˳�
                    AspCode = AspCode & Mid(s, i) 
                    Exit For 
                '��ĸ
                ElseIf checkABC(Sx) = True Or(Sx = "_" And Zc <> "") Or Zc <> "" Then
                    Zc = Zc & Sx 
                    s1 = LCase(Mid(s & " ", i + 1, 1)) 
                    If InStr("abcdefghijklmnopqrstuvwxyz0123456789", s1) = 0 And Not(s1 = "_" And Zc <> "") Then'����ж�
                        tempS = Mid(s, i + 1) 

                        If InStr("|function|sub|", "|" & LCase(Zc) & "|") Then
                            '������ʼ
                            If YesFunction = False And LCase(UpWord) <> "end" Then
                                YesFunction = True 
                                DimList = replace(getFunDimName(tempS),"$","")
                                StartFunction = True 
                            ElseIf YesFunction = True And LCase(UpWord) = "end" Then                        '�����һ������
                                YesFunction = False 
                            End If 
                        ElseIf YesFunction = True And LCase(Zc) = "var" Then
                            DimList = DimList & "," & replace(getVarName(tempS),"$","")
                        ElseIf YesFunction = True Then
                            '�ų���������ÿһ������
                            If StartFunction = False Then
                                Zc = replaceDim2(DimList, Zc) 
                            End If 
                            StartFunction = False 
                        End If 
                        UpWord = Zc                                                                     '��ס��ǰ����
                        AspCode = AspCode & Zc 
                        Zc = "" 
                    End If 
                Else
                    AspCode = AspCode & Sx 
                End If 
            Else
                AspCode = AspCode & Sx 
            End If 
            doEvents 
        Next 
        AspCode = RTrim(AspCode)                                                        'ȥ���ұ߿ո�
        AspCode = AspCode & vbCrLf                                                      'Asp����
        doEvents 
    Next 
    phpCodeConfusion = AspCode 
End Function 

%> 
