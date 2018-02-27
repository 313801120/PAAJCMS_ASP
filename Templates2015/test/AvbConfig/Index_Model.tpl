#默认网页模板列表# start
【加载顶部框架】
<text>
{$Include file='head.html' block='top'$} 
</text>
---------------------
【】
<dIv>


<text>


<!--#循环自定义数组#-->
{$ForArray arraylist='aa|bb|cc' splitstr='|' nloop='11' default='[_2016年07月14日 07时02分]' operate='true' deletedefault='true'$}
<!--#test start#-->
<!--#[_2016年07月14日 07时02分] start#-->



 




<!--#读Left栏目#-->
<sPAn class="testspan">{$ReadColumeSetTitle title='[$fortitle$]' style='312' moreclass='leftmore' morestr='More' moreurl=' ' stylevalue='0' value='[$读出内容 block=\'BlockName网站公告\' file=\'\'$]'$}</sPAn>
 

[$forid$]/[$forcount$]=[$fortitle$]   , [$forid$]

<!--#[_2016年07月14日 07时02分] end#-->
<!--#test end#-->


<text default='<R#读出内容BlockName网站公告 start#>'/>
这里面放内容, 第三种调用方法
<text default='<R#读出内容BlockName网站公告 end#>'/>
</text>



</dIv>
---------------------
【加载底部框架】
<text>
{$Include file='Foot.html' block='foot'$} 
</text>
---------------------
#默认网页模板列表# end




#Source参数列表# start

#Source参数列表# end
[处理动作设置]{生成并IE打开}
[配置设置]{#dialogbackground='默认' dialogborder='默认' modulebackground='默认' moduleborder='默认' }
[样式前缀设置]{test_}
[帮助信息设置]{}
[Css保存文件设置]{style.css}
[ASP动作处理设置]{不处理动作}
[动作文件设置]{end
\Config\ASPAction\格式化HTML.Asp
【拷贝文件夹】/../Jquery【|】[$模板路径$][$网站目录名称$]/Jquery}
[默认HTML设置]{1、Template.Html}
[默认CSS设置]{10、2016style.css}
[自定义模板]{}
[自定义CSS]{}
[自定义模板动作]{}
[自定义CSS动作]{}
[txtIsCodeHtmlTemplate]{1}
[txtIsCodeCssTemplate]{1}
[txtIsDeleteRepeatCssClass]{1}
[txtCopyImageEncryption]{|left|nav|module| }

