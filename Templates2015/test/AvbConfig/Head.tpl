#Ĭ����ҳģ���б�# start
��������ʼ��ǩ��
<text>
<!--#top start#-->
</text>
---------------------
#���̶���ʼ��



<style>
.topbox{
    position:fixed;
    left:0px;
    top:0px;
    height:160px;
    width:100%;
    background-color:#FFFFFF;
    z-index: 1;
}
.topboxspace{
    height:160px;
    width:100%;
}
</style>
<body>
<text>
<div class="topboxspace"></div>
<div class="topbox">
</text>
---------------------
������Bar��


<dIv>



<style>
.lightbox{

}
.lightbox .black_overlay{
    display: none;
    position: absolute;
    top: 0%;
    left: 0%;
    width: 100%;
    height: 100%;
    background-color: black;
    z-index:1001;
    -moz-opacity: 0.3;
    opacity:.30;
    filter: alpha(opacity=30);
}
.lightbox .white_content {
    display: none;
    position: absolute;
    top: 100px;
    left: 25%;
    width: 450px;
    height: 490px;
    padding: 1px;
    border: 10px solid #000000;
    background-color:#FFFFFF;
    z-index:1002;
    overflow:hidden;
}
/*������*/
.topbarbox{
    width:100%;
    background:#F7F7F7;
    color:#666666;
    height:26px;
    line-height:26px;
    border-bottom:1px solid #D8D8D8;
    font-size:12px;
}
/*Ĭ��������ʽ*/
.topbarbox a{font-size:12px;color:#666666}
.topbarbox a:hover{text-decoration:none;color:#333333}
/*topbar���*/
.topbarbox .topbar{
    width:980px;
    margin:0 auto;
}
/*��߲���*/
.topbarbox .topbar .left{
    float:left;
}
    /*΢��*/
.topbarbox .topbar .left .weixin{
    color:#333333;
    background-image:url(icon_gz.gif);
    background-position: 0px -75px;
    background-repeat:no-repeat;
    padding-left:24px;
    display:block;
    float:left;
}
    /*��ɫ��ע*/
.topbarbox .topbar .left a.greengz{
    background-image:url(icon_gz.gif);
    background-position: 0px -234px;
    background-repeat:no-repeat;
    width:39px;
    height:16px;
    display:block;
    float:left;
    margin:4px 0 0 4px;
}
    /*΢��*/
.topbarbox .topbar .left .weibo{
    color:#333333;
    background-image:url(icon_gz.gif);
    background-position: 0px 3px;
    background-repeat:no-repeat;
    padding-left:24px;
    display:block;
    float:left;
}
    /*��ɫ��ע*/
.topbarbox .topbar .left a.redgz{
    background-image:url(icon_gz.gif);
    background-position: 0px -196px;
    background-repeat:no-repeat;
    width:39px;
    height:16px;
    display:block;
    float:left;
    margin:4px 0 0 4px;
}
    /*�����*/
.topbarbox .topbar .left .line{
    float:left;
    display:block;
    margin:5px 10px 0 0;
    color:#FF0000;
    background:green;
}
#localtime{
    float:left;
    display:block;
}
/*�ұ߲���*/
.topbarbox .topbar .right{
    float:right;
}
.topbarbox .topbar .right a.weiboico{
    background-image:url(icon_gz.gif);
    background-position: 0px -42px;
    background-repeat:no-repeat;
    padding-left:28px;
}
.topbarbox .topbar .right a.tqqico{
    background-image:url(icon_gz.gif);
    background-position: 0px -118px;
    background-repeat:no-repeat;
    padding-left:28px;
}
</style>
<body>
<text>
<sPAn class="testspan">{$copyTemplateMaterial dir='TopBar/1/' isHandle='true'$}</sPAn>


<div class="lightbox">
    <div id="light" class="white_content">
        <sPAn class="testspan">{$MainInfo title='΢�Ź�ע��ͼ' showtitle='΢�Ź�ע��ͼ' default='[_΢�Ź�ע��ͼ2015��01��28�� 09ʱ59��]' autoadd='true'$}</sPAn>
<!--#test start#-->
<!--#[_΢�Ź�ע��ͼ2015��01��28�� 09ʱ59��] start#-->
<img src="nwwx.jpg">
<!--#[_΢�Ź�ע��ͼ2015��01��28�� 09ʱ59��] end#-->
<!--#test end#--> 
    </div>
    <div id="fade" class="black_overlay" onClick="CloseLightBox('light','fade');"></div>
</div>

<div class="topbarbox">
    <div class="topbar">
        <div class="left">
            <span class="weixin">΢��</span><a href="#" class="greengz" onClick="ShowLightBox('light','fade');"></a><span class="line"></span>
            <span class="weibo">΢��</span><a href="http://d.weibo.com/" rel="nofollow" target="_blank" class="redgz"></a><span class="line"></span>
            
            <span>{$MainInfo title='������˾��ӭ��' showtitle='������˾��ӭ��' default='��ӭ���ʣ�΢ս�Թ���' autoadd='true'$}</span>
        </div>
        <div class="right">
<sPAn class="testhidde">{$MainInfo title='��վ�����ı�TopBar' showtitle='��վ�����ı�TopBar' default='[_��վ�����ı�TopBar2015��01��26�� 09ʱ07��]' autoadd='true'$}
<!--#test start#-->
<!--#[_��վ�����ı�TopBar2015��01��26�� 09ʱ07��] start#-->
{$HrefA content='�����ղ�' title='�����ղ�' type='�ղ�' $} |
{$HrefA content='��Ϊ��ҳ' title='��Ϊ��ҳ' type='��Ϊ��ҳ' $} | 
{$HrefA href='http://d.weibo.com/' title='����΢��' rel='nofollow' class='weiboico' target='_blank'$} |
{$HrefA href='http://t.qq.com/' title='��Ѷ΢��' rel='nofollow' class='tqqico' target='_blank'$}
<!--#[_��վ�����ı�TopBar2015��01��26�� 09ʱ07��] end#-->
<!--#test end#--> 
</sPAn>  
<!--#test start#-->
<a href="#">�����ղ�</a> | 
<a href="#">��Ϊ��ҳ</a> | 
<a href="#">��վ��ͼ</a> | 
<a href="#" class="weiboico">����΢��</a> | 
<a href="#" class="tqqico">��Ѷ΢��</a>
<!--#test end#-->         
        </div>
    </div>
</div>
</text>


</dIv>
---------------------
��Top��
<dIv>



<style>
.top3 {
    height: 150px;
    width: 980px;
    margin: auto;
    vertical-align: top;
    font-size: 20px;
    font-family:"΢���ź�", "����", "����";
}
.top31{
    font-size: 20px;
    width: 200px;
    height: auto;
    float: right;
    border-top-color: #003399;
    border-right-color: #003399;
    border-bottom-color: #003399;
    border-left-color: #003399;
    font-family:"΢���ź�", "����", "����";
    line-height: 30px;
    margin-top: 80px;
    background-image: url(top1.png);
    background-repeat: no-repeat;
    text-align: center;
    background-position: 20px -5px;
}
.top31 div{
    font-family:Verdana, Arial, Helvetica, sans-serif;
    font-size:18px;
    font-weight:bold;
    line-height: 30px;
    color: #606060;
}
.top32 {
    float: left;
    width: 600px;
    color: #4c4c4c;
    margin-top: 30px;
}
.top1right {
    float: right;
    width: 230px;
}
.top31a {
    font-family:"΢���ź�", "����", "����";
    font-size: 24px;
    color:#c40000;
    font-weight: bold;
}
.top321 {
    float: left;
    height: 120px;
    width: 150px;
}
.top322 {
    padding-top: 40px;
    line-height: 35px;
}
.p4 {
    font-family:"΢���ź�", "����", "����";
}
</style>
<body>
<text>
<sPAn class="testspan">{$copyTemplateMaterial dir='ģ�鹦���б�/�����Զ���/����/' isHandle='true'$}</sPAn>


<div class="top3">
    <div class="top32">
        <div class="top321"><img src="jy_03.png" width="142" height="104" /></div>
         <div class="top322">
         
<sPAn class="testspan">{$MainInfo title='������˾����' showtitle='������˾����' default='[_������˾����2014��12��05�� 11ʱ24��]' autoadd='true'$}</sPAn>
<!--#test start#-->
<!--#[_������˾����2014��12��05�� 11ʱ24��] start#-->
            <p class="p4">�����Զ���</p>
            <p>���ʽྻ�յ�����ϵͳ���ɵ�һƷ��</p>
<!--#[_������˾����2014��12��05�� 11ʱ24��] end#-->
<!--#test end#-->
           
         </div>
    </div>
    <div class="top31">
<sPAn class="testhidde">{$MainInfo title='������������' showtitle='������������' default='[_������������2014��12��05�� 11ʱ26��]' autoadd='true'$}</sPAn>
<!--#test start#-->
<!--#[_������������2014��12��05�� 11ʱ26��] start#-->
      <p>��������:</p>
       <div>400-880-5200</div>
<!--#[_������������2014��12��05�� 11ʱ26��] end#-->
<!--#test end#-->

    </div>
    <div class="clear"></div>
</div>
<div class="clear"></div>
</text>





</dIv>
---------------------
��Nav��
<dIv>


<style>
body, div, p,img,dl, dt, dd, ul, ol, li, h1, h2, h3, h4, h5, h6, pre, form, fieldset, input, textarea, blockquote { padding:0px; margin:0px; }
li { list-style-type:none; }
.navbox{width:980px;  margin:0 auto 0 auto}
.clear { clear:both; }
/*����*/
#nav501{
background-color:#00386D;
}
.nav501{
    width:100%;
    background-color: #00386D;
    height:44px;
    line-height:44px;
    font-weight:bold;
}
.nav501 li{float:left;position:relative;}
.nav501 .left{
    width:0px;
    overflow: hidden;
}
.nav501 .line{
    width:0px;
    overflow: hidden;
}
.nav501 .right{
    width:0px;
    overflow: hidden;
    float:right;
}
.nav501 a{
    font-size:14px;
    color:#FFFFFF;
    text-decoration:none;
    height:44px;
    margin:0 1px 0 0;
    padding:0 6px;
    line-height:44px;
    display:block;
    text-align:center;
}
.nav501 a:hover{
    background-color: #FFC509;
    text-decoration:none;
    color:#00386D;
}
.nav501 .focus{
    font-size:14px;
    text-align:center;
    text-decoration:none;
    color:#00386D;
    height:44px;
    margin:0 1px 0 0;
    padding:0 6px;
    line-height:44px;
    background-color: #FFC509;
}
.nav501 .focus a{
    font-size:14px;
    text-align:center;
    text-decoration:none;
    color:#00386D;
    height:44px;
    padding:0px;
    margin:0px;
}
.nav501 li ul {display:none;
background-color:#00386D;
margin-top:44px;}
.nav501 li:hover ul {display: block; position: absolute; top:0px;left:0;} 
.nav501 li ul li {border-top:1px solid #FFFFFF;
}
.nav501 li:hover ul li a {display:block; width:110px; text-align:center;}
.nav501 li:hover ul li a:hover {}
/*End*/

</style>
<body>
<text>
<div id="nav501">
    <div class="navbox">    
        <ul class="nav501">
            <li class=left></li>

<!--#�Զ��嵼���б�#-->
<sPAn class="testspan">{$ColumnList topnumb='6' default='[_2016��02��23�� 17ʱ15��]'$}</sPAn>
 
<!--#[_2016��02��23�� 17ʱ15��]
[list]<li><a href='[$url$]'>[$columnname$]</a></li>
            <li class=line></li>
[/list]
[list-focus]<li class=focus><a href='[$url$]'>[$columnname$]</a></li>
            <li class=line></li>
[/list-focus]
#-->
 

        </ul>
        <div class="clear"></div> 
    </div>
</div>
</text>
</dIv>
---------------------
��Banner��
<dIv>
<style>
.banner6{
    height:439px;overflow:hidden; position:relative;width: 99%;
}
#bfocus7 {
    width:100%;
    height:100%;
    overflow:hidden;
    position:relative;
}
#bfocus7 ul {
    height:100%;
    position:absolute;
}
#bfocus7 ul li {
    width:100%;
    height:100%;
    float:left;
    height:439px;
} 
#bfocus7 ul li a{
    height:439px;
    display:block;
}
#bfocus7 ul li div {
    position:absolute;
    overflow:hidden;
}
#bfocus7 .btn {
    position:absolute;
    width:500px;
    height:14px;
    right:50%;
    bottom:8%;
    text-align:right;
}
#bfocus7 .btn span {
    display:inline-block;
    _display:inline;
    _zoom:1;
    width:14px;
    height:14px;
    _font-size:0;
    margin-left:5px;
    cursor:pointer;
    background:url(i-19.png) no-repeat;
}
#bfocus7 .btn .spanon {
    background:url(i-18.png) no-repeat;
}
#bfocus7 .preNext {
    width:45px;
    height:100px;
    position:absolute;
    top:40%;
    background:url(sprite.png) no-repeat 0 0;
    cursor:pointer;
}
#bfocus7 .pre {
    left:10%;
}
#bfocus7 .next {
    right:10%;
    background-position:right top;
}
</style>
<body>
<text>
<sPAn class="testspan">{$copyTemplateMaterial dir='ģ�鹦���б�/Banner/106/' isHandle='true'$}</sPAn>

<script type="text/javascript"  src="Banner_Style106.js"></script>
<div class="banner6">
    <div id="bfocus7">
        <ul>

<!--#ϸ���б� ASPPHPͨ��#-->
<sPAn class="testspan">{$ArticleList columnname='banner' topnumb='6' default='[_2016��02��23�� 14ʱ07��]'$}</sPAn>
<!--#[_2016��02��23�� 14ʱ07��]
[list]<li style="background:url([$bigimage$]) 50% no-repeat;"><a href="[$url$]"></a></li>
[/list]
#-->

 
        </ul>
    </div>
</div>

</text>



</dIv>
---------------------
����Ʒ������
<style>
.searchdialog{
    color:#000000;
    font-size:12px; 
    line-height:22px;
    width:980px;
    margin:10px auto;
}
.searchdialog a{
    font-size:12px;
    line-height:22px;
    color:#333333; 
    font-weight:normal;
}
.searchdialog a:hover{text-decoration:underline;color:#000000}
.searchdialog .left{
    float:left;
    width:660px;
    height:34px;
    line-height:34px;
    overflow-y:hidden;
}
.searchdialog .right{
    float:right;
}
</style>
<style>
.search{
}
.search .search_txt{
    float:left;
    border:2px solid #00386D;
    border-right:none;
    height:25px;
    line-height:25px;
    width:200px;
    color:#666666;font-size:12px;
    background-image:url(search.gif);
    background-repeat:no-repeat;
    background-position:0 2px;
    padding-left:30px;
}
.search .search_txt:hover{
    color:#000000;
}

.search .search_btn{
    float:left;
    height:29px;width:50px;color:#FFFFFF;
    border:2px solid #00386D;
    border-left:none;    
    background-color:#00386D;
    cursor:pointer;
}
.search .search_btn:hover{
    background-color:#00549F;
}
</style>
<dIv class='searchdialog'>
    <text>
    <div class="left">
        �ؼ��ʣ�&nbsp;

<!--#����ͳ��#-->
<sPAn class="testspan">{$SearchStatList topnumb='6' addsql='order by sortrank asc' default='[_2016��02��23�� 17ʱ07��]' through='true'$}</sPAn>
<!--#[_2016��02��23�� 17ʱ07��]
[list]<a href="[$WEB_VIEWURL$]?act=Search&wd=[$title$]">[$title$]</a> &nbsp;
[/list]
#-->


    </div>
    <div class="right">




 
<sPAn class="testspan">{$copyTemplateMaterial dir='ģ�鹦���б�/Search������/Ĭ�ϰ�/' isHandle='true'$}</sPAn>

<div class="search"> 
<form action='[$WEB_VIEWURL$]?act=Search' method='post' name='f'">
<input type="text" name="wd" id="wd" class="search_txt" onFocus="if(this.value=='������ؼ���')this.value='';" onBlur="if(this.value=='')this.value='������ؼ���';" value="[$glb_searchKeyWord default='������ؼ���'$]" /><input type="submit" name="btnSearch" id="btnSearch" value="����" class="search_btn" />
</form>
</div>
 



    </div>
    <div class="clear"></div>
    </text>
</dIv>
---------------------
#���̶�������
<text>
</div>
</text>
---------------------
������������ǩ��
<text>
<!--#top end#-->
</text>
---------------------
#Ĭ����ҳģ���б�# end




#Source�����б�# start

#Source�����б�# end
[����������]{���ɲ�IE��}
[��������]{#dialogbackground='Ĭ��' dialogborder='Ĭ��' modulebackground='Ĭ��' moduleborder='Ĭ��' }
[��ʽǰ׺����]{h_}
[������Ϣ����]{}
[Css�����ļ�����]{style.css}
[ASP������������]{��������}
[�����ļ�����]{}
[Ĭ��HTML����]{1��Template.Html}
[Ĭ��CSS����]{9��style.css}
[�Զ���ģ��]{}
[�Զ���CSS]{}
[�Զ���ģ�嶯��]{}
[�Զ���CSS����]{}
[txtIsCodeHtmlTemplate]{1}
[txtIsCodeCssTemplate]{1}
[txtIsDeleteRepeatCssClass]{1}
[txtCopyImageEncryption]{}
