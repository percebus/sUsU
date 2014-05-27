<!DOCTYPE html>
<html lang="es">
    <head>
<!--#include file="head.asp"-->

<!-- jQuery -->
<script type="text/javascript" src="http://jquery.com/src/jquery-latest.js"></script>
<!-- <script type="text/javascript" src="src/inc/scripts/client/jQuery.js"></script> -->


<!-- datePicker -->
<script type="text/javascript" src="http://jqueryui.com/ui/jquery.ui.core.js"      ></script>
<script type="text/javascript" src="http://jqueryui.com/ui/jquery.ui.widget.js"    ></script>
<script type="text/javascript" src="http://jqueryui.com/ui/jquery.ui.datepicker.js"></script>
<!-- <link type="text/css" rel="stylesheet" media="all" href="src/cfg/CSS/jQuery/ui-lightness/jquery.ui.all.css" /> -->
<link type="text/css" rel="stylesheet" media="all" href="src/cfg/CSS/jQuery/ui-lightness/jquery-ui-1.8.2.custom.css"/>


<!-- WYSIWYG -->
<!--<script type="text/javascript" src="http://ckeditor.com/sites/default/files/js/CKEditor/js_d6fe38b679adcc34d4bfed75359b6e7a.js"></script>-->
<script type="text/javascript" src="http://ckeditor.com/apps/ckeditor/3.3.1/ckeditor.js?1277707288"></script>
<script type="text/javascript" src="http://ckeditor.com/apps/ckeditor/3.3.1/_source/lang/_languages.js?1277707288"></script>
        <script type="text/javascript">
        <!--//--><![CDATA[//><!--
            jQuery.extend(Drupal.settings, { "basePath": "/", "ckeditorPath": "apps/ckeditor/3.3.1", "ckfinderPath": "apps/ckfinder/2.0" });
        //--><!]]>
        </script>
<script type="text/javascript" src="http://a.cksource.com/c/1/misc/jquery.tooltip.pack.js"></script>

        <script type="text/javascript">
        <!--
            var $jQuery = jQuery.noConflict();
                $jQuery(document).ready(function()
                                        {
                                            $jQuery("a, input").mouseover(function()
                                                                          {
                                                                              $jQuery(this).fadeTo("fast",<%=imgAlphaFinal/100%>,null);
                                                                              $jQuery(this).css("background-color","<%=highLightColor%>");
                                                                          });

                                            $jQuery("a, input").mouseout(function()
                                                                         {
                                                                             $jQuery(this).fadeTo("fast",<%=imgAlphaInitial/100%>,null);
                                                                             $jQuery(this).css("background-color","");
                                                                         });
                                            $jQuery(".datePicker").datepicker(
                                                                              {
                                                                                  showButtonPanel:   true,

                                                                                  showOtherMonths:   true,
                                                                                  selectOtherMonths: true,

                                                                                  showWeek:          true,
                                                                                  firstDay:          1,

                                                                                  showOn:            'button',
                                                                                  buttonImage:       'src/cfg/img/calendar.gif'
                                                                              });
//s                                          $jQuery$(".WYSIWYG").wysiwyg();
                                        });
        -->
        </script>

<link id="feed" rel="alternate" type="application/rss+xml" title="sUsU-feed" href="http://<%=request.ServerVariables("HTTP_HOST")%><%=appFolder%>search.asp?format=feed"   />

    </head>
    <body background="src/cfg/img/BG.gif">
        <div id="wrapper">
            <div id="header" class="rounded">
<!--
                <header>
-->
                    <h1>
                        <a href="./">
                            SuSubasta.com
                            <img src="src/cfg/img/home.png" alt="Home" width="25"  />
                        </a>
                    </h1>
                    <h2 style="display:block; width:auto; float:right">
<%
    garbage = sessionState()
%>
                        <br /><br />
                        <hr />
                        <form action="search.asp" method="get">
                            <input type="text" name="q" id="textfield" value="Buscar" onFocus="this.value=''" />
                            <input type="submit" value=" s&Ugrave;s&Uacute; " /> 
<!--
                        | Contáctanos | Mapa del sitio |
-->
                        </form>
                    </h2>
<!--
                </header>
-->
            </div>

            <div id="nav">
<!--
                <nav>
-->
                    <div id="navCategories" class="menu">
                        <h3>Categor&iacute;as</h3>
                        <ul>
<%
    categoriesSQLstring = buildSQLstring( "SELECT", DBMS, "ItemCategories", "", "", "", false, "", "", "", "*", "", 5, "Name", true )
%><!-- <%=categoriesSQLstring%> --><%
    oRS.source         = categoriesSQLstring
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)
%>
                            <li><a href="search.asp?category=<%=oRS.fields.item("ID").value%>"><%=oRS.fields.item("Name").value%></a></li>
<%
            oRS.moveNext()
        loop
    oRS.close()
%>
                            <li><a href="search.asp">Ver todas</a></li>
                        </ul>
                    </div>
                    <div id="navUtilities" class="menu">
                        <h3>Ligas &uacute;tiles</h3>
                        <ul>
<%
    for i = lBound(menuItems) to uBound(menuItems)
%>
                            <li><a href="<%=menuItemsLinks(i)%>"><%=menuItems(i)%></a></li>
<%
    next
%>
                        </ul>
                    </div>
<iframe width="100%" height="200" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" src="http://www.lasallechihuahua.edu.mx/ILS/scripts/JCystems/JCoFTP/flshow.asp?folder=/ILS/site/public/JC/sUsU/pics&allowNavigation=0&format=3Dflip&BGcolor=FFD8AF"></iframe>
<!--
                </nav>
-->
            </div>

            <div id="sectionMain" class="rounded">
<!--
                <section id="main">
-->