<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )
    view = "user"
%>
<!--#include file="src/header.asp"-->

        BienVenido! <%=sessionGET( appName & "usr" )%>, elige una opci&oacute;n:
        <ul>
<%
    if (   cInt(  sessionGET( appName & "usrSecurityLevelCode" )  )   <=   minDBadminSecurity   ) then
%>
            <li><a href="admin.asp">** ADMINISTRADOR **</a></li>
<%
    end if
%>
            <li><a href="setPassword.asp">Cambiar ContraSe&ntilde;a</a></li>
<!--
            <li><a href="setEMail.asp"   >Cambiar eMail</a></li>
-->
            <li><a href="personas.asp"   >Administrar Identidades</a></li>
            <li><a href="search.asp"     >Comprar</a></li>
<!--
            <li><a href="buy.asp"        >Ofertar en una subasta activa</a></li>
-->
            <li><a href="bids.asp"       >Administrar mis subastas</a></li>
<!--
            <li><a href="finance.asp"    >Administrar mi informaci&oacute;n Bancaria</a></li>
            <li><a href="summary.asp"    >Ver un listado de mis movimientos</a></li>
-->
        </ul>

<!--#include file="src/footer.asp"-->