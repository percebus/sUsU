<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   )               ,  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" )               ,  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrSecurityLevelCode" ),  ( appFolder & "login.asp" )  )
    if (   cInt(  sessionGET( appName & "usrSecurityLevelCode" )  )   >   minDBadminSecurity   ) then
response.redirect("login.asp")
    end if
%>

<!--#include file="src/header.asp"-->

    <ul>
       <li><a href="adminTables.asp">AdminTables</a></li>
       <li><a href="adminTablesRecords.asp?table=Stats&R=1&W=0&RW=0&X=0&URL=">Stats</a></li>
    </ul>
    <a href="src/layers.png" target="_blank"><img width="100%" border="0" src="src/layers.png" /></a>

<!--#include file="src/footer.asp"-->