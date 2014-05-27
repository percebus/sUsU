<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )
    view = "setPassword"
%>
<!--#include file="src/header.asp"-->

        <table border="1">
            <form id="setPasswordForm" name="setPasswordForm" action="src/setPassword.asp" method="post" onSubmit="return form_validate_required(this, 'los siguientes campos estan vacíos:\n', fieldRequired, fieldDescription);">
                <tr><td>ContraSeña</td><td><input id="pwd" name="pwd" type="password" /></td></tr>
                <tr><td>Confirmar ContraSeña</td><td><input id="pwd2" name="pwd2" type="password" /></td></tr>
                <tr><td colspan="2"><input type="submit" value=" >>> " style="width:100%" /></td></tr>
	        </form>
        </table>

<!--#include file="src/footer.asp"-->