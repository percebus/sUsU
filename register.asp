<!--#include file="src/inc.asp"   -->
<%
    view = "register"
%>
<!--#include file="src/header.asp"-->

        <table border="1">
            <form id="registerForm" name="registerForm" action="src/register.asp" method="post" onSubmit="return form_validate_required(this, 'los siguientes campos estan vacíos:\n', fieldRequired, fieldDescription);">
                <tr><td>Nombre de Usuario</td><td><input id="userName" name="userName" type="text" /></td></tr>
                <tr><td>Nick</td><td><input id="nick" name="nick" type="text" /></td></tr>
                <tr><td>Alias</td><td><input id="alias" name="alias" type="text" /></td></tr>
                <tr><td><input type="eMail" id="eMail" name="eMail" value="tu@eMail" size="10" onClick="this.value=''" onBlur="checkEMails(this, document.getElementById('eMail2') )" /></td><td><input type="eMail" id="eMail2" name="eMail2" value="tu@eMailNuevamente"  onClick="this.value=''" onBlur="checkEMails(this, document.getElementById('eMail') )" size="10" /></td></tr>
                <tr><td colspan="2"><input type="submit" value=" >>> " style="width:100%" /></td></tr>
	        </form>
        </table>

<!--#include file="src/footer.asp"-->