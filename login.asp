<!--#include file="src/inc.asp"-->
<%
    view      = "login"
    usedTries = setDefault( sessionGET("tries")+1, 1 )
%>
<!--#include file="src/header.asp"-->
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td align="center" valign="middle" height="100%">
<%
        if (   ( session("tries") >= tries )   or   (  absInt( request.cookies("session")("tries") )  >=  tries  )   ) then
%>
                    <table border="0" align="center" class="alert">
                        <tr>
                            <td height="100">Demasiados intentos, intente de nuevo m&aacute;s tarde</td>
                        </tr>
                    </table>
<%
        else
%>
                    <form name="form" method="post" action="src/login.asp" target="_self">
                        <input id="referer" name="referer" type="hidden" value="<%=referer%>" maxlength="20" />
                        <input id="folder"  name="folder"  type="hidden" value="<%=folder%>"  maxlength="20" />
                        <input id="method"  name="method"  type="hidden" value="<%=method%>"  maxlength="20" />
                        <table width="500" align="center" class="box">
                            <thead>
                                <tr>
                                    <td colspan="2" align="center"></td>
                                </tr>
                            </thead>
                            <tr>
                                <td align="right">Usuario</td>
                                <td><input id="usr" name="usr" type="text" maxlength="20" /></td>
                            </tr>
                            <tr>
                                <td align="right">Contrase&ntilde;a</td>
                                <td><input id="pwd"  name="pwd"    type="password" maxlength="20" /></td>
                            </tr>
                            <tr>
                                <td align="right"<%if ( usedTries > 1 ) then%> class="alert"<%end if%>><%=usedTries%> / <%=tries%></td>
                                <td align="left"><input id="logIn" name="logIn" type="submit" value="Log In" /></td>
                            </tr>
                        </table>
                    </form>
<%
        end if
%>
                </td>
            </tr>
        </table>
        No tienes cuenta? <a href="register.asp">Reg&iacute;strate!</a>
<!--#include file="src/footer.asp"-->