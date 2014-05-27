<!--#include file="inc.asp"-->
<%
 call sessionPUT( appName & "usr"    , "" )
 call sessionPUT( appName & "persona", "" )

    for i = lBound(usrVars) to uBound(usrVars)
call sessionPUT( appName & "usr" & usrVars(i), "" )
    next

    for i = lBound(personaVars) to uBound(personaVars)
call sessionPUT( appName & "persona" & personaVars(i), "" )
    next

response.redirect("../")
%>