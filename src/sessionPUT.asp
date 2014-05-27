<!--#include file = "inc.asp" -->
<%
    vars = split( request("vars"), "," )
    vals = split( request("vals"), "," )
    for iVar = lBound(vars) to uBound(vars)
call sessionPUT( appName & vars(iVar), vals(iVar) )
    next
response.redirect(  setDefault( request("URL"), request.ServerVariables("HTTP_REFERER") )  )
%>