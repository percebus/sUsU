<%@ language = "VBscript" %>
<%
    response.buffer = false
%>
<!--#include file = "cfg/dim.asp"  -->
<!--#include file = "code/lib.asp" -->
<!--#include file = "cfg/cfg.asp"  -->
<%
    isMobileBrowser = false
    isWAPbrowser    = false
    if (   inStr(  lCase( request.serverVariables("HTTP_ACCEPT") ),  lCase("WAP")  )   >   0   ) then
        isMobileBrowser = true
        isWAPbrowser    = true
    elseif (   inStr(  lCase( request.serverVariables("ALL_HTTP") ),  lCase("operamini")  )   >   0   ) then
        isMobileBrowser = true
    elseif (   iInArray(  left( request.serverVariables("HTTP_USER_AGENT"), 4 ),  mobileBrowsers  )   >   0   ) then
        isMobileBrowser = true
    end if
%>
<!--#include file = "cfg/DB.asp"      -->
<!--#include file = "cfg/var.asp"     -->
<!--#include file = "cfg/session.asp" -->