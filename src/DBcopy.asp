<!--#include file="cfg/cfg.asp" -->
<%
    dim fs,f
    set fs = server.createObject("scripting.fileSystemObject")
%><%=DB%><%
    set f  = fs.GetFile( DB )
   call f.copy( DBcopyToFolder & "/sUsUx.mdb", false )
    set f  = nothing
    set fs = nothing
%>