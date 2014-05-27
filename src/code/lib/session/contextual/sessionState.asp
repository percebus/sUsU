<%
    function sessionState()
        tempVal = false
        if not (   len(  sessionGET( appName & "usr" )  )   >   0   ) then
            tempVal = true
%>
        <a href="login.asp">Iniciar Sesión | Registrarse</a>
<%
        else
%>
        <a href="user.asp"><%=sessionGET( appName & "usr" )%></a> | 
        <a href="personas.asp">
<%
            if (   len(  sessionGET( appName & "persona" )  )   >   0   ) then
%>
		    <%=sessionGET( appName & "persona" )%>
<%
            else
%>
            elegir identidad
<%
            end if
%>
        </a> | 
        <a href="src/logOut.asp">Cerrar Sesión</a>
<%
        end if
        sessionSate = tempVal
    end function
%>