<%
    function sessionGET(var)
        if (useCookieSessions) then
            sessionGET = request.cookies("session")(var)
        else
            sessionGET =                    session(var)
        end if
    end function
%>