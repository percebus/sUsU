<%
    sub sessionPUT(var, val)
        if (useCookieSessions) then
            response.cookies("session")(var) = val
        else
                                session(var) = val
        end if
    end sub
%>