<%
    sub setQueryVarTextType(byRef var, requestVar)
        tempVal = ""
        tempVal = cStr( request(requestVar) )
        var     = tempVal
        if ( len(var) > 0 ) then
            call addQuery( query, queryVars, queryValues, requestVar, var )
        end if
    end sub
%>