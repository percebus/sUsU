<%
    sub setQueryVarUIDType(byRef var, requestVar)
        tempVal = -1
        if (  len( request(requestVar) )  >  0  ) then
            tempVal = cInt( request(requestVar) )
        end if
        var = tempVal
        if not( var > 0 ) then
            var = ""
        else
            call addQuery( query, queryVars, queryValues, requestVar, var )
        end if
    end sub
%>