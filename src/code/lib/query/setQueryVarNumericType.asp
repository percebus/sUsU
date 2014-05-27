<%
    sub setQueryVarNumericType(byRef var, requestVar)
        tempVal = ""
        if (  len( request(requestVar) )  >  0  ) then
            tempVal = cInt( request(requestVar) )
        end if
        var = tempVal
        if ( len(var) > 0 ) then
            call addQuery( query, queryVars, queryValues, requestVar, var )
        end if
    end sub
%>