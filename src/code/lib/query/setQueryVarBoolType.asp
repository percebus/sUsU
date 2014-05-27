<%
    sub setQueryVarBoolType(byRef var, requestVar)
        tempVal = ""
        if (  len( request(requestVar) )  >  0  ) then
            tempVal = cBool( request(requestVar) )
        end if
        var = tempVal
        if ( len(var) > 0 ) then
            call addQuery( query, queryVars, queryValues, requestVar, absInt(var) )
        end if
    end sub
%>