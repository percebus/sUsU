<%
    sub addQuery(byRef query, byRef queryVars, byRef queryValues, var, val)
        queryElements = uBound(queryVars) + 1
        reDim preserve queryVars(  queryElements)
        reDim preserve queryValues(queryElements)
                       queryVars(  queryElements) = var
                       queryValues(queryElements) = val

        query = query & "&" & var & "=" & val
    end sub
%>