<%
    function absInt(x)
        tempVal = ""
        if ( len(x) > 0 ) then
            tempVal = abs( cInt(x) )
        end if
        absInt = tempVal
    end function
%>