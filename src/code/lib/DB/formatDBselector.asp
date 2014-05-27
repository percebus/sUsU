<%
    function formatDBselector(var)
        tempVal = var
        if not(  left( tempVal, 1 )  =  "["  ) then
            tempVal = "[" & tempVal
        end if
        if not(  right( tempVal, 1 )  =  "]"  ) then
            tempVal = tempVal & "]"
        end if
        formatDBselector = tempVal
    end function
%>