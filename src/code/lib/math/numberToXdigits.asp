<%
    function numberToXdigits( val, digits )
        do while ( len(val) < digits )
            val = "0" & val
        loop
        numberToXdigits = val
    end function
%>