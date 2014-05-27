<%
    function formatTimeG2P(theTime)
        formatTimeG2P = numberToXdigits( hour(theTime), 2 ) & ":" & numberToXdigits( minute(theTime), 2 ) & ":" & numberToXdigits( second(theTime), 2 )
    end function
%>