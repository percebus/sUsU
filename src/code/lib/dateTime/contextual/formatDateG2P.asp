<%
    function formatDateG2P(theDate)
        formatDateG2P = year(theDate) & "/" & numberToXDigits( month(theDate), 2 ) & "/" & numberToXDigits( day(theDate), 2 )
    end function
%>