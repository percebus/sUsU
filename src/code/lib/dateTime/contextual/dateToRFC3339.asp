<%
    function dateToRFC3339(val)
        tempVal = dateAdd(  "h", ( serverGMT * -1 ), val  )
        tempVal = year(val) & "-" & numberToXdigits( month(val), 2 ) & "-" & numberToXdigits( day(val), 2 ) & "T" & numberToXdigits( hour(val), 2 )  & ":" & numberToXdigits( minute(val), 2 ) & ":" & numberToXdigits( second(val), 2 )  & "Z"
        dateToRFC3339 = tempVal
    end function
%>