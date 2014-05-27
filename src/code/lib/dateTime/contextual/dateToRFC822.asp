<%
    function dateToRFC822(val)
        dim monthNames(12)
            monthNames( 1) = "Jan"
            monthNames( 2) = "Feb"
            monthNames( 3) = "Mar"
            monthNames( 4) = "Apr"
            monthNames( 5) = "May"
            monthNames( 6) = "Jun"
            monthNames( 7) = "Jul"
            monthNames( 8) = "Aug"
            monthNames( 9) = "Sep"
            monthNames(10) = "Oct"
            monthNames(11) = "Nov"
            monthNames(12) = "Dec"

        dim weekDaysNames(7)
            weekDaysNames(1) = "Mon"
            weekDaysNames(2) = "Tue"
            weekDaysNames(3) = "Wed"
            weekDaysNames(4) = "Thu"
            weekDaysNames(5) = "Fri"
            weekDaysNames(6) = "Sat"
            weekDaysNames(7) = "Sun"

        tempVal = dateAdd(  "h", ( serverGMT * -1 ), val  )
        tempVal = weekDaysNames( datePart("w", val, 2) ) & ", " & numberToXdigits( day(val), 2 ) & " " & monthNames( month(val) ) & " " & year(val) & " " & numberToXdigits( hour(val), 2 ) & ":" & numberToXdigits( minute(val), 2 ) & ":" & numberToXdigits( second(val), 2 ) & " GMT"
        dateToRFC822 = tempVal
    end function
%>