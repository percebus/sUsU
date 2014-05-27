<%
    function data2array( RS, SQLstring )
        hasRecords = false
        RS.source         = SQLstring
        RS.cursorType     = 0
        RS.cursorLocation = 2
        RS.lockType       = 1
        RS.open()
            'columns, records
            if not(RS.EOF) then
                hasRecords = true
                tempArray  = RS.getRows()
            end if
        RS.close()
        if (hasRecords) then
            data2array = tempArray
        else
            data2array = ""
        end if
    end function
%>