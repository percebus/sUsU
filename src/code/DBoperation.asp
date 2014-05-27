<%
    operation = ""
    if     (   ( DBoID > 0 )   and   not(  len( request("RW") )  >  0  )   ) then
        operation = "SELECT"
    elseif (   ( DBoID > 0 )   and    doEditFields  )                        then
        operation = "UPDATE"
    else
        operation = "INSERT"
    end if
%>