<%
    test1 = array(1, 2, 3)
    test2 = array("text", test1, 3)

    for i = lBound(test2) to uBound(test2)
        if   isNumeric( test2(i) ) then
%>
        <%= test2(i) * 3 %><br />
<%
        elseif isArray( test2(i) ) then
            currentTest2 = test2(i)
            for j = lBound( test2(i) ) to uBound( test2(i) )
%>
        <%=currentTest2(j)%>, 
<%
            next
%>
        <br />
<%
        else
%>
       <%=test2(i)%><br />
<%
        end if
    next
%>
*******************************************************************
<form method="get" target="testResult">
    <input type="checkbox" name="vehicle" value="Bike" /> I have a bike<br />
    <input type="checkbox" name="vehicle" value="Car" /> I have a car<br />
    <input type="checkbox" name="vehicle" value="Airplane" />I have an airplane<br />
    <input type="submit" value=">>>" />
</form>
******************************************************************
<br />
<%
    dim test3(0,3)
        test3(0,0) = "juan"
        test3(0,1) = "guerrero"
        test3(0,2) = "cortazar"
        test3(0,3) = "ULSA"

    for i = lBound(test3) to uBound(test3)
        for j = lBound(test3,2) to uBound(test3,2)
%>
        <%=test3(i,j)%>,
<%
        next
%>
        <br />
<%
    next
%>