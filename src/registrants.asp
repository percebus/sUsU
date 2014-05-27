<!--#include file="inc.asp"-->
<%
'***** CFG ************************************************************
    tableName        = "Registrants"

    qVars  = array( "Date"              )
    qVals  = array( formatDateG2P(date) )
    qValsT = array( "date"              )

    filterVars       = array("code")

    orderByVars      = array( "Date", "Time" )
    orderByVarsASC   = array(  false, false  )
'**********************************************************************

        oRS.source           = buildSQLstring( "SELECT", DBMS, tableName, "", "", "", true, "", "", "", filterVars, "", "", orderByVars, orderByVarsASC )
        oRS.cursorType       = 0
        oRS.cursorLocation   = 2
        oRS.lockType         = 1
        oRS.open()
            do while not(oRS.EOF)
%><%=oRS.fields.item("code").value%><br /><%
                oRS.moveNext()
            loop
        oRS.close()

    set oRS = nothing
%>