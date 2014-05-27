<!--#include file="inc.asp" -->
<%
'***** CFG ******************************************************
    DBtable         = "dummy"

    queryFilter     = array( "Text" , "Date" , "Time", "Numeric", "Bool"     )
    queryFilterAs   = array( "Texto", "Fecha", "Hora", "Número" , "Booleano" )

    queryUpdates    = array( "Text"    , "Numeric", "Date"              , "Time"               , "Bool", "Binary" )
    queryUpdatesV   = array( "percebus", "123"    ,  formatDateG2P(date),  formatTimeG2P(time) ,  true , NULL     )
    queryUpdatesVT  = array( "text"    , "num"    , "date"              , "time"               , "bool", "null"   )

    queryCriteria   = "*"
    queryCriteriaV  = ""
    queryCriteriaVT = ""

    queryOrderBy    = array( "ID"    )
    queryOrderByASC = array(  false  )
'***************************************************************

    set oSQLexe = server.createObject("ADODB.command"  )
        oSQLexe.activeConnection = DBeditConnection

    selectSQL     = buildSQLstring( "SELECT", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
    insertSQL     = buildSQLstring( "INSERT", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
    updateSQL     = buildSQLstring( "UPDATE", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
    deleteSQL     = buildSQLstring( "DELETE", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
    procedureSQL  = buildSQLstring( "SP"    , DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
    procedureSQL2 = buildSQLstring( "SP"    , DBMS, DBtable, ""           , queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
%>
<html>
    <body>
        Example:
        <p>
            <table border="1">
                <tr><td>SELECT</td><td nowrap="true">buildSQLstring( "<span style="color:red">SELECT</span                    >", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
                <tr><td>INSERT</td><td nowrap="true">buildSQLstring( "<span style="color:red">INSERT</span                    >", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
                <tr><td>UPDATE</td><td nowrap="true">buildSQLstring( "<span style="color:red">UPDATE</span                    >", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
                <tr><td>DELETE</td><td nowrap="true">buildSQLstring( "<span style="color:red">DELETE</span                    >", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
                <tr><td>SP    </td><td nowrap="true">buildSQLstring( "<span style="color:red">SP&nbsp;&nbsp;&nbsp;&nbsp;</span>", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
                <tr><td>SP    </td><td nowrap="true">buildSQLstring( "<span style="color:red">SP&nbsp;&nbsp;&nbsp;&nbsp;</span>", DBMS, DBtable, queryCriteria, queryCriteriaV, queryCriteriaVT, true, queryUpdates, queryUpdatesV, queryUpdatesVT, queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )</td></tr>
            </table>
        </p>
        <hr />
        <p>
            Where:
            <table border="1">
                <tr><td>queryFilter  </td><td nowrap="true">array( "NickName", "Alias", "eMail"  )</td></tr>
                <tr><td>queryFilterAs</td><td nowrap="true">array( "apodo"   , ""     , "correo" )</td></tr>
                <tr><td colspan="2"><hr /></td></tr>

                <tr><td>queryCriteria  </td><td nowrap="true">array( "Alias"   , "eMail"                 )</td></tr>
                <tr><td>queryCriteriaV </td><td nowrap="true">array( "Percebus", "renee.karime@gMail.com")</td></tr>
                <tr><td>queryCriteriaVT</td><td nowrap="true">array( "text"    , "text"                  )</td></tr>
                <tr><td colspan="2"><hr /></td></tr>

                <tr><td>queryUpdates  </td><td nowrap="true">array( "Code", "SecurityLevelID" )</td></tr>
                <tr><td>queryUpdatesV </td><td nowrap="true">array( "JC1" , 3                 )</td></tr>
                <tr><td>queryUpdatesVT</td><td nowrap="true">array( "text", "num"             )</td></tr>
                <tr><td colspan="2"><hr /></td></tr>

                <tr><td>queryOrderBy   </td><td nowrap="true">array( "Alias" )</td></tr>
                <tr><td>queryOrderByASC</td><td nowrap="true">array(  true   )</td></tr>
                <tr><td colspan="2"><hr /></td></tr>
            </table>
        </p>
        <hr />
        <p>
            Resulting:
            <table border="1">
                <tr><td nowrap="true"><%=selectSQL%></td></tr>
                <tr><td nowrap="true"><%=insertSQL%></td></tr>
                <tr><td nowrap="true"><%=updateSQL%></td></tr>
                <tr><td nowrap="true"><%=deleteSQL%></td></tr>
                <tr><td nowrap="true"><%=procedureSQL%></td></tr>
                <tr><td nowrap="true"><%=procedure2SQL%></td></tr>
            </table>
        </p>
        <p><%=DBstring%></p>
        <p>
            Collecting...<br />
<%
    queryFilter     = array( "Text" , "Date" , "Time", "Numeric", "Bool"     )
    queryFilterAs   = array( "Texto", "Fecha", "Hora", "Número" , "Booleano" )

    queryCriteria   = "*"
    queryCriteriaV  = ""
    queryCriteriaVT = ""

    queryOrderBy    = array( "ID"    )
    queryOrderByASC = array(  false  )
 
    selectSQL       = buildSQLstring( "SELECT", DBMS, "dummy", queryCriteria, queryCriteriaV, queryCriteriaVT, true, "", "", "", queryFilter, queryFilterAs, "", queryOrderBy, queryOrderByASC )
%><%=selectSQL%><br /><%
    oRS.source         = selectSQL
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)
            for i = lBound(queryFilter) to uBound(queryFilter)
                currentField = setDefault( queryFilterAs(i), queryFilter(i) )
%>
                <%=oRS.fields.item(currentField).value%> |
<%
            next
            oRS.moveNext
%>
                <br />
<%
        loop
    oRS.close()
%>
        </p>
        <hr />
        <p>
            Inserting...<br />
<%
    queryUpdates   = array( "Text"    , "Numeric", "Date"              , "Time"               , "Bool", "Binary" )
    queryUpdatesV  = array( "percebus", "123"    ,  formatDateG2P(date),  formatTimeG2P(time) ,  true ,  NULL    )
    queryUpdatesVT = array( "text"    , "num"    , "date"              , "time"               , "bool", "null"   )

    SQLstring       = buildSQLstring( "INSERT", DBMS, "dummy", "", "", "", true, queryUpdates, queryUpdatesV, queryUpdatesVT, "", "", "", "", "" )
%><%=SQLstring%><br /><%
    oSQLexe.commandText = SQLstring
    oSQLexe.execute
%>
        </p>
        <p>
            Collecting...<br />
<%
    oRS.source         = selectSQL
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        maxNumber = setDefault( sessionGET( appName & "maxNumber" ),  0  )
        do while not(oRS.EOF)
            for i = lBound(queryFilter) to uBound(queryFilter)
                currentField = setDefault( queryFilterAs(i), queryFilter(i) )
%>
                <%=oRS.fields.item(currentField).value%> |
<%
            next
            oRS.moveNext
%>
                <br />
<%
        loop
    oRS.close()
    maxNumber = maxNumber + 1
    call sessionPUT( appName & "maxNumber", maxNumber )
%>
        </p>
        <p>
            UpDating....<br />
<%
    queryFilter     = array( "Text" , "Date" , "Time", "Numeric", "Bool"     )
    queryFilterAs   = array( "Texto", "Fecha", "Hora", "Número" , "Booleano" )

    queryUpdates    = array( "Numeric", "Date"              , "Time"               , "Bool" , "Binary" )
    queryUpdatesV   = array( maxNumber,  formatDateG2P(date),  formatTimeG2P(time) ,  false , NULL     )
    queryUpdatesVT  = array( "num"    , "date"              , "time"               , "bool" , "null"   )

    queryCriteria   = array( "Numeric" )
    queryCriteriaV  = array( "123"     )
    queryCriteriaVT = array( "num"     )

    SQLstring       = buildSQLstring( "UPDATE", DBMS, "dummy", queryCriteria, queryCriteriaV, queryCriteriaVT, false, queryUpdates, queryUpdatesV, queryUpdatesVT, "", "", "", "", "" )
%><%=SQLstring%><br /><%
    oSQLexe.commandText = SQLstring
    oSQLexe.execute
%>
        </p>
        <p>
            Collecting...<br />
<%
    oRS.source         = selectSQL
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)
            for i = lBound(queryFilter) to uBound(queryFilter)
                currentField = setDefault( queryFilterAs(i), queryFilter(i) )
%>
                <%=oRS.fields.item(currentField).value%> |
<%
            next
            oRS.moveNext
%>
                <br />
<%
        loop
    oRS.close()
%>
        </p>
        <p>
            Deleting....<br />
<%
    queryCriteria   = array( "Numeric"     )
    queryCriteriaV  = array( maxNumber - 1 )
    queryCriteriaVT = array( "num"         )

    SQLstring       = buildSQLstring( "DELETE", DBMS, "dummy", queryCriteria, queryCriteriaV, queryCriteriaVT, false, "", "", "", "", "", "", "", "" )
%><%=SQLstring%><br /><%
    oSQLexe.commandText = SQLstring
'   oSQLexe.execute
%>
        </p>


    </body>
</html>
<%
    set oRS     = nothing
    set oSQLexe = nothing
%>