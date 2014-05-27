<!--#include file="src/inc.asp"-->
<%
'    call varIsRequired(  sessionGET( appName & "usr"   )               ,  ( appFolder & "login.asp" )  )
'    call varIsRequired(  sessionGET( appName & "usrID" )               ,  ( appFolder & "login.asp" )  )
'    call varIsRequired(  sessionGET( appName & "usrSecurityLevelCode" ),  ( appFolder & "login.asp" )  )
'    if (   cInt(  sessionGET( appName & "usrSecurityLevelCode" )  )   >   minDBadminSecurity   ) then
'response.redirect("login.asp")
'    end if
    call varIsRequired( request("table"),  ( appFolder & "adminTables.asp" )  )

'***** CFG *************************************
    tableName         = request("table")
    DefaultNamesTable = "DefaultNames"
    R                 = absInt( request("R")  )
    W                 = absInt( request("W")  )
    RW                = absInt( request("RW") )
    X                 = absInt( request("X")  )
    URLwhenDone       = setDefault( request("URL"), appFolder & "adminTablesRecords.asp?table=" & tableName & "&R=" & R & "&W=" & W & "&RW=" & RW & "&X=" & X & "&doRW=0" & "&URL=" )
    for iURLcode = lBound(URLdecodes) to uBound(URLdecodes)
        URLwhenDone = replace( URLwhenDone, URLdecodes(iURLcode), URLencodes(iURLcode) )
    next
'***********************************************

    set oDB2SQL = new DB2SQL
        oDB2SQL.DBo           = DefaultNamesTable
        oDB2SQL.fVars         = array( "TableName", "TableRename" )
        oDB2SQL.qVars         = array( "TableName" )
        oDB2SQL.qVals         = array(  tableName  )
        oDB2SQL.qValsT        = array( "text"      )
        oDB2SQL.orderBy       = array( "TableName" )
        oDB2SQL.orderByASC    = array(  true       )
            defaultTableNames = data2array( oRS, oDB2SQL.SQLstring )
            tableRename = ""
            if ( isArray(defaultTableNames) ) then
                for iTableName = lBound(defaultTableNames,2) to uBound(defaultTableNames,2)
                    if ( defaultTableNames(0,iTableName) = tableName ) then
                        tableRename = defaultTableNames(1,iTableName)
                    end if
                next
            end if
            if not( len(TableRename) > 0 ) then
                tableRename = tablePickPrefix & tableName
            end if
    set oDB2SQL = nothing
%>

<!--#include file="src/header.asp"-->

        <div id="navBreadcrums">
            <nav id="breadcrums">
                <ul>
                    <li><a href="adminTables.asp">Tablas</a></li>
                    <li><%=tableName%></li>
                </ul>
            </nav>
        </div>

<%
    oRS.source         = buildSQLstring( "SELECT", DBMS, tableRename, "", "", "", false, "", "", "", "*", "", "", orderByVars, orderByVarsASC )
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
%>
        <table border="1" cellpadding="10" cellspacing="10">
            <thead>
<%
        if ( cBool(W) ) then
%>
                <tr>
                    <th colspan="3" align="center">
                        <a href="adminTablesRecordsSet.asp?table=<%=tableName%>&R=<%=R%>&W=<%=W%>&RW=<%=RW%>&X=<%=X%>&doRW=1&URL=<%=URLwhenDone%>">Agregar Registro</a>
                    </th>
                </tr>
<%
        end if
%>
                <tr>
<%
        for each currentField in oRS.fields
%>
                    <th><%=currentField.name%></th>
<%
        next
%>
                </tr>
            </thead>
<%
        if not(oRS.EOF) then
%>
            <tbody>
<%
            do while not(oRS.EOF)
%>
                <tr>
<%
                for each currentField in oRS.fields
%>
                    <td nowrap="true">
<%
                    if ( cBool(R) ) then
%>
                        <a href="adminTablesRecordsSet.asp?table=<%=tableName%>&ID=<%=oRS.fields.item("ID").value%>&R=<%=R%>&W=<%=W%>&RW=<%=RW%>&X=<%=X%>&doRW=0&URL=<%=URLwhenDone%>">
<%
                    end if
%>
                            <%=currentField.value%>
                        </a>
                    </td>
<%
                next
                if ( cBool(X) ) then
%>
                    <td><a href="<%=currentDeleteQuery%>">X</a></td>
<%
                end if
%>
                </tr>
<%
                oRS.moveNext
            loop
%>
            </tbody>
<%
        end if
%>
        </table>
<%
    oRS.close()
%>

<!--#include file="src/footer.asp"-->