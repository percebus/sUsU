<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   )               ,  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" )               ,  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrSecurityLevelCode" ),  ( appFolder & "login.asp" )  )
    if (   cInt(  sessionGET( appName & "usrSecurityLevelCode" )  )   >   minDBadminSecurity   ) then
response.redirect("login.asp")
    end if

'***** CFG *************************************

    tableName      = tableAdminTables
    qVars          = "" ' array( "InActive", "MinSecurityLevelID"                       )
    qOperators     = "" ' array( "not"     , ">"                                        )
    qVals          = "" ' array( "null"    , sessionGET( appName & "usrSecurityLevel" ) ) 
    qValsT         = "" ' array( "null"    , "num"                                      )

    orderByVars    = array( "DisplayOrder", "TableName" )
    orderByVarsASC = array(     true      ,    true     )
'***********************************************
%>
<!--#include file="src/header.asp"-->

        <div id="navBreadcrums">
            <nav id="breadcrums">
                <ul>
                    <li>Tablas</li>
                </ul>
            </nav>
        </div>

        <div id="navMain">
           <nav id="main">
               <ul>
<%
    SQLstring          = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", "*", "", "", orderByVars, orderByVarsASC )
%><!-- <%=SQLstring%> --><%
    oRS.source         = SQLstring
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)
%>
                   <li><a href="adminTablesRecords.asp?table=<%=oRS.fields.item("TableName").value%>&R=<%=absInt( oRS.fields.item("R").value )%>&W=<%=absInt( oRS.fields.item("W").value )%>&RW=<%=absInt( oRS.fields.item("RW").value )%>&X=<%=absInt( oRS.fields.item("X").value )%>&URL="><%=oRS.fields.item("DisplayName").value%></a></li>
<%
            oRS.moveNext
        loop
    oRS.close()
%>
                </ul>
            </nav>
        </nav>

<!--#include file="src/footer.asp"-->