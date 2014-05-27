<!--#include file="inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )

'***** CFG *****************************************************
    set oSQLexe = server.createObject("ADODB.command"  )
        oSQLexe.activeConnection = DBeditConnection

    tableName = "Accounts"

    qVars  = array(  "Code"                        )
    qVals  = array(  sessionGET( appName & "usr" ) )
    qValsT = array( "text"                         )

    UDvars  = array( "PassWord"     )
    UDvals  = array( request("pwd") )
    UDvalsT = array( "text"         )
'***************************************************************
    oSQLexe.commandText = buildSQLstring( "UPDATE", DBMS, tableName, qVars, qVals, qValsT, false, UDvars, UDvals, UDvalsT, "", "", "", "", "" )
    oSQLexe.execute

    set oSQLexe = nothing
response.redirect("../user.asp")
%>