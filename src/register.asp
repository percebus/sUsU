<!--#include file="inc.asp"-->
<%
'***** CFG *****************************************************
    set oSQLexe = server.createObject("ADODB.command"  )
        oSQLexe.activeConnection = DBeditConnection

    tableName = "Registrants"

    code    = "H" & hour(time) & "U" & left( request.form("userName"), 2 ) & right( "Y" & year(date), 2  )  & "N" & right( request.form("nick"), 2 ) & "M" & month(date) & "A" & left( request.form("alias"), 2 ) & "D" & day(date) & "R" & absInt( rnd * 100 ) & "m" & minute(time)
    URLcode = "HTTP://" & request.ServerVariables("HTTP_HOST") & appFolder & "src/activate.asp?code=" & code

    eMailSubject = "sUsU: WelCome!"
    eMailMSG     = "<a href=""" & URLcode & """ target=""_blank"">" & URLcode & "</a>"

    insertVars        = array( "Code", "UserName"               , "NickName"          , "Alias"              , "eMail"               )
    insertValues      = array(  code , request.form("userName") , request.form("nick"), request.form("alias"), request.form("eMail") )
    insertValuesTypes = array( "text", "text"                   , "text"              , "text"               , "text"                )
'***************************************************************

    call sessionPUT( appName & "usr", "" )

    oSQLexe.commandText = buildSQLstring( "INSERT", DBMS, tableName, "", "", "", true, insertVars, insertValues, insertValuesTypes, "", "", "", "", "" )
    oSQLexe.execute

    sendTo = array( "", request("eMail") )
    if (   eMail( eMailerCarrier, eMailerSMTP, eMailerPort, eMailerUSR, eMailerPWD, eMailerFromEMail, eMailerFromName, sendTo, , eMailerAdmins, ";", true, eMailSubject, eMailMSG, "" )   ) then
response.redirect("../wait.asp")
    else
response.redirect("../error.asp")
    end if
%>