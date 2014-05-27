<!--#include file="inc.asp"-->
<%
'***** CFG ************************************************************
    SELECTtableName = "Registrants"
    INSERTtableName = "Accounts"

    qVars  = array(         "Code"  )
    qVals  = array( request("code") )
    qValsT = array(         "text"  )
'**********************************************************************

    set oSQLexe = server.createObject("ADODB.command"  )
        oSQLexe.activeConnection = DBeditConnection

    accountExists = false
        oRS.source         = buildSQLstring( "SELECT", DBMS, SELECTtableName, qVars, qVals, qValsT, true, "", "", "", "*", "", "", "", "" )
        oRS.cursorType     = 0
        oRS.cursorLocation = 2
        oRS.lockType       = 1
        oRS.open()
            if not(oRS.EOF) then
                accountExists  = true
                usr            = oRS.fields.item( "UserName" ).value
                usrAlias       = oRS.fields.item( "Alias"    ).value
                usrNickname    = oRS.fields.item( "NickName" ).value
                addVars        = array( "RegistrantID"             , "Code" , "Alias" , "NickName" , "eMail"                        )
                addValues      = array( oRS.fields.item("ID").value,  usr   , usrAlias, usrNickName, oRS.fields.item("eMail").value )
                addValuesTypes = array( "num"                      , "text" , "text"  , "text"     , "text"                         )

                oSQLexe.commandText = buildSQLstring( "INSERT", DBMS, INSERTtableName, "", "", "", true, addVars, addValues, addValuesTypes, "", "", "", "", "" )
                oSQLexe.execute

call sessionPUT( appName & "usr"         , usr         )
call sessionPUT( appName & "usrAlias"    , usrAlias    )
call sessionPUT( appName & "usrNickname" , usrNickname )
            end if
        oRS.close()

    set oRS     = nothing
    set oSQLexe = nothing

    if (accountExists) then
response.redirect("../setPassWord.asp")
    else
response.redirect("../error.asp")
    end if
%>