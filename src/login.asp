<!--#include file = "inc.asp" -->
<%
'***** CFG **************************************************
    tableName  = "Q_Accounts"

    usr        = lCase( request.form("usr") )
    pwd        = request.form("pwd")

    qVars      = array( "PassWord", "Code" )
    qVals      = array(  pwd      ,  usr   )
    qValsT     = array( "text"    , "text" )

    filterVars = usrVars
'************************************************************

    if (   inStr(  lCase(referer),  lCase( appFolder & "wml" )  )   >   0   ) then
        view = "WML"
    else
        view = "HTML"
    end if
    if (   not( session("tries") < tries )   or   not(  absInt( request.cookies("session")("tries") )  <  tries  )   ) then
response.redirect( "../login.asp" )
    end if

    if ( len(usr) > 0 ) then
        userExists = false
            oRS.source           = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", filterVars, "", "", "", "" )
            oRS.cursorType       = 0
            oRS.cursorLocation   = 2
            oRS.lockType         = 1
            oRS.open()
                if not(oRS.EOF) then
call sessionPUT( appName & "usr"  , usr )
                    for i = lBound(filterVars) to uBound(filterVars)
call sessionPUT( appName & "usr" & filterVars(i), oRS.fields.item( filterVars(i) ).value )
                    next
                    userExists = true
                end if
            oRS.close()
        set oRS = nothing
        if (userExists) then
response.redirect( "../personas.asp" )
        end if

    'If failed, tries += 1
        if (  session("tries")  <  absInt( request.cookies("session")("tries") )  ) then
                                session("tries") = absInt( request.cookies("session")("tries") )
        end if
        if (  session("tries")  >  absInt( request.cookies("session")("tries") )  ) then
            response.cookies("session")("tries") =                            session("tries")
        end if
                                session("tries") =                            session("tries")   + 1
            response.cookies("session")("tries") = absInt( request.cookies("session")("tries") ) + 1

        if (   not( session("tries") < tries )   or   not(  absInt( request.cookies("session")("tries") )  <  tries  )   ) then
                                session("tries") = tries
            response.cookies("session")("tries") = tries
            response.cookies("session").expires  = cDate( now + 1 / daysBlocked )
        end if
    end if
response.redirect( "../login.asp" )
%>