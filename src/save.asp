<!--#include file = "inc.asp"  -->
<%
'***** CFG ******************************
    response.buffer = true
    statsDBoName    = "Stats"
'****************************************

    gotoURL = setDefault( request("URL"), appFolder & "noURL.asp" )
%><%=gotoURL%><br /><%
        qVars  = split( request("criterias")     , "," )
        qValsT = split( request("criteriasTypes"), "," )

    dim qVals()
  reDim qVals( uBound(qVars) )
    for iVar = lBound(qVars) to uBound(qVars)
        qVals(iVar) = request( qVars(iVar) )
    next

        eVars  = split( request("params")              , "," )
        eValsT = split( request("paramsTypes")         , "," )
        eValsO = split( request("paramsOriginalValues"), "," )

    hasValues = false
    dim eVals()
  reDim eVals( uBound(eVars) )
    for iVar = lBound(eVars) to uBound(eVars)
        eVals(iVar) = request( eVars(iVar) )
        if (  len( eVals(iVar) )  >  0  ) then
            hasValues = true
        end if
    next

    hasChanges = false
    if (  len( request("paramsOriginalValues") )  >  0  ) then
        eFVarsS  = ""
        eFValsS  = ""
        eFValsTS = ""
        for iVar = lBound(eVars) to uBound(eVars)
            if not( eVars(iVar) = "ID" ) then
                if not( eVals(iVar) = eValsO(iVar) ) then
                    hasChanges = true
                    if ( len(eFVarsS) > 0 ) then
                        eFVarsS  = eFVarsS  & ","
                        eFValsS  = eFValsS  & ","
                        eFValsTS = eFValsTS & ","
                    end if
                    eFVarsS  = eFVarsS  & eVars( iVar)
                    eFValsS  = eFValsS  & eVals( iVar)
                    eFValsTS = eFValsTS & eValsT(iVar)
                end if
            end if
        next
        eFVars  = split( eFVarsS , "," )
        eFVals  = split( eFValsS , "," )
        eFValsT = split( eFValsTS, "," )
    else
        if (hasValues) then
            hasChanges = true
            eFVars     = eVars
            eFVals     = eVals
            eFValsT    = eValsT
        end if
    end if

%><%=request("criterias")%><br /><%
%><%=request("criteriasTypes")%><br /><%

%><%=request("params")%><br /><%
%><%=request("paramsTypes")%><br /><%
%><%=request("paramsOriginalValues")%><br /><%
%><%=eFVarsS%><br /><%
%><%=eFValsS%><br /><%
%><%=eFValsTS%><br /><%

    if (hasChanges) then
        set oSQLexe = server.createObject("ADODB.command"  )
            oSQLexe.activeConnection = DBeditConnection
            select case uCase( request("operation") )
                case "INSERT"
                    SQLstring = buildSQLstring( request("operation"), DBMS, request("DBo"), ""   , ""   , ""    , false, eFVars, eFVals, eFValsT, "", "", "", "", "" )
                case "UPDATE"
                    SQLstring = buildSQLstring( request("operation"), DBMS, request("DBo"), qVars, qVals, qValsT, false, eFVars, eFVals, eFValsT, "", "", "", "", "" )
            end select
%><%=SQLstring%><%
            oSQLexe.commandText = SQLstring
            oSQLexe.execute()

            if (  cBool( request("stat") )  ) then

                on error resume next

                sVars  = array(               "UserAccountID"  ,                       "PersonaID"  ,         "DBo" ,         "Operation" ,                  "Criteria"              ,                 "Params"                 ,                         "RequestURL"   ,                         "RequestMethod"  ,                         "IPsource"    ,                          "UserBrowser"    ,                          "UserCPU"    ,                          "UserPixels"    ,                         "UserOS"      ,                          "AllData" )',                         "AllRaw"   )
                sVals  = array( sessionGET( appName & "usrID" ), sessionGET( appName & "personaID" ), request("DBo"), request("operation"), replace( request("criterias"), ",", "|" ), replace(  join( eFVars, "," ), ",", "|" ), request.serverVariables("HTTP_REFERER"), request.ServerVariables("REQUEST_METHOD"), request.serverVariables("REMOTE_ADDR"), request.serverVariables("HTTP_USER_AGENT"), request.serverVariables("HTTP_UA_CPU"), request.serverVariables("HTTP_UA_PIXELS"), request.serverVariables("HTTP_UA_CPU"), request.serverVariables("ALL_HTTP"))', request.serverVariables("ALL_RAW") )
                sValsT = array(               "num"            ,                       "num"        ,         "text",         "text"      ,                  "text"                  ,                  "text"                  ,                         "text"         ,                         "text"           ,                         "text"        ,                          "text"           ,                          "text"       ,                          "text"          ,                         "text"        ,                          "text"    )',                         "text"     )

                oSQLexe.commandText = buildSQLstring( "INSERT", DBMS, statsDBoName, "", "", "", false, sVars, sVals, sValsT, "", "", "", "", "" )
                oSQLexe.execute()

            end if

        set oSQLexe = nothing
    end if

    for iURLcode = lBound(URLencodes) to uBound(URLencodes)
        gotoURL = replace( gotoURL, URLencodes(iURLcode), URLdecodes(iURLcode) )
    next
%><%=gotoURL%><%
response.redirect(gotoURL)
%>