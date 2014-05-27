<%
    if (   len(  sessionGET( appName & "usr" )  )   >   0   ) then
        isSessionU2D = true
        for i = lBound(usrVars) to uBound(usrVars)
            if not (   len(  sessionGET( appName & "usr" & usrVars(i) )  )   >   0   ) then
                isSessionU2D = false
            end if
        next

        if not(isSessionU2D) then

    '***** CFG ***********************************************************
           tableName  = "Q_Accounts"

           qVars      = array(                        "Code"   )
           qVals      = array(  sessionGET( appName & "usr" )  )
           qValsT     = array(                        "text"   )

    '*********************************************************************

            oRS.source         = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", "*", "", "", "", "" )
            oRS.cursorType     = 0
            oRS.cursorLocation = 2
            oRS.lockType       = 1
            oRS.open()
                if not(oRS.EOF) then
                    for i = lBound(usrVars) to uBound(usrVars)
call sessionPUT( appName & "usr" & usrVars(i), oRS.fields.item( usrVars(i) ).value )
                    next
                end if
            oRS.close()

        end if

    end if   
%>