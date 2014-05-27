<!--#include file="src/inc.asp"-->
<%
'    call varIsRequired(  sessionGET( appName & "usr"   )               ,  ( appFolder & "login.asp" )  )
'    call varIsRequired(  sessionGET( appName & "usrID" )               ,  ( appFolder & "login.asp" )  )
'    call varIsRequired(  sessionGET( appName & "usrSecurityLevelCode" ),  ( appFolder & "login.asp" )  )
'    if (   cInt(  sessionGET( appName & "usrSecurityLevelCode" )  )   >   minDBadminSecurity   ) then
'response.redirect("login.asp")
'    end if
    call varIsRequired( request("table"),  ( appFolder & "adminTables.asp" )  )

'***** CFG *************************************************************************
    view         = "adminTableSet"
    tableName    = request("table")
    DBoID        = request("ID")
    action       = defaultAction
    R            = request("R")
    W            = request("W")
    RW           = request("RW")
    X            = request("X")
    doEditFields = false
    URLwhenDone = request("URL")
    for iURLcode = lBound(URLdecodes) to uBound(URLdecodes)
        URLwhenDone = replace( URLwhenDone, URLdecodes(iURLcode), URLencodes(iURLcode) )
    next
    if ( cBool(W) or cBool(RW) ) then
        doEditFields = setDefault(  cBool( request("doRW") ),  doEditFields  )
    end if
'***********************************************************************************

    operation = ""
    if ( cBool(doEditFields) ) then
        if ( DBoID > 0 ) then
            operation = "UPDATE"
        else
            operation = "INSERT"
        end if
    else
        operation = "SELECT"
    end if

    set oDB2SQL = new DB2SQL
        oDB2SQL.DBo    = tableFKs
                         '             0                       1
        oDB2SQL.fVars      = array( tableFKsLocalColumn, tableFKsRemoteObject )
        oDB2SQL.qVars      = array( tableFKsLocalObject, tableFKsRemoteColumn )
        oDB2SQL.qVals      = array( tableName          , "ID"                 )
        oDB2SQL.qValsT     = array( "text"             , "text"               )
        oDB2SQL.orderBy    = array( tableFKsLocalColumn )
        oDB2SQL.orderByASC = array( true                )
             FKs           = data2array( oRS, oDB2SQL.SQLstring )
    set oDB2SQL = nothing

    qVars  = array( "ID"  )
    qValsT = array( "num" )
    if (  ( operation = "SELECT" )  or  ( operation = "UPDATE" )  ) then
        qVals  = array( DBoID )
    else
        qVals  = array( "0"   )
    end if

    oRS.source         = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", "*", "", "", "", "" )
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()

        set oDB2SQL = new DB2SQL
            oDB2SQL.fVars   = defaultDBfields
            oDB2SQL.orderBy = defaultDBfieldsOrder
            oDB2SQL.orderBy = defaultDBfieldsOrderASC

        set oRS2                  = server.createObject("ADODB.Recordset")
            oRS2.activeConnection = DBreadConnection

          dim data()
        reDim data(0)
        tableFieldsNames = ""
        tableFieldsTypes = ""
        formFieldsTypes  = ""
        defaultValues    = ""

        iCurrentField    = -1
        for each currentField in oRS.fields
            iCurrentField = iCurrentField + 1

        ' ***** DB Table Field Name
            if ( len(tableFieldsNames) > 0 ) then
                tableFieldsNames = tableFieldsNames & ","
            end if
            tableFieldsNames = tableFieldsNames & currentField.name

        ' ***** Form Field Data
            reDim preserve data(iCurrentField)
            currentData = ""
            if ( isArray(FKs) ) then
                for iFK = lBound(FKs,2) to uBound(FKs,2)
                    if ( currentField.name = FKs(0,iFK) ) then
                        oDB2SQL.DBo = FKs(1,iFK)
                        currentData = data2array( oRS2, oDB2SQL.SQLstring )
                    end if
                next
            end if
            data(iCurrentField) = currentData

        ' ***** DB Table Field Type
            currentTypeSQL = ""
            iFoundType     = -1
            iType          = lBound(typesVars)
            do until(  ( iType > uBound(typesVars) )  or  ( len(currentTypeSQL) > 0 )  )
                if (  isArray( typesVars(iType) )  ) then
                    tempVar = iInArray( currentField.type, typesVars(iType) )
                    if not( tempVar < 0 ) then
                        iFoundType     = iType
                        currentTypeSQL = typesVals(iType)
                    end if
                else
                    if ( typesVars(iType) = currentField.type ) then
                        iFoundType     = iType
                        currentTypeSQL = typesVals(iType)
                    end if
                end if
                iType = iType +1
            loop
            if ( len(tableFieldsTypes) > 0 ) then
                tableFieldsTypes = tableFieldsTypes & ","
            end if
            tableFieldsTypes = tableFieldsTypes & currentTypeSQL


        ' ***** Form Field Type
            if ( len(formFieldsTypes) > 0 ) then
                formFieldsTypes = formFieldsTypes & ","
            end if
            if ( isArray(currentData) ) then
                formFieldsTypes = formFieldsTypes & "combo"
            elseif (  ( uCase(currentField.name) = "ID" )  or  ( lCase(currentField.name) = "inactive" )  ) then
                formFieldsTypes = formFieldsTypes & "label"
            else
                if ( iFoundType > 0 ) then
                    if ( operation = "SELECT" ) then
                        formFieldsTypes = formFieldsTypes & typesUIsR(iFoundType)
                    else
                        formFieldsTypes = formFieldsTypes & typesUIsW(iFoundType)
                    end if
                else
                    if ( operation = "SELECT" ) then
                        formFieldsTypes = formFieldsTypes & "label"
                    else
                        formFieldsTypes = formFieldsTypes & "text"
                    end if
                end if
            end if

        ' ***** Form Field Default Values
            if not(oRS.EOF) then
                if ( len(defaultValues) > 0 ) then
                    defaultValues = defaultValues & ","
                end if
                if ( len(currentField.value) > 0 ) then
                    defaultValues = defaultValues & replace( currentField.value, ",", ";" )
                else
                    defaultValues = defaultValues & emptyFieldMask
                end if
            end if

        next

    oRS.close()

    if (  ( operation = "INSERT" )  or  ( operation = "UPDATE" )  ) then
        fieldsTypesDisabled = false
    else
        fieldsTypesDisabled = true
    end if
    filedsIDs   = split( tableFieldsNames, "," )
    DBtypes     = split( tableFieldsTypes, "," )
    fieldsTypes = split( formFieldsTypes , "," )
    fieldsData  = data
    if ( len(defaultValues) > 0 ) then
        defaultValues       = replace( defaultValues, emptyFieldMask, "" )
        fieldsDefaultValues = split( defaultValues, "," )
    else
        fieldsDefaultValues = ""
    end if
    fieldsLinks        = ""
    fieldsLinksTargets = ""
%>

<!--#include file="src/header.asp"-->


        <div id="navBreadcrums">
            <nav id="breadcrums">
                <ul>
                    <li><a href="adminTables.asp">Tablas</a></li>
                    <li><a href="adminTablesRecords.asp?table=<%=tableName%>&R=<%=R%>&W=<%=W%>&RW=<%=RW%>&X=<%=X%>"><%=tableName%></a></li>
                    <li>
                        <%=operation%>
<%
    if ( operation = "UPDATE" ) then
%>
                        ID:<%=DBoID%>
<%
    end if
%>
                    </li>
                </ul>
            </nav>
        </div>

        <table>
            <form action="<%=action%>" method="post">
<!-- 
    DEBUG
    * <%=tableFieldsNames%>
    * <%=tableFieldsTypes%>
    * <%=formFieldsTypes%>
    * <%=defaultValues%>
    * DISABLED: <%=fieldsTypesDisabled%>
    * URL: <%=URLwhenDone%>
-->
                <tbody>
<%
                      '( DOMIDs   , DOMnames, displays, HTMLtypes  , HTMLtypesDisabled, useMultiples, sizes, regExps, data      , defaultValues      , DOMclasses, CSSs, tabIndexes, disableds          , JScripts, optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, hRefs      , targets            )
    garbage = buildForm( filedsIDs, ""      , ""      , fieldsTypes, ""               , ""          , ""   , ""     , fieldsData, fieldsDefaultValues, ""        , ""  , ""        , fieldsTypesDisabled, ""      , ""            , ""               , ""              , ""            , fieldsLinks, fieldsLinksTargets )
%>
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="2" align="center">
<%
    if ( operation = "SELECT" ) then
        if ( cBool(RW) ) then
%>
                        [ <a href="adminTablesRecordsSet.asp?table=<%=tableName%>&ID=<%=DBoID%>&R=<%=R%>&W=<%=W%>&RW=<%=RW%>&X=<%=X%>&doRW=1&URL=<%=URLwhenDone%>">Modificar datos</a> ]
<%
        else
%>
                        
<%
        end if
    else
        if ( operation = "UPDATE" ) then
            DBo            = tableName
            criterias      = "ID"
            criteriasTypes = "num"
%>
                            <input type="hidden" name="ID"                   value="<%=DBoID%>"          />
                            <input type="hidden" name="criterias"            value="<%=criterias%>"      />
                            <input type="hidden" name="criteriasTypes"       value="<%=criteriasTypes%>" />
<%
        else
            DBo = tableName
        end if
        params      = tableFieldsNames
        paramsTypes = tableFieldsTypes
%>
                            <input type="hidden" name="operation"            value="<%=operation%>"     />
                            <input type="hidden" name="DBo"                  value="<%=DBo%>"           />
                            <input type="hidden" name="params"               value="<%=params%>"        />
                            <input type="hidden" name="paramsOriginalValues" value="<%=defaultValues%>" />
                            <input type="hidden" name="paramsTypes"          value="<%=paramsTypes%>"   />
                            <input type="hidden" name="URL"                  value="<%=URLwhenDone%>"   />

                            <input type="hidden" name="stat"                 value="1"              />

                            <input type="submit" value=" > > > " style="width:100%" />
<%
    end if
%>

                        </td>
                    </tr>
                </tfoot>
            </form>
        </table>

<!--#include file="src/footer.asp"-->