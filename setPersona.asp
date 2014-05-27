<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )

'***** CFG *************************************************************************
    view         = "setPersona"
    DBoID        = request("ID")
    action       = defaultAction
    doEditFields = cBool( request("RW") )
    URLwhenDone  = appFolder & "personas.asp"
%>
<!--#include file="src/code/DBoperation.asp"    -->
<%
    set oDB2SQL = new DB2SQL
        oDB2SQL.fVars      = defaultDBfields
        oDB2SQL.orderBy    = defaultDBfieldsOrder
        oDB2SQL.orderBy    = defaultDBfieldsOrderASC

        oDB2SQL.DBo        = "UI_TypesOfPersonas"
            typesOfPersons = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo        = "UI_CitiesG2P"
            cities         = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo        = "UI_Languages"
            languages      = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo        = "UI_Companies"
            companies      = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo        = "UI_AddressesG2P"
            addresses      = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo        = "UI_Phones"
            phones         = data2array( oRS, oDB2SQL.SQLstring )

' PERSONAS
        oDB2SQL.DBo        = "Q_PersonasAccounts"
        oDB2SQL.qVars      = array(             "UserAccountID"      )
        oDB2SQL.qVals      = array( sessionGET( appName & "usrID" )  )
        oDB2SQL.qValsT     = array(             "num"              )
            personas       = data2array( oRS, oDB2SQL.SQLstring )

    set oDB2SQL = nothing

'*****************************************************************************************
%>
<!--#include file="src/header.asp"-->
<%
    if ( operation = "SELECT" ) then
        fieldsDisableds = true
        action          = ""
    end if

    fieldsLinksTargets  = "_self"

    fieldsDisplays      = array ( "Nombre Corto", "Apellido | Empresa"  , "Apellido Materno", "Nombre"   , "Es persona física", "Título"      , "Fecha de Nacimiento", "Ciudad de Nacimiento" , "Sexo"       , "Dirección Principal", "Idioma Principal"         ,"2do idioma"               , "Empesa Principal")
    filedsIDs           = array ( "ShortName"   , "FamilyName"          , "MiddleName"      , "GivenName", "isPhisicalPerson" , "TypeOfPerson", "DateOfBirth"        , "CityOfBirthID"        , "wasBornMale", "MainAddressID"      , "MainLanguageID"           ,"2ndLanguageID"            , "MainCompanyID"   )
    fieldsClasses       = array ( ""            , ""                    , ""                , ""         , ""                 , ""            , "datePicker"         , ""                     , ""           , ""                   , ""                         ,""                         , ""                )
    fieldsTypes         = array ( "text"        , "text"                , "text"            , "text"     , "radio"            , "combo"       , "date"               , "combo"                , "radio"      , "combo"              , "combo"                    ,"combo"                    , "combo"           )
    fieldsTypesDisabled = array ( "label"       , "label"               , "label"           , "label"    , ""                 , ""            , "label"              , ""                     , ""           , ""                   , ""                         ,""                         , ""                )
    fieldsData          = array ( ""            , ""                    , ""                , ""         ,  ToF               , typesOfPersons, ""                   ,  cities                , sexs         ,  addresses           ,  languages                 , languages                 ,  companies        )
    fieldsLinks         = array ( ""            , ""                    , ""                , ""         , ""                 , URLtitles     , ""                   , URLsuggestions & "city", ""           , URLaddresses         , URLsuggestions & "language",URLsuggestions & "language", ""                )
    if (  ( operation = "SELECT" )  or  ( operation = "UPDATE" )  ) then
        oRS.source         = buildSQLstring( "SELECT", DBMS, "Personas", array("ID", "UserAccountID"), array(DBoID, sessionGET( appName & "usrID" ) ), array("num", "num"), false, "", "", "", "*", "", "", "", "" )
        oRS.cursorType     = 0
        oRS.cursorLocation = 2
        oRS.lockType       = 1
        oRS.open()
            if not(oRS.EOF) then
                   dim fieldsDefaultValues ()
                 reDim fieldsDefaultValues ( uBound(filedsIDs) )

                for iFieldsIDs = lBound(filedsIDs) to uBound(filedsIDs)
                    fieldsDefaultValues(iFieldsIDs) = oRS.fields.item( filedsIDs(iFieldsIDs) ).value
                next
            else
response.redirect("setPersona.asp")
            end if
        oRS.close()
    else
        fieldsDefaultValues = ""
        fieldsDisableds     = ""
    end if

    fieldsNames       = "" ' use fieldsIDs

    filedsUseMultples = ""
    fieldsSizes       = "" ' use default
    fieldsRegExp      = ""

    fieldsTabsIndexes = "" ' use default
    fieldsCSSs        = ""

    optionsHeaders    = ""
    optionsSeparators = " "
    optionsRepeaters  = ", "
    optionsFooters    = ""
%>
        <fieldset>
            <legend style="text-align:center">Personas | Identidades</legend>
            <table>
                <form action="<%=action%>" method="post">
<!--
                    <thead>
                        <tr><th colspan="2"></th></tr>
                    </thead>
-->
                    <tbody>
<%
                      '( DOMIDs   , DOMnames   , displays      , HTMLtypes  , HTTPtypesDisabled  . useMultiples     , sizes      , regExps      , data      , defaultValues      , DOMclasses   , CSSs      , tabIndexes      , disableds      , JScripts      , optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, hRefs      , targets            )
    garbage = buildForm( filedsIDs, fieldsNames, fieldsDisplays, fieldsTypes, fieldsTypesDisabled, fieldsUseMultiple, fieldsSizes, fieldsRegExps, fieldsData, fieldsDefaultValues, fieldsClasses, fieldsCSSs, fieldsTabIndexes, fieldsDisableds, fieldsJScripts, optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, fieldsLinks, fieldsLinksTargets )
%>

                        <tr><td colspan="2" align="center"><hr />Datos de Facturaci&oacute;n<hr /></td></tr>
<%
    fieldsDisplays      = array ( "RFC"        , "Nombre"        , "Telefono"      , "Dirección"       , "Enviar Factura a esta dirección" )
    filedsIDs           = array ( "BillingCode", "BillingName"   , "BillingPhoneID", "BillingAddressID", "SendBillToAddressID"             )
    fieldsNames         = "" ' use fieldsIDs
    fieldsTypes         = array ( "text"       , "text"          , "combo"         , "combo"           , "combo"                           )
    fieldsTypesDisabled = array ( "label"      , "label"         , ""              , ""                , ""                                )
    fieldsData          = array ( ""           , ""              ,  phones         ,  addresses        ,  addresses                        )
    fieldsLinks         = array ( ""           , ""              , URLphones       , URLaddresses      , appFolder & "addresses.asp"       )
    if (  ( operation = "SELECT" )  or  ( operation = "UPDATE" )  ) then
        oRS.source         = buildSQLstring( "SELECT", DBMS, "Personas", array("ID", "UserAccountID"), array(DBoID, sessionGET( appName & "usrID" ) ), array("num", "num"), false, "", "", "", "*", "", "", "", "" )
        oRS.cursorType     = 0
        oRS.cursorLocation = 2
        oRS.lockType       = 1
        oRS.open()
            if not(oRS.EOF) then
                 reDim fieldsDefaultValues ( uBound(filedsIDs) )

                for iFieldsIDs = lBound(filedsIDs) to uBound(filedsIDs)
                    fieldsDefaultValues(iFieldsIDs) = oRS.fields.item( filedsIDs(iFieldsIDs) ).value
                next
            else
'response.redirect("setPersona.asp")
            end if
        oRS.close()
    else
        fieldsDefaultValues = ""
        fieldsDisableds     = ""
    end if

    garbage = buildForm( filedsIDs, fieldsNames, fieldsDisplays, fieldsTypes, fieldsTypesDisabled, fieldsUseMultiple, fieldsSizes, fieldsRegExps, fieldsData, fieldsDefaultValues, fieldsClasses, fieldsCSSs, fieldsTabIndexes, fieldsDisableds, fieldsJScripts, optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, fieldsLinks, fieldsLinksTargets )
%>
                    </tbody>
                    <tfoot>
                        <tr>
                            <td colspan="2" align="center">
<%
    if ( operation = "SELECT" ) then
%>
                        [ <a href="setPersona.asp?ID=<%=DBoID%>&RW=1">Modificar mis datos</a> ]
<%
    else
        if ( operation = "UPDATE" ) then
            DBo            = "Personas"
            params         = join(  array( "ShortName", "FamilyName", "MiddleName", "GivenName", "isPhisicalPerson", "TypeOfPerson", "DateOfBirth", "CityOfBirthID", "wasBornMale", "MainAddressID", "MainLanguageID", "2ndLanguageID", "MainCompanyID", "BillToID", "BillingCode", "BillingName", "BillingPhoneID", "BillingAddressID", "SendBillToAddressID" ),  ","  )
            paramsTypes    = join(  array( "text"     , "text"      , "text"      , "text"     , "bool"            , "num"         , "date"       , "num"          , "bool"       , "num"          , "num"           , "num"          , "num"          , "num"     , "text"       , "text"       , "num"           , "num"             , "num"                 ),  ","  )

            criterias      = join(  array( "ID" , "UserAccountID" ),  ","  )
            criteriasTypes = join(  array( "num", "num"           ),  ","  )
%>
                                <input type="hidden" name="ID"            value="<%=DBoID%>"           />
                                <input type="hidden" name="criterias"      value="<%=criterias%>"      />
                                <input type="hidden" name="criteriasTypes" value="<%=criteriasTypes%>" />
<%
        else
            DBo         = "Personas"
            params      = join(  array ( "ShortName", "FamilyName", "MiddleName", "GivenName", "isPhisicalPerson", "TypeOfPerson", "UserAccountID", "DateOfBirth", "CityOfBirthID", "wasBornMale", "MainAddressID", "MainLanguageID", "2ndLanguageID", "MainCompanyID", "BillToID", "BillingCode", "BillingName", "BillingPhoneID", "BillingAddressID", "SendBillToAddressID" ),  ","  )
            paramsTypes = join(  array ( "text"     , "text"      , "text"      , "text"     , "bool"            , "num"         , "num"          , "date"       , "num"          , "bool"       , "num"          , "num"           , "num"          , "num"          , "num"     , "text"       , "text"       , "num"           , "num"             , "num"                 ),  ","  )
        end if
%>
                                <input type="hidden" name="UserAccountID" value="<%=sessionGET( appName & "usrID" )%>" />
                                <input type="hidden" name="operation"     value="<%=operation%>"                       />
                                <input type="hidden" name="DBo"           value="<%=DBo%>"                             />
                                <input type="hidden" name="params"        value="<%=params%>"                          />
                                <input type="hidden" name="paramsTypes"   value="<%=paramsTypes%>"                     />
                                <input type="hidden" name="URL"           value="<%=URLwhenDone%>"                     />

                                <input type="hidden" name="stat"          value="1"                                    />

                                <input type="submit" value=" > > > " style="width:100%" />
<%
    end if
%>
                            </td>
                        </tr>
                    </tfoot>
                </form>
            </table>
        </fieldset>

<!--#include file="src/footer.asp"-->