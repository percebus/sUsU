<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )

'***** CFG *************************************************************************
    view         = "setBid"
    DBoID        = request("ID")
    action       = defaultAction
    doEditFields = cBool( request("RW") )
    URLwhenDone  = appFolder & "bids.asp"
%>
<!--#include file="src/code/DBoperation.asp"    -->
<%
    set oDB2SQL = new DB2SQL
        oDB2SQL.fVars   = defaultDBfields
        oDB2SQL.orderBy = defaultDBfieldsOrder
        oDB2SQL.orderBy = defaultDBfieldsOrderASC

        oDB2SQL.DBo     = "UI_ItemCategories"
            categories  = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo     = "UI_TypesOfBids"
            typesOfBids = data2array( oRS, oDB2SQL.SQLstring )

        oDB2SQL.DBo     = "UI_Currencies"
            currencies  = data2array( oRS, oDB2SQL.SQLstring )

' PERSONAS
        oDB2SQL.DBo     = "Q_PersonasAccounts"
        oDB2SQL.qVars   = array(             "UserAccountID"      )
        oDB2SQL.qVals   = array( sessionGET( appName & "usrID" )  )
        oDB2SQL.qValsT  = array(             "num"              )
            personas    = data2array( oRS, oDB2SQL.SQLstring )

    set oDB2SQL = nothing

'*****************************************************************************************
%>
<!--#include file="src/header.asp"-->
<%
    fieldsDisplays      = array ( "Persona | Identidad", "Categoría"                , "Descripci&oacute;n Breve", "Tipo de Subasta Inicial"   , "Tipo de Subasta Actual"    , "Ascendente", "Pagar n cifras mas abajo de la m&aacute;s alta", "Se muestran las ofertas", "Se muestran almenos totales", "Oferta Inicial", "Moneda"                   , "Pago Inicial"               , "Precio de Reserva", "Escal&oacute;n por oferta", "Precio de Compra Directa", "Valor Estimado", "Moneda del Valor Estimado", "Fecha de Creación", "Fecha de Inicio", "Fecha de Cierre", "Descripción", "HTML"     )
    filedsIDs           = array ( "OfererAccountID"    , "ItemCategoryID"           , "Brief"                   , "OriginalTypeOfBid"         , "LastTypeOfBid"             , "Ascendant" , "StepsFromLastHigherBidder"                     , "AllBidsAreVisible"      , "HighestBidIsVisible"        , "OpeningBid"    , "CurrencyID"               , "PaymentPercentBeforeSending", "ReservePrice"     , "BidStep"                  , "BuyoutPrice"             , "EstimatedPrice", "EstimatedPriceCurrencyID" , "CreatedDate"      , "OpeningDate"    , "ClosingDate"    , "Description", "HTML"     )
    DBtypes             = array ( "num"                , "num"                      , "text"                    , "num"                       , "num"                       , "bool"      , "num"                                           , "bool"                   , "bool"                       , "num"           , "num"                      , "num"                        , "num"              , "num"                      , "num"                     , "num"           , "num"                      , "date"             , "date"           , "date"           , "text"       , "text"     )
    fieldsClasses       = array ( ""                   , ""                         , ""                        , ""                          , ""                          , ""          , ""                                              , ""                       , ""                           , ""              , ""                         , ""                           , ""                 , ""                         , ""                        , ""              , ""                         , ""                 , "datePicker"     , "datePicker"     , ""           , "ckeditor" )
    fieldsTypes         = array ( "combo"              , "combo"                    , "text"                    , "combo"                     , "combo"                     , "radio"     , "num"                                           , "radio"                  , "radio"                      , "num"           , "combo"                    , "combo"                      , "num"              , "num"                      , "num"                     , "num"           , "combo"                    , "label"            , "date"           , "date"           , "textarea"   , "textarea" )
    fieldsTypesDisabled = array ( ""                   , ""                         , "label"                   , ""                          , ""                          , ""          , "label"                                         , ""                       , ""                           , "label"         , ""                         , "label"                      , "label"            , "label"                    , "label"                   , "label"         , ""                         , "label"            , "label"          , "label"          , "label"      , "label"    )
    fieldsData          = array ( personas             , categories                 , ""                        , typesOfBids                 , typesOfBids                 , ToF         , ""                                              , ToF                      , ToF                          , ""              ,  currencies                , percents                     , ""                 , ""                         , ""                        , ""              ,  currencies                , ""                 , ""               , ""               , ""           , ""         )
    fieldsLinks         = array ( URLpersonas          , URLsuggestions & "category", ""                        , ""                          , URLsuggestions & "typeOfBid", ""          , ""                                              , ""                       , ""                           , ""              , URLsuggestions & "currency", ""                           , ""                 , ""                         , ""                        , ""              , URLsuggestions & "currency", ""                 , ""               , ""               , ""           , ""         )
    fieldsLinksTargets  = "_self"
    if (  ( operation = "SELECT" )  or  ( operation = "UPDATE" )  ) then
        oRS.source         = buildSQLstring( "SELECT", DBMS, "Bids", array("ID"), array( DBoID ), array( "num" ), true, "", "", "", "*", "", "", "", "" )
        oRS.cursorType     = 0
        oRS.cursorLocation = 2
        oRS.lockType       = 1
        oRS.open()
            if not(oRS.EOF) then
                  dim fieldsDefaultValues (                   )
                reDim fieldsDefaultValues ( uBound(filedsIDs) )
                if not(doEditFields) then
                   fieldsDisableds = true
                   action          = ""
                end if
                for iFieldsIDs = lBound(filedsIDs) to uBound(filedsIDs)
                    fieldsDefaultValues(iFieldsIDs) = oRS.fields.item( filedsIDs(iFieldsIDs) ).value
                next
            else
response.redirect("setBid.asp")
            end if
        oRS.close()
    end if

    fieldsNames         = "" ' use fieldsIDs

    filedsUseMultples   = ""
    fieldsSizes         = 50 ' use default
    fieldsRegExp        = ""

    fieldsTabsIndexes   = "" ' use default
    fieldsCSSs          = ""

    optionsHeaders      = ""
    optionsSeparators   = " "
    optionsRepeaters    = ", "
    optionsFooters      = ""

    if ( operation = "SELECT" ) then
%>
        <div>
            <ul>
                <li><a href="bidItems.asp?bid=<%=    DBoID%>">Art&iacute;culo(s) de la Subasta</a></li>
                <li><a href="bidPayments.asp?bid=<%= DBoID%>">Formas de Pago</a></li>
                <li><a href="bidEULA.asp?bid=<%=     DBoID%>">E.U.L.A.</a></li>
                <li><a href="bidHistorial.asp?bid=<%=DBoID%>">Historial de la subasta</a></li>
            </ul>
        </div>
<%
    end if
%>
        <table border="1" cellpadding="10" cellspacing="10">
            <form action="<%=action%>" method="post">
                <tbody>
<%
                      '( DOMIDs   , DOMnames   , displays      , HTMLtypes  , HTMLtypesDisabled  , useMultiples     , sizes      , regExps      , data      , defaultValues      , DOMclasses   , CSSs      , tabIndexes      , disableds      , JScripts      , optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, hRefs      , targets            )
    garbage = buildForm( filedsIDs, fieldsNames, fieldsDisplays, fieldsTypes, fieldsTypesDisabled, fieldsUseMultiple, fieldsSizes, fieldsRegExps, fieldsData, fieldsDefaultValues, fieldsClasses, fieldsCSSs, fieldsTabIndexes, fieldsDisableds, fieldsJScripts, optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, fieldsLinks, fieldsLinksTargets )
%>
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="2" align="center">
<%
    if ( operation = "SELECT" ) then
%>
                        [ <a href="setBid.asp?ID=<%=DBoID%>&RW=1">Modificar mis datos</a> ]
<%
    else
        if ( operation = "UPDATE" ) then
            DBo            = "Bids"
            criterias      = "ID"
            criteriasTypes = "num"
%>
                            <input type="hidden" name="ID"             value="<%=DBoID%>"          />
                            <input type="hidden" name="criterias"      value="<%=criterias%>"      />
                            <input type="hidden" name="criteriasTypes" value="<%=criteriasTypes%>" />
<%
        else
            DBo = "Bids"
        end if
        params        = join( filedsIDs, "," )
        paramsTypes   = join( DBtypes  , "," )
%>
                            <input type="hidden" name="operation"   value="<%=operation%>"   />
                            <input type="hidden" name="DBo"         value="<%=DBo%>"         />
                            <input type="hidden" name="params"      value="<%=params%>"      />
                            <input type="hidden" name="paramsTypes" value="<%=paramsTypes%>" />
                            <input type="hidden" name="URL"         value="<%=URLwhenDone%>" />

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