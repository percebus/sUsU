<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )

'***** CFG *************************************************
    view           = "bids"

    displayVars    = array( "ID", "OfererFullName"       , "Brief"            , "OpeningBid"      , "CurrencyName" )
    displayVarsL   = array( "ID", "Identidad del Ofertor", "Breve Descripción", "Cantidad Inicial", "Moneda"       )

    tableName      = "Q_Bids"
    qVars          = array(                  "UserAccountID"  )
    qVals          = array(  sessionGET( appName & "usrID" )  )
    qValsT         = array(                    "num"          )
    filterVars     = "*"
    orderByVars    = array( "CreatedDate" )
    orderByVarsASC = array(  false        )
'***********************************************************
%>
<!--#include file="src/header.asp"-->

        <table border="1" cellpadding="5" cellspacing="5">
            <thead>
                <tr>
                    <th colspan="<%=( uBound(displayVarsL) +3 )%>">Seleccione una Subasta</th>
                </tr>
                <tr>
                    <th>Editar</th>
<%
    for i = lBound(displayVarsL) to uBound(displayVarsL)
%>
                    <th><%=displayVarsL(i)%></th>
<%
    next
%>
                    <th>Elegir</th>
                </tr>

                <tr><td colspan="<%=( uBound(displayVars) +3 )%>" align="center">[ <a href="setBid.asp">Agregar Subasta</a> ]</td></tr>

            </thead>
            <tbody>
<%
    oRS.source         = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", filterVars, "", "", orderByVars, orderByVarsASC )
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)
%>
                <tr>
                    <td nowrap="true" align="center">[ <a href="setBid.asp?ID=<%=oRS.fields.item("ID").value%>">+</a> ]</td>
<%
            for i = lBound(displayVars) to uBound(displayVars)
%>
                    <td align="center"><a href="bid.asp?ID=<%=oRS.fields.item("ID").value%>"><%=oRS.fields.item( displayVars(i) ).value%></a></td>
<%
            next
%>
                    <td align="center">[ <a href="bid.asp?ID=<%=oRS.fields.item("ID").value%>"> > </a> ]</td>
                </tr>
<%
            oRS.moveNext
        loop
    oRS.close()
%>
            </tbody>
        </table>

<!--#include file="src/footer.asp"-->