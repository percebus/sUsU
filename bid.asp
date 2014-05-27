<!--#include file="src/inc.asp"-->
<%
'***** CFG *************************************************
    view  = "bid"

    DBoID = request("ID")

'   displayVars  = array( "ID, Brief, ItemCategoryID, ClosingDate, MinutesFromClosing, CurrencyName, OpeningBid, BuyOutPrice" )
    qVars  = array( "ID"  )
    qVals  = array( DBoID )
    qValsT = array( "num" )
'***********************************************************
    call varIsRequired(  DBoID, ( appFolder & "search.asp" )  )
%>

<!--#include file="src/header.asp"-->

<%
    isBidOpen          = true
    tableName          = "Q_Bids-TopBid2nd"
    SQLstring          = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", "*", "", "", "", "" )
%><!-- <%=SQLstring%> --><%
    oRS.source         = SQLstring
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        if not(oRS.EOF) then
            isBidOpen       = setDefault(  cBool( oRS.fields.item( "AllBidsAreVisible" ).value ),  isBidOpen  )
            stepsFromMaxBid = setDefault(     oRS.fields.item("StepsFromLastHigherBidder").value, 0 )
%>

        <div>
            <div style="display:block; width:auto; float:right; position:relative">
<%
            forceNewPage = cBool(  setDefault( oRS.fields.item("GalleryForceNewPage").value, true )  )
            galleryName  = setDefault( oRS.fields.item("GalleryName").value, "megazine" )
            if (  not(forceNewPage)  and  not( lCase(galleryName) = "megazine" )  )then
%>
<iframe src="http://laSalleChihuahua.edu.mx/ILS/scripts/JCystems/JCoFTP/<%=galleryName%>.asp?folder=/ILS/site/public/JC/sUsU/bids/<%=oRS.fields.item("ID").value%>&allowNavigation=0&BGcolor=8BCAFE<%=oRS.fields.item("GalleryParams").value%>" width="200" height="200" frameborder="0" marginheight="0" marginwidth="0" scrolling="no"></iframe>
<%
            else
%>
                <a href="http://laSalleChihuahua.edu.mx/ILS/scripts/JCystems/JCoFTP/<%=galleryName%>.asp?folder=/ILS/site/public/JC/sUsU/bids/<%=oRS.fields.item("ID").value%>&allowNavigation=0&BGcolor=8BCAFE<%=oRS.fields.item("GalleryParams").value%>" target="_blank">
                    <img src="<%=imgsFolder%><%=oRS.fields.item("ID").value%>/1.jpg" width="200" border="0" style="float:left; position:relative" /><br />
                    [ Ver Galer&iacute;a Flash ]
                </a>
<%
            end if
%>
            </div>
            ID: <%=oRS.fields.item("ID").value%> | 
            Categor&iacute;a; <a href="search.asp?category=<%=oRS.fields.item("ItemCategoryID").value%>"><%=oRS.fields.item("ItemCategoryName").value%></a><br />
            Breve descripc&iacute;on: <%=oRS.fields.item("Brief").value%>
            <ul>
                <li>Ofertor: <%=oRS.fields.item("OfererFullName").value%></li>
                <li>Fechas <%=oRS.fields.item("OpeningDate").value%> - <%=oRS.fields.item("ClosingDate").value%></li>
<!--
                <li>Tipo ascendente: <%=cBool( oRS.fields.item("Ascendant").value )%></li>
-->
<%
            if (  len( oRS.fields.item("MinNumberOfBidders").value )  >  0  ) then
%>
                <li>N&uacute;mero m&iacute;nimo de postores: <%=oRS.fields.item("MinNumberOfBidders").value%></li>
<%
            end if
            if (  len( oRS.fields.item("MaxBidsPerBidder").value )  >  0  ) then
%>
                <li>M&aacute;ximas pujas por usuario: <%=oRS.fields.item("MaxBidsPerBidder").value%></li>
<%
            end if
            if (  len( oRS.fields.item("StepsFromLastHigherBidder").value )  >  0  ) then
%>
                <li style="color:green">Ganador Paga <%=oRS.fields.item("StepsFromLastHigherBidder").value%> cifra(s) abajo de la m&aacute;s alta</li>
<%
            end if
%>
                <li>Las pujas son visibles: <%=cBool( oRS.fields.item("AllBidsAreVisible").value )%></li>
                <li>Al menos la cifra m&aacute;s alta es visible: <%=cBool( oRS.fields.item("HighestBidIsVisible").value )%></li>
<%
            if (  len( oRS.fields.item("OpeningBid").value )  >  0  ) then
%>
                <li>Oferta empieza: $<%=oRS.fields.item("OpeningBid").value%> <%=oRS.fields.item("CurrencyName").value%></li>
<%
            end if
            if (  len( oRS.fields.item("BidStep").value )  >  0  ) then
%>
                <li>Incremento por puja: $<%=oRS.fields.item("BidStep").value%> <%=oRS.fields.item("CurrencyName").value%></li>
<%
            end if
            if (  len( oRS.fields.item("ReservePrice").value )  >  0  ) then
%>
                <li>Precio de reserva: $<%=oRS.fields.item("ReservePrice").value%> <%=oRS.fields.item("CurrencyName").value%></li>
<%
            end if
            if (  len( oRS.fields.item("BuyOutPrice").value )  >  0  ) then
%>
                <li>Precio BuyOut: $<%=oRS.fields.item("BuyOutPrice").value%> <%=oRS.fields.item("CurrencyName").value%></li>
<%
            end if
%>
             </ul>
             Description: <%=oRS.fields.item("description").value%><br />
            <%=oRS.fields.item("HTML").value%><br />
        </div>
        <div>
<%
            openingBid   = setDefault( oRS.fields.item("OpeningBid"  ).value, 0      )
            bidStep      = setDefault( oRS.fields.item("bidStep"     ).value, 1      )
            maxBid       = setDefault( oRS.fields.item("MaxBid"      ).value, 0      )
            maxBid2      = setDefault( oRS.fields.item("2ndMaxBid"   ).value, maxBid )
            buyoutPrice  = setDefault( oRS.fields.item("BuyOutPrice" ).value, 0      )
            reservePrice = setDefault( oRS.fields.item("ReservePrice").value, 0      )

            if (   len(  sessionGET( appName & "usrID" )  )   >   0   ) then
%>
            <div style="display:block; width:auto">
<%
                if not(  ( buyoutPrice > 0 )  and  not( maxBid < buyoutPrice )  ) then
'               if not(  sessionGET( appName & "usrID" )  =  oRS.fields.item("UserAccountID").value  ) then
%>
                <table>
<%
                    if (  cBool( oRS.fields.item("Ascendant").value )  ) then
%>
                    <form action="<%=defaultAction%>" method="post">
                        <tr>
                            <td align="right">
                                Ofertar: <select name="Ammount">
<%
                        if ( maxBid > 0 ) then
                            startingBid = maxBid + bidStep
                        else
                            startingBid = openingBid
                        end if
                        startingBid = setDefault( startingBid, 1 )

                        endingBid = startingBid + 100
                        for iBid = startingBid to endingBid
%>
                                    <option value="<%=iBid%>"><%=iBid%></option>
<%
                        next
%>
                                </select>
							</td>
                            <td><input type="submit" value="<%=oRS.fields.item("CurrencyName").value%>" /></td>
                        </tr>
<%
                        DBo         = "BidsUsers"
                        operation   = "INSERT"
                        params      = join(  array( "BidID", "UserID", "Ammount" ),  ","  )
                        paramsTypes = join(  array( "num"  , "num"   , "num"     ),  ","  )
%>
                        <input type="hidden" name="UserID"        value="<%=sessionGET( appName & "usrID" )%>" />
                        <input type="hidden" name="BidID"         value="<%=DBoID%>"                           />

                        <input type="hidden" name="operation"     value="<%=operation%>"                       />
                        <input type="hidden" name="DBo"           value="<%=DBo%>"                             />
                        <input type="hidden" name="params"        value="<%=params%>"                          />
                        <input type="hidden" name="paramsTypes"   value="<%=paramsTypes%>"                     />
                        <input type="hidden" name="URL"           value="<%=appFolder%>bid.asp?ID=<%=DBoID%>"  />

                        <input type="hidden" name="stat"          value="1"                                    />
                    </form>
<%
                    end if

                    if ( buyoutPrice > 0 ) then
                        DBo         = "BidsUsers"
                        operation   = "INSERT"
                        params      = join(  array( "BidID", "UserID", "Ammount" ),  ","  )
                        paramsTypes = join(  array( "num"  , "num"   , "num"     ),  ","  )
%>
                    <form action="<%=defaultAction%>" method="post">
                        <input type="hidden" name="UserID"        value="<%=sessionGET( appName & "usrID" )%>" />
                        <input type="hidden" name="BidID"         value="<%=DBoID%>"                           />
                        <input type="hidden" name="Ammount"       value="<%=buyoutPrice%>"                     />

                        <input type="hidden" name="operation"     value="<%=operation%>"                       />
                        <input type="hidden" name="DBo"           value="<%=DBo%>"                             />
                        <input type="hidden" name="params"        value="<%=params%>"                          />
                        <input type="hidden" name="paramsTypes"   value="<%=paramsTypes%>"                     />
                        <input type="hidden" name="URL"           value="<%=appFolder%>bid.asp?ID=<%=DBoID%>"  />

                        <input type="hidden" name="stat"          value="1"                                    />
                        <tr>
                            <td>o COMPRAR AHORA!!!</td><td><input type="submit" value="$<%=buyoutPrice%>&nbsp;<%=oRS.fields.item("CurrencyName").value%>" /></td>
                        </tr>
                    </form>
<%
                    end if
%>
                </table>
            </div>
<%
                else
%>
                Este item ya ha sido vendido
<%
                end if
'               end if
            else
%>
            [ <a href="login.asp">Inicia sesi&oacute;n para comenzar a ofrecer</a> ]
<%
            end if

            if true then ' (   (  cBool( oRS.fields.item("Ascendant").value )  )   and   (  cBool( oRS.fields.item("HighestBidIsVisible").value )  )   ) then
%>
            <style type="text/css">
            <!--
<%
                if ( stepsFromMaxBid > 0 ) then
%>
                #showTopBid2nd
<%
                else
%>
                #showTobBid
<%
                end if
%>
                {
                    color: green;
                }
            -->
            </style>

            <h1 style="float:right; display:block; text-align:right">
                <span id="showTopBid"   >TOP $<%=maxBid%></span><br />
                <span id="showTopBid2nd">2nd TOP $<%=maxBid2%></span><br />
                <span id="showReserve">M&iacute;nimo $<%=reservePrice%> - $<%=maxBid%> = $<%= abs( reservePrice - maxBid )%></span><br />
<%
                if ( buyoutPrice > 0 ) then
%>
                <span id="showBuyOut">M&aacute;ximo $<%=buyoutPrice %> - $<%=maxBid%> = $<%= abs( buyoutPrice  - maxBid )%></span>
<%
                end if
%>
            </h1>
<%
            end if
%>
        </div>
<%
        else
%>
        Lo sentimos, de momento no tenemos art&iacute;culos disponibles bajo esa categor&iacute;a
<%
        end if
    oRS.close()

    qVars          = array( "BidID" )
    qVals          = array(  DBoID  )
    qValsT         = array(  "num"  )

    orderByVars    = array( "Date" )
    orderByVarsASC = array( false )

    tableName      = "Q_BidsUsers"
    SQLstring      = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, false, "", "", "", "*", "", "", orderByVars, orderByVarsASC )
%><!-- <%=SQLstring%> --><%
    oRS.source         = SQLstring
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        if not(oRS.EOF) then
%>
        <table>
            <thead>
                <tr>
                    <th>Fecha</th>
                    <th>Usuario</th>
                    <th>Cifra</th>
                </tr>
            </thead>
            <tbody>
<%
            do until(oRS.EOF)
                if (isBidOpen) then
%>
                <tr>
                    <td              ><%=oRS.fields.item("Date").value%></td>
                    <td              ><%=oRS.fields.item("UserDisplay").value%></td>
                    <td align="right"><%=oRS.fields.item("Ammount").value%></td>
                </tr>
<%
                else
%>
                <tr>
                    <td              ><%=oRS.fields.item("Date").value%></td>
                    <td              >???</td>
                    <td align="right">$XX.XX</td>
                </tr>
<%
                end if
                oRS.moveNext
            loop
%>
            </tbody>
        </table>
<%
        end if
    oRS.close()
%>

<!--#include file="src/footer.asp"-->