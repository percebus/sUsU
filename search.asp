<!--#include file="src/inc.asp"-->
<%
'***** CFG *************************************************
    view           = "search"
    searchPageSize = 4
    itemCategory   = request("category")

'   displayVars  = array( "ID, Brief, ItemCategoryID, ClosingDate, MinutesFromClosing, CurrencyName, OpeningBid, BuyOutPrice" )
    if ( len(itemCategory) > 0 ) then
        qVars  = array( "ItemCategoryID" )
        qVals  = array(  itemCategory    )
        qValsT = array(    "num"         )
    else
        qVars  = ""
        qVals  = ""
        qValsT = ""
    end if

    format = request("format")
'***********************************************************

    if ( format = "feed" ) then
        response.contentType = "text/xml"
%><?xml version="1.0" encoding="ISO-8859-1" ?>
<rss xmlns:media="http://search.yahoo.com/mrss/" xmlns:atom="http://www.w3.org/2005/Atom" version="2.0">
    <channel>
        <title>sUsUbasta <%=ItemCategoryName%></title>
        <description>Antiguedades a subastar</description>
        <docs>http://www.rssboard.org/rss-specification</docs>
        <link>http://<%=request.ServerVariables("HTTP_HOST")%><%=request.ServerVariables("URL")%></link>
        <pubDate><%=dateToRFC822(date)%></pubDate>
        <lastBuildDate><%=dateToRFC822(date)%></lastBuildDate>
<%
    else
%>
<!--#include file="src/header.asp"-->
<%
    end if
    tableName = "Q_BidsSearch"
    pattern   = replace( request("q"), " ", "%" )
    if ( len(pattern) > 0 ) then
        searchSQLstring = "SELECT * FROM [Q_BidsSearch] WHERE [Brief] LIKE '%" & pattern & "%' OR [Description] LIKE '%" & pattern & "%' OR [ItemCategoryName] LIKE '%" & pattern & "%'"
    else
        searchSQLstring = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, true, "", "", "", "*", "", "", "", "" )
    end if

    oRS.source         = searchSQLstring
    oRS.cursorType     = 3 ' 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    if not( format = "feed" ) then
        oRS.pageSize   = searchPageSize
    end if
    oRS.open()
        if not(oRS.EOF) then
            if not( format = "feed" ) then
                totalPages       = oRS.pageCount
                currentPage      = setDefault( request("page"), 1 )
                if ( currentPage < 1 ) then
                    currentPage  = 1
                elseif( cInt(currentPage) > cInt(totalPages) ) then
                    currentPage  = totalPages
                end if
                oRS.absolutePage = currentPage
                maxRows          = currentPage * searchPageSize
            end if
            doContinue = true
            do until ( oRS.EOF or not(doContinue) )
                if ( format = "feed" ) then
%>
        <item>
            <guid isPermaLink="true">http://<%=request.ServerVariables("HTTP_HOST")%><%=appFolder%>bid.asp?ID<%=oRS.fields.item("ID").value%></guid>
            <title><![CDATA[<%=oRS.fields.item("Brief").value%>]]></title>
            <description><![CDATA[<%=oRS.fields.item("Description").value%>]]></description>
            <pubDate><%=dateToRFC822( formatDateG2P(oRS.fields.item("ClosingDate").value) )%></pubDate>
            <atom:updated><%=dateToRFC3339(now)%></atom:updated>
            <link>http://<%=request.ServerVariables("HTTP_HOST")%><%=appFolder%>bid.asp?ID=<%=oRS.fields.item("ID").value%></link>
            <enclosure     url="http://<%=request.ServerVariables("HTTP_HOST")%><%=imgsFolder%><%=oRS.fields.item("ID").value%>/1.jpg" type="image/jpeg" length="0" />
            <media:content url="http://<%=request.ServerVariables("HTTP_HOST")%><%=imgsFolder%><%=oRS.fields.item("ID").value%>/1.jpg" type="image/jpeg">
                <media:title>Foto</media:title>
                <media:description type='html'>
                    <![CDATA[
                        <p>
                            <img src="<%=imgsFolder%><%=oRS.fields.item("ID").value%>/1.jpg" width="130" border="0" style="float:left; position:relative" />
                        </p>
                    ]]>
                </media:description>
            </media:content>
        </item>
<%
                else
%>

        <div id="bid<%=oRS.fields.item("ID").value%>" class="bid" style="clear:left">
            <div class="img" style="display:block; width:auto; float:left; position:relative">
                <a href="bid.asp?ID=<%=oRS.fields.item("ID").value%>"><img src="<%=imgsFolder%><%=oRS.fields.item("ID").value%>/1.jpg" width="130" border="0" style="float:left; position:relative" /></a>
            </div>
            <div class="details" style="padding:10px; float:left; position:relative">
                <span class="typeOfBid"            >Tipo de Subasta: <a href="http://en.wikipedia.org/wiki/Auction" target="_blank"><%=oRS.fields.item("LastTypeOfBidName").value%></a></span><br />
                <span class="brief"                >Descripci&oacute;n Breve: <a href="bid.asp?ID=<%=oRS.fields.item("ID").value%>"><%=oRS.fields.item("Brief").value%></a></span><br />
                <span class="category"             >Categor&iacute;a: <a href="search.asp?category=<%=oRS.fields.item("ItemCategoryID").value%>"><%=oRS.fields.item("ItemCategoryName").value%></a></span><br />
                <span class="closingDate"          >Fecha de Cierre: <%=oRS.fields.item("ClosingDate").value%></span>
                <span class="minutesForClosingDate">( faltan <%=oRS.fields.item("MinutesFromClosing").value%> minutos )</span><br />
                <span class="openingBid"           >Desde $<%=oRS.fields.item("OpeningBid").value%>&nbsp;<%=oRS.fields.item("CurrencyName").value%></span><br />
                <span class="buyoutPrice"          >Compra inmediata $<%=oRS.fields.item("BuyOutPrice").value%>&nbsp;<%=oRS.fields.item("CurrencyName").value%></span>
            </div>
            <div style="clear:left">
                <hr />
            </div>
        </div>
<%
                end if
                oRS.moveNext
                if (  not( format = "feed" )  and  ( cInt(oRS.absolutePosition) > cInt(maxRows) )  ) then
                     doContinue = false
                end if
            loop

            q = ""
            if (  len( request("q") )  >  0  ) then
                q = "&q=" & request("q")
            end if
            if ( len(itemCategory) > 0 ) then
                q = "&category=" & itemCategory
            end if

%>
        <div style="text-align:center">
<%
            if ( currentPage > 1 ) then
%>
            <a href="search.asp?nothing=<%=q%>">1 |&lt;</a>
            <a href="search.asp?page=<%=currentPage -1%><%=q%>"><%=currentPage -1%> &lt;</a>
<%
            end if
%>
            <%=currentPage%>
<%
            if ( cInt(currentPage) < cInt(oRS.pageCount) ) then
%>
            <a href="search.asp?page=<%=currentPage +1%><%=q%>">&gt; <%=currentPage +1%></a>
            <a href="search.asp?page=<%=oRS.pageCount%><%=q%>">&gt;| <%=oRS.pageCount%></a>
<%
           end if
%>
        </div>
<%
        else
            if not( format = "feed" ) then
%>
        Lo sentimos, de momento no tenemos art&iacute;culos disponibles bajo esa categor&iacute;a
<%
            end if
        end if
    oRS.close()
    if ( format = "feed" ) then
%>
    </channel>
</rss>
<%
    else
%>
<!--#include file="src/footer.asp"-->
<%
    end if
%>