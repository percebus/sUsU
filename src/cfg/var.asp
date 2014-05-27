<%
    GMTdifference     = serverGMT-GMT
    nowGMT            = formatDateG2P(  dateAdd( "h", GMTdifference, now() )  )

'    if not(debugRequested) then
'        on error resume next
'    end if

    if not (  len( session("tries") )  >  0  ) then
        session("tries") = 0
    end if
    if not (  len ( request.cookies("session")("tries") )  > 0  ) then
        response.cookies("session")("tries") = 0
    end if
%>