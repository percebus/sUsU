<%
    function formatDBdataTypes( fieldsValues, fieldsValuesTypes, DBMS )
        if ( isArray(fieldsValues) and isArray(fieldsValuesTypes) ) then
            fieldsValuesWithDBformat = fieldsValues
            for i = lBound(fieldsValuesWithDBformat) to uBound(fieldsValuesWithDBformat)
                select case replace( lCase(DBMS), " ", "" )
                    case "msaccess"
                       select case lCase( fieldsValuesTypes(i) )
                           case "num"
                               fieldsValuesWithDBformat(i) =       fieldsValuesWithDBformat(i)
 
                           case "autonumeric"

                           case "text"
                               fieldsValuesWithDBformat(i) = "'" & fieldsValuesWithDBformat(i) & "'"

                            case "date"
                               fieldsValuesWithDBformat(i) = "'" & fieldsValuesWithDBformat(i) & "'"

                           case "time"
                               fieldsValuesWithDBformat(i) = "'" & fieldsValuesWithDBformat(i) & "'"

                           case "bool"
                               if (  cBool( fieldsValuesWithDBformat(i) )  ) then
                                   fieldsValuesWithDBformat(i) = "TRUE"
                               else
                                   fieldsValuesWithDBformat(i) = "FALSE"
                               end if

                           case "null"
                               fieldsValuesWithDBformat(i) = "NULL"

                           case else
                       end select

                   case else

                end select
            next
        end if
        formatDBdataTypes = fieldsValuesWithDBformat
    end function
%>