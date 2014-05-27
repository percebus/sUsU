<%
    function buildSQLstring( action, DBMS, DBobjectName, objectFieldsNames, objectFieldsValues, objectFieldsValuesTypes, useOR, objectFieldsUpDateNames, objectFieldsUpDateValues, objectFieldsUpDateValuesTypes, objectFieldsNamesFilter, objectFieldsNamesFilterAliases, limitRowCount, objectFieldsNamesOrderBy, objectFieldsNamesOrderByASC )
        DBMS       = replace( lCase(DBMS), " ", "" )
        objectName = formatDBselector(DBobjectName)
        tempVal    = ""
        action     = uCase(action)

        if (  ( action = "INSERT" )  OR  ( action = "UPDATE" )  ) then
            if ( isArray(objectFieldsUpDateValuesTypes) ) then
                objectFieldsUpDateValuesWithFormat = formatDBdataTypes( objectFieldsUpDateValues, objectFieldsUpDateValuesTypes, DBMS )
            end if
        end if

        whereSQLstring = ""
        if ( isArray(objectFieldsNames) ) then
            if ( isArray(objectFieldsValues) ) then
                if ( isArray(objectFieldsValuesTypes) ) then
                    objectFieldsValuesWithDBformat = formatDBdataTypes( objectFieldsValues, objectFieldsValuesTypes, DBMS )
                end if

                for i = lBound(objectFieldsNames) to uBound(objectFieldsNames)
                    if (  len( objectFieldsValues(i) )  >  0  ) then
                        if ( len(whereSQLstring) > 0 ) then
                            if ( cBool(useOR) ) then
                                whereSQLstring = whereSQLstring & " OR "
                            else
                                whereSQLstring = whereSQLstring & " AND "
                            end if
                        end if
                        whereSQLstring = whereSQLstring  &  objectName & "." & formatDBselector( objectFieldsNames(i) )  &  " = "  &  objectFieldsValuesWithDBformat(i)
                    end if
                 next
            end if
            whereSQLstring = " WHERE " & whereSQLstring
        end if

        select case action
            case "SELECT"
                tempVal    = action & " "
                currentVal = ""
                if ( len(limitRowCount) > 0 ) then
                    if (  (  DBMS = "msaccess" )  or  (  DBMS = "mssqlserver" )  ) then
                        tempVal = tempVal & "TOP " & limitRowCount & " "
                    elseif (  DBMS = "firebird" ) then
                        tempVal = tempVal & "FIRST " & limitRowCount & " "
                    end if
                end if
                if ( isArray(objectFieldsNamesFilter) ) then
                    for i = lBound(objectFieldsNamesFilter) to uBound(objectFieldsNamesFilter)
                        if (  len( objectFieldsNamesFilter(i) )  >  0  ) then
                            if ( len(currentVal) > 0 ) then
                                currentVal = currentVal & ", "
                            end if
                            currentVal = currentVal & objectName & "." & formatDBselector( objectFieldsNamesFilter(i) )
                            if ( isArray(objectFieldsNamesFilterAliases) ) then
                                if (  len( objectFieldsNamesFilterAliases(i) )  >  0  ) then
                                    currentVal = currentVal & " AS " & objectFieldsNamesFilterAliases(i)
                                end if
                            end if
                        end if
                    next
                    tempVal = tempVal & currentVal
                else
                    tempVal = tempVal & "*"
                end if

                tempVal = tempVal & " FROM " & objectName & whereSQLstring

                currentVal = ""
                if ( isArray(objectFieldsNamesOrderBy) ) then
                    for i = lBound(objectFieldsNamesOrderBy) to uBound(objectFieldsNamesOrderBy)
                        if (  len( objectFieldsNamesOrderBy(i) )  >  0  ) then
                            if ( len(currentVal) > 0 ) then
                                currentVal = currentVal & ", "
                            end if
                            currentVal = currentVal & formatDBselector( objectFieldsNamesOrderBy(i) )
                            if not(  cBool( objectFieldsNamesOrderByASC(i) )  ) then
                                currentVal = currentVal & " DESC"
                            end if
                        end if
                    next
                    if ( len(currentVal) > 0 ) then
                        currentVal = " ORDER BY " & currentVal
                    end if
                end if

                tempVal = tempVal & currentVal & ";"

            case "INSERT"
                tempVal    = action & " INTO " & objectName
                currentVal = ""
                if ( isArray(objectFieldsUpDateNames) ) then
                    tempVal = tempVal & " ("
                    for i = lBound(objectFieldsUpDateNames) to uBound(objectFieldsUpDateNames)
                        if (  len( objectFieldsUpDateValues(i) )  >  0  ) then
                            if ( len(currentVal) > 0 ) then
                                currentVal = currentVal & ", "
                            end if
                            currentVal = currentVal & formatDBselector( objectFieldsUpDateNames(i) )
                        end if
                    next
                    tempVal = tempVal & currentVal & ")"
                end if

                currentVal = ""
                if ( isArray(objectFieldsUpDateValuesWithFormat) ) then
                    tempVal = tempVal & " VALUES ("
                    for i = lBound(objectFieldsUpDateValuesWithFormat) to uBound(objectFieldsUpDateValuesWithFormat)
                        if (  len( objectFieldsUpDateValues(i) )  >  0  ) then
                            if ( len(currentVal) > 0 ) then
                                currentVal = currentVal & ", "
                            end if
                            currentVal = currentVal & objectFieldsUpDateValuesWithFormat(i)
                        end if
                    next
                    tempVal = tempVal & currentVal & ")"
                end if

                tempVal = tempVal & ";"

            case "UPDATE"
                tempVal    = action & " " & objectName
                currentVal = ""
                if ( isArray(objectFieldsUpDateValues) ) then
                    tempVal = tempVal & " SET "
                    if ( isArray(objectFieldsUpDateNames) ) then
                        for i = lBound(objectFieldsUpDateNames) to uBound(objectFieldsUpDateNames)
                            if (  len( objectFieldsUpDateValues(i) )  >  0  ) then
                                if ( len(currentVal) > 0 ) then
                                    currentVal = currentVal & ", "
                                end if
                                currentVal = currentVal  &  formatDBselector( objectFieldsUpDateNames(i) )  &  " = "  &  objectFieldsUpDateValuesWithFormat(i)
                            end if
                        next
                    end if
                end if

                tempVal = tempVal & currentVal & whereSQLstring & ";"

            case "DELETE"
                tempVal = action & " FROM " & objectName & whereSQLstring & ";"

            case "SP"
                tempVal = "EXEC " & DBobjectName & " "
                if ( isArray(objectFieldsValuesWithDBformat) ) then
                    for i = lBound(objectFieldsValuesWithDBformat) to uBound(objectFieldsValuesWithDBformat)
                        if ( isArray(objectFieldsNames) ) then
                            tempVal = tempVal & "@" & objectFieldsNames(i) & " = "
                        end if
                        tempVal = tempVal & objectFieldsValuesWithDBformat(i)
                        if ( i < uBound(objectFieldsValuesWithDBformat) ) then
                            tempVal = tempVal & ", "
                        end if
                    next
                end if
                tempVal = tempVal & ";"

            case else
        end select

        buildSQLstring = tempVal

    end function
%>