<%
    function buildForm( DOMIDs, DOMnames, displays, HTMLtypes, HTMLtypesDisabled, useMultiples, sizes, regExps, data, defaultValues, DOMclasses, CSSs, tabIndexes, disableds, JScripts, optionsHeaders, optionsSeparators, optionsRepeaters, optionsFooters, hRefs, targets )
        formSize = 0
        if ( isArray(DOMIDs) ) then
            formSize = uBound(DOMIDs)
        elseif ( isArray(DOMnames) ) then
            formSize = uBound(DOMnames)
        elseif ( isArray(displays) ) then
            formSize = uBound(displays)
        end if

        for iID = 0 to formSize
            currentID   = ""
            currentName = ""
            if ( isArray(DOMIDs) ) then
                currentID   = DOMIDs(iID)
            end if
            if ( isArray(DOMnames) ) then
                currentName = DOMnames(iID)
            end if
            if     not( len(currentID)   > 0 ) then
                currentID   = currentName
            elseif not( len(currentName) > 0 ) then
                currentName = currentID
            end if

'           currentDisplay  = setDefault( displays(iID), currentName  )
            currentDisplay  = setDefault(  iArray2text( displays, iID ),  currentID  )

            disabled        = ""
            currentDisabled = setDefault(  iArray2text( disableds, iID ),  false     )
            if (currentDisabled) then
                disabled    = "disabled=""true"""
                currentType = setDefault(  iArray2text( HTMLtypesDisabled, iID ),  iArray2text( HTMLtypes, iID )  )
            else
                currentType = iArray2text( HTMLtypes, iID )
            end if

            mainType = currentType
            select case currentType
                case "list"
                    mainType = "select"
                case "combo"
                    mainType = "select"
                case "checkbox"
                    mainType = "text"
                case "radio"
                    mainType = "options"
            end select

            currentSize         = iArray2text( sizes        , iID )
            currentRegexp       = iArray2text( regExps      , iID )

            currentDefaultValue = iArray2text( defaultValues, iID )
            currentSize         = iArray2text( sizes        , iID )

            currentData         = data(iID)

            currentClass        = iArray2text( DOMclasses   , iID )
            currentCSS          = iArray2text( CSSs         , iID )

            currentDisabled = false

            currentJScripts = iArray2text( JScripts, iID )

            currentHRef     = setDefault( iArray2text( hRefs  , iID ), "#"     )
            currentTarget   = setDefault( iArray2text( targets, iID ), "_self" )

%>
            <tr>
                <td>
                    <a name="field_<%=currentName%>" href="<%=currentHRef%>" target="<%=currentTarget%>"><%=currentDisplay%></a></td>
                <td>
<%

            select case mainType
                case "label"
%>
                    <%=currentDefaultValue%>
<%
                case "select" ' or list
                    multiple = ""
                    select case currentType
                        case "list"
                            currentSize = 5
                            currentMultiple = iArray2text( useMultiples, iID )
                            if ( cBool(currentMultiple) ) then
                                multiple = "multiple=""true"""
                            end if
                        case "combo"
                            currentSize = 1
                    end select
%>
                        <select id="<%=currentID%>" name="<%=currentName%>" class="<%=currentClass%>" <%=multiple%> size="<%=currentSize%>" style="<%=currentCSS%>" tabindex="<%=tabIndexes%>" <%=disabled%> <%=currentJScripts%>>
                            <option value=""></option>
<%
                    if ( isArray(currentData) ) then
                        for iCurrentData = lBound(currentData,2) to uBound(currentData,2)
                            selected = ""
                            if (  trim(currentDefaultValue)  =  trim( currentData(0,iCurrentData) )  ) then
                                selected = "selected=""true"""
                            end if
									
%>
                            <option value="<%=currentData(0,iCurrentData)%>" <%=selected%>><%=currentData(1,iCurrentData)%></option>
<%
                        next
                    end if
%>
                        </select>
<%
                case "options"
                    currentHeader    = iArray2text( optionsHeaders   , iID )
                    currentSeparator = iArray2text( optionsSeparators, iID )
                    currentRepeater  = iArray2text( optionsRepeaters , iID )
                    currentFooter    = iArray2text( optionsFooter    , iID )
%>
                    <%=currentHeader%>
<%
                    if ( isArray(currentData) ) then
                        for iCurrentData = lBound(currentData,2) to uBound(currentData,2)
                            checked = ""
                            if (  trim( absInt(currentDefaultValue) )  =  trim( currentData(0,iCurrentData) )  ) then
                                checked = "checked=""true"""
                            end if
									
%>
                        <label>
                            <input id="<%=currentID%>.<%=currentData(0,iCurrentData)%>" name="<%=currentName%>" class="<%=currentClass%>" type="<%=currentType%>" style="<%=currentCSS%>" tabindex="<%=tabIndexes%>" value="<%=currentData(0,iCurrentData)%>" <%=checked%> <%=disabled%> <%=currentJScripts%>/>
                            <%=currentSeparator%>
                            <%=currentData(1,iCurrentData)%>
                        </label>
<%
                            if ( iCurrentData < uBound(currentData,2) ) then
%>
                    <%=currentRepeater%>
<%
                            end if
                        next
                    end if
%>
                    <%=currentFooter%>
<%
                case "textarea"
                    currentSize = setDefault( currentSize, 5 )
%>
                        <textarea id="<%=currentID%>" name="<%=currentName%>" class="<%=currentClass%>" style="<%=currentCSS%>" tabindex="<%=tabIndexes%>" cols="<%=currentSize%>" <%=disabled%> <%=currentJScripts%>><%=currentDefaultValue%></textarea>
<%
                case else
                    currentSize = setDefault( currentSize, 30 )
                    checked     = ""
                    select case currentType
                        case "file"
                            accept = ""
                            MIMEs  = ""
'                           if ( isArray(data) ) then
'                               for i = lBound(data) to uBound(data)
'                                   MIMEs = MIMEs & data(iID)
'                                   if ( i < uBound(data) ) then
'                                       MIMEs = MIMEs & ", "
'                                   end if
'                               next
'                               accept = "accept=""" & MIMEs & """"
'                           end if
                        case "checkbox"
                            if ( len(currentDefaultValue) > 0 ) then
                                if ( cBool(currentDefaultValue) ) then
                                    checked = "checked=""true"""
                                end if
                            end if
                            currentDefaultValue = 1
                    end select
%>
                        <input id="<%=currentID%>" name="<%=currentName%>" class="<%=currentClass%>" style="<%=currentCSS%>" tabindex="<%=tabIndexes%>" type="<%=currentType%>" maxlength="<%=currentSize%>" value="<%=currentDefaultValue%>" <%=checked%> <%=accept%> <%=disabled%> <%=currentJScripts%> />
<%
            end select
%>
                </td>
            </tr>
<%
        next
    end function
%>