<!--#include file="src/inc.asp"-->
<%
    call varIsRequired(  sessionGET( appName & "usr"   ),  ( appFolder & "login.asp" )  )
    call varIsRequired(  sessionGET( appName & "usrID" ),  ( appFolder & "login.asp" )  )

'***** CFG *************************************************
    view         = "personas"

    displayVars  = array( "ShortName"   , "FamilyName"        , "MiddleName"      , "GivenName", "TypeOfPersonName" )
    displayVarsL = array( "Nombre Corto", "Apellido | Empresa", "Apellido Materno", "Nombre"   ,"Tipo de Persona"   )

    tableName    = "Q_Personas"
    qVars        = array(                "UserAccountID"    )
    qVals        = array(  sessionGET( appName & "usrID" )  )
    qValsT       = array(                    "num"          )
    filterVars   = "*"
'***********************************************************
%>
<!--#include file="src/header.asp"-->
        <table border="1" cellpadding="5" cellspacing="5">
            <thead>
                <tr>
                    <th colspan="<%=( uBound(displayVarsL) +3 )%>">Seleccione una Persona | Identidad</th>
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

                <tr><td colspan="<%=( uBound(displayVars) +3 )%>" align="center">[ <a href="setPersona.asp">Agregar Identidad</a> ]</td></tr>

            </thead>
            <tbody>
<%
    oRS.source         = buildSQLstring( "SELECT", DBMS, tableName, qVars, qVals, qValsT, true, "", "", "", filterVars, "", "", "", "" )
    oRS.cursorType     = 0
    oRS.cursorLocation = 2
    oRS.lockType       = 1
    oRS.open()
        do while not(oRS.EOF)

%>
                <tr>
                    <td nowrap="true" align="center">[ <a href="setPersona.asp?ID=<%=oRS.fields.item("ID").value%>">+</a> ]</td>
<%
            for i = lBound(displayVars) to uBound(displayVars)
%>
                    <td align="center"><a href="src/sessionPUT.asp?vars=persona,personaID,personaFamilyName,personaMiddleName,personaGivenName,personaTypeOf&vals=<%=oRS.fields.item("Display").value%>,<%=oRS.fields.item("ID").value%>,<%=oRS.fields.item("FamilyName").value%>,<%=oRS.fields.item("MiddleName").value%>,<%=oRS.fields.item("GivenName").value%>,<%=oRS.fields.item("TypeOfPersonName").value%>&URL=<%=appFolder%>user.asp"><%=oRS.fields.item( displayVars(i) ).value%></a></td>
<%
            next
%>
                    <td align="center">[ <a href="src/sessionPUT.asp?vars=persona,personaID,personaFamilyName,personaMiddleName,personaGivenName,personaTypeOf&vals=<%=oRS.fields.item("Display").value%>,<%=oRS.fields.item("ID").value%>,<%=oRS.fields.item("FamilyName").value%>,<%=oRS.fields.item("MiddleName").value%>,<%=oRS.fields.item("GivenName").value%>,<%=oRS.fields.item("TypeOfPersonName").value%>&URL=<%=appFolder%>user.asp"> > </a> ]</td>
                </tr>
<%
            oRS.moveNext
        loop
    oRS.close()
%>
            </tbody>
        </table>

<!--#include file="src/footer.asp"-->