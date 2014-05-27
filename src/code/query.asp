<%
    query = ""

    CSSfolder = setDefault( request("CSSfolder"), defaultCSSfolder )
    CSS       = setDefault( request("CSS"      ), defaultCSS       )

    call setQueryVarTextType( accessFrom, "accessFrom" )

    usr = session("usr")
    call setQueryVarTextType( usr, "usr" )

    age    = session("age")
    isMale = session("isMale")
    call setQueryVarNumericType( age   , "age"    )
    call setQueryVarBoolType   ( isMale, "isMale" )

    call setQueryVarUIDType( section, "section" )
    call setQueryVarUIDType( grade  , "grade"   )
    call setQueryVarUIDType( group  , "group"   )

    call setQueryVarUIDType( quiz           , "quiz"            )
    call setQueryVarUIDType( quizApplication, "quizApplication" )
    call setQueryVarUIDType( quizUID        , "quizUID"         )
    call setQueryVarUIDType( part           , "part"            )
%>