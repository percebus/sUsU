<%
    function getFirstFileURL( filesPath, extensions, preferredName )
        preferredName = setDefault( preferredName, "" )

        if not(  right( filesPath, 1 )  =  "/"  ) then
            filesPath = filesPath & "/"
        end if

        set fs = server.createObject("scripting.fileSystemObject")
        set f  = fs.getFolder( server.mapPath(filesPath) )

                 fileWithExtensions = ""
        preferredFileWithExtensions = ""
        for each fileFound in f.files
            if not( len(preferredFileWithExtensions) > 0 ) then
                iExtension = -1
                iExtension = iInArray( fs.getExtensionName(fileFound.name), extensions )
                if not( iExtension < 0 ) then
                    if ( fileFound.name = preferredName ) then
                        preferredFileWithExtensions = filesPath & fileFound.name
                                 fileWithExtensions = preferredFileWithExtensions
                    else
                        if not( len(fileWithExtensions) > 0 ) then
                             fileWithExtensions = filesPath & fileFound.name
                        end if
                    end if
                end if
            end if
        next

        set f  = nothing
        set fs = nothing

        getFirstFileURL = setDefault(  setDefault( preferredFileWithExtensions, fileWithExtensions ),  defaultBidPic  )

    end function
%>