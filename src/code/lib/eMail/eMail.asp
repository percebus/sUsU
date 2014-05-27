<%
    function eMail (typeOfMail, SMTP, port, USR, PWD, from, fromName, addresses, CCs, BCCs, escString, isHTML, subject, msg, attachment)

        select case typeOfMail
		    case "Persits"
                set mail          = server.createObject("Persits.mailSender")
                    mail.host     = SMTP 'enter valid SMTP host
                    mail.userName = USR
                    mail.password = PWD
                    mail.from     = from
                    mail.fromName = fromName
                if ( isArray(addresses) ) then
                    for i = lBound(addresses) to uBound(addresses)
                        mail.AddAddress addresses(i)
                    next
                end if			
                if ( isArray(CCs) ) then
                    for i = lBound(CCs) to uBound(CCs)
                        mail.AddCC CCs(i)
                    next
                end if
                if ( isArray(BCCs) ) then
                    for i = lBound(BCCs) to uBound(BCCs)
                        mail.AddBCC BCCs(i)
                    next
                end if
                mail.subject       = subject
                mail.isHTML        = isHTML
                mail.body          = msg 
'               mail.addAttachment = attachment
                mail.send
            case "CDONTS"
                set mail = server.createObject("CDONTS.newMail")
                    mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/sendusing")      = 2 
                    mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")     = SMTP 'Name or IP of remote SMTP server
                    mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port 'Server port
                    mail.configuration.fields.upDate
                    mail.from = from
                if ( isArray(addresses) ) then
                    sendTo = ""
                    for i = lBound(addresses) to uBound(addresses)
                        sendTo = sendTo & ";" & addresses(i)
                    next
                    mail.to = sendTo
                end if
                if ( isArray(CCs) ) then
                    sendTo = ""
                    for i = lBound(CCs) to uBound(CCs)
                        sendTo = sendTo & ";" & CCs(i)
                    next
                    mail.CC = sendTo
                end if
                if ( isArray(BCCs) ) then
                    sendTo = ""
                    for i = lBound(BCCs) to uBound(BCCs)
                        sendTo = sendTo & ";" & BCCs(i)
                    next
                    mail.BCC = sendTo
                end if
                mail.subject    = subject
                mail.bodyFormat = 0
                mail.mailFormat = 0
                mail.body       = msg
               'mail.attachment = attachment
                mail.send
            case "CDO"
                set mail = server.createObject("CDO.message")
                   'mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/sendusing")  = 2
                   'mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP 'Name or IP of remote SMTP server
                   'mail.configuration.fields.item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") 	= port 'Server port
                   'mail.configuration.fields.upDate
                    mail.from = from
                if ( isArray(addresses) ) then
                    sendTo = ""
                    for i = lBound(addresses) to uBound(addresses)
                        sendTo = sendTo & ";" & addresses(i)
                    next
                    mail.to = sendTo
                end if
                if ( isArray(CCs) ) then
                    sendTo = ""
                    for i = lBound(CCs) to uBound(CCs)
                        sendTo = sendTo & ";" & CCs(i)
                    next
                    mail.CC = sendTo
                end if
                if ( isArray(BCCs) ) then
                    sendTo = ""
                    for i = lBound(BCCs) to uBound(BCCs)
                        sendTo = sendTo & ";" & BCCs(i)
                    next
                    mail.BCC = sendTo
                end if
                mail.subject = subject
                if ( isHTML ) then 
                    mail.HTMLbody = msg
                else
                    mail.textBody = msg
                end if
               'mail.attachment = attachment
                mail.send
        end select
        set mail = nothing
        if (err) then
            eMail = false
        else
            eMail = true 
        end if
	end function
%>