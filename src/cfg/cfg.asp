<%
    'option explicit ' If activated, declare vars at dim.asp
    response.expires         = -1
    response.expiresAbsolute = date - 2

    const debugFolder       = ""' "/www.laSalleChihuahua.edu.mx"
          appFolder         = debugFolder & "/JC/sUsU/"
         imgsFolder         = debugFolder & "/ILS/site/public/JC/sUsU/bids/"
    const appName           = "sUsU"
    const appDisplayName    = "sUsU"

    const useCookieSessions = false
          session.timeOut   = 20

    const defaultCSSfolder  = "src/cfg/css/default"
    const defaultCSS        = "src/cfg/css/default/all.css"

    const tries             = 5
          usrVars           = array( "ID", "Code", "Alias", "NickName", "SecurityLevelID", "SecurityLevelCode", "SecurityLevelName" )
          personaVars       = array( "ID", "FamilyName", "MiddleName", "GivenName", "TypeOf" )

          weekDays          = array( "", "L", "M", "Mi", "J", "V", "S", "D" )
    const serverGMT         = -7 ' Difference with GMT-0
          GMT               = -7 ' Difference between the server time and the desired time

          menuItems         = array( "", "Programa de Protección al comprador", "Programa de Reputación", "Pagos con PayPal" )
          menuItemsLinks    = array( "", "protection.asp"                     , "reputation.asp"        , "PayPal.asp"       )

          imgFormats        = array( "JPG", "PNG", "GIF" )
          bidsFolder        = "/ILS/site/public/JC/sUsU/bids/"
          defaultBidPic     = "src/cfg/img/default.png"

    const defaultAction     = "src/save.asp"

      dim ToF(1,1)
          ToF(0,0) = 1
          ToF(1,0) = "si"
          ToF(0,1) = 0
          ToF(1,1) = "no"

          highlightColor    = "#FFFF99"
    const imgAlphaInitial   = 100
    const imgAlphaFinal     =  50
          URLdecodes        = array( "?",  "&" )
      dim URLencodes()
    reDim URLencodes( uBound(URLdecodes) )
          for iURLcode = lBound(URLdecodes) to uBound(URLdecodes)
              URLencodes(iURLcode) = server.URLencode( URLdecodes(iURLcode) )
          next

'***********************************************************************************
'**** DB ***************************************************************************
'***********************************************************************************

          defaultDBfields         = array( "ID" , "Display" )
          defaultDBfieldsOrder    = array(        "Display" )
          defaultDBfieldsOrderASC = array(          true    )

          emptyFieldMask          = "_-*/|\*_-"

      dim sexs(1,1)
          sexs(0,0) = 1
          sexs(1,0) = "Masculino"
          sexs(0,1) = 0
          sexs(1,1) = "Femenino"

      dim percents(1,10)
          for iPercent = 0 to 10
              percents(0,iPercent) =     ( iPercent * 10 )
              percents(1,iPercent) = cStr( iPercent * 10 ) & "%"
          next

          URLpersonas       = appFolder & "personas.asp"
          URLphones         = appFolder & "phones.asp"
          URLaddresses      = appFolder & "addresses.asp"
          URLsuggestions    = appFolder & "suggestions.asp?add="
          URLtitles         = appFolder & "titles.asp"

'***** ADO DB data-types ********************
    const adEmpty                        =   0
    const adError                        =  10

    const adBool                         =  11

    const adNumeric                      = 131
    const     adNumericVar               = 139
    const     adInt                      =   3
    const         adIntUnsigned          =  19
    const         adIntTiny              =  16
    const             adIntTinyUnsigned  =  17
    const         adIntSmall             =   2
    const             adIntSmallUnsigned =  18
    const         adIntBig               =  20
    const             adIntBigUnsigned   =  21
    const     adDec                      =  14
    const         adDecSingle            =   4
    const         adDecDouble            =   5
    const         adCurrency             =   6

    const adGUID                         =  72

    const adDate                         =   7
    const     adDateDB                   = 133
    const     adDateDBtime               = 134
    const         adDateDBtimestamp      = 135
    const     adDateFileTime             =  64

    const adChar                         = 129
    const     adCharVar                  = 200
    const         adCharVarLong          = 201
    const     adCharW                    = 130
    const         adCharWvar             = 202
    const             adCharWvarLong     = 203

    const adBstring                      =   8

    const adBinary                       = 128
    const     adBinaryVar                = 204
    const         adBinaryVarLong        = 205

    const adChapter                      = 136

    const adCOMdispatch                  =   9
    const adCOMunknown                   =  13

    const adUser                         = 132

    const adVar                          =  12
    const     adVarProp                  = 138


    typesNum     = array( adNumeric, adNumericVar, adInt, adIntUnsigned, adIntTiny, adIntTinyUnsigned, adIntSmall, adIntSmallUnsigned, adIntBig, adIntBigUnsigned, adDec, adDecSingle, adDecDouble, adCurrency, adGUID )
    typesDate    = array( adDate, adDateDB, adDateDBtime, adDateDBtimestamp, adDateFileTime )
    typesText    = array( adChar, adCharVar, adCharVarLong, adBstring, adCharW, adCharWvar, adCharWvarLong )
    typesBinary  = array( adBinary, adBinaryVar, adBinaryVarLong )
    typesVariant = array( adVar, adVarProp )
    typesCOM     = array( adCOMdisptach, adCOMunknown )

    typesVars    = array( adEmpty, adError, adBool    , typesNum, adGUID , typesDate, typesText, typesBinary, typesCOM, adUser , typesVariant, adChapter )
    typesVals    = array( "empty", "error", "bool"    , "num"   , "GUID" , "date"   , "text"   , "binary"   , "COM"   , "user" , "var"       , "chapter" )
    typesUIsR    = array( "label", "label", "checkbox", "label" , "label", "label"  , "label"  , "label"    , "label" , "label", "label"     , "label"   )
    typesUIsW    = array( "label", "label", "checkbox", "text"  , "label", "date"   , "text"   , "text"     , "text"  , "text" , "text"      , "text"    )

'***** DBMS *************************************************************************
      DBMS                     = "MS Access"
          tableAdminTables     = "Q_AdminTables"
          minDBadminSecurity   = 3
          tableFKs             = "QS_Relationships"
          tableFKsLocalObject  = "LocalObject"
          tableFKsLocalColumn  = "LocalColumn"
          tableFKsRemoteObject = "RemoteObject"
          tableFKsRemoteColumn = "RemoteColumn"
          tablePickPrefix      = "UI_"
          tableQueryPrefix     = "Q_"

          DB               = server.mapPath( "/ILS/site/private/JC/sUsUx.mdb" ) ' server.mapPath( appFolder & "src/DB.mdb" ) 'use server.mapPath(DB) if DB is file-based, ej: access, excel, dbase, paradox, etc.
          DBcopyToFolder   = server.mapPath( "/ILS/site/private/JC/"         )

    const DBconnector      = "Provider=Microsoft.Jet.OLEDB.4.0"
   'const DBconnector      = "Driver={Microsoft Access Driver (*.mdb)}" 'for ODBC connections

    const DBedit           = "Jet OLEDB:Database Password=123456789"
    const DBread           = "Jet OLEDB:Database Password=123456789"

          DBstring         = DBconnector & "; Data Source=" & DB & ";"
          DBeditConnection = DBstring & DBedit '>>> IF LEN<=0 THEN NO DB EDIT
          DBreadConnection = DBstring & DBread '>>> IF LEN<=0 THEN NO DB READ; Reenforces security, but "edit" can be used as well

'***** eMail **********************************************************************
    const eMailerCarrier   = "Persits"
    const eMailerLanguage  = "es-mx"
    const eMailerSMTP      = "smtp01.zabco.net"
    const eMailerPort      = 25
    const eMailerUSR       = ""
    const eMailerPWD       = ""
    const eMailerFromEMail = "no-repy@sUsU.com"
    const eMailerFromName  = "sUsU"
          eMailerAdmins    = array( "", "A01301149@ITESM.mx", "A00385674@ITESM.mx", "A00974436@ITESM.mx", "A00933994@ITESM.mx" )
%>