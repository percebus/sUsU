<%
    class DB2SQL
    '************** Properties **************
        private pDBMS
        private pDBo
        private pOperation

        private pCriteriaVars
        private pCriteriaOperators
        private pCriteriaValues
        private pCriteriaValuesTypes
        private pCriteriaOR

        private pEditVars
        private pEditValues
        private pEditValuesTypes

        private pFilterVars
        private pFilterAliases
        private pLimitRows

        private pOrderByVars
        private pOrderByBarsASC
	
    '*********** Event Handlers *************
        private sub class_Initialize()
            pDBMS                = "MS Access"

            pDBo                 = ""
            pOperation           = "SELECT"

            pCriteriaVars        = ""
            pCriteriaOperators   = ""
            pCriteriaValues      = ""
            pCriteriaValuesTypes = ""
            pCriteriaOR          = false

            pEditVars            = ""
            pEditValues          = ""
            pEditValuesTypes     = ""

            pFilterVars          = "*"
            pFilterAliases       = ""
            pLimitRows           = ""

            pOrderByVars         = ""
            pOrderByBarsASC      = true
        end sub

        private sub class_Terminate()
            
        end sub

	
'************ property let **************
        public property let DBMS(theText)
            pDBMS = theText
        end property

        public property let DBo(theText)
            pDBo = theText
        end property

        public property let operation(theText)
            pOperation = theText
        end property


        public property let qVars(textOrArray)
            pCriteriaVars = textOrArray
        end property

        public property let qVals(textOrArray)
            pCriteriaValues = textOrArray
        end property

        public property let qValsT(textOrArray)
            pCriteriaValuesTypes = textOrArray
        end property

        public property let qOR(theBool)
            pCriteriaOR = theBool
        end property


        public property let eVars(textOrArray)
            pEditVars = textOrArray
        end property

        public property let eVals(textOrArray)
            pEditValues = textOrArray
        end property

        public property let eValsT(textOrArray)
            pEditValuesTypes = textOrArray
        end property


        public property let fVars(textOrArray)
            pFilterVars = textOrArray
        end property

        public property let fAliases(textOrArray)
            pFilterAliases = textOrArray
        end property

        public property let limitRows(theNumber)
            pLimitRows = theNumber
        end property


        public property let orderBy(textOrArray)
            pOrderByVars = textOrArray
        end property

        public property let orderByASC(theBool)
            pOrderByVarsASC = theBool
        end property


'************ property get **************
        public property get SQLstring()
            SQLstring = buildSQLstring( pOperation, pDBMS, pDBo, pCriteriaVars, pCriteriaValues, pCriteriaValuesTypes, pCriteriaOR, pEditVars, pEditValues, pEditValuesTypes, pFilterVars, pFilterValues, pLimitRows, pOrderVarsBy, pOrderVarsByASC )
        end property

	end class
%>