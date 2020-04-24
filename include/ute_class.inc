<%

'
'   Public Properties    (R = read, W = write)
'
'      DBName      R/W  Name of Database. For display purpose only.
'
'      HeadLine    R    Headline of page (e.g. to be used as <title>)
'      TableName   R    Name of current table
'
'      ReadOnly    W    display table in readonly mode                    (default: false)
'
'      ListTables  W    display toolbutton to list all tables             (default: true)
'      Filters     W    display toolbutton to define and activate filters (default: true)
'      Export      W    display toolbutton to export all data to CSV file (default: true)
'      SQL         W    display toolbutton to show current sql statement  (default: true)
'      Definitions W    display toolbutton to show field defintions       (default: true)
'
'      ImageDir    R/W  name of image directory, must end with "/"  (default: images/ )
'
'   Public Functions
'      Init    (sDSN)   must be called *before* any other HTML output
'      Draw    ()       writes complete HTML code
'      getHTML ()       returns complete HTML code
'
'---------------------------------------------------------------------------

Class clsUTE



	 '-----------------------------------------------------------------------
    ' Private Member Variables
    '
    Private m_DB            ' database connection object
    Private m_RS            ' recordset object
    Private m_RSForm        ' recordset object for form view
    
    Private m_sSQL          ' SQL statement being used to read data from db

    Private m_nMode         ' View mode: mdTable, mdForm, mdExport
    Private m_nFormMode     ' Form mode: mdEdit, mdInsert, mdDelete

    Private m_sDSN          ' odbc connect string
    Private m_sDBName       ' database name (for display purpose only)
    Private m_sTable        ' table name
    Private m_nPage         ' current page
    Private m_nPageSize     ' size of current page
    Private m_bSortFields       ' sort fields alphabetically (columns) ?
    Private m_bViewDefinitions  ' show field definitions ?
    Private m_bViewSQL              ' show sql statement ?
    Private m_bAutoPKDetection  ' detect primary key fields ?
    Private m_bSubmitted            ' form was submitted
    Private m_bTableLookUp		'Enable or Disable Lookups in the Table List

    Private m_bReadOnly     ' display table in readonly mode

    Private m_bListTables           ' display toolbutton to list all tables within db
    Private m_bFilters              ' display toolbutton to define and activate filters
    Private m_bShowExportLink   ' display toolbutton to export all data to CSV file
    Private m_bShowSQL      ' display toolbutton to show current sql statement
    Private m_bShowDefLink      ' display toolbutton to show field defintions

    Private m_PrimaryKeyFields()    ' array holding names of all primary key fields
    Private m_PrimaryKeyTypes() ' array holding types of all primary key fields
    Private m_StandardFields()  ' array holding names of all other fields
    Private m_StandardTypes()   ' array holding types of all other fields
    

    Private m_SortFields()      ' array holding names of the fields for the ORDER BY clause
    Private m_SortFieldsOrder() ' array holding the directions for the ORDER BY clause

    Private m_nNumberOfFilters  ' number of filters
    Private m_FilterFields()    ' array holding names of fields of the filters
    Private m_FilterCompares()  ' array holding the comparisons for the filters
    Private m_FilterValues()    ' array holding the values of the filters
    Private m_FilterCombines()  ' array holding the combinitions of the filters

    Private m_sHeadLine             ' Headline

    Private m_sUTEScript        ' name of UTE script file
    Private m_sIMAGEDir			' name of image directory
    
    Private m_sMainMenu

    Private m_nRecord               ' number of record to be edited/deleted
    Private m_ErrorField()
    Private m_ErrorMessage()
    
    Private  m_FieldsLookupTable() ' array holding SQL LookupTable
	Private  m_FieldsLabel()		  ' array holding the Fields Label
	Private  m_FieldsHTMLCode()		'array holding HTML Code like a link to other windows
	Private  m_FieldsRichEdit()	'array holding the fields that will be treated as RichEditors
	Private  m_FieldsDefaultValue() 'array holding the default values of the fields
	Private  m_FieldsHookBrowse() 'array holding the fields that use Browse control
	Private  f_UserID
	Private  LoginAdmin



    '-----------------------------------------------------------------------
    ' Property Functions
    '


    ' ---- HeadLine (read) ----
    Property Get HeadLine()
        HeadLine = m_sHeadLine
    End Property

    ' ---- DBName (read/write) ----
    Property Let DBName(s)
        m_sDBName = s
    End Property

    Property Get DBName()
        DBName = m_sDBName
    End Property

    ' ---- TableName (read) ----
    Property Get TableName()
        TableName = m_sTable
    End Property

    ' ---- ReadOnly (write) ----
    Property Let ReadOnly(b)
        m_bReadOnly = b
    End Property

    ' ---- ListTables (write) ----
    Property Let ListTables(b)
        m_bListTables = b
    End Property

    ' ---- Definitions (write) ----
    Property Let Definitions(b)
        m_bShowDefLink = b
    End Property

    ' ---- SQL (write) ----
    Property Let SQL(b)
        m_bShowSQL = b
    End Property
    
    ' ---- Filter (write) ----
    Property Let Filters(b)
        m_bFilters = b
    End Property

    ' ---- Export (write) ----
    Property Let Export(b)
        m_bShowExportLink = b
    End Property

    ' ---- ImageDir (read/write)----
    Property Let ImageDir(s)
        m_sIMAGEDir = s
    End Property

    Property Get ImageDir()
        ImageDir = m_sIMAGEDir
    End Property
    
    '----- Main Menu link page
    Property Let MainMenu(s)
        m_sMainMenu = s
    End Property    

	'------ Disable or Enable Lookups in the Table List
	Property Let TableLookup(b)
		m_bTableLookUp = b
	End Property

    '-----------------------------------------------------------------------
    ' Private Member Functions
    '

    ''----------------------------------------------------------------------
    '' Name:     Class_Initialize
    ''           ================
    '' 
    '' Constructor.
    ''
    ''----------------------------------------------------------------------
    Private Sub Class_Initialize()

		
		f_UserID=Session("f_UserID")	
		LoginAdmin=Cbool(Session("LoginAdmin"))
	
		
        Set m_DB           = Server.CreateObject("ADODB.Connection")
        Set m_RS           = Server.CreateObject("ADODB.Recordset")
        

        m_nMode            = DEF_MODE
        m_nFormMode        = DEF_FORM_MODE

        m_sDSN             = ""
        m_sDBName          = ""
        m_sTable           = ""
        m_sSQL             = ""
        m_nPage            = DEF_PAGE
        m_nPageSize        = DEF_PAGE_SIZE
        m_bSortFields      = DEF_SORT_FIELDS
        m_bViewDefinitions = DEF_VIEW_DEFINITIONS
        m_bViewSQL         = DEF_VIEW_SQL
        m_bAutoPKDetection = DEF_PK_DETECTION
        m_bSubmitted       = False

        m_bReadOnly        = DEF_READONLY 

        m_bListTables      = DEF_LIST_TABLES
        m_bShowDefLink     = DEF_SHOW_DEF_LINK
        m_bShowExportLink  = DEF_EXPORT_LINK
        m_bShowSQL         = DEF_SQL_LINK
        m_bFilters         = DEF_FILTERS
        m_nNumberOfFilters = DEF_NUM_FILTER

        m_sUTEScript       = Request.ServerVariables("SCRIPT_NAME")
        m_sIMAGEDir        = DEF_IMAGE_DIR

        Redim m_PrimaryKeyFields(0)
        Redim m_PrimaryKeyTypes(0)
        Redim m_StandardFields(0)
        Redim m_StandardTypes(0)

        Redim m_SortFields(0)
        Redim m_SortFieldsOrder(0)

        Redim m_FilterFields(0)
        Redim m_FilterCompares(0)
        Redim m_FilterValues(0)
        Redim m_FilterCombines(0)

        Redim m_ErrorField(0)
        Redim m_ErrorMessage(0)
        
        Redim m_FieldsLookupTable(0)
        Redim m_FieldsLabel(0)
        Redim m_FieldsHTMLCode(0)
        Redim m_FieldsRichEdit(0)
        Redim m_FieldsDefaultValue(0)
        Redim m_FieldsHookBrowse(0)
        

    End Sub

    ''----------------------------------------------------------------------
    '' Name:     ClearAllTags
    ''           ===============
    '' 
    '' Clear all HTML Tags from a string
    ''
    '' Parameter: 
    ''			String
    ''
    '' return value:
    ''      string      
    ''
    ''----------------------------------------------------------------------
	Function clearAllTags(s)
		Dim re
		Set re = New RegExp
		re.Pattern = "(<[^>]*>)"
		re.Global = True
		re.IgnoreCase = True
		clearAllTags = re.Replace(s, "")
	End Function


    ''----------------------------------------------------------------------
    '' Name:     RTESafe
    ''           ===============
    '' 
    '' Replace improper characters for the Rich Editor text
    ''
    '' Parameter: 
    ''			String
    ''
    '' return value:
    ''      string      
    ''
    ''----------------------------------------------------------------------
	function RTESafe(strText)
		'returns safe code for preloading in the RTE
		dim tmpString
		
		tmpString = trim(strText)
		if len(tmpString) > 0 then 
			'convert all types of single quotes
			tmpString = replace(tmpString, chr(145), chr(39))
			tmpString = replace(tmpString, chr(146), chr(39))
			tmpString = replace(tmpString, "'", "&#39;")
			
			'convert all types of double quotes
			tmpString = replace(tmpString, chr(147), chr(34))
			tmpString = replace(tmpString, chr(148), chr(34))
		'	tmpString = replace(tmpString, """", "\""")
			
			'replace carriage returns & line feeds
			tmpString = replace(tmpString, chr(10), " ")
			tmpString = replace(tmpString, chr(13), " ")
		else
			tmpString = ""
		end if
		
		RTESafe = tmpString
	end function

    ''----------------------------------------------------------------------
    '' Name:     Class_Terminate
    ''           ===============
    '' 
    '' Destructor.
    ''
    ''----------------------------------------------------------------------
    Private Sub Class_Terminate()
        Redim m_PrimaryKeyFields(0)
        Redim m_StandardFields(0)
        Set m_RS = Nothing
        Set m_DB = Nothing
        
        if IsObject(m_RSForm) then
            Set m_RSForm = Nothing
        end if
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     SetURLParameter
    ''           ===============
    '' 
    '' Sets a given parameter to a URL parameter string. If the parameter is
    '' already present in the URL string it will be updated, otherwise it will
    '' simply be added.
    ''
    '' Parameter: 
    ''      sURL        URL string to set/update the parameter in
    ''      sName       name of parameter to be set
    ''      sValue      value of parameter to be set
    ''
    '' return value:
    ''      string      new URL parameter string
    ''
    ''----------------------------------------------------------------------
    Private Function SetURLParameter (sURL, sName, sValue)
        Dim iPos
        Dim sLink, sReturn, sLeft, sRight

        sLink = sURL
        sReturn = ""

        if InStr(sLink, sName & "=") <> 0 then
            ' update exisiting parameter
            iPos = InStr(sLink, sName & "=")
            sLeft = Left(sLink, iPos+Len(sName))

            sRight = Right(sLink, Len(sLink) - (iPos + Len(sName)))
            if InStr(sRight, "&") <> 0 then
                ' at least one following parameter
                iPos = InStr(sRight, "&")
                sRight = Right(sRight, Len(sRight)-iPos+1)

                sReturn = sReturn & sLeft & sValue & sRight
            else
                ' no following parameter
                sReturn = sReturn & sLeft & sValue
            end if
        else
            ' add parameter
            if Len(sLink) <> 0 then
                sReturn = sReturn & sLink & "&" & sName & "=" & sValue
            else
                sReturn = sReturn & sName & "=" & sValue
            end if
        end if

        SetURLParameter = sReturn
    End Function
    
    
    
    ''----------------------------------------------------------------------
    '' Name:     AddHookBrowse
    ''           =========
    '' 
    '' Marks a Field for AddHookBrowse. 
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName
    ''				ute.AddHookBrowse("Field Name")
    '' return value:
    '' 
    ''
    ''----------------------------------------------------------------------
    
	Public Sub AddHookBrowse (Params)
	
		Redim Preserve m_FieldsHookBrowse(UBound(m_FieldsHookBrowse)+1)
		m_FieldsHookBrowse(UBound(m_FieldsHookBrowse))= Params
	
	End Sub    
	
    ''----------------------------------------------------------------------
    '' Name:     AddRichEdit
    ''           =========
    '' 
    '' Marks a Field for RichEdit. 
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName
    ''				ute.AddRichEdit("Field Name")
    '' return value:
    '' 
    ''
    ''----------------------------------------------------------------------
    
	Public Sub AddRichEdit (Params)
	
		Redim Preserve m_FieldsRichEdit(UBound(m_FieldsRichEdit)+1)
		m_FieldsRichEdit(UBound(m_FieldsRichEdit))= Params
	
	End Sub    	
    

    ''----------------------------------------------------------------------
    '' Name:     AddLookUpTable
    ''           =========
    '' 
    '' Inserts a LookUpTable for a Field. It can be a Value Lookup or SQL Query
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName=Lookup(value1[|label1];value2[|label2];value3[|label3])
	''				FieldName=Lookup(1;2;3)  
	''				FieldName=Lookup(1|good;2|very good;3|excelent)
	''				ute.AddLookUpTable("Category=sql(select * from categories  where categories.id=v_Field{MasterTable_ReplaceField_with_Value:SlaveTable_ShowField})")    ''
    '' return value:
    '' 
    ''
    ''----------------------------------------------------------------------
    
	Public Sub AddLookUpTable (Params)
	
		Redim Preserve m_FieldsLookupTable(UBound(m_FieldsLookupTable)+1)
		m_FieldsLookupTable(UBound(m_FieldsLookupTable))= Params
	
	End Sub
	
	
    ''----------------------------------------------------------------------
    '' Name:     AddHTMLCode 
    ''           =========
    '' 
    '' adds an HTML Code at the right of the  Field
    '' 
    ''
    '' Parameter: 
    ''      Params  
	''				ute.AddHTMLCode("Logo=<a href=OpenDownloadDir.asp> Dir of Images Logos </a>")    ''
    '' return value:
    '' 
    ''
    ''----------------------------------------------------------------------
    
	Public Sub AddHTMLCode(Params)
	
		Redim Preserve m_FieldsHTMLCode(UBound(m_FieldsHTMLCode)+1)
		m_FieldsHTMLCode(UBound(m_FieldsHTMLCode))= Params
	
	End Sub	
	
		''--------------------------------------------------------------------------
		'' Name:     LookUpTable
		''           =================
		'' 
		'' Returns a value if the FieldName has a LookupTable options like Drop Down Menu or SQL Code
		'' 
		''
		'' Parameter: 
		''      The name of the field
		''
		'' return value: Number
		''     NO_LOOKUP_TABLE	 = 1
		''	   VALUES_TABLE		 = 2
		''     SQL_QUERY		 = 3
		''
		''--------------------------------------------------------------------------
		Private Function LookupTable ( FieldName )
		dim i,substr

		for i=1 to UBound(m_FieldsLookupTable)
			substr=Mid(m_FieldsLookupTable(i),1,InStr(1,ucase(m_FieldsLookupTable(i)),"=")-1 )
			if   ucase(substr) = ucase(FieldName) then
				if InStr(1,ucase(m_FieldsLookupTable(i)),"SQL") > 0 then
					LookupTable=SQL_QUERY
				elseif InStr(1,ucase(m_FieldsLookupTable(i)),"REPLACE") > 0 then
					LookupTable=REPLACE_VALUES
				else
					LookupTable=VALUES_TABLE
				end if
				exit function
			end if

		next

		LookupTable=NO_LOOKUP_TABLE

		end Function

		''--------------------------------------------------------------------------
		'' Name:     LookUpTableGetParams
		''           =================
		'' 
		'' Returns the string of the parameters
		'' 
		''
		'' Parameter: 
		''      The name of the field
		''
		'' return value: String
		''     
		''
		''--------------------------------------------------------------------------
		Private Function LookUpTableGetParams ( FieldName )
		dim i

		for i=1 to UBound(m_FieldsLookupTable)
			if InStr(1,m_FieldsLookupTable(i),FieldName) > 0 then
				LookUpTableGetParams=m_FieldsLookupTable(i)
				exit function
			end if

		next

		End Function	
	
    ''----------------------------------------------------------------------
    '' Name:     AddFieldsLabel
    ''           =========
    '' 
    '' Inserts a Label for the Field. 
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName=Label text	
    ''
    '' return value:
    ''		The label name or else ""
    ''
    ''----------------------------------------------------------------------	
	Public Sub AddFieldsLabel (Params)
	
		Redim Preserve m_FieldsLabel(UBound(m_FieldsLabel)+1)
		m_FieldsLabel(UBound(m_FieldsLabel))= Params	
	
	End Sub
	
    ''----------------------------------------------------------------------
    '' Name:     AddDefaultValue
    ''           =========
    '' 
    '' Inserts a defualt value for the Field. 
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName=Default Value	
    ''
    '' return value:
    ''		Nothing
    ''
    ''----------------------------------------------------------------------	
	Public Sub AddDefaultValue (Params)
	
		Redim Preserve m_FieldsDefaultValue(UBound(m_FieldsDefaultValue)+1)
		m_FieldsDefaultValue(UBound(m_FieldsDefaultValue))= Params	
	
	End Sub	
	
  
    ''----------------------------------------------------------------------
    '' Name:     GetDefaultValue
    ''           =========
    '' 
    '' Inserts a Field name to search. 
    '' 
    ''
    '' Parameter: 
    ''      Params  FieldName to be searched
    ''
    '' return value:
    ''		if found then returns the Default Value else empty string
    ''
    ''----------------------------------------------------------------------	
	Public Function GetDefaultValue(Fieldname)
	
		dim i

		for i=1 to UBound(m_FieldsDefaultValue)
			if InStr(1,ucase(m_FieldsDefaultValue(i)),ucase(FieldName)) > 0 then
				GetDefaultValue=Mid(m_FieldsDefaultValue(i),InStr(1,m_FieldsDefaultValue(i),"=")+1,Len(m_FieldsDefaultValue(i)))
				exit function
			end if	
		next
		GetDefaultValue=""	
	
	End Function
	
	Public Function FieldHasALabel(Fieldname)
	
		dim i

		for i=1 to UBound(m_FieldsLabel)
			if InStr(1,ucase(m_FieldsLabel(i)),ucase(FieldName)) > 0 then
				FieldHasALabel=Mid(m_FieldsLabel(i),InStr(1,m_FieldsLabel(i),"=")+1,Len(m_FieldsLabel(i)))
				exit function
			end if	
		next
		FieldHasALabel=""	
	
	End Function

    ''----------------------------------------------------------------------
    '' Name:     buildLink
    ''           =========
    '' 
    '' Returns a string containing all UTE relevant URL parameters, such
    '' as tablename, page or primary key fields.
    ''
    '' Parameter: 
    ''      sCurrent    current URL string
    ''
    '' return value:
    ''      string      link
    ''
    ''----------------------------------------------------------------------
    Private Function buildLink (sCurrent)
        Dim i
        Dim sReturn, sDefs, sSQL, sSort, sSubmitted

        ' preserve current URL string
        sReturn = sCurrent

        sSQL = "0"
        if m_bViewSQL then sSQL = "1"
        sDefs = "0"
        if m_bViewDefinitions then sDefs = "1"
        sSort = "0"
        if m_bSortFields then sSort = "1"
        sSubmitted = "0"
        if m_bSubmitted then sSubmitted = "1"

        ' set all UTE URL params
        sReturn = SetURLParameter(sReturn, sParamTable,      m_sTable)
        sReturn = SetURLParameter(sReturn, sParamMode,       CStr(m_nMode))
        sReturn = SetURLParameter(sReturn, sParamFormMode,   CStr(m_nFormMode))
        sReturn = SetURLParameter(sReturn, sParamPage,       CStr(m_nPage))
        sReturn = SetURLParameter(sReturn, sParamPageSize,   CStr(m_nPageSize))
        sReturn = SetURLParameter(sReturn, sParamSQL,        sSQL)
        sReturn = SetURLParameter(sReturn, sParamDefs,       sDefs)
        sReturn = SetURLParameter(sReturn, sParamSortFields, sSort)
        sReturn = SetURLParameter(sReturn, sParamSubmitted,  sSubmitted)
        sReturn = SetURLParameter(sReturn, sParamRecord,     CStr(m_nRecord))

        ' add primary key fields
        for i = 1 to UBound(m_PrimaryKeyFields)
            sReturn = SetURLParameter(sReturn, sParamPKey & CStr(i), m_PrimaryKeyFields(i))
        next

        ' add sort fields
        for i = 1 to UBound(m_SortFields)
            sReturn = SetURLParameter(sReturn, sParamSort & CStr(i),    m_SortFields(i))
            sReturn = SetURLParameter(sReturn, sParamSortDir & CStr(i), m_SortFieldsOrder(i))
        next

        buildLink = sReturn

    End Function


    ''----------------------------------------------------------------------
    '' Name:     GetLink
    ''           =======
    '' 
    '' Adds the given parameter to a compelte UTE link. An UTE link includes
    '' all possible URL parameters and is used to switch table pages or open
    '' the record form.
    ''
    '' Parameter: 
    ''      sScript     name of script to be called, e.g. ute.asp
    ''      sCurrent    current link, if "" the function build a complete new link
    ''      sParam      name of parameter to be set
    ''      sValue      value to be set
    ''
    '' return value:
    ''      string      complete link
    ''
    ''----------------------------------------------------------------------
    Private Function GetLink ( sScript, sCurrent, sParam, sValue )

        Dim iPos
        Dim sReturn, sLeft, sRight

        ' check if we already have a complete UTE URL string ?
        if InStr(sCurrent, sParamMode) <> 0 then
            ' use current link
            iPos = InStr(sCurrent, "?")
            sReturn = "&" & Right(sCurrent, Len(sCurrent)-iPos)
        else
            ' build new link
            sReturn = "&" & buildLink(sCurrent)
        end if

        sReturn = SetURLParameter(sReturn, sParam, sValue)

        ' add script name, repleace leading "&" by "?"
        sReturn = sScript & "?" & Right(sReturn, Len(sReturn)-1)

        GetLink = sReturn
    End Function


    ''----------------------------------------------------------------------
    '' Name:     RemoveParameter
    ''           ===============
    '' 
    '' Removes parameter from given URL string.
    ''
    '' Parameter: 
    ''      sLink       string containing the link
    ''      sParam      parameter name
    ''
    '' return value:
    ''      string
    ''
    ''----------------------------------------------------------------------
    Private Function RemoveParameter ( sLink, sParam )
        Dim sLeft, sRight, sReturn
        Dim iPos
        sReturn = sLink

        while InStr(sReturn, sParam & "=") <> 0
            iPos   = InStr(sReturn, sParam & "=")
            sLeft  = Left(sReturn, iPos-1)
            iPos   = InStr(iPos, sReturn, "&")
            sRight = ""
            if iPos <> 0 then sRight = Right(sReturn, Len(sReturn)-iPos)
            sReturn = sLeft & sRight
        wend

        if Right(sReturn, 1) = "&" then sReturn = Left(sReturn, Len(sReturn)-1)

        RemoveParameter = sReturn
    End Function

    
    ''----------------------------------------------------------------------
    '' Name:     RemoveCountedParameters
    ''           =======================
    '' 
    '' Removes so called "counted parameters" like "pkey[n]" or "sort[n]"
    '' from the given link. The start counter defines the start value for [n].
    '' E.g.: sParam = sort, nStartCounter = 2
    '' -> removes all sort2, sort3, sort4, ... from the link
    ''
    '' Parameter: 
    ''      sLink       string containing the link
    ''      sParam      parameter name
    ''      nStarCount  start counter
    ''
    '' return value:
    ''      string
    ''
    ''----------------------------------------------------------------------
    Private Function RemoveCountedParameters ( sLink, sParam, nStartCount )
        Dim sReturn
        Dim n
        n = nStartCount
        sReturn = sLink

        while InStr(sReturn, sParam & CStr(n)) <> 0
            sReturn = RemoveParameter(sReturn, sParam & CStr(n))
            n = n + 1
        wend

        RemoveCountedParameters = sReturn
    End Function

    
    ''----------------------------------------------------------------------
    '' Name:     GetParameter
    ''           ============
    '' 
    '' Gets all parameters from URL and throw excaption if neccessary.
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub GetParameter()

        Dim i, j
        Dim sTemp
        Dim sError

        ' ---- ODBC connect string ----
        if m_sDSN = "" then
            '"Invalid ODBC Connection String"
            sError = STR_ERR_1001 
            err.Raise vbObjectError + 1001, "ute_table", sError
        end if

        ' ---- Tablename ----
        if Request.QueryString(sParamTable) <> "" then
            m_sTable = Request.QueryString(sParamTable)
        else
            m_nMode = MD_DATABASE
        end if


        ' ---- Tablename ----
        m_sTable = Request.QueryString(sParamTable)
        if m_bListTables then
            ' if no table set, display all tables within database
            if Request.QueryString(sParamTable) = "" then
                m_nMode = MD_DATABASE
            end if
        else
            ' if no table set throw error
            if Request.QueryString(sParamTable) = "" then
                '"Missing ""%1"" URL parameter."
                sError = Replace(STR_ERR_1002, "%1", sParamTable) 
                err.Raise vbObjectError + 1002, "ute_table", sError
            end if
        end if

        ' ---- Mode ----
        if Request.QueryString(sParamMode) <> "" then
            sTemp = Request.QueryString(sParamMode)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamMode) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nMode = CInt(sTemp)
            if (m_nMode < MD_DATABASE) or (m_nMode > MD_FILTER) then
                '"Invalid ""%1"" URL parameter. Must be ""1"", ""2"" or ""3""."
                sError = Replace(STR_ERR_1004, "%1", sParamMode)
                err.Raise vbObjectError + 1004, "ute_table", sError
            end if
        end if

        ' ---- FormMode ----
        if Request.QueryString(sParamFormMode) <> "" then
            sTemp = Request.QueryString(sParamFormMode)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamFormMode) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nFormMode = CInt(sTemp)
            if (m_nFormMode < 1) or (m_nFormMode > 3) then
                '"Invalid ""%1"" URL parameter. Must be ""1"", ""2"" or ""3""."
                sError = Replace(STR_ERR_1004, "%1", sParamFormMode)
                err.Raise vbObjectError + 1004, "ute_table", sError
            end if
        end if

        ' ---- Page ----
        if Request.QueryString(sParamPage) <> "" then
            sTemp = Request.QueryString(sParamPage)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamPage) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nPage = CInt(sTemp)
        end if

        ' ---- Page Size ----
        if Request.QueryString(sParamPageSize) <> "" then
            sTemp = Request.QueryString(sParamPageSize)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamPageSize) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nPageSize = CInt(sTemp)
        end if

        ' ---- Sort Fields Alphabetically ----
        if Request.QueryString(sParamSortFields) <> "" then
            sTemp = Request.QueryString(sParamSortFields)
            if (sTemp <> "0") and (sTemp <> "1") then
                '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                sError = Replace(STR_ERR_1005, "%1", sParamSortDir) 
                sError = Replace(sError, "%2", "0") 
                sError = Replace(sError, "%3", "1")
                err.Raise vbObjectError + 1005, "ute_table", sError
            end if
            m_bSortFields = (sTemp = "1")
        end if

        ' ---- View SQL Statement ----
        if Request.QueryString(sParamSQL) <> "" then
            sTemp = Request.QueryString(sParamSQL)
            if (sTemp <> "0") and (sTemp <> "1") then
                '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                sError = Replace(STR_ERR_1005, "%1", sParamSQL) 
                sError = Replace(sError, "%2", "0") 
                sError = Replace(sError, "%3", "1")
                err.Raise vbObjectError + 1005, "ute_table", sError
            end if
            m_bViewSQL = (sTemp = "1")
        end if

        ' ---- View Field Definitions ----
        if Request.QueryString(sParamDefs) <> "" then
            sTemp = Request.QueryString(sParamDefs)
            if (sTemp <> "0") and (sTemp <> "1") then
                '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                sError = Replace(STR_ERR_1005, "%1", sParamDefs) 
                sError = Replace(sError, "%2", "0") 
                sError = Replace(sError, "%3", "1")
                err.Raise vbObjectError + 1005, "ute_table", sError
            end if
            m_bViewDefinitions = (sTemp = "1")
        end if

        ' ---- Submitted ----
        if Request.QueryString(sParamSubmitted) <> "" then
            sTemp = Request.QueryString(sParamSubmitted)
            if (sTemp <> "0") and (sTemp <> "1") then
                '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                sError = Replace(STR_ERR_1005, "%1", sParamSubmitted) 
                sError = Replace(sError, "%2", "0") 
                sError = Replace(sError, "%3", "1")
                err.Raise vbObjectError + 1005, "ute_table", sError
            end if
            m_bSubmitted = (sTemp = "1")
        end if

        ' ---- Record ----
        if Request.QueryString(sParamRecord) <> "" then
            sTemp = Request.QueryString(sParamRecord)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamRecord) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nRecord = CInt(sTemp)
        end if


        ' ---- Primary Keys ----
        i = 1
        while Request.QueryString(sParamPKey & CStr(i)) <> ""
            ' switch off auto primary key detection 
            m_bAutoPKDetection = False
            ' "-1" will force SetPrimaryKeyFieldType to set the field type properly
            AddPrimaryKeyField Request.QueryString(sParamPKey & CStr(i)), -1
            i = i + 1
        wend


        ' ---- Sort Field ----
        ' This can be either "sort" (for compatebility purpose) or "sort[n]"
        if Request.QueryString(sParamSort) <> "" then
            AddSortField Request.QueryString(sParamSort)
            if Request.QueryString(sParamSortDir) <> "" then
                sTemp = LCase(Request.QueryString(sParamSortDir))
                if (sTemp <> SORT_ASC) and (sTemp <> SORT_DESC) then
                    '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                    sError = Replace(STR_ERR_1005, "%1", sParamSortDir) 
                    sError = Replace(sError, "%2", SORT_ASC) 
                    sError = Replace(sError, "%3", SORT_DESC)
                    err.Raise vbObjectError + 1005, "ute_table", sError
                end if
                AddSortOrder sTemp
            else
                ' default sort order
                AddSortOrder DEF_SORT_DIR
            end if
        else
            i = 1
            while Request.QueryString(sParamSort & CStr(i)) <> ""
                AddSortField Request.QueryString(sParamSort & CStr(i))

                if Request.QueryString(sParamSortDir & CStr(i)) <> "" then
                    sTemp = LCase(Request.QueryString(sParamSortDir & CStr(i)))
                    if (sTemp <> SORT_ASC) and (sTemp <> SORT_DESC) then
                        '"Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."
                        sError = Replace(STR_ERR_1005, "%1", sParamSortDir & CStr(i)) 
                        sError = Replace(sError, "%2", SORT_ASC) 
                        sError = Replace(sError, "%3", SORT_DESC)
                        err.Raise vbObjectError + 1005, "ute_table", sError
                    end if
                    AddSortOrder sTemp
                else
                    ' default sort order
                    AddSortOrder DEF_SORT_DIR
                end if

                i = i + 1
            wend
        end if 

        ' ---- Number of Filters ----
        if Request.QueryString(sParamFilterCount) <> "" then
            sTemp = Request.QueryString(sParamFilterCount)
            if not IsNumeric(sTemp) then
                '"Invalid ""%1"" URL parameter. Must be numeric."
                sError = Replace(STR_ERR_1003, "%1", sParamFilterCount) 
                err.Raise vbObjectError + 1003, "ute_table", sError
            end if
            m_nNumberOfFilters = CInt(sTemp)
        end if

        ' ---- General Filter Parameters ----
        if Request.QueryString(sParamFilterCompare & "1") <> "" then

            ' ---- Filter Compares ----
            i = 1
            while (Request.QueryString(sParamFilterCompare & CStr(i)) <> "") and (i <= m_nNumberOfFilters)
                Redim Preserve m_FilterCompares(UBound(m_FilterCompares)+1)
                m_FilterCompares(UBound(m_FilterCompares)) = Request.QueryString(sParamFilterCompare & CStr(i))
                i = i + 1
            wend

            ' ---- Filter Fields ----
            for j = 1 to i
                Redim Preserve m_FilterFields(UBound(m_FilterFields)+1)
                m_FilterFields(UBound(m_FilterFields)) = Request.QueryString(sParamFilterField & CStr(j))
            next

            ' ---- Filter Values ----
            for j = 1 to i
                Redim Preserve m_FilterValues(UBound(m_FilterValues)+1)
                m_FilterValues(UBound(m_FilterValues)) = Request.QueryString(sParamFilterValue & CStr(j))
            next

            ' ---- Filter Combines ----
            for j = 1 to i - 1
                Redim Preserve m_FilterCombines(UBound(m_FilterCombines)+1)
                m_FilterCombines(UBound(m_FilterCombines)) = Request.QueryString(sParamFilterCombine & CStr(j))
            next
            
        end if

    End Sub


    ''----------------------------------------------------------------------
    '' Name:     AddPrimaryKeyField
    ''           ==================
    '' 
    '' Add's a primary key field to the array
    ''
    '' Parameter:                
    ''      sField      name of field
    ''      nType       type of field
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub AddPrimaryKeyField (sField, nType)
        Redim Preserve m_PrimaryKeyFields(UBound(m_PrimaryKeyFields)+1)
        Redim Preserve m_PrimaryKeyTypes(UBound(m_PrimaryKeyFields))
        m_PrimaryKeyFields(UBound(m_PrimaryKeyFields)) = sField
        m_PrimaryKeyTypes(UBound(m_PrimaryKeyTypes)) = nType
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     AddStandardField
    ''           ================
    '' 
    '' Add's a "standard" field to the array
    ''
    '' Parameter: 
    ''      sField      name of field
    ''      nType       type of field
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub AddStandardField (sField, nType)
        Redim Preserve m_StandardFields(UBound(m_StandardFields)+1)
        Redim Preserve m_StandardTypes(UBound(m_StandardFields))
        m_StandardFields(UBound(m_StandardFields)) = sField
        m_StandardTypes(UBound(m_StandardTypes)) = nType
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     AddSortField
    ''           ============
    '' 
    '' Add's a field the table should be sorted after to the array
    ''
    '' Parameter:                
    ''      sField      name of field
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub AddSortField (sField)
        Redim Preserve m_SortFields(UBound(m_SortFields)+1)
        m_SortFields(UBound(m_SortFields)) = sField
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     AddSortOrder
    ''           ============
    '' 
    '' Add's the sort order of a field to the array
    ''
    '' Parameter:                
    ''      sField      name of field
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub AddSortOrder (sOrder)
        Redim Preserve m_SortFieldsOrder(UBound(m_SortFieldsOrder)+1)
        m_SortFieldsOrder(UBound(m_SortFieldsOrder)) = sOrder
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     SetPrimaryKeyFieldType
    ''           ======================
    '' 
    '' Sets the type of a primary key field to the types array if the
    '' current type is -1. This will be used if there are primary keys
    '' being set via URL. In this case we need to set the type afterwards.
    ''
    '' Parameter:                
    ''      sField      Name of the field
    ''      nType       Type of the field
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub SetPrimaryKeyFieldType ( sField, nType )
        Dim i
        for i = 1 to UBound(m_PrimaryKeyFields)
            if (m_PrimaryKeyFields(i) = sField) and (m_PrimaryKeyFields(i) = -1) then
                m_PrimaryKeyTypes(i) = nType
            end if
        next
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     IsKnownPrimaryKey
    ''           =================
    '' 
    '' Checks if the given field is already known as primary key
    ''
    '' Parameter: 
    ''      sField      Name of field
    ''
    '' return value:
    ''      boolean
    ''
    ''----------------------------------------------------------------------
    Private Function IsKnownPrimaryKey ( sField ) 
        Dim i
        Dim bReturn
        bReturn = False
        for i = 1 to UBound(m_PrimaryKeyFields)
            if m_PrimaryKeyFields(i) = sField then bReturn = True
        next
        IsKnownPrimaryKey = bReturn
    End Function


    ''----------------------------------------------------------------------
    '' Name:     IsPrimaryKey_inDBSchema
    ''           =======================
    '' 
    '' Checks if the given field is defined in db schema
    ''
    '' Parameter: 
    ''      sField      name of field
    ''
    '' return value:
    ''      boolean
    ''
    ''----------------------------------------------------------------------
    Private Function IsPrimaryKey_inDBSchema ( sField )
        Dim bReturn
        bReturn = False

        Dim rsSchema
        Set rsSchema = Server.CreateObject("ADODB.Recordset")
        rsSchema.CursorType = adOpenDynamic
        
        ' Getting the adSchemaPrimaryKeys will only supported by oledb 
        ' providers, not by simple ODBC connections. They will throw an
        ' error.
        on error resume next
        Set rsSchema = m_DB.openSchema(adSchemaPrimaryKeys)
        if Err = 0 then
            do while (not rsSchema.EOF) and (not bReturn)
                if LCase(rsSchema("TABLE_NAME")) = LCase(m_sTable) then
                    if LCase(rsSchema("COLUMN_NAME")) = LCase(sField) then
                        bReturn = True
                    end if
                end if
                rsSchema.MoveNext
            loop
            rsSchema.Close
        end if
        on error goto 0

        Set rsSchema = Nothing

        IsPrimaryKey_inDBSchema = bReturn
    End Function

    
    ''----------------------------------------------------------------------
    '' Name:     PrintSchema
    ''           ===========
    '' 
    '' For debug purpose only ! Prints the conntents of the given schema.
    ''
    '' Parameter: 
    ''
    '' return value:
    ''
    ''----------------------------------------------------------------------
    Private Sub PrintSchema

        Dim rsSchema, fld
        Set rsSchema = Server.CreateObject("ADODB.Recordset")
        rsSchema.CursorType = adOpenDynamic

        on error resume next
        Set rsSchema = m_DB.openSchema(adSchemaPrimaryKeys)
        'Set rsSchema = m_DB.openSchema(adSchemaIndexes)
        'Set rsSchema = m_DB.openSchema(adSchemaColumns)
        'Set rsSchema = m_DB.openSchema(adSchemaTables)
        'Set rsSchema = m_DB.openSchema(adSchemaProviderTypes)

        if err = 0 then
            while not rsSchema.EOF
                response.write "-----------------------------------------------------------------------<br>" & vbCrLf
                for each fld in rsSchema.Fields
                    response.write fld.name & ": " & fld.value & "<br>" & vbCrLf
                next
                rsSchema.MoveNext
            wend
        end if
        on error goto 0

        response.end

    End Sub

    
    ''----------------------------------------------------------------------
    '' Name:     SortFields
    ''           ==========
    '' 
    '' Sort given array ascending
    ''
    '' Parameter:
    ''      fields      array hoöding the fields to be sorted
    ''      types       array holding the types of the fields
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub SortFields ( ByRef fields, ByRef types )
        Dim pa
        Dim pb
        Dim temp
        ' standard bubble sort
        for pa = 1 to UBound(fields) - 1
            for pb = 1 to UBound(fields) - pa
                if fields(pb) > fields(pb + 1) then 
                    ' swap fields
                    temp = fields(pb)
                    fields(pb) = fields(pb + 1)
                    fields(pb + 1) = temp
                    ' swap types
                    temp = types(pb)
                    types(pb) = types(pb + 1)
                    types(pb + 1) = temp
                end if
            next
        next
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     AnalyzeTable
    ''           ============
    '' 
    '' Analyzing Table for Primary Key Fields and "normal" Fields.
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub AnalyzeTable()

        Dim fld
        Dim rsTemp
        Set rsTemp = Server.CreateObject("ADODB.Recordset")
        rsTemp.Open "[" & m_sTable & "]", m_DB, adOpenStatic, adLockReadOnly, adCmdTable

        for each fld in rsTemp.fields
						
            if m_bAutoPKDetection then
                ' check if field is in schema marked as primary key
                if IsPrimaryKey_inDBSchema(fld.name) then
                    if Not(IsKnownPrimaryKey(fld.name)) then 
                        AddPrimaryKeyField fld.name, fld.type
                    end if
                ' treat field as primary key field if the following flags are set:
                ' -> (KeyColumn) OR (Fixed and ((not Updateable) and (not UnknownUpdateable)))
                elseif ((fld.attributes and adFldKeyColumn) <> 0) or _
                    ( _
                    ((fld.attributes and adFldFixed) <> 0) and _
                    ( _
                        ((fld.attributes and adFldUpdatable) = 0) and _
                        ((fld.attributes and adFldUnknownUpdatable) = 0)) _
                    ) then
                    if Not(IsKnownPrimaryKey(fld.name)) then 
                        AddPrimaryKeyField fld.name, fld.type
                    end if
                ' this is no primary key field
                else
                    if Not(IsKnownPrimaryKey(fld.name)) then 
                        AddStandardField fld.name, fld.type
                    end if
                end if
            else
                ' the primary keys have been set via URL
                if IsKnownPrimaryKey(fld.name) then 
                    ' we need to set the type since we only got the name from the URL
                    SetPrimaryKeyFieldType fld.name, fld.type
                else
                    AddStandardField fld.name, fld.type
                end if
            end if
        next

        rsTemp.Close
        Set rsTemp = Nothing

        ' sort fields (ascending) in array if wanted
        if m_bSortFields then
            SortFields m_PrimaryKeyFields, m_PrimaryKeyTypes
            SortFields m_StandardFields, m_StandardTypes
        end if

    End Sub


    ''----------------------------------------------------------------------
    '' Name:     getPoweredBy
    ''           ============
    '' 
    '' Returns HTML code for "powered by UTE"
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      string
    ''
    ''----------------------------------------------------------------------
    Private Function getPoweredBy ()
        Dim sReturn
        sReturn = _
            "<a href=""" & sUTELink & """ target=""_blank"" class=""ute_link"" " & _
            "title=""" & sUTELink & """>" & sUTELongName & "</a>"
        sReturn = Replace(STR_POWERED_BY, "%1", sReturn)
        sReturn = Replace(sReturn,        "%2", sUTEVersion)
        getPoweredBy = "<div class=""ute_powered_by"">" & sReturn & "</div>"
    End Function


    ''--------------------------------------------------------------------------
    '' Name:     AddWhere
    ''           ========
    '' 
    '' Adds WHERE clause to SQL Statement
    ''
    '' Parameter: 
    ''      sName       name of field
    ''      nType       type of field
    ''      sValue      value of field, if empty the value is taken from the field object
    ''      sCompare    comparison like "=" or ">"
    ''      sCombine    combinition of clauses like "AND" or "OR"
    ''      bFirst      is the the first where clause ?
    ''
    '' return value:
    ''      string
    ''
    ''--------------------------------------------------------------------------
    Private Function AddWhere ( sName, nType,  sValue, sCompare, sCombine, bFirst )
        Dim sReturn, sSepChar

        sSepChar = ""
        select case nType
            case adBSTR, adVariant, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                sSepChar = "'"
            case adDate, adDBDate, adDBTime, adDBTimeStamp
                sSepChar = "'"
            case else
                sSepChar = ""
        end select

        if bFirst then
            sReturn = " WHERE "
        else
            sReturn = " " & sCombine & " "
        end if
        
        select case nType
            case adSingle, adDouble, adCurrency
                sValue = Replace(sValue, ",", ".")
            case adBSTR, adVariant, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                sValue = Replace(sValue, "'", "''")
        end select

        AddWhere = sReturn & sName & " " & sCompare & " " & sSepChar & sValue & sSepChar

    End Function


    ''----------------------------------------------------------------------
    '' Name:     getFieldType
    ''           ============
    '' 
    '' Returns the type of the given field
    ''
    '' Parameter: 
    ''      sName       Name of the field
    ''
    '' return value:
    ''      Inbteger
    ''
    ''----------------------------------------------------------------------
    Private Function getFieldType ( sName )
        Dim i, bFound
        Dim nReturn
        nReturn = 0

        i = 0
        while (not bFound) and (i < UBound(m_PrimaryKeyFields))
            i = i + 1
            if m_PrimaryKeyFields(i) = sName then
                nReturn = m_PrimaryKeyTypes(i)
                bFound = True
            end if
        wend

        i = 0
        while (not bFound) and (i < UBound(m_StandardFields))
            i = i + 1
            if m_StandardFields(i) = sName then
                nReturn = m_StandardTypes(i)
                bFound = True
            end if
        wend

        getFieldType = nReturn

    End Function


    ''----------------------------------------------------------------------
    '' Name:     getFilter
    ''           =========
    '' 
    '' Returns complete Filter SQL statement
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      string
    ''
    ''----------------------------------------------------------------------
    Private Function getFilter ()
        Dim sReturn, sCombine
        Dim i, bFirst
        bFirst = True

        sReturn  = ""
        sCombine = ""

        for i = 1 to m_nNumberOfFilters
            
            if i > 1 then sCombine = m_FilterCombines(i-1)

            sReturn = sReturn & AddWhere( _
                    m_FilterFields(i), _
                    getFieldType(m_FilterFields(i)), _
                    m_FilterValues(i), _
                    m_FilterCompares(i), _
                    sCombine, _
                    bFirst)
                        
            bFirst = False
        next

        GetFilter = sReturn
    End Function


    ''----------------------------------------------------------------------
    '' Name:     getAllRecordsFromDB
    ''           ===================
    '' 
    '' Creates SQL statement to get all records from table, opens
    '' and configures recordset.
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Private Sub getAllRecordsFromDB ()
        Dim i
        Dim s
        Dim sUserID,rsTemp
    
        m_sSQL = "SELECT * FROM " & m_sTable

        ' if no sort field is given select one
        if UBound(m_SortFields) = 0 then
            if UBound(m_PrimaryKeyFields) <> 0 then 
                AddSortField m_PrimaryKeyFields(1) 
                AddSortOrder SORT_DESC
            elseif UBound(m_StandardFields) <> 0 then 
                AddSortField m_StandardFields(1)
                AddSortOrder SORT_DESC
            end if
        end if

        ' add WHERE clause
        if UBound(m_FilterFields) >= m_nNumberOfFilters then
            m_sSQL = m_sSQL & getFilter
        end if

        ' add ORDER BY clause
        if UBound(m_SortFields) <> 0 then
            s = " ORDER BY "
            for i = 1 to UBound(m_SortFields)
                s = s & m_SortFields(i)
                if m_SortFieldsOrder(i) = SORT_DESC then s = s & " DESC"
                s = s & ", "
            next
            ' cut trailing ", "
            s = Left(s, Len(s)-2)
            
            if LoginAdmin=false then
				Set rsTemp = Server.CreateObject("ADODB.Recordset")
				rsTemp.Open "select " & f_RelationField & " from " & tbl_UsersTable & " where id=" & f_UserID, m_DB, adOpenStatic,adLockReadOnly														
				if rsTemp.EOF = false then
					sUserID=rsTemp.Fields.Item(f_RelationField).value				
				end if
				rsTemp.close
				Set rsTemp = Nothing            
				
				'Search to Find if m_sSQL has already the clause WHERE
				if UBound(m_FilterFields) >= m_nNumberOfFilters then
					m_sSQL = m_sSQL & " AND " & f_RelationField & "=" & Cstr(sUserID)
				else
					m_sSQL = m_sSQL & " Where " & f_RelationField & "=" & Cstr(sUserID)
				end if
            end if
            m_sSQL = m_sSQL & s
        end if 

        on error resume next
        m_RS.Open m_sSQL, m_DB, adOpenStatic

        if Err <> 0 then
            if UBound(m_FilterCompares) <> 0 then
                ' redirect to filter page and display original error message
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamMode, MD_FILTER)
                s = getLink(m_sUTEScript, s, sParamSubmitted, "0")
                s = getLink(m_sUTEScript, s, sParamFilterError, Server.URLEncode(Err.description))
                Response.Redirect s
            else
                Response.Write CStr(Hex(Err)) & ": " & Err.Description
                Response.End
            end if
        end if

        m_RS.PageSize  = m_nPageSize

        if m_nPage > m_RS.PageCount then
            m_nPage = m_RS.PageCount
        end if

        if m_nPage <> 0 then
            m_RS.AbsolutePage = m_nPage
        end if

    End Sub


    '-----------------------------------------------------------------------
    ' Inlcude all mode specific private class functions
    '
%>
<!--#include file ="ute_class_database.inc"-->
<!--#include file ="ute_class_table.inc"-->
<!--#include file ="ute_class_form.inc"-->
<!--#include file ="ute_class_export.inc"-->
<!--#include file ="ute_class_filter.inc"-->
<%

    '-----------------------------------------------------------------------
    ' Public Member Functions
    '

    ''----------------------------------------------------------------------
    '' Name:     Init
    ''           ====
    '' 
    '' Read all paramters, analyze table and prepares HTML output.
    ''
    '' Parameter: 
    ''      sDSN        ODBC connection string
    ''      bReadOnly   Display table in readonly mode
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Public Sub Init( sDSN )

        m_sDSN = sDSN



        ' read all other parameters from URL
        GetParameter()

        ' are there filters to be set ?
        if m_nMode = MD_FILTER then
            if m_bSubmitted then
                ' set filter and redirect to table
                UpdateFilter
            end if
        end if

        ' calculate headline 
        select case m_nMode
            case MD_DATABASE
                m_sHeadLine = m_sDBName
            case MD_TABLE
                m_sHeadLine = m_sTable
            case MD_FILTER
                m_sHeadLine = STR_DEF_FILTER
            case MD_FORM
                select case m_nFormMode
                    case MD_INSERT
                        m_sHeadLine = STR_INSERT
                    case MD_EDIT
                        m_sHeadLine = STR_EDIT
                    case MD_DELETE
                        m_sHeadLine = STR_DELETE
                end select
        end select

        ' open database connection
        m_DB.Open m_sDSN



    End Sub

    '-----------------------------------------------------------------------
    ' Public Member Functions
    '

    ''----------------------------------------------------------------------
    '' Name:     InitSecondPart
    ''           ====
    '' 
    '' Break the Init Section just to cach the user parameters before it will call Export, Filter etc
    ''
    '' Parameter: 
    ''		none
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
	Public Sub InitSecondPart
	
        if m_nMode <> MD_DATABASE then

            ' get all fields from table
            AnalyzeTable

            if m_nMode <> MD_FILTER then
                ' load all records from db
                getAllRecordsFromDB
            end if

            if m_nMode = MD_EXPORT then 
                ' create csv data and send it to the response stream
                ExportToStream
            end if

            if m_nMode = MD_FORM then
                ' update record and redirect to table
                Update
            end if

        end if	
	
	
	End Sub

    ''----------------------------------------------------------------------
    '' Name:     Draw
    ''           ====
    '' 
    '' Writes entire HTML code directly to stream.
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      none
    ''
    ''----------------------------------------------------------------------
    Public Sub Draw()
    
		
		InitSecondPart()
		
        select case m_nMode
            case MD_DATABASE
                Response.Write buildHTML_Database
            case MD_TABLE
                Response.Write buildHTML_Table
            case MD_FORM
                Response.Write buildHTML_Form
            case MD_FILTER
                Response.Write buildHTML_Filter
        end select
    End Sub


    ''----------------------------------------------------------------------
    '' Name:     getHTML
    ''           =======
    '' 
    '' Returns entire HTML code as string.
    ''
    '' Parameter: 
    ''      none
    ''
    '' return value:
    ''      string      entire UTE HTML code
    ''
    ''----------------------------------------------------------------------
    Public Function getHTML()
        select case m_nMode
            case MD_DATABASE
                getHTML = buildHTML_Database
            case MD_TABLE
                getHTML = buildHTML_Table
            case MD_FORM
                getHTML = buildHTML_Form
            case MD_FILTER
                getHTML = buildHTML_Filter
        end select
    End Function


End Class

%>