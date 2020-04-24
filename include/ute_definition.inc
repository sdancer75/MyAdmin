

<!--#include file ="ute_language_el.asp"-->
<!--#include file ="ute_adolib.inc"-->
<!--#include file ="ute_class.inc"-->
<!--#include file ="WebSiteAccount.asp"-->
<%
'---------------------------------------------------------------------------
' Definitions
'
Const sUTELongName         = "Paradox Interactive"
Const sUTEShortName        = "MyAdmin"
Const sUTEVersion          = "v2.10"
Const sUTELink             = "http://www.paradoxinteractive.gr"

'---------------------------------------------------------------------------
' "public" URL Parameter
'
Const sParamTable          = "name"                ' name of table
Const sParamPKey           = "pkey"                ' e.g. pkey1, pkey2, ... pkeyn
Const sParamSortFields     = "sorted"              ' sort fields alphabetically (1=true, 0=false, default=0)

'---------------------------------------------------------------------------
' "private" URL Parameter
'
Const sParamPage           = "page"                ' current page
Const sParamPageSize       = "pagesize"            ' page size (default=10)
Const sParamSort           = "sort"                ' sort field
Const sParamSortDir        = "sortdir"             ' sort direction  (asc, desc, default=asc) 
Const sParamMode           = "mode"                ' MD_TABLE, MD_FORM, MD_EXPORT
Const sParamFormMode       = "formmode"            ' MD_INSERT, MD_EDIT, MD_DELETE
Const sParamRecord         = "record"              ' editing or deleting record
Const sParamDefs           = "definitions"         ' show field definitions (1=true, 0=false, default=0)
Const sParamSQL            = "sql"                 ' show sql statement  (1=true, 0=false, default=0)
Const sParamSubmitted      = "submitted"           ' flag to signal that page was submitted (1=true, 0=false)
Const sParamFilterCount    = "fltcount"            ' number of filters
Const sParamFilterField    = "fltfield"            ' field to filter for, e.g.: fltfield1, fltfield1, ... fltfieldn
Const sParamFilterCompare  = "fltcomp"             ' comparison for filter, e.g.: fltcomp1, fltcomp2, ... fltcompn
Const sParamFilterValue    = "fltvalue"            ' value to filter for, e.g.: fltvalue1, fltvalue2, ... fltvaluen
Const sParamFilterCombine  = "fltcomb"             ' combine filters, e.g.: fltcomb1, fltcomb2, ... fltcombn
Const sParamFilterError    = "flterror"            ' error message

'---------------------------------------------------------------------------
' form fields
'
Const sFormUTEFieldPrefix  = "?"                   ' all UTE own form fields statr with this character. This must
                                                   ' not be a valid SQL fieldname character to ensure integrity
Const sFormButton          = "ute_form_button"     ' name of ok and cancel buttons in form
Const sFormIdentField      = "ute_ident_field"     ' name of field to identify a record to be edited
Const sFormIdentType       = "ute_ident_type"      ' type of field to identify a record to be edited
Const sFormIdentValue      = "ute_ident_value"     ' value of field to identify a record to be edited
Const sFormCount           = "ute_count"           ' number of filters 
Const sFormField           = "ute_field"           ' name of "field" input in filter form
Const sFormCompare         = "ute_compare"         ' name of "compare" input in filter form
Const sFormValue           = "ute_value"           ' name of "value" input in filter form
Const sFormCombine         = "ute_combine"         ' name of "combine" input in filter form

'---------------------------------------------------------------------------
' default values
'
Const DEF_PAGE             = 1                     ' default page
Const DEF_PAGE_SIZE        = 10                    ' default number of records per page
Const DEF_SORT_FIELDS      = False                 ' default sort fields alphabetically (columns)
Const DEF_SORT_DIR         = "asc"                 ' default sort direction: SORT_ASC
Const DEF_VIEW_DEFINITIONS = False                 ' default show field definitions
Const DEF_VIEW_SQL         = False                 ' default show sql statement
Const DEF_PK_DETECTION     = True                  ' default primary key detection

Const DEF_MODE             = 1                     ' default view mode: MD_TABLE
Const DEF_FORM_MODE        = 1                     ' default form mode: MD_INSERT

Const DEF_READONLY         = False                 ' default readonly
Const DEF_LIST_TABLES      = True                  ' default list all table in DB
Const DEF_SHOW_DEF_LINK    = True                  ' default show view definitions link
Const DEF_EXPORT_LINK      = True                  ' default export data link
Const DEF_SQL_LINK         = True                  ' default show sql link
Const DEF_FILTERS          = True                  ' default show and activate filters

Const DEF_IMAGE_DIR        = "images/"             ' directory where the images are
Const DEF_USER_DIR         = "../UserFiles"        ' User Dir for the FileManagement.asp 

Const DEF_EXPORT_SEP       = ","                   ' seperator chacarter between values
Const DEF_EXPORT_VAL       = """"                  ' chararcter a value is placed into

Const DEF_MAX_INPUT_LENGTH = 58                    ' max length of INPUT
Const DEF_MEMO_COLS        = 50                    ' number of cols of TEXTAREA
Const DEF_MEMO_ROWS        = 8                     ' number of rows of TEXTAREA
Const DEF_MEMO_COL_WIDTH   = 300                   ' width of MEMO column in table

Const DEF_NUM_FILTER       = 1                     ' default number of filters
Const DEF_MAX_FILTER       = 10                    ' maximum number of filters
Const DEF_MAX_FILTER_LEN   = 50                    ' max length of filter INPUT
Const DEF_FILTER_SIZE      = 30                    ' size of filter INPUT

'---------------------------------------------------------------------------
' view modes and form modes
'
Const MD_DATABASE          = 0                     ' list tables of current database
Const MD_TABLE             = 1                     ' table view mode
Const MD_FORM              = 2                     ' form view mode
Const MD_EXPORT            = 3                     ' export mode
Const MD_FILTER            = 4                     ' display filter form

Const MD_INSERT            = 1                     ' insert mode
Const MD_EDIT              = 2                     ' edit mode
Const MD_DELETE            = 3                     ' delete mode

'---------------------------------------------------------------------------
' other defines
'
Const SORT_ASC             = "asc"
Const SORT_DESC            = "desc"

'---------------------------------------------------------------------------
' FileOpen iomode Values
'
Const fsoForReading        = 1
Const fsoForWriting        = 2
Const fsoForAppending      = 8
Const fsoCreateIfNotExist  = True

'--------------My definitions ----------------------------------------------
Const MAX_FIELDS_TO_SHOW = 10
Const NO_LOOKUP_TABLE	 = 1
Const VALUES_TABLE		 = 2
Const SQL_QUERY			 = 3
Const REPLACE_VALUES	 = 4


'==================== User Defined Variables ============================



Const tbl_UsersTable					="Users"
Const f_Username						="username"
Const f_Password						="password"
Const f_RelationField					="CSN"					


Const html_username					="Username"
Const html_password					="Password"
Const html_submit						="submit"

Const msg_NoValidLogin				="No Valid Username or Password"
Const sAuthAccessTitle				=""
Const sLoginStartPage					="admin.asp"


Const FIELD_MAX_LEN_TO_MAKE_TEXTAREA    =150



%>