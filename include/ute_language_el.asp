<%

Session.LCID = 1032 ' Greek  

'Session.LCID = &h0409 ' English (United States) 
'Session.LCID = &h0809 ' English (United Kingdom) 
'Session.LCID = &h0c09 ' English (Australian) 
'Session.LCID = &h1009 ' English (Canadian) 
'Session.LCID = &h1409 ' English (New Zealand) 
'Session.LCID = &h1809 ' English (Ireland) 
'Session.LCID = &h1c09 ' English (South Africa) 
'Session.LCID = &h2009 ' English (Jamaica) 
'Session.LCID = &h2409 ' English (Caribbean) 
'Session.LCID = &h2809 ' English (Belize) 
'Session.LCID = &h2c09 ' English (Trinidad) 
'Session.LCID = &h3009 ' English (Zimbabwe) 
'Session.LCID = &h3409 ' English (Philippines)
'----- custom --------
Const STR_HOOKBROWSE_TITLE			   = " Επιλογή Αρχείου "
Const sMainMenuTitle					="Η Σελίδα του Διαχειριστή"

' ---- Application ----
Const STR_SORT_ASC                      = "αύξουσα ταξινόμηση"
Const STR_SORT_DESC                     = "φθίσουσα ταξινόμηση"

Const STR_DATABASE                      = "Βάση δεδομένων"
Const STR_DB_TITLE                      = "%1"

Const STR_INSERT                        = "Νέα εγγραφή"
Const STR_EDIT                          = "Επεξεργασία εγγραφής"
Const STR_DELETE                        = "Διαγραφή εγγραφής"

Const STR_DEF_FILTER                    = "Ορισμός φίλτρου"
Const STR_NUM_FILTER                    = "Αρ. Φίλτρων:"

Const STR_NON_VIEW                      = "none-viewable data"

Const STR_OK                            = "Αποθήκευση"
Const STR_CANCEL                        = "’κυρο"
Const STR_CLEAR                         = "Καθαρισμός"

Const STR_PAGES                         = "Σελίδα:"
Const STR_NEXT_PAGE                     = "μετάβαση στην επόμενη σελίδα"
Const STR_PREV_PAGE                     = "μετάβαση στην προηγούμενη σελίδα"
Const STR_REC_COUNT                     = "Εγγραφές ανα σελίδα:"
Const STR_ALL                           = "Προβολή όλων"

Const STR_RECORDS                       = "Προβολή εγγραφής %1 - %2 από %3 συνολικά"

Const STR_POWERED_BY                    = "powered by %1 %2"
Const STR_FILTER                        = "Ορισμός φίλτρου"
Const STR_LIST_TABLES                   = "Προβολή πινάκων της Βάσης Δεδομένων"
Const STR_EXPORT                        = "Αποθήκευση ως CSV (Excel) αρχείο"
Const STR_DEF_SHOW                      = "Show Field Definitions"
Const STR_DEF_HIDE                      = "Hide Field Definitions"
Const STR_SQL_SHOW                      = "Show current SQL Statment"
Const STR_SQL_HIDE                      = "Hide current SQL Statment"

Const STR_DEF_NAME                      = "Name"
Const STR_DEF_TYPE                      = "Type"
Const STR_DEF_DEFINEDSIZE               = "Definied Size"
Const STR_DEF_PRECISION                 = "Precision"
Const STR_DEF_ATTRIBUTES                = "Attributes"

' ---- ADO -----
Const STR_ADO_KEY                       = "Key" 
Const STR_ADO_MAYDEFER                  = "may defer" 
Const STR_ADO_UPDATEABLE                = "updatable"
Const STR_ADO_UNKNOWNUPDATEABLE         = "unknown updatable"
Const STR_ADO_FIXED                     = "fixed"
Const STR_ADO_ISNULLABLE                = "can be set to NULL"
Const STR_ADO_MAYBENULL                 = "may be NULL"
Const STR_ADO_LONG                      = "long"
Const STR_ADO_ROWID                     = "Row ID"
Const STR_ADO_ROWVERSION                = "Row Version"
Const STR_ADO_CACHEDEFERRED             = "Cache deferred"

Const STR_ADO_TYPE_EMPTY                = "Empty"
Const STR_ADO_TYPE_TINYINT              = "TinyInt"
Const STR_ADO_TYPE_SMALLINT             = "SmallInt"
Const STR_ADO_TYPE_INTEGER              = "Integer"
Const STR_ADO_TYPE_BIGINT               = "BigInt"
Const STR_ADO_TYPE_UNSIGNEDTINYINT      = "UnsignedTinyInt"
Const STR_ADO_TYPE_UNSIGNEDSMALLINT     = "UnsignedSmallInt"
Const STR_ADO_TYPE_UNSIGNEDINT          = "UnsignedInt"
Const STR_ADO_TYPE_UNSIGNEDBIGINT       = "UnsignedBigInt"
Const STR_ADO_TYPE_SINGLE               = "Single"
Const STR_ADO_TYPE_DOUBLE               = "Double"
Const STR_ADO_TYPE_CURRENCY             = "Currency"
Const STR_ADO_TYPE_DECIMAL              = "Decimal"
Const STR_ADO_TYPE_NUMERIC              = "Numeric"
Const STR_ADO_TYPE_BOOLEAN              = "Boolean"
Const STR_ADO_TYPE_ERROR                = "Error"
Const STR_ADO_TYPE_USERDEFINED          = "UserDefined"
Const STR_ADO_TYPE_VARIANT              = "Variant"
Const STR_ADO_TYPE_IDISPATCH            = "IDispatch"
Const STR_ADO_TYPE_IUNKNOWN             = "IUnknown"
Const STR_ADO_TYPE_GUID                 = "GUID"
Const STR_ADO_TYPE_DBDATE               = "DBDate"
Const STR_ADO_TYPE_DBTIME               = "DBTime"
Const STR_ADO_TYPE_DBTIMESTAMP          = "DBTimeStamp"
Const STR_ADO_TYPE_BSTR                 = "BSTR"
Const STR_ADO_TYPE_CHAR                 = "Char"
Const STR_ADO_TYPE_VARCHAR              = "VarChar"
Const STR_ADO_TYPE_LONGVARCHAR          = "LongVarChar"
Const STR_ADO_TYPE_WCHAR                = "WChar"
Const STR_ADO_TYPE_VARWCHAR             = "VarWChar"
Const STR_ADO_TYPE_LONGVARWCHAR         = "LongVarWChar"
Const STR_ADO_TYPE_BINARY               = "Binary"
Const STR_ADO_TYPE_VARBINARY            = "VarBinary"
Const STR_ADO_TYPE_LONGVARBINARY        = "LongVarBinary"
Const STR_ADO_TYPE_CHAPTER              = "Chapter"
Const STR_ADO_TYPE_PROPVARIANT          = "PropVariant"
Const STR_ADO_TYPE_UNKONWN              = "Unknown"
Const STR_REQUIREDFIELD					= "Υποχρεωτικό πεδίο"
Const STR_SELECT_DROPDOWN				= "Παρακαλώ επιλέξτε"
Const STR_BACK_TO_MAIN_MENU				= "[ΕΠΙΣΤΡΟΦΗ ΣΤΟ ΚΕΝΤΡΙΚΟ ΜΕΝΟΥ]"

' ---- Error Messages ----
Const STR_ERR_1001                      = "Invalid ODBC Connection String"
Const STR_ERR_1002                      = "Missing ""%1"" URL parameter."
Const STR_ERR_1003                      = "Invalid ""%1"" URL parameter. Must be numeric."
Const STR_ERR_1004                      = "Invalid ""%1"" URL parameter. Must be ""1"", ""2"" or ""3""."
Const STR_ERR_1005                      = "Invalid ""%1"" URL parameter. Must be either ""%2"" or ""%3""."

%>