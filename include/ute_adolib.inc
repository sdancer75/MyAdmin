<%


' Microsofts ADOVBS.INC must be included before !!!


'---------------------------------------------------------------------------
' Additional DataTypeEnum Values
'
Const adChapter = 136
Const adPropVariant = 138


'---------------------------------------------------------------------------
' These field types won't be displayed
' and are not editable
'
dim   ExcludeTypes()
redim ExcludeTypes(0)


''--------------------------------------------------------------------------
'' Name:     AddExcludeType
''           ==============
'' 
'' Inserts field types not to be displayed by UTE.
''
'' Parameter: 
''      nType       integer     Type of field.
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Sub AddExcludeType ( nType )
    Redim Preserve ExcludeTypes(UBound(ExcludeTypes)+1)
    ExcludeTypes(UBound(ExcludeTypes)) = nType
End Sub

AddExcludeType adEmpty            ' 0
AddExcludeType adError            ' 10
AddExcludeType adUserDefined      ' 132
AddExcludeType adVariant          ' 12
AddExcludeType adIDispatch        ' 9
AddExcludeType adIUnknown         ' 13
AddExcludeType adBinary           ' 128
AddExcludeType adVarBinary        ' 204
AddExcludeType adLongVarBinary    ' 205
AddExcludeType adChapter          ' 136
AddExcludeType adPropVariant      ' 138


''--------------------------------------------------------------------------
'' Name:     GetTypeString
''           =============
'' 
'' Returns the Type of an ADODB.Field as string
''
'' Parameter: 
''      nType       integer     Type of field.
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
function GetTypeString ( nType )
    dim sReturn
    select case nType
        case adEmpty
            sReturn = STR_ADO_TYPE_EMPTY
        case adTinyInt
            sReturn = STR_ADO_TYPE_TINYINT
        case adSmallInt
            sReturn = STR_ADO_TYPE_SMALLINT
        case adInteger
            sReturn = STR_ADO_TYPE_INTEGER
        case adBigInt
            sReturn = STR_ADO_TYPE_BIGINT
        case adUnsignedTinyInt
            sReturn = STR_ADO_TYPE_UNSIGNEDTINYINT
        case adUnsignedSmallInt
            sReturn = STR_ADO_TYPE_UNSIGNEDSMALLINT
        case adUnsignedInt
            sReturn = STR_ADO_TYPE_UNSIGNEDINT
        case adUnsignedBigInt
            sReturn = STR_ADO_TYPE_UNSIGNEDBIGINT
        case adSingle
            sReturn = STR_ADO_TYPE_SINGLE
        case adDouble
            sReturn = STR_ADO_TYPE_DOUBLE
        case adCurrency
            sReturn = STR_ADO_TYPE_CURRENCY
        case adDecimal
            sReturn = STR_ADO_TYPE_DECIMAL
        case adNumeric
            sReturn = STR_ADO_TYPE_NUMERIC
        case adBoolean
            sReturn = STR_ADO_TYPE_BOOLEAN
        case adError
            sReturn = STR_ADO_TYPE_ERROR
        case adUserDefined
            sReturn = STR_ADO_TYPE_USERDEFINED
        case adVariant
            sReturn = STR_ADO_TYPE_VARIANT
        case adIDispatch
            sReturn = STR_ADO_TYPE_IDISPATCH
        case adIUnknown
            sReturn = STR_ADO_TYPE_IUNKNOWN
        case adGUID
            sReturn = STR_ADO_TYPE_GUID
        case adDBDate
            sReturn = STR_ADO_TYPE_DBDATE
        case adDBTime
            sReturn = STR_ADO_TYPE_DBTIME
        case adDBTimeStamp
            sReturn = STR_ADO_TYPE_DBTIMESTAMP
        case adBSTR
            sReturn = STR_ADO_TYPE_BSTR
        case adChar
            sReturn = STR_ADO_TYPE_CHAR
        case adVarChar
            sReturn = STR_ADO_TYPE_VARCHAR
        case adLongVarChar
            sReturn = STR_ADO_TYPE_LONGVARCHAR
        case adWChar
            sReturn = STR_ADO_TYPE_WCHAR
        case adVarWChar
            sReturn = STR_ADO_TYPE_VARWCHAR
        case adLongVarWChar
            sReturn = STR_ADO_TYPE_LONGVARWCHAR
        case adBinary
            sReturn = STR_ADO_TYPE_BINARY
        case adVarBinary
            sReturn = STR_ADO_TYPE_VARBINARY
        case adLongVarBinary
            sReturn = STR_ADO_TYPE_LONGVARBINARY
        case adChapter
            sReturn = STR_ADO_TYPE_CHAPTER
        case adPropVariant
            sReturn = STR_ADO_TYPE_PROPVARIANT
        case else
            sReturn = STR_ADO_TYPE_UNKONWN
    end select

    GetTypeString = sReturn

end function


''--------------------------------------------------------------------------
'' Name:     GetAttributesString
''           ===================
'' 
'' Returns the Attributes of an ADODB.Field as string
''
'' Parameter: 
''      nAttrib     integer     bitmasked attributes
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
function GetAttributesString ( nAttrib )
    dim sReturn
    sReturn = ""
    if (nAttrib and adFldKeyColumn) then        sReturn = sReturn & STR_ADO_KEY & ", "
    if (nAttrib and adFldMayDefer) then         sReturn = sReturn & STR_ADO_MAYDEFER & ", "
    if (nAttrib and adFldUpdatable) then        sReturn = sReturn & STR_ADO_UPDATEABLE & ", "
    if (nAttrib and adFldUnknownUpdatable) then sReturn = sReturn & STR_ADO_UNKNOWNUPDATEABLE & ", "
    if (nAttrib and adFldFixed) then            sReturn = sReturn & STR_ADO_FIXED & ", "
    if (nAttrib and adFldIsNullable) then       sReturn = sReturn & STR_ADO_ISNULLABLE & ", "
    if (nAttrib and adFldMayBeNull) then        sReturn = sReturn & STR_ADO_MAYBENULL & ", "
    if (nAttrib and adFldLong) then             sReturn = sReturn & STR_ADO_LONG & ", "
    if (nAttrib and adFldRowID) then            sReturn = sReturn & STR_ADO_ROWID & ", "
    if (nAttrib and adFldRowVersion) then       sReturn = sReturn & STR_ADO_ROWVERSION & ", "
    if (nAttrib and adFldCacheDeferred) then    sReturn = sReturn & STR_ADO_CACHEDEFERRED & ", "
    if sReturn <> "" then sReturn = Left(sReturn, Len(sReturn)-2)
    GetAttributesString = sReturn
end function


''--------------------------------------------------------------------------
'' Name:     Attrib
''           ======
'' 
'' Checks, if a specified attribute is set
''
'' Parameter: 
''      nField      integer     bitmasked attribute
''      nAttrib     integer     attribute to be checked for
''
'' return value:
''      boolean
''
''--------------------------------------------------------------------------
function Attrib ( nField, nAttrib )
    Attrib = ((nField and nAttrib) <> 0)
end function


''--------------------------------------------------------------------------
'' Name:     NotAttrib
''           =========
'' 
'' Checks, if a specified attribute is not set
''
''      nField      integer     bitmasked attribute
''      nAttrib     integer     attribute to be checked for
''
'' return value:
''      boolean
''
''--------------------------------------------------------------------------
function NotAttrib ( nField, nAttrib )
    NotAttrib = (not Attrib(nField, nAttrib))
end function


''--------------------------------------------------------------------------
'' Name:     IsExcluded
''           ==========
'' 
'' Checks, if given Type is in Exclude List
''
'' Parameter: 
''      nType       integer     Field type
''
'' return value:
''      boolean
''
''--------------------------------------------------------------------------
function IsExcluded ( nType )
    Dim bReturn
    Dim ii
    bReturn = false
    for ii = 1 to UBound(ExcludeTypes)
        if nType = ExcludeTypes(ii) then bReturn = True
    next
    IsExcluded = bReturn
end function


%>