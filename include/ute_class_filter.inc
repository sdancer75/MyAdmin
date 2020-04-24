<%




''--------------------------------------------------------------------------
'' Name:     UpdateFilter
''           ============
'' 
'' Reads filter settings from form, passes them to the URL and redirects
'' to ute page in table view mode.
''
'' Parameter: 
''      none
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub UpdateFilter

    Dim bValid, i, s

    s = Request.QueryString
    s = getLink(m_sUTEScript, s, sParamMode, MD_TABLE)
    s = getLink(m_sUTEScript, s, sParamFormMode,  CStr(DEF_FORM_MODE))
    s = getLink(m_sUTEScript, s, sParamSubmitted, "0")

    ' get form data if "OK" was clicked
    if Request.Form(sFormUTEFieldPrefix & sFormButton) <> STR_CANCEL then

        ' reset all table view settings
        s = RemoveParameter(s, sParamPage)
        s = RemoveParameter(s, sParamPageSize)
        s = RemoveParameter(s, sParamSort)
        s = RemoveParameter(s, sParamSortDir)
        s = RemoveParameter(s, sParamFilterError)

        ' remove all current filter settings
        s = RemoveCountedParameters(s, sParamFilterField,   1)
        s = RemoveCountedParameters(s, sParamFilterCompare, 1)
        s = RemoveCountedParameters(s, sParamFilterValue,   1)
        s = RemoveCountedParameters(s, sParamFilterCombine, 1)

        if Request.Form(sFormUTEFieldPrefix & sFormButton) = STR_CLEAR then

            ' set number of filters to default
            s = RemoveParameter(s, sParamFilterCount)

        elseif Request.Form(sFormUTEFieldPrefix & sFormButton) = STR_OK then

            ' add filters to URL
            bValid = True
            for i = 1 to m_nNumberOfFilters
                if bValid then
                    if Request.Form(sFormField & CStr(i)) <> "" then
                        s = getLink(m_sUTEScript, s, sParamFilterField & CStr(i), _
                            Server.URLEncode(Request.Form(sFormField & CStr(i))))
                        s = getLink(m_sUTEScript, s, sParamFilterCompare & CStr(i), _
                            Server.URLEncode(Request.Form(sFormCompare & CStr(i))))
                        s = getLink(m_sUTEScript, s, sParamFilterValue & CStr(i), _
                            Server.URLEncode(Request.Form(sFormValue & CStr(i))))
                    else
                        bValid = False
                    end if
                end if
            next
            ' add filter combinations to URL
            for i = 1 to m_nNumberOfFilters - 1
                s = getLink(m_sUTEScript, s, sParamFilterCombine & CStr(i), Request.Form(sFormCombine & CStr(i)))
            next

        end if

    end if

    Response.Redirect s

End Sub


''--------------------------------------------------------------------------
'' Name:     AddOption
''           =========
'' 
'' Creates HTML code for OPTION tag.
''
'' Parameter: 
''      sValue      value to be set
''      sText       text to be displayed
''      sSelected   selected value
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function AddOption ( sValue, sText, sSelected )
    Dim sReturn,sLabel
    
    sReturn = "<option value='" & sValue & "'"
    if sValue = sSelected then sReturn = sReturn & " selected"
    
    if sText<>"=" then
		sLabel=FieldHasALabel(sText)
   		if sLabel<>"" then    
			sReturn = sReturn & ">" & sLabel & "</option>"
		else
			sReturn = sReturn & ">" & sText & "</option>"
		end if
	else
		sReturn = sReturn & ">" & sText & "</option>"
	end if
    
    AddOption = sReturn
End Function


''--------------------------------------------------------------------------
'' Name:     buildFilterCountList
''           ====================
'' 
'' Creates HTML code for number of filters selections.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildFilterCountList()
    Dim i, s, sReturn

    s = Request.QueryString

    sReturn = _
        "<select class=""ute_navigation"" name=""" & sFormCount & """ " & _
        "onChange=""jumpPage('parent', this, 0)"">" & vbCrLf

    for i = 1 to DEF_MAX_FILTER
        s = getLink(m_sUTEScript, s, sParamFilterCount, CStr(i))
        sReturn = sReturn & "<option value='" & s & "'" 
        if i = m_nNumberOfFilters then sReturn = sReturn & " selected"
        sReturn = sReturn & ">" & CStr(i) & "</option>" & vbCrLf
    next

    sReturn = sReturn & "</select>"

    buildFilterCountList = sReturn

End Function


''--------------------------------------------------------------------------
'' Name:     buildFilterFieldList
''           ====================
'' 
'' Creates HTML code for FIELD selection.
''
'' Parameter: 
''      nFilter     number of current filter
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildFilterFieldList ( nFilter )
    Dim i, sField, sReturn

    sField = ""
    if nFilter <= UBound(m_FilterFields) then sField = m_FilterFields(nFilter)

    sReturn = _
        "<select class=""ute_navigation"" name=""" & sFormField & CStr(nFilter) & """>" & vbCrLf & _
        "<option value=''></option>" & vbCrLf

    for i = 1 to UBound(m_PrimaryKeyFields)
        sReturn = sReturn & AddOption(m_PrimaryKeyFields(i), m_PrimaryKeyFields(i), sField) & vbCrLf
    next
    for i = 1 to UBound(m_StandardFields)
        sReturn = sReturn & AddOption(m_StandardFields(i), m_StandardFields(i), sField) & vbCrLf
    next

    sReturn = sReturn & "</select>"

    buildFilterFieldList = sReturn

End Function


''--------------------------------------------------------------------------
'' Name:     buildFilterCompareList
''           ======================
'' 
'' Creates HTML code for compare selection.
''
'' Parameter: 
''      nFilter     number of current filter
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildFilterCompareList ( nFilter )
    Dim sCompare, sReturn

    sCompare = ""
    if nFilter <= UBound(m_FilterCompares) then sCompare = m_FilterCompares(nFilter)

    sReturn = _
        "<select class=""ute_navigation"" name=""" & sFormCompare & CStr(nFilter) & """>" & vbCrLf & _
        AddOption("=",      "=",      sCompare) & vbCrLf & _
        AddOption(">",      "&gt;",   sCompare) & vbCrLf & _
        AddOption(">=",     "&gt;=",  sCompare) & vbCrLf & _
        AddOption("<",      "&lt;",   sCompare) & vbCrLf & _
        AddOption("<=",     "&lt;=",  sCompare) & vbCrLf & _
        AddOption("LIKE",   "LIKE",   sCompare) & vbCrLf & _
        AddOption("IS",     "IS",     sCompare) & vbCrLf & _
        AddOption("IS NOT", "IS NOT", sCompare) & vbCrLf & _
        "</select>"

    buildFilterCompareList = sReturn

End Function


''--------------------------------------------------------------------------
'' Name:     buildFilterCombineList
''           ======================
'' 
'' Creates HTML code for connect selection.
''
'' Parameter: 
''      nFilter     number of current filter
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildFilterCombineList ( nFilter )
    Dim sReturn, sCombine

    sCombine = ""
    if nFilter <= UBound(m_FilterCombines) then sCombine = m_FilterCombines(nFilter)

    sReturn = _
        "<select class=""ute_navigation"" name=""" & sFormCombine & CStr(nFilter) & """>" & vbCrLf & _
        AddOption("AND", "AND", sCombine) & vbCrLf & _
        AddOption("OR",  "OR",  sCombine) & vbCrLf & _
        "</select>"

    buildFilterCombineList = sReturn

End Function


''--------------------------------------------------------------------------
'' Name:     buildHTML_Filter
''           ================
'' 
'' Creates entire UTE HTML code for filter view mode.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildHTML_Filter()

    Dim i, s, sValue, sField

    ' build complete string to handle form
    s = Request.QueryString
    s = getLink(m_sUTEScript, s, sParamSubmitted, "1")

    sValue = _
        "<script language=""JavaScript"">" & vbCrLf & _
        "<!--" & vbCrLf & _
        "   function jumpPage(targ, selObj, restore)" & vbCrLf & _
        "   {" & vbCrLf & _
        "       eval(targ + "".location='"" + selObj.options[selObj.selectedIndex].value + ""'"");" & vbCrLf & _
        "       if (restore) selObj.selectedIndex=0;" & vbCrLf & _
        "   }" & vbCrLf & _
        "//-->" & vbCrLf & _
        "</script>" & vbCrLf & _
        vbCrLf & _
        "<p><span class=""ute_headline"">" & m_sTable & "</span></p>" & vbCrLf & _
        "<p><span class=""ute_subheadline"">" & m_sHeadline & "</span></p>" & vbCrLf

    if Request.QueryString(sParamFilterError) <> "" then
        sValue = sValue & _
            "<p><span class=""ute_form_error"">" & _
            Request.QueryString(sParamFilterError) & _
            "</span></p>" & vbCrLf
    end if

    sValue = sValue & _
        "<form method=""post"" action=""" & s & """>" & vbCrLf & _
        "<table>" & vbCrLf & _
        "<tr><td class=""ute_navigation"" align=""left"">" & vbCrLf & _
        "<table><tr>" & vbCrLf & _
        "<td class=""ute_navigation"">" & STR_NUM_FILTER & "</td>" & vbCrLf & _
        "<td class=""ute_navigation"">" & buildFilterCountList & "</td>" & vbCrLf & _
        "</tr></table>" & vbCrLf & _
        "</td></tr>" & vbCrLf & _
        "<tr><td>&nbsp;</td></tr>" & vbCrLf

    for i = 1 to m_nNumberOfFilters
        sField = ""
        if i <= UBound(m_FilterValues) then sField = m_FilterValues(i)

        sValue = sValue & _
            "<tr><td class=""ute_navigation"" align=""left"">" & vbCrLf & _
            "<table><tr><td class=""ute_navigation"">" & vbCrLf & _
                buildFilterFieldList(i) & vbCrLf & _
                buildFilterCompareList(i) & vbCrLf & _
                "<input class=""ute_navigation"" type=""text"" name=""" & sFormValue & CStr(i) & """ " & _
                    "maxlength=""" & CStr(DEF_MAX_FILTER_LEN) & """ " & _
                    "size=""" & CStr(DEF_FILTER_SIZE) & """ value=""" & sField & """>" & vbCrLf & _
            "</td></tr></table>" & vbCrLf & _
            "</td></tr>" & vbCrLf
        if i < m_nNumberOfFilters then
            sValue = sValue & _
                "<tr><td align=""middle""><br>" & vbCrLf & _
                    buildFilterCombineList(i) & vbCrLf & _
                "<br><br></td></tr>" & vbCrLf
        end if
    next

    sValue = sValue & _
        "<tr><td align=""middle"">" & vbCrLf & _
        "<br><br><br>&nbsp;" & vbCrLf & _
        "<input type=""submit"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_ok"" " & _
                "value=""" & STR_OK & """>" & vbCrLf & _
        "&nbsp;" & vbCrLf & _
        "<input type=""submit"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_cancel"" " & _
                "value=""" & STR_CLEAR & """>" & vbCrLf & _
        "&nbsp;" & vbCrLf & _
        "<input type=""submit"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_cancel"" " & _
                "value=""" & STR_CANCEL & """>" & vbCrLf & _
        "&nbsp;" & vbCrLf & _
        "</td></tr>" & vbCrLf

    sValue = sValue & _
        "</table>" & vbCrLf & _
        "</form>" & vbCrLf

    buildHTML_Filter = sValue

End Function

%>