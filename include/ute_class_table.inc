<%



''--------------------------------------------------------------------------
'' Name:     InsertField
''           ===========
'' 
'' Returns HTML code for column header (i.e. field name)
''
'' Parameter: 
''      Field           field object
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
Private Function InsertField ( Field )

    Dim s
    Dim sValue
    Dim sOrderDir
    Dim sOrderDescr
    Dim sOrderPic
    Dim bPicLink
    Dim sLinkPrefix
    Dim sLinkSuffix
    Dim sImg
    Dim sSort, sSortDir
    Dim sLabel

    sSort    = m_SortFields(1)
    sSortDir = m_SortFieldsOrder(1)

    if (sSort <> Field.name) or (IsExcluded(Field.type)) then
        sOrderDir   = DEF_SORT_DIR
        if sOrderDir = SORT_ASC then
            sOrderDescr = STR_SORT_ASC
        else
            sOrderDescr = STR_SORT_DESC
        end if
        sOrderPic   = m_sIMAGEDir & "sort_none.gif"
        bPicLink    = False
    else
        if sSortDir = SORT_ASC then
            sOrderDir   = SORT_DESC
            sOrderDescr = STR_SORT_DESC
            sOrderPic   = m_sIMAGEDir & "sort_asc.gif"
        else '     -> SORT_DESC
            sOrderDir   = SORT_ASC
            sOrderDescr = STR_SORT_ASC
            sOrderPic   = m_sIMAGEDir & "sort_desc.gif"
        end if
        bPicLink = True
    end if

    sLinkPrefix = ""
    sLinkSuffix = ""

    if Not IsExcluded(Field.type) then
        s = Request.QueryString
        s = getLink(m_sUTEScript, s, sParamPage,    "1")            ' switch to page one !
        s = getLink(m_sUTEScript, s, sParamSort & "1",    Field.name)
        s = getLink(m_sUTEScript, s, sParamSortDir & "1", sOrderDir)
        s = RemoveCountedParameters(s, sParamSort, 2)
        s = RemoveCountedParameters(s, sParamSortDir, 2)
        sLinkPrefix = "<a href=""" & s & """ title=""" & sOrderDescr & """ class=""ute_link"">"
        sLinkSuffix = "</a>"
    end if

	sLabel=FieldHasALabel(Field.name)
	if sLabel<>"" then
		sValue = sLinkPrefix & sLabel & sLinkSuffix
	else
		sValue = sLinkPrefix & Field.name & sLinkSuffix
	end if

    if sOrderPic <> "" then
        sImg = "<img src=""" & sOrderPic & """ border=""0"" alt=""" & sOrderDescr & """ " &_
                "width=""10"" height=""9"">"
        sValue = sValue & "&nbsp;" & sLinkPrefix & sImg & sLinkSuffix
    end if 

    InsertField = sValue & vbCrLf

End Function


''--------------------------------------------------------------------------
'' Name:     InsertFieldValue
''           ================
'' 
'' Returns HTML code for field value
''
'' Parameter: 
''      Field       field object
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Function InsertFieldValue ( Field ) 
    Dim sField
    Dim sValue
    Dim Lookup,LookupParams
    Dim rsTemp,getParam1,getParam2,ParamError,ParamStr,i
    Dim Attrib,AllowNull,flag,pos,nextparam,substr,sLabel
    Dim case_labels(),case_values,sUserID,pkvalue,ForceShowAll

    if IsExcluded(Field.type) then
        sField = _
            "<center>" & _
                "<img src=""" & m_sIMAGEDir & "exclude.gif"" border=""0"" alt=""" & STR_NON_VIEW & """ " & _
                "width=""16"" height=""16"">" & _
            "</center>"
    else
        select case Field.type

            case adBoolean
                if CBool(Field.value) then
                    sValue = _
                        "<center>" & _
                            "<img src=""" & m_sIMAGEDir & "bool_true.gif"" border=""0"" alt=""" & CStr(True) & """ " & _
                            "width=""12"" height=""12"">" & _
                        "</center>"
                else
                    sValue = _
                        "<center>" & _
                            "<img src=""" & m_sIMAGEDir & "bool_false.gif"" border=""0"" alt=""" & CStr(False) & """ " & _
                            "width=""12"" height=""12"">" & _
                        "</center>"
                end if 

            case else
                if IsNull(Field.value) then
                    sValue = ""
                else
					Lookup=LookupTable(field.name)
					if  (Lookup = SQL_QUERY) and (m_bTableLookUp = True) then
						LookupParams=LookUpTableGetParams(field.name)
						pos=InStr(1,LookupParams,"(" )
						substr=Mid(LookupParams,pos+1,Len(LookupParams)-pos-1 )
						if substr <> "" Then
							
								ForceShowAll=False
								'replace the v_Field{field} parameter
								ParamError=False
								getParam1=InStr(1,ucase(SubStr),ucase("v_field{"))
								if getParam1 > 0 then
														
									getParam2=InStr(1,ucase(SubStr),ucase("}"))
									if getParam2 > 0 then
										ParamStr=Mid(SubStr,getParam1,(getParam2-getParam1)+1)
										SubStr=Replace(SubStr,ParamStr,"")
													
										'Remove pkvalue
										getParam1=InStr(1,ucase(SubStr),ucase("pkvalue{"))									
										if getParam1 > 0 then										
												getParam2=InStr(1,ucase(SubStr),ucase("}"))	
												if getParam2 > 0 then
													SubStr=Replace(SubStr,Mid(SubStr,getParam1,(getParam2-getParam1)+1),"")
																
												end if
										end if	
													
										'Remove ShowAll tag if exists
										getParam1=InStr(1,ucase(SubStr),ucase("showall"))
													
										if getParam1 > 0 then
											ForceShowAll=true										
											SubStr=Replace(SubStr,Mid(SubStr,getParam1,Len("showall")),"")
										end if									
																				
													
										getParam1=InStr(1,ParamStr,"{")
										if getParam1 > 0 then
											getParam2=InStr(1,ParamStr,"}")
											if getParam2 > 0 then
												ParamStr=Mid(ParamStr,getParam1+1,(getParam2-getParam1)-1)
												if ucase(Mid(ParamStr,1,InStr(1,ParamStr,":")-1)) = ucase(Field.Name) then
														'If admin come here
														if (LoginAdmin=True) or (ForceShowAll=true) then
																SubStr=SubStr & Cstr(Field.Value)
														else
																
																if ucase(Mid(ParamStr,1,InStr(1,ParamStr,":")-1)) = ucase(pkvalue) then
																	Set rsTemp = Server.CreateObject("ADODB.Recordset")
																	rsTemp.Open "select " & f_RelationField & " from " & tbl_UsersTable & " where id=" & f_UserID, m_DB, adOpenStatic,adLockReadOnly														
																	if rsTemp.EOF = false then
																		sUserID=rsTemp.Fields.Item(f_RelationField).value
																		SubStr = SubStr + Cstr(sUserID)
																	end if
																	rsTemp.close
																	Set rsTemp = Nothing
																	
																																		
																else																
																	SubStr=SubStr & Cstr(Field.Value)
																end if
																	
														end if
												end if
											else
												ParamError=True
											end if
										else
											ParamError=True
										end if
									Else
														
										ParamError=True
									End if
							
								end if
								
							
								
							
							if ParamError=False then
									Set rsTemp = Server.CreateObject("ADODB.Recordset")
									rsTemp.Open substr, m_DB, adOpenStatic,adLockReadOnly
									SubStr=	Mid(ParamStr,InStr(1,ParamStr,":")+1,Len(ParamStr))					
									
									if rsTemp.EOF = false then
									
											For i = 0 To rsTemp.Fields.Count - 1
												if ucase(rsTemp.Fields.Item(i).Name) = ucase(SubStr) then
													sValue = rsTemp.Fields.Item(i) & " <font size=1 color=blue>(Filter Value =  " & CStr(Field) &  ")</font>"
													exit for
												end if
											Next	
									end if			

									
																		
									'The following code does not work if the FieldName = NAME
									'sValue = rsTemp.Fields(SubStr)
						
									rsTemp.Close
									Set rsTemp = Nothing						
							else
							
								sValue = "Syntax Error on SQL Code"
							end if
						
						end if
					elseif (Lookup = VALUES_TABLE) and (m_bTableLookUp = True) then
					
						flag=true
						LookupParams=LookUpTableGetParams(field.name)
						pos=InStr(1,LookupParams,"(" )
						Redim case_labels(0)
						Redim case_values(0)
						While Flag=true and pos > 0 
							nextparam=InStr(pos+1,LookupParams,";")
							if nextparam=0 then
								nextparam=InStr(pos+1,LookupParams,")")
						    end if
						    if nextparam > 0 then
								substr=Mid(LookupParams,pos+1,(nextparam-1)-pos)
								if InStr(1,SubStr,"|")=0 then
									Redim Preserve case_labels(UBound(case_labels)+1)
									case_labels(UBound(case_labels))= SubStr										
								else
									Redim Preserve case_labels(UBound(case_labels)+1)
									case_labels(UBound(case_labels))=Mid(SubStr,1,InStr(1,SubStr,"|")-1)

									Redim Preserve case_values(UBound(case_values)+1)
									case_values(UBound(case_values))=Mid(SubStr,InStr(1,SubStr,"|")+1, Len(SubStr))										 									

									
								end if
							end if
							
							pos=InStr(pos+1,LookupParams,";")
							if pos=0 then
								flag=false
							end if
						Wend
						
						'Now print the values
						if UBound(case_values)>0 then
							sValue=Field.value
							for i=1 to UBound(case_Labels)
								if Cstr(case_labels(i))=CStr(field.value) then
									sValue=case_values(i) & " <font size=1 color=blue>(Filter Value =  " & CStr(Field) &  ")</font>"
									exit for
								end if							
							next												
						
						else
							sValue=Field.value
							for i=1 to UBound(case_Labels)
								if i=CInt(field.value) then
									sValue=case_labels(i) & " <font size=1 color=blue>(Filter Value =  " & CStr(Field) &  ")</font>"
									exit for
								end if							
							next 
							
						
						end if
					
					else
                
						sValue = ClearAllTags(RTESafe(Field.value))
						if sValue = "" then sValue = "&nbsp;"
						if Len(sValue) > 50 then
							sValue=Mid(sValue,1,50)+"<font color=blue> ....more </font>"
						end if
					end if
                end if
        end select
        sField = sValue
    end if

    InsertFieldValue = sField

End Function


''--------------------------------------------------------------------------
'' Name:     InsertFieldDefinition
''           =====================
'' 
'' Returns HTML code for field definition
''
'' Parameter: 
''      Field           field object
''      bIsPrimaryKey   this is a primary key field
''      sStyle          name of CSS style class
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Function InsertFieldDefinition ( Field, bIsPrimaryKey, sStyle ) 
    Dim sReturn

    sReturn = _
        "<tr>" & vbCrLf & _
        "<td class=""" & sStyle & """>"

    if bIsPrimaryKey then sReturn = sReturn & "<i>"
    sReturn = sReturn & field.name
    if bIsPrimaryKey then sReturn = sReturn & "</i>"
    sReturn = sReturn & "</td>" & vbCrLf & _

        "<td class=""" & sStyle & """>" & GetTypeString(field.type) & "</td>" & vbCrLf & _
        "<td class=""" & sStyle & """>" & CStr(field.definedsize) & "</td>" & vbCrLf & _
        "<td class=""" & sStyle & """>" & CStr(field.precision) & "</td>" & vbCrLf & _
        "<td class=""" & sStyle & """>" & GetAttributesString(field.attributes) & "</td>" & vbCrLf & _
        "</tr>" & vbCrLf

    InsertFieldDefinition = sReturn
End Function


''--------------------------------------------------------------------------
'' Name:     getNavigation
''           =============
'' 
'' Returns HTML code for navigation within table.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function getNavigation

    Dim sNavPage, sNavStatistic, sNavPageSize

    Dim s, sValue
    Dim i, nTo, nFrom, nColCount

    Dim b10, b25, b50, bAll, bPrev, bNext, bCurrent

    nColCount = UBound(m_PrimaryKeyFields) + UBound(m_StandardFields)
    if not m_bReadOnly then nColCount = nColCount + 1

    b10   = (m_nPageSize <> 10)
    b25   = (m_nPageSize <> 25)
    b50   = (m_nPageSize <> 50)
    bAll  = (m_nPageSize < m_RS.RecordCount)
    bPrev = (m_nPage > 1)
    bNext = (m_nPage < m_RS.PageCount)

    nFrom = 0
    if m_RS.RecordCount > 0 then nFrom = 1 + ((m_nPage - 1) * m_nPageSize)
    nTo   = m_nPageSize + ((m_nPage - 1) * m_nPageSize)
    if nTo > m_RS.RecordCount then
        nTo = m_RS.RecordCount
    end if


    ' page navigation
    sNavPage = _
        "<table><tr>" & vbCrLf & _
        "<td class=""ute_navigation"">" & STR_PAGES & "</td>" & vbCrLf & _
        "<td class=""ute_navigation"">" & vbCrLf

    s = Request.QueryString
    if bPrev then
        s = getLink(m_sUTEScript, s, sParamPage, CStr(m_nPage-1))
        sNavPage = sNavPage & "<a href=""" & s & """ title=""" & STR_PREV_PAGE & """ " & _
            "class=""ute_link"">&lt;&lt;</a>&nbsp;" & vbCrLf
    else
        sNavPage = sNavPage & "<span class=""ute_navigation_passive"">&lt;&lt;</span>&nbsp;" & vbCrLf
    end if

    sNavPage = sNavPage & "<select class=""ute_navigation"" name=""pages"" onChange=""jumpPage('parent', this, 0)"">" & vbCrLf 
    for i = 1 to m_RS.PageCount
        bCurrent = (i = m_nPage)
        s = getLink(m_sUTEScript, s, sParamPage, CStr(i))
        if not bCurrent then
            sNavPage = sNavPage & "<option value='" & s & "'>" & i & "</option>" & vbCrLf
        else
            sNavPage = sNavPage & "<option value='" & s & "' selected>" & i & "</option>" & vbCrLf
        end if
    next
    sNavPage = sNavPage & "</select>" & vbCrLf

    if bNext then
        s = getLink(m_sUTEScript, s, sParamPage, CStr(m_nPage+1))
        sNavPage = sNavPage & "<a href=""" & s & """ title=""" & STR_NEXT_PAGE & """ " & _
            "class=""ute_link"">&gt;&gt;</a>" & vbCrLf
    else
        sNavPage = sNavPage & "<span class=""ute_navigation_passive"">&gt;&gt;</span>" & vbCrLf
    end if

    sNavPage = sNavPage & _
        "</td>" & vbCrLf & _
        "</tr></table>" & vbCrLf


    ' statistics
    sNavStatistic = Replace(STR_RECORDS,   "%1", CStr(nFrom))
    sNavStatistic = Replace(sNavStatistic, "%2", CStr(nTo))
    sNavStatistic = Replace(sNavStatistic, "%3", CStr(m_RS.RecordCount))


    ' page size
    sNavPageSize = _
        "<table><tr>" & vbCrLf & _
        "<td class=""ute_navigation"">" & STR_REC_COUNT & "</td>" & vbCrLf & _
        "<td class=""ute_navigation"">" & vbCrLf

    sNavPageSize = sNavPageSize & "<select class=""ute_navigation"" name=""psize"" onChange=""jumpPage('parent', this, 0)"">" & _
                    vbCrLf 

    s = Request.QueryString
    if m_RS.RecordCount > 10 then
        s = getLink(m_sUTEScript, s, sParamPage,     "1")
        s = getLink(m_sUTEScript, s,  sParamPageSize, "10")
        if b10 then
            sNavPageSize = sNavPageSize & "<option value='" & s & "'>10</option>" & vbCrLf
        else
            sNavPageSize = sNavPageSize & "<option value='" & s & "' selected>10</option>" & vbCrLf
        end if
    end if
        
    if m_RS.RecordCount > 25 then
        s = getLink(m_sUTEScript, s, sParamPage,     "1")
        s = getLink(m_sUTEScript, s,  sParamPageSize, "25")
        if b25 then
            sNavPageSize = sNavPageSize & "<option value='" & s & "'>25</option>" & vbCrLf
        else
            sNavPageSize = sNavPageSize & "<option value='" & s & "' selected>25</option>" & vbCrLf
        end if
    end if
        
    if m_RS.RecordCount > 50 then
        s = getLink(m_sUTEScript, s, sParamPage,     "1")
        s = getLink(m_sUTEScript, s,  sParamPageSize, "50")
        if b50 then
            sNavPageSize = sNavPageSize & "<option value='" & s & "'>50</option>" & vbCrLf
        else
            sNavPageSize = sNavPageSize & "<option value='" & s & "' selected>50</option>" & vbCrLf
        end if
    end if

    if m_RS.RecordCount <> 0 then
        s = getLink(m_sUTEScript, s, sParamPage,     "1")
        s = getLink(m_sUTEScript, s,  sParamPageSize, CStr(m_RS.RecordCount))
        if bAll then
            sNavPageSize = sNavPageSize & "<option value='" & s & "'>" & STR_ALL & "</option>" & vbCrLf
        else
            sNavPageSize = sNavPageSize & "<option value='" & s & "' selected>" & STR_ALL & "</option>" & vbCrLf
        end if
    end if

    sNavPageSize = sNavPageSize & "</select>" & vbCrLf

    sNavPageSize = sNavPageSize & _
        "</td>" & vbCrLf & _
        "</tr></table>" & vbCrLf


    ' build everything
    sValue = _
        "<table class=""ute_table"">" & vbCrLf & _
        "<tr><td class=""ute_navigation"" colspan=""" & CStr(nColCount) & """>" & vbCrLf & _
        "<table width=""100%""><tr>" & vbCrLf & _
        "<td width=""33%"" class=""ute_navigation"" align=""left"">"   & sNavPage      & "</td>" & vbCrLf & _
        "<td width=""33%"" class=""ute_navigation"" align=""center"">" & sNavStatistic & "</td>" & vbCrLf & _
        "<td width=""33%"" class=""ute_navigation"" align=""right"">"  & sNavPageSize  & "</td>" & vbCrLf & _
        "</tr></table>" & vbCrLf & _
        "</td></tr>" & vbCrLf & _
        "</table>"

    getNavigation = sValue

End Function


''--------------------------------------------------------------------------
'' Name:     getToolBar
''           ==========
'' 
'' Returns HTML code for the toolbar on top of the table.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function getToolBar()
    Dim s, sValue

    sValue = ""

    if m_bListTables or _
       m_bFilters or _
       m_bShowExportLink or _
       m_bShowSQL or _
       m_bShowDefLink then

        sValue = _
            "<table class=""ute_table"">" & vbCrLf & _
            "<tr><td class=""ute_toolbar"">" & vbCrLf

        ' list all tables within database
        if m_bListTables then
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamTable, "")
            s = getLink(m_sUTEScript, s, sParamMode,  MD_DATABASE)
            s = RemoveParameter(s, sParamPage)
            s = RemoveParameter(s, sParamPageSize)
            s = RemoveParameter(s, sParamSQL)
            s = RemoveParameter(s, sParamDefs)
            s = RemoveParameter(s, sParamSort)
            s = RemoveParameter(s, sParamSortDir)
            s = RemoveParameter(s, sParamSubmitted)
            s = RemoveParameter(s, sParamFilterCount)
            s = RemoveCountedParameters(s, sParamSort,          1)
            s = RemoveCountedParameters(s, sParamSortDir,       1)
            s = RemoveCountedParameters(s, sParamPKey,          1)
            s = RemoveCountedParameters(s, sParamFilterField,   1)
            s = RemoveCountedParameters(s, sParamFilterCompare, 1)
            s = RemoveCountedParameters(s, sParamFilterValue,   1)
            s = RemoveCountedParameters(s, sParamFilterCombine, 1)

            sValue = sValue & _
                "<a href=""" & s & """ " & _
                    "onMouseover=""SelectImage('Database_up','Database');""" & _
                    "onMouseout=""SelectImage('Database_normal','Database');"">" & _
                        "<img src=""" & m_sIMAGEDir & "database.gif"" name=""Database"" border=""0"" " & _
                        "height=""25"" width=""25"" title=""" & STR_LIST_TABLES & """>" & _
                "</a>" & vbCrLf
        end if

        ' define filter
        if m_bFilters then
            if UBound(m_FilterCompares) <> 0 then
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamMode, MD_FILTER)
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Filter_down','Filter');""" & _
                        "onMouseout=""SelectImage('Filter_down_sel','Filter');"">" & _
                        "<img src=""" & m_sIMAGEDir & "filter_down_sel.gif"" name=""Filter"" border=""0"" " & _
                        "height=""25"" width=""25"" title=""" & STR_FILTER & """>" & _
                    "</a>" & vbCrLf
            else
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamMode, MD_FILTER)
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Filter_up','Filter');""" & _
                        "onMouseout=""SelectImage('Filter_normal','Filter');"">" & _
                        "<img src=""" & m_sIMAGEDir & "filter.gif"" name=""Filter"" border=""0"" " & _
                        "height=""25"" width=""25"" title=""" & STR_FILTER & """>" & _
                    "</a>" & vbCrLf
            end if
        end if

        ' export data link
        if m_bShowExportLink then
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamMode, MD_EXPORT)
            sValue = sValue & _
                "<a href=""" & s & """ " & _
                    "onMouseover=""SelectImage('Down_up','Down');""" & _
                    "onMouseout=""SelectImage('Down_normal','Down');"">" & _
                    "<img src=""" & m_sIMAGEDir & "down.gif"" name=""Down"" border=""0"" " & _
                    "height=""25"" width=""25"" title=""" & STR_EXPORT & """>" & _
                "</a>" & vbCrLf
        end if

        ' show/hide sql statement
        if (m_bShowSQL) and (m_RS.RecordCount > 0) then
            if m_bViewSQL then
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamSQL, "0")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('SQL_down','SQL');""" & _
                        "onMouseout=""SelectImage('SQL_down_sel','SQL');"">" & _
                            "<img src=""" & m_sIMAGEDir & "sql_down_sel.gif"" name=""SQL"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_SQL_HIDE & """>" & _
                    "</a>" & vbCrLf
            else
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamSQL, "1")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('SQL_up','SQL');""" & _
                        "onMouseout=""SelectImage('SQL_normal','SQL');"">" & _
                            "<img src=""" & m_sIMAGEDir & "sql.gif"" name=""SQL"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_SQL_SHOW & """>" & _
                    "</a>" & vbCrLf
            end if
        end if

        ' show/hide field definitions
        if (m_bShowDefLink) and (m_RS.RecordCount > 0) then
            if m_bViewDefinitions then
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamDefs, "0")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Definition_down','Definition');""" & _
                        "onMouseout=""SelectImage('Definition_down_sel','Definition');"">" & _
                            "<img src=""" & m_sIMAGEDir & "definition_down_sel.gif"" name=""Definition"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_DEF_HIDE & """>" & _
                    "</a>" & vbCrLf
            else
                s = Request.QueryString
                s = getLink(m_sUTEScript, s, sParamDefs, "1")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Definition_up','Definition');""" & _
                        "onMouseout=""SelectImage('Definition_normal','Definition');"">" & _
                            "<img src=""" & m_sIMAGEDir & "definition.gif"" name=""Definition"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_DEF_SHOW & """>" & _
                    "</a>" & vbCrLf
            end if
        end if


		sValue = sValue + "<div align=right><a href=" + m_sMainMenu + "><Font size=2 face=tahoma color=black><b>" & STR_BACK_TO_MAIN_MENU & "</font></B></a></div>"
        sValue = sValue & _
            "</td></tr>" & vbCrLf & _
            "</table>"
    end if

    getToolBar = sValue

End Function


''--------------------------------------------------------------------------
'' Name:     buildHTML_Table
''           ===============
'' 
'' Creates entire UTE HTML code for table view mode.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildHTML_Table()
    Dim i
    Dim nCounter
    Dim s
    Dim sStyle
    Dim sValue
    Dim MaxFieldsToShow

    ' add javascript code
    sValue = _
        "<script language=""JavaScript"">" & vbCrLf & _
        "<!--" & vbCrLf & _
        "   browserName = navigator.appName;" & vbCrLf & _
        "   browserVer = parseInt(navigator.appVersion);" & vbCrLf & _
        vbCrLf & _
        "   if (browserName == ""Netscape"" && browserVer >= 3)" & vbCrLf & _
        "   {" & vbCrLf & _
        "   version = ""n3"";" & vbCrLf & _
        "   }" & vbCrLf & _
        "   else if (browserName == ""Microsoft Internet Explorer"" && browserVer >= 3)" & vbCrLf & _
        "   {" & vbCrLf & _
        "   version = ""n3"";" & vbCrLf & _
        "   }" & vbCrLf & _
        "   else version = ""n2"";" & vbCrLf & _
        vbCrLf & _
        "   if (version == ""n3"")" & vbCrLf & _
        "   {" & vbCrLf & _
        "       // create image objects" & vbCrLf & _
        "       Add_normal              = new Image();" & vbCrLf & _
        "       Add_active              = new Image();" & vbCrLf & _
        "       Edit_normal             = new Image();" & vbCrLf & _
        "       Edit_active             = new Image();" & vbCrLf & _
        "       Delete_normal           = new Image();" & vbCrLf & _
        "       Delete_active           = new Image();" & vbCrLf & _
        "       Database_normal         = new Image();" & vbCrLf & _
        "       Database_up             = new Image();" & vbCrLf & _
        "       Filter_normal           = new Image();" & vbCrLf & _
        "       Filter_up               = new Image();" & vbCrLf & _
        "       Filter_down             = new Image();" & vbCrLf & _
        "       Filter_down_sel         = new Image();" & vbCrLf & _
        "       Down_normal             = new Image();" & vbCrLf & _
        "       Down_up                 = new Image();" & vbCrLf & _
        "       SQL_normal              = new Image();" & vbCrLf & _
        "       SQL_up                  = new Image();" & vbCrLf & _
        "       SQL_down                = new Image();" & vbCrLf & _
        "       SQL_down_sel            = new Image();" & vbCrLf & _
        "       Definition_normal       = new Image();" & vbCrLf & _
        "       Definition_up           = new Image();" & vbCrLf & _
        "       Definition_down         = new Image();" & vbCrLf & _
        "       Definition_down_sel     = new Image();" & vbCrLf & _
        vbCrLf & _
        "       // assign actual images to image objects" & vbCrLf & _
        "       Add_normal.src          = """ & m_sIMAGEDir & "btnAdd.gif"";" & vbCrLf & _
        "       Add_active.src          = """ & m_sIMAGEDir & "btnAddSel.gif"";" & vbCrLf & _
        "       Edit_normal.src         = """ & m_sIMAGEDir & "btnEdit.gif"";" & vbCrLf & _
        "       Edit_active.src         = """ & m_sIMAGEDir & "btnEditSel.gif"";" & vbCrLf & _
        "       Delete_normal.src       = """ & m_sIMAGEDir & "btnDelete.gif"";" & vbCrLf & _
        "       Delete_active.src       = """ & m_sIMAGEDir & "btnDeleteSel.gif"";" & vbCrLf & _
        "       Database_normal.src     = """ & m_sIMAGEDir & "database.gif"";" & vbCrLf & _
        "       Database_up.src         = """ & m_sIMAGEDir & "database_up.gif"";" & vbCrLf & _
        "       Filter_normal.src       = """ & m_sIMAGEDir & "filter.gif"";" & vbCrLf & _
        "       Filter_up.src           = """ & m_sIMAGEDir & "filter_up.gif"";" & vbCrLf & _
        "       Filter_down.src         = """ & m_sIMAGEDir & "filter_down.gif"";" & vbCrLf & _
        "       Filter_down_sel.src     = """ & m_sIMAGEDir & "filter_down_sel.gif"";" & vbCrLf & _
        "       Down_normal.src         = """ & m_sIMAGEDir & "down.gif"";" & vbCrLf & _
        "       Down_up.src             = """ & m_sIMAGEDir & "down_up.gif"";" & vbCrLf & _
        "       SQL_normal.src          = """ & m_sIMAGEDir & "sql.gif"";" & vbCrLf & _
        "       SQL_up.src              = """ & m_sIMAGEDir & "sql_up.gif"";" & vbCrLf & _
        "       SQL_down.src            = """ & m_sIMAGEDir & "sql_down.gif"";" & vbCrLf & _
        "       SQL_down_sel.src        = """ & m_sIMAGEDir & "sql_down_sel.gif"";" & vbCrLf & _
        "       Definition_normal.src   = """ & m_sIMAGEDir & "definition.gif"";" & vbCrLf & _
        "       Definition_up.src       = """ & m_sIMAGEDir & "definition_up.gif"";" & vbCrLf & _
        "       Definition_down.src     = """ & m_sIMAGEDir & "definition_down.gif"";" & vbCrLf & _
        "       Definition_down_sel.src = """ & m_sIMAGEDir & "definition_down_sel.gif"";" & vbCrLf & _

        "   }" & vbCrLf & _
        "" & vbCrLf & _
        "   function SelectImage(img_src, img_name)" & vbCrLf & _
        "   {" & vbCrLf & _
        "       if (version == ""n3"")" & vbCrLf & _
        "       {" & vbCrLf & _
        "           imgOn = eval(img_src + "".src"");" & vbCrLf & _
        "           document [img_name].src = imgOn;" & vbCrLf & _
        "       }" & vbCrLf & _
        "   }" & vbCrLf & _
        "" & vbCrLf & _
        "   function jumpPage(targ, selObj, restore)" & vbCrLf & _
        "   {" & vbCrLf & _
        "       eval(targ + "".location='"" + selObj.options[selObj.selectedIndex].value + ""'"");" & vbCrLf & _
        "       if (restore) selObj.selectedIndex=0;" & vbCrLf & _
        "   }" & vbCrLf & _
        "//-->" & vbCrLf & _
        "</script>" & vbCrLf & _
        vbCrLf & _
        "<p><span class=""ute_headline"">" & m_sHeadline & "</span></p>" & vbCrLf & _
        vbCrLf


    ' add toolbar
    sValue = sValue & getToolBar & vbCrLf

    ' display SQL statement
    if m_bViewSQL then
        sValue = sValue & _
            "<table class=""ute_table"">" & vbCrLf & _
            "<tr><td class=""ute_navigation"">" & vbCrLf & _
            m_sSQL & vbCrLf & _
            "</td></tr></table>"  & vbCrLf
    end if

    ' add naviation bar
    sValue = sValue & getNavigation & vbCrLf

    ' add table
    sValue = sValue & "<table class=""ute_table"">" & vbCrLf
    sValue = sValue & "<tr>" & vbCrLf

    if not m_bReadOnly then
        ' add insert record link
        s = Request.QueryString
        s = getLink(m_sUTEScript, s, sParamMode,     MD_FORM)
        s = getLink(m_sUTEScript, s, sParamFormMode, MD_INSERT)
        sValue = sValue & _
        "<td class=""ute_header"" width=""30"">" & vbCrLf & _
        "<a href=""" & s & """ title=""" & STR_INSERT & """ " & _
            "onMouseover=""SelectImage('Add_active','Add');""" & _
            "onMouseout=""SelectImage('Add_normal','Add');"">" & _
            "<img src=""" & m_sIMAGEDir & "btnAdd.gif"" border=""0"" alt=""" & STR_INSERT & """ " & _
            "name=""Add"" width=""12"" height=""12""></a>" & vbCrLf & _
        "</td>" & vbCrLf
    end if 
    
    if UBound(m_StandardFields) > MAX_FIELDS_TO_SHOW then
	   MaxFieldsToShow=MAX_FIELDS_TO_SHOW
	else
	   MaxFieldsToShow=UBound(m_StandardFields)
	end if    

    ' add primary keys fields to header
    for i = 1 to UBound(m_PrimaryKeyFields)
        sValue = sValue & _
            "<td class=""ute_header_pk"">" & vbCrLf & _
            InsertField (m_RS.fields(m_PrimaryKeyFields(i))) & _
            "</td>" & vbCrLf
    next

    ' add standard fields to header
    for i = 1 to MaxFieldsToShow
        select case m_RS.fields(m_StandardFields(i)).Type
            ' MEMO
            case adLongVarChar, adLongVarWChar
                sValue = sValue & _
                    "<td class=""ute_header"" width=""" & CStr(DEF_MEMO_COL_WIDTH) & """>" & vbCrLf & _
                    InsertField (m_RS.fields(m_StandardFields(i))) & _
                    "</td>" & vbCrLf
            case else
                sValue = sValue & _
                    "<td class=""ute_header"">" & vbCrLf & _
                    InsertField (m_RS.fields(m_StandardFields(i))) & _
                    "</td>" & vbCrLf
        end select
    next

    sValue = sValue & "</tr>" & vbCrLf

    ' insert values
    nCounter = 1
    Do Until m_RS.EOF or (nCounter > m_nPageSize)

        if (nCounter mod 2) = 0 then
            sStyle = "ute_content_even"
        else
            sStyle = "ute_content_odd"
        end if

        sValue = sValue & "<tr class=""" & sStyle & """>" & vbCrLf

        if not m_bReadOnly then
            ' add edit record link
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamMode,     MD_FORM)
            s = getLink(m_sUTEScript, s, sParamFormMode, MD_EDIT)
            s = getLink(m_sUTEScript, s, sParamRecord,   CStr(nCounter + ((m_nPage - 1) * m_nPageSize)))
            sValue = sValue & _
                "<td class=""ute_header"" width=""30"">" & vbCrLf & _
                "<a href=""" & s & """ title=""" & STR_EDIT & """ " & _
                    "onMouseover=""SelectImage('Edit_active','Edit" & CStr(nCounter)& "');""" & _
                    "onMouseout=""SelectImage('Edit_normal','Edit" & CStr(nCounter)& "');"">" & _
                    "<img src=""" & m_sIMAGEDir & "btnEdit.gif"" border=""0"" alt=""" & STR_EDIT & """ " & _
                    "name=""Edit" & CStr(nCounter)& """ width=""12"" height=""12""></a>" & vbCrLf
            ' add delete record link
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamMode,     MD_FORM)
            s = getLink(m_sUTEScript, s, sParamFormMode, MD_DELETE)
            s = getLink(m_sUTEScript, s, sParamRecord,   CStr(nCounter + ((m_nPage - 1) * m_nPageSize)))
            sValue = sValue & _
                "<a href=""" & s & """ title=""" & STR_DELETE & """ " & _
                    "onMouseover=""SelectImage('Delete_active','Delete" & CStr(nCounter)& "');""" & _
                    "onMouseout=""SelectImage('Delete_normal','Delete" & CStr(nCounter)& "');"">" & _
                    "<img src=""" & m_sIMAGEDir & "btnDelete.gif"" border=""0"" alt=""" & STR_DELETE & """ " & _
                    "name=""Delete" & CStr(nCounter)& """ width=""12"" height=""12""></a>" & vbCrLf & _
                "</td>" & vbCrLf
        end if

        ' add primary keys field value
        for i = 1 to UBound(m_PrimaryKeyFields)
            sValue = sValue & _
                "<td class=""" & sStyle & "_pk"">" & _
                InsertFieldValue (m_RS.fields(m_PrimaryKeyFields(i))) & _
                "</td>" & vbCrLf
        next

        ' add standard field value
        

		
        for i = 1 to MaxFieldsToShow
            select case m_RS.fields(m_StandardFields(i)).Type
                ' MEMO
                case adLongVarChar, adLongVarWChar
                    sValue = sValue & _
                        "<td class=""" & sStyle & """ width=""" & CStr(DEF_MEMO_COL_WIDTH) & """>" & _
                        InsertFieldValue (m_RS.fields(m_StandardFields(i))) & _
                        "</td>" & vbCrLf
                case else
                    sValue = sValue & _
                        "<td class=""" & sStyle & """>" & _
                        InsertFieldValue (m_RS.fields(m_StandardFields(i))) & _
                        "</td>" & vbCrLf
            end select
        next

        sValue = sValue & "</tr>" & vbCrLf

        nCounter = nCounter + 1
        m_RS.MoveNext
    Loop

    sValue = sValue & "</table>" & vbCrLf

    ' add naviation bar
    sValue = sValue & getNavigation & vbCrLf

    ' add "Powered By"
    sValue = sValue & getPoweredBy & vbCrLf


    ' show field definitions if link is enabled and selected
    if (m_bShowDefLink) and (m_bViewDefinitions) and (m_RS.RecordCount > 0) then

        sValue = sValue & _
            "<p><table class=""ute_table"">" & vbCrLf & _
            "<tr>" & vbCrLf & _
            "<td class=""ute_header"">" & STR_DEF_NAME & "</td>" & vbCrLf & _
            "<td class=""ute_header"">" & STR_DEF_TYPE & "</td>" & vbCrLf & _
            "<td class=""ute_header"">" & STR_DEF_DEFINEDSIZE & "</td>" & vbCrLf & _
            "<td class=""ute_header"">" & STR_DEF_PRECISION & "</td>" & vbCrLf & _
            "<td class=""ute_header"">" & STR_DEF_ATTRIBUTES & "</td>" & vbCrLf & _
            "</tr>" & vbCrLf

        m_RS.MoveFirst
        nCounter = 0

        if UBound(m_PrimaryKeyFields) > 0 then
            for i = 1 to UBound(m_PrimaryKeyFields)
                nCounter = nCounter + 1
                if (nCounter mod 2) = 0 then
                    sStyle = "ute_content_even"
                else
                    sStyle = "ute_content_odd"
                end if
                sValue = sValue & InsertFieldDefinition (m_RS.fields(m_PrimaryKeyFields(i)), True, sStyle)
            next
        end if

        if UBound(m_StandardFields) > 0 then
            for i = 1 to UBound(m_StandardFields)
                nCounter = nCounter + 1
                if (nCounter mod 2) = 0 then
                    sStyle = "ute_content_even"
                else
                    sStyle = "ute_content_odd"
                end if
                sValue = sValue & InsertFieldDefinition (m_RS.fields(m_StandardFields(i)), False, sStyle)
            next
        end if

        sValue = sValue & "</table></p>"

    end if

    buildHTML_Table = sValue

End Function

%>