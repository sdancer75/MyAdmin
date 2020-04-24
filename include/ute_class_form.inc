
<%




''--------------------------------------------------------------------------
'' Name:     getSQLStatement
''           ==================
'' 
'' Returns SQL Statement to select/delete specific record. The statement is
'' compiled by the hidden form "ident" fields.
''
'' Parameter: 
''      sCmd        "select *" or "delete"
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
Private Function getSQLStatement ( sCmd )
    Dim i
    Dim bFirst
    Dim sSQL, sField, nType, sValue
    Dim curField

    i = 1
    bFirst = True
    sSQL = sCmd & " FROM " & m_sTable
    while Request.Form(sFormUTEFieldPrefix & sFormIdentField & CStr(i)) <> ""
        sField = Request.Form(sFormUTEFieldPrefix & sFormIdentField & CStr(i))
        nType  = CInt(Request.Form(sFormUTEFieldPrefix & sFormIdentType & CStr(i)))
        sValue = Request.Form(sFormUTEFieldPrefix & sFormIdentValue & CStr(i))
        Set curField = m_RS(sField)
        sSQL = sSQL & AddWhere (sField, nType, sValue, "=", "AND", bFirst)
        bFirst = False
        i = i + 1
    wend

    getSQLStatement = sSQL

End Function


''--------------------------------------------------------------------------
'' Name:     IsErrorField
''           ============
'' 
'' Is this an Error Field ?
''
'' Parameter: 
''      sName       name of field to be checked
''
'' return value:
''      boolean
''
''--------------------------------------------------------------------------
Private Function IsErrorField ( sName )
    Dim i, bReturn
    bReturn = False
    for i = 1 to UBound(m_ErrorField)
        if m_ErrorField(i) = sName then bReturn = True
    next
    IsErrorField = bReturn
End Function


''--------------------------------------------------------------------------
'' Name:     GetErrorNumber
''           ==============
'' 
'' Returns the Error Description
''
'' Parameter: 
''      sName       name of field the err descr. should be returned
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
Private Function GetErrorDescription ( sName )
    Dim i, sReturn
    sReturn = ""
    for i = 1 to UBound(m_ErrorField)
        if m_ErrorField(i) = sName then sReturn = sReturn & m_ErrorMessage(i) & "<br>"
    next
    ' cut traling <br>
    if sReturn <> "" then sReturn = Left(sReturn, Len(sReturn)-4)
    GetErrorDescription = sReturn
End Function


''--------------------------------------------------------------------------
'' Name:     PutError
''           ========
'' 
'' Put Error into Error Management
''
'' Parameter: 
''      errField        name of field that returns an error
''      errNumber       error number
''      errMessage      error description
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub PutError ( errField )
    Dim e
    Dim nError
    ' the errors collections of the connection object contains all occured errors
    For Each e In m_DB.Errors
        nError = UBound(m_ErrorField) + 1
        Redim Preserve m_ErrorField(nError)
        m_ErrorField(nError)   = errField
        Redim Preserve m_ErrorMessage(nError)
        m_ErrorMessage(nError) = e.Description
    Next
    Err.Clear
    m_DB.Errors.Clear
End Sub


''--------------------------------------------------------------------------
'' Name:     InsertFieldForm
''           ===============
'' 
'' Return HTML code for a single field (incl, heading, form element and definitions
''
'' Parameter: 
''      field           field object
''      bPrimaryKey     this is a primary key field
''			
''
'' return value:
''      string          HTML code
''
''--------------------------------------------------------------------------
Private Function InsertFieldForm ( field, bPrimaryKey )

    Dim sReturn, sValue
    Dim sStyle, sStyleForm
    Dim nSize, nMaxLength
    Dim nMemoCols, nMemoRows
    Dim sChecked
    Dim Attrib,AllowNull,Lookup,flag,pos,LookupParams,nextparam,substr,sLabel
    Dim rsTemp,getParam1,getParam2,ParamError,ParamStr,i,pkvalue,ForceShowAll,sUserID
   

    sStyle = "ute_header"
    if bPrimaryKey then sStyle = sStyle & "_pk"

    sStyleForm = "ute_form_value"
    if IsErrorField (field.name) then sStyleForm = "ute_form_error"

    ' field value
    if m_bSubmitted then
        if bPrimaryKey then
            if IsNull(field.value) then
                sValue = ""
            else
                select case field.type
                    case else
                        sValue = CStr(field.value)
                end select
            end if
        else
            sValue = Trim(Request.Form("z"+field.name))
        end if
    else
		if m_nFormMode = 1 then 'if new record then set the default values if they exists
		
			sValue = Cstr(GetDefaultValue(field.name))
		else
			if IsNull(field.value) then
			    sValue = ""
			else
			    select case field.type
			        case else
			            sValue = CStr(field.value)
			    end select
			end if
		end if
        
    end if
    ' html endcode field value
    sValue = Server.HTMLEncode(sValue)

    sLabel=FieldHasALabel(Field.name)
	if sLabel<>"" then
		sReturn  =  "<tr>" &  "<td class=""" & sStyle & """>" & sLabel & "</td>" & "<td class=""" & sStyleForm & """>"
	else
		sReturn  =  "<tr>" &  "<td class=""" & sStyle & """>" & field.name & "</td>" & "<td class=""" & sStyleForm & """>"
	end if
	
    

    if  (NotAttrib(field.Attributes, adFldUpdatable) and NotAttrib(field.Attributes, adFldUnknownUpdatable)) or _ 
        ((NotAttrib(field.Attributes, adFldUpdatable) and bPrimaryKey)) or _
        (m_nFormMode = MD_DELETE) then 
        ' this field is not editable

        if IsExcluded(field.type) then
            ' this field is not editable by ute
            sReturn = sReturn & _
                "<img src=""" & m_sIMAGEDir & "exclude.gif"" border=""0"" alt=""" & STR_NON_VIEW & """ " & _
                "width=""16"" height=""16"">"
        else
            ' display field value
            sReturn = sReturn & sValue
        end if
    else
        ' this field is editable

        select case field.type
            ' VARCHAR
            case adBSTR, adVariant, adChar, adVarChar, adWChar, adVarWChar
                nMaxLength = field.DefinedSize
                if nMaxLength > DEF_MAX_INPUT_LENGTH then 
                    nSize = DEF_MAX_INPUT_LENGTH
                else
                    nSize = nMaxLength
                end if

            ' MEMO
            case adLongVarChar, adLongVarWChar
                nMemoCols = DEF_MEMO_COLS
                nMemoRows = DEF_MEMO_ROWS

            ' ELSE 
            case else
                nMaxLength = field.Precision
                if nMaxLength > DEF_MAX_INPUT_LENGTH then
                    nSize = DEF_MAX_INPUT_LENGTH
                else
                    nSize = nMaxLength
                end if

        end select
        
        Attrib=GetAttributesString(field.attributes)
        AllowNull=InStr(1,Attrib,Cstr(STR_ADO_ISNULLABLE))
   
        if IsExcluded(field.type) then
            sReturn = sReturn & _
                "<img src=""" & m_sIMAGEDir & "exclude.gif"" border=""0"" alt=""" & STR_NON_VIEW & """ " & _
                "width=""16"" height=""16"">"
        else
            if  ((field.type = adLongVarChar) or (field.type = adLongVarWChar) or (nMaxLength > FIELD_MAX_LEN_TO_MAKE_TEXTAREA)) and (field.type <> adBoolean) then
                ' MEMO -> TEXTAREA
                flag=false
                for i=1 to UBound(m_FieldsRichEdit)
					if ucase(m_FieldsRichEdit(i)) = ucase(Field.name) then
						flag=true
						exit for
					end if
                next
                
                
                if flag=true then
                
					sReturn=sReturn & vbCrLf & "<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf &_
					"function submitForm() {" & vbCrLf &_
					   	"updateRTEs();" & vbCrLf &_
					   	"return true;" & vbCrLf &_
					"}" & vbCrLf & vbCrLf &_
					"initRTE(""images/"", """", """", false);" & vbCrLf &_
					"</script>" & vbCrLf &_
					"<script language=""JavaScript"" type=""text/javascript"">" & vbCrLf &_
					"writeRichText('z" & field.name & "', '" & RTESafe(field.value) & "', 520, 200, true, false);" & vbCrLf &_
					"</script>" & vbCrLf
					
                
                else
                
					sReturn = sReturn & "<textarea name=""z" & field.name & """ cols=" & CStr(DEF_MEMO_COLS) & _
					    " rows=" & CStr(DEF_MEMO_ROWS) & ">" & sValue & "</textarea>"
				end if
                
                
                
                
                if AllowNull=0 then
                
                sReturn=sReturn+"<IMG SRC=""images/icon_required.gif"" ALIGN=Middle HSPACE=2>"
                
                end if
            elseif (field.type = adBoolean) then
                ' -> CHECKBOX
                sChecked = ""
                if CBool(field.value) then sChecked = " checked"
                sReturn = sReturn & "<input type=""checkbox"" name=""z" & field.name & """" & sChecked & ">" 
                if AllowNull=0 then
                
                sReturn=sReturn+"<IMG SRC=""images/icon_required.gif"" ALIGN=Middle HSPACE=2>"
                
                end if 		
            elseif (field.type = adCurrency) then
                ' ->Currency
                sReturn = sReturn & "<input type=""text"" name=""z" & field.name & """ & value=""" & (sValue) & """" & ">" 
                if AllowNull=0 then
                
                sReturn=sReturn+"<IMG SRC=""images/icon_required.gif"" ALIGN=Middle HSPACE=2>"
                
                end if 		                	
            else
                ' DropDown Lookup Table Menu
                Lookup=LookupTable(field.name)
                if  Lookup <> NO_LOOKUP_TABLE then
					'Values LookupTable
					if Lookup = VALUES_TABLE then
						sReturn = sReturn & "<select id=""z" & field.name & """ name=""z" & field.name & """>"					
						if AllowNull<>0 then
							sReturn = sReturn & "<option value="""">-- " & STR_SELECT_DROPDOWN & " --</option>"
						end if
						flag=true
						LookupParams=LookUpTableGetParams(field.name)
						pos=InStr(1,LookupParams,"(" )
						While Flag=true and pos > 0 
							nextparam=InStr(pos+1,LookupParams,";")
							if nextparam=0 then
								nextparam=InStr(pos+1,LookupParams,")")
						    end if
						    if nextparam > 0 then
								substr=Mid(LookupParams,pos+1,(nextparam-1)-pos)
								if InStr(1,SubStr,"|")=0 then
									sReturn = sReturn & "<option value=""" & SubStr  
									if sValue=SubStr then
										sReturn = sReturn & " selected "
									end if
									sReturn = sReturn & """>" &  SubStr & "</option>"								
								else
									if Trim(ucase(Mid(SubStr,1,InStr(1,SubStr,"|")-1))) = Trim(ucase(sValue)) then
										sReturn = sReturn & "<option selected value=""" & Mid(SubStr,1,InStr(1,SubStr,"|")-1) & """>" & Mid(SubStr,InStr(1,SubStr,"|")+1, Len(SubStr)) & "</option>"
									else
										sReturn = sReturn & "<option value=""" & Mid(SubStr,1,InStr(1,SubStr,"|")-1) & """>" & Mid(SubStr,InStr(1,SubStr,"|")+1, Len(SubStr)) & "</option>"																	
									end if
								end if
							end if
							
							pos=InStr(pos+1,LookupParams,";")
							if pos=0 then
								flag=false
							end if
						Wend
						sReturn=sReturn & "</select>"
					elseif Lookup = SQL_QUERY then
					'SQL LookupTable

						LookupParams=LookUpTableGetParams(field.name)
						pos=InStr(1,LookupParams,"(" )
						substr=Mid(LookupParams,pos+1,Len(LookupParams)-pos-1 )
						if substr <> "" Then						
							
							getParam1=InStr(1,ucase(SubStr),ucase("showall"))
							ForceShowAll=CBool(getParam1 > 0)
							'if non admin the show all the records specified in the where clause
							'if the showall tag exist then force admin mode
							
							If (LoginAdmin=False) and (ForceShowAll=false) then
									
									'replace the v_Field{field} parameter
									ParamError=False
									getParam1=InStr(1,ucase(SubStr),ucase("v_field{"))
									if getParam1 > 0 then
										
										getParam2=InStr(1,ucase(SubStr),ucase("}"))
										if getParam2 > 0 then
											ParamStr=Mid(SubStr,getParam1,(getParam2-getParam1)+1)
											SubStr=Replace(SubStr,ParamStr,"")
											
											'Get & Remove pkvalue

											getParam1=InStr(1,ucase(SubStr),ucase("pkvalue{"))
											getParam1=getParam1+Len("pkvalue{")
											if getParam1 > 0 then										
													getParam2=InStr(getparam1,ucase(SubStr),ucase("}"))
													if getParam2 > 0 then
														pkvalue=Mid(SubStr,getParam1,(getParam2-getParam1))	
														getParam1=InStr(1,ucase(SubStr),ucase("pkvalue{"))											
														SubStr=Replace(SubStr,Mid(SubStr,getParam1,(getParam2-getParam1)+1),"")
													end if
											end if	
											
		
											
											
											getParam1=InStr(1,ParamStr,"{")
											if getParam1 > 0 then
												getParam2=InStr(1,ParamStr,"}")
												if getParam2 > 0 then
													ParamStr=Mid(ParamStr,getParam1+1,(getParam2-getParam1)-1)
													if ucase(Mid(ParamStr,1,InStr(1,ParamStr,":")-1)) = ucase(Field.Name) then
														'if new record in user mode then 
														'display his own records only
														if (LoginAdmin=False) and (ForceShowAll=false) then
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
																	
																	'Remove WHERE clause
																	getParam1=InStr(1,ucase(SubStr),ucase("where"))																	
																	if getParam1 > 0 then																									
																			SubStr=Replace(SubStr,Mid(SubStr,getParam1,(Len(SubStr)-getParam1)+1),"")														
																	end if
																	
																	SubStr=SubStr + "where " + f_RelationField + "="
																	Set rsTemp = Server.CreateObject("ADODB.Recordset")																	
																	rsTemp.Open "select " & f_RelationField & " from " & tbl_UsersTable & " where id=" & f_UserID, m_DB, adOpenStatic,adLockReadOnly														
																	if rsTemp.EOF = false then
																		sUserID=rsTemp.Fields.Item(f_RelationField).value
																		SubStr = SubStr + Cstr(sUserID)
																	end if
																	rsTemp.close
																	Set rsTemp = Nothing																	
																																					
																	'SubStr=SubStr & Cstr(Field.Value)
																end if																		
																
														
														else										
															SubStr=SubStr & Cstr(Field.Value)
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
							else
							'if Admin then he can see all - Delete the Where clause
									ParamError=False
									getParam1=InStr(1,ucase(SubStr),ucase("v_field{"))
									if getParam1 > 0 then
										
										getParam2=InStr(1,ucase(SubStr),ucase("}"))
										if getParam2 > 0 then
											ParamStr=Mid(SubStr,getParam1,(getParam2-getParam1)+1)
											SubStr=Replace(SubStr,ParamStr,"")
											
											'Remove ShowAll tag if exists
											getParam1=InStr(1,ucase(SubStr),ucase("showall"))
									
											if getParam1 > 0 then										
												SubStr=Replace(SubStr,Mid(SubStr,getParam1,Len("showall")),"")
											end if												
																		
											getParam1=InStr(1,ParamStr,"{")
											if getParam1 > 0 then
												getParam2=InStr(1,ParamStr,"}")
												if getParam2 > 0 then
													ParamStr=Mid(ParamStr,getParam1+1,(getParam2-getParam1)-1)
												end if											
											end if
											getParam1=InStr(1,ucase(SubStr),ucase("pkvalue{"))
											getParam1=getParam1+Len("pkvalue{")
											if getParam1 > 0 then										
													getParam2=InStr(1,ucase(SubStr),ucase("}"))
													if getParam2 > 0 then
														pkvalue=Mid(SubStr,getParam1,(getParam2-getParam1))												
													end if
											end if											
											SubStr=Mid(SubStr,1,InStr(1,ucase(SubStr),"WHERE")-2)
										end if
									end if
									
									
							
							
							end if
							
							if ParamError=False then
									
									'if forceShowAll and User then showAll from his own records
									if (LoginAdmin=false) and (ForceShowAll) then
									
										substr = substr & " Where " & f_RelationField & " = 120"
										
										
									end if									
									
									Set rsTemp = Server.CreateObject("ADODB.Recordset")
									rsTemp.Open substr, m_DB, adOpenStatic,adLockReadOnly
									

									
									SubStr=	Mid(ParamStr,InStr(1,ParamStr,":")+1,Len(ParamStr)-InStr(1,ParamStr,":"))					
									
									
									sReturn = sReturn & "<select id=""z" & field.name & """ name=""z" & field.name & """>"					
									if AllowNull<>0 then
										sReturn = sReturn & "<option value="""">-- " & STR_SELECT_DROPDOWN & " --</option>"
									end if	
									
									if rsTemp.EOF = False then
										While rsTemp.EOF = False								
											For i = 0 To rsTemp.Fields.Count - 1
												if ucase(rsTemp.Fields.Item(i).Name) = ucase(SubStr) then
													
													if (LoginAdmin=true) or (ForceShowAll) then
																
																sReturn = sReturn & "<option value=""" & Cstr(rsTemp.Fields.Item(pkvalue).value) & """" 
													
																if (rsTemp.Fields.Item(pkvalue).value <> "") then
																	if field.value <> "" then
																
																		if Cstr(rsTemp.Fields.Item(pkvalue).value) = Cstr(field.value) then
																				sReturn = sReturn & " selected "
																		end if
																		
																	end if
																	
																end if
																sReturn = sReturn & ">" &  CStr(rsTemp.Fields.Item(i).value) & "</option>"

													else
																sReturn = sReturn & "<option value=""" & Cstr(rsTemp.Fields.Item(pkvalue)) & """" 
													
																if (sValue=SubStr) or  Cstr(rsTemp.Fields.Item(pkvalue).value) = Cstr(field.value) then
																		sReturn = sReturn & " selected "
																end if
																sReturn = sReturn & ">" &  CStr(rsTemp.Fields.Item(i)) & "</option>"											
													
													end if
																					
													
													
													exit for
												end if
											Next
											rsTemp.MoveNext
										Wend				
									end if

									sReturn=sReturn & "</select>"
																		
									'The following code does not work if the FieldName = NAME
									'sValue = rsTemp.Fields(SubStr)
						
									rsTemp.Close
									Set rsTemp = Nothing						
							else
							
								sReturn = sReturn & "Syntax Error on SQL Code"
							end if					
					
					end if
				 end if	
                
                
                else
                '--> Input for sure
					sReturn = sReturn & "<input type=""text"" name=""z" & field.name & """ maxlength=" & _       
					    CStr(nMaxLength) & " size=" & CStr(nSize) & " value=""" & sValue & """>"
					    
					for i=1 to UBound(m_FieldsHookBrowse)
						if ucase(m_FieldsHookBrowse(i)) = ucase(Field.Name) then
						
						sReturn = sReturn & " <INPUT type=""button"" value=""" & STR_HOOKBROWSE_TITLE & """ id=""btn" & field.name & """ name=""btn" & field.name & """ onclick=""javascript: win=window.open('FileManager/FileManager.asp?field=z" & field.name & "','_blank','toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=830,height=500,left=0,top=0');win.window.focus()"">"
						
						end if
					
					next
                
                end if
                
                
                
                if AllowNull=0 then                
					sReturn=sReturn+"<IMG SRC=""images/icon_required.gif"" ALIGN=Middle HSPACE=2>"                
                end if                    
            end if 
         

            ' put error message into form
            if IsErrorField(field.name) then
                sReturn = sReturn & "&nbsp;" & GetErrorDescription(field.name)
            end if 

        end if 
    end if 
	

	'add HTML code at the end of the field (if exists)	
	for i=1 to UBound(m_FieldsHTMLCode)
		getparam1=InStr(1,m_FieldsHTMLCode(i),"=")
		SubStr=Mid(m_FieldsHTMLCode(i),1, getparam1-1)
		if (InStr(1,ucase(SubStr),ucase(Field.Name)) > 0) and Len(Field.Name)=Len(SubStr) then
			SubStr=Mid(m_FieldsHTMLCode(i),getparam1+1,Len(m_FieldsHTMLCode(i))-getParam1)		
			sReturn = sReturn & SubStr			
			exit for
		end if

	next	
	
    sReturn = sReturn & "</td>" & vbCrLf
    

    if m_bViewDefinitions then
        sReturn = sReturn & _
            "<td class=""ute_form_def"">" & GetTypeString(field.type) & "</td>" & _
            "<td class=""ute_form_def"">" & GetAttributesString(field.attributes) & "</td>"
    end if

    sReturn = sReturn & "</tr>" & vbCrLf

    InsertFieldForm = sReturn

End Function




''--------------------------------------------------------------------------
'' Name:     InsertIdentFields
''           =================
'' 
'' Return HTML code for hidden form fields. These fields contain the field
'' names and values to identify the current record.
''
'' Parameter: 
''      none
''
'' return value:
''      string
''
''--------------------------------------------------------------------------
Private Function InsertIdentFields ()
    Dim i
    Dim nCount
    Dim curField
    Dim sField, sType, sValue
    Dim sReturn

    sReturn = ""

    if m_nFormMode <> MD_INSERT then

        ' do we already have some ident fields in form ?
        if Request.Form(sFormUTEFieldPrefix & sFormIdentField & "1") <> "" then
            ' yes, so use these values
            i = 1
            while Request.Form(sFormUTEFieldPrefix & sFormIdentField & CStr(i)) <> ""
                sField = Request.Form(sFormUTEFieldPrefix & sFormIdentField & CStr(i))
                sType  = Request.Form(sFormUTEFieldPrefix & sFormIdentType & CStr(i))
                sValue = Request.Form(sFormUTEFieldPrefix & sFormIdentValue & CStr(i))
                sReturn = sReturn & _
                    "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentField & CStr(i) & """ " & _
                        "value=""" & sField & """>" & vbCrLf & _
                    "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentType & CStr(i) & """ " & _
                        "value=""" & sType & """>" & vbCrLf & _
                    "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentValue & CStr(i) & """ " & _
                        "value=""" & sValue & """>" & vbCrLf
                i = i + 1
            wend
        else
            ' no, create ident fields from primary key fields or all fields
            nCount = 0
            if UBound(m_PrimaryKeyFields) > 0 then
                ' use Primary Keys for identification
                for i = 1 to UBound(m_PrimaryKeyFields)
                    set curField = m_RS(m_PrimaryKeyFields(i))
                    if (not IsExcluded(curField.Type)) and not (IsNull(curField.Value)) then
                        nCount = nCount + 1
                        if (curField.Type = adBoolean) then
                            if curField.value then
                                sValue = CStr(True)
                            else
                                sValue = CStr(False)
                            end if
                        else
                            sValue = CStr(curField.Value)
                        end if
                        sReturn = sReturn & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentField & CStr(nCount) & _
                                """ value=""" & m_PrimaryKeyFields(i) & """>" & vbCrLf & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentType & CStr(nCount) & _
                                """ value=""" & m_PrimaryKeyTypes(i) & """>" & vbCrLf & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentValue & CStr(nCount) & _
                                """ value=""" & sValue & """>" & vbCrLf
                    end if
                next
            else
                ' use all fields for identification
                for i = 1 to UBound(m_StandardFields)
                    set curField = m_RS(m_StandardFields(i))
                    if (not IsExcluded(curField.Type)) and not (IsNull(curField.Value)) then
                        nCount = nCount + 1
                        if (curField.Type = adBoolean) then
                            if curField.value then
                                sValue = CStr(True)
                            else
                                sValue = CStr(False)
                            end if
                        else
                            sValue = CStr(curField.Value)
                        end if
                        sReturn = sReturn & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentField & CStr(nCount) & _
                                """ value=""" & m_StandardFields(i) & """>" & vbCrLf & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentType & CStr(nCount) & _
                                """ value=""" & m_StandardTypes(i) & """>" & vbCrLf & _
                            "<input type=""hidden"" name=""" & sFormUTEFieldPrefix & sFormIdentValue & CStr(nCount) & _
                                """ value=""" & sValue & """>" & vbCrLf
                    end if
                next
            end if 
        end if 
    end if

    InsertIdentFields = sReturn
End Function


''--------------------------------------------------------------------------
'' Name:     InitForm
''           ========
'' 
'' Initialises a recordset for the form. This is either created as a new
'' (empty) one, or as a copy from the current record from the table view.
''
'' Parameter: 
''      none
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub InitForm ()

    Dim s

    ' if cancel button was pressed redirect back to table
    if Request.Form(sFormUTEFieldPrefix & sFormButton) = STR_CANCEL then
        s = Request.QueryString
        s = getLink(m_sUTEScript, s, sParamMode,      CStr(MD_TABLE))
        s = getLink(m_sUTEScript, s, sParamFormMode,  CStr(DEF_FORM_MODE))
        s = getLink(m_sUTEScript, s, sParamSubmitted, "0")
        Response.Redirect s
    end if

    ' get RecordSet for this form
    if m_nFormMode = MD_INSERT then
        ' create new RecordSet
        Set m_RSForm = Server.CreateObject("ADODB.Recordset")
        m_RSForm.Open m_sTable, m_DB, adOpenStatic, adLockPessimistic, adCmdTable        
        m_RSForm.AddNew
    else
        if not m_bSubmitted then
            ' identify record by "record" URL param
            m_RS.Move m_nRecord - 1, adBookmarkFirst
            Set m_RSForm = m_RS
        else
            ' identify record by hidden form fields
            Set m_RSForm = Server.CreateObject("ADODB.Recordset")
            m_RSForm.Open getSQLStatement("SELECT *"), m_DB, adOpenStatic, adLockPessimistic, adCmdText
        end if
    end if

End Sub


''--------------------------------------------------------------------------
'' Name:     IsFormField
''           ===========
'' 
'' Returns TRUE if thegiven field name is a field on the current submitted form.
''
'' Parameter: 
''      name        name of field
''
'' return value:
''      boolean
''
''--------------------------------------------------------------------------
Private Function IsFormField ( name )
    Dim bReturn
    Dim field
    bReturn = False
    for each field in Request.Form
        if (field = name) then bReturn = True
    next
    IsFormField = bReturn
End Function



''---------------------------------------------------------
''Name	: TwoDigits
''
''Returns a string with 2 digits length
''--------------------------------------------------------

function TwoDigits(zahl)
     zahl = trim(cstr(zahl))
     if len(zahl)=1 then
             twodigits = "0" & zahl
     else
             twodigits = zahl

     end if
end function



''--------------------------------------------------------------------------
'' Name:     UpdateRecordSet
''           ===============
'' 
'' Put form values into record set
''
'' Parameter: 
''      rsUpdt      recordset to be updated
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub UpdateRecordSet ( byref rsUpdt )

    Dim field,DeclaredDate,SQLISOFormat 

    ' own error handling
    On Error Resume Next

    'for each field in Request.Form
    for each field in rsUpdt.fields

        select case field.type
            case adBoolean
                rsUpdt(field.name) = (Request.Form("z"+field.name) = "on")  
                
            case adDBTimeStamp     
                ' is this a field being set in the form ?
                if IsFormField("z"+field.name) then
                    if Request.Form("z"+field.name) = "" then
                        rsUpdt(field.name) = NULL
                    else
						'format dd/mm/yy
						DeclaredDate = Request.Form("z"+field.name)
						'full date / time
						'year(DeclaredDate) & "-"& twodigits(month(DeclaredDate)) & "-" & twodigits(day(DeclaredDate)) & " " & twodigits(hour(now)) & ":" & twodigits(minute(now)) & ":" & twodigits(second(now))												
						
						SQLISOFormat = year(DeclaredDate) & "-"& twodigits(month(DeclaredDate)) & "-" & twodigits(day(DeclaredDate))											
                        
                        rsUpdt(field.name) = SQLISOFormat

                    end if
                end if            
            case else
                ' is this a field being set in the form ?
                if IsFormField("z"+field.name) then
                    if Request.Form("z"+field.name) = "" then                        
                        rsUpdt(field.name) = NULL                          
                    else
                        rsUpdt(field.name) = Request.Form("z"+field.name)
                    end if
                end if
        end select
    
        if Err <> 0 then PutError field.name

    next

    if UBound(m_ErrorField) = 0 then 
        ' make update permanent in DB
        rsUpdt.Update
        if Err <> 0 then PutError ""
    end if

    ' disable own error handling
    On Error Goto 0

End Sub


''--------------------------------------------------------------------------
'' Name:     Update
''           ======
'' 
'' Updates record in database. If successful redirect to ute page in table
'' view mode (else do not redirect, this will lead to display the form and
'' show the error)
''
'' Parameter: 
''      none
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub Update
    Dim s
    
    ' initialise form
    InitForm

    if m_bSubmitted then

        select case m_nFormMode
            case MD_INSERT
                ' Insert New Record
                UpdaterecordSet m_RSForm

            case MD_EDIT
                ' Edit Record
                UpdaterecordSet m_RSForm

            case MD_DELETE
                ' Delete Record
                m_DB.Execute getSQLStatement("DELETE")
        end select

        ' If everything is ok return to table view
        if UBound(m_ErrorField) = 0 then 
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamMode,      CStr(MD_TABLE))
            s = getLink(m_sUTEScript, s, sParamFormMode,  CStr(DEF_FORM_MODE))
            s = getLink(m_sUTEScript, s, sParamSubmitted, "0")
            if m_nPageSize = m_RS.RecordCount-1 then
                s = getLink(m_sUTEScript, s, sParamPageSize, CStr(m_RS.RecordCount-1))
            end if
            Response.Redirect s
        end if

    end if

End Sub


''--------------------------------------------------------------------------
'' Name:     getFormToolBar
''           ==============
'' 
'' Returns HTML code for the toolbar on top of the form.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function getFormToolBar( nColCount )
    Dim s, sValue, sSubmitted

    sValue = ""

    if m_bShowDefLink then

        sValue = "<tr><td class=""ute_toolbar"" colspan=""" & CStr(nColCount)& """>" & vbCrLf

        s = Request.QueryString

        ' show/hide field definitions
        if (m_bShowDefLink) and (m_RS.RecordCount > 0) then
            if m_bViewDefinitions then
                if m_bSubmitted then
                    sSubmitted = "1"
                else
                    sSubmitted = "0"
                end if
                s = getLink(m_sUTEScript, s, sParamSubmitted, sSubmitted)
                s = getLink(m_sUTEScript, s, sParamDefs, "0")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Definition_down','Definition');""" & _
                        "onMouseout=""SelectImage('Definition_down_sel','Definition');"">" & _
                            "<img src=""" & m_sIMAGEDir & "definition_down_sel.gif"" name=""Definition"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_DEF_HIDE & """>" & _
                    "</a>" & vbCrLf
            else
                s = getLink(m_sUTEScript, s, sParamDefs, "1")
                s = getLink(m_sUTEScript, s, sParamSubmitted, "0")
                sValue = sValue & _
                    "<a href=""" & s & """ " & _
                        "onMouseover=""SelectImage('Definition_up','Definition');""" & _
                        "onMouseout=""SelectImage('Definition_normal','Definition');"">" & _
                            "<img src=""" & m_sIMAGEDir & "definition.gif"" name=""Definition"" border=""0"" " & _
                            "height=""25"" width=""25"" title=""" & STR_DEF_SHOW & """>" & _
                    "</a>" & vbCrLf
            end if
        end if

        sValue = sValue & "</td></tr>" & vbCrLf

    end if

    getFormToolBar = sValue

End Function


''--------------------------------------------------------------------------
'' Name:     buildHTML_Form
''           ==============
'' 
'' Creates entire UTE HTML code for form view mode.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildHTML_Form()
    Dim i, nColCount
    Dim s, sValue, sSubmitted,flag

    if m_bViewDefinitions then
        nColCount = 4
    else
        nColCount = 2
    end if

    s = Request.QueryString
    s = getLink(m_sUTEScript, s, sParamSubmitted, "1")

    ' add javascript code
    sValue = _
		"<script language=""JavaScript"" type=""text/javascript"" src=""richtext.js""></script>" & vbCrLf & _
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
        "       Definition_normal       = new Image();" & vbCrLf & _
        "       Definition_up           = new Image();" & vbCrLf & _
        "       Definition_down         = new Image();" & vbCrLf & _
        "       Definition_down_sel     = new Image();" & vbCrLf & _
        vbCrLf & _
        "       // assign actual images to image objects" & vbCrLf & _
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
        "//-->" & vbCrLf & _
        "</script>" & vbCrLf & _
        "<p><span class=""ute_headline"">" & m_sTable & "</span></p>" & vbCrLf & _
        "<p><span class=""ute_subheadline"">" & m_sHeadline & "</span></p>" & vbCrLf

    for i = 1 to UBound(m_ErrorField)
        if m_ErrorField(i) = "" then
            sValue = sValue & _
                "<p><span class=""ute_form_error"">" & _
                m_ErrorMessage(i) & _
                "</span></p>" & vbCrLf
        end if
    next


	

	sValue = sValue & _
	    "<form method=""post"" action=""" & s & """>" & vbCrLf & _
	    InsertIdentFields & vbCrLf & _
	    "<table class=""ute_form"">" & vbCrLf


    ' add toolbar
    sValue = sValue & getFormToolBar(nColCount) & vbCrLf

    ' insert all primary keys
    for i = 1 to UBound(m_PrimaryKeyFields)
        sValue = sValue & InsertFieldForm (m_RSForm.fields(m_PrimaryKeyFields(i)), True)
    next

    ' insert all other fields
    for i = 1 to UBound(m_StandardFields)
        sValue = sValue & InsertFieldForm (m_RSForm.fields(m_StandardFields(i)), False)
    next
    

	
    sValue = sValue & _
        "<tr>" & vbCrLf & _
        "<td colspan=""" & CStr(nColCount) & """>" & vbCrLf & _
        "<IMG SRC=""images/icon_required.gif"" ALIGN=Middle HSPACE=2>" & Cstr(STR_REQUIREDFIELD) &_
        "<br><br>&nbsp;" & vbCrLf 
        
		'If a rich editor is enabled then change the way the form is submited
		flag=false
		if UBound(m_FieldsRichEdit) > 0 and m_nFormMode<>MD_DELETE then 
				flag=true			
		end if
        
        
        
        if flag=true then
        
			sValue = sValue & "<input type=""button"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_ok"" " & _
			        "value=""" & STR_OK & """  onclick=""javascript:submitForm();document.forms[0].action='" & s & "';document.forms[0].submit(); "" >" & vbCrLf        
        
        else
			sValue = sValue & "<input type=""submit"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_ok"" " & _
			        "value=""" & STR_OK & """>" & vbCrLf 
		end if
                
                
        sValue=sValue & "&nbsp;" & vbCrLf & _
        "<input type=""submit"" name=""" & sFormUTEFieldPrefix & sFormButton & """ class=""ute_btn_cancel"" " & _
                "value=""" & STR_CANCEL & """>" & vbCrLf & _
        "&nbsp;" & vbCrLf & _        
        "</td>" & vbCrLf & _
        "</tr>" & vbCrLf

    sValue = sValue & _
        "</table>" & vbCrLf & _
        "</from>"
	
    buildHTML_Form = sValue

End Function


%>
