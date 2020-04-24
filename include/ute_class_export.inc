<%



''--------------------------------------------------------------------------
'' Name:     ExportToStream
''           ==============
'' 
'' Writes all data in current recordset as comma separated text to the response
'' stream.
''
'' Parameter: 
''      none
''
'' return value:
''      none
''
''--------------------------------------------------------------------------
Private Sub ExportToStream ()

    Dim i
    Dim sLine,sLabel

    ' Clear out the existing HTTP header information
    Response.Buffer = TRUE
    Response.Clear
    Response.ContentType = "text/csv"
    Response.AddHeader "Content-Disposition", "inline;filename=" & m_sTable & ".csv"


    ' set RecordSet to get all data
    m_RS.PageSize     = m_RS.RecordCount
    m_RS.AbsolutePage = 1
    m_RS.MoveFirst

    ' write Header
    sLine = ""
    for i = 1 to UBound(m_PrimaryKeyFields)
        sLabel=FieldHasALabel(m_PrimaryKeyFields(i))
   		if sLabel<>"" then
			sLine = sLine & DEF_EXPORT_VAL & sLabel & DEF_EXPORT_VAL & DEF_EXPORT_SEP 
		else       
			sLine = sLine & DEF_EXPORT_VAL & m_PrimaryKeyFields(i) & DEF_EXPORT_VAL & DEF_EXPORT_SEP
		end if
    next

    for i = 1 to UBound(m_StandardFields)
        sLabel=FieldHasALabel(m_StandardFields(i))
		if sLabel<>"" then
			sLine = sLine & DEF_EXPORT_VAL & sLabel & DEF_EXPORT_VAL & DEF_EXPORT_SEP        
		else
			sLine = sLine & DEF_EXPORT_VAL & m_StandardFields(i) & DEF_EXPORT_VAL & DEF_EXPORT_SEP
		end if
        
    next

    if sLine <> "" then
        sLine = Left(sLine, Len(sLine)-1)
        Response.Write sLine & vbCrLf
    end if 

    Do Until m_RS.EOF
        sLine = ""
        for i = 1 to UBound(m_PrimaryKeyFields)
            sLine = sLine & DEF_EXPORT_VAL & m_RS(m_PrimaryKeyFields(i)) & DEF_EXPORT_VAL & DEF_EXPORT_SEP
        next
        for i = 1 to UBound(m_StandardFields)
            sLine = sLine & DEF_EXPORT_VAL & m_RS(m_StandardFields(i)) & DEF_EXPORT_VAL & DEF_EXPORT_SEP
        next
        if sLine <> "" then
            sLine = Left(sLine, Len(sLine)-1)
            Response.Write sLine & vbCrLf
        end if 
        m_RS.MoveNext
    Loop

    Response.End

End Sub

%>