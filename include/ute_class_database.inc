<%
'---------------------------------------------------------------------------
'
'   Project:    UTE - (U)niversal ASP (T)able (E)ditor
'
'   Module:     UTE class - View Database Functions
'
'   Version:    3.00
'
'   Comments:   This module does the following things:
'                   1. defines all functions being needed in
'                      database mode
'
'---------------------------------------------------------------------------
'
'   (c) in 2000-2003 by Tom Wellige                    
'   http://www.wellige.com  mailto:tom@wellige.com     
'                                               
'   This project is released under the "GNU General Public License (GPL)" 
'   http://www.gnu.org/licenses/gpl.html
'
'   and is maintained on SourceForge at
'   http://sourceforge.net/projects/ute-asp/
'
'   and can also be found on CodeProject at
'   http://www.codeproject.com/asp/ute.asp
'
'---------------------------------------------------------------------------


''--------------------------------------------------------------------------
'' Name:     buildHTML_Database
''           ==================
'' 
'' Creates entire UTE HTML code for database mode.
''
'' Parameter: 
''      none
''
'' return value:
''      string      HTML code
''
''--------------------------------------------------------------------------
Private Function buildHTML_Database()
    Dim sValue, s
    Dim rstSchema

    if m_sDBName <> "" then
        s = Replace(STR_DB_TITLE, "%1", m_sDBName)
    else
        s = Replace(STR_DB_TITLE, "%1", STR_DATABASE)
    end if

    sValue = _
        "<p><span class=""ute_headline"">" & s & "</span></p>" & vbCrLf & _
        "<table width=""100%""><tr>" & vbCrLf & _
        "<td class=""ute_navigation"">" & vbCrLf & _
        "<ul class=""ute_table_list"">" & vbCrLf

    Set rstSchema = m_DB.OpenSchema(adSchemaTables)
    Do Until rstSchema.EOF
        if UCase(rstSchema("TABLE_TYPE")) = "TABLE" then
            s = Request.QueryString
            s = getLink(m_sUTEScript, s, sParamTable, rstSchema("TABLE_NAME"))
            s = getLink(m_sUTEScript, s, sParamMode,  MD_TABLE)
            sValue = sValue & "<li><a href=""" & s & """>" & rstSchema("TABLE_NAME") & "</a>" & vbCrLf
        end if
        rstSchema.MoveNext
    Loop
    rstSchema.Close
    Set rstSchema = Nothing

    sValue = sValue & _
        "</ul>" & vbCrLf & _
        "</td></tr></table>" & vbCrLf & _
        getPoweredBy

    buildHTML_Database = sValue

End Function

%>