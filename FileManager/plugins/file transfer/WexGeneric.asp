<%
	' WebExplorer Generic File Transfer Plugin
	' This plugin requires Microsoft Data Access Components (MDAC) 2.5+ 
	' to be installed on the server. See http://www.microsoft.com/data

	' Standard Class for File Transfer Plugin
	Class pluginFileTransfer
		Public path
		Public uploadedFileName, uploadedFileSize
		Public contentType
		
		' Create objects required by the plugin
		Private Sub Class_Initialize()
			'FSO is already available so the following line is commented out
			'Set FSO = server.CreateObject ("Scripting.FileSystemObject")
			Set stream = Server.CreateObject("ADODB.Stream")
		End Sub
		
		' Destroy objects
		Private Sub Class_Terminate()
			'Set FSO = Nothing
			Set stream = Nothing
		End Sub
	
		' Upload the posted file
		' Return values: 0 - success, 1 - no file sent, 2 - path not found, 3 - write error, 4 - extension not allowed
		Public Function Upload()
			Dim file

			stream.Type = 1
			stream.Open
			
			ParsePost()
			
			If uploadedFileSize<=0 Then
				Upload = 1
			ElseIf not FSO.FolderExists(path) Then 
				Upload = 2
			Else
				If CheckExtension(uploadedFileName) Then
					on error resume next
					
					stream.SaveToFile FSO.BuildPath(path, uploadedFileName), 2
				
					If err.Number<>0 Then
						Upload = 3
					Else
						Upload = 0
					End If
				Else
					Upload = 4
				End If
			End If
			
			stream.Close
		End Function
		
		' Download the file with the given name at current path
		' Return values: 0 - success, 1 - file not found, 2 - read error
		Public Function Download(fileName)
			Const chunk = 65535
			
			Dim filePath, fileSize, i
			
			filePath = FSO.BuildPath(path, fileName)
		
			If not FSO.FileExists(filePath) Then 
				Download = 1
			Else
				stream.Type = 1
				stream.Open 

				on error resume next					
				stream.LoadFromFile(filePath)
				
				fileSize = stream.Size
				
				If err.Number<>0 Then
					Download = 2
				Else
					Response.AddHeader "Content-Disposition","attachment;filename="  & fileName
					Response.AddHeader "Content-Length", fileSize
					Response.CharSet = "UTF-8" 
					Response.ContentType = "application/x-msdownload"

					For i = 1 To fileSize \ chunk
						If Not Response.IsClientConnected Then Exit For
						Response.BinaryWrite stream.Read(chunk)
					Next

					If fileSize Mod chunk > 0 Then 
						If Response.IsClientConnected Then 
							Response.BinaryWrite stream.Read(fileSize Mod chunk) 
						End If 
					End If 

					Download = 0
				End If
				
				stream.Close 
			End If
		End Function

' --- Internal variables and functions specific to this plugin only (non standard) ---
		
		Private stream

		' Parses the posted data to extract file data and info
		Private Sub ParsePost()
			Dim biData, sInputName
			Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
			Dim nPosFile, nPosBound
			
			Dim tmpStream
	
			biData = Request.BinaryRead(Request.TotalBytes)

			nPosBegin = 1
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
			
			If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
			 
			vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
			nDataBoundPos = InstrB(1, biData, vDataBounds)
			
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			
			If nPosFile <> 0 And nPosFile < nPosBound Then
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				uploadedFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				uploadedFileName = Right(uploadedFileName, Len(uploadedFileName)-InStrRev(uploadedFileName, "\"))
	
				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
					
				contentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				
				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
		
				Set tmpStream = Server.CreateObject("ADODB.Stream")
				tmpStream.Type = 1
				tmpStream.Open
				tmpStream.Write biData
				tmpStream.Position = nPosBegin-1
				tmpStream.CopyTo stream, nPosEnd-nPosBegin
				Set tmpStream = Nothing
				
				uploadedFileSize = stream.Size
			End If
		End Sub

		' String to byte string conversion
		Private Function CByteString(sString)
			Dim nIndex
			For nIndex = 1 to Len(sString)
				CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
			Next
		End Function
	
		' Byte string to string conversion
		Private Function CWideString(bsString)
			Dim nIndex
			CWideString =""
			For nIndex = 1 to LenB(bsString)
				CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
			Next	
		End Function
	End Class
%>