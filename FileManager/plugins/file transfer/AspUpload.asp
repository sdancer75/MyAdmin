<%
	' Persits Software's AspUpload File Transfer Plugin
	' This plugin requires Persits Software's AspUpload 
	' to be installed on the server. See http://www.aspupload.com

	' Standard Class for File Transfer Plugin
	Class pluginFileTransfer
		Public path
		Public uploadedFileName, uploadedFileSize
		Public contentType
		
		' Create objects required by the plugin
		' The 3rd party component etc.
		Private Sub Class_Initialize()
			'FSO is already available so the following line is commented out
			'Set FSO = server.CreateObject ("Scripting.FileSystemObject")
			Set aspUpload = Server.CreateObject("Persits.Upload.1")
		End Sub
		
		' Destroy objects
		Private Sub Class_Terminate()
			'Set FSO = Nothing
			Set aspUpload = Nothing
		End Sub
	
		' Upload the posted file
		' Return values: 0 - success, 1 - no file sent, 2 - path not found, 3 - write error, 4 - extension not allowed
		' The public variable path (save location) should be set prior to calling this function.
		' The public variables uploadedFileName, uploadedFileSize and contentType should be set before exiting this function.
		Public Function Upload()
			Dim file
			
			on error resume next
			aspUpload.SaveToMemory
			Set file = aspUpload.Files(1)
						
			If file is nothing Then
				Upload = 1
			ElseIf not FSO.FolderExists(path) Then 
				Upload = 2
			Else
				uploadedFileName = file.ExtractFileName
				
				If CheckExtension(uploadedFileName) Then
					file.SaveAs FSO.BuildPath(path, uploadedFileName)
					
					If err.Number<>0 Then
						Upload = 3
					Else
						uploadedFileSize = file.Size
						contentType = file.ContentType
						Upload = 0
					End If				
				Else
					Upload = 4
				End If
			End If
		End Function
		
		' Download the file with the given name at current path
		' Return values: 0 - success, 1 - file not found, 2 - read error
		Public Function Download(fileName)
			Dim filePath
			filePath = FSO.BuildPath(path, fileName)
		
			If not FSO.FileExists(filePath) Then 
				Download = 1
			Else
				on error resume next					
			
				aspUpload.SendBinary filePath, true, "application/octet-stream", true

				If err.Number<>0 Then
					Download = 2
				Else
					Download = 0
				End If
			End If
		End Function

' --- Internal variables and functions specific to this plugin only (non standard) ---
		' Here, the private variables, functions  specific to this plugin can be defined.
		Private aspUpload
	End Class
%>