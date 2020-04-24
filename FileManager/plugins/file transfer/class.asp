<%
	' File Transfer Plugin Name
	' Required component
	' URL to the component

	' Standard Class for File Transfer Plugin
	Class pluginFileTransfer
		Public path
		Public uploadedFileName, uploadedFileSize
		Public contentType
		
		' Create objects required by the plugin
		' The 3rd party component etc.
		Private Sub Class_Initialize()

		End Sub
		
		' Destroy objects
		Private Sub Class_Terminate()

		End Sub
	
		' Upload the posted file
		' Return values: 0 - success, 1 - no file sent, 2 - path not found, 3 - write error, 4 - extension not allowed
		' The public variable path (save location) should be set prior to calling this function.
		' The public variables uploadedFileName, uploadedFileSize and contentType should be set before exiting this function.
		Public Function Upload()

		End Function
		
		' Download the file with the given name at current path
		' Return values: 0 - success, 1 - file not found, 2 - read error
		Public Function Download(fileName)

		End Function

' --- Internal variables and functions specific to this plugin only (non standard) ---
		' Here, the private variables, functions  specific to this plugin can be defined.

	End Class
%>