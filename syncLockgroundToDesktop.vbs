'''''''''''''''''''''''''''''''''''''''''''''
' Scripter: Eric Olsen						'
' Version: 0.14								'
' Script Purpose:							'
'		Take any new pictures added to the	'
'	lock screen and add them to the desktop	'
'	backgrounds as well.					'
'											'
' To Do:									'
'	Simplify configuration, maybe move to	'
'	an external file						'
'''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''
' Settings are further down, DO NOT change	'
'	anything between here and the next note	'
'''''''''''''''''''''''''''''''''''''''''''''

' Global Objects -- don't change these!
Set oFS = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
' -- or this
pathSource =  oShell.ExpandEnvironmentStrings("%LocalAppData%") & "\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"

' Configuration -- you can change these ;)
'--------------------------------------------
' pathDestination is where the pictures should be saved to. If you don't want it in your User Profile under Pictures\Wallpapers
'	put a ' in front of the next line, and remove the ' in front of the one after, then change the path between ""
pathDestination = oShell.ExpandEnvironmentStrings("%UserProfile%") & "\Pictures\Wallpapers"
'Const pathDestination = "C:\Pictures\Wallpapers"
'--------------------------------------------
' The rest of the settings are mostly for fine-tuning, enabling logging, and what-not
Const MIN_PIC_SIZE = 400000 ' in bytes - smallest file to consider possibly a picture
Const MIN_PIC_WIDTH = 1680 ' Minimum width of the picture to be considered for a background
Const MIN_PIC_HEIGHT = 1050 ' Minimum height of the picture to be considered for a background
Const LOG_TO_FILE = False ' True or False, write details of what's happening to a file
Const LOG_FILE = "syncLockgroundToDesktop.log"
Const MAX_LOG_SIZE = 40000 ' in bytes - how big to allow each log to grow to (each, log and old log if enabled)
Const LOG_BACKUP = True ' True or False, copy old log to LOG_FILE_OLD
Const LOG_FILE_OLD = "syncLockgroundToDesktop_old.log"
Const LOG_TO_SCREEN = False ' True or False, pop-up a window for each line (I believe it pauses the script at that line too)
'--------------------------------------------
' Don't change anything below this!!!
'--------------------------------------------
' Global Constants
Const FOR_APPENDING = 8 '1 = read only, 2 = overwrite, 8 = append

Call Main()

Sub Main()
	if LOG_TO_FILE AND LOG_BACKUP then
		' Check the log size, clear it out if too large
		If oFS.FileExists(LOG_FILE) Then
			If oFS.GetFile(LOG_FILE).Size > MAX_LOG_SIZE Then
				oFS.CopyFile LOG_FILE, LOG_FILE_OLD, True
				oFS.DeleteFile LOG_FILE
			End If 'oFS.GetFile(LOG_FILE).Size > MAX_LOG_SIZE
		End If 'oFS.FileExists(LOG_FILE)
	End if 'LOG_BACKUP
	
	log "---=== STARTED AT " & Now & " ===---"
	'''' EXISTING GUIDs
	' Load the folder
	If NOT oFS.FolderExists(pathDestination) Then oFS.CreateFolder(pathDestination)
	Set oFolder = oFS.GetFolder(pathDestination)
	' Count the files and set the array size
	Redim existingGuidList(oFolder.Files.Count - 1)
	' Load destination (existing) file GUIDs into an array (to see what I already have done)
	n = 0
	For Each file In oFolder.Files
		existingGuidList(n) = getGuid(file)
		n = n + 1
	Next
	
	'''' NEW GUIDs
	' Load the folder
	Set oFolder = oFS.GetFolder(pathSource)
	' Check source file (potentially new) GUIDs against existing file GUIDs array (what might need to be copied)
	For Each file In oFolder.Files
		log "Found " & getGuid(file)
		' Skip any small files, they're likely not pictures
		If file.size > MIN_PIC_SIZE Then
			log "	It's larger than " & MIN_PIC_SIZE & " bytes, so it's probably a picture"
			' It's probably a picture, load it to check the dimensions
			Set pic = LoadPicture(file)
			log "	Loaded as a picture"
			' Skip anything that's either obviously not a background, or is designed for a phone (portrait)
			log "	Dimensions are: " & Round(pic.Width/26.4583) & "x" & Round(pic.Height/26.4583)
			If Round(pic.Width/26.4583) >= MIN_PIC_WIDTH AND Round(pic.Height/26.4583) >= MIN_PIC_HEIGHT Then
				log "		Dimensions are >= " & MIN_PIC_WIDTH & "x" & MIN_PIC_HEIGHT & ", we should be saving this one as a background! (if it's not already there)"
				' Dimensions are good, check if we already have this one
				newGuid = getGuid(file)
				newPic = True
				For Each guid In existingGuidList
					' If it's in the list of existing pics, don't copy it again
					If newGuid = guid Then
						log "	Apparently it's already saved as a background, don't copy a duplicate..."
						newPic = False
					End If 'newGuid = guid
				Next
				' If it wasn't found, then copy it over
				If newPic = True Then
					copy file
				End If 'newPic = True
			Else
				log "	Dimensions are < " & MIN_PIC_WIDTH & "x" & MIN_PIC_HEIGHT & ", skipping"
			End If 'pic.Width >= MIN_PIC_WIDTH AND pic.Height >= MIN_PIC_HEIGHT
		Else
			log "	It's smaller than " & MIN_PIC_SIZE & " bytes, skipping"
		End If 'file.size > MIN_PIC_SIZE
	Next 'For Each file In oFolder.Files
	log "---=== ENDED AT " & Now & " ===---"
End Sub 'Main()

Function getGuid(ByVal file)
	Dim result
	' Strip down to just the file name
	'file = RIGHT(file, (LEN(file) - (LEN(path) + 1)))
	file = Split(file,"\")
	' Strip off the extension, if any
	result = Split(file(uBound(file)),".")(0)
	' Return the GUID
	getGuid = result
End Function 'getGuid(file)

Function copy(ByVal file)
	Dim finalDestination
	finalDestination = pathDestination & "\" & getGuid(file) & ".jpg"
	log "	It's a NEW PICTURE!!! Copying " & file & " to " & finalDestination
	' Copy the picture from the Lock Screen backgrounds to the Desktop Wallpapers folder
	oFS.CopyFile file, finalDestination
	'msgbox "Copy " & file & " to " & destination & "\" & getGuid(file) & ".jpg"
End Function 'move(guid)

Function log(ByVal line)
	if LOG_TO_SCREEN Then WScript.echo line
	if LOG_TO_FILE Then
		Set file = oFS.OpenTextFile(LOG_FILE, FOR_APPENDING, True)
		file.WriteLine(line)
	End if ' LOG_TO_FILE
End Function 'log(line)