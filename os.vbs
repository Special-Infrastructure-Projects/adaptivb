dim fso
set fso = CreateObject("Scripting.FileSystemObject")
function listdir(directory)
	Set folder = fso.GetFolder(directory)
	Set files = folder.Files
	redim arr(-1)
	for each file in files
		redim preserve arr(ubound(arr)+1)
		arr(ubound(arr)) = file
	next
	listdir = (arr)
end function
function mkdir(name)
	fso.CreateFolder name
end function
function exists(name)
	e = fso.FolderExists(name)
	if e = 0 then
		exists = False
	else
		exists = True
	end if
end function
function getlogin()
	Set wshShell = CreateObject( "WScript.Shell" )
	getlogin = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )	
end function
