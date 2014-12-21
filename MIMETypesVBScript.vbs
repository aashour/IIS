Option Explicit

'---------- CONFIGURATION ----------
Dim OutputFileName, IISAddress, Properties
OutputFileName = "IISMimeTypes.xml"
IISAddress = "IIS://localhost/MimeMap"
Properties = "MimeMap"
'---------- END CONFIGURATION ----------

Dim Version, Author, VersionDate, Year
Version = "1.1"
Author = "cAfonso"
VersionDate = "09/12/2008"
Year = "2008"

' Echo to user
Wscript.echo " ------------------------------"
Wscript.echo " IIS MIMETypes Exporter " & Version & " (" & VersionDate & ")"
Wscript.echo " Copyright (C) " & Year & " " & Author
Wscript.echo " "
Wscript.echo " Begin..."

' Objects that will need to write to the file
Dim objFSO, objShell, objTextFile, objFile

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

Wscript.echo " Created FileSystemObject..."

' Does the file exist? If so, yield a notice
If objFSO.FileExists(OutputFileName) Then
   Wscript.Echo " NOTE: File " & OutputFileName & " already exists and will be overwritten!"
Else
   Set objFile = objFSO.CreateTextFile(OutputFileName)
   Wscript.echo " Created file " & OutputFileName & "..."
End If 

set objFile = nothing

' OpenTextFile Method needs a Const value; ForAppending = 8, ForReading = 1, ForWriting = 2
Const OpenMethod = 2

' Open the file
Set objTextFile = objFSO.OpenTextFile(OutputFileName, OpenMethod, True)

Wscript.echo " Openned file for writing..."
' This script lists the MIME types for an IIS Server.
' To use this script, just double-click or execute it from a command line (cscript Filename.vbs)

Dim mimeMapEntry, allMimeMaps, mimeMap

' Get the mimemap object.
Set mimeMapEntry = GetObject(IISAddress)
allMimeMaps = mimeMapEntry.GetEx(Properties)

Wscript.echo " Got IIS MIMETypes..."

' Sort the elements
Dim objSortedList, Seed, Disambig
Set objSortedList = CreateObject( "System.Collections.Sortedlist" )
' litle "trick": IIS can have more than one extension with the same description, sorted list doesn't allow multiple values mapped to the same key. Let's add spaces to the key and them trim the output.
Disambig = " "

For Each mimeMap In allMimeMaps
	If objSortedList.Contains(mimeMap.MimeType) = -1 then
		Disambig = Disambig & " "
		objSortedList.Add mimeMap.MimeType & Disambig, mimeMap.Extension
	Else
		objSortedList.Add mimeMap.MimeType, mimeMap.Extension
	End If
Next

objTextFile.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
objTextFile.WriteLine("<?xml-stylesheet type=""text/xsl"" href=""MIMETypes.xsl""?>")
objTextFile.WriteLine("<IISRegisteredMimeTypes>")

For Seed = 0 To objSortedList.Count - 1
	objTextFile.WriteLine(" 	<RegisteredType>")
	objTextFile.Write(" 		<MimeType>")
	objTextFile.Write(Trim(objSortedList.GetKey(Seed)))
	objTextFile.WriteLine("</MimeType>")
	objTextFile.Write(" 		<Extension>")
	objTextFile.Write(objSortedList.GetByIndex(Seed))
	objTextFile.WriteLine("</Extension>")
	objTextFile.WriteLine(" 	</RegisteredType>")
Next

objTextFile.WriteLine("</IISRegisteredMimeTypes>")

Wscript.echo " MIMETypes successfully exported to file!"

objTextFile.Close

Wscript.echo " All done!"
Wscript.echo " "
Wscript.echo " ------------------------------"

WScript.Quit



