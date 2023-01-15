' ----------------------------------------------------------------------
'   Script for renaming several file names at once
'   (c) 2019 Markus J. Auvo
'
'   The script works in similar fashion as find-and-replace (CTRL+H)
'   function in e.g. Microsoft Windows Office applications.
'
'	GUIDE:
'
'   The user is prompted for following:
'   - absolute path of the folder containing the files
'   - string to be found and replaced
'   - new string
' ----------------------------------------------------------------------

OPTION EXPLICIT

' ----------------------------------------------------------------------
'  Constant and variable declarations
' ----------------------------------------------------------------------

Const title = "FIND & REPLACE"

Dim FSO:    Set FSO = CreateObject("Scripting.FileSystemObject")

Dim strFolderPath   ' Absolute path to folder
Dim objFolderTarget ' Target folder
Dim objFileList     ' Files in the folder

Dim stringFrom      ' String to be replaced
Dim stringTo        ' New string

' ----------------------------------------------------------------------
'  Methods
' ----------------------------------------------------------------------

'
' Displays messages in console window without a line break
'
Sub Whisper(msg)
    WScript.StdOut.Write msg
End Sub

'
' Displays messages in console window with a line break
'
Sub Say(msg)
    WScript.Echo msg 
End Sub

'
' Displays messages in message box
'
Sub Shout (msg)
    MsgBox msg, 0, "Message"
End Sub

'
' Gets a folder at a given path -- if folder is found
'
Function getFolder(absPath)
	Whisper "Searching for folder..."
    If (FSO.FolderExists(absPath)) Then
        Set objFolderTarget = FSO.GetFolder(absPath)
        Set objFileList = objFolderTarget.Files
        Say "FOUND!!"
        Say "-- Folder: " & VBTab & objFolderTarget.Name
        Say VBTab & VBTab & objFolderTarget.Files.Count & " files" & VBCrLf
    Else
        Say "NOT FOUND"
        Shout "Completed!"
        Wscript.Quit
    End If
End Function

'
' Iterates through every file in the older
'
Function iterateThrough(sFind, sReplace)
    Dim fileItem
    Dim tempFile

    For Each fileItem In objFileList
        tempFile = fileItem.Name
        tempFile = Replace(tempFile, sFind, sReplace)
        If(tempFile<>fileItem.Name) Then
            Whisper "..." & fileItem.Name & VBTab & "-->" & VBTab
            fileItem.Move(fileItem.ParentFolder & "\" & tempFile)
            Say tempFile
        End If
    Next
End Function

' ----------------------------------------------------------------------
'  OKAY...LET'S ROCK AND ROLL !!
' ----------------------------------------------------------------------

'
' User input: the absolute path to a file folder
'
Do While strFolderPath = Empty
    strFolderPath = InputBox("Enter the absolute path to a file folder",title)
Loop

getFolder(strFolderPath)

'
' User input: Find a (sub)string
'
Do While stringFrom = Empty
    stringFrom = InputBox("Find what",title)
Loop

'
' User input: Replace with a (sub)string
' --may also be an empty (sub)string
'
stringTo = InputBox("Replace with",title)
If (stringTo = Empty) Then
    stringTo = ""
End If

'
' Do the Magic!!
'
iterateThrough stringFrom, stringTo

'
' Voila!
'
Shout("Completed!")

