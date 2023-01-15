' ----------------------------------------------------------------------
'   A code template for vbscript files.
'   (c) 2019-2023 Markus J. Auvo
'
'   GUIDE:
'
'   Simply comment out unnecessary parts.
' ----------------------------------------------------------------------

OPTION EXPLICIT

' ----------------------------------------------------------------------
'  Constants and Variables
' ----------------------------------------------------------------------

' STRINGS

' Strings for writing to a file
Dim targetFileDir:      targetFileDir = "<absolute folder path>"
Dim targetFile:         targetFile = "<filename>"

' OBJECTS

Dim objFSO:             Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objFileToWrite:     Set objFileToWrite = objFSO.OpenTextFile(targetFileDir + targetFile, 2, true)

Dim shellAPP:           Set shellAPP = CreateObject("Shell.Application")
Dim folderNamespace:    Set folderNamespace = shellAPP.Namespace("<absolute folder path>")

' ----------------------------------------------------------------------
'  Methods
' ----------------------------------------------------------------------

'
' Displays a message in console window without a line break
'
Sub WriteToConsole(msg)
    WScript.StdOut.Write msg
End Sub

'
' Displays a message in console window with a line break
'
Sub WriteToConsoleNewLine(msg)
    WScript.Echo msg 
End Sub

'
' Displays a message in console window without a line break
' and moves the cursor back to the beginning of the line.
'
Sub WriteToConsoleLineReturn(msg)
    WScript.StdOut.Write msg & chr(13)
End Sub

'
' Displays a message in a message box
'
Sub WriteToDialog(msg)
    MsgBox msg, 0, "Message"
End Sub

'
' Writes a line of strings to a file
'
Sub WriteToFile(dataline)
    objFileToWrite.WriteLine(dataline)
End Sub

