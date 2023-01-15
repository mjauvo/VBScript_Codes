' ----------------------------------------------------------------------
'   Script for printing the contents of a movie directory to a text file.
'   (c) 2019-2023 Markus J. Auvo
'
'   VERSION 3
'
'   The script prints out the movie contents from several root directories.
'   (e.g. M:\, N:\) including its subdirectories into an output file (.CSV).
'   Output file is then used as imported data source in an Excel file.
'
'   In this case, the movies have been are organized into category folders
'   with names beginning with an underscore _. This script iterates through
'   following video file types: avi, mkv, mp4
'
'   The output file will be placed into a given location.
'   - "outputFile"  :   Name of the output file
'   - "outputDir"   :   Location of the output file
' ----------------------------------------------------------------------

OPTION EXPLICIT

' ----------------------------------------------------------------------
'  Constants and Variables
' ----------------------------------------------------------------------

' Dictionary object for several root directories,
' in case several locations are used
Dim rootKey
Dim rootDirs
Set rootDirs = CreateObject("Scripting.Dictionary")

rootDirs.Add "U:\", "CINEMA I"
rootDirs.Add "V:\", "CINEMA II"

' Runtime variables

Const outputDir =   <directory where the output file is to be generated>
Const outputFile =  <name of the output file>

Const COL_TITLE =       "Title"
Const COL_YEAR =        "Year"
Const COL_GENRE =       "Genre"
Const COL_LENGTH =      "Length"
Const COL_SIZE =        "Size"
Const COL_FILE_TYPE =   "File Type"
Const COL_WIDTH =       "Width"
Const COL_HEIGHT =      "Height"
Const COL_SUBTITLE =    "Subtitle"
Const COL_FOLDER =      "Folder"
Const COL_FILENAME =    "File"

Const MOVIE_FILE =                0
Const MOVIE_SIZE =                1
Const MOVIE_FILE_TYPE =           2
Const MOVIE_YEAR =               15
Const MOVIE_GENRE =              16
Const MOVIE_TAGS =               18
Const MOVIE_TITLE =              21
Const MOVIE_LENGTH =             27
Const MOVIE_FILE_NAME =         165 ' File name including the file extension
Const MOVIE_FOLDER_CURRENT =    190 ' Current folder of the file
Const MOVIE_FOLDER_STRUCTURE =  191 ' Folder structure of the file w/o the file
Const MOVIE_FULL_PATH =         194 ' Full path of the file
Const MOVIE_SUBTITLE =          215
Const MOVIE_FRAME_HEIGHT =      314 ' Frame height
Const MOVIE_FRAME_WIDTH =       316 ' Frame width

Dim folderNamespace
Dim inputDir
Dim strRootDirectory                ' Absolute path to root directory
Dim intFileCount: intFileCount = 0  ' Reset the counter for movie file items

Dim shellAPP:           Set shellAPP = CreateObject("Shell.Application")

Dim objFSO:             Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objOutputFile:      Set objOutputFile = objFSO.CreateTextFile(outputDir & outputFile, 2, True)

Dim objROOT
Dim objCategoryFolderList
Dim intCategoryFolderCount: intCategoryFolderCount = 0  ' Reset the category folder counter

' ----------------------------------------------------------------------
'  Methods
' ----------------------------------------------------------------------

'
' Display messages in console window without a line break
'
Sub WriteToConsole(msg)
    WScript.StdOut.Write msg
End Sub

'
' Display messages without a line break in console window
' and moves the cursor back to the beginning of the line.
'
Sub WriteToConsoleR(msg)
    WScript.StdOut.Write msg & chr(13)
End Sub

'
' Display messages with a line break in console window
'
Sub WriteToConsoleNL(msg)
    WScript.Echo msg 
End Sub

'
' Display a message in a dialog box
'
Sub WriteToDialog(msg)
    MsgBox msg, 0, "Message"
End Sub

'
' Write a line of strings into a file
'
Sub WriteToFile(text)
    objOutputFile.WriteLine(text)
End Sub

' ----------------------------------------------------------------------
'  Functions
' ----------------------------------------------------------------------

' Get the current time
'
' This solution was copied and modified from Stackoverflow post reply
' Jan 30 '12 at 19:15 by user MBu
' edited by user Jeremy Odekirk at Feb 11 '19 at 21:42
' https://stackoverflow.com/questions/9063340/find-time-with-millisecond-using-vbscript
' 
Function getCurrentTime()
    Dim t: t = Timer

    ' Int() behaves exactly like Floor() function, i.e. it returns the biggest integer lower than function's argument
    Dim temp: temp = Int(t)

    Dim Milliseconds: Milliseconds = Int((t-temp) * 1000)
    Dim Seconds: Seconds = temp mod 60

    temp    = Int(temp/60)

    Dim Minutes: Minutes = temp mod 60
    Dim Hours: Hours   = Int(temp/60)

    ' Let's format it
    Dim strTime
    strTime =           String(2 - Len(Hours), "0") & Hours & ":"
    strTime = strTime & String(2 - Len(Minutes), "0") & Minutes & ":"
    strTime = strTime & String(2 - Len(Seconds), "0") & Seconds & "."
    strTime = strTime & String(4 - Len(Milliseconds), "0") & Milliseconds

    getCurrentTime = strTime
End Function

Function initializeOutputFile
        ' Write column headers into output file
        WriteToConsole "Column headers for output file..."
        WriteToFile(COL_TITLE & ";" & COL_YEAR & ";" & COL_GENRE & ";" & COL_LENGTH & ";" & COL_SIZE & ";" & COL_FILE_TYPE & ";" & COL_WIDTH & ";" & COL_HEIGHT & ";" & COL_SUBTITLE & ";" & COL_FOLDER & ";" & COL_FILENAME)
        WriteToConsoleNL "[OK]"
        WriteToConsoleNL ""
End Function

'
' Get the root directory at a given path -- if it is found
'
Function getRootDirectory(rootPath)
    WriteToConsole "Target root '" & rootDirs(rootPath) & "'..."
    If (objFSO.FolderExists(rootPath)) Then
        Set objROOT = objFSO.getFolder(rootPath)
        Set objCategoryFolderList = objROOT.Subfolders
        WriteToConsoleNL "[OK]"
    Else
        WriteToDialog "[FAIL]"
        Wscript.Quit
    End If
End Function

'
' Iterate through category folders in a given root directory
'
Function iterateThroughCategoryFolders
    Dim objCategoryFolder

    For Each objCategoryFolder In objCategoryFolderList
        ' A category folder begins with an underscore
        If(Left(objCategoryFolder.Name, 1) = "_") Then
            ' Increment category folder counter
            intCategoryFolderCount = intCategoryFolderCount + 1
            ' Go through files in the category folder, if there are any.
            If(objCategoryFolder.Files.count > 0) Then
                iterateThroughFiles(objCategoryFolder)
            End If
            ' Go through subfolders in the category folder, if there are any.
            If(objCategoryFolder.Subfolders.count > 0) Then
                iterateThroughSubfolders(objCategoryFolder)
            End If
        End If
    Next

    ' No category folders were found
    If(intCategoryFolderCount < 1) Then
        WriteToDialog "NO MOVIE CATEGORIES FOUND!"
    End If
End Function

'
' Iterate through subfolders in a category folder
'
Function iterateThroughSubfolders(parentFolder)
    Dim objSubFolder
    Dim subfolder

    For Each objSubFolder In parentFolder.Subfolders
        ' Go through files in the category subfolder, if there are any.
        If(objSubFolder.Files.count > 0) Then
            iterateThroughFiles(objSubFolder)
        End If
        ' Go through subfolders in the this folder, if there are any.
        If(objSubFolder.Subfolders.count > 0) Then
            iterateThroughSubfolders(objSubFolder)
        End If
    Next
End Function

'
' Iterate through movie files in a subfolder and 
' write them to the output file.
'
' Movie files with extension AVI, MKV and MP4 are processed.
'
Function iterateThroughFiles(subfolder)
    Dim objFile
    Dim fileName
    Dim strFileExt
    Set folderNamespace = shellAPP.Namespace(subfolder.Path)

    ' Get the movie file properties
    '
    ' This solution was copied and modified from Stackoverflow post reply
    ' Jul 31 '14 at 8:04 by user MC ND.
    ' https://stackoverflow.com/questions/25050807/how-can-i-use-vbscript-to-read-the-attributes-of-an-mp4-file
    ' 
    Dim headers, i, aHeaders(330)
        For i = 0 to 329
            aHeaders(i) = folderNamespace.GetDetailsOf(folderNamespace.Items, i)
        Next

    For Each fileName in folderNamespace.Items
        If (LCase(Right(fileName,4))=".avi" OR LCase(Right(fileName,4))=".mkv" OR LCase(Right(fileName,4))=".mp4") Then 

            '
            ' Gather information
            '

            Dim movieTitle
            movieTitle = folderNamespace.GetDetailsOf(fileName, MOVIE_TITLE)

            Dim movieYear
            movieYear = folderNamespace.GetDetailsOf(fileName, MOVIE_YEAR)

            Dim movieGenre
            movieGenre = folderNamespace.GetDetailsOf(fileName, MOVIE_GENRE)

            Dim movieLength
            movieLength = folderNamespace.GetDetailsOf(fileName, MOVIE_LENGTH)

            Dim movieSize
            movieSize = folderNamespace.GetDetailsOf(fileName, MOVIE_SIZE)

            Dim movieFileType
            movieFileType = folderNamespace.GetDetailsOf(fileName, MOVIE_FILE_TYPE)

            Dim movieFrameWidth
            movieFrameWidth = folderNamespace.GetDetailsOf(fileName, MOVIE_FRAME_WIDTH)

            Dim movieFrameHeight
            movieFrameHeight = folderNamespace.GetDetailsOf(fileName, MOVIE_FRAME_HEIGHT)

            Dim movieSubtitle
            movieSubtitle = folderNamespace.GetDetailsOf(fileName, MOVIE_SUBTITLE)

            Dim movieFolderStructure
            movieFolderStructure = folderNamespace.GetDetailsOf(fileName, MOVIE_FOLDER_STRUCTURE)

            Dim movieFile
            movieFile = folderNamespace.GetDetailsOf(fileName, MOVIE_FILE)

            '
            ' Process the information
            '

            ' Title
            If (len(movieTitle) < 1) Then
                movieTitle = movieFile
            Else
                movieTitle = chr(34) & movieTitle & chr(34)
            End If

            ' Year
            If (len(movieYear) < 1) Then
                movieYear = "--"
            End If

            ' Genre
            If (len(movieGenre) < 1) Then
                movieGenre = "--"
            End If

            ' Length
            If (len(movieLength) < 1) Then
                movieLength = "--"
            End If

            ' Size
            If (len(movieSize) < 1) Then
                movieSize = "--"
            End If

            ' File type
            If (len(movieFileType) < 1) Then
                movieFileType = "--"
            Else
                movieFileType = Left(movieFileType, 3)
            End If

            ' Frame width
            If (len(movieFrameWidth) < 1) Then
                movieFrameWidth = "--"
            End If

            ' Frame height
            If (len(movieFrameHeight) < 1) Then
                movieFrameHeight = "--"
            End If

            ' Subtitle
            If (len(movieSubtitle) < 1) Then
                movieSubtitle = "--"
            End If

            ' Folder structure
            If (len(movieFolderStructure) < 1) Then
                movieFolderStructure = "--"
            Else
                rootKey = Left(movieFolderStructure, 3)
                movieFolderStructure = Replace(movieFolderStructure, rootKey, rootDirs(rootKey) & "  -->  ")

                'movieFolderStructure = Right(movieFolderStructure, len(movieFolderStructure)-2)
                movieFolderStructure = Replace(movieFolderStructure, "\", "  -->  ")
            End If

            ' File
            If (len(movieFile) < 1) Then
                movieFile = "--"
            End If

            ' Write to file
            WriteToFile(movieTitle & ";" & movieYear & ";" & movieGenre & ";" & movieLength & ";" & movieSize & ";" & movieFileType & ";" & movieFrameWidth & ";" & movieFrameHeight & ";" & movieSubtitle & ";" & movieFolderStructure & ";" & movieFile)

            intFileCount = intFileCount + 1
        End If

        ' Display the operation progress
        WriteToConsoleR "Items written to file: " & intFileCount
    Next
End Function

Sub finishIt(sTime, eTime, title_count)
    Dim dialogMsg

    dialogMsg = "COMPLETED!!" & Chr(10) & Chr(10)
    dialogMsg = dialogMsg & "Started: " & sTime & Chr(10)
    dialogMsg = dialogMsg & "Finished: " & eTime & Chr(10) & Chr(10)
    dialogMsg = dialogMsg & title_count & " titles written to file."

    WriteToDialog dialogMsg
End Sub


' ----------------------------------------------------------------------
'  Main Processing
' ----------------------------------------------------------------------

'
' Do the Magic!!
'
Dim startTime
Dim endTime

startTime = Time()

initializeOutputFile

Dim inputDirKey

For Each inputDirKey in rootDirs.keys
    strRootDirectory = inputDirKey
    getRootDirectory(strRootDirectory)
    iterateThroughCategoryFolders
Next

endTime = Time()

'
' Voila!
'
finishIt startTime, endTime, intFileCount
