Option Explicit

Dim currentDirectory, outputFilePath, folder, folderName, lastUnderscoreIndex, newFolderName, newFolderPath

' Create a file system object
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

' Set the current directory path to the current directory
currentDirectory = fso.GetAbsolutePathName(".")

' Set the output file path to "FolderList.txt" in the current directory
outputFilePath = fso.BuildPath(currentDirectory, "FolderList.txt")

' Create the output text file
Dim outputFile : Set outputFile = fso.CreateTextFile(outputFilePath, True)
outputFile.WriteLine("Folder List:")

' Get a collection of folders in the current directory
Dim folders : Set folders = fso.GetFolder(currentDirectory).SubFolders

' Loop through each folder in the collection
For Each folder In folders
    folderName = folder.Name
    
    ' Find the last underscore in the folder name
    lastUnderscoreIndex = InStrRev(folderName, "_")
    
    If lastUnderscoreIndex > 0 Then
        ' Remove all characters before the final underscore, including the final underscore
        newFolderName = Mid(folderName, lastUnderscoreIndex + 1)
        
        ' Build the new folder path in the current directory
        newFolderPath = currentDirectory & "\" & newFolderName
        
        ' Rename the folder
        folder.Name = newFolderName
        
        ' Check if the folder name contains "po"
        If InStr(1, newFolderName, "po", vbTextCompare) = 0 Then
            ' If not, create subfolders "RED", "WHITE", and "BLUE" within the renamed folder
            fso.CreateFolder(newFolderPath & "\RED")
            fso.CreateFolder(newFolderPath & "\WHITE")
            fso.CreateFolder(newFolderPath & "\BLUE")
        End If

        ' Write the folder name to the output text file
        outputFile.WriteLine(folder.Name)
    End If
Next

' Close the output text file
outputFile.Close

' Clean up
Set fso = Nothing

WScript.Echo "Folders renamed, subfolders created, and folder list written to FolderList.txt. - With love, from Josh"
