'Just Testing
'Sample Function Call:
'Call fn_mergeFiles("C:\temp1","txt")

' Gets all the files in the specified folder
' Reads the content of every file and writes them into one single file

Public Function fn_mergeFiles(folderPath,fileExtension)
	Const ForReading =1
	startTime = Now()
	
	'On Error Resume Next
	
	Dim line_count: line_count = 0
	Dim objFSO, oFolder, filename, objOutputFile, strComputer, objWMIService, objFile, objExtension
	Dim FileList, final_strTxt, orig_fileName, orig_fileName_1
	Dim objTextFile, strText, arr_strTxt, int_1

	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	Set oFolder = objFSO.GetFolder(folderPath)		'Get the Folder Name into the object
	
	filename = "files_combined.txt"
	Set objOutputFile = objFSO.CreateTextFile(""&folderPath&"\"&filename)  ' Create a new file where all the Tic File contents will be placed into
	
	strComputer = "."   'Provide the  Hostname or . refers to the computer where this script is running
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")  ' create a object for Windows Management Instrumentation
	
	'An Associators Of query does pretty much what the name implies: it enables you to associate two WMI classes. 
	'In this case, we’re associating Win32_Directory (the class that lets us manage folders) with CIM_DataFile (the class that lets us manage files).
	'CIM: Common Information Model
	
	Set FileList = objWMIService.ExecQuery _ 
    ("ASSOCIATORS OF {Win32_Directory.Name='"&folderPath&"'} Where " _
        & "ResultClass = CIM_DataFile") 

    For Each objFile In FileList  'Loop through every file in the folder
    	final_strTxt = ""

    	If objFile.extension = fileExtension Then
    		file_count = file_count +1
	    	'Get the file name:
	    	orig_fileName = Split(objFile.Name, "\",-1,1)
		    orig_fileName_1 = orig_fileName(UBound(orig_fileName))

	    	'Open file for reading, Read the whole file and then split them into single lines at every new line
		    Set objTextFile = objFSO.OpenTextFile(objFile.Name, ForReading)
		    strText = objTextFile.ReadAll
		    arr_strTxt =Split(strText, vbCrLf, -1 ,1)
		    
		    'For every line append the file name at the end		    
		    For int_1 = 0 To UBound(arr_strTxt) -1
		    	If line_count = 0 Then
		    		final_strTxt = arr_strTxt(int_1) & "|" & orig_fileName_1
		    		line_count = 1
		    	Else
		    		final_strTxt = final_strTxt & vbCrLf & arr_strTxt(int_1) & "|" & orig_fileName_1	
		    	End If
		    Next
		    objOutputFile.WriteLine final_strTxt
		End If
	Next
 
	objOutputFile.Close
	
	If file_count = 0 Then  'If there are no files to merge then delete the empty file that was created'
		objFSO.DeleteFile(""&folderPath&"\"&filename)
		Set oFolder = Nothing
		Set objFSO = Nothing
		Exit Function
	End If
	
	If Err.Number <> 0 Then 
		'if any error occurs, then return the err.description'
		fn_mergeFiles = Err.Description
		Err.Number = 0
	Else
		'if no errors then return True '
		fn_mergeFiles = True
	End If
	
	Set oFolder = Nothing
	Set objFSO = Nothing
	
	On Error Goto 0
	endTime = Now()
	Msgbox "Time taken to extract and load "&file_count&" files into Access DB: "&DateDiff("s", startTime, endTime)&" seconds"
End Function
