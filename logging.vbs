Option Explicit

Public Function logging(intLogLevel, logMessage)
	'This function is used to log messages into log file in the current directory'
	'intLogLevel: can be set to INFO,ERROR,DEBUG(1,2,3 respectively), if any other value is passed it defaults to DEBUG
	'logMessage: is the log message that wil be logged in the log file
	
	'******************************
		'CONSTANTS'		
	'******************************
	Const FORAPPENDING = 8
	Const INFO = 1
	Const ERRORS = 2
	Const DEBUGGING = 3
	Const FILEEXTENSION = ".log"
	Const LOGFILENAME = "file_validation"

	'******************************
		'Variables
	'******************************
	Dim oFso,oFile
	Dim oSh
	Dim getCurrentDir, logFile
	
	Set oSh = CreateObject("WScript.Shell")
	getCurrentDir = oSh.CurrentDirectory  'Get the current working directory'
	logFile = getCurrentDir&"\"&LOGFILENAME&"_"&GetCurrentFormattedDateTime("date")&FILEEXTENSION
	
	'Creare a new file'
	Set oFso = CreateObject("Scripting.FileSystemObject")
	If oFso.FileExists(logFile) Then
		Set oFile = oFso.OpenTextFile(logFile,FORAPPENDING)
	Else
		Set oFile = oFso.CreateTextFile(logFile,True)
	End If
	
	Select Case intLogLevel
		Case 1 :
			logMessage = ""&GetCurrentFormattedDateTime("default")&":INFO--	"&logMessage	
		Case 2 :
			logMessage = ""&GetCurrentFormattedDateTime("default")&":ERROR--	"&logMessage
		Case 3 :
			logMessage = ""&GetCurrentFormattedDateTime("default")&":DEBUG--	"&logMessage
		Case Else:
			logMessage = ""&GetCurrentFormattedDateTime("default")&":DEBUG--	InValid Log level set; Defaulting to DEBUG:	"&logMessage
	End Select
	
	oFile.WriteLine(logMessage)
	oFile.Close
	
	Set oFso = Nothing
	Set oSh = Nothing
	
End Function

Public Function GetCurrentFormattedDateTime(ret_dte)

	'This function is used to format the current date and time'
	'ret_dte : if the value "date" is passed, then it return YYYYMMDD else YYYYMMDDHHMMSS'
	
	Dim strCurrentDate
	Dim strYear, strMonth, strDay
	Dim strHour, strMin, strSeconds
	
	strCurrentDate = Now()
	
	strHour = DatePart("h", strCurrentDate)
  	strMin = DatePart("n", strCurrentDate)
  	strSeconds = DatePart("s", strCurrentDate)
	
	strYear = Year(strCurrentDate)
	
	' Append 0 to months less than 2 digits
    If Month(strCurrentDate) < 10 Then
        strMonth = "0" & Month(strCurrentDate)
    Else
        strMonth = Month(strCurrentDate)
    End If
    ' Append 0 to days less than 2 digits
    If Day(strCurrentDate) < 10 Then
        strDay = "0" & Day(strCurrentDate)
    Else
        strDay = Day(strCurrentDate)
    End If
    
    If strHour < 10 Then
	  	strHour = "0" &strHour
  	End If
  
  	If strMin < 10 Then
  		strMin = "0" &strMin
  	End If
  
  	If strSeconds < 10 Then
	  	strSeconds = "0" &strSeconds
  	End If
    ' Return new date
    If ret_dte = "date" Then
		  GetCurrentFormattedDateTime = strYear & strMonth & strDay
	  Else
		  GetCurrentFormattedDateTime = strYear & strMonth & strDay & strHour & strMin & strSeconds
	  End If
End Function
