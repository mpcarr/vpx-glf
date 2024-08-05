'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class GlfDebugLogFile
	Private Filename
	Private TxtFileStream

	Public default Function init()
        Filename = cGameName + "_" & GetTimeStamp & "_debug_log.txt"
	  Set Init = Me
	End Function
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		 & LZ(Month(CurrTime),  2) & "-" _
		 & LZ(Day(CurrTime),	2) & "_" _
		 & LZ(Hour(CurrTime),   2) & "_" _
		 & LZ(Minute(CurrTime), 2) & "_" _
		 & LZ(Second(CurrTime), 2) & "_" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If glf_debugEnabled = True Then
			Dim FormattedMsg, Timestamp, fso, logFolder, TxtFileStream
	
			' Create a FileSystemObject
			Set fso = CreateObject("Scripting.FileSystemObject")
			
			' Define the log folder path
			logFolder = "glf_logs"
	
			' Check if the log folder exists, if not, create it
			If Not fso.FolderExists(logFolder) Then
				fso.CreateFolder logFolder
			End If
	
			' Open the log file for appending
			Set TxtFileStream = fso.OpenTextFile(logFolder & "\" & Filename, 8, True)
			
			' Get the current timestamp
			Timestamp = GetTimeStamp
			
			' Format the message
			FormattedMsg = Timestamp & ": " & label & ": " & message
			
			' Write the formatted message to the log file
			TxtFileStream.WriteLine FormattedMsg
			
			' Close the file stream
			TxtFileStream.Close
			
			' Print the message to the debug console
			Debug.Print label & ": " & message
		End If
	End Sub
End Class

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************