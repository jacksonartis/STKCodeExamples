' This script will take a custom report with just one data provider and a set of time steps and turn it into a scalar file
'
' Author: Jackson Artis
' Date: 16 Oct 2020

' read report
set oArgs = wscript.arguments
path = oArgs(0)
set oArgs = nothing
fileArray = getFileToArray(path)

' get number of data points in report
fileLength = UBound(fileArray)


counter = 0
numberSubtractor = 0
currentLine = trim(fileArray(counter))
currentWordArray = split(currentLine, " ")


first = CStr(currentWordArray(0))

while first <> "Event"
	counter = counter + 1
	numberSubtractor = numberSubtractor + 1
	currentLine = trim(fileArray(counter))
	currentWordArray = split(currentLine, " ")
	If ubound(currentWordArray) > 0 Then
		first = CStr(currentWordArray(0))
	End If 
	
wend 

counter = counter + 2
passnumber = 1
numberOfDataPoints = fileLength - numberSubtractor - 1


' write header
fileStr = "stk.v.12.0" & vbcrlf
fileStr = fileStr & vbcrlf
fileStr = fileStr & "BEGIN Data" & vbcrlf
fileStr = fileStr & "TimeFormat	UTCG" & vbcrlf
fileStr = fileStr & "NumberOfIntervals 1   " & vbcrlf
fileStr = fileStr & "	BEGIN Interval" & vbcrlf
fileStr = fileStr & "	  NumberOfPoints " & numberOfDataPoints & vbcrlf
fileStr = fileStr & "	  BEGIN TimeValues"& vbcrlf

Dim newArray(100)
for i = counter to ubound(fileArray)
	'get values
	
	thisData = trim(fileArray(i))
	thisDataArray = split(thisData, " ")
	limit = ubound(thisDataArray)
	arrayCounter = 0 
	for j = 0 to limit
		If thisDataArray(j) <> "" Then
			newArray(arrayCounter) = thisDataArray(j)
			arrayCounter = arrayCounter + 1
		End If
	next 
	
	
	'fileStr = fileStr & vbTab & thisDataArray(0) & vbTab & thisDataArray(1) & vbTab & thisDataArray(2) & vbTab & thisDataArray(3) & vbTab & thisDataArray(4) & vbcrlf
	fileStr = fileStr & vbTab & vbTab & newArray(2) & " " & newArray(3) & " " & newArray(4) & " " & newArray(8) & vbTab & passnumber & vbcrlf	
	passnumber = passnumber + 1
next 	



' write footer
fileStr = fileStr & vbcrlf
fileStr = fileStr & "	END TimeValues" & vbcrlf 
fileStr = fileStr & "  END Interval" & vbcrlf   
fileStr = fileStr & "END Data" & vbcrlf   


' write everything to report
writeStrToFile path, fileStr

' the end


' ***** read report and write it to array *****
Function getFileToArray(path)
   set fso = createObject("Scripting.FileSystemObject")
   set fIn = fso.OpenTextFile(path, 1)
   i = 0
   while not fIn.AtEndOfStream
      redim preserve fileArray(i)
      fileArray(i) = fIn.ReadLine
      i = i + 1
   wend
   set fIn = nothing
   getFileToArray = fileArray
End Function

' ***** write everything to report *****
Sub writeStrToFile(path,string) 'saves a string to the given path
   dim fso, f
   set fso = CreateObject("Scripting.FileSystemObject")
   set f = fso.CreateTextFile(path)

   f.write string

   set f = nothing
   set fso = nothing
End Sub