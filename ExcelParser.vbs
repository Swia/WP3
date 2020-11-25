Dim dateToSearch, dateFromCell
Dim objFSO, objFSO2, objFile, outFile, inFile2, strLine, endFlag, checkCounter   
Dim i, k, j, result,  finalResult, GUIDcounter, errorCounter, errArr, lastCommaChecker 
Set objExcel = CreateObject("Excel.Application") 
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\BezrukovMG\Documents\VBS\v_iv_room_fias.xlsx") 
'dateToSearch = FormatDateTime(Now, vbShortDate) 
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFSO2 = CreateObject("Scripting.FileSystemObject")
outFile = "C:\Users\BezrukovMG\Documents\VBS\result.txt"
'inFile2 = "D:\loads\VBS\list.txt"
Set objFile = objFSO.CreateTextFile(outFile, True)
'Set objFile2 = objFSO2.OpenTextFile(inFile2)
Set objSheet = objWorkbook.Worksheets(1)
checkCounter = 2
lastCommaChecker = false 
'WScript.Echo "Finished." 
'objExcel.Application.Quit 'close excel
'WScript.Quit
objFile.Write "{" & vbCRLF 
objFile.Write Chr(34) & "orderid" & Chr(34) & ": " & Chr(34) &  "5fbbc599769b841c5e556fed" & Chr(34) & "," & vbCRLF 
objFile.Write Chr(34) & "pathvalues" & Chr(34) & ":  [" & vbCRLF 
'objFile.Write "{" & vbCRLF 
do while objSheet.Cells(checkCounter,1).Value <> ""
	'MsgBox CStr(objSheet.Cells(checkCounter,2).Value)
	For errorCounter = 3 to 15 
		If (objSheet.Cells(checkCounter,errorCounter).Value) = "1" Then 
			errArr = Split (objSheet.Cells(1,errorCounter).Value , "_")
			'MsgBox errArr (1)
			'MsgBox errArr (2)
			If lastCommaChecker Then
				objFile.Write ", " & vbCRLF
			End if 
			
			objFile.Write   	"{ "  & vbCRLF 
			objFile.Write		Chr(34) & "path" & Chr(34)& ": " & Chr(34) & "/root/fiasGuid" & Chr(34) & " ," & vbCRLF 
			objFile.Write		Chr(34) & "value" & Chr(34) & ":  [" & Chr(34) &_
								CStr(objSheet.Cells(checkCounter,2).Value) & Chr(34) & "] " & vbCRLF 
			objFile.Write		" },"& vbCRLF
			objFile.Write   	"{ "  & vbCRLF 
			objFile.Write		Chr(34) & "path"& Chr(34) & ": " & Chr(34) & "/root/fiasErrType" & Chr(34) & " ," & vbCRLF 
			objFile.Write		Chr(34) & "value" & Chr(34) & ":  [" & Chr(34) &_
								errArr (1) & Chr(34)& "] " & vbCRLF 
			objFile.Write		" },"& vbCRLF
			objFile.Write   	"{ "  & vbCRLF 
			objFile.Write		Chr(34)& "path"& Chr(34) & ": " & Chr(34) & "/root/fiasErrCode" & Chr(34) & " ," & vbCRLF 
			objFile.Write		Chr(34) & "value" & Chr(34) & ":  [" & Chr(34) &_
								errArr (2) & Chr(34) & "] " & vbCRLF 
			objFile.Write		" }"
			lastCommaChecker = true
		End if	
	Next  
	
	checkCounter = checkCounter + 1 
loop

objFile.Write " ]"& vbCRLF 
objFile.Write " }"& vbCRLF 
	'For 2 to checkCounter
	'For ForCounter = 2 to 6
	'MsgBox CStr(objSheet.Cells(2+GUIDcounter,2).Value)
	'GUIDcounter = GUIDcounter + 1 	
	'Next
'
'	Do While NOT objFile2.AtEndOfStream	
'		strLine = objFile2.ReadLine
'		endFlag = 0
'		i = 2	
		
'		do while endFlag = 0
'			If (CStr(strLine) = CStr(objSheet.Cells(i,1).Value)) Then 
'				endFlag = 1
'				j = 2 
				'objFile.Write objSheet.Cells(i,1).Value & " - " & Year(objSheet.Cells(i,2).Value) & vbCRLF
'				objFile.Write Year(objSheet.Cells(i,2).Value) & vbCRLF
'			Elseif checkCounter = i Then
'				endFlag = 1
'				objFile.Write strLine & " - " & " no record" & vbCRLF
'			else 
'				i = i + 1
'			end if		
'		 loop 
'	loop 	
'MsgBox checkCounter


objFile.Close
'objFile2.Close
Set objFile = Nothing
'Set inFile2 = Nothing
'Set objFSO = Nothing
'Set objFSO2 = Nothing
 
WScript.Echo "Finished." 
objExcel.Application.Quit 'close excel
WScript.Quit