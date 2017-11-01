


Public gCurrentStep
'Nov 1
Public sUndefinedSnippets
'Nov 1
Public iUndefined, iSteps,iScenarios,iUndefinedForThisScenario,iUndefinedScenarios,sResultsAll,sResultsScenario
iUndefined=0
iSteps=0
'Nov 1
sUndefinedSnippets= vbLf&"You can implement step definitions for undefined steps with these snippets:" &vbLf
'Function to include files
Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub


'Read all the step definitions and load them for execution
'All the step definitions should be placed under the step_defs folder of the framework
ReadAllStepDefs "step_defs"


ReadAndRunFeatures


'Read all the features in the folder specified and run them
Private Sub ReadAndRunFeatures
	Dim oFolder,Folders,Item,x
	Set oFolder = CreateObject("Scripting.FileSystemObject")

	x= oFolder.GetAbsolutePathName(".")
	
	Set Folders = oFolder.GetFolder (x)
		
	For Each Item In Folders.Files
		If InStr(1,Item.Name,".feature")>0 Then
		Wscript.Echo Item.Name
			'Run a feature
			ReadFeatureFile Item.Name
		End if		
	Next	
End Sub


Sub ReadAllStepDefs(FolderName)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Root = objFSO.GetAbsolutePathName(".")
	FolderName = Root&"/"&FolderName
	'Check if the folder exists
	If not objFSO.FolderExists(FolderName) Then
		Exit Sub
	End If
	Set objFolder=objFSO.GetFolder(FolderName)
	
	For Each Item in objFolder.Files
		IncludeFile(FolderName&"/"&Item.Name)
	Next
	Set objFSO = Nothing
End Sub


Sub ReadFeatureFile(strFileName)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFileName)
	
	While Not objFile.AtEndOfStream
		Line = Replace(Trim(objFile.ReadLine),vbTab,"")
		If Line<>"" Then
			Select Case Split(Line," ")(0)
				Case "Feature:":
					'Nov 1 
					Wscript.Echo Line
					sResultFeed=sResultFeed &vbLf & "<div class='feature_heading'>"& Line &"</div>"
				Case "Given":
								gCurrentStep = "Given"
								ExecuteStep(Line)
				Case "Then":
								gCurrentStep = "Then"
								ExecuteStep(Line)
				Case "When":
								gCurrentStep = "When"
								ExecuteStep(Line)
				Case "And":
								Line = gCurrentStep&" "&Right(Line,Len(Line)-3)
								ExecuteStep(Line)
				Case "But":
								Line = gCurrentStep&" "&Right(Line,Len(Line)-3)
								ExecuteStep(Line)
				Case "Scenario:":
					Wscript.Echo vbLf & " " & Line
					iScenarios=iScenarios+1
					If iUndefinedForThisScenario>0 then
						iUndefinedScenarios=iUndefinedScenarios+1
					End If
					iUndefinedForThisScenario=0
					sResultsAll=sResultsAll &vbLf & "<div class='scenario_heading'>"& Line &"</div>"
				Case "Background:":
				Case "Scenario Outline:":
			End Select
		End If
		
	Wend
	Set objFSO = Nothing
	Set objFile = Nothing
End Sub



'**********************************************************************************
' The functions are part of cucumber.vbs to understand and execute the gherkin commands
'
'
'



'Execute the steps or generate templates
Sub ExecuteStep(StrStep)
	iSteps=iSteps+1
	'Nov 1 
	Wscript.Echo "  " & StrStep
	On Error Resume Next
	Func=GenerateFuncWithArgs(StrStep)
	Execute Func
	If Err.Number=13 Then
	iUndefined=iUndefined+1
	iUndefinedForThisScenario=iUndefinedForThisScenario+1
		'Nov 1
		sUndefinedSnippets=sUndefinedSnippets &vbLf& "Sub "& GenerateFuncDefWithArgs(StrStep) &vbLf &vbTab &"'Your code here" &vbLf& "End Sub" &vbLf
		'Wscript.Echo &vbLf& "Sub "& GenerateFuncDefWithArgs(StrStep) &vbLf &vbTab &"'Your code here" &vbLf& "End Sub"
		sResultFeed=sResultFeed &vbLf & "<div class='step_undefined'>"& strStep &"</div>"
	End If
	On Error Goto 0
End Sub


'Generates the function text to be implemented
Function GenerateFuncDefWithArgs(StrStep)
	StepText = StrStep
	ArgCount=0
	Args=""
	ArrStep = Split(StepText,"""")
	For Iter=1 To UBound(ArrStep) Step 2
		ArgCount=ArgCount+1
		StepText=Replace(StepText,""""&ArrStep(Iter)&"""","")
		Args=Args&",Arg"&ArgCount
	Next 
	If Args<>"" Then
		Args=Right(Args,Len(Args)-1)
		StepText = Replace(Replace(Trim(StepText)," ","_")&"("&Args&")","__","_")
		GenerateFuncDefWithArgs=StepText
	Else
		StepText = Replace(Trim(StepText)," ","_")
		GenerateFuncDefWithArgs=StepText
	End If
End Function

'Generates the function text to be executed
Function GenerateFuncWithArgs(StrStep)
	StepText = StrStep
	Args=""
	ArrStep = Split(StepText,"""")
	For Iter=1 To UBound(ArrStep) Step 2
		StepText=Replace(StepText,""""&ArrStep(Iter)&"""","")
		Args=Args&","&""""&ArrStep(Iter)&""""
	Next
	If Args<>"" Then
		Args=Right(Args,Len(Args)-1)
		StepText = Replace(Replace(Trim(StepText)," ","_")&" "&Args&"","__","_")
		GenerateFuncWithArgs=StepText
	Else
		StepText = Replace(Trim(StepText)," ","_")
		GenerateFuncWithArgs=StepText
	End If
	
End Function

'Nov 1
Function GenerateResultFile(sResultFeed)

	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("ResultTemplate.html", ForReading)
	strText = objFile.ReadAll
	objFile.Close
	strNewText = Replace(strText, "<result>", sResultFeed)
	objFSO.CreateTextFile("Result.html")
	Set objFile = objFSO.OpenTextFile("Result.html", ForWriting)
	objFile.WriteLine strNewText
	objFile.Close


End Function

GenerateResultFile(sResultsAll)

If iUndefinedForThisScenario>0 then
	iUndefinedScenarios=iUndefinedScenarios+1
End If


WScript.Echo vbLf 

If iUndefinedScenarios>0 then
	'Nov 1
	If iScenarios>0 Then
		WScript.Echo  iScenarios & " scenarios (" & iUndefinedScenarios & " undefined)"
	Else
		WScript.Echo iScenarios & " scenario  (" & iUndefinedScenarios & " undefined)"
	End If 
Else
	'Nov 1
	If iScenarios>0 Then
		WScript.Echo  iScenarios & " scenarios"
	Else
		WScript.Echo iScenarios & " scenario"
	End If 
End If

If iUndefined>0 Then
	WScript.Echo  iSteps & " steps ("&iUndefined & " undefined)"
Else
	WScript.Echo iSteps & " steps "
End If
'Nov 1
If iUndefined>0 Then
	WScript.Echo sUndefinedSnippets
End If

a=Split(s,"""")
For i=1 To UBound(a) Step 2
	MsgBox a(i)
Next