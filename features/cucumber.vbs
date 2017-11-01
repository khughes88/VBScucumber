'**********************************************************************************
' The functions are part of cucumber.vbs which is the core execution driver for the feature files
'
'
'


Public gCurrentStep


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
					Wscript.Echo "Running feature: " & Split(Line," ")(1)
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
	On Error Resume Next
	Func=GenerateFuncWithArgs(StrStep)
	Execute Func
	If Err.Number=13 Then
		Wscript.Echo "Sub "& GenerateFuncDefWithArgs(StrStep) &vbLf &vbTab &"'Your code here" &vbLf& "End Sub"
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



a=Split(s,"""")
For i=1 To UBound(a) Step 2
	MsgBox a(i)
Next