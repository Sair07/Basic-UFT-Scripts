Dim almURL, almLauncherPath
Dim StepCount
StepCount = 1
' Get the current date and time
Dim currentTime
currentTime = Now

' Format the current time to make it file-path friendly
Dim formattedTime
formattedTime = Replace(currentTime, ":", "_")  ' Replace colon with underscore
formattedTime = Replace(formattedTime, " ", "_") ' Replace space with underscore
formattedTime = Replace(formattedTime, "/", "-") ' Replace slash with dash
formattedTime = Replace(formattedTime, "\", "-")

almURL = "https://dev-testalm.cytiva.net/qcbin"
almLauncherPath = "C:\Users\cyt.sb3644810\Desktop\ALM Client Launcher 3.1\ALM Client Launcher 3.1.exe"

' Launch ALM using ALM Client Launcher
' Run the command using Shell
Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run """" & almLauncherPath & """"
Set objShell = Nothing

Sub UpdateRunStep(TestStatus, TestDescription, stepExpected, stepActual, screenshot)
	
	Set myCurentRun = QCUtil.CurrentRun	
	Set	myStepFactory = myCurentRun.StepFactory
	myStepFactory.AddItem("Test Step")
	Set myStepList = myStepFactory.NewList("")
	stepID = myStepList.Count
	Set  RunStep = myStepList.Item(stepID)	
	Set attachFactory = RunStep.Attachments	
	myStepList.Item(stepID).Field("ST_STATUS") = TestStatus
	myStepList.Item(stepID ).Field("ST_DESCRIPTION") = TestDescription
	myStepList.Item(stepID).Field("ST_EXPECTED") = stepExpected
	myStepList.Item(stepID ).Field("ST_ACTUAL") = stepActual
	'add attachment to the current step in Run
	Set attachment = attachFactory.AddItem (Null)
	attachment.FileName = screenshot
	attachment.Type = 1
	attachment.Post
	myStepList.Post
	'StepCount = StepCount + 1
	'code to add Attachment to the RUN
	'Set objAttachmentFactory = myCurentRun.Attachments
	'Set objAttachment = objAttachmentFactory.AddItem(Null)
	'objAttachment.FileName = attachment
	'objAttachment.Type = 1
	'objAttachment.Post()
	
End Sub

If UIAWindow("ALM Client Launcher").Exist Then
	UIAWindow("ALM Client Launcher").Activate
	'Reporter.ReportEvent micPass, "ALM Client Launcher Launched", "ALM Client Launcher opened successfully."	
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Launch ALM Client Launcher_" + formattedTime + ".png"	
    	Desktop.CaptureBitmap FilePath
    	UpdateRunStep "Passed", "Launching ALM Client Launcher." ,"ALM Client Launcher Launched", "ALM Client Launcher opened successfully.", FilePath    
	FilePath = ""    	
Else 
	'Reporter.ReportEvent micFail, "ALM Client Launcher Launch Failed", "ALM Client Launcher did not open."
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Launch ALM Client Launcher_Failed_" + formattedTime + ".png"	
    	Desktop.CaptureBitmap FilePath
    	UpdateRunStep "Passed", "Launching ALM Client Launcher." ,"ALM Client Launcher Launched", "ALM Client Launcher opened successfully.", FilePath    
	FilePath = ""  
	Abort
End If

UIAWindow("ALM Client Launcher").UIAObject("m_urlEditBoxPanel").Click
UIAWindow("ALM Client Launcher").UIAObject("m_dropdownButtonPictureBox").Click
UIAWindow("ALM Client Launcher").UIAEdit("m_urlTextBox").SetValue almURL

wait(2)
UIAWindow("ALM Client Launcher").UIAObject("m_goButtonPictureBox").Click
'wait for the ALM login page to load
wait(5)
'update the step with screenshot
FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Launch ALM Login_" + formattedTime + ".png"	
Desktop.CaptureBitmap FilePath
UpdateRunStep "Passed", "Launching ALM Login Page." ,"ALM Login Page", "ALM Login Page opened successfully.", FilePath    
FilePath = "" 
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAEdit("m_user").UIAEdit("m_user_EmbeddableTextBox").SetValue "admin.user"
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAEdit("m_password").UIAEdit("[Editor] Edit Area").SetValue "5QWU&.vW4$"
wait(2)
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAButton("AUTHENTICATE").Click
wait(2)
If Err.Number = 0 Then
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Authenticated successfully" + formattedTime + ".png"	
	Desktop.CaptureBitmap FilePath
	UpdateRunStep "Passed", "Authenticate" ,"Validate ALM connected", "Authenticated successfully.", FilePath    
	FilePath = "" 
Else
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Authenticated Failed" + formattedTime + ".png"	
	Desktop.CaptureBitmap FilePath
	UpdateRunStep "Failed", "Authenticate" ,"Validate ALM connected", "Authenticated Failed.", FilePath    
	FilePath = ""
End  IF

'select domain & project
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAComboBox("m_domains").UIAButton("[Editor] dropdown button").Click
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAComboBox("m_domains").UIAEdit("[Editor] Edit Area").SetValue "LIFESCIENCES"
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAComboBox("m_projects").UIAButton("[Editor] dropdown button").Click
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAComboBox("m_projects").UIAEdit("[Editor] Edit Area").SetValue "FlexFactory"
'login
UIAWindow("MainForm").UIATab("m_loginTabControl").UIAButton("LOGIN").Click
wait(5)
If Err.Number = 0 Then
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Logged in successfully" + formattedTime + ".png"	
	Desktop.CaptureBitmap FilePath
	UpdateRunStep "Passed", "Log in to ALM Project" ,"Logged in to ALM", "Logged into ALM successfully.", FilePath    
	FilePath = "" 
Else
	FilePath = "C:\Users\cyt.sb3644810\Desktop\Test screenshot\Logged in Failed" + formattedTime + ".png"	
	Desktop.CaptureBitmap FilePath
	UpdateRunStep "Failed", "Log- in to ALM Project" ,"Logged in to ALM", "Failed to Login.", FilePath    
	FilePath = ""
	Abort
End  IF

