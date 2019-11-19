'File Name: Meterpreter_Defender.vbs
'Version: v1.1, 11/19/2019
'Author: Justin Grimes, 11/18/2019

'--------------------------------------------------
'Declare the variables to be used in this script.
'Undefined variables will halt script execution.
Option Explicit

dim oShell, oShell2, oFSO, scriptName, tempFile, appPath, logPath, strComputerName, fgcPath, i, tempData, defaultPerimiterFile, _
 strUserName, strSafeDate, strSafeTime, strDateTime, logFileName, strEventInfo, objLogFile, tempDir, tempDir0, tempDir1, _
 mailFile, objDangerHashCache, oFile, tempOutput, companyName, companyAbbr, companyDomain, toEmail, cacheData, mpdMode, _
 executionLimit, objFGCFile, sleepTime, infected, continuous
'--------------------------------------------------

  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The "scriptName" is the filename of this script.
  scriptName = "Meterpreter_Defender.vbs"
  ' The "appPath" is the full absolute path for the script directory, with trailing slash.
  appPath = "\\SERVER\AutomationScripts\Meterpreter_Defender\"
  ' The "logPath" is the full absolute path for where network-wide logs are stored.
  logPath = "\\SERVER\Logs"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "Company"
  ' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
  ' to have been sent by "COMPUTERNAME@domain.com"
  companyDomain = "Company.com"
  ' The "toEmail" is a valid email address where notifications will be sent.
  toEmail = "IT@Company.com"
  ' The "mpdMode" is the mode type for Meterpreter_Payload_Detection.exe. 
  ' To enable detection without any remediation, specify "IDS".
  ' To enable detection AND remediation (kill infected process), specify "IPS".
  mpdMode = "IDS"
  'This application runs in a loop. Each loop takes approximately 5 minutes to complete.
  'The "executionLimit" sets the number of loops which are performed before the entire application
  'is restarted.
  executionLimit = 3
  'The "sleepTime" is the amount of time in ms that the loop will wait for Meterpreter detection to occur.
  'Meterpreter detection happens constantly until Meterpreter_Payload_Detection.exe is restarted at the end of each loop.
  sleepTime = 300000
  'Setting "continuous" allows the script to run in the background indefinately or die after a designated amount of time.
  'The execution duration of this application is determined by the number of loop iteration & duration of the sleepTimer.
  'Increased execution time leads to an increase in resources required to scan cache report files for infection. 
  continuous = TRUE
  ' ----------

'--------------------------------------------------
'Set global variables for the session.
Set oShell = WScript.CreateObject("WScript.Shell")
Set oShell2 = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
tempDir0 = "C:\Program Files\Meterpreter_Defender"
tempDir1 = tempDir0 & "\Cache"
tempDir = tempDir1 & "\" & strComputerName
tempFile = tempDir & "\" & strComputerName & "-Cache.dat"
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
logFileName = logPath & "\" & strComputerName & "-" & strDateTime & "-Meterpreter_Defender.txt"
mailFile = tempDir & "\" & strComputerName & "-Meterpreter_Defender_Warning.mail"
i = 0
'--------------------------------------------------

'--------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated.
'Returns FALSE if the application is not elevated.
Function isUserAdmin()
  On Error Resume Next
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If Err.number = 0 Then 
    isUserAdmin = TRUE
  Else
    isUserAdmin = FALSE
  End If
  Err.Clear
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to restart the script with admin priviledges if required.
Function restartAsAdmin()
    oShell2.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to read files into memory as a string like PHP's file_get_contents.
'Inspired by https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Function fileGetContents(fgcPath) 
  'Set a handle to the file to be opened.
  Set objFGCFile = oFSO.OpenTextFile(fgcPath, 1)
  'Read the contents of the file into a string.
  fileGetContents = objFGCFile.ReadAll
  'Close the handle to the file we opened earlier in the function.
  objFGCFile.Close
  'Clean up unneeded memory.
  objFGCFile = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to verify the tempDir and clear the previous tempFile file and create a new one.
'Start by making C:\Program Files\Ransomware_Defender.
'Then make C:\Program Files\Ransomware_Defender\Cache.
'Then verify the cache files inside.
Function clearCache()
  If Not oFSO.FolderExists(tempDir0) Then
    oFSO.CreateFolder(tempDir0)
  End If
  If oFSO.FolderExists(tempDir0) Then
    If Not oFSO.FolderExists(tempDir1) Then
      oFSO.CreateFolder(tempDir1)
    End If
    If oFSO.FolderExists(tempDir1) Then
      If Not oFSO.FolderExists(tempDir) Then
        oFSO.CreateFolder(tempDir)
      End If
      If oFSO.FolderExists(tempDir) Then
        If oFSO.FileExists(tempFile) Then
          oFSO.DeleteFile(tempFile)
        End If
      End If
    End If
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
Function createLog(strEventInfo)
  If Not strEventInfo = "" Then
    Set objLogFile = oFSO.CreateTextFile(logFileName, TRUE)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a Warning.mail file. Use to prepare an email before calling sendEmail().
Function createEmail()
  If oFSO.FileExists(mailFile) Then
    oFSO.DeleteFile(mailFile)
  End If
  If Not oFSO.FileExists(mailFile) Then
    oFSO.CreateTextFile(mailFile)
  End If
  Set oFile = oFSO.CreateTextFile(mailFile, TRUE)
  oFile.Write "To: " & toEmail & vbNewLine & "From: " & strComputerName & "@" & companyDomain & vbNewLine & _
   "Subject: " & companyAbbr & " Meterpreter Defender Warning!!!" & vbNewLine & "This is an automatic email from the " & _ 
   companyName & " Network to notify you that a Meterpreter payload was detected on a domain workstation." & _
   vbNewLine & vbNewLine & "Please log-in and verify that the equipment listed below is secure." & vbNewLine & _
   vbNewLine & "USER NAME: " & strUserName & vbNewLine & "WORKSTATION: " & strComputerName & vbNewLine & _
   "This check was generated by " & strComputerName & " and is performed when Windows boots." & vbNewLine & vbNewLine & _
   "Script: """ & scriptName & """" 
  oFile.close
End Function
'--------------------------------------------------

'--------------------------------------------------
Function searchCache(cacheData)
  searchCache = FALSE
  If InStr(cacheData, "Meterpreter Process Found") > 0 Then
    searchCache = TRUE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to sleep script execution for 5 minutes.
Function searchSleep()
  WScript.Sleep(sleepTime)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for running SendMail to send a prepared Warning.mail email message.
Function sendEmail() 
  oShell.run "c:\Windows\System32\cmd.exe /c " & appPath & "sendmail.exe " & mailFile, 0, FALSE
End Function
'--------------------------------------------------

'--------------------------------------------------
Function launchMPD(mpdMode)
  oShell.Run "c:\Windows\System32\cmd.exe /c " & appPath & "Meterpreter_Payload_Detection.exe " & _
   mpdMode & " > """ & tempFile & """", 0, FALSE
End Function
'--------------------------------------------------

'--------------------------------------------------
Function killMPD()
  oShell.Run "c:\Windows\System32\cmd.exe /c taskkill /f /im Meterpreter_Payload_Detection.exe", 0, TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
'The main logic of the script which makes use of the code & functions above.

'Nake sure the script is being run with elevated priviledges.
If Not isUserAdmin() Then
  'Restart the script with elevated priviledges if needed.
  restartAsAdmin()
Else
  'Run until the executionLimit is reached before restarting the entire application (~15m worth of scanning).
  Do While i <= executionLimit
    WScript.Sleep(15000)
    'Verify that required directories exist & re-create a fresh cache file.
    clearCache()
    'Start Meterpreter_Payload_Detection.exe.
    launchMPD(mpdMode)
    'Sleep for 5 minutes to give Meterpreter_Payload_Detection.exe time to conduct a scan.
    searchSleep()
    'Load the Meterpreter_Payload_Detection.exe output from the temporary cache file into memory.
    cacheData = fileGetContents(tempFile)
    'Search the Meterpreter_Payload_Detection output contained in the cache file for indication of compromise.
    infected = searchCache(cacheData)
    'Check if any indication of compromise was detected.
    If infected Then
      'An indication of compromise was detected on the last iteration of the loop.
      'Create a logfile and copy the data containing IOC to the logfile.
      createLog("The workstation " & strComputerName & " has detected and mitigated a Meterpreter Payload! " & _
       "The relevant detection information is below." & vbNewLine & vbNewLine & cacheData)
      'Create an Warning.mail email file.
      createEmail()
      'Send the Warning.mail email file using Sendmail.exe.
      sendEmail()
    End If
    'Clean up the cacheData so it doesn't consume memory until the next iteration of the loop.
    cacheData = NULL
    'Increment the execution counter.
    i = i + 1
    'Kill the Meterpreter_Payload_Detection.exe process so that we can start a new one on the next loop.
    killMPD()
  Loop
  'Check if the "continuous" config entry is set so the script will restart automatically if required.
  If continuous Then 
    'Restart the script.
    restartAsAdmin()
  End If
End If
'Kill the current instance of the scipt to reset all variables & handles.
WScript.Quit()
'--------------------------------------------------