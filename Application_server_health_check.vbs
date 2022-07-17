'Description    : This script executes the health check up activity and send auto email
'Author : Sheetal Sapkal


'------------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Dim mI, mA
Dim mTitleBar, mDaysKeep, mMsg, mDir
Dim oWMIService, oService, oShell, oEvent, oFSO, oDrive, oFile, oLogFile, FSO,oFolder 
Dim mColListOfServices, mColLoggedEvents,f,fl
Dim mComputerName, mServicesList, status
Dim mEventLogFileName, mLogFileName,mFile, filecount
Dim mYYYY, mMM, mDD, mHH, mNN, mSS, mTimeDate
Dim mDayofToday, mDaysofWeek, mWeekDays
Dim mDrive, mFilePath, mFileName, mExtn, mReportData
Dim mFreeSize, mTotalSize, mUsedSize, mcurFree, minCapacity

Const ForReading = 1

mComputerName = "."
mLogFileName = "C:\temp\Dailycheck.htm"
Set oFSO = wscript.CreateObject("Scripting.FileSystemObject")
Set oLogFile = oFSO.OpenTextFile(mLogFileName, 2, True)

oLogFile.WriteLine "<HTML>"
oLogFile.WriteLine "<HEAD>"
oLogFile.WriteLine "<BODY>"

oLogFile.WriteLine "<b><centre>"
oLogFile.WriteLine " Application server health check " & "<br>"
oLogFile.WriteLine FormatDateTime(Now(),1) & ", " & FormatDateTime(Now(),3) & "</center></b><br>"
oLogFile.WriteLine "<hr>"


'------------------------------------------------------------------------------------------------------------------------------------------------------
'WARNING    : No Risk.
'------------------------------------------------------------------------------------------------------------------------------------------------------
oLogFile.WriteLine "<br><font color=black><b><u>Window service status</u></b>"
Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & mComputerName & "\root\cimv2")
Set mColListOfServices = oWMIService.ExecQuery("Select * from Win32_Service ")
For Each oService In mColListOfServices
'Dim mTitleBar
Dim objWMIService, objItem, objService, colListOfServices
'Dim mComputerName, mServicesList

'Mention your service name inside double quote of LCase
If LCase(oService.Name) = LCase("Mention your service name here") Then 
	   status = oService.State
           If status = "Stopped" Then 
                 oLogFile.WriteLine "<br><b><font color =red>" & oService.State & ":  " & oService.Name &"</font></b>"
           ELSE
                 oLogFile.WriteLine "<br>" & oService.State & ":  " & oService.Name
           END IF 
    END IF




Next
oLogFile.WriteLine "<br>"



'Description    : script checks the Disk Space of the given Drive.

'------------------------------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------------------------------------
oLogFile.WriteLine "<br><u><b>Diskspace Status - C:\</u></b>"
mDrive = "C"
minCapacity = 0.10
Set oFSO = wscript.CreateObject("Scripting.FileSystemObject")
Set oDrive = oFSO.GetDrive(mDrive)
mFreeSize = oDrive.FreeSpace
mTotalSize = oDrive.TotalSize
mcurFree = mFreeSize / mTotalSize
oLogFile.WriteLine "<br>Required Disk Space     : " & minCapacity * 100 & "%"
oLogFile.WriteLine "<br>Available Disk Space    : " & FormatNumber(mFreeSize / 1073741824, 2) & " GBytes"
oLogFile.WriteLine "<br><font color=green><b> Available Disk Space (%): " & FormatPercent(mcurFree, 2)&" </b></font>"


If mcurFree < minCapacity Then oLogFile.WriteLine "<br><font color=red>TOO LOW DISK SPACE</font>"
oLogFile.WriteLine "<br>"

'------------------------------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------------------------------------
oLogFile.WriteLine "<br><u><b>Diskspace Status - E:\</u></b>"
mDrive = "E"
minCapacity = 0.10
Set oFSO = wscript.CreateObject("Scripting.FileSystemObject")
Set oDrive = oFSO.GetDrive(mDrive)
mFreeSize = oDrive.FreeSpace
mTotalSize = oDrive.TotalSize
mcurFree = mFreeSize / mTotalSize
oLogFile.WriteLine "<br>Required Disk Space     : " & minCapacity * 100 & "%"
oLogFile.WriteLine "<br>Available Disk Space    : " & FormatNumber(mFreeSize / 1073741824, 2) & " GBytes"
oLogFile.WriteLine "<br><font color=green><b> Available Disk Space (%): " & FormatPercent(mcurFree, 2)&" </b></font>"

If mcurFree < minCapacity Then oLogFile.WriteLine "<br><font color=red>TOO LOW DISK SPACE </font>"
oLogFile.WriteLine "<br>"



Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile(mLogFileName, ForReading)

Do Until oFile.AtEndOfStream
    mReportData = mReportData & vbCr & oFile.ReadLine
Loop
'for i = count-8 to count-1

'mReportData =mReportData & list(i) & "<br>"

'Next 

oFile.Close
Set oFSO = Nothing
Set oFile = Nothing

'---------------------------------------------------------------
' Mailing Section
'---------------------------------------------------------------
Dim oEmail, oConfig
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSendUsingPort = 2
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Const cdoBasic = 1
Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"

Set oEmail = CreateObject("CDO.Message")
Set oConfig = CreateObject("CDO.Configuration")
oConfig.Fields.Item(cdoSendUsingMethod) = cdoSendUsingPort
'mention SMTP server name below in double quote
oConfig.Fields.Item(cdoSMTPServer) = "mention SMTP server name below in double quote"
oConfig.Fields.Item(cdoSMTPServerPort) = 25
oConfig.Fields.Update
Set oEmail.Configuration = oConfig
oEmail.To = ""
oEmail.From = ""
oEmail.Subject = "Health check up"
oEmail.HtmlBody = mReportData
oEmail.Send 
Set oEmail = Nothing
Set oConfig = Nothing
'wscript.echo "Mail Sent"

