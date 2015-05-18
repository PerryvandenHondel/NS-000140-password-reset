''	---------------------------------------------------------------------------
''
''
''	SCRIPT
''		password-reset.vbs
''
''
''	SCRIPT_ID
''		140
''
'' 
''	DESCRIPTION
''		Reset a password of an account, extra unlock
''		Save password to a file to be sent to the requester of the reset
''
''
''	VERSION
''		01	2015-05-13	Initial version
''
'' 
''	FUNCTIONS AND SUBS
''		Function DsQueryGetDn
''		Function EncloseWithDQ
''		Function GetRandomCharString
''		Function GetScriptName
''		Function RunCommand
''		Sub DeleteFile
''		Sub ScriptDone
''		Sub ScriptInit
''		Sub ScriptRun
''		Sub ScriptUsage
'' 
''
''	---------------------------------------------------------------------------''



Option Explicit



Const	NEW_PASSWORD_LENGTH = 		12
Const 	FOR_WRITING =				2



Dim		gobjFso			'' Global object for File System Object
Dim		gstrRootDse
Dim		gstrUserName
Dim		gstrRef



Sub DeleteFile(sPath)
	''
	''	DeleteFile()
	''	
	''	Delete a file specified as "d:\folder\filename.ext"
	''
	''	sPath	The name of the file to delete.
	''
   	Dim oFSO
   	
   	Set oFSO = CreateObject("Scripting.FileSystemObject")
   	If oFSO.FileExists(sPath) Then
   		On Error Resume Next
		oFSO.DeleteFile sPath, True
		If Err.Number = 0 Then
			WScript.Echo "INFO: Deleted " & sPath
		End If
		
   	End If
   	Set oFSO = Nothing
End Sub '' DeleteFile



Function GetRandomCharString(ByVal intLen)
	''
	''	Returns a string of random chars of intLen length.
	'' 	                           12345678901234567890
	''	GetRandomCharString(20) >> 12ghyUjHsdbeH5fDsYt6
	''
	
	Dim		strValidChars	'' String with valid chars
	Dim		i				'' Random position of strValidChars
	Dim		r				'' Function return value
	Dim		x				'' Loop counter
	
	strValidChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz!@#$"
	
	For x = 1 to intLen
		Randomize
		'' intNumber = Int(1 - Len(strValidChars) * Rnd + 1)
		i = Int(Len(strValidChars) * Rnd + 1)
		
		r = r & Mid(strValidChars, i, 1)
	Next
	GetRandomCharString = r
End Function '' of Function GetRandomCharString



Function EncloseWithDQ(ByVal s)
	''
	''	Returns an enclosed string s with double quotes around it.
	''	Check for exising quotes before adding adding.
	''
	''	s > "s"
	''
	
	If Left(s, 1) <> Chr(34) Then
		s = Chr(34) & s
	End If
	
	If Right(s, 1) <> Chr(34) Then
		s = s & Chr(34)
	End If

	EncloseWithDQ = s
End Function '' of Function EncloseWithDQ



Function IsAccountLockedDn(ByVal strDn)
	''
	''	Check if an account is locked using command line tool ADFIND.EXE
	''
	''	Result:
	''		True		Account is locked
	''		False		Account is not locked.
	''
	''	Source:
	''		http://stackoverflow.com/questions/11795294/detect-if-an-active-directory-user-account-is-locked-using-ldap-in-python
	''
	Dim		strCommand
	Dim		strFilter
	Dim		objShell
	Dim		objExec
	Dim		strOutput
	Dim		intLockoutTimeValue
	Dim		blnResult
	
	strFilter = "(lockoutTime>=1)"
	strCommand = "adfind.exe -b " & EncloseWithDQ(strDn) & " -f " & EncloseWithDQ(strFilter) & " lockoutTime "
	strCommand = strCommand & " -csv -nocsvheader -nodn -nocsvq"

	'WScript.Echo strCommand
	
	Set objShell = CreateObject("WScript.Shell")
	Set objExec = objShell.Exec(strCommand)
	Do
		strOutput = objExec.Stdout.ReadLine()
		'WScript.Echo "strOutput: " & strOutput
		
		If Len(strOutput) > 0 Then
			'intLockoutTimeValue = Int(strOutput)
			blnResult = True
		Else	
			'intLockoutTimeValue = 0
			blnResult = False
		End If
	Loop While Not objExec.Stdout.atEndOfStream
	Set objExec = Nothing
	Set objShell = Nothing
	
	'WScript.Echo "intLockoutTimeValue=" & intLockoutTimeValue
	IsAccountLockedDn = blnResult
End Function '' of Function IsAccountLockedDn


Function DsQueryGetDn(ByVal strRootDse, ByVal strUserName)
	''
	''	Use the DSQUERY.EXE command to find a DN of a CN in a specific AD set by strRootDse
	''
	''		strRootDse:  	DC=domain,DC=ext
	''		strUserName: 	sAMAccountName
	''
	''		Returns: 	The DN of blank if not found.
	''
	
	Dim		c			''	Command
	Dim		r			''	Result
	Dim		objShell
	Dim		objExec
	Dim		strOutput
	
	If InStr(strUserName, "CN=") > 0 Then
		'' When the strCN already contains a Distinguished Name (DN), result = strUserName
		r = EncloseWithDQ(strUserName)
	Else
		'' No, we must search for the DN based on the CN
	
		c = "dsquery.exe "
		c = c & "* "
		c = c & strRootDse & " "
		c = c & "-filter (sAMAccountName=" & strUserName & ")"

		Set objShell = CreateObject("WScript.Shell")
		Set objExec = objShell.Exec(c)
		
		Do
			strOutput = objExec.Stdout.ReadLine()
		Loop While Not objExec.Stdout.atEndOfStream

		Set objExec = Nothing
		Set objShell = Nothing
		If Len(strOutput) > 0 Then
			r = EncloseWithDQ(strOutput)  '' BEWARE: r contains now " around the string, see "CN=name,OU=name,DC=domain,DC=nl"
		Else
			WScript.Echo "ERROR Could not find the Distinguished Name for " & strUserName & " in " & strRootDse
			r = ""
		End If
	End If
	DsQueryGetDn = r
End Function '' DsQueryGetDn



Function RunCommand(sCommandLine)
	''
	''	RunCommand(sCommandLine)
	''
	''	Run a DOS command and wait until execution is finished before the script can commence further.
	''
	''	Input
	''		sCommandLine	Contains the complete command line to execute 
	''
	Dim oShell
	Dim sCommand
	Dim	nReturn

	Set oShell = WScript.CreateObject("WScript.Shell")
	sCommand = "CMD /c " & sCommandLine
	' 0 = Console hidden, 1 = Console visible, 6 = In tool bar only
	'LogWrite "RunCommand(): " & sCommandLine
	nReturn = oShell.Run(sCommand, 6, True)
	Set oShell = Nothing
	RunCommand = nReturn 
End Function '' RunCommand



Function GetScriptName()
	''
	''	Returns the script name
	''	Removes script versioning (script-00)
	''	Removes .vbs extension
	''
	Dim	strReturn

	strReturn = WScript.ScriptName
	strReturn = Replace(strReturn, ".vbs", "")			'' Set the script name (without .vbs extention)

	If Mid(strReturn, Len(strReturn) - 2, 1) = "-" Then
		strReturn = Left(strReturn, Len(strReturn) - 3)
	End If
	GetScriptName = strReturn
End Function '' GetScriptName()



Sub ScriptUsage()
	WScript.Echo
	WScript.Echo "Usage:"
	WScript.Echo vbTab & GetScriptName() & " /rootdse:<rootdse> /user:<sam account name> /ref:<referencenumber>"
	WScript.Echo 
	WScript.Echo "Options:"
	WScript.Echo vbTab & "/rootdse               Root DSE of the domain to access"
	WScript.Echo vbTab & "/user                  User name of account"
	WScript.Echo vbTab & "/ref                   Reference of the request number"
	WScript.Echo 
	Wscript.Echo "Example:"
	Wscript.Echo vbTab & GetScriptName() & " /rootdse:DC=productie,DC=spoor,DC=nl /user:testuser /ref:653342"
	WScript.Echo 
	WScript.Quit(0)
End Sub '' ScriptUsage()



Sub ScriptInit()
	Dim 	colNamedArguments
	Dim		intArgumentCount
	
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
	
	Set colNamedArguments = WScript.Arguments.Named
	intArgumentCount = WScript.Arguments.Named.Count
	
	WScript.Echo intArgumentCount
	
	If intArgumentCount <> 3 Then
		Call ScriptUsage()
		Set colNamedArguments = Nothing
	ElseIf intArgumentCount = 3 Then
		gstrRootDse = WScript.Arguments.Named("rootdse")
		gstrUserName = WScript.Arguments.Named("user")
		gstrRef = WScript.Arguments.Named("ref")
	End If
	
	Set colNamedArguments = Nothing
End Sub '' of Sub ScriptInit



Sub ScriptRun()
	Dim		strDn
	Dim		strCommand
	Dim		strNewPassword
	Dim		strFile
	Dim		objFile
	
	
	WScript.Echo "ScriptRun()"
	WScript.Echo gstrRootDse
	WScript.Echo gstrUserName
	WScript.Echo gstrRef
	
	strDn = DsQueryGetDn(gstrRootDse, gstrUserName)
	WScript.Echo strDn
	If Len(strDn) > 0 Then
		'' We have found a DN of the user.
		
		strFile = gstrRef & ".txt"
		Call DeleteFile(strFile)
	
		Set objFile = gobjFso.OpenTextFile(strFile, FOR_WRITING, True)
		
		'' https://technet.microsoft.com/en-us/library/cc782255%28v=ws.10%29.aspx

		'' dsmod.exe user userdn -pwd newpassword 
		strNewPassword = GetRandomCharString(NEW_PASSWORD_LENGTH)
		strCommand = "dsmod.exe user " & EncloseWithDQ(strDn) & " -pwd " & strNewPassword
		
		objFile.WriteLine "Password reset done under " & gstrRef & " at " & Now()
		objFile.WriteLine
		objFile.WriteLine "New initial password is: " & strNewPassword
		WScript.Echo strCommand
		Call RunCommand(strCommand)
	
		'' dsmod.exe user userdn -mustchpwd yes
		strCommand = "dsmod.exe user " & EncloseWithDQ(strDn) & " -mustchpwd yes"
		WScript.Echo strCommand
		objFile.WriteLine
		objFile.WriteLine "User must change password at next logon"
		Call RunCommand(strCommand)
			
		If IsAccountLockedDn(strDn) = True Then
			'' The account is locked out, unlock it!
			'' dsmod.exe user userdn -disabled no
			strCommand = "dsmod.exe user " & EncloseWithDQ(strDn) & " -disabled no"
			'WScript.Echo strCommand
			objFile.WriteLine
			objFile.WriteLine "Extra: account is unlocked"
			Call RunCommand(strCommand)
		End If
		
		objFile.Close
		Set objFile = Nothing
	End If
End Sub '' of Sub ScriptRun



Sub ScriptDone()
	Set gobjFso = Nothing
End Sub '' of Sub ScriptDone



Call ScriptInit()
Call ScriptRun()
Call ScriptDone()
WScript.Quit(0)


'' End of Script