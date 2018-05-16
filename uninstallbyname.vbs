' uninstallbyname.vbs

' 22/12/16  dce  add DryRun
' 24/12/16  dce  refactor to read all the registry, not just the ones that msiexec knows about
' 05/01/17  dce  if no match found exit 0
'                add /SP- to Inno setup uninstall command
' 27/03/17  dce  big rewrite to do two pass handling
'                and handle parameter passing
' 30/03/17  dce  handle msiexec command subsitution properly
' 19/04/17  dce  runaway count 4 to handle old Java installations

Sub Explain
WScript.Echo "cscript //nologo uninstallbyname.vbs [options] ""ProductName"""
WScript.Echo " "
WScript.Echo "cycle through the registry uninstall keys looking for:"
WScript.Echo "  * a match"
WScript.Echo "  * an uninstall string"
WScript.Echo "if we find a relevant uninstall string, ammend it to be silent, and execute it"
WScript.Echo "we only process:"
WScript.Echo "  * msiexec"
WScript.Echo "  * inno setup"
WScript.Echo "if the uninstall command gives an error, then quit with that error."
WScript.Echo "to cope with the scenario where there is an msiexec which calls some other command which we don't handle, we do two passes, "
WScript.Echo "the first one where we ignore commands we don't handle and the second pass where we return an error if any are left."
WScript.Echo " "
WScript.Echo "options: "
WScript.Echo "ProductName : Display Name as it shows up in 'Programs and Features' or 'Add or Remove Programs'"
WScript.Echo "              match starts from character 1, if Display Name contains spaces, it must be enclosed in quotes"
WScript.Echo "         /? : this help text"
WScript.Echo "     /debug : show progress"
WScript.Echo "    /dryrun : don't actually do anything, /debug is implied"
WScript.Echo " "
WScript.Echo "errors: "
WScript.Echo " 0 = uninstall command found and executed, it returned 0 (we've no idea if that indicates success)"
WScript.Echo " 0 = no match found, we tried to uninstall something which isn't there, and that's OK"
WScript.Echo " ? = this script will exit with any errors returned by the specific uninstall routine"
WScript.Echo "-1 = incorrect number of parameters, probably none or search string unquoted"
WScript.Echo "-2 = runaway count exceeded, check your search string."
WScript.Echo "-3 = you asked us to process something which has an uninstall string we can't handle"
End Sub

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
strComputer = "." 
strKey1 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
strKey2 = "Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"

blnDebugMode = False  'show progress
blnDryRun    = False  'don't actually do anything

intCompareType = 1   ' 0 = binary, 1 = text
strUninstallCmd = ""
intMinProdLen   = 5  ' minimum Product Name length
intUninstallErr = 0
intSuccessCount = 0
intRunawayCount = 0  ' to prevent runaway
intRunawayMax   = 4

' check we've been passed enough arguments
Set objArgs = Wscript.Arguments
if objArgs.Count < 1 Then Wscript.Quit -1

' arguments start from first = 0
For intArg = 0 to objArgs.Count -1
    thisArg = Lcase(objArgs(intArg))
    ' InStr([start,]stringtosearch, forsearchstring[,compareType])
    If InStr(1, thisArg, "/dry", intCompareType) = 1 Then 
        blnDryRun    = True  'don't actually do anything
        blnDebugMode = True  'show progress, implied by DryRun
    ElseIf InStr(1, thisArg, "/debug", intCompareType) = 1 Then 
        blnDebugMode = True  'show progress
    ElseIf InStr(1, thisArg, "/?", intCompareType) = 1 Then 
        Explain
        Wscript.Quit 0
    Else
        ' if called without spaces we might get here twice
        If strFindProduct <> "" Then 
            If blnDebugMode Then WScript.Echo "too many parameters"
            Wscript.Quit -1
        End If
        strFindProduct = thisArg
    End If
Next

' and we must have a string to search for by now
If Len(strFindProduct) < intMinProdLen Then 
    If blnDebugMode Then WScript.Echo """ProductName"" missing or too short"
    Wscript.Quit -1
End If

If blnDebugMode Then WScript.Echo "   Debug Mode: ON"
If blnDryRun    Then WScript.Echo "  DryRun Mode: ON"
If blnDebugMode Then WScript.Echo "searching for: " & strFindProduct

' Wscript.Quit -99

strRegKeys = Array( _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", _
    "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" _
)

' process through 1 time
FindUninstallCmd(1)

' if that failed, then quit here
If intUninstallErr <> 0 Then 
    If blnDebugMode Then WScript.Echo "uninstall err: " & intUninstallErr
    Wscript.Quit intUninstallErr
End If

' and then do it again looking for things we don't handle
FindUninstallCmd(2)

If blnDebugMode Then WScript.Echo "-------------------------------------------------------------------------------"

' if we've not found anything by now, then give up
If strUninstallCmd = "" Then 
    If blnDebugMode Then WScript.Echo " no match for: " & strFindProduct
    WScript.Quit 0
End If

' if everything completed with error or don't care, then we're done
If intRunawayCount = intSuccessCount Then
    If blnDebugMode Then WScript.Echo "success uninstall: " & strFindProduct
    WScript.Quit 0
End If

' otherwise quit with the last uninstall error we had
If blnDebugMode Then WScript.Echo "quit with err: " & intUninstallErr
Wscript.Quit intUninstallErr

Function FindUninstallCmd (intPass)
    For Each strKey in strRegKeys
        If blnDebugMode Then WScript.Echo "-------------------------------------------------------------------------------"
        If blnDebugMode Then WScript.Echo "searching key: " & strKey
        Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 

        ' get all the registry keys into an array
        objReg.EnumKey HKLM, strKey, arrSubkeys 

        ' spin through the uninstall keys array, looking for a match on DisplayName
        For Each strSubkey In arrSubkeys 
            ' get the display name, GetStringValue returns 0 if we get a hit.
            intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, "DisplayName", strDisplayName) 
            If intRet1 <> 0 Then ' if we didn't get a value, then try the quiet display name
                objReg.GetStringValue HKLM, strKey & strSubkey, "QuietDisplayName", strDisplayName 
            End If 

            ' look for a match from the start (1) of the string?
            If InStr(1, strDisplayName, strFindProduct, intCompareType) = 1 Then
                If blnDebugMode Then WScript.Echo vbCRLF & "pass: " & intPass & " found: " & strDisplayName

                ' if there's a QuietUninstallString, then get that, if we try to read a value that doesn't exist that's OK
                objReg.GetStringValue HKLM, strKey & strSubkey, "QuietUninstallString", strUninstallCmd 
                If strUninstallCmd <> "" Then ProcessUninstall strUninstallCmd,intPass 

                ' and try the default UninstallString, we might have deleted it by now
                objReg.GetStringValue HKLM, strKey & strSubkey, "UninstallString", strUninstallCmd 
                If strUninstallCmd <> "" Then ProcessUninstall strUninstallCmd,intPass
            End If
        Next 
    Next
End Function


' -------------------------------------------------------------------------------------------------------
Function ProcessUninstall(strUninstallCmd,intPass)
    'see if it's one of the things we can deal with and process it
    intRunawayCount = intRunawayCount + 1
    If blnDebugMode Then 
        WScript.Echo "runaway count: " & intRunawayCount & " pass " & intPass
        WScript.Echo "uninstall raw: " & strUninstallCmd
    End If
    If intRunawayCount > intRunawayMax Then 
        If blnDebugMode Then WScript.Echo "runaway protection count too large, quitting..."
        Wscript.Quit -2
    End If

    Set objShell = WScript.CreateObject("WScript.Shell")

    If InStr(1, strUninstallCmd, "msiexec", intCompareType) Then
        ' MSIEXEC gives us things like:
        ' MsiExec.exe /X{07C69B3A-62B3-41BF-82EE-B3A87BD6EA0C}
        ' or even 
        ' MsiExec.exe /I{07C69B3A-62B3-41BF-82EE-B3A87BD6EA0C} <-- invokes the interactive install routine
        ' we want to change all these to /qn /x

        ' we have a (hopefully) valid uninstall command, but we need to make it silent and non interactive (/qn /x)
        If blnDebugMode Then WScript.Echo "uninstall raw: " & strUninstallCmd
        strUninstallCmd = Replace (strUninstallCmd, "/i", "/x",1 ,-1 , intCompareType) ' sometimes it only gives us interactive
        If InStr(1, strUninstallCmd, "/qn", intCompareType) = 0 Then
            strUninstallCmd = Replace (strUninstallCmd, "/x", "/qn /x",1 ,-1 , intCompareType) ' and add silent
        End If   
        If blnDebugMode Then WScript.Echo "uninstall cmd: " & strUninstallCmd
        ' run the uninstaller, in a normal window, wait for it to complete
        If NOT blnDryRun Then intUninstallErr = objShell.Run (strUninstallCmd, 1, true) 
        ' if that worked OK, then carry on - otherwise quit
        If blnDebugMode Then WScript.Echo "uninstall err: " & intUninstallErr

    ElseIf InStr(1, strUninstallCmd, "\unins0", intCompareType) Then
        ' Inno setup gives us things like this:
        ' "C:\Program Files (x86)\KeePass Password Safe 2\unins000.exe" /SILENT
        ' or even this
        ' "C:\WINDOWS\SysWOW64\unins000.exe" /SILENT
        ' we want to make these ones /SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART

        ' we have a (hopefully) valid uninstall command, but we need to make it silent and non interactive
        strUninstallCmd = Replace (strUninstallCmd, " /SILENT", "",1 ,-1 , intCompareType) 
        strUninstallCmd = strUninstallCmd & " /SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
        If blnDebugMode Then WScript.Echo "uninstall cmd: " & strUninstallCmd
        ' run the uninstaller, in a normal window, wait for it to complete
        If NOT blnDryRun Then intUninstallErr = objShell.Run (strUninstallCmd, 1, true) 
        ' if that worked OK, then carry on - otherwise quit
        If blnDebugMode Then WScript.Echo "uninstall err: " & intUninstallErr

    Else
        if intPass = 2 Then 
            ' other things may have their own commands e.g.
            ' "c:\Program Files\AVG\Avg2013 Admin\Common\avgmfapx.exe" /ProductType=Admin /AppMode=SETUP /Uninstall
            ' "c:\Program Files (x86)\Common Files\Adobe AIR\Versions\1.0\Resources\Adobe AIR Updater.exe" -arp:uninstall
            ' we can't code for all of these, so we'll just give up with an error
            If blnDebugMode Then WScript.Echo "this script only processes msiexec / inno setup, giving up"
            intUninstallErr = -3
            If blnDebugMode Then WScript.Echo "uninstall err: " & intUninstallErr
        Else
            If blnDebugMode Then WScript.Echo "ignored at pass = " & intPass
        End If
    End If
    
    ' count the number of successes
    If intUninstallErr = 0 Then intSuccessCount = intSuccessCount + 1
    
End Function

