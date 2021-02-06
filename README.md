# uninstallbyname
vbs to be called from wpkg on Windows to neatly handle uninstalling programs by their name

`cscript //nologo uninstallbyname.vbs [options] "ProductName"`

cycle through the registry uninstall keys looking for:

* a match
* an uninstall string

if we find a relevant uninstall string, ammend it to be silent, and execute it

we only process:

* msiexec
* inno setup

if the uninstall command gives an error, then quit with that error.

to cope with the scenario where there is an msiexec which calls some other command which we don't handle, we do two passes,
the first one where we ignore commands we don't handle and the second pass where we return an error if any are left.

option      | description
------------|-----------------------------------------------------------------------------------------
ProductName | Display Name as it shows up in 'Programs and Features' or 'Add or Remove Programs'.  Match starts from character 1, if Display Name contains spaces, it must be enclosed in quotes
/? | this help text
/debug | show progress
/dryrun | don't actually do anything, /debug is implied

-# |errors
---|-------------------------------------------------------------------------------------
0 | uninstall command found and executed, it returned 0 (we've no idea if that indicates success)
0 | no match found, we tried to uninstall something which isn't there, and that's OK
? | this script will exit with any errors returned by the specific uninstall routine
-1 | incorrect number of parameters, probably none or search string unquoted
-2 | runaway count exceeded, check your search string.
-3 | you asked us to process something which has an uninstall string we can't handle

# TODO:
we need to add code to the msiexec option: to suppress reboots
msiexec /i "C:\package.msi" REBOOT=ReallySuppress
