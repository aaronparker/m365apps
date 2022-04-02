'=======================================================================================================
' Name: OffScrub_O15msi.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2011-2015 Microsoft Corporation
' Script to remove (scrub) Office 2013 MSI products
' when a regular uninstall is no longer possible
'=======================================================================================================
Option Explicit

Dim sDefault
'=======================================================================================================
'[INI] Section for script behavior customizations

'Pre-configure the SKU's to remove.
'Only for use without command line parameters
'Example: sDefault = "CLIENTALL"
sDefault = "" 

'DO NOT CUSTOMIZE BELOW THIS LINE!
'=======================================================================================================


Const SCRIPTVERSION = "2.10"
Const SCRIPTFILE    = "OffScrub_O15msi.vbs"
Const SCRIPTNAME    = "OffScrub_O15msi"
Const RETVALFILE    = "ScrubRetValFile.txt"
Const OVERSION      = "15.0"
Const OVERSIONMAJOR = "15"
Const OREF          = "Office15"
Const OREGREF       = "OFFICE15."
Const ONAME         = "Office 2013 MSI"
Const OPACKAGE      = "PackageRefs"
Const OFFICEID      = "000000FF1CE}"
Const HKCR          = &H80000000
Const HKCU          = &H80000001
Const HKLM          = &H80000002
Const HKU           = &H80000003
Const FOR_WRITING   = 2
Const PRODLEN       = 12
Const COMPPERMANENT = "00000000000000000000000000000000"
Const UNCOMPRESSED  = 38
Const SQUISHED      = 20
Const COMPRESSED    = 32
Const REG_ARP       = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const VB_YES        = 6
Const MSIOPENDATABASEREADONLY = 0
Const LYNC_ALL      = "{4A2C120F-307B-4400-B239-F29ADB54D3C6}{5CFD6599-10E5-4CF0-B6E1-BF39D30A64F8}{5CFD6599-10E5-4CF0-B6E1-BF39D30A64F8}{BF3AC8BA-1A0F-42AD-8B65-4250617AF682}{3475BF22-3564-4EF3-A633-C5F3F4582392}{263BA91B-7782-4EEB-A4FC-7BD554CAF1F3}{AA256AE1-6B6A-48E6-9957-B38F92CA614B}{D79732A1-BB17-4789-AE75-69D61261E305}{C7B887F2-07CA-4903-93A2-9B4E16E4EABD}{81BE0B17-563B-45D4-B198-5721E6C665CD}{11298539-8073-4D54-B6A0-88D4FA512E5C}{C192041D-2861-4E02-9F43-4041858A58F1}{7023C711-0E65-471E-8048-12C455968841}{58A013B1-1613-4978-881A-FCA43710C84A}{7FD6C049-9777-4B51-91FF-B19D79ADF439}{D3001D99-675B-44DF-A8EB-A7BB6F864DB7}{0C5EA724-8649-47FA-B505-75B35390378D}{13DE0C92-2AE4-48D0-8CC8-58D5E327BDCB}{E7EC16E6-C220-41C0-9C91-5E7702B8EC86}{1B10C75C-70E1-460E-B07B-D7DFF365D80F}{331977BC-B246-46B4-8829-1D52F41C8C7B}{D8255EF2-0BB2-4AF1-A662-5EBACD179475}{DD069437-C92B-4C1C-A992-14F6C7E12C2C}{E9E30DB3-8D72-43A0-B1B8-A6F8261D20D6}{545B7E32-E254-40E1-8935-91C61E3D02C2}{70409E9E-AFAE-4C05-AE57-F83B89819434}{1D6E3225-753D-41AD-A2C4-68684700F592}{217AA75D-82C0-4C49-9252-A0E6F9661688}{5AB81CD4-7C78-420C-AAAC-855C4BADBDDD}{AA595672-6515-4961-B81F-485F86627C76}{C9F2C38C-21F0-4687-8C7D-51AA02CE8C98}{DD80DED6-700D-4CC5-B2A9-C64A1AD155B9}{88257193-EC61-4152-8AB1-A5FB4BE638D7}{7D9109C3-58A9-4AFD-A1D3-47E7D811726E}{71C6D199-5B8E-41E7-BA36-D99F66E0072E}{1CFE7869-777D-4563-8161-2C75ED95B621}{FE25DDB2-5766-4A9E-86D2-2B709CC8F65D}{621F7793-1C51-45BA-899F-41557946B0E3}{B31017AA-FBF8-4003-8785-EC789C2AE0C2}{11849FBC-C416-4742-8279-17C3A2C85F72}{4F380D4B-A84D-45C7-AF58-59EA2AEDF35A}{81BE0B17-563B-45D4-B198-5721E6C665CD}"

Const ERROR_SUCCESS                 = 0   'Bit #1.  0 indicates Success. Script completed successfully
Const ERROR_FAIL                    = 1   'Bit #1.  Failure bit. Indicates an overall script failure.
                                          'RESERVED bit! Returned when process is killed from task manager
Const ERROR_REBOOT_REQUIRED         = 2   'Bit #2.  Reboot bit. If set a reboot is required
Const ERROR_USERCANCEL              = 4   'Bit #3.  User Cancel bit. Controlled cancel from script UI
Const ERROR_STAGE1                  = 8   'Bit #4.  Informational. Msiexec based install was not possible
Const ERROR_STAGE2                  = 16  'Bit #5.  Critical. Not all of the intended cleanup operations could be applied
Const ERROR_INCOMPLETE              = 32  'Bit #6.  Pending file renames (del on reboot) - OR - Removal needs to run again after a system reboot.
Const ERROR_DCAF_FAILURE            = 64  'Bit #7.  Critical. Da capo al fine (second attempt) still failed.
Const ERROR_ELEVATION_USERDECLINED  = 128 'Bit #8.  Critical script error. User declined to allow mandatory script elevation
Const ERROR_ELEVATION               = 256 'Bit #9.  Critical script error. The attempt to elevate the process did not succeed
Const ERROR_SCRIPTINIT              = 512 'Bit #10. Critical script error. Initialization failed
Const ERROR_RELAUNCH                = 1024'Bit #11. Critical script error. This is a temporary value and must not be the final return code
Const ERROR_UNKNOWN                 = 2048'Bit #12 Critical script error. Script did not complete in a well defined state
Const ERROR_ALL                     = 4095'Full BitMask
Const ERROR_USER_ABORT              = &HC000013A 'RESERVED. Dec -1073741510. Critical error. Returned when user aborts with <Ctrl>+<Break> or closes the cmd window
Const ERROR_SUCCESS_CONFIG_COMPLETE = 1728
Const ERROR_SUCCESS_REBOOT_REQUIRED = 3010

'=======================================================================================================
Dim oFso, oMsi, oReg, oWShell, oWmiLocal, oShellApp
Dim ComputerItem, Item, LogStream, TmpKey
Dim arrTmpSKUs, arrDeleteFiles, arrDeleteFolders, arrMseFolders, arrVersion
Dim dicKeepProd, dicKeepLis, dicApps, dicKeepFolder, dicDelRegKey, dicKeepReg
Dim dicInstalledSku, dicRemoveSku, dicKeepSku, dicSrv, dicCSuite, dicCSingle
Dim f64, fLegacyProductFound, fCScript
Dim sTmp, sSkuRemoveList, sWinDir, sWICacheDir, sMode
Dim sAppData, sTemp, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles
Dim sAllusersProfile, sOSinfo, sOSVersion, sCommonProgramFilesX86, sProfilesDirectory
Dim sProgramData, sLocalAppData, sOInstallRoot, sScriptDir, sNotepad
Dim iVersionNT, iError

'=======================================================================================================
'Main
'=======================================================================================================
'Configure defaults
Dim sLogDir : sLogDir = ""
Dim sMoveMessage: sMoveMessage = ""
Dim fClearAddinReg	: fClearAddinReg = False
Dim fRemoveOse      : fRemoveOse = False
Dim fRemoveOspp     : fRemoveOspp = False
Dim fRemoveAll      : fRemoveAll = False
Dim fRemoveC2R      : fRemoveC2R = False
Dim fRemoveAppV     : fRemoveAppV = False
Dim fRemoveCSuites  : fRemoveCSuites = False
Dim fRemoveCSingle  : fRemoveCSingle = False
Dim fRemoveSrv      : fRemoveSrv = False
Dim fRemoveLync     : fRemoveLync = False
Dim fKeepUser       : fKeepUser = True  'Default to keep per user settings
Dim fSkipSD         : fSkipSD = False 'Default to not Skip the Shortcut Detection
Dim fKeepSG         : fKeepSG = False 'Default to not override the SoftGrid detection
Dim fDetectOnly     : fDetectOnly = False
Dim fQuiet          : fQuiet = False
Dim fBasic          : fBasic = False
Dim fNoCancel       : fNoCancel = False
Dim fPassive        : fPassive = True
Dim fNoReboot       : fNoReboot = False 'Default to offer reboot prompt if needed
Dim fNoElevate      : fNoElevate = False
Dim fElevated       : fElevated = False
Dim fTryReconcile   : fTryReconcile = False
Dim fC2rInstalled   : fC2rInstalled = False
Dim fRebootRequired : fRebootRequired = False
Dim fReturnErrorOrSuccess : fReturnErrorOrSuccess = False
Dim fEndCurrentInstalls : fEndCurrentInstalls = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fForce          : fForce = False
'CAUTION! -> "fForce" will kill running applications which can result in data loss! <- CAUTION
Dim fLogInitialized : fLogInitialized = False
Dim fBypass_Stage1  : fBypass_Stage1 = True 'Component Detection
Dim fBypass_Stage2  : fBypass_Stage2 = False 'Setup
Dim fBypass_Stage3  : fBypass_Stage3 = False 'Msiexec
Dim fBypass_Stage4  : fBypass_Stage4 = False 'CleanUp

'Create required objects
Set oWmiLocal   = GetObject("winmgmts:{(Debug)}\\.\root\cimv2")
Set oWShell     = CreateObject("Wscript.Shell")
Set oShellApp   = CreateObject("Shell.Application")
Set oFso        = CreateObject("Scripting.FileSystemObject")
Set oMsi        = CreateObject("WindowsInstaller.Installer")
Set oReg        = GetObject("winmgmts:\\.\root\default:StdRegProv")

'Get environment path info
sAppData            = oWShell.ExpandEnvironmentStrings("%appdata%")
sLocalAppData       = oWShell.ExpandEnvironmentStrings("%localappdata%")
sTemp               = oWShell.ExpandEnvironmentStrings("%temp%")
sAllUsersProfile    = oWShell.ExpandEnvironmentStrings("%allusersprofile%")
RegReadValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory", sProfilesDirectory, "REG_EXPAND_SZ"
If NOT oFso.FolderExists(sProfilesDirectory) Then 
    sProfilesDirectory  = oFso.GetParentFolderName(oWShell.ExpandEnvironmentStrings("%userprofile%"))
End If
sProgramFiles       = oWShell.ExpandEnvironmentStrings("%programfiles%")
'Deferred until after architecture check
'sProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%")

sCommonProgramFiles = oWShell.ExpandEnvironmentStrings("%commonprogramfiles%")
'Deferred until after architecture check
'sCommonProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%")

sProgramData        = oWSHell.ExpandEnvironmentStrings("%programdata%")
sWinDir             = oWShell.ExpandEnvironmentStrings("%windir%")
sWICacheDir         = sWinDir & "\" & "Installer"
sScrubDir           = sTemp & "\" & SCRIPTNAME
sNotepad            = sWinDir & "\notepad.exe"
'sScriptDir only required for OSPP cleanup
sScriptDir          = wscript.ScriptFullName
sScriptDir          = Left(sScriptDir, InStrRev(sScriptDir, "\"))

' Get current script host
fCScript = UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C"

'Detect if we're running on a 64 bit OS
Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
For Each Item In ComputerItem
    f64 = Instr(Left(Item.SystemType,3),"64") > 0
    If f64 Then Exit For
Next
If f64 Then sProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%programfiles(x86)%")
If f64 Then sCommonProgramFilesX86 = oWShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%")

'Get OS details and VersionNT
Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_OperatingSystem")
For Each Item in ComputerItem 
    sOSinfo = sOSinfo & Item.Caption 
    sOSinfo = sOSinfo & Item.OtherTypeDescription
    sOSinfo = sOSinfo & ", " & "SP " & Item.ServicePackMajorVersion
    sOSinfo = sOSinfo & ", " & "Version: " & Item.Version
    sOsVersion = Item.Version
    sOSinfo = sOSinfo & ", " & "Codepage: " & Item.CodeSet
    sOSinfo = sOSinfo & ", " & "Country Code: " & Item.CountryCode
    sOSinfo = sOSinfo & ", " & "Language: " & Item.OSLanguage
Next

'Build the VersionNT number
arrVersion = Split(sOsVersion, Delimiter(sOsVersion))
iVersionNt = CInt(arrVersion (0)) * 100 + CInt(arrVersion (1))

'Check if we're running as 32 bit process on a 64 bit OS
If InStr(LCase(wscript.path), "syswow64") > 0 Then RelaunchAs64Host

fElevated = CheckRegPermissions
If NOT fElevated AND NOT fNoElevate Then
    'Try to relaunch elevated
    RelaunchElevated

    ' can't relaunch. Exit out
    SetError ERROR_ELEVATION
    If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C" Then
        If Not fLogInitialized Then CreateLog
        Log "Error: Insufficient registry access permissions - exiting"
    End If
    SetRetVal iError
    'Undo temporary entries created in ARP
    TmpKeyCleanUp
    'wscript.quit 3
    ExitScript
End If

' set retval for file based logic
'--------------------------------
' value needs to be kept on 'user abort'
SetRetVal ERROR_USER_ABORT

' create dictionary objects
'--------------------------
Set dicKeepProd = CreateObject("Scripting.Dictionary")
Set dicInstalledSku = CreateObject("Scripting.Dictionary")
Set dicRemoveSku = CreateObject("Scripting.Dictionary")
Set dicKeepSku = CreateObject("Scripting.Dictionary")
Set dicKeepLis = CreateObject("Scripting.Dictionary")
Set dicKeepFolder = CreateObject("Scripting.Dictionary")
Set dicApps = CreateObject("Scripting.Dictionary")
Set dicDelRegKey = CreateObject("Scripting.Dictionary")
Set dicKeepReg = CreateObject("Scripting.Dictionary")
Set dicSrv = CreateObject("Scripting.Dictionary")
Set dicCSuite = CreateObject("Scripting.Dictionary")
Set dicCSingle = CreateObject("Scripting.Dictionary")

'Create the temp folder
If Not oFso.FolderExists(sScrubDir) Then oFso.CreateFolder sScrubDir

'Set the default logging directory
sLogDir = sScrubDir

'Call the command line parser
ParseCmdLine

'Ensure CScript as engine
If NOT fCScript AND NOT fQuiet Then RelaunchAsCScript

'Get Office Install Folder
If NOT RegReadValue(HKLM,"SOFTWARE\Microsoft\Office\"&OVERSION&"\Common\InstallRoot","Path",sOInstallRoot,"REG_SZ") Then 
    sOInstallRoot = sProgramFiles & "\Microsoft Office\"&OREF
End If

'Ensure integrity of WI metadata which could fail used APIs otherwise
EnsureValidWIMetadata HKCU,"Software\Classes\Installer\Products",COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Products",COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products",COMPRESSED
EnsureValidWIMetadata HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components",COMPRESSED
EnsureValidWIMetadata HKCR,"Installer\Components",COMPRESSED

'Add initial known .exe files that might need to be closed
dicApps.Add "communicator.exe", "communicator.exe"
'Adding setup.exe to the hard list of processes that are shut down will potentially break wrappers that invoke OffScrub
'dicApps.Add "setup.exe", "setup.exe"
Select Case OVERSIONMAJOR
Case "12"
Case "14"
    dicApps.Add "bcssync.exe","bcssync.exe"
    dicApps.Add "officesas.exe","officesas.exe"
    dicApps.Add "officesasscheduler.exe","officesasscheduler.exe"
    dicApps.Add "msosync.exe","msosync.exe"
    dicApps.Add "onenotem.exe","onenotem.exe"
Case "15"
Case Else
End Select

'-------------------
'Stage # 0 - Basics |
'-------------------
'Build a list with installed/registered Office products
sTmp = "Stage # 0 " & chr(34) & "Basics" & chr(34) & " (" & Time & ")"
LogH2 vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf

FindInstalledOProducts
If dicInstalledSku.Count > 0 Then Log "Found registered product(s): " & Join(RemoveDuplicates(dicInstalledSku.Items),",") 

'Validate the list of products we got from the command line if applicable
ValidateRemoveSkuList

'Log detection results
If dicRemoveSku.Count > 0 Then Log "Product(s) to be removed: " & Join(RemoveDuplicates(dicRemoveSku.Items),",")
sMode = "Selected " & ONAME & " products"
If NOT dicRemoveSku.Count > 0 Then sMode = "Orphaned " & ONAME & " products"
If fRemoveAll Then sMode = "All " & ONAME & " products"
Log "Final removal mode: " & sMode
Log "Remove OSE service: " & fRemoveOse 

'Log preview mode if applicable
If fDetectOnly Then Log "*************************************************************************"
If fDetectOnly Then Log "*                          PREVIEW MODE                                 *"
If fDetectOnly Then Log "* All uninstall and delete operations will only be logged not executed! *"
If fDetectOnly Then Log "*************************************************************************" & vbCrLf

'Check if there are legacy products installed
CheckForLegacyProducts
If fLegacyProductFound Then Log "Found legacy Office products that will not be removed." Else Log "No legacy Office products found."

'Cache .msi files
If dicRemoveSku.Count > 0 Then CacheMsiFiles

'Log Sku/Prod detection results
LogSkuResults

'UnPin Shortcuts
If NOT fSkipSD AND dicRemoveSku.Count > 0 Then
    On Error Resume Next
    LogH1 "UnPin shortcuts"
    CleanShortcuts sAllUsersProfile, False, True
    CleanShortcuts sProfilesDirectory, False, True
    On Error Goto 0
End If 'NOT SkipSD


'--------------------------------
'Stage # 1 - Component Detection |
'--------------------------------
sTmp = "Stage # 1 " & chr(34) & "Component Detection" & chr(34) & " (" & Time & ")"
LogH2 vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage1 OR fForce Then
    'Build a list with files which are installed/registered to a product that's going to be removed
    Log "Prepare for CleanUp stages."
    Log "Identifying removable elements. This can take several minutes."
    ScanComponents 
Else
    Log "Not running Component Detection in default removal."
End If

'End all running Office applications
If fForce OR fQuiet OR fPassive Then CloseOfficeApps

'----------------------
'Stage # 2 - Setup.exe |
'----------------------
sTmp = "Stage # 2 " & chr(34) & "Setup.exe" & chr(34) & " (" & Time & ")"
LogH2 vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage2 Then
    SetupExeRemoval
Else
    Log "Skipping Setup.exe because bypass was requested."
End If

'------------------------
'Stage # 3 - Msiexec.exe |
'------------------------
sTmp = "Stage # 3 " & chr(34) & "Msiexec.exe" & chr(34) & " (" & Time & ")"
LogH2 vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage3 Then
    MsiexecRemoval
Else
    Log "Skipping Msiexec.exe because bypass was requested."
End If

'--------------------
'Stage # 4 - CleanUp |
'--------------------
'Removal of files and registry settings
sTmp = "Stage # 4 " & chr(34) & "CleanUp" & chr(34) & " (" & Time & ")"
LogH2 vbCrLf & sTmp & vbCrLf & String(Len(sTmp),"=") & vbCrLf
If Not fBypass_Stage4 Then
    
    'Office Source Engine
    If fRemoveOse Then
        LogH1 "Office Source Engine CleanUp"
        RemoveOSE
    End If

    'Local Installation Source (MSOCache)
    LogH1 "Local Installation Source CleanUp"
    WipeLIS
    
    'Obsolete files
    LogH1 "File CleanUp"
    If fRemoveAll Then 
        FileWipeAll 
    Else 
        FileWipeIndividual
    End If
    
    'Empty Folders
    LogH1 "Folder CleanUp"
    DeleteEmptyFolders
    
    'Restore Explorer if needed
    If fForce OR fQuiet OR fPassive Then RestoreExplorer
    
    'Registry data
    LogH1 "Registry CleanUp"
    RegWipe
    
    'Wipe orphaned files from Windows Installer cache
    LogH1 "MSI Cache - orphaned files CleanUp"
    MsiClearOrphanedFiles
    
    'Temporary .msi files in scrubcache
    LogH1 "Temporary files CleanUp"
    DeleteMsiScrubCache
    
    'Temporary files
    DelScrubTmp
    
Else
    Log "Skipping CleanUp because bypass was requested."
End If

If Not sMoveMessage = "" Then Log vbCrLf & "Please remove this folder after next reboot: " & sMoveMessage


ExitScript

'-------------------------------------------------------------------------------
'   ExitScript
'
'   Returncode and reboot handler 
'-------------------------------------------------------------------------------
Sub ExitScript
    Dim sPrompt

    ' Update cached error and quit
    '-----------------------------
    SetRetVal iError
    Log vbCrLf & "For detailed logging please refer to the log in folder " & chr(34) & sScrubDir & chr(34) & vbCrLf

    ' log result
    If CBool(iError AND ERROR_INCOMPLETE) Then 
        LogH2 "Removal result: " & iError & " - INCOMPLETE. Uninstall requires a system reboot to complete."
    Else
        sTmp = " - SUCCESS"
        If CBool(iError AND ERROR_USERCANCEL) Then sTmp = " - USER CANCELED"
        If CBool(iError AND ERROR_FAIL) Then sTmp = " - FAIL"
        LogH2 "Removal result: " & iError & sTmp
    End If
    If CBool(iError AND ERROR_FAIL) Then
        If CBool(iError AND ERROR_REBOOT_REQUIRED) Then Log " - Reboot required"
        If CBool(iError AND ERROR_USERCANCEL) Then Log " - User cancel"
        If CBool(iError AND ERROR_STAGE1) Then Log " - Msiexec failed"
        If CBool(iError AND ERROR_STAGE2) Then Log " - Cleanup failed"
        If CBool(iError AND ERROR_INCOMPLETE) Then Log " - Removal incomplete. Rerun after reboot needed"
        If CBool(iError AND ERROR_DCAF_FAILURE) Then Log " - Second attempt cleanup still incomplete"
        If CBool(iError AND ERROR_ELEVATION_USERDECLINED) Then Log " - User declined elevation"
        If CBool(iError AND ERROR_ELEVATION) Then Log " - Elevation failed"
        If CBool(iError AND ERROR_SCRIPTINIT) Then Log " - Initialization error"
        If CBool(iError AND ERROR_RELAUNCH) Then Log " - Unhandled error during relaunch attempt"
        If CBool(iError AND ERROR_UNKNOWN) Then Log " - Unknown error"
        ' ERROR_USER_ABORT is only valid for the temporary cached error file
        'If CBool(iError AND ERROR_USER_ABORT) Then Log " - Process terminated by user"
    End If

    ' Check if we need to show a simplified return code
    ' 0 = Success
    ' Non Zero = Error
     If CBool(iError AND ERROR_FAIL) AND fReturnErrorOrSuccess Then
        Dim fOverallSuccess
        fOverallSuccess = True
        If CBool(iError AND ERROR_USERCANCEL) Then fOverallSuccess = False
        If CBool(iError AND ERROR_STAGE2) Then fOverallSuccess = False
        If CBool(iError AND ERROR_DCAF_FAILURE) Then fOverallSuccess = False
        If CBool(iError AND ERROR_ELEVATION_USERDECLINED) Then fOverallSuccess = False
        If CBool(iError AND ERROR_ELEVATION) Then fOverallSuccess = False
        If CBool(iError AND ERROR_SCRIPTINIT) Then fOverallSuccess = False
        If CBool(iError AND ERROR_RELAUNCH) Then fOverallSuccess = False
        If CBool(iError AND ERROR_UNKNOWN) Then fOverallSuccess = False

        sTmp = "ReturnErrorOrSuccess switch has been set. The current value return code translates to: "
        If fOverallSuccess Then 
            iError = ERROR_SUCCESS
            Log sTmp & "SUCCESS"
        Else
            Log sTmp & "ERROR"
        End If
    End If

    LogH2 "Removal end." 

    ' Reboot handling
    If fRebootRequired Then
        Log vbCrLf & "A restart is required to complete the operation!"
        sPrompt = "In order to complete uninstall, a system reboot is necessary. Would you like to reboot now?"
        If NOT (fQuiet OR fPassive OR fNoReboot) Then
            If MsgBox(sPrompt, vbYesNo, SCRIPTNAME & " - Reboot Required") = VB_YES Then
                Dim colOS, oOS
                Dim oWmiReboot
                Set oWmiReboot = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2")
                Set colOS = oWmiReboot.ExecQuery ("Select * from Win32_OperatingSystem")
                For Each oOS in colOS
                    oOS.Reboot()
                Next
            End If
        End If
    End If

    If NOT fQuiet Then
        For Each Item in Wscript.Arguments
            If Item = "UAC" Then 
                wscript.stdout.write "Press <Enter> to close this window"
                sTemp = wscript.stdin.read(1)
            End If
        Next 'Argument
    End If

    wscript.quit iError
End Sub 'ExitScript
'=======================================================================================================
'=======================================================================================================

'Stage 0 - 4 Subroutines
'=======================================================================================================

'Office configuration products are listed with their configuration product name in the "Uninstall" key
'To identify an Office configuration product all of these condiditions have to be met:
' - "SystemComponent" does not have a value of "1" (DWORD) 
' - "OPACKAGE" (see constant declaration) entry exists and is not empty
' - "DisplayVersion" exists and the 2 leftmost digits are "OVERSIONMAJOR"
Sub FindInstalledOProducts
    Dim ArpItem, File
    Dim sCurKey, sValue, sConfigName, sProdC, sCVHValue
    Dim sProductCodeList, sProductCode 
    Dim arrKeys, arrMultiSzValues
    Dim fSystemComponent0, fPackages, fDisplayVersion, fReturn, fCategorized

    If dicInstalledSku.Count > 0 Then Exit Sub 'Already done from InputBox prompt
    
    'Handle orphaned products to get them added to the detection scope
    If fTryReconcile Then
        For Each File in oFso.GetFolder(sWICacheDir).Files
            If Len(File.Name)>3 Then
                Select Case LCase(Right(File.Name, 4))
                Case ".msi"
                    sProductCode = ""
                    sProductCode = GetMsiProductCode(File.Path)
                    If InScope(sProductCode) Then
                        If NOT RegKeyExists(HKLM, REG_ARP & sProductCode) Then
                            'Ensure the orphaned item is getting removed
                            If Len(sSkuRemoveList) > 0 Then
                                sSkuRemoveList = sSkuRemoveList & "," & GetProductID(Mid(sProductCode, 11, 4))
                            Else
                                sSkuRemoveList = GetProductID(Mid(sProductCode, 11, 4))
                            End If
                            'Add to ScrubDir
                            oFso.CopyFile File.Path,sScrubDir & "\" & sProductCode & ".msi",True
                            'Register the product with MSI
                            MsiRegisterProduct File.Path
                        End If 'NOT sProductCode
                    End If 'InScope
                Case Else
                End Select
            End If '>3
        Next 'File
    End If 'fTryReconcile

    'Locate standalone Office products that have no configuration product entry and create a
    'temporary configuration entry
    ReDim arrTmpSKUs(-1)
    If RegEnumKey(HKLM, REG_ARP, arrKeys) Then
        For Each ArpItem in arrKeys
            If InScope(ArpItem) Then
                sCurKey = REG_ARP & ArpItem & "\"
                fSystemComponent0 = Not (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
                If fSystemComponent0 Then
                    RegReadValue HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ"
                    Redim arrMultiSzValues(0)
                    'Logic changed to drop the LCID identifier
                    'sConfigName = GetProductID(Mid(ArpItem,11,4)) & "_" & CInt("&h" & Mid(ArpItem,16,4))
                    sConfigName = OREGREF & GetProductID(Mid(ArpItem, 11, 4))
                    If NOT RegKeyExists(HKLM, REG_ARP&sConfigName) Then
                        'Create a new ARP item
                        ReDim Preserve arrTmpSKUs(UBound(arrTmpSKUs) + 1)
                        arrTmpSKUs(UBound(arrTmpSKUs)) = sConfigName
                        oReg.CreateKey HKLM, REG_ARP & sConfigName
                        arrMultiSzValues(0) = sConfigName
                        oReg.SetMultiStringValue HKLM, REG_ARP & sConfigName, OPACKAGE, arrMultiSzValues
                        arrMultiSzValues(0) = ArpItem
                        oReg.SetStringValue HKLM, REG_ARP & sConfigName, "Comment", "Temporary OffScrub generated key. Please delete this key!"
                        oReg.SetMultiStringValue HKLM, REG_ARP & sConfigName, "ProductCodes", arrMultiSzValues
                        oReg.SetStringValue HKLM, REG_ARP & sConfigName, "DisplayVersion", sValue
                        oReg.SetStringValue HKLM, REG_ARP & sConfigName, "DisplayName", SCRIPTNAME & "_" & sConfigName
                        oReg.SetDWordValue HKLM, REG_ARP & sConfigName, "SystemComponent", 0
                    Else
                        'Update the existing temporary ARP item
                        fReturn = RegReadValue(HKLM, REG_ARP&sConfigName, "ProductCodes", sProdC, "REG_MULTI_SZ")
                        If NOT InStr(sProdC, ArpItem) > 0 Then sProdC = sProdC & chr(34) & ArpItem
                        oReg.SetMultiStringValue HKLM, REG_ARP & sConfigName, "ProductCodes", Split(sProdC, Chr(34))
                    End If 'RegKeyExists
                End If 'fSystemComponent0
            End If 'InScope
        Next 'ArpItem
    End If 'RegEnumKey
    
    'Find the configuration products
    If RegEnumKey(HKLM, REG_ARP, arrKeys) Then
        For Each ArpItem in arrKeys
            sCurKey = REG_ARP & ArpItem & "\"
            sValue = ""
            fSystemComponent0 = NOT (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
            fPackages = RegReadValue(HKLM, sCurKey, OPACKAGE, sValue, "REG_MULTI_SZ")
            fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ")
            If fDisplayVersion Then
                If Len(sValue) > 1 Then
                    fDisplayVersion = (Left(sValue, 2) = OVERSIONMAJOR)
                Else
                    fDisplayVersion = False
                End If
            End If
            'fSystemComponent0 filter causes issues if the ARP entries have been hidden 
            'If (fSystemComponent0 AND fPackages AND fDisplayVersion) Then
            If (fPackages AND fDisplayVersion) Then
                If InStr(ArpItem,".") > 0 Then sConfigName = UCase(Mid(ArpItem, InStr(ArpItem, ".") + 1)) Else sConfigName = UCase(ArpItem)
                If NOT dicInstalledSku.Exists(sConfigName) Then dicInstalledSku.Add sConfigName, sConfigName

                'Categorize the SKU
                'Four categories are available: ClientSuite, ClientSingleProduct, Server, C2R
                If RegReadValue(HKLM, REG_ARP & OREGREF & sConfigName, "ProductCodes", sProductCodeList, "REG_MULTI_SZ") Then
                    fCategorized = False
                    For Each sProductCode in Split(sProductCodeList, Chr(34))
                        If Len(sProductCode) = 38 Then
                            If Mid(sProductCode, 11, 1) = "1" Then
                                'Server product
                                If NOT dicSrv.Exists(UCase(sConfigName)) Then dicSrv.Add UCase(sConfigName), sConfigName
                                fCategorized = True
                                Exit For
                            Else
                                Select Case Mid(sProductCode, 11, 4)
                                'Client Suites
                                Case "000F","0011","0012","0013","0014","0015","0016","0017","0018","0019","001A","001B","0029","002B","002E","002F","0030","0031","0033","0035","0037","003D","0044","0049","0061","0062","0066","006C","006D","006F","0074","00A1","00A3","00A9","00BA","00CA","00E0","0100","0103","011A"
                                    If NOT dicCSuite.Exists(UCase(sConfigName)) Then dicCSuite.Add UCase(sConfigName), sConfigName
                                    fCategorized = True
                                    Exit For

                                Case "007E", "008F", "008C", "24E1", "237A"
                                    If NOT dicKeepProd.Exists(sProductCode) Then dicKeepProd.Add sProductCode, sConfigName 
                                    fC2rInstalled = True
                                Case Else
                                End Select
                            End If
                        End If 'Len 38
                    Next 'sProductCode
                    If NOT fCategorized Then
                        If NOT dicCSingle.Exists(UCase(sConfigName)) Then dicCSingle.Add UCase(sConfigName), sConfigName
                    End If 'fCategorized
                End If 'RegReadValue "ProductCodes"
            End If
        Next 'ArpItem
    End If 'RegEnumKey
End Sub 'FindInstalledOProducts
'=======================================================================================================

'Check if there are Office products from previous versions on the computer
Sub CheckForLegacyProducts
    Const OLEGACY = "78E1-11D2-B60F-006097C998E7}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}"
    Dim Product
    
    'Set safe default
    fLegacyProductFound = True
    
    For Each Product in oMsi.Products
        If Len(Product) = 38 Then
            'Handle O09 - O11 Products
            If InStr(OLEGACY, UCase(Right(Product, 28))) > 0 Then
                'Found legacy Office product. Keep flag in default and exit
                Exit Sub
            End If
            If UCase(Right(Product,PRODLEN)) = OFFICEID Then
                Select Case Mid(Product,4,2)
                Case "12", "14"
                    'Found legacy Office product. Keep flag in default and exit
                    Exit Sub
                Case Else
                End Select
            End If
        End If '38
    Next 'Product
    fLegacyProductFound = False
    
End Sub 'CheckForLegacyProducts
'=======================================================================================================

'Create clean list of Products to remove.
'Strip off bad & empty contents
Sub ValidateRemoveSkuList
    Dim Sku, Key, sProductCode, sProductCodeList
    Dim arrRemoveSKUs
    
    If fRemoveAll Then
        'Remove all mode
        For Each Key in dicInstalledSku.Keys
            dicRemoveSku.Add Key,dicInstalledSku.Item(Key)
        Next 'Key
    Else
        'Remove individual products or preconfigured configurations mode
        
        'Ensure to have a string with no unexpected contents
        sSkuRemoveList = Replace(sSkuRemoveList,";",",")
        sSkuRemoveList = Replace(sSkuRemoveList," ","")
        sSkuRemoveList = Replace(sSkuRemoveList,Chr(34),"")
        While InStr(sSkuRemoveList,",,")>0
            sSkuRemoveList = Replace(sSkuRemoveList,",,",",")
        Wend
        
        'Prepare 'remove' and 'keep' dictionaries to determine what has to be removed
        
        'Initial pre-fill of 'keep' dic
        For Each Key in dicInstalledSku.Keys
            dicKeepSku.Add Key,dicInstalledSku.Item(Key)
        Next 'Key
        
        'Determine contents of keep and remove dic
        'Individual products
        arrRemoveSKUs = Split(UCase(sSkuRemoveList),",")
        For Each Sku in arrRemoveSKUs
            If Sku = "OSE" Then fRemoveOse = True
            If dicKeepSku.Exists(Sku) Then
                'A Sku to remove has been passed in
                'remove the item from the keep dic
                dicKeepSku.Remove(Sku)
                'Now add it to the remove dic
                If NOT dicRemoveSku.Exists(Sku) Then dicRemoveSku.Add Sku,Sku
            End If
        Next 'Sku

        'Client Suite Category
        If fRemoveCSuites Then
            For Each Key in dicInstalledSku.Keys
                If dicCSuite.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveCSuites
        
        'Client Single/Standalone Category
        If fRemoveCSingle Then
            For Each Key in dicInstalledSku.Keys
                If dicCSingle.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveCSingle
        
        'Server Category
        If fRemoveSrv Then
            For Each Key in dicInstalledSku.Keys
                If dicSrv.Exists(Key) Then
                    If dicKeepSku.Exists(Key) Then dicKeepSku.Remove(Key)
                    If NOT dicRemoveSku.Exists(Key) Then dicRemoveSku.Add Key,Key
                End If
            Next 'Key
        End If 'fRemoveSrv
        
        If NOT dicKeepSku.Count > 0 Then fRemoveAll = True

    End If 'fRemoveAll

    'Fill the KeepProd dic
    For Each Sku in dicKeepSku.Keys
        If RegReadValue(HKLM, REG_ARP & OREGREF & Sku, "ProductCodes", sProductCodeList, "REG_MULTI_SZ") Then
            For Each sProductCode in Split(sProductCodeList, chr(34))
                If Len(sProductCode) = 38 Then
                    If NOT dicKeepProd.Exists(sProductCode) Then
                    	dicKeepProd.Add sProductCode, Sku
                    	' also add the UpgradeCode 
                    	If Not dicKeepProd.Exists(GetUpgradeCode(sProductCode)) Then dicKeepProd.Add GetUpgradeCode(sProductCode), Sku & "_UpgradeCode"
                    End If
                End If '38
            Next 'sProductCode 
        End If
    Next 'Sku
        
    If fRemoveAll OR fRemoveOse Then CheckRemoveOSE
    If fRemoveAll OR fRemoveOspp Then CheckRemoveOspp

End Sub 'ValidateRemoveSkuList
'=======================================================================================================

'Check if OSE service can be scrubbed
Sub CheckRemoveOSE
    Const O11 = "6000-11D3-8CFE-0150048383C9}"
    Dim Product
    
    If fRemoveOse Then Exit Sub
    For Each Product in oMsi.Products
        If Len(Product) = 38 Then
            If UCase(Right(Product, 28)) = O11 Then 
                'Found Office 2003 Product. Set flag to not remove the OSE service
                Exit Sub
            End If
            If UCase(Right(Product, PRODLEN))= OFFICEID Then
                Select Case Mid(Product, 4, 2)
                Case "12","14","15","16","17"
                    If NOT Mid(Product, 4, 2) = OVERSIONMAJOR Then
	                    'Found another Office product. Set flag to keep the OSE service
                        fRemoveOse = False
                        Exit Sub
                    End If
                Case Else
                End Select
            End If
        End If '38
    Next 'Product
    fRemoveOse = True
End Sub 'CheckRemoveOSE
'=======================================================================================================

'Check if OSPP service can be scrubbed
Sub CheckRemoveOSPP
    Dim Product
    
    If NOT CInt(OVERSIONMAJOR) > 12 Then 
        fRemoveOspp = False
        Exit Sub
    End If

    If fRemoveOspp Then Exit Sub

    If fC2rInstalled Then
        fRemoveOspp = False
        Exit Sub
    End If

    For Each Product in oMsi.Products
        If Len(Product) = 38 Then
            If UCase(Right(Product, PRODLEN)) = OFFICEID Then
                Select Case Mid(Product, 4, 2)
                Case "14","15","16","17"
                    If NOT Mid(Product, 4, 2) = OVERSIONMAJOR Then
						'Found another Office product. Set flag to keep the OSPP service
                        fRemoveOspp = False
                        Exit Sub
                    End If
                Case Else
                End Select
            End If
        End If '38
    Next 'Product
    fRemoveOspp = True
End Sub 'CheckRemoveOSPP
'=======================================================================================================

'Cache .msi files for products that will be removed in case they are needed for later file detection
Sub CacheMsiFiles
    Dim Product
    Dim sMsiFile
    
    'Non critical routine for failures.
    'Errors will be logged but must not fail the execution
    On Error Resume Next
    LogH1 "Cache .msi files to temporary Scrub folder"
    'Cache the files
    For Each Product in oMsi.Products
        'Ensure valid GUID length
        If CheckDeleteEx(Product) Then
            CheckError "CacheMsiFiles"
            sMsiFile = oMsi.ProductInfo(Product,"LocalPackage") : CheckError "CacheMsiFiles"
            LogOnly " - " & Product & ".msi"
            If oFso.FileExists(sMsiFile) Then oFso.CopyFile sMsiFile,sScrubDir & "\" & Product & ".msi",True
            CheckError "CacheMsiFiles"
        End If 'InScope
    Next 'Product

    Err.Clear
End Sub 'CacheMsiFiles
'=======================================================================================================

'Build a list of all files that will be deleted
Sub ScanComponents
    Const MSIINSTALLSTATE_LOCAL = 3

    Dim FileList, RegList, ComponentID, CompClient, Record, qView, MsiDb, CompVerbose
    Dim Processes, Process, Prop, prod
    Dim sQuery, sSubKeyName, sPath, sFile, sMsiFile, sCompClient, sComponent, sCompReg
    Dim fRemoveComponent, fAffectedComponent, fIsPermanent, fIsFile, fIsFolder
    Dim i, iProgress, iCompCnt, iRemCnt
    Dim dicFLError, oDic, oFolderDic, dicCompPath
    Dim hDefKey

    'Logfile
    Set FileList = oFso.OpenTextFile(sScrubDir & "\FileList.txt",FOR_WRITING,True,True)
    Set RegList = oFso.OpenTextFile(sScrubDir & "\RegList.txt",FOR_WRITING,True,True)
    Set CompVerbose = oFso.OpenTextFile(sScrubDir & "\CompVerbose.txt",FOR_WRITING,True,True)
    
    'FileListError dic
    Set dicFLError = CreateObject("Scripting.Dictionary")
    
    Set oDic = CreateObject("Scripting.Dictionary")
    Set oFolderDic = CreateObject("Scripting.Dictionary")
    Set dicCompPath = CreateObject("Scripting.Dictionary")

    'Prevent that API errors fail script execution
    On Error Resume Next

    iCompCnt = oMsi.Components.Count
    If NOT Err = 0 Then
        'API failure
        Log "Error during components detection. Cannot complete this task."
        Err.Clear
        Exit Sub
    End If

    'Ensure to not divide by zero
    If iCompCnt = 0 Then iCompCnt = 1
    LogOnly " Scanning " & iCompCnt & " components"
    'Enum all Components
    For Each ComponentID In oMsi.Components
        CompVerbose.WriteLine vbCrLf & "Checking Component: " & ComponentID
        
        'Progress bar
        i = i + 1
        If iProgress < (i / iCompCnt) * 100 Then 
            If fCScript Then wscript.stdout.write "." : LogStream.Write "."
            iProgress = iProgress + 1
            If iProgress = 35 OR iProgress = 70 Then Log ""
        End If

        'Check if all ComponentClients will be removed
        sCompClient = ""
        iRemCnt = 0
        fIsPermanent = False
        fRemoveComponent = False 'Flag to track if the component will be completely removed
        fAffectedComponent = False 'Flag to track if some clients remain installed who have a none shared location
        dicCompPath.RemoveAll
        Err.Clear
        For Each CompClient In oMsi.ComponentClients(ComponentID)
            CompVerbose.Write " CompClient " & CompClient & "-> "
            If Err = 0 Then
                'Ensure valid guid length
                If Len(CompClient) = 38 Then
                    fRemoveComponent = InScope(CompClient)
                    If fRemoveComponent OR (CompClient = "{00000000-0000-0000-0000-000000000000}") Then
                        sPath = ""
                        sPath = LCase(oMsi.ComponentPath(CompClient,ComponentID))
                        sPath = Replace(sPath,"?",":")
                        'Scan for msidbComponentAttributesPermanent flag
                        If CompClient = "{00000000-0000-0000-0000-000000000000}" Then
                            fIsPermanent = True
                            iRemCnt = iRemCnt + 1
                        End If
                        If fRemoveComponent Then fRemoveComponent = CheckDelete(CompClient)
                        CompVerbose.Write "CheckDelete: " & fRemoveComponent & "; "
                        If fRemoveComponent Then
                            iRemCnt = iRemCnt + 1
                            fAffectedComponent = True
                            'Since the scope remains within one Office family the keypath for the component
                            'is assumed to be identical
                            If sCompClient = "" Then sCompClient = CompClient
                        ' flag the CompClient entry for removal
                            sCompReg = "Installer\Components\"&GetCompressedGuid(ComponentID)&"\"&GetCompressedGuid(CompClient)
                            If NOT dicDelRegKey.Exists(sCompReg) Then
                                dicDelRegKey.Add sCompReg,HKCR
                                RegList.WriteLine HiveString(HKCR)&"\"&sCompReg
                            End If
                            sCompReg = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"&GetCompressedGuid(ComponentID)&"\"&GetCompressedGuid(CompClient)
                            If NOT dicDelRegKey.Exists(sCompReg) Then
                                dicDelRegKey.Add sCompReg,HKLM
                                RegList.WriteLine HiveString(HKCR)&"\"&sCompReg
                            End If
                        Else
                            If NOT dicCompPath.Exists(sPath) Then dicCompPath.Add sPath,CompClient
                        End If
                        CompVerbose.WriteLine "AffectedComponent: " & fAffectedComponent
                        CompVerbose.WriteLine " CompClient now set to: " & sCompClient
                    Else
                        CompVerbose.Write "InScope: " & fRemoveComponent & "; "
                    End If
                Else
                    CompVerbose.WriteLine "Error: Invalid metadata"
                    If NOT dicFLError.Exists("Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient) Then _
                        dicFLError.Add "Error: Invalid metadata found. ComponentID: "&ComponentID &", ComponentClient: "&CompClient, ComponentID
                End If '38
            Else
                CompVerbose.WriteLine "Error: " & Err.number & " " & Err.Description
                Err.Clear
            End If 'Err = 0
        Next 'CompClient
        
        'Determine if the component resources go away
        sPath = ""
        fRemoveComponent = fAffectedComponent AND (iRemCnt = oMsi.ComponentClients(ComponentID).Count)
        CompVerbose.WriteLine " Component goes away: " & fRemoveComponent
' This caused unintentional removals
'        If NOT fRemoveComponent AND fAffectedComponent Then
'            'Flag as removable if component has a unique keypath
'            sPath = LCase(oMsi.ComponentPath(sCompClient,ComponentID))
'            sPath = Replace(sPath,"?",":")
'            fRemoveComponent = NOT dicCompPath.Exists(sPath)
'        End If
        If fRemoveComponent Then
            'Check msidbComponentAttributesPermanent flag
            If fIsPermanent AND NOT fForce Then fRemoveComponent = False
            CompVerbose.WriteLine " msidbComponentAttributesPermanent: " & NOT fRemoveComponent
        End If

        If fRemoveComponent Then
            CompVerbose.WriteLine " RESULT: Component IN SCOPE for removal"
            fIsFile = False : fIsFolder = False

            'Component resources go away for this product
            Err.Clear
            'Add the component registration key to ensure removal
            sCompReg = "Installer\Components\"&GetCompressedGuid(ComponentID)&"\"
            If NOT dicDelRegKey.Exists(sCompReg) Then
                dicDelRegKey.Add sCompReg,HKCR
                RegList.WriteLine HiveString(HKCR)&"\"&sCompReg
            End If
            sCompReg = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"&GetCompressedGuid(ComponentID)&"\"
            If NOT dicDelRegKey.Exists(sCompReg) Then
                dicDelRegKey.Add sCompReg,HKLM
                RegList.WriteLine HiveString(HKCR)&"\"&sCompReg
            End If
            'Get the component path
            If sPath = "" Then
                sPath = LCase(oMsi.ComponentPath(sCompClient,ComponentID))
                sPath = Replace(sPath,"?",":")
            End If
            CompVerbose.WriteLine " Path: " & sPath
            If Len(sPath) > 4 Then
                If Left(sPath,1) = "0" Then
                    'Registry keypath

                    Select Case Left(sPath,2)
                    Case "00"
                        sPath = Mid(sPath,5)
                        hDefKey = HKCR
                    Case "01"
                        sPath = Mid(sPath,5)
                        hDefKey = HKCU
                    Case "02","22"
                        sPath = Mid(sPath,5)
                        hDefKey = HKLM
                    Case Else
                        '
                    End Select
                    
                    'Go for the safe way and just reset the default entry
                    'compared to deleting the whole key
                    If Right(sPath,1) = "\" Then sPath = sPath & "(Default)"

                    If NOT dicDelRegKey.Exists(sPath) Then
                        dicDelRegKey.Add sPath,hDefKey
                        RegList.WriteLine HiveString(hDefKey)&"\"&sPath
                    End If
                Else
                
                    'File or Folder
                    If oFso.FileExists(sPath) OR oFso.FolderExists(sPath) Then
                        If Right(sPath,1) = "\" Then
                            fIsFolder = True
                            CompVerbose.WriteLine " Folder check OK"
                        Else
                            fIsFile = True
                            CompVerbose.WriteLine " File check OK"
                        End If
                        If fIsFile Then sPath = oFso.GetFile(sPath).ParentFolder
                        If Not oFolderDic.Exists(sPath) Then
                            oFolderDic.Add sPath,sPath
                            FileList.WriteLine sPath & vbTab & "(FOLDER)"
                        End If
                        'Get the .msi file
                        If oFso.FileExists(sScrubDir & "\" & sCompClient & ".msi") Then
                            sMsiFile = sScrubDir & "\" & sCompClient & ".msi"
                        Else
                            sMsiFile = oMsi.ProductInfo(sCompClient,"LocalPackage")
                        End If
                        CompVerbose.WriteLine " Set msi file to : " & sMsiFile
                        If Not Err = 0 Then
                            CompVerbose.WriteLine " Error: Failed to obtain .msi file for product " & sCompClient
                            If NOT dicFLError.Exists("Failed to obtain .msi file for product "&sCompClient) Then _
                                dicFLError.Add "Failed to obtain .msi file for product "&sCompClient, ComponentID
                            Err.Clear
                        End If
                        CompVerbose.Write " Open .msi file for reading returned: "
                        Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEREADONLY)
                        
                        If Err = 0 Then
                            CompVerbose.WriteLine " SUCCESS"
                            'Get the component name from the 'Component' table
                            sQuery = "SELECT `Component`,`ComponentId` FROM Component WHERE `ComponentId` = '" & ComponentID &"'"
                            Set qView = MsiDb.OpenView(sQuery) : qView.Execute
                            Set Record = qView.Fetch()
                            If Not Record Is Nothing Then sComponent = Record.Stringdata(1)
                            CompVerbose.WriteLine " Obtained ComponentId as: " & sComponent

                            'Get filenames from the 'File' table
                            sQuery = "SELECT `Component_`,`FileName` FROM File WHERE `Component_` = '" & sComponent &"'"
                            Set qView = MsiDb.OpenView(sQuery) : qView.Execute
                            Set Record = qView.Fetch()
                            Do Until Record Is Nothing
                                'Read the filename
                                sFile = Record.StringData(2)
                                If InStr(sFile,"|") > 0 Then sFile = Mid(sFile,InStr(sFile,"|")+1,Len(sFile))
                                'sFile = sPath & "\" & sFile
                                CompVerbose.WriteLine "  File: " & sPath& "\" & sFile
                                If Not oDic.Exists(sPath & "\" & sFile) Then 
                                    'Exception handler
                                    fAdd = True
                                    Select Case UCase(sFile)
                                    Case "FPERSON.DLL"
                                        'Catch exception caused by changed .msi keypath authoring logic for smart tags
                                        For Each prod in oMsi.Products
                                            If NOT Checkdelete(prod) Then
                                                If oMsi.FeatureState(prod, "MSTagPluginNamesFiles") = MSIINSTALLSTATE_LOCAL Then
                                                    fAdd = False
                                                    Exit For
                                                End If
                                            End If
                                        Next 'prod
                                    Case Else
                                    End Select
                                    If fAdd Then
                                        CompVerbose.WriteLine "  Added as new file to dictionary"
                                        oDic.Add sPath & "\" & sFile,sFile
                                        FileList.WriteLine sFile & vbTab & sPath & "\" & sFile
                                        If Len(sFile)>4 Then
                                            sFile = LCase(sFile)
                                            If Right(sFile,4) = ".exe" Then
                                                If NOT dicApps.Exists(sFile) Then
                                                    Select Case sFile
                                                    Case "setup.exe","ose.exe","osppsvc.exe","explorer.exe"
                                                    Case Else
                                                        dicApps.Add sFile,LCase(sPath) & "\" & sFile
                                                        CompVerbose.WriteLine "  Added to the list of processes that need to be closed."
                                                    End Select
                                                End If 'dicApps.Exists
                                            End If '.exe
                                        End If 'Len > 4
                                    End If 'fAdd
                                End If 'oDic.Exists
                                Set Record = qView.Fetch()
                            Loop
                            Set Record = Nothing
                            qView.Close
                            Set qView = Nothing
                        Else
                            CompVerbose.WriteLine " Error: Could not read from .msi file"
                            If NOT dicFLError.Exists("Error: Could not read from .msi file: "&sMsiFile) Then _
                                dicFLError.Add "Error: Could not read from .msi file: "&sMsiFile, ComponentID
                            Err.Clear
                        End If 'Err = 0
                    Else
                        CompVerbose.WriteLine " Error: File check FAILED"
                    End If 'FileExists(sPath)
                End If
            End If 'Len(sPath) > 4
        Else
            CompVerbose.WriteLine " RESULT: Component NOT in scope for removal"
            If fAffectedComponent Then
                'Add the path to the 'Keep' dictionary
                Err.Clear
                For Each CompClient In oMsi.ComponentClients(ComponentID)
                    'Get the component path
                    sPath = "" : sPath = LCase(oMsi.ComponentPath(CompClient,ComponentID))
                    sPath = Replace(sPath,"?",":")
                
                    If Len(sPath) > 4 Then
                        If Left(sPath,1) = "0" Then
                            'Registry keypath

                            Select Case Left(sPath,2)
                            Case "00"
                                sPath = Mid(sPath,5)
                                hDefKey = HKCR
                            Case "01"
                                sPath = Mid(sPath,5)
                                hDefKey = HKCU
                            Case "02","22"
                                sPath = Mid(sPath,5)
                                hDefKey = HKLM
                            Case Else
                                '
                            End Select
                            If NOT dicKeepReg.Exists(LCase(sPath)) Then
                                dicKeepReg.Add LCase(sPath),hDefKey
                            End If
                        Else
                            'File keypath
                            If oFso.FileExists(sPath) Then
                                If NOT dicKeepFolder.Exists(LCase(sPath)) Then dicKeepFolder.Add LCase(sPath)
                                sPath = LCase(oFso.GetFile(sPath).ParentFolder) & "\"
                                If NOT dicKeepFolder.Exists(sPath) Then AddKeepFolder sPath
                            End If
                            'Folder keypath
                            If oFso.FolderExists(sPath) Then AddKeepFolder sPath
                        End If 'Is Registry
                    End If 'sPath > 4
                Next 'CompClient
            End If 'fAffectedComponent
        End If 'fRemoveComponent
        Err.Clear
    Next 'ComponentID
    On Error Goto 0
    
    Log " Done" & vbCrLf
    If dicFLError.Count > 0 Then LogOnly Join(dicFLError.Keys,vbCrLf)
    If Not oFolderDic.Count = 0 Then arrDeleteFolders = oFolderDic.Keys Else Set arrDeleteFolders = Nothing
    If Not oDic.Count = 0 Then arrDeleteFiles = oDic.Keys Else Set arrDeleteFiles = Nothing
End Sub 'ScanComponents
'=======================================================================================================


'Try to remove the products by calling setup.exe
Sub SetupExeRemoval
    Dim OseService, Service, TextStream
    Dim iSetupCnt, RetVal
    Dim Sku, sConfigFile, sUninstallCmd, sCatalyst, sDll, sDisplayLevel, sNoCancel

    iSetupCnt = 0
    If Not dicRemoveSku.Count > 0 Then
        Log " Nothing to remove for Setup.exe"
        Exit Sub
    End If
    
    'Ensure that the OSE service is *installed, *not disabled, *running under System context.
    'If validation fails exit out of this sub.
    Set OseService = oWmiLocal.Execquery("Select * From Win32_Service Where Name like 'ose%'")
    If OseService.Count = 0 Then Exit Sub
    For Each Service in OseService
        If (Service.StartMode = "Disabled") AND (Not Service.ChangeStartMode("Manual")=0) Then Exit Sub
        If (Not Service.StartName = "LocalSystem") AND (Service.Change( , , , , , , "LocalSystem", "")) Then Exit Sub
    Next 'Service
    
    For Each Sku in dicRemoveSku.Keys
        If Sku="CLICK2RUN" Then
            'Already done
        Else
            'Create an "unattended" config.xml file for uninstall
            If fQuiet Then sDisplayLevel = "None" Else sDisplayLevel="Basic"
            If fNoCancel Then sNoCancel="Yes" Else sNoCancel="No"
            Set TextStream = oFso.OpenTextFile(sScrubDir & "\config.xml",FOR_WRITING,True,True)
            TextStream.Writeline "<Configuration Product=""" & Sku & """>"
            TextStream.Writeline "<Display Level=""" & sDisplayLevel & """ CompletionNotice=""No"" SuppressModal=""Yes"" NoCancel=""" & sNoCancel & """ AcceptEula=""Yes"" />"
            TextStream.Writeline "<Logging Type=""Verbose"" Path=""" & sLogDir & """ Template=""Microsoft Office " & Sku & " Setup(*).txt"" />"
            TextStream.Writeline "<Setting Id=""MSIRESTARTMANAGERCONTROL"" Value=""Disable"" />"
            TextStream.Writeline "<Setting Id=""SETUP_REBOOT"" Value=""Never"" />"
            TextStream.Writeline "</Configuration>"
            TextStream.Close
            Set TextStream = Nothing
        
            'Ensure path to setup.exe is valid to prevent errors
            sDll = ""
            If RegReadValue(HKLM,REG_ARP & OREGREF & Sku,"UninstallString",sCatalyst,"REG_SZ") Then
                If InStr(LCase(sCatalyst),"/dll")>0 Then sDll = Right(sCatalyst,Len(sCatalyst)-InStr(LCase(sCatalyst),"/dll")+2)
                If InStr(sCatalyst,"/")>0 Then sCatalyst = Left(sCatalyst,InStr(sCatalyst,"/")-1)
                sCatalyst = Trim(Replace(sCatalyst,Chr(34),""))
                If NOT oFso.FileExists(sCatalyst) Then
                    sCatalyst = sCommonProgramFiles & "\" & OREF & "\Office Setup Controller\setup.exe"
                    If NOT oFso.FileExists(sCatalyst) AND f64 Then
                        sCatalyst = sCommonProgramFilesX86 & "" & OREF & "\Office Setup Controller\setup.exe"
                    End If
                End If
                If oFso.FileExists(sCatalyst) Then
                    sUninstallCmd = Chr(34) & sCatalyst & Chr(34) & " /uninstall " & Sku & " /config " & Chr(34) & sScrubDir & "\config.xml" & Chr(34) & sDll 
                    iSetupCnt = iSetupCnt + 1
                    Log " - Run Setup.exe to remove " & Sku '& vbCrLf & sUninstallCmd 
                    If Not fDetectOnly Then 
                        On Error Resume Next
                        ' end other instances of setup
                        EndCurrentInstalls
                        ' call uninstall
                        RetVal = oWShell.Run(sUninstallCmd,0,True) : CheckError "SetupExeRemoval"
                        Log " - Setup.exe returned: " & SetupRetVal(Retval) & " (" & RetVal & ")" & vbCrLf
                        fRebootRequired = fRebootRequired OR (RetVal = "3010")
                        On Error Goto 0
                    Else
                        Log " -> Removal suppressed in preview mode."
                    End If
                Else
                    Log " Error: Office setup.exe appears to be missing"
                End If 'RetVal = 0) AND oFso.FileExists
            End If 'RegReadValue
        End If 
    Next 'Sku
    If iSetupCnt = 0 Then Log " Nothing to remove for setup."
End Sub 'SetupExeRemoval
'=======================================================================================================

'Invoke msiexec to remove individual .MSI packages
Sub MsiexecRemoval

    Dim Product
    Dim i
    Dim sCmd, sReturn, sMsiProp
    Dim fRegWipe

    fRegWipe = False

    Select Case OVERSIONMAJOR
    Case "11"
        sMsiProp = " REBOOT=ReallySuppress NOLOCALCACHEROLLBACK=1"
    Case "12"
        fRegWipe = True
        sMsiProp = " REBOOT=ReallySuppress NOREMOVESPAWN=True"
    Case "14"
        fRegWipe = True
        sMsiProp = " REBOOT=ReallySuppress NOREMOVESPAWN=True"
    Case "15"
        fRegWipe = True
        sMsiProp = " REBOOT=ReallySuppress NOREMOVESPAWN=True"
    Case Else
    End Select

    'Clear up ARP first to avoid possible custom action dependencies
    If fRegWipe Then RegWipeARP

    'Check MSI registered products
    'Office System does only support per machine installation so it's sufficient to use Installer.Products
    i = 0
    sMsiProp = " MSIRESTARTMANAGERCONTROL=Disable" & sMsiProp
    For Each Product in oMsi.Products
        If CheckDeleteEx(Product) Then
            i = i + 1 
            Log " Calling msiexec.exe to remove " & Product
            sCmd = "msiexec.exe /x" & Product & sMsiProp
            If fQuiet AND NOT fBasic Then 
                sCmd = sCmd & " /q"
            Else
                sCmd = sCmd & " /qb-"
            End If
            sCmd = sCmd & " /l*v+ "&chr(34)&sLogDir&"\Uninstall_"&Product&".log"&chr(34)
            If NOT fDetectOnly Then 
                ' end other instances of setup
                EndCurrentInstalls

                'Execute the uninstall
                LogOnly " - Calling msiexec with '"&sCmd&"'"
                sReturn = oWShell.Run(sCmd, 0, True)
                Log " - msiexec returned: " & SetupRetVal(sReturn) & " (" & sReturn & ")" & vbCrLf
                fRebootRequired = fRebootRequired OR (sReturn = "3010") OR (sReturn = "1618")
            Else
                Log "  -> Removal suppressed in preview mode."
                LogOnly "  -> Command: "&sCmd
            End If
        End If 'InScope
    Next 'Product
    If i = 0 Then Log " Nothing to remove for msiexec"
End Sub 'MsiexecRemoval
'=======================================================================================================

'Remove the OSE (Office Source Engine) service
Sub RemoveOSE
    On Error Resume Next
    DeleteService "ose"
    'Delete the folder
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\Source Engine"
    'Delete the registration
    RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\ose\"
End Sub 'RemoveOSE
'=======================================================================================================


'File cleanup operations for the Local Installation Source (MSOCache)
Sub WipeLIS
    Const LISROOT = "MSOCache\All Users\"
    Dim LogicalDisks, Disk, Folder, SubFolder, MseFolder, File, Files
    Dim arrSubFolders
    Dim sFolder
    Dim fRemoveFolder
    
    'LogH1 "LIS CleanUp"
    'Search all hard disks
    Set LogicalDisks = oWmiLocal.ExecQuery("Select * From Win32_LogicalDisk WHERE DriveType=3")
    For Each Disk in LogicalDisks
        If oFso.FolderExists(Disk.DeviceID & "\" & LISROOT) Then
            Set Folder = oFso.GetFolder(Disk.DeviceID & "\" & LISROOT)
            For Each Subfolder in Folder.Subfolders
                If Len(Subfolder) > 37 Then
                    If fRemoveAll Then 
                        If  (Mid(Subfolder.Name,27,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)=OVERSIONMAJOR) OR _
                            LCase(Right(Subfolder.Name,7)) = OVERSIONMAJOR &".data" Then DeleteFolder Subfolder.Path
                    Else
                        If  (Mid(Subfolder.Name,27,PRODLEN) = OFFICEID AND Mid(SubFolder.Name,4,2)=OVERSIONMAJOR) AND _
                            CheckDelete(UCase(Left(Subfolder.Name,38))) AND _
                            UCase(Right(Subfolder,1))= UCase(Left(Disk.DeviceID,1))Then DeleteFolder Subfolder.Path
                    End If
                End If 'Len > 37
            Next 'Subfolder
            If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                sFolder = Folder.Path
                Set Folder = Nothing
                SmartDeleteFolder sFolder
            End If
        End If 'oFso.FolderExists
    Next 'Disk
    
    'MSECache
    If EnumFolders(sProgramFiles,arrSubFolders) Then
        For Each SubFolder in arrSubFolders
            If UCase(Right(SubFolder,9))="\MSECACHE" Then
                ReDim arrMseFolders(-1)
                Set Folder = oFso.GetFolder(SubFolder)
                GetMseFolderStructure Folder
                For Each MseFolder in arrMseFolders
                    If oFso.FolderExists(MseFolder) Then
                        fRemoveFolder = False
                        Set Folder = oFso.GetFolder(MseFolder)
                        Set Files = Folder.Files
                        For Each File in Files
                            If (LCase(Right(File.Name,4))=".msi") Then
                                If CheckDelete(ProductCode(File.Path)) Then 
                                    fRemoveFolder = True
                                    Exit For
                                End If 'CheckDelete
                            End If
                        Next 'File
                        Set Files = Nothing
                        Set Folder = Nothing
                        If fRemoveFolder Then SmartDeleteFolder MseFolder
                    End If 'oFso.FolderExists(MseFolder)
                Next 'MseFolder
            End If
        Next 'SubFolder
    End If 'oFso.FolderExists
End Sub 'WipeLis
'=======================================================================================================

'Wipe files and folders as documented in KB 928218
Sub FileWipeAll
    Dim sFolder
    Dim Folder, Subfolder
    
    If fForce OR fQuiet OR fPassive Then CloseOfficeApps
    
    'Handle other services.
    Select Case OVERSIONMAJOR
    Case "11"
    Case "12"
    Case "14"
        DeleteService "odserv"
        DeleteService "Microsoft Office Groove Audit Service"
        DeleteService "Microsoft SharePoint Workspace Audit Service"
    Case "15"
    Case Else
    End Select

    'User specific files
    If NOT fKeepUser Then
        'Delete files that should be backed up before deleting them
        CopyAndDeleteFile sAppdata & "\Microsoft\Templates\Normal.dotm"
        CopyAndDeleteFile sAppdata & "\Microsoft\Templates\Normalemail.dotm"
        sFolder = sAppdata & "\microsoft\document building blocks"
        If oFso.FolderExists(sFolder) Then 
            Set Folder = oFso.GetFolder(sFolder)
            For Each Subfolder In Folder.Subfolders
                If oFso.FileExists(Subfolder & "\blocks.dotx") Then CopyAndDeleteFile Subfolder & "\blocks.dotx"
            Next 'Subfolder
            Set Folder = Nothing
        End If 'oFso.FolderExists(sFolder)
    End If  
    
    'Run the individual filewipe from component detection first
    FileWipeIndividual
    If fC2rInstalled AND NOT fForce Then Exit Sub
    
    'Take care of the rest
    LogH2 "General computer specific files"
    DeleteFolder sOInstallRoot
    DeleteFolder sCommonProgramFiles & "\Microsoft Shared\" & OREF
    DeleteFile sAllUsersProfile & "\Application Data\Microsoft\Office\Data\opa"&OVERSIONMAJOR&".dat"
    DeleteFile sAllUsersProfile & "\Application Data\Microsoft\Office\Data\opa"&OVERSIONMAJOR&".bak"
    DeleteFile sAllUsersProfile & "\Microsoft\Office\Data\opa"&OVERSIONMAJOR&".dat"
    DeleteFile sAllUsersProfile & "\Microsoft\Office\Data\opa"&OVERSIONMAJOR&".bak"
    If (fRemoveOspp OR fForce) AND CInt(OVERSIONMAJOR) > 12 Then
        If CInt(OVERSIONMAJOR) = 15 Then CleanOSPP
        DeleteService "osppsvc"
        DeleteFolder sCommonProgramFiles & "\Microsoft Shared\OfficeSoftwareProtectionPlatform"
        DeleteFolder sAllUsersProfile & "\Microsoft\OfficeSoftwareProtectionPlatform"
    End If
    Select Case OVERSIONMAJOR
    Case "12"
    Case "14"
        DeleteFile oWShell.SpecialFolders("AllUsersStartup")&"\OfficeSAS.lnk"
        DeleteFile oWShell.SpecialFolders("Startup")&"\OneNote 2010 Screen Clipper and Launcher.lnk"
    Case "15"
    Case Else
    End Select
    DeleteEmptyFolder sCommonProgramFiles & "\Microsoft Shared\" & OREF
    DeleteEmptyFolder sCommonProgramFiles & "\Microsoft Shared\" 
    DeleteEmptyFolder sProgramFiles & "\Microsoft Office\" & OREF
    DeleteEmptyFolder sProgramFiles & "\Microsoft Office\"

End Sub 'FileWipeAll
'=======================================================================================================

'Wipe individual files & folders related to SKU's that are no longer installed
Sub FileWipeIndividual
    Dim LogicalDisks, Disk,sc
    Dim File, Files, XmlFile, scFiles, oFile, Folder, SubFolder, Processes, Process, item
    Dim sFile, sFolder, sPath, sConfigName, sContents, sProductCode, sLocalDrives,sScQuery
    Dim sValue, sScRoots
    Dim arrSubfolders, arrShortCutRoots
    Dim fKeepFolder, fDeleteSC
    Dim iRet,iCnt,iPos
    
    LogH2 "Individual files"
    If IsArray(arrDeleteFiles) Then
        If fForce OR fQuiet Then
            Log " Doing Action: StopOSE"
            iRet = StopService("ose")
            Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like 'ose%.exe'")
            For Each Process in Processes
                LogOnly " - Running process : " & Process.Name
                Log " -> Ending process: " & Process.Name
                iRet = Process.Terminate()
            Next 'Process
            LogOnly " End Action: StopOSE"
            CloseOfficeApps
        End If
        'Wipe individual files detected earlier
        LogH2 "Remove left behind files"
        For Each sFile in arrDeleteFiles
            If oFso.FileExists(sFile) Then DeleteFile sFile
        Next 'File
    End If 'IsArray
    
    'Wipe Catalyst in commonfiles
    LogH2 "Office Setup Controller - Commonfiles"
    sFolder = sCommonProgramFiles & "\microsoft shared\" & OREF & "\Office Setup Controller\"
    If EnumFolderNames(sFolder,arrSubFolders) Then
        For Each SubFolder in arrSubFolders
            sPath = sFolder & SubFolder
            If InStr(SubFolder, ".") > 0 Then sConfigName = UCase(Left(SubFolder, InStr(SubFolder, ".") - 1)) Else sConfigName = UCase(Subfolder)
            If GetFolderPath(sPath) Then
                Set Folder = oFso.GetFolder(sPath)
                Set Files = Folder.Files
                fKeepFolder = False
                For Each File In Files
                    If Len(File.Name)>3 Then
                        If (LCase(Right(File.Name,4))=".xml") Then
                            If Len(File.Name) >= Len(sConfigName) Then
                                If (UCase(Left(File.Name,Len(sConfigName)))=sConfigName) Then
                                    Set XmlFile = oFso.OpenTextFile(File,1)
                                    sContents = XmlFile.ReadAll
                                    Set XmlFile = Nothing
                                    sProductCode = ""
                                    On Error Resume Next
                                    sProductCode = Mid(sContents,InStr(sContents,"ProductCode=")+Len("ProductCode=")+1,38)
                                    On Error Goto 0
                                    If Len(sProductCode) = 38 Then
                                        If CheckDelete(sProductCode) Then DeleteFile File.Path Else fKeepFolder = True
                                    End If
                                End If 'sConfigName
                            End If 'Len >=
                        End If '.xml
                    End If 'Len(File.Name)>3
                Next 'File
                Set Files = Nothing
                Set Folder = Nothing
                If Not fKeepFolder Then DeleteFolder sPath
            End If 'GetFolderPath
        Next 'SubFolder
    End If 'EnumFolderNames
    
    'Wipe Shortcuts
    If NOT fSkipSD Then
        On Error Resume Next
        LogH2 "Shortcuts"
        CleanShortcuts sAllUsersProfile, True, False
        CleanShortcuts sProfilesDirectory, True, False
        On Error Goto 0
    End If 'NOT SkipSD
    Err.Clear
        
End Sub 'FileWipeIndividual
'=======================================================================================================

'-------------------------------------------------------------------------------
'   CleanShortcuts
'
'   Recursively search all profile folders for Office shortcuts in scope 
'-------------------------------------------------------------------------------
Sub CleanShortcuts (sFolder, fDelete, fUnPin)
    Dim oFolder, fld, file, sc, item
    Dim fDeleteSC

	Set oFolder = oFso.GetFolder(sFolder)
	' exclude system protected link folders
    If CBool(oFolder.Attributes AND 1024) Then Exit Sub

    On Error Resume Next
    For Each fld In oFolder.SubFolders
        If Err <> 0 Then
		    CheckError "CleanShortcuts: " & vbTab & sFolder
        Else
            CleanShortcuts fld.Path, fDelete, fUnPin
        End If
	Next
    For Each file In oFolder.Files
		If LCase(Right(file.Path, 4)) = ".lnk" Then
            fDeleteSC = False
            LogOnly " check file: " & file.Path
            set sc = oWShell.CreateShortcut(file.Path)
            If Err <> 0 Then
		        CheckError "CleanShortcutsSC: " & vbTab & sFolder
            Else
                'Compare if the shortcut target is in the list of executables that will be removed
                'LogOnly "  - SC.TargetPath: " & sc.TargetPath
                If Len(sc.TargetPath) > 0 Then
                    If InStr(sc.TargetPath,"{") > 0 Then
                        'Handle Windows Installer shortcuts
                        If Len(sc.TargetPath) >= InStr(sc.TargetPath,"{") + 37 Then
                            If CheckDelete(Mid(sc.TargetPath, InStr(sc.TargetPath,"{"), 38)) Then fDeleteSC = True
                        End If
                    Else
                        'Handle regular shortcuts
                        If NOT fBypass_Stage1 Then
                            ' Compare against results from component scan
                            For Each item in dicApps.Items
                                If LCase(sc.TargetPath) = item Then
                                    LogOnly "  - removing shortcut per match from component detection: " & file.Path
                                    fDeleteSC = True
                                    Exit For
                                End If
                            Next 'item
                        Else
                        End If
                        If NOT oFso.FileExists(sc.TargetPath) Then
                            ' Shortcut target does not exist
                            If InStr (sc.TargetPath, OREF) > 0 Then
                                LogOnly "  - removing Office shortcut with non-existent target: " & file.Path & " - " & sc.TargetPath
                                fDeleteSC = True
                            Else
                                'LogOnly "  - keep orphaned SC as target is not in scope: " & sc.TargetPath
                            End If
                        Else
                            'LogOnly "  - keep SC as shortcut target does still exist: " & sc.TargetPath
                        End If
                    End If
                End If
            End If
            If fDeleteSC Then 
                If Not IsArray(arrDeleteFolders) Then ReDim arrDeleteFolders(0)
                sFolder = file.Drive & file.Path
                If Not arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder Then
                    ReDim Preserve arrDeleteFolders(UBound(arrDeleteFolders) + 1)
                    arrDeleteFolders(UBound(arrDeleteFolders)) = sFolder
                End If
                If fUnPin OR fDelete Then 
                    If oFso.FileExists(sc.TargetPath) Then
                        UnPin file
                    Else
                        sc.TargetPath = sNotepad
                        sc.Save
                        UnPin file
                    End If
                End If
                If fDelete Then DeleteFile file.Path
                fDeleteSC = False
            End If 'fDeleteSC
        End If
	Next
    On Error Goto 0
End Sub 'CleanShortcuts

'-------------------------------------------------------------------------------
'   UnPin
'
'   Unpins a shortcut from the taskbar or start menu 
'-------------------------------------------------------------------------------
Sub UnPin(file)
    Dim fldItem, verb

    On Error Resume Next
    Set fldItem = oShellApp.NameSpace(file.ParentFolder.Path).ParseName(file.Name)
    For Each verb in fldItem.Verbs
        Select Case LCase(Replace(verb, "&", ""))
        Case "unpin from taskbar", "von taskleiste lösen", "détacher du barre des tâches", "détacher de la barre des tâches", "desanclar de la barra de tareas", "ta bort från aktivitetsfältet", "frigør fra proceslinje", "frigør fra proceslinjen", "desanclar de la barra de tareas", "odepnout z hlavního panelu", "van de taakbalk losmaken", "poista kiinnitys tehtäväpalkista", "rimuovi dalla barra delle applicazioni"
            LogOnly "unpinning Office shortcut from taskbar: " & file.Name 
            verb.DoIt
        Case "unpin from start menu", "vom startmenü lösen", "désépingler du menu démarrer", "supprimer du menu démarrer", "détacher du menu démarrer", "détacher de la menu démarrer", "odepnout z nabídky start", "frigør fra menuen start", "van het menu start losmaken", "losmaken van menu start", "poista kiinnitys käynnistä-valikosta", "irrota aloitusvalikosta"
            LogOnly "unpinning Office shortcut from start menu: " & file.Name 
            If iVersionNT > 600 Then verb.DoIt
        End Select
        Select Case Replace(verb, "&", "")
        Case "从「开始」菜单解锁", "從 [開始] 功能表取消釘選", "タスク バーに表示しない(K)", "작업 표시줄에서 제거(K)", "Открепить от панели задач", "Ξεκαρφίτσωμα από το μενού Έναρξη", "‏‏בטל הצמדה לתפריט התחלה"
            LogOnly "unpinning Office shortcut: " & file.Name 
            verb.DoIt
        End Select
    Next
    On Error Goto 0
End Sub
'=======================================================================================================

Sub DelScrubTmp
    
    On Error Resume Next
    If oFso.FolderExists(sScrubDir & "\ScrubTmp") Then oFso.DeleteFolder sScrubDir & "\ScrubTmp",True

End Sub 'DelScrubTmp
'=======================================================================================================

'Ensure there are no unexpected .msi files in the scrub folder
Sub DeleteMsiScrubCache
    Dim Folder, File, Files
    
    Set Folder = oFso.GetFolder(sScrubDir) : CheckError "DeleteMsiScrubCache"
    Set Files = Folder.Files
    For Each File in Files
        CheckError "DeleteMsiScrubCache"
        If LCase(Right(File.Name,4))=".msi" Then
            CheckError "DeleteMsiScrubCache"
            DeleteFile File.Path : CheckError "DeleteMsiScrubCache"
        End If
    Next 'File
End Sub 'DeleteMsiScrubCache
'=======================================================================================================

Sub MsiClearOrphanedFiles
    Const USERSIDEVERYONE = "s-1-1-0"
    Const MSIINSTALLCONTEXT_ALL = 7
    Const MSIPATCHSTATE_ALL = 15

    'Error handling inlined
    On Error Resume Next

    Dim Patch, AllPatches, Product, AllProducts
    Dim File, Files, Folder
    Dim sFName, sLocalMsp, sLocalMsi, sPatchList, sMsiList

    Set Folder = oFso.GetFolder(sWinDir & "\Installer")
    Set Files = Folder.Files

    'Get a complete list of patches
    Err.Clear
    Set AllPatches = oMsi.PatchesEx("",USERSIDEVERYONE,MSIINSTALLCONTEXT_ALL,MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        CheckError "MsiClearOrphanedFiles (msp)"
    Else
        'Fill a comma separated stringlist with all .msp patchfiles
        For Each Patch in AllPatches
            sLocalMsp = "" : sLocalMsp = LCase(Patch.Patchproperty("LocalPackage")) : CheckError "MsiClearOrphanedFiles (msp)"
            sPatchList = sPatchList & sLocalMsp & ","
        Next 'Patch

        'Delete all non referenced .msp files from %windir%\installer
        For Each File in Files
            sFName = "" : sFName = LCase(File.Path)
            If LCase(Right(sFName,4)) = ".msp" Then
                If Not InStr(sPatchList,sFName) > 0 Then
                    'While this is an orphaned file keep the scope of Office only
                    If InStr(UCase(MspTargets(File.Path)),OFFICEID)>0 Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4))
        Next 'File
    End If 'Err=0

    'Get a complete list products
    Err.Clear
    Set AllProducts = oMsi.ProductsEx("",USERSIDEVERYONE,MSIINSTALLCONTEXT_ALL)
    If Err <> 0 Then
        CheckError "MsiClearOrphanedFiles (msi)"
    Else
        'Fill a comma separated stringlist with all .msi files
        For Each Product in AllProducts
            sLocalMsi = "" : sLocalMsi = LCase(Product.InstallProperty("LocalPackage")) : CheckError "MsiClearOrphanedFiles (msi)"
            sMsiList = sMsiList & sLocalMsi & ","
        Next 'Product

        'Delete all non referenced .msi files from %windir%\installer
        For Each File in Files
            sFName = "" : sFName = LCase(File.Path)
            If LCase(Right(sFName,4)) = ".msi" Then
                If Not InStr(sMsiList,sFName) > 0 Then
                    'While this is an orphaned file keep the scope of Office only
                    If UCase(Right(ProductCode(File.Path),PRODLEN))=OFFICEID Then DeleteFile File.Path
                End If
            End If 'LCase(Right(sFName,4)) = ".msi"
        Next 'File
    End If 'Err=0

End Sub 'MsiClearOrphanedFiles
'=======================================================================================================

Sub RegWipe
    Dim Item, Name, Sku, key
    Dim hDefKey, sSubKeyName, sCurKey, value, sValue, sGuid
    Dim fkeep, fSystemComponent0, fPackages, fDisplayVersion
    Dim arrKeys, arrNames, arrTypes, arrMultiSzValues, arrMultiSzNewValues
    Dim arrTestNames, arrTestTypes
    Dim i, iLoopCnt, iPos
    Dim fDelReg
    
    'Wipe registry data
    
    'User Profile settings
    LogH2 "User Policies"
    RegDeleteKey HKCU,"Software\Policies\Microsoft\Office\" & OVERSION & "\"
    If NOT fKeepUser Then
        RegDeleteKey HKCU,"Software\Microsoft\Office\" & OVERSION & "\"
        LogH2 "User Settings"
    End If 'fKeepUser
    
    'Computer specific settings
    If fClearAddinReg Then
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Outlook\"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Word\"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Excel\"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\PowerPoint\"
    End If
    If (fRemoveAll AND NOT fC2rInstalled) OR (fRemoveAll AND fForce) Then
        LogH2 "Machine Settings"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\" & OVERSION & "\"
        If fRemoveOse OR fForce Then
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office Test\"
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\", "LastAccessInstall", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\", "MID", False
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Excel\Addins\Microsoft.PerformancePoint.Planning.Client.Excel\"
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\InfoPath\Converters\Import\InfoPath.DesignerExcelImport\Versions\", OVERSION, False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\InfoPath\Converters\Import\InfoPath.DesignerWordImport\Versions\", OVERSION, False
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\MEWord12\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\Word12\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Export\Word97\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\MEWord12\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\Word12\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\Word97\"
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "GrooveMonitor", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "LobiServer", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "BCSSync", False
            RegDeleteKey HKLM,"SYSTEM\CurrentControlSet\Services\Outlook\"
        End If
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\OffDiag\Location\", OVERSIONMAJOR, False
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\Software\Microsoft\Office\" & OVERSION & "\"
        RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\Common\OffDiag\Location\", OVERSIONMAJOR, False
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\OfficeCustomizeWizard\" & OVERSION & "\"
        RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\Install\SOFTWARE\Microsoft\OfficeCustomizeWizard\" & OVERSION & "\"
        
        Select Case OVERSIONMAJOR
        Case "11"
            'Jet_Replication
            sValue = ""
            If RegReadValue (HKCR, "CLSID\{CC2C83A6-9BE4-11D0-98E7-00C04FC2CAF5}\InprocServer32", "SystemDB", sValue, "REG_SZ") Then
                If Len(sValue) > Len(sOInstallRoot) Then
                    If LCase(Left(sValue, Len(sOInstallRoot))) = LCase (sOInstallRoot) Then RegDeleteKey HKCR, "CLSID\{CC2C83A6-9BE4-11D0-98E7-00C04FC2CAF5}\InprocServer32\"
                End If
            End If
        Case "12"
        Case "14"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform_Test\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Office\Common\ActiveX Compatibility\{00024512-0000-0000-C000-000000000046}\"
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Office\OneNote\Adapters\", "{456B0D0E-49DD-4C95-8DB6-175F54DE69A3}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{42042206-2D85-11D3-8CFF-005004838597}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{0006F045-0000-0000-C000-000000000046}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{7CCA70DB-DE7A-4FB7-9B2B-52E2335A3B5A}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{D66DC78C-4F61-447F-942B-3FB6980118CF}", False
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}\"
            'Groove Extensions 
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks\", "{B5A7F190-DDA6-4420-B3BA-52453494E6CD}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{99FD978C-D287-4F50-827F-B2C658EDA8E7}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{AB5C5600-7E6E-4B06-9197-9ECEF74D31CC}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{920E6DB1-9907-4370-B3A0-BAFC03D81399}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{16F3DD56-1AF5-4347-846D-7C10C4192619}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{2916C86E-86A6-43FE-8112-43ABE6BF8DCC}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{72853161-30C5-4D22-B7F9-0BBC1D38A37E}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{6C467336-8281-4E60-8204-430CED96822D}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{2A541AE1-5BF6-4665-A8A3-CFA9672E4291}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{B5A7F190-DDA6-4420-B3BA-52453494E6CD}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{A449600E-1DC6-4232-B948-9BD794D62056}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{3D60EDA7-9AB4-4DA8-864C-D9B5F2E7281D}", False
            RegDeleteValue HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\", "{387E725D-DC16-4D76-B310-2C93ED4752A0}", False
            RegDeleteKey HKLM,"SOFTWARE\Classes\*\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX\"
            RegDeleteKey HKLM,"SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX\"
            RegDeleteKey HKLM,"SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX\"
            RegDeleteKey HKLM,"SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX\"
            RegDeleteKey HKLM,"SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\XXX Groove GFS Context Menu Handler XXX\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 1 (GFS Unread Stub)\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 2 (GFS Stub)\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 2.5 (GFS Unread Folder)\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 3 (GFS Folder)\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Groove Explorer Icon Overlay 4 (GFS Unread Mark)\"
            RegDeleteKey HKLM,"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{72853161-30C5-4D22-B7F9-0BBC1D38A37E}\"
        
        Case "15"

        Case Else
        End Select

        'Win32Assemblies
        LogH2 "Win32Assemblies"
        If RegEnumKey(HKCR,"Installer\Win32Assemblies\",arrKeys) Then
            For Each Item in arrKeys
                If InStr(UCase(Item),OREF)>0 Then RegDeleteKey HKCR,"Installer\Win32Assemblies\"&Item & "\"
            Next 'Item
        End If 'RegEnumKey
        'Groove blocks reinstall if it locates groove.exe over this key
        If RegKeyExists(HKCR,"GrooveFile\Shell\Open\Command\") Then
            sValue = ""
            RegReadValue HKCR,"GrooveFile\Shell\Open\Command\","",sValue,"REG_SZ"
            If InStr(sValue,"\"&OREF&"\")>0 Then RegDeleteKey HKCR,"GrooveFile\"
        End If 'RegKeyExists
    End If 'fRemoveAll

    Select Case OVERSIONMAJOR
    Case "11"
        For iLoopCnt = 1 to 3
            Select Case iLoopCnt
            Case 1
                'CIW - HKCU
                sSubKeyName = "Software\Microsoft\OfficeCustomizeWizard\" & OVERSION & "\RegKeyPaths\"
                hDefKey = HKCU
            Case 2 
                'CIW - HKLM
                sSubKeyName = "SOFTWARE\Microsoft\OfficeCustomizeWizard\" & OVERSION & "\RegKeyPaths\"
                hDefKey = HKLM
            Case 3
                'Add/Remove Programs
                sSubKeyName = REG_ARP
                hDefKey = HKLM
            End Select
        
            If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
                For Each Item in arrKeys
                    'OFFICEID id
                    If Len(Item)>37 Then
                        sGuid = UCase(Left(Item,38))
                        If Right(sGuid,PRODLEN)=OFFICEID Then
                            If CheckDelete(sGuid) Then 
                                RegDeleteKey hDefKey, sSubKeyName & Item & "\"
                            End If
                        End If 'Right(Item,PRODLEN)=OFFICEID
                    End If 'Len(Item)>37
                Next 'Item
                If iLoopCnt < 3 Then
                    If RegEnumValues(hDefKey,sSubKeyName,arrNames,arrTypes) Then
                        i = 0
                        For Each Name in arrNames
                            If RegReadValue(hDefKey,sSubKeyName,Name,sValue,arrTypes(i)) Then
                                If sValue = sGuid Then RegDeleteValue hDefKey, sSubKeyName, Name, False
                            End If
                            i = i + 1
                        Next
                    End If
                End If
            End If
            If NOT RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then RegDeleteKey hDefKey,"Software\Microsoft\OfficeCustomizeWizard\11.0\"
            If NOT RegEnumKey(hDefKey,"Software\Microsoft\OfficeCustomizeWizard\11.0\",arrKeys) Then RegDeleteKey hDefKey,"Software\Microsoft\OfficeCustomizeWizard\"
        Next 'iLoopCnt
    Case "12"
        'Add/Remove Programs
        RegWipeARP 
    Case "14"
        'Add/Remove Programs
        RegWipeARP 
    Case Else
    End Select

    'UpgradeCodes, WI config, WI global config
    For iLoopCnt = 1 to 5
        Select Case iLoopCnt
        Case 1
            LogH2 "HKLM UpgradeCodes"
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\"
            hDefKey = HKLM
        Case 2 
            LogH2 "HKCR UpgradeCodes"
            sSubKeyName = "Installer\UpgradeCodes\"
            hDefKey = HKCR
        Case 3
            LogH2 "HKLM Products"
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\"
            hDefKey = HKLM
        Case 4 
            LogH2 "HKCR Features"
            sSubKeyName = "Installer\Features\"
            hDefKey = HKCR
        Case 5 
            LogH2 "HKCR Products"
            sSubKeyName = "Installer\Products\"
            hDefKey = HKCR
        Case Else
            sSubKeyName = ""
            hDefKey = ""
        End Select
        If RegEnumKey(hDefKey,sSubKeyName,arrKeys) Then
            For Each Item in arrKeys
                'Ensure we have the expected length for a compressed GUID
                If Len(Item)=32 Then
                    'Expand the GUID
                    sGuid = GetExpandedGuid(Item) 
                    'Check if it's an Office key
                    If CheckDeleteEx (sGuid) Then
                        RegDeleteKey hDefKey,sSubKeyName & Item & "\"
                    Else
                        If iLoopCnt < 3 Then
                            'Enum all entries
                            RegEnumValues hDefKey,sSubKeyName & Item,arrNames,arrTypes
                            If IsArray(arrNames) Then
                                'Delete entries within removal scope
                                For Each Name in arrNames
                                    If Len(Name)=32 Then
                                        sGuid = GetExpandedGuid(Name)
                                        If CheckDelete(sGuid) Then RegDeleteValue hDefKey, sSubKeyName & Item & "\", Name, True
                                    Else
                                        'Invalid data -> delete the value
                                        RegDeleteValue hDefKey, sSubKeyName & Item & "\", Name, True
                                    End If
                                Next 'Name
                            End If 'IsArray(arrNames)
                            'If all entries were removed - delete the key
                            RegEnumValues hDefKey,sSubKeyName & Item,arrNames,arrTypes
                            If Not IsArray(arrNames) Then RegDeleteKey hDefKey, sSubKeyName & Item & "\"
                        Else 'iLoopCnt >= 3
                            If CheckDelete(sGuid) Then RegDeleteKey hDefKey, sSubKeyName & Item & "\"
                        End If 'iLoopCnt < 3
                    End If 'fRemoveAll
                End If 'Len(Item)=32
            Next 'Item
        End If 'RegEnumKey
    Next 'iLoopCnt

    'Components
    Log " - Global Components"
    sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            'Ensure we have the expected length for a compressed GUID
            If Len(Item)=32 Then
                If RegEnumValues(HKLM,sSubKeyName & Item,arrNames,arrTypes) Then
                    If IsArray(arrNames) Then
                        For Each Name in arrNames
                            If Len(Name)=32 Then
                                sGuid = GetExpandedGuid(Name)
                                If CheckDelete(sGuid) Then
                                    RegDeleteValue HKLM, sSubKeyName & Item & "\", Name, False
                                    'Check if the key is now empty
                                    If NOT RegEnumValues(HKLM,sSubKeyName & Item,arrTestNames,arrTestTypes) Then
                                        If NOT dicDelRegKey.Exists(sSubKeyName&Item&"\") Then dicDelRegKey.Add sSubKeyName&Item&"\",HKCR
                                    End If
                                End If
                            End If '32
                        Next 'Name
                    End If 'IsArray
                End If 'RegEnumValues
            End If '32
        Next 'Item
    End If 'RegEnumKey

    'Published Components
    Log " - Published Components"
    sSubKeyName = "Installer\Components\"
    If RegEnumKey(HKCR,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            'Ensure we have the expected length for a compressed GUID
            If Len(Item)=32 Then
                If RegEnumValues(HKCR,sSubKeyName & Item,arrNames,arrTypes) Then
                    If IsArray(arrNames) Then
                        For Each Name in arrNames
                            If RegReadValue (HKCR,sSubKeyName & Item, Name, sValue,"REG_MULTI_SZ") Then
                                arrMultiSzValues = Split(sValue,chr(34))
                                If IsArray(arrMultiSzValues) Then
                                    i = -1
                                    ReDim arrMultiSzNewValues(-1)
                                    fDelReg = False
                                    For Each value in arrMultiSzValues
                                        If Len(value) > 19 Then
                                            sGuid = ""
                                            If GetDecodedGuid(Left(value,SQUISHED),sGuid) Then
                                                If CheckDelete(sGuid) Then
                                                    fDelReg = True
                                                Else
                                                    i = i + 1 
                                                    ReDim Preserve arrMultiSzNewValues(i)
                                                    arrMultiSzNewValues(i) = value
                                                End If 'CheckDelete
                                            End If 'decode
                                        End If '19
                                    Next 'Value
                                    If NOT (i = -1) Then
                                        If NOT fDetectOnly Then 
                                            If NOT UBound(arrMultiSzValues) = i Then oReg.SetMultiStringValue HKCR,sSubKeyName & Item,Name,arrMultiSzNewValues
                                        End If
                                    Else
                                        If fDelReg Then
                                            RegDeleteValue HKCR,sSubKeyName & Item & "\", Name, False
                                            'Check if the key is now empty
                                            If NOT RegEnumValues(HKCR,sSubKeyName & Item,arrTestNames,arrTestTypes) Then
                                                If NOT dicDelRegKey.Exists(sSubKeyName&Item&"\") Then dicDelRegKey.Add sSubKeyName&Item&"\",HKCR
                                            End If
                                        End If 'DelReg
                                    End If
                                End If 'IsArray
                            End If
                        Next 'Name
                    End If 'IsArray
                End If 'RegEnumValues
            End If '32
        Next 'Item
    End If 'RegEnumKey

    'Delivery
    Log " - Delivery"
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If Len(Item) > 37 Then
                If fRemoveAll Then
                    If (Mid(Item,27,PRODLEN)=OFFICEID AND Mid(Item,4,2)=OVERSIONMAJOR) OR _
                       LCase(Right(Item,7))=OVERSIONMAJOR&".data" Then RegDeleteKey HKLM,sSubKeyName & Item & "\"
                Else
                    If (Mid(Item,27,PRODLEN)=OFFICEID AND Mid(Item,4,2)=OVERSIONMAJOR) AND _
                       CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item & "\"
                End If
            End If '37
        Next 'Item
    End If 'RegEnumKey
    
    'Registration
    Log " - HKLM Registration"
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\"&OVERSION&"\Registration\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If Len(Item)>37 Then
                If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item & "\"
            End If
        Next 'Item
    End If 'RegEnumKey
    
    'User Preconfigurations
    Log " - HKLM User Preconfigurations"
    hDefKey = HKLM
    sSubKeyName = "SOFTWARE\Microsoft\Office\"&OVERSION&"\User Settings\"
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            If Len(Item)>37 Then
                If CheckDelete(UCase(Left(Item,38))) Then RegDeleteKey HKLM,sSubKeyName & Item & "\"
            End If
        Next 'Item
    End If 'RegEnumKey

    'Known Keypath settings
    Log " - Detcted KeyPath settings"
    For Each key in dicDelRegKey.Keys
        If Right(key,1) = "\" Then
            RegDeleteKey dicDelRegKey.Item(key),key
        Else
            iPos = InStrRev(Key,"\")
            If iPos > 0 Then RegDeleteValue dicDelRegKey.Item(key), Left(key,iPos - 1), Mid(key,iPos+1), False
        End If
    Next

    'Temporary entries in ARP
    TmpKeyCleanUp
End Sub 'RegWipe
'=======================================================================================================

'Clean up Add/Remove Programs registry
Sub RegWipeARP

    Dim Item, Name, Sku, key
    Dim sSubKeyName, sCurKey, sValue, sGuid
    Dim fkeep, fSystemComponent0, fPackages, fDisplayVersion
    Dim arrKeys

    'Add/Remove Programs
    sSubKeyName = REG_ARP
    If RegEnumKey(HKLM,sSubKeyName,arrKeys) Then
        For Each Item in arrKeys
            '*0FF1CE*
            If Len(Item) > 37 Then
                sGuid = UCase(Left(Item, 38))
                If CheckDeleteEx(sGuid) Then RegDeleteKey HKLM, sSubKeyName & Item
            End If 'Len(Item)>37
            
            'Config entries
            sCurKey = sSubKeyName & Item & "\"
            fSystemComponent0 = Not (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
            fPackages = RegReadValue(HKLM,sCurKey,OPACKAGE,sValue,"REG_MULTI_SZ")
            fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ")
            If fDisplayVersion AND Len(sValue) > 1 Then
                fDisplayVersion = (Left(sValue, 2) = OVERSIONMAJOR)
            End If
            If (fPackages AND fDisplayVersion) Then
                fKeep = False
                If Not fRemoveAll Then
                    For Each Sku in dicKeepSku.Keys
                        If UCase(Item) = OREGREF & Sku Then
                            fkeep = True
                            Exit For
                        End If
                    Next 'Sku
                End If
                If Not fkeep Then RegDeleteKey HKLM, sSubKeyName & Item
            End If
        Next 'Item
    End If 'RegEnumKey

End Sub 'RegWipeARP
'=======================================================================================================

'Clean up temporary registry keys
Sub TmpKeyCleanUp
    Dim TmpKey
    
    If fLogInitialized Then Log " - temporary OffScrub registry entries"
    If IsArray(arrTmpSKUs) Then
        For Each TmpKey in arrTmpSKUs
            oReg.DeleteKey HKLM, REG_ARP & TmpKey
        Next 'Item
    End If 'IsArray
End Sub 'TmpKeyCleanUp

'=======================================================================================================
' Helper Functions
'=======================================================================================================

'Create a log with the results of the SKU detection
Sub LogSkuResults
    Dim SkuLog, SkuKey , p

    On Error Resume Next 'Don't fail on logging
    
    Set SkuLog = oFso.OpenTextFile(sScrubDir & "\SkuLog.txt",FOR_WRITING,True,True)
    
    SkuLog.WriteLine "Installed SKUs (All):"
    SkuLog.WriteLine "====================="
    For Each SkuKey in dicInstalledSku.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Server SKUs:"
    SkuLog.WriteLine          "============"
    For Each SkuKey in dicSrv.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Client Suite SKUs:"
    SkuLog.WriteLine          "=================="
    For Each SkuKey in dicCSuite.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Client Standalone SKUs:"
    SkuLog.WriteLine          "======================="
    For Each SkuKey in dicCSingle.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Installed Products (All):"
    SkuLog.WriteLine          "========================="
    For Each p in oMsi.Products
        If InScope(p) Then
            SkuLog.Write " - " & p & " - "
            SkuLog.Write oMsi.ProductInfo(p, "ProductName")
            SkuLog.WriteLine " "
        End If
    Next 'Product

    SkuLog.WriteLine vbCrLf & "***************************************************************************************************" & vbCrLf

    SkuLog.WriteLine vbCrLf & "SKUs to keep:"
    SkuLog.WriteLine          "============="
    For Each SkuKey in dicKeepSku.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Products to keep:"
    SkuLog.WriteLine          "================="
    For Each p in dicKeepProd.Keys
        SkuLog.Write " - " & p & " - "
        SkuLog.Write oMsi.ProductInfo(p, "ProductName")
        SkuLog.WriteLine " "
    Next 'Key

    SkuLog.WriteLine vbCrLf & "***************************************************************************************************" & vbCrLf

    SkuLog.WriteLine vbCrLf & "SKUs to remove:"
    SkuLog.WriteLine          "==============="
    For Each SkuKey in dicRemoveSku.Keys
        SkuLog.WriteLine " - " & SkuKey
    Next 'Key

    SkuLog.WriteLine vbCrLf & "Products to remove:"
    SkuLog.WriteLine          "==================="
    For Each p in oMsi.Products
        If CheckDeleteEx(p) Then
            SkuLog.Write " - " & p & " - "
            SkuLog.Write oMsi.ProductInfo(p, "ProductName")
            SkuLog.WriteLine " "
        End If 'InScope
    Next 'Product

    SkuLog.Close
    Set SkuLog = Nothing

End Sub 'LogSkuResults
'=======================================================================================================

'End all running instances of applications that will be removed
Sub CloseOfficeApps
    Dim Processes, Process, prop
    Dim fWait
    Dim iRet
    
    On Error Resume Next
    
    fWait = False
    Log " Doing Action: CloseOfficeApps"

    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If dicApps.Exists(LCase(Process.Name)) Then
            Log " - End process " & Process.Name
            iRet = Process.Terminate()
            CheckError "CloseOfficeApps: " & "Process.Name"
        Else
            For Each prop in Process.Properties_
                If prop.Name = "ExecutablePath" Then 
                    If InStr(UCase(prop.Value), UCase(sOInstallRoot)) > 0 Then
                        Log " - End process '" & Process.Name
                        iRet = Process.Terminate()
                        CheckError "CloseOfficeApps: " & "Process.Name"
                        fWait = True
                    End If 
                End If 'ExcecutablePath
            Next 'prop
        End If
    Next 'Process
    If fWait Then
        wscript.sleep 10000
    End If
    LogOnly " End Action: CloseOfficeApps"
End Sub 'CloseOfficeApps
'=======================================================================================================

'Ensure Windows Explorer is restarted if needed
Sub RestoreExplorer
    Dim Processes, Result, oAT, DateTime, JobID
    Dim sCmd
    
    'Non critical routine. Don't fail on error
    On Error Resume Next
    wscript.sleep 1000
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='explorer.exe'")
    If Processes.Count < 1 Then 
        oWShell.Run "explorer.exe"
        'To handle this in case of System context, schedule and run as interactive task
        If iVersionNT > 502 Then
            'Vista and later
            oWShell.Run "SCHTASKS /Create /TN OffScrEx /TR explorer /SC ONCE /ST 12:00 /IT",0,True
            oWShell.Run "SCHTASKS /Run /TN OffScrEx", 0, True
            oWShell.Run "SCHTASKS /Delete /TN OffScrEx /F", 0, False
        Else
            Set oAT = oWmiLocal.Get("Win32_ScheduledJob")
            Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
            DateTime.SetVarDate DateAdd("n",1,Now),True
            Result = oAT.Create("explorer.exe", DateTime.Value, , , , True, JobID)
        End If 'iVersionNT
    End If
End Sub 'RestoreExploer
'=======================================================================================================

'Returns the delimiter for a passed in string
Function Delimiter (sVersion)

    Dim iCnt, iAsc

    Delimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            Delimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function
'=======================================================================================================

'Check registry access permissions. Failure will terminate the script
Function CheckRegPermissions
    Const KEY_QUERY_VALUE       = &H0001
    Const KEY_SET_VALUE         = &H0002
    Const KEY_CREATE_SUB_KEY    = &H0004
    Const DELETE                = &H00010000

    Dim sSubKeyName
    Dim fReturn

    CheckRegPermissions = True
    sSubKeyName = "Software\Microsoft\Windows\"
    oReg.CheckAccess HKLM, sSubKeyName, KEY_QUERY_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_SET_VALUE, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, KEY_CREATE_SUB_KEY, fReturn
    If Not fReturn Then CheckRegPermissions = False
    oReg.CheckAccess HKLM, sSubKeyName, DELETE, fReturn
    If Not fReturn Then CheckRegPermissions = False

End Function 'CheckRegPermissions
'=======================================================================================================

'Check if a product will be removed
Function CheckDeleteEx (sProductCode)
        
    CheckDeleteEx = False
    If CheckDelete (sProductCode) Then
        CheckDeleteEx = True
        Exit Function
    End If
    If (fRemoveAll AND NOT fC2rInstalled) OR (fRemoveAll AND fForce) Then
        CheckDeleteEx = InScope(sProductCode) AND NOT dicKeepProd.Exists(UCase(sProductCode))
    End If
End Function 'CheckDelete
'=======================================================================================================

'Check if an Office product is still registered with a SKU that stays on the computer
Function CheckDelete (sProductCode)
        
    'Ensure valid GUID length
    If NOT Len(sProductCode) = 38 Then
        CheckDelete = False
        Exit Function
    End If

    'If it's a non Office ProductCode exit with false right away
    CheckDelete = InScope(sProductCode)
    If Not CheckDelete Then Exit Function
    If dicKeepProd.Exists(UCase(sProductCode)) Then CheckDelete = False

End Function 'CheckDelete
'=======================================================================================================

'Check if ProductCode is in scope
Function InScope(sProductCode)

    Dim fInScope
    Dim sProd

    fInScope = False
    If Len(sProductCode) = 38 Then
        sProd = UCase(sProductCode)
        Select Case OVERSIONMAJOR
        Case "11"
            If Right(sProd,PRODLEN)=OFFICEID Then fInScope = True
        Case "12"
            If Right(sProd,PRODLEN)=OFFICEID AND Mid(sProd,4,2) = OVERSIONMAJOR Then fInScope = True
        Case "14"
            If Right(sProd,PRODLEN)=OFFICEID AND Mid(sProd,4,2) = OVERSIONMAJOR Then fInScope = True
        Case "15"
            If Right(sProd,PRODLEN)=OFFICEID AND Mid(sProd,4,2) = OVERSIONMAJOR Then 
                Select Case Mid(sProd, 11, 4)
                Case "007E", "008F", "008C", "24E1", "237A"
                    ' C2R products - keep them
                Case Else
                    fInScope = True
                End Select
            End If
        Case Else
        End Select
    End If '38

    InScope = fInScope
End Function 'InScope
'=======================================================================================================

'Register an orphaned .msi product as installed for MSI
Sub MsiRegisterProduct (sMsiFile)

    Dim sDisplayVersion, sCurKey, sDisplayName, sLang, sProductCode, sTmpKey
    Dim iCnt

    'Create a temporary keys to simulate an installed product
    sProductCode = ""
    sProductCode = GetMsiProductCode(sMsiFile)
    sDisplayVersion = GetMsiProductVersion(sMsiFile)
    If sDisplayVersion = "" Then sDisplayVersion = OVERSION & ".0000.0000"
    sDisplayName = GetMsiProductName(sMsiFile)
    If sDisplayName = "" Then sDisplayName = sProductCode
    Select Case OVERSIONMAJOR
    Case "9","10","11"
        sLang = CInt("&h" & Mid(sProductCode,6,4))
    Case "12","14"
        sLang = CInt("&h" & Mid(sProductCode,16,4))
    Case Else
    End Select

    For iCnt = 1 To 3
        Select Case iCnt
        Case 1
            sCurKey = REG_ARP & sProductCode
            oReg.CreateKey HKLM,sCurKey
        Case 2
            sCurKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\" & GetCompressedGuid(sProductCode)
            oReg.CreateKey HKLM,sCurKey
            oReg.CreateKey HKLM,sCurKey & "\Features"
            oReg.CreateKey HKLM,sCurKey & "\InstallProperties"
            oReg.CreateKey HKLM,sCurKey & "\Patches"
            oReg.CreateKey HKLM,sCurKey & "\Usage"
            sCurKey = sCurKey & "\InstallProperties"
            oReg.SetStringValue HKLM,sCurKey,"LocalPackage",sMsiFile
        Case 3
            sCurKey = "Installer\Products\" & GetCompressedGuid(sProductCode)
            sTmpKey = sCurKey
            oReg.CreateKey HKCR,sCurKey
            oReg.SetDWordValue HKCR,sCurKey,"AdvertiseFlags",388
            oReg.SetDWordValue HKCR,sCurKey,"Assignment",1
            oReg.SetDWordValue HKCR,sCurKey,"AuthorizedLUAApp",0
            oReg.SetStringValue HKCR,sCurKey,"Clients",":"
            oReg.SetDWordValue HKCR,sCurKey,"DeploymentFlags",3
            oReg.SetDWordValue HKCR,sCurKey,"InstanceType",0
            oReg.SetDWordValue HKCR,sCurKey,"Language",sLang
            oReg.SetStringValue HKCR,sCurKey,"PackageCode",GetMsiPackageCode(sMsiFile)
            oReg.SetStringValue HKCR,sCurKey,"ProductName",sDisplayName
            oReg.SetDWordValue HKCR,sCurKey,"VersionMinor",0
            sCurKey = sTmpKey & "\SourceList"
            oReg.CreateKey HKCR,sCurKey
            oReg.SetExpandedStringValue HKCR,sCurKey,"LastUsedSource",sScrubDir
            oReg.SetStringValue HKCR,sCurKey,"PackageName",Mid(sMsiFile,InstrRev(sMsiFile,"\")+1)
            sCurKey = sTmpKey & "\SourceList\Media"
            oReg.CreateKey HKCR,sCurKey
            oReg.SetStringValue HKCR,sCurKey,"1",OREF & ";1"
            oReg.SetStringValue HKCR,sCurKey,"DiskPrompt",sDisplayName
            sCurKey = sTmpKey & "\SourceList\Net"
            oReg.CreateKey HKCR,sCurKey
            oReg.SetExpandedStringValue HKCR,sCurKey,"1",sScrubDir

        Case Else
        End Select
        If iCnt <3 Then
            oReg.SetStringValue HKLM,sCurKey,"Comments",""
            oReg.SetStringValue HKLM,sCurKey,"Contact",""
            oReg.SetStringValue HKLM,sCurKey,"DisplayName",sDisplayName
            oReg.SetStringValue HKLM,sCurKey,"DisplayVersion",sDisplayVersion
            oReg.SetDWordValue HKLM,sCurKey,"EstimatedSize",0
            oReg.SetStringValue HKLM,sCurKey,"HelpLink",""
            oReg.SetStringValue HKLM,sCurKey,"HelpTelephone",""
            oReg.SetStringValue HKLM,sCurKey,"InstallDate","20100101"
            If f64 Then
                oReg.SetStringValue HKLM,sCurKey,"InstallLocation",sProgramFilesX86
            Else
                oReg.SetStringValue HKLM,sCurKey,"InstallLocation",sProgramFiles
            End If
            oReg.SetStringValue HKLM,sCurKey,"InstallSource",sScrubDir
            oReg.SetDWordValue HKLM,sCurKey,"Language",sLang
            oReg.SetExpandedStringValue HKLM,sCurKey,"ModifyPath","MsiExec.exe /X" & sProductCode
            oReg.SetDWordValue HKLM,sCurKey,"NoModify",1
            oReg.SetStringValue HKLM,sCurKey,"Publisher","Microsoft Corporation"
            oReg.SetStringValue HKLM,sCurKey,"Readme",""
            oReg.SetStringValue HKLM,sCurKey,"Size",""
            oReg.SetDWordValue HKLM,sCurKey,"SystemComponent",0
            oReg.SetExpandedStringValue HKLM,sCurKey,"UninstallString","MsiExec.exe /X" & sProductCode
            oReg.SetStringValue HKLM,sCurKey,"URLInfoAbout",""
            oReg.SetStringValue HKLM,sCurKey,"URLUpdateInfo",""
            oReg.SetDWordValue HKLM,sCurKey,"Version",0
            oReg.SetDWordValue HKLM,sCurKey,"VersionMajor",OVERSIONMAJOR
            oReg.SetDWordValue HKLM,sCurKey,"VersionMinor",0
            oReg.SetDWordValue HKLM,sCurKey,"WindowsInstaller",1
        End If '< 3
    Next 'iCnt

End Sub 'MsiRegisterProduct
'=======================================================================================================

'Obtain the ProductCode (GUID) from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode
Function GetMsiProductCode(sMsiFile)
    
    Dim MsiDb,Record
    Dim qView
    
    On Error Resume Next
    
    GetMsiProductCode = ""
    Set Record = Nothing
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEREADONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductCode'")
    qView.Execute
    Set Record = qView.Fetch
    GetMsiProductCode = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductCode
'=======================================================================================================

'Obtain the ProductVersion from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode
Function GetMsiProductVersion(sMsiFile)
    
    Dim MsiDb,Record
    Dim qView
    
    On Error Resume Next
    
    GetMsiProductVersion = ""
    Set Record = Nothing
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEREADONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductVersion'")
    qView.Execute
    Set Record = qView.Fetch
    GetMsiProductVersion = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductVersion
'=======================================================================================================

'Obtain the ProductVersion from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode
Function GetMsiProductName(sMsiFile)
    
    Dim MsiDb,Record
    Dim qView
    
    On Error Resume Next
    
    GetMsiProductName = ""
    Set Record = Nothing
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEREADONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductName'")
    qView.Execute
    Set Record = qView.Fetch
    GetMsiProductName = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductVersion
'=======================================================================================================

'Obtain the PackageCode (GUID) from a .msi package
'The function will the .msi'S SummaryInformation stream
Function GetMsiPackageCode(sMsiFile)

    On Error Resume Next

    Const PID_REVNUMBER = 9
    
    GetMsiPackageCode = ""
    GetMsiPackageCode = GetCompressedGuid(oMsi.SummaryInformation(sMsiFile,MSIOPENDATABASEREADONLY).Property(PID_REVNUMBER))

End Function 'GetMsiPackageCode
'=======================================================================================================

'Returns a string with a list of ProductCodes from the summary information stream
Function MspTargets (sMspFile)
    Const MSIOPENDATABASEMODE_PATCHFILE = 32
    Const PID_TEMPLATE                  =  7

    Dim Msp
    'Non critical routine. Don't fail on error
    On Error Resume Next
    MspTargets = ""
    If oFso.FileExists(sMspFile) Then
        Set Msp = Msi.OpenDatabase(WScript.Arguments(0),MSIOPENDATABASEMODE_PATCHFILE)
        If Err = 0 Then MspTargets = Msp.SummaryInformation.Property(PID_TEMPLATE)
    End If 'oFso.FileExists(sMspFile)
End Function 'MspTargets
'=======================================================================================================

'Return the ProductCode {GUID} from a .MSI package
Function ProductCode(sMsi)
    Const MSIUILEVELNONE = 2 'No UI
    Dim MsiSession

    On Error Resume Next
    'Non critical routine. Don't fail on error
    If oFso.FileExists(sMsi) Then
        oMsi.UILevel = MSIUILEVELNONE
        Set MsiSession = oMsi.OpenPackage(sMsi,1)
        ProductCode = MsiSession.ProductProperty("ProductCode")
        Set MsiSession = Nothing
    Else
        ProductCode = ""
    End If 'oFso.FileExists(sMsi)
End Function 'ProductCode
'=======================================================================================================

Function GetUpgradeCode(sGuid)

    'Ensure Valid Length
    If NOT Len(sGuid) = 38 Then Exit Function
    
    GetUpgradeCode = "{00" & Mid(sGuid, 4, 2) & "0000-" & Mid(sGuid, 11, 4) & "-0000-" & Mid(sGuid, 21, 1) & "000-" & Mid(sGuid, 26, 1) & "000000FF1CE}"

End Function 'GetUpgradeCode
'=======================================================================================================

Function GetExpandedGuid (sGuid)
    Dim i

    'Ensure valid length
    If NOT Len(sGuid) = 32 Then Exit Function

    GetExpandedGuid = "{" & StrReverse(Mid(sGuid,1,8)) & "-" & _
                       StrReverse(Mid(sGuid,9,4)) & "-" & _
                       StrReverse(Mid(sGuid,13,4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i + 1),1)
	    Else
		    GetExpandedGuid = GetExpandedGuid & mid(sGuid,(i - 1),1)
	    End If
    Next
    GetExpandedGuid = GetExpandedGuid & "}"
End Function
'=======================================================================================================

'Converts a GUID into the compressed format
Function GetCompressedGuid (sGuid)
    Dim sCompGUID
    Dim i
    
    'Ensure Valid Length
    If NOT Len(sGuid) = 38 Then Exit Function

    sCompGUID = StrReverse(Mid(sGuid,2,8))  & _
                StrReverse(Mid(sGuid,11,4)) & _
                StrReverse(Mid(sGuid,16,4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function
'=======================================================================================================

'Unsquish GUID
Function GetDecodedGuid(sEncGuid, sGuid)

Dim sDecode, sTable, sHex, iChr
Dim arrTable
Dim i, iAsc, pow85, decChar
Dim lTotal
Dim fFailed

    fFailed = False

    sTable =    "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," & _
                "0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff," & _
                "0xff,0x00,0xff,0xff,0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,0x09,0x0a,0x0b,0xff," & _
                "0x0c,0x0d,0x0e,0x0f,0x10,0x11,0x12,0x13,0x14,0x15,0xff,0xff,0xff,0x16,0xff,0x17," & _
                "0x18,0x19,0x1a,0x1b,0x1c,0x1d,0x1e,0x1f,0x20,0x21,0x22,0x23,0x24,0x25,0x26,0x27," & _
                "0x28,0x29,0x2a,0x2b,0x2c,0x2d,0x2e,0x2f,0x30,0x31,0x32,0x33,0xff,0x34,0x35,0x36," & _
                "0x37,0x38,0x39,0x3a,0x3b,0x3c,0x3d,0x3e,0x3f,0x40,0x41,0x42,0x43,0x44,0x45,0x46," & _
                "0x47,0x48,0x49,0x4a,0x4b,0x4c,0x4d,0x4e,0x4f,0x50,0x51,0x52,0xff,0x53,0x54,0xff"
    arrTable = Split(sTable,",")
    lTotal = 0 : pow85 = 1
    For i = 0 To 19
        fFailed = True
        If i Mod 5 = 0 Then
            lTotal = 0 : pow85 = 1
        End If ' i Mod 5 = 0
        iAsc = Asc(Mid(sEncGuid,i+1,1))
        sHex = arrTable(iAsc)
        If iAsc >=128 Then Exit For
        If sHex = "0xff" Then Exit For
        iChr = CInt("&h"&Right(sHex,2))
        lTotal = lTotal + (iChr * pow85)
        If i Mod 5 = 4 Then sDecode = sDecode & DecToHex(lTotal)
        pow85 = pow85 * 85
        fFailed = False
    Next 'i
    If NOT fFailed Then sGuid = "{"&Mid(sDecode,1,8)&"-"& _
                                Mid(sDecode,13,4)&"-"& _
                                Mid(sDecode,9,4)&"-"& _
                                Mid(sDecode,23,2) & Mid(sDecode,21,2)&"-"& _
                                Mid(sDecode,19,2) & Mid(sDecode,17,2) & Mid(sDecode,31,2) & Mid(sDecode,29,2) & Mid(sDecode,27,2) & Mid(sDecode,25,2) &"}"

    GetDecodedGuid = NOT fFailed

End Function 'GetDecodedGuid
'=======================================================================================================

'Convert a long decimal to hex
Function DecToHex(lDec)
    
    Dim sHex
    Dim iLen
    Dim lVal, lExp
    Dim arrChr
  
    arrChr = Array("0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F")
    sHex = ""
    lVal = lDec
    lExp = 16^10
    While lExp >= 1
        If lVal >= lExp Then
            sHex = sHex & arrChr(Int(lVal / lExp))
            lVal = lVal - lExp * Int(lVal / lExp)
        Else
            sHex = sHex & "0"
            If sHex = "0" Then sHex = ""
        End If
        lExp = lExp / 16
    Wend

    iLen = 8 - Len(sHex)
    If iLen > 0 Then sHex = String(iLen,"0") & sHex
    DecToHex = sHex
End Function
'=======================================================================================================

'Ensures that only valid metadata entries exist to avoid API failures
Sub EnsureValidWIMetadata (hDefKey,sKey,iValidLength)

Dim arrKeys
Dim SubKey

If Len(sKey) > 1 Then
    If Right(sKey,1) = "\" Then sKey = Left(sKey,Len(sKey)-1)
End If

If RegEnumKey(hDefKey,sKey,arrKeys) Then
    For Each SubKey in arrKeys
        If NOT Len(SubKey) = iValidLength Then
            RegDeleteKey hDefKey,sKey & "\" & SubKey & "\"
        End If
    Next 'SubKey
End If

End Sub 'EnsureValidWIMetadata
'=======================================================================================================

'-------------------------------------------------------------------------------
'   CleanOSPP
'
'   Clean out licenses from the Office Software Protection Platform 
'-------------------------------------------------------------------------------
Sub CleanOSPP
    Dim oProductInstances, pi
    Dim sCleanOSPP, sCmd, sRetVal

    CONST OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663"  'Office 2013

    sCleanOSPP = "x64\CleanOSPP.exe"
    If Not f64 Then sCleanOSPP = "x86\CleanOSPP.exe"
    If oFso.FileExists(sScriptDir & sCleanOSPP) Then
        sCmd = sScriptDir & sCleanOSPP
        Log "   Running: " & sCmd
        On Error Resume Next
        sRetVal = oWShell.Run(sCmd, 0, True)
        Log "   Return value: " & sRetVal
        On Error Goto 0
        Exit Sub
    End If
    
    On Error Resume Next
    ' Initialize the software protection platform object with a filter on Office 2013 products
    If iVersionNT > 601 Then
        Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM SoftwareLicensingProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
    Else
        Set oProductInstances = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '" & OfficeAppId & "' " & "AND PartialProductKey <> NULL")
    End If

    ' Remove all licenses
    For Each pi in oProductInstances
        If NOT IsNull(pi) Then
            pi.UninstallProductKey( pi.ProductKeyID)
        End If
    Next 'pi

End Sub 'CleanOSPP
'=======================================================================================================


'Create a backup copy of the file in the ScrubDir then delete the file
Sub CopyAndDeleteFile(sFile)
    Dim File
    
    'Error handling inlined
    On Error Resume Next
    If oFso.FileExists(sFile) Then
        Set File = oFso.GetFile(sFile)
        If Not oFso.FolderExists(sScrubDir & "\" & File.ParentFolder.Name) Then oFso.CreateFolder sScrubDir & "\" & File.ParentFolder.Name
        If Not fDetectOnly Then
            LogOnly " - Backing up file: " & sFile
            oFso.CopyFile sFile,sScrubDir & "\" & File.ParentFolder.Name & "\" & File.Name,True : CheckError "CopyAndDeleteFile"
            Set File = Nothing
            DeleteFile(sFile)
        Else
            LogOnly " - Simulate CopyAndDelete file: " & sFile
        End If
    End If 'oFso.FileExists
End Sub 'CopyAndDeleteFile
'=======================================================================================================

'Wrapper to delete a file
Sub DeleteFile(sFile)
    Dim File
    Dim sFileName, sNewPath
    
    On Error Resume Next

    If dicKeepFolder.Exists(LCase(sFile)) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & sFile
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element: " & sFile
            LogOnly "   Remaining applications will need a repair!"
        End If
    End If
    If f64 Then
        If dicKeepFolder.Exists(LCase(Wow64Folder(sFile))) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & sFile
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element: " & sFile
            LogOnly "   Remaining applications will need a repair!"
        End If
        End If
    End If

    If oFso.FileExists(sFile) Then
        LogOnly " - Delete file: " & sFile
        If Not fDetectOnly Then oFso.DeleteFile sFile,True
        If Err <> 0 Then
            CheckError "DeleteFile"
            If fForce Then
                'Try to move the file and delete from there
                Set File = oFso.GetFile(sFile)
                sFileName = File.Name
                sNewPath = sScrubDir & "\ScrubTmp"
                Set File = Nothing
                If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
                'Move the file
                LogOnly " - Move file to: " & sNewPath & "\" & sFileName
                oFso.MoveFile sFile,sNewPath & "\" & sFileName
                If Err <> 0 Then 
                    CheckError "DeleteFile (move)"
                End If 'Err <> 0
            Else
                fRebootRequired = True
            End If 'fForce
        End If 'Err <> 0
    End If 'oFso.FileExists
End Sub 'DeleteFile
'=======================================================================================================

'64 bit aware wrapper to return the requested folder 
Function GetFolderPath(sPath)
    GetFolderPath = True
    If oFso.FolderExists(sPath) Then Exit Function
    If f64 AND oFso.FolderExists(Wow64Folder(sPath)) Then
        sPath = Wow64Folder(sPath)
        Exit Function
    End If
    GetFolderPath = False
End Function 'GetFolderPath
'=======================================================================================================

'Enumerates subfolder names of a folder and returns True if subfolders exist
Function EnumFolderNames (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolderNames = Len(sSubFolders)>0
End Function 'EnumFolderNames
'=======================================================================================================

'Enumerates subfolders of a folder and returns True if subfolders exist
Function EnumFolders (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders)>0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders,Len(sSubFolders)-1),","))
    EnumFolders = Len(sSubFolders)>0
End Function 'EnumFolders
'=======================================================================================================

Sub GetMseFolderStructure (Folder)
    Dim SubFolder
    
    For Each SubFolder in Folder.SubFolders
        ReDim Preserve arrMseFolders(UBound(arrMseFolders)+1)
        arrMseFolders(UBound(arrMseFolders)) = SubFolder.Path
        GetMseFolderStructure SubFolder
    Next 'SubFolder
End Sub 'GetMseFolderStructure
'=======================================================================================================

'Wrapper to delete a folder 
Sub DeleteFolder(sFolder)
    Dim Folder
    Dim sDelFolder, sFolderName, sNewPath
    
    'Ensure trailing "\"
    sFolder = sFolder & "\"
    While InStr(sFolder,"\\")>0
        sFolder = Replace(sFolder,"\\","\")
    Wend

    If dicKeepFolder.Exists(LCase(sFolder)) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & sFolder
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element: " & sFolder
            LogOnly "   Remaining applications will need a repair!"
        End If
    End If
    If f64 Then
        If dicKeepFolder.Exists(LCase(Wow64Folder(sFolder))) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & sFolder
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element: " & sFolder
            LogOnly "   Remaining applications will need a repair!"
        End If
        End If
    End If
    
    'Strip trailing "\"
    If Len(sFolder) > 1 Then
        sFolder = Left(sFolder,Len(sFolder)-1)
    End If

    On Error Resume Next
    If oFso.FolderExists(sFolder) Then 
        sDelFolder = sFolder
    ElseIf f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        sDelFolder = Wow64Folder(sFolder)
    Else
        Exit Sub
    End If
    If Not fDetectOnly Then 
        LogOnly " - Delete folder: " & sDelFolder
        oFso.DeleteFolder sDelFolder,True
    Else
        LogOnly " - Simulate delete folder: " & sDelFolder
    End If
    If Err <> 0 Then
        CheckError "DeleteFolder"
        'Try to move the folder and delete from there
        Set Folder = oFso.GetFolder(sDelFolder)
        sFolderName = Folder.Name
        sNewPath = sScrubDir & "\ScrubTmp"
        Set Folder = Nothing
        'Ensure we stay within the same drive
        If Not oFso.FolderExists(sNewPath) Then oFso.CreateFolder(sNewPath)
        'Move the folder
        LogOnly " - Moving folder to: " & sNewPath & "\" & sFolderName
        oFso.MoveFolder sFolder,sNewPath & "\" & sFolderName
        If Err <> 0 Then
            CheckError "DeleteFolder (move)"
        End If 'Err <> 0
    End If 'Err <> 0
End Sub 'DeleteFolder
'=======================================================================================================

'Delete empty folder structures
Sub DeleteEmptyFolders
    Dim Folder
    Dim sFolder
    
    ' cosmetic' task don't fail on error
    On Error Resume Next
    If Not IsArray(arrDeleteFolders) Then Exit Sub
    'LogH2 "Empty Folder Cleanup"
    For Each sFolder in arrDeleteFolders
        If oFso.FolderExists(sFolder) Then
            Set Folder = oFso.GetFolder(sFolder)
            If CBool(Folder.Attributes AND 1024) Then
                'exclude protected folder
            Else
                If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                    Set Folder = Nothing
                    SmartDeleteFolder sFolder
                End If
            End If
        End If
    Next 'sFolder
    CheckError "DeleteEmptyFolders"
    On Error Goto 0
End Sub 'DeleteEmptyFolders
'=======================================================================================================

'Delete indivdual empty folder structures
Sub DeleteEmptyFolder (sFolder)
    Dim Folder
    
    ' cosmetic' task don't fail on error
    On Error Resume Next
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        If CBool(Folder.Attributes AND 1024) Then
            'exclude protected folder
        Else
            If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then 
                Set Folder = Nothing
                SmartDeleteFolder sFolder
            End If
        End If
    End If

    CheckError "DeleteEmptyFolder"
    On Error Goto 0
End Sub 'DeleteEmptyFolder
'=======================================================================================================

'Wrapper to delete a folder and remove the empty parent folder structure
Sub SmartDeleteFolder(sFolder)
    If oFso.FolderExists(sFolder) Then 
        If Not fDetectOnly Then
            LogOnly "  Request SmartDelete for folder: " & sFolder
            SmartDeleteFolderEx sFolder
        Else
            LogOnly "  Simulate request SmartDelete for folder: " & sFolder
        End If
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then 
        If Not fDetectOnly Then 
            LogOnly "Request SmartDelete for folder: " & Wow64Folder(sFolder)
            SmartDeleteFolderEx Wow64Folder(sFolder)
        Else
            LogOnly "Simulate request SmartDelete for folder: " & Wow64Folder(sFolder)
        End If
    End If
End Sub 'SmartDeleteFolder
'=======================================================================================================

'Executes the folder delete operation
Sub SmartDeleteFolderEx(sFolder)
    Dim Folder
    
    On Error Resume Next
    DeleteFolder sFolder : CheckError "SmartDeleteFolderEx"
    On Error Goto 0
    Set Folder = oFso.GetFolder(oFso.GetParentFolderName(sFolder))
    If (Folder.Subfolders.Count = 0) AND (Folder.Files.Count = 0) Then SmartDeleteFolderEx(Folder.Path)
End Sub 'SmartDeleteFolderEx
'=======================================================================================================

'Adds the folder structure to the 'KeepFolder' dictionary
Sub AddKeepFolder(sPath)

    Dim Folder

    'Ensure trailing "\"
    sPath = LCase(sPath) & "\"
    While InStr(sPath,"\\")>0
        sPath = Replace(sPath,"\\","\")
    Wend

    If NOT dicKeepFolder.Exists (sPath) Then
        dicKeepFolder.Add sPath,sPath
    Else
        Exit Sub
    End If
    sPath = LCase(oFso.GetParentFolderName(sPath)) & "\"
    If oFso.FolderExists(sPath) Then AddKeepFolder(sPath)
End Sub
'=======================================================================================================

'Handles additional folder-path operations on 64 bit environments
Function Wow64Folder(sFolder)
    If LCase(Left(sFolder,Len(sWinDir & "\System32"))) = LCase(sWinDir & "\System32") Then 
        Wow64Folder = sWinDir & "\syswow64" & Right(sFolder,Len(sFolder)-Len(sSys32Dir))
    ElseIf LCase(Left(sFolder,Len(sProgramFiles))) = LCase(sProgramFiles) Then 
        Wow64Folder = sProgramFilesX86 & Right(sFolder,Len(sFolder)-Len(sProgramFiles))
    Else
        Wow64Folder = "?" 'Return invalid string to ensure the folder cannot exist
    End If
End Function 'Wow64Folder
'=======================================================================================================

Function HiveString(hDefKey)
    On Error Resume Next
    Select Case hDefKey
        Case HKCR : HiveString = "HKEY_CLASSES_ROOT"
        Case HKCU : HiveString = "HKEY_CURRENT_USER"
        Case HKLM : HiveString = "HKEY_LOCAL_MACHINE"
        Case HKU  : HiveString = "HKEY_USERS"
        Case Else : HiveString = hDefKey
    End Select
End Function
'=======================================================================================================

Function RegKeyExists(hDefKey,sSubKeyName)
    Dim arrKeys
    RegKeyExists = False
    If oReg.EnumKey(hDefKey,sSubKeyName,arrKeys) = 0 Then RegKeyExists = True
End Function
'=======================================================================================================

Function RegValExists(hDefKey,sSubKeyName,sName)
    Dim arrValueTypes, arrValueNames
    Dim i

    RegValExists = False
    If Not RegKeyExists(hDefKey,sSubKeyName) Then Exit Function
    If sName = "" Then
        RegValExists = True
        Exit Function
    End If
    If oReg.EnumValues(hDefKey,sSubKeyName,arrValueNames,arrValueTypes) = 0 AND IsArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    End If 'oReg.EnumValues
End Function
'=======================================================================================================

'Read the value of a given registry entry
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim Item
    Dim arrValues
    
    Select Case UCase(sType)
        Case "1","REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "2","REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "7","REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey,sSubKeyName,sName,arrValues)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,arrValues)
            If RetVal = 0 Then sValue = Join(arrValues,chr(34))
        
        Case "4","REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then 
                RetVal = oReg.GetDWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
            End If
        
        Case "3","REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case "11","REG_QWORD"
            RetVal = oReg.GetQWORDValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        
        Case Else
            RetVal = -1
    End Select 'sValue
    
    RegReadValue = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

'Enumerate a registry key to return all values
Function RegEnumValues(hDefKey,sSubKeyName,arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    If f64 Then
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames32,arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrNames64,arrTypes64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32,"\") & "\" & Join(arrNames64,"\")),"\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32,"\") & "\" & Join(arrTypes64,"\")),"\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey,sSubKeyName,arrNames,arrTypes)
    End If 'f64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'=======================================================================================================

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey,sSubKeyName,arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If f64 Then
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey,Wow64Key(hDefKey, sSubKeyName),arrKeys64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32,"\") & "\" & Join(arrKeys64,"\")),"\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey,sSubKeyName,arrKeys)
    End If 'f64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'=======================================================================================================

'Wrapper around oReg.DeleteValue to handle 64 bit
Sub RegDeleteValue(hDefKey, sSubKeyName, sName, fRegMultiSZ)
    Dim sWow64Key,sRealName
    Dim iRetVal
    
    sRealName = sName
    If UCase(sName) = "(DEFAULT)" Then sRealName = ""

    If dicKeepReg.Exists(LCase(sSubKeyName & sName)) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element. Remaining applications will need a repair!"
        End If
    End If
    If f64 Then
        If dicKeepReg.Exists(LCase(Wow64Key(hDefKey, sSubKeyName) & sName)) Then
            If NOT fForce Then
                LogOnly " - Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
                Exit Sub
            Else
                LogOnly " - Enforced delete of still required keypath element. Remaining applications will need a repair!"
            End If
        End If
    End If

    If RegValExists(hDefKey,sSubKeyName,sRealName) Then
        On Error Resume Next
        If RegReadValue(hDefKey,sSubKeyName,sName,sValue,"REG_MULTI_SZ") Then
            LogOnly " - Disallowing unsafe delete of REG_MULTI_SZ: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
            Exit Sub
        End If
        If Not fDetectOnly Then 
            LogOnly " - Delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName & " -> " & sName
            iRetVal = 0
            iRetVal = oReg.DeleteValue(hDefKey, sSubKeyName, sRealName)
            CheckError "RegDeleteValue"
            If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
        Else
            LogOnly " - Simulate delete registry value: " & HiveString(hDefKey) & "\" & sSubKeyName & " -> " & sName
        End If
        On Error Goto 0
    End If 'RegValExists
    If f64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegValExists(hDefKey,sWow64Key,sRealName) Then
            On Error Resume Next
            If RegReadValue(hDefKey,sSubKeyName,sName,sValue,"REG_MULTI_SZ") Then
                LogOnly " - Disallowing unsafe delete of REG_MULTI_SZ: " & HiveString(hDefKey) & "\" & sSubKeyName & sName
                Exit Sub
            End If
            If Not fDetectOnly Then 
            LogOnly " - Delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key & " -> " & sName
                iRetVal = 0
                iRetVal = oReg.DeleteValue(hDefKey, sWow64Key, sRealName)
                CheckError "RegDeleteValue"
                If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
            Else
                LogOnly " - Simulate delete registry value: " & HiveString(hDefKey) & "\" & sWow64Key & " -> " & sName
            End If
            On Error Goto 0
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteValue
'=======================================================================================================

'Wrappper around RegDeleteKeyEx to handle 64bit scenrios
Sub RegDeleteKey(hDefKey, sSubKeyName)
    Dim sWow64Key
    
    'Ensure trailing "\"
    sSubKeyName = sSubKeyName & "\"
    While InStr(sSubKeyName,"\\")>0
        sSubKeyName = Replace(sSubKeyName,"\\","\")
    Wend

    If dicKeepReg.Exists(LCase(sSubKeyName)) Then
        If NOT fForce Then
            LogOnly " - Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName
            Exit Sub
        Else
            LogOnly " - Enforced delete of still required keypath element. Remaining applications will need a repair!"
        End If
    End If
    If f64 Then
        If dicKeepReg.Exists(LCase(Wow64Key(hDefKey, sSubKeyName))) Then
            If NOT fForce Then
                LogOnly " - Disallowing the delete of still required keypath element: " & HiveString(hDefKey) & "\" & sSubKeyName
                Exit Sub
            Else
                LogOnly " - Enforced delete of still required keypath element. Remaining applications will need a repair!"
            End If
        End If
    End If
    
    If Len(sSubKeyName) > 1 Then
        'Strip of trailing "\"
        sSubKeyName = Left(sSubKeyName,Len(sSubKeyName)-1)
    End If
    
    If RegKeyExists(hDefKey, sSubKeyName) Then
        If Not fDetectOnly Then
            LogOnly " - Delete registry key: " & HiveString(hDefKey) & "\" & sSubKeyName
            On Error Resume Next
            RegDeleteKeyEx hDefKey, sSubKeyName
            On Error Goto 0
        Else
            LogOnly " - Simulate delete registry key: " & HiveString(hDefKey) & "\" & sSubKeyName
        End If
    End If 'RegKeyExists
    If f64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegKeyExists(hDefKey,sWow64Key) Then
            If Not fDetectOnly Then
                LogOnly " - Delete registry key: " & HiveString(hDefKey) & "\" & sWow64Key
                On Error Resume Next
                RegDeleteKeyEx hDefKey, sWow64Key
                On Error Goto 0
            Else
                LogOnly " - Simulate delete registry key: " & HiveString(hDefKey) & "\" & sWow64Key
            End If
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteKey
'=======================================================================================================

'Recursively delete a registry structure
Sub RegDeleteKeyEx(hDefKey, sSubKeyName) 
    Dim arrSubkeys
    Dim sSubkey
    Dim iRetVal

    On Error Resume Next
    oReg.EnumKey hDefKey, sSubKeyName, arrSubkeys
    If IsArray(arrSubkeys) Then 
        For Each sSubkey In arrSubkeys 
            RegDeleteKeyEx hDefKey, sSubKeyName & "\" & sSubkey 
        Next 
    End If 
    If Not fDetectOnly Then 
        iRetVal = 0
        iRetVal = oReg.DeleteKey(hDefKey,sSubKeyName)
        If NOT (iRetVal=0) Then LogOnly "     Delete failed. Return value: "&iRetVal
    End If
End Sub 'RegDeleteKeyEx
'=======================================================================================================

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos

    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case HKLM
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
        
    End Select 'hDefKey
End Function 'Wow64Key
'=======================================================================================================

'Remove duplicate entries from a one dimensional array
Function RemoveDuplicates(Array)
    Dim Item
    Dim oDic
    
    Set oDic = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not oDic.Exists(Item) Then oDic.Add Item,Item
    Next 'Item
    RemoveDuplicates = oDic.Keys
End Function 'RemoveDuplicates
'=======================================================================================================

'End running instances of setup
Sub EndCurrentInstalls ()
    Dim Processes, Process
    Dim iRet

    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name like '%setup%' OR Name like '%install%'")
    For Each Process in Processes
        If fEndCurrentInstalls Then
            Log " - End process " & Process.Name
            iRet = Process.Terminate()
            CheckError "EndCurrentInstalls: " & Process.Name
        Else
            Log " - Skip termination of process: " & Process.Name
        End If
    Next 'Process
    StopService "msiserver"
End Sub 'EndCurrentInstalls
'=======================================================================================================

'Uses WMI to stop a service
Function StopService(sService)
    Dim Services, Service
    Dim sQuery
    Dim iRet

    On Error Resume Next
    
    iRet = 0
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    'Stop the service
    For Each Service in Services
        If UCase(Service.State) = "STARTED" Then iRet = Service.StopService
        If UCase(Service.State) = "RUNNING" Then iRet = Service.StopService

    Next 'Service
    StopService = (iRet = 0)
End Function 'StopService
'=======================================================================================================

'Delete a service
Sub DeleteService(sService)
    Dim Services, Service, Processes, Process
    Dim sQuery, sStates
    Dim iRet
    
    On Error Resume Next
    
    sStates = "STARTED;RUNNING"
    sQuery = "Select * From Win32_Service Where Name='" & sService & "'"
    Set Services = oWmiLocal.Execquery(sQuery)
    
    'Stop and delete the service
    For Each Service in Services
        Log " Found service " & sService & " in state " & Service.State
        If InStr(sStates,UCase(Service.State))>0 Then iRet = Service.StopService()
        'Ensure no more instances of the service are running
        Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name='" & sService & ".exe'")
        For Each Process in Processes
            iRet = Process.Terminate()
        Next 'Process
        If Not fDetectOnly Then 
            Log " - Deleting Service -> " & sService
            iRet = Service.Delete()
        Else
            Log " - Simulate deleting Service -> " & sService
        End If
    Next 'Service
    Set Services = Nothing
    Err.Clear

End Sub 'DeleteService
'=======================================================================================================

'Translation for setup.exe error codes
Function SetupRetVal(RetVal)
    Select Case RetVal
        Case 0 : SetupRetVal = "Success"
        Case 30001,1 : SetupRetVal = "AbstractMethod"
        Case 30002,2 : SetupRetVal = "ApiProhibited"
        Case 30003,3  : SetupRetVal = "AlreadyImpersonatingAUser"
        Case 30004,4 : SetupRetVal = "AlreadyInitialized"
        Case 30005,5 : SetupRetVal = "ArgumentNullException"
        Case 30006,6 : SetupRetVal = "AssertionFailed"
        Case 30007,7 : SetupRetVal = "CABFileAddFailed"
        Case 30008,8 : SetupRetVal = "CommandFailed"
        Case 30009,9 : SetupRetVal = "ConcatenationFailed"
        Case 30010,10 : SetupRetVal = "CopyFailed"
        Case 30011,11 : SetupRetVal = "CreateEventFailed"
        Case 30012,12 : SetupRetVal = "CustomizationPatchNotFound"
        Case 30013,13 : SetupRetVal = "CustomizationPatchNotApplicable"
        Case 30014,14 : SetupRetVal = "DuplicateDefinition"
        Case 30015,15 : SetupRetVal = "ErrorCodeOnly - Passthrough for Win32 error"
        Case 30016,16 : SetupRetVal = "ExceptionNotThrown"
        Case 30017,17 : SetupRetVal = "FailedToImpersonateUser"
        Case 30018,18 : SetupRetVal = "FailedToInitializeFlexDataSource"
        Case 30019,19 : SetupRetVal = "FailedToStartClassFactories"
        Case 30020,20 : SetupRetVal = "FileNotFound"
        Case 30021,21 : SetupRetVal = "FileNotOpen"
        Case 30022,22 : SetupRetVal = "FlexDialogAlreadyInitialized"
        Case 30023,23 : SetupRetVal = "HResultOnly - Passthrough for HRESULT errors"
        Case 30024,24 : SetupRetVal = "HWNDNotFound"
        Case 30025,25 : SetupRetVal = "IncompatibleCacheAction"
        Case 30026,26 : SetupRetVal = "IncompleteProductAddOns"
        Case 30027,27 : SetupRetVal = "InstalledProductStateCorrupt"
        Case 30028,28 : SetupRetVal = "InsufficientBuffer"
        Case 30029,29 : SetupRetVal = "InvalidArgument"
        Case 30030,30 : SetupRetVal = "InvalidCDKey"
        Case 30031,31 : SetupRetVal = "InvalidColumnType"
        Case 30032,31 : SetupRetVal = "InvalidConfigAddLanguage"
        Case 30033,33 : SetupRetVal = "InvalidData"
        Case 30034,34 : SetupRetVal = "InvalidDirectory"
        Case 30035,35 : SetupRetVal = "InvalidFormat"
        Case 30036,36 : SetupRetVal = "InvalidInitialization"
        Case 30037,37 : SetupRetVal = "InvalidMethod"
        Case 30038,38 : SetupRetVal = "InvalidOperation"
        Case 30039,39 : SetupRetVal = "InvalidParameter"
        Case 30040,40 : SetupRetVal = "InvalidProductFromARP"
        Case 30041,41 : SetupRetVal = "InvalidProductInConfigXml"
        Case 30042,42 : SetupRetVal = "InvalidReference"
        Case 30043,43 : SetupRetVal = "InvalidRegistryValueType"
        Case 30044,44 : SetupRetVal = "InvalidXMLProperty"
        Case 30045,45 : SetupRetVal = "InvalidMetadataFile"
        Case 30046,46 : SetupRetVal = "LogNotInitialized"
        Case 30047,47 : SetupRetVal = "LogAlreadyInitialized"
        Case 30048,48 : SetupRetVal = "MissingXMLNode"
        Case 30049,49 : SetupRetVal = "MsiTableNotFound"
        Case 30050,50 : SetupRetVal = "MsiAPICallFailure"
        Case 30051,51 : SetupRetVal = "NodeNotOfTypeElement"
        Case 30052,52 : SetupRetVal = "NoMoreGraceBoots"
        Case 30053,53 : SetupRetVal = "NoProductsFound"
        Case 30054,54 : SetupRetVal = "NoSupportedCulture"
        Case 30055,55 : SetupRetVal = "NotYetImplemented"
        Case 30056,56 : SetupRetVal = "NotAvailableCulture"
        Case 30057,57 : SetupRetVal = "NotCustomizationPatch"
        Case 30058,58 : SetupRetVal = "NullReference"
        Case 30059,59 : SetupRetVal = "OCTPatchForbidden"
        Case 30060,60 : SetupRetVal = "OCTWrongMSIDll"
        Case 30061,61 : SetupRetVal = "OutOfBoundsIndex"
        Case 30062,62 : SetupRetVal = "OutOfDiskSpace"
        Case 30063,63 : SetupRetVal = "OutOfMemory"
        Case 30064,64 : SetupRetVal = "OutOfRange"
        Case 30065,65 : SetupRetVal = "PatchApplicationFailure"
        Case 30066,66 : SetupRetVal = "PreReqCheckFailure"
        Case 30067,67 : SetupRetVal = "ProcessAlreadyStarted"
        Case 30068,68 : SetupRetVal = "ProcessNotStarted"
        Case 30069,69 : SetupRetVal = "ProcessNotFinished"
        Case 30070,70 : SetupRetVal = "ProductAlreadyDefined"
        Case 30071,71 : SetupRetVal = "ResourceAlreadyTracked"
        Case 30072,72 : SetupRetVal = "ResourceNotFound"
        Case 30073,73 : SetupRetVal = "ResourceNotTracked"
        Case 30074,74 : SetupRetVal = "SQLAlreadyConnected"
        Case 30075,75 : SetupRetVal = "SQLFailedToAllocateHandle"
        Case 30076,76 : SetupRetVal = "SQLFailedToConnect"
        Case 30077,77 : SetupRetVal = "SQLFailedToExecuteStatement"
        Case 30078,78 : SetupRetVal = "SQLFailedToRetrieveData"
        Case 30079,79 : SetupRetVal = "SQLFailedToSetAttribute"
        Case 30080,80 : SetupRetVal = "StorageNotCreated"
        Case 30081,81 : SetupRetVal = "StreamNameTooLong"
        Case 30082,82 : SetupRetVal = "SystemError"
        Case 30083,83 : SetupRetVal = "ThreadAlreadyStarted"
        Case 30084,84 : SetupRetVal = "ThreadNotStarted"
        Case 30085,85 : SetupRetVal = "ThreadNotFinished"
        Case 30086,86 : SetupRetVal = "TooManyProducts"
        Case 30087,87 : SetupRetVal = "UnexpectedXMLNodeType"
        Case 30088,88 : SetupRetVal = "UnexpectedError"
        Case 30089,89 : SetupRetVal = "Unitialized"
        Case 30090,90 : SetupRetVal = "UserCancel"
        Case 30091,91 : SetupRetVal = "ExternalCommandFailed"
        Case 30092,92 : SetupRetVal = "SPDatabaseOverSize"
        Case 30093,93 : SetupRetVal = "IntegerTruncation"
        'msiexec return values
        Case 1259 : SetupRetVal = "APPHELP_BLOCK"
        Case 1601 : SetupRetVal = "INSTALL_SERVICE_FAILURE"
        Case 1602 : SetupRetVal = "INSTALL_USEREXIT"
        Case 1603 : SetupRetVal = "INSTALL_FAILURE"
        Case 1604 : SetupRetVal = "INSTALL_SUSPEND"
        Case 1605 : SetupRetVal = "UNKNOWN_PRODUCT"
        Case 1606 : SetupRetVal = "UNKNOWN_FEATURE"
        Case 1607 : SetupRetVal = "UNKNOWN_COMPONENT"
        Case 1608 : SetupRetVal = "UNKNOWN_PROPERTY"
        Case 1609 : SetupRetVal = "INVALID_HANDLE_STATE"
        Case 1610 : SetupRetVal = "BAD_CONFIGURATION"
        Case 1611 : SetupRetVal = "INDEX_ABSENT"
        Case 1612 : SetupRetVal = "INSTALL_SOURCE_ABSENT"
        Case 1613 : SetupRetVal = "INSTALL_PACKAGE_VERSION"
        Case 1614 : SetupRetVal = "PRODUCT_UNINSTALLED"
        Case 1615 : SetupRetVal = "BAD_QUERY_SYNTAX"
        Case 1616 : SetupRetVal = "INVALID_FIELD"
        Case 1618 : SetupRetVal = "INSTALL_ALREADY_RUNNING"
        Case 1619 : SetupRetVal = "INSTALL_PACKAGE_OPEN_FAILED"
        Case 1620 : SetupRetVal = "INSTALL_PACKAGE_INVALID"
        Case 1621 : SetupRetVal = "INSTALL_UI_FAILURE"
        Case 1622 : SetupRetVal = "INSTALL_LOG_FAILURE"
        Case 1623 : SetupRetVal = "INSTALL_LANGUAGE_UNSUPPORTED"
        Case 1624 : SetupRetVal = "INSTALL_TRANSFORM_FAILURE"
        Case 1625 : SetupRetVal = "INSTALL_PACKAGE_REJECTED"
        Case 1626 : SetupRetVal = "FUNCTION_NOT_CALLED"
        Case 1627 : SetupRetVal = "FUNCTION_FAILED"
        Case 1628 : SetupRetVal = "INVALID_TABLE"
        Case 1629 : SetupRetVal = "DATATYPE_MISMATCH"
        Case 1630 : SetupRetVal = "UNSUPPORTED_TYPE"
        Case 1631 : SetupRetVal = "CREATE_FAILED"
        Case 1632 : SetupRetVal = "INSTALL_TEMP_UNWRITABLE"
        Case 1633 : SetupRetVal = "INSTALL_PLATFORM_UNSUPPORTED"
        Case 1634 : SetupRetVal = "INSTALL_NOTUSED"
        Case 1635 : SetupRetVal = "PATCH_PACKAGE_OPEN_FAILED"
        Case 1636 : SetupRetVal = "PATCH_PACKAGE_INVALID"
        Case 1637 : SetupRetVal = "PATCH_PACKAGE_UNSUPPORTED"
        Case 1638 : SetupRetVal = "PRODUCT_VERSION"
        Case 1639 : SetupRetVal = "INVALID_COMMAND_LINE"
        Case 1640 : SetupRetVal = "INSTALL_REMOTE_DISALLOWED"
        Case 1641 : SetupRetVal = "SUCCESS_REBOOT_INITIATED"
        Case 1642 : SetupRetVal = "PATCH_TARGET_NOT_FOUND"
        Case 1643 : SetupRetVal = "PATCH_PACKAGE_REJECTED"
        Case 1644 : SetupRetVal = "INSTALL_TRANSFORM_REJECTED"
        Case 1645 : SetupRetVal = "INSTALL_REMOTE_PROHIBITED"
        Case 1646 : SetupRetVal = "PATCH_REMOVAL_UNSUPPORTED"
        Case 1647 : SetupRetVal = "UNKNOWN_PATCH"
        Case 1648 : SetupRetVal = "PATCH_NO_SEQUENCE"
        Case 1649 : SetupRetVal = "PATCH_REMOVAL_DISALLOWED"
        Case 1650 : SetupRetVal = "INVALID_PATCH_XML"
        Case 3010 : SetupRetVal = "SUCCESS_REBOOT_REQUIRED"
        Case Else : SetupRetVal = "Unknown Return Value"
    End Select
End Function 'SetupRetVal
'=======================================================================================================

Function GetProductID(sProdID)
        Dim sReturn
        
        Select Case sProdId
        
        Case "000F" : sReturn = "MONDO"
        Case "0010" : sReturn = "WEBFLDRS"
        Case "0011" : sReturn = "PROPLUS"
        Case "0012" : sReturn = "STANDARD"
        Case "0013" : sReturn = "BASIC"
        Case "0014" : sReturn = "PRO"
        Case "0015" : sReturn = "ACCESS"
        Case "0016" : sReturn = "EXCEL"
        Case "0017" : sReturn = "SharePointDesigner"
        Case "0018" : sReturn = "PowerPoint"
        Case "0019" : sReturn = "Publisher"
        Case "001A" : sReturn = "Outlook"
        Case "001B" : sReturn = "Word"
        Case "001C" : sReturn = "AccessRuntime"
        Case "001F" : sReturn = "Proof"
        Case "0020" : sReturn = "O2007CNV"
        Case "0021" : sReturn = "VisualWebDeveloper"
        Case "0026" : sReturn = "ExpressionWeb"
        Case "0029" : sReturn = "Excel"
        Case "002A" : sReturn = "Office64"
        Case "002B" : sReturn = "Word"
        Case "002C" : sReturn = "Proofing"
        Case "002E" : sReturn = "Ultimate"
        Case "002F" : sReturn = "HomeAndStudent"
        Case "0028" : sReturn = "IME"
        Case "0030" : sReturn = "Enterprise"
        Case "0031" : sReturn = "ProfessionalHybrid"
        Case "0033" : sReturn = "Personal"
        Case "0035" : sReturn = "ProfessionalHybrid"
        Case "0037" : sReturn = "PowerPoint"
        Case "0038" : sReturn = "OlTimeZoneTool"
        Case "003A" : sReturn = "PrjStd"
        Case "003B" : sReturn = "PrjPro"
        Case "003D" : sReturn = "SINGLEIMAGE"
        Case "0043" : sReturn = "OFFICE32"
        Case "0044" : sReturn = "InfoPath"
        Case "0045" : sReturn = "XWEB"
        Case "0048" : sReturn = "OLC"
        Case "0049" : sReturn = "ACADEMIC"
        Case "004A" : sReturn = "OWC11"
        Case "0051" : sReturn = "VISPRO"
        Case "0052" : sReturn = "VisView"
        Case "0053" : sReturn = "VisStd"
        Case "0054" : sReturn = "VisMUI"
        Case "0055" : sReturn = "VisMUI"
        Case "0057" : sReturn = "VISIO"
        Case "0061" : sReturn = "CLICK2RUN"
        Case "0062" : sReturn = "CLICK2RUN"
        Case "0066" : sReturn = "CLICK2RUN"
        Case "006C" : sReturn = "CLICK2RUN"
        Case "006D" : sReturn = "CLICK2RUN"
        Case "006E" : sReturn = "Shared"
        Case "006F" : sReturn = "OFFICE"
        Case "0070" : sReturn = "OOBE"
        Case "0074" : sReturn = "STARTER"
        Case "007A" : sReturn = "OLC" 'Outlook Connector
        Case "007C" : sReturn = "OSCFB" 'Outlook Social Connector for FaceBook
        Case "007D" : sReturn = "OSCWL" 'Outlook Social Connector for Windows Live Messenger
        Case "007F" : sReturn = "OLC" 'Outlook Social Connector
        Case "008A" : sReturn = "RecentDocs"
        Case "008B" : sReturn = "SmallBusinessBasics"
        Case "00A1" : sReturn = "ONENOTE"
        Case "00A3" : sReturn = "OneNoteHomeStudent"
        Case "00A4" : sReturn = "OWC11"
        Case "00A7" : sReturn = "CPAO"
        Case "00A9" : sReturn = "InterConnect"
        Case "00AF" : sReturn = "PPtView"
        Case "00B0" : sReturn = "ExPdf"
        Case "00B1" : sReturn = "ExXps"
        Case "00B2" : sReturn = "ExPdfXps"
        Case "00B4" : sReturn = "PrjMUI"
        Case "00B5" : sReturn = "PrjtMUI"
        Case "00B9" : sReturn = "AER"
        Case "00BA" : sReturn = "Groove"
        Case "00CA" : sReturn = "SmallBusiness"
        Case "00E0" : sReturn = "Outlook"
        Case "00D1" : sReturn = "ACE"
        Case "0100" : sReturn = "OfficeMUI"
        Case "0101" : sReturn = "OfficeXMUI"
        Case "0103" : sReturn = "PTK"
        Case "0114" : sReturn = "GrooveSetupMetadata"
        Case "0115" : sReturn = "SharedSetupMetadata"
        Case "0116" : sReturn = "SharedSetupMetadata"
        Case "0117" : sReturn = "AccessSetupMetadata"
        Case "011A" : sReturn = "SendASmile"
        Case "011D" : sReturn = "ProPlusSubscription"
        Case "011F" : sReturn = "OLConnect"
        Case "0126" : sReturn = "WWLIBCXM"
        
        Case "1014" : sReturn = "STS"
        Case "1015" : sReturn = "WSSMUI"
        Case "1032" : sReturn = "PJSVRAPP"
        Case "104B" : sReturn = "SPS"
        Case "104E" : sReturn = "SPSMUI"
        Case "107F" : sReturn = "OSrv"
        Case "1080" : sReturn = "OSrv"
        Case "1088" : sReturn = "lpsrvwfe"
        Case "10D7" : sReturn = "IFS"
        Case "10D8" : sReturn = "IFSMUI"
        Case "10EB" : sReturn = "DLCAPP"
        Case "10F5" : sReturn = "XLSRVAPP"
        Case "10F6" : sReturn = "XlSrvWFE"
        Case "10F7" : sReturn = "DLC"
        Case "10F8" : sReturn = "SlSrvMui"
        Case "10FB" : sReturn = "OSrchWFE"
        Case "10FC" : sReturn = "OSRCHAPP"
        Case "10FD" : sReturn = "OSrchMUI"
        Case "1103" : sReturn = "DLC"
        Case "1104" : sReturn = "LHPSRV"
        Case "1105" : sReturn = "PIA"
        Case "1106" : sReturn = "GRVMGMTSRV"
        Case "1109" : sReturn = "GSERVERRELAY"
        Case "110D" : sReturn = "OSERVER"
        Case "110F" : sReturn = "PSERVER"
        Case "1110" : sReturn = "WSS"
        Case "1121" : sReturn = "SPSSDK"
        Case "1122" : sReturn = "SPSDev"
        Case "1163" : sReturn = "SCC" 'SharePoint Client Components
        Case Else : sReturn = sProdID
        
        End Select 'sProdId
    GetProductID = sReturn
End Function 'GetProductID

'-------------------------------------------------------------------------------
'   LogH
'
'   Write a header log string to the log file
'-------------------------------------------------------------------------------
Sub LogH (sLog)
    LogStream.WriteLine ""
    sLog = sLog & vbCrLf & String(Len(sLog), "=")
    If NOT fQuiet AND fCScript Then wscript.echo ""
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'Logh

'-------------------------------------------------------------------------------
'   LogH1
'
'   Write a header log string to the log file
'-------------------------------------------------------------------------------
Sub LogH1 (sLog)
    LogStream.WriteLine ""
    sLog = sLog & vbCrLf & String(Len(sLog), "-")
    If NOT fQuiet AND fCScript Then wscript.echo ""
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine sLog
End Sub 'LogH1

'-------------------------------------------------------------------------------
'   LogH2
'
'   Write w/o indent Cmd window and the log file
'-------------------------------------------------------------------------------
Sub LogH2 (sLog)
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    LogStream.WriteLine ""
    LogStream.WriteLine sLog
End Sub 'LogH2

'-------------------------------------------------------------------------------
'   Log
'
'   Echos the log string to the Cmd window and the log file
'-------------------------------------------------------------------------------
Sub Log (sLog)
    If NOT fQuiet AND fCScript Then wscript.echo sLog
    If sLog = "" Then
        LogStream.WriteLine
    Else
        LogStream.WriteLine "   " & Time & ": " & sLog
    End If
End Sub 'Log

'-------------------------------------------------------------------------------
'   LogOnly
'
'   Commits the log string to the log file
'-------------------------------------------------------------------------------
Sub LogOnly (sLog)
    If sLog = "" Then
        LogStream.WriteLine
    Else
        LogStream.WriteLine "   " & Time & ": " & sLog
    End If
End Sub 'Log
'=======================================================================================================

Sub CheckError(sModule)
    If Err <> 0 Then 
        LogOnly "   " & Now & " - " & sModule & " - Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
               "; Err# (Dec): " & Err & "; Description : " & Err.Description
    End If 'Err = 0
    Err.Clear
End Sub
'=======================================================================================================

'Command line parser
Sub ParseCmdLine

    Dim iCnt, iArgCnt
    Dim arrArguments
    Dim sArg0, sArguments
    
    iArgCnt = Wscript.Arguments.Count
    If iArgCnt > 1 Then
        If wscript.Arguments(1) = "UAC" Then
            If wscript.arguments.count = 2 Then iArgCnt = 0
        End If
    End If

    sArguments = ""
    If iArgCnt = 0 Then
        If sDefault = "" Then
            'Create the log
            CreateLog
            Log "No argument specified. Preparing user prompt" & vbCrLf
            fPassive = False
            FindInstalledOProducts
            If dicInstalledSku.Count > 0 Then sDefault = Join(RemoveDuplicates(dicInstalledSku.Items),",") Else sDefault = "CLIENTALL"
            sDefault = InputBox("Enter a list of " & ONAME & " products to remove" & vbCrLf & vbCrLf & _
                    "Examples:" & vbCrLf & _
                    "CLIENTALL" & vbTab & "-> all Client products" & vbCrLf & _
                    "SERVER" & vbTab & "-> all Server products" & vbCrLf & _
                    "ALL" & vbTab & vbTab & "-> all Server & Client products" & vbCrLf & _
                    "ProPlus,PrjPro" & vbTab & "-> ProPlus and Project" & vbCrLf &_
                    "?" & vbTab & vbTab & "-> display Help", _
                    SCRIPTFILE & " - " & ONAME & " remover", _
                    sDefault)

            If IsEmpty(sDefault) Then 'User cancelled
                Log "User cancelled. CleanUp & Exit."
                'Undo temporary entries created in ARP
                TmpKeyCleanUp
                'wscript.quit 1602
                SetError ERROR_USERCANCEL
                ExitScript
            End If 'IsEmpty(sDefault)
            Log "Answer from prompt: " & sDefault & vbCrLf
        End If
        sDefault = Trim(UCase(Trim(Replace(sDefault, Chr(34), ""))))
        arrArguments = Split(Trim(sDefault), " ")
        If UBound(arrArguments) = -1 Then ReDim arrArguments(0)
    Else
        ReDim arrArguments(iArgCnt - 1)
        For iCnt = 0 To (iArgCnt - 1)
            arrArguments(iCnt) = UCase(Wscript.Arguments(iCnt))
            sArguments = sArguments & arrArguments(iCnt) & " "
        Next 'iCnt
    End If 'iArgCnt = 0

    'Handle the SKU list
    sArg0 = Replace(arrArguments(0), "/", "")
    If Left(sArg0, 1) = "-" Then sArg0 = Mid(sArg0, 2)

    Select Case UCase(sArg0)
    
    Case "?"
        ShowSyntax
    
    Case "ALL"
        fRemoveAll = True
        fRemoveOse = False
    
    Case "CLIENTSUITES"
        fRemoveCSuites = True
        fRemoveOse = False
    
    Case "CLIENTSTANDALONE"
        fRemoveCSingle = True
        fRemoveOse = False

    Case "CLIENTALL"
        fRemoveCSuites = True
        fRemoveCSingle = True
        fRemoveOse = False
    
    Case "SERVER"
        fRemoveSrv = True
        fRemoveOse = False

    Case "ALL,OSE"
        fRemoveAll = True
        fRemoveOse = True
    
    Case Else
        fRemoveAll = False
        fRemoveOse = False
        sSkuRemoveList = sArg0
    
    End Select
    
    For iCnt = 0 To UBound(arrArguments)

        Select Case arrArguments(iCnt)
        
        Case "?", "/?", "-?"
            ShowSyntax
        
        Case "/B", "/BYPASS"
            If UBound(arrArguments) > iCnt Then
                If InStr(arrArguments(iCnt + 1), "1") > 0 Then fBypass_Stage1 = True
                If InStr(arrArguments(iCnt + 1), "2") > 0 Then fBypass_Stage2 = True
                If InStr(arrArguments(iCnt + 1), "3") > 0 Then fBypass_Stage3 = True
                If InStr(arrArguments(iCnt + 1), "4") > 0 Then fBypass_Stage4 = True
            End If
        
        Case "/CLEARADDINREG", "/CLEARADDINSREG"
        	fClearAddinReg = True
        
        Case "/D", "/DELETEUSERSETTINGS"
            fKeepUser = False
        
        ' Option to ensure that other setup/install processes are terminated
        ' to avoid any conflict with the Office removal process
        Case "/ECI", "/ENDCURRENTINSTALLS"
            fEndCurrentInstalls = True
        
        Case "/FR", "/FASTREMOVE"
            fBypass_Stage1 = True
            fSkipSD = True
        
        Case "/F", "/FORCE"
            fForce = True
        
        Case "/K", "/KEEPUSERSETTINGS"
            fKeepUser = True
        
        Case "/KEEPSG", "/KEEPSOFTGRID"
            fKeepSG = True

        Case "/KEEPLYNC"
            fRemoveLync = False

        Case "/REMOVELYNC"
            fRemoveLync = True
        
        Case "/L", "/LOG"
            fLogInitialized = False
            If UBound(arrArguments) > iCnt Then
                If oFso.FolderExists(arrArguments(iCnt + 1)) Then 
                    sLogDir = arrArguments(iCnt + 1)
                Else
                    On Error Resume Next
                    oFso.CreateFolder(arrArguments(iCnt + 1))
                    If Err <> 0 Then sLogDir = sScrubDir Else sLogDir = arrArguments(iCnt + 1)
                End If
            End If
        
        Case "/NOCANCEL"
            fNoCancel = True
        
        Case "/NOREBOOT"
            fNoReboot = True
        
        Case "/NE", "/NOELEVATE"
            fNoElevate = True
        
        Case "/O", "/OSE"
            fRemoveOse = True
        
        Case "/P", "/PREVIEW", "/DETECTONLY"
            fDetectOnly = True
        
        Case "/PASSIVE", "/QB-"
            fPassive = True
        
        Case "/Q", "/QUIET"
            fQuiet = True
        
        Case "/QB"
            fQuiet = True
            fBasic = True

        Case "/QND"
            fBypass_Stage1 = True
            fBypass_Stage2 = True
            fBypass_Stage3 = True
            fRemoveOse = True
            fRemoveOspp = True
            fRemoveAll = True
            fSkipSD = True
            fForce = True
        
        Case "/R", "/RECONCILE"
            fTryReconcile = True
        
        Case "/REMOVEOSPP" , "/CLEANOSPP"
            fRemoveOspp = True
        
        Case "/RETERRORSUCCESS", "/RETURNERRORORSUCCESS", "/REOS"
            fReturnErrorOrSuccess = True

        Case "/S", "/SKIPSD", "/SKIPSHORTCUTDETECTION"
            fSkipSD = True
        
        Case "/SC", "/SCANCOMPONENTS"
            fBypass_Stage1 = False
        
        Case Else
        
        End Select
    Next 'iCnt
    If Not fLogInitialized Then CreateLog
    LogH2 "Arguments: " & sArguments & vbCrLf

End Sub 'ParseCmdLine
'=======================================================================================================

'-------------------------------------------------------------------------------
'   CreateLog
'
'   Create the removal log file
'-------------------------------------------------------------------------------
Sub CreateLog
    Dim DateTime
    Dim sLogName
    
    On Error Resume Next
    ' create the log file
    Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
    DateTime.SetVarDate Now, True
    sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    sLogName = sLogName &  "_" & Left(DateTime.Value, 14)
    sLogName = sLogName & "_ScrubLog.txt"
    Err.Clear
    Set LogStream = oFso.CreateTextFile(sLogName, True, True)
    If Err <> 0 Then 
        Err.Clear
        sLogDir = sScrubDir
        sLogName = sLogDir & "\" & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
        sLogName = sLogName &  "_" & Left(DateTime.Value, 14)
        sLogName = sLogName & "_ScrubLog.txt"
        Set LogStream = oFso.CreateTextFile(sLogName, True, True)
    End If
    On Error Goto 0

    LogH2 "Microsoft Customer Support Services - " & ONAME & " Removal Utility" & vbCrLf & vbCrLf & _
        "Version: " & vbTab & SCRIPTVERSION & vbCrLf & _
        "64 bit OS: " & vbTab & f64 & vbCrLf & _
        "Removal start: " & vbTab & Time & vbCrLf & _
        "OS Details: " & sOSinfo & vbCrLf
    fLogInitialized = True
End Sub 'CreateLog

'-------------------------------------------------------------------------------
'   RelaunchAs64Host
'
'   Relaunch self with 64 bit CScript host
'-------------------------------------------------------------------------------
Sub RelaunchAs64Host
    Dim Argument, sCmd
    Dim fQuietRelaunch

    fQuietRelaunch = False
    sCmd = Replace(LCase(wscript.Path), "syswow64", "sysnative") & "\cscript.exe " & Chr(34) & WScript.scriptFullName & Chr(34)
    If fQuiet Then fQuietRelaunch = True
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmd = sCmd  &  " " & chr(34) & Argument & chr(34)
            Select Case UCase(Argument)
            Case "/Q", "/QUIET"
                fQuietRelaunch = True
            End Select
        Next 'Argument
    End If
    sCmd = sCmd & " /ChangedHostBitness"
    If fQuietRelaunch Then
        sCmd = Replace (sCmd, "\cscript.exe", "\wscript.exe")
        Wscript.Quit CLng(oWShell.Run (sCmd, 0, True))
    Else
        Wscript.Quit CLng(oWShell.Run (sCmd, 1, True))
    End If

End Sub 'RelaunchAs64Host

'-------------------------------------------------------------------------------
'   RelaunchAsCScript
'
'   Relaunch self with Cscript as host
'-------------------------------------------------------------------------------
Sub RelaunchAsCScript
    Dim Argument
    Dim sCmdLine
    
    sCmdLine = "cmd.exe /c " & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    Log "Relaunching with CScript as host. Full command: " & sCmdLine
    Wscript.Quit CLng(oWShell.Run(sCmdLine, 1, True))

End Sub 'RelaunchAsCScript

'-------------------------------------------------------------------------------
'   RelaunchElevated
'
'   Relaunch the script with elevated permissions
'-------------------------------------------------------------------------------
Sub RelaunchElevated
    Dim Argument, Process, Processes
    Dim iParentProcessId, iSpawnedProcessId
    Dim sCmdLine, sRetValFile, sValue
    Dim oShell

    SetError ERROR_RELAUNCH
    ' Shell object for relaunch
    Set oShell = CreateObject("Shell.Application")
    ' build command line for relaunch
    sCmdLine = Chr(34) & WScript.ScriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            Select Case UCase(Argument)
            Case "/Q","/QUIET"
                'Don't try to relaunch in quiet mode
                Exit Sub
                SetError ERROR_ELEVATION_FAILED
            Case "UAC"
                'Already tried elevated relaunch
                SetError ERROR_ELEVATION_FAILED
                Exit Sub
            Case Else
                sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
            End Select
        Next 'Argument
    End If
    ' prep work to get the return value from the elevated process
    iParentProcessId = GetMyProcessId 
    
    ' launch the elevated instance
    oShell.ShellExecute "cscript.exe", sCmdLine & " /NoElevate UAC", "", "runas", 1
    ' get the process id of the spawned instance
    WScript.Sleep 500
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process WHERE ParentProcessId='" & iParentProcessId & "'")
    If Processes.Count > 0 Then
        For Each Process in Processes
		    iSpawnedProcessId = Process.ProcessId
		    Exit For
        Next 'Process
        ' monitor the tasklist to detect the end of the spawned process
        While oWmiLocal.ExecQuery("Select * From Win32_Process WHERE ProcessId='" & iSpawnedProcessId & "'").Count > 0
            WScript.Sleep 3000
        Wend
        ' get the return value from the file
        Wscript.Quit GetRetValFromFile
    End If
    ' elevation failed (user declined)
    SetError ERROR_ELEVATION_USERDECLINED
End Sub 'RelaunchElevated

'-------------------------------------------------------------------------------
'   GetMyProcessId
'
'   Returns the process id of the own process
'-------------------------------------------------------------------------------
Function GetMyProcessId()
    Dim iParentProcessId

    iParentProcessId = 0
    ' try to obtain from creating a new cscript instance
    On Error Resume Next
    iParentProcessId = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & oWShell.Exec("cscript.exe").ProcessId & "'").ParentProcessId
    On Error Goto 0
    If iParentProcessId > 0 Then
        ' succeeded to obtain the process id
        GetMyProcessId = iParentProcessId
        Exit Function
    End If

    ' failed to obtain the id from the creation of a new instance
    ' get it from enum of Win32_Process
    Dim Process, Processes
    Err.Clear
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process WHERE Name='cscript.exe' AND CommandLine like '%" & SCRIPTNAME & "%'")
    For Each Process in Processes
        iParentProcessId = Process.ProcessId
        Exit For
    Next
    GetMyProcessId = iParentProcessId
End Function 'GetMyProcessId


'-------------------------------------------------------------------------------
'   SetError
'
'   Set error bit(s) 
'-------------------------------------------------------------------------------
Sub SetError(ErrorBit)
    iError = iError OR ErrorBit
    Select Case ErrorBit
    Case ERROR_DCAF_FAILURE, ERROR_STAGE2, ERROR_ELEVATION_USERDECLINED, ERROR_ELEVATION, ERROR_SCRIPTINIT
        iError = iError OR ERROR_FAIL
    End Select
End Sub

'-------------------------------------------------------------------------------
'   ClearError
'
'   Unset error bit(s) 
'-------------------------------------------------------------------------------
Sub ClearError(ErrorBit)
    iError = iError AND (ERROR_ALL - ErrorBit)
    Select Case ErrorBit
    Case ERROR_ELEVATION_USERDECLINED, ERROR_ELEVATION, ERROR_SCRIPTINIT
        iError = iError AND (ERROR_ALL - ERROR_FAIL)
    End Select
End Sub

'-------------------------------------------------------------------------------
'   SetRetVal
'
'   Write return value to file
'-------------------------------------------------------------------------------
Sub SetRetVal(iError)
    Dim RetValFileStream
    
    'don't fail script execution if writing the return value to file fails
    On Error Resume Next 

    Set RetValFileStream = oFso.createTextFile(sScrubDir & "\" & RETVALFILE, True, True)
    RetValFileStream.Write iError
    RetValFileStream.Close
    On Error Goto 0
End Sub 'SetRetVal

'-------------------------------------------------------------------------------
'   GetRetValFromFile
'
'   Read return value from file.
'   Used to ensure return value can get obtained from an elevated process
'-------------------------------------------------------------------------------
Function GetRetValFromFile ()
    Dim RetValFileStream
    Dim iRetValFromFile

    On Error Resume Next 'don't fail script execution when getting the return value from file fails

    If oFso.FileExists(sScrubDir & "\" & RETVALFILE) Then
        Set RetValFileStream = oFso.OpenTextFile(sScrubDir & "\" & RETVALFILE, 1, False, -2)
        GetRetValFromFile = RetValFileStream.ReadAll
        RetValFileStream.Close
        Exit Function
    End If
    Err.Clear
    On Error Goto 0
    GetRetValFromFile = ERROR_UNKNOWN
End Function 'GetRetValFromFile

'-------------------------------------------------------------------------------
'   ShowSyntax
'
'   Show the expected syntax for the script usage
'-------------------------------------------------------------------------------
Sub ShowSyntax
    TmpKeyCleanUp
    Wscript.Echo vbCrLf & _
             SCRIPTFILE & " V " & SCRIPTVERSION & vbCrLf & _
             "Copyright (c) Microsoft Corporation. All Rights Reserved" & vbCrLf & vbCrLf & _
             "Microsoft Customer Support Services - " & ONAME & " Removal Utility" & vbCrLf & vbCrLf & _
             SCRIPTFILE & " helps to remove " & ONAME & " Server & Client products" & vbCrLf & vbCrLf & _
             "Usage:" & vbTab & SCRIPTFILE & " [List of config ProductIDs] [Options]" & vbCrLf & vbCrLf & _
             vbTab & "/?                          ' Displays this help"& vbCrLf &_
             vbTab & "/Log [LogfolderPath]        ' Custom folder for log files" & vbCrLf & _
             vbTab & "/SkipSD                     ' Skips the ShortcutDetection in local profiles" & vbCrLf & _
             vbTab & "/NoCancel                   ' Setup.exe and Msiexec.exe have no Cancel button" & vbCrLf &_
             vbTab & "/Quiet                      ' Script, Setup.exe and Msiexec.exe run quiet with no UI" & vbCrLf &_
             vbTab & "/Preview                    ' Run this script to preview what would get removed"& vbCrLf & vbCrLf & _
             "Examples:"& vbCrLf & _
             vbTab & SCRIPTFILE & " CLIENTSUITES    ' Remove all " & ONAME & " Client suites/products" & vbCrLf &_
             vbTab & SCRIPTFILE & " CLIENTALL       ' Remove all " & ONAME & " Client products" & vbCrLf &_
             vbTab & SCRIPTFILE & " SERVER          ' Remove all " & ONAME & " Server products" & vbCrLf &_
             vbTab & SCRIPTFILE & " ALL             ' Remove all " & ONAME & " Server & Client products" & vbCrLf &_
             vbTab & SCRIPTFILE & " ProPlus,PrjPro  ' Remove ProPlus and Project" & vbCrLf
    Wscript.Quit
End Sub 'ShowSyntax
'=======================================================================================================