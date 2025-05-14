' This adds the MSI uninstall string to the CarbonBlack reg key so that uninst.exe can launch MSI if needed


' Writes a REG_SZ value to the local computer's registry using WMI.
' Parameters:
'   RootKey - The registry hive (see http://msdn.microsoft.com/en-us/library/aa390788(VS.85).aspx for a list of possible values).
'   Key - The key that contains the desired value.
'   Value - The value that you want to set.
'   ValueStr - The value string
'   RegType - The registry bitness: 32 or 64.
'
Function WriteRegStr(RootKey, Key, ValueName, ValueStr, RegType)
    Dim oCtx, oLocator, oReg, oInParams, oOutParams, ret

    Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    oCtx.Add "__ProviderArchitecture", RegType

    Set oLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set oReg = oLocator.ConnectServer("", "root\default", "", "", , , , oCtx).Get("StdRegProv")
    ret = oReg.CreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\CarbonBlack")

    Set oInParams = oReg.Methods_("SetStringValue").InParameters
    oInParams.hDefKey = RootKey
    oInParams.sSubKeyName = Key
    oInParams.sValueName = ValueName
    oInParams.sValue = ValueStr

    Set oOutParams = oReg.ExecMethod_("SetStringValue", oInParams, , oCtx)

    WriteRegStr = oOutParams.Returnvalue

End Function

Function Is64BitPlatform()
	Dim WshShell
	Dim WshProcEnv
	Dim system_architecture

	Set WshShell =  CreateObject("WScript.Shell")
	Set WshProcEnv = WshShell.Environment("Process")

	process_architecture= WshProcEnv("PROCESSOR_ARCHITECTURE") 

	If process_architecture = "x86" Then    
			system_architecture= WshProcEnv("PROCESSOR_ARCHITEW6432")

			If system_architecture = ""  Then    
					system_architecture = "x86"
			End if    
	Else    
			system_architecture = process_architecture    
	End If

	If system_architecture = "x86" Then
      Is64BitPlatform = False
  Else
      Is64BitPlatform = True
  End If
End Function

Dim regType
Dim valueStr

If Is64BitPlatform() Then
    regType = 64
Else
    regType = 32
End If

Const HKEY_LOCAL_MACHINE = &H80000002
valueStr = Session.Property("CustomActionData")

Call WriteRegStr(HKEY_LOCAL_MACHINE, "SOFTWARE\CarbonBlack", "MSIProductCode", valueStr, regType)