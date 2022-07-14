' check_rdp_licenses Nagios plugin para servidores de licencias RDP.
' Copyright (C) 2022  Ramón Román Castro <ramonromancastro@gmail.com>
' 
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

Option Explicit

' -----------------------------------------------------------------------
' NAGIOS PLUGINS CODE
' -----------------------------------------------------------------------

Const VERSION = "0.10"
Const nagios_OK = 0
Const nagios_WARNING = 1
Const nagios_CRITICAL = 2
Const nagios_UNKNOWN = 3

Const plugin_ignoreTemporaryWarnings = true

Dim nagios_return_message:nagios_return_message = Array("OK", "WARNING", "CRITICAL", "UNKNOWN")
Dim nagios_return_code:nagios_return_code = nagios_OK
Dim nagios_message:nagios_message = ""
Dim nagios_perf:nagios_perf = ""
Dim nagios_threshold_warning:nagios_threshold_warning = 80
Dim nagios_threshold_critical:nagios_threshold_critical = 90
Dim nagios_threshold_regexp

Set nagios_threshold_regexp = New RegExp
nagios_threshold_regexp.IgnoreCase = True
nagios_threshold_regexp.Global = True
nagios_threshold_regexp.Pattern = "^(~|[0-9]+|@[0-9]+):?([0-9]*)$"

Function nagios_check_threshold(value)
	Dim objMatches, objMatch
	nagios_check_threshold = nagios_OK
End Function

Sub set_nagios_threshold_warning(threshold)
	If NOT nagios_threshold_regexp.Test(threshold) Then
		suffix_nagios_message "Invalid threshold format" 
		set_nagios_return_code nagios_UNKNOWN
		nagios_exit
	End If
	nagios_threshold_warning = threshold
End Sub

Sub set_nagios_threshold_critical(threshold)
	If NOT nagios_threshold_regexp.Test(threshold) Then
		suffix_nagios_message "Invalid threshold format" 
		set_nagios_return_code nagios_UNKNOWN
		nagios_exit
	End If
	nagios_threshold_critical = threshold
End Sub

Sub set_nagios_return_code(return_code)
	If (return_code = nagios_CRITICAL) Then
		nagios_return_code = nagios_CRITICAL
	ElseIf (return_code = nagios_WARNING) And (nagios_return_code <> nagios_CRITICAL) Then
		nagios_return_code = nagios_WARNING
	ElseIf (return_code = nagios_UNKNOWN) And (nagios_return_code <> nagios_CRITICAL) AND (nagios_return_code <> nagios_WARNING) Then
		nagios_return_code = nagios_UNKNOWN
	ElseIf (return_code = nagios_OK) And (nagios_return_code <> nagios_CRITICAL) AND (nagios_return_code <> nagios_WARNING) AND (nagios_return_code <> nagios_UNKNOWN) Then
		nagios_return_code = nagios_OK
	End If
End Sub

Function get_nagios_return_code
	get_nagios_return_code = nagios_return_code
End Function

Sub suffix_nagios_message(message)
	nagios_message = nagios_message & message
End Sub

Sub suffix_nagios_perf(perf)
	nagios_perf = nagios_perf & perf & " "
End Sub

Sub prefix_nagios_message(message)
	nagios_message = message & nagios_message
End Sub

Function nagios_exit()
	If (nagios_perf <> "") Then
		WScript.Echo nagios_return_message(nagios_return_code) & ": " & nagios_message & "|" & nagios_perf
	Else
		WScript.Echo nagios_return_message(nagios_return_code) & ": " & nagios_message
	End If
	WScript.Quit nagios_return_code
End Function

' -----------------------------------------------------------------------
' PLUGIN CODE
' -----------------------------------------------------------------------

Function ProductVersionID2String(value)
	Select Case value
		Case 0
			ProductVersionID2String = "Not supported"
		Case 1
			ProductVersionID2String = "Not supported"
		Case 2
			ProductVersionID2String = "Windows Server 2008"
		Case 3
			ProductVersionID2String = "Windows Server 2008 R2"
		Case 4
			ProductVersionID2String = "Windows Server 2012/Windows Server 2012 R2"
		Case 5
			ProductVersionID2String = "Windows Server 2016"
		Case 6
			ProductVersionID2String = "Windows Server 2019"
		Case Else
			ProductVersionID2String = "Unknown (" & value & ")"
		End Select
End Function

Function ProductType2String(value)
	Select Case value
		Case 0
			ProductType2String = "Per device"
		Case 1
			ProductType2String = "Per user"
		Case 2
			ProductType2String = "Not valid"
		Case Else
			ProductType2String = "Unknown (" & value & ")"
		End Select
End Function

Function KeyPackType2String(value)
	Select Case value
		Case 0
			KeyPackType2String = "Unknown"
		Case 1
			KeyPackType2String = "Retail"
		Case 2
			KeyPackType2String = "Volume"
		Case 3
			KeyPackType2String = "Concurrent"
		Case 4
			KeyPackType2String = "Temporary"
		Case 5
			KeyPackType2String = "Open"
		Case 6
			KeyPackType2String = "Not supported"
		Case Else
			KeyPackType2String = "Unknown (" & value & ")"
		End Select
End Function

Dim objWMIService, objWMIArray, objWMIObject
Dim obj_Win32_TSLicenseKeyPack_Issued, obj_Win32_TSLicenseKeyPack_Total, obj_Win32_TSLicenseKeyPack_Perf
Dim tmpStatus, tmpIndex, strKey, tmpPerf, tmpWarning, tmpCritical

On Error Resume Next

Set obj_Win32_TSLicenseKeyPack_Total = CreateObject("Scripting.Dictionary")
Set obj_Win32_TSLicenseKeyPack_Issued = CreateObject("Scripting.Dictionary")
Set obj_Win32_TSLicenseKeyPack_Perf = CreateObject("Scripting.Dictionary")

Set ObjWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
If Err.Number <> 0 Then
	suffix_nagios_message "Unable to connect to the RD License Server." 
	set_nagios_return_code nagios_UNKNOWN
	nagios_exit
End If

Set objWMIArray = ObjWMIService.ExecQuery("SELECT * FROM Win32_TSLicenseKeyPack WHERE KeyPackType <> 0 AND KeyPackType <> 6") 

For Each objWMIObject In objWMIArray
	tmpIndex = ProductVersionID2String(objWMIObject.ProductVersionID) & " " & KeyPackType2String(objWMIObject.KeyPackType) & " (" & ProductType2String(objWMIObject.ProductType) & ")"
	tmpPerf = objWMIObject.ProductVersionID & "_" & objWMIObject.KeyPackType & "_" & objWMIObject.ProductType
        If objWMIObject.TotalLicenses < 0 Then
          objWMIObject.TotalLicenses = 0
        End If
	If NOT obj_Win32_TSLicenseKeyPack_Total.Exists(tmpIndex) Then
		obj_Win32_TSLicenseKeyPack_Total.Add tmpIndex, objWMIObject.TotalLicenses
		obj_Win32_TSLicenseKeyPack_Issued.Add tmpIndex, objWMIObject.IssuedLicenses
		obj_Win32_TSLicenseKeyPack_Perf.Add tmpIndex, tmpPerf
	Else
		obj_Win32_TSLicenseKeyPack_Total.Item(tmpIndex) = obj_Win32_TSLicenseKeyPack_Total.Item(tmpIndex) + objWMIObject.TotalLicenses
		obj_Win32_TSLicenseKeyPack_Issued.Item(tmpIndex) = obj_Win32_TSLicenseKeyPack_Issued.Item(tmpIndex) + objWMIObject.IssuedLicenses
	End If
Next

For Each strKey in obj_Win32_TSLicenseKeyPack_Total.Keys
	tmpStatus = nagios_OK
    tmpWarning = 0
    tmpCritical = 0
	If (InStr(strKey,"Temporary") = 0) OR (InStr(strKey,"Temporary") > 0 AND NOT plugin_ignoreTemporaryWarnings) Then
		If obj_Win32_TSLicenseKeyPack_Total.Item(strKey) <= 0 AND obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) > 0 Then
			tmpStatus = nagios_CRITICAL
		ElseIf obj_Win32_TSLicenseKeyPack_Total.Item(strKey) <= 0 AND obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) <= 0 Then
			tmpStatus = nagios_OK
		Else
			tmpWarning = CInt(nagios_threshold_warning*obj_Win32_TSLicenseKeyPack_Total.Item(strKey)/100)
			tmpCritical = CInt(nagios_threshold_critical*obj_Win32_TSLicenseKeyPack_Total.Item(strKey)/100)
			If obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) > tmpCritical Then
				tmpStatus = nagios_CRITICAL
			ElseIf obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) > tmpWarning Then
				tmpStatus = nagios_WARNING
			End If
		End If
	End If
	suffix_nagios_message vbNewLine & nagios_return_message(tmpStatus) & ": " & strKey & ": " & obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) & "/" & obj_Win32_TSLicenseKeyPack_Total.Item(strKey) & " (issued/total)"
	suffix_nagios_perf "'" & obj_Win32_TSLicenseKeyPack_Perf.Item(strKey) & "'=" & obj_Win32_TSLicenseKeyPack_Issued.Item(strKey) & ";" & tmpWarning & ";" & tmpCritical & ";0;" & obj_Win32_TSLicenseKeyPack_Total.Item(strKey)
	set_nagios_return_code tmpStatus
Next

Select Case get_nagios_return_code
	Case nagios_OK
		prefix_nagios_message "All license key packs are ok"
	Case nagios_WARNING
		prefix_nagios_message "One or more license key packs are in warning state"
	Case nagios_CRITICAL
		prefix_nagios_message "One or more license key packs are in critical state"
	Case nagios_UNKNOWN
		prefix_nagios_message "Unable to obtain License Key Pack information." 
End Select
 
nagios_exit
