Attribute VB_Name = "ModifyRegistry"
Option Compare Database
Public Const REG_PATH_NAME = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Article Tracker\Path"
Public Const REG_NOTIFICATIONS = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Article Tracker\Notifications"

'Add a trusted location to the registry so dynamic content is enabled
Public Function AddTrustedLocation()
If RegKeyExists(REG_PATH_NAME) = False Then
Dim myWS As Object
Set myWS = CreateObject("WScript.Shell")
Dim Path As String
Dim subfolders As String
Dim datemod As String
Dim description As String

Path = REG_PATH_NAME
subfolders = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Article Tracker\AllowSubfolders"
datemod = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Article Tracker\Date"
description = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Access\Security\Trusted Locations\Article Tracker\Description"

myWS.regWrite Path, "C:\", "REG_SZ"
myWS.regWrite subfolders, &H1, "REG_DWORD"
myWS.regWrite datemod, date, "REG_SZ"
myWS.regWrite description, "C:\ drive trusted location to facilitate Article Tracking system", "REG_SZ"

Else
'do nothing
End If
End Function

'Check if notifications are muted by user
Public Function IsMuted()
Dim myWS As Object
Set myWS = CreateObject("WScript.Shell")

If RegKeyExists(REG_NOTIFICATIONS) = True Then
    IsMuted = myWS.regread(REG_NOTIFICATIONS)
Else
    IsMuted = 0
End If
End Function

'mutes notifications for greater usability
Public Function MuteNotifications()
Dim myWS As Object
Set myWS = CreateObject("WScript.Shell")

'write with value 1
myWS.regWrite REG_NOTIFICATIONS, &H1, "REG_DWORD"

End Function

'unmutes notifications on the 25th day or greater of every month
'so that the next month's notifications can be seen
Public Function UnmuteNotifications()
Dim myWS As Object
Set myWS = CreateObject("WScript.Shell")

'if the key does not exist, write to it with val 0
'so that notifications are unmuted by default
If RegKeyExists(REG_NOTIFICATIONS) = False Then
myWS.regWrite REG_NOTIFICATIONS, &H0, "REG_DWORD"
End If
Debug.Print "is muted = " & IsMuted & vbCrLf
'if it does exist, and it is end of month, set notifications to be unmuted
If DatePart("d", date) > 15 Then
If IsMuted > 0 Then
myWS.regWrite REG_NOTIFICATIONS, &H0, "REG_DWORD"
Debug.Print "after unmuting: " & IsMuted & vbCrLf
End If
End If

End Function

'returns True if the registry key i_RegKey was found
'and False if not
Function RegKeyExists(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  
  'try to read the registry key
  myWS.regread i_RegKey
  
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  
  'key was not found
  RegKeyExists = False
End Function


