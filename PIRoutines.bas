Attribute VB_Name = "PIRoutines"
Option Explicit

Public Function ConnectServers() As Boolean
   Dim AccessLevel As Long
   Dim rc1 As Long
   Dim rc2 As Long
   Dim rc3 As Long
   Dim rc4 As Long
   Dim Password As String
   
   On Error GoTo errorHandler
   
   ' Password can be blank, so make sure it is a string before passing it to
   ' the login function. If it is blank, and you pass in the cell's value
   ' directly, you get an "empty" instead of a null string ("")
   Password = CStr(ActiveSheet.Range("rangeSourcePassword").Value)
   
   ' Connect to the Source PI Server
   rc1 = piut_setservernode(ActiveSheet.Range("rangeSourceServer").Value)
   rc2 = piut_login(ActiveSheet.Range("rangeSourceUser").Value, Password, AccessLevel)
   
   ' Password can be blank, so make sure it is a string before passing it to
   ' the login function. If it is blank, and you pass in the cell's value
   ' directly, you get an "empty" instead of a null string ("")
   Password = CStr(ActiveSheet.Range("rangeTargetPassword").Value)
   
   ' Connect to the Target PI Server
   rc3 = piut_setservernode(ActiveSheet.Range("rangeTargetServer").Value)
   rc4 = piut_login(ActiveSheet.Range("rangeTargetUser").Value, Password, AccessLevel)
   
   If rc1 = SUCCESS And rc2 = SUCCESS And rc3 = SUCCESS And rc4 = SUCCESS Then
      ConnectServers = True
   Else
      ConnectServers = False
   End If
   Exit Function
errorHandler:
   ConnectServers = False
End Function

Public Sub SetActiveServer(ByVal ServerID As Integer)
   Dim ReturnCode As Long
   On Error Resume Next
   If ServerID = SOURCE_SERVER Then
      ReturnCode = piut_setservernode(ActiveSheet.Range("rangeSourceServer").Value)
   ElseIf ServerID = TARGET_SERVER Then
      ReturnCode = piut_setservernode(ActiveSheet.Range("rangeTargetServer").Value)
   End If

End Sub


Public Function GetSystemDigitalStateCode(ByVal StateString As String) As Long
   Dim ReturnCode As Long
   Dim DigCode As Long
   Dim PointNum As Long
   
   On Error GoTo errorHandler
   
   ReturnCode = pipt_digcode(DigCode, StateString)
   If ReturnCode = SUCCESS Then
      If DigCode <> 0 Then
         GetSystemDigitalStateCode = -DigCode
      Else
         GetSystemDigitalStateCode = 0
      End If
   Else
      GetSystemDigitalStateCode = 0
   End If
   Exit Function
errorHandler:
      GetSystemDigitalStateCode = 0
End Function

Public Function GetDigitalStateCodeForTag(ByVal TagName As String, ByVal StateString As String) As Long
   Dim ReturnCode As Long
   Dim DigCode As Long
   Dim PointNum As Long
   
   On Error GoTo errorHandler
   
   PointNum = GetPointNumber(TagName)
   ReturnCode = pipt_digcodefortag(PointNum, DigCode, StateString)
   If ReturnCode = SUCCESS Then
      If DigCode <> 0 Then
         GetDigitalStateCodeForTag = -DigCode
      Else
         GetDigitalStateCodeForTag = 0
      End If
   Else
      GetDigitalStateCodeForTag = 0
   End If
   Exit Function
errorHandler:
      GetDigitalStateCodeForTag = 0
End Function

Public Function GetDigitalStateString(ByVal StateCode As Long) As String
   Dim ReturnCode As Long
   Dim StateStr As String * 80
   
   On Error GoTo errorHandler
   
   ReturnCode = pipt_digstate(StateCode, StateStr, 80)
   If ReturnCode = SUCCESS Then
      GetDigitalStateString = TrimNulls(StateStr)
   Else
      GetDigitalStateString = ""
   End If
   Exit Function
errorHandler:
   GetDigitalStateString = ""
End Function

Public Function GetPointNumber(TagName As String) As Long
   Dim ReturnCode As Long
   Dim PointNum As Long
   Dim Tag As String * 80
   
   On Error GoTo errorHandler
   
   Tag = Trim(TagName)
   ReturnCode = pipt_findpoint(Tag, PointNum)
   If ReturnCode = SUCCESS Then
      GetPointNumber = PointNum
   Else
      ' If the point was not found, return a bogus point number
      ' Yes, the returned value is just like the famous zip code.
      GetPointNumber = -90210
   End If
   Exit Function
errorHandler:
   GetPointNumber = -90210
End Function

Public Function GetPIServerTime() As Long
   Dim ReturnCode As Long
   Dim Stime As Long
   
   On Error GoTo errorHandler
   ReturnCode = pitm_servertime(Stime)
   If ReturnCode = 1 Then
      GetPIServerTime = Stime
   Else
      GetPIServerTime = GetPITime(Now)
   End If
   Exit Function
errorHandler:
   GetPIServerTime = GetPITime(Now)
End Function

Public Function GetPITime(Timestamp As Date) As Long
   Dim ReturnCode As Long
   Dim TimeString As String * 19
   Dim RelativeTime As Long
   Dim TimeVal As Long
   
   On Error GoTo errorHandler
   
   ' Format the date in the way PI expects to see it.
   TimeString = Format(Timestamp, "dd-mmm-yy hh:mm:ss")
   
   ReturnCode = pitm_parsetime(TimeString, RelativeTime, TimeVal)
   If ReturnCode = SUCCESS Then
      GetPITime = TimeVal
   Else
      GetPITime = -9999
   End If
   Exit Function
errorHandler:
   GetPITime = -9999
End Function

Public Function GetTimeString(TimeVal As Long) As String
   Dim ReturnCode As Long
   Dim TimeString As String * 19
   
   On Error GoTo errorHandler
      
   pitm_formtime TimeVal, TimeString, 19
   GetTimeString = TrimNulls(TimeString)
   Exit Function
errorHandler:
   GetTimeString = "1-Jan-70 00:00:00"
End Function

Public Function TrimNulls(ByVal str As String) As String
   Dim NullPosition As Integer
   
   On Error GoTo errorHandler
   
   NullPosition = InStr(1, str, vbNullChar)
   If NullPosition > 0 Then
      TrimNulls = Trim(Left(str, NullPosition - 1))
   Else
      TrimNulls = Trim(str)
   End If
   Exit Function
errorHandler:
   TrimNulls = str
End Function

Public Function IsDigitalTag(ByVal TagName As String) As Boolean
   Dim PointNum As Long
   Dim ReturnCode As Long
   Dim PointType As String * 1
   
   On Error GoTo errorHandler
   
   PointNum = GetPointNumber(TagName)
   If PointNum > 0 Then
      ReturnCode = pipt_pointtype(PointNum, PointType)
      If ReturnCode = SUCCESS Then
         If PointType = "D" Then
            IsDigitalTag = True
         Else
            IsDigitalTag = False
         End If
      Else
         IsDigitalTag = False
      End If
   Else
      IsDigitalTag = False
   End If
   Exit Function
errorHandler:
   IsDigitalTag = False
End Function

Public Function TagExists(ByVal TagName As String) As Boolean
   Dim ReturnCode As Long
   Dim PointNum As Long
   Dim Tag As String * 80
   
   On Error GoTo errorHandler
   
   Tag = Trim(TagName)
   ReturnCode = pipt_findpoint(Tag, PointNum)
   If ReturnCode = SUCCESS Then
      TagExists = True
   Else
      TagExists = False
   End If
   Exit Function
errorHandler:
      TagExists = False
End Function


