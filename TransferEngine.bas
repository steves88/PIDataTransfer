Attribute VB_Name = "TransferEngine"
Option Explicit

Public Sub TransferData()
   Dim rangeTagName As Range
   
   On Error GoTo errorHandler
   
   Application.Cursor = xlWait
   If ConnectServers() Then
      Set rangeTagName = ActiveSheet.Range("rangeTagsStart").Offset(1, 0)
      Do While Trim(rangeTagName.Value) <> ""
         ' If the target tag name is blank, then assume that the target
         ' name is the same as the source name.
         If Trim(rangeTagName.Offset(0, 1).Value) = "" Then
            ProcessTagValues rangeTagName.Value, rangeTagName.Value
         Else
            ProcessTagValues rangeTagName.Value, rangeTagName.Offset(0, 1).Value
         End If
         Set rangeTagName = rangeTagName.Offset(1, 0)
      Loop
   Else
      Application.Cursor = xlDefault
      MsgBox "Unable to connect to Source and Target PI servers.", vbExclamation, APP_NAME
   End If
   Application.Cursor = xlDefault
   Exit Sub
errorHandler:
   Application.Cursor = xlDefault
End Sub

Private Sub ProcessTagValues(ByVal SourceTagName As String, ByVal TargetTagName As String)
   Dim ReturnCode As Long
   Dim Times() As Long
   Dim Values() As Single
   Dim DigStates() As Long
   Dim SourcePointNum As Long
   Dim TargetPointNum As Long
   Dim Count As Long
   Dim Index As Long
   Dim ErrorCount As Integer
   Dim IsSourceTagDigital As Boolean
   Dim ItemCount As Long
   
   On Error GoTo errorHandler
   
   ItemCount = 0
   ErrorCount = 0
   SetActiveServer SOURCE_SERVER
   SourcePointNum = GetPointNumber(SourceTagName)
   If SourcePointNum > 0 Then
      IsSourceTagDigital = IsDigitalTag(SourceTagName)
      Count = ActiveSheet.Range("rangeMaxDataPoints").Value
      
      ReDim Values(Count)
      ReDim DigStates(Count)
      ReDim Times(Count)
      
      Times(0) = GetPITime(ActiveSheet.Range("rangeStartTime").Value)
      Times(Count - 1) = GetPITime(ActiveSheet.Range("rangeEndTime").Value)
      ReturnCode = piar_compvalues(SourcePointNum, Count, Times(0), Values(0), DigStates(0), 0)
      If ReturnCode = SUCCESS Then
         Application.StatusBar = Count & " archive values found for tag " & SourceTagName
         SetActiveServer TARGET_SERVER
         TargetPointNum = GetPointNumber(TargetTagName)
         If TargetPointNum > 0 Then
            If IsDigitalTag(TargetTagName) Then
               ' Make sure the source tag is also digital. The flag is set above when
               ' the source PI server is the active connection.
               If IsSourceTagDigital Then
                  ConvertDigStates SourceTagName, TargetTagName, DigStates
                  Count = Count - 1
                  For Index = 0 To Count
                     Application.StatusBar = "Writing value " & Index + 1 & " of " & Count & " to target PI server..."
                     ReturnCode = piar_putvalue(TargetPointNum, 0, DigStates(Index), Times(Index), 0)
                     If ReturnCode <> SUCCESS Then
                        ErrorCount = ErrorCount + 1
                        If ErrorCount > ActiveSheet.Range("rangeMaxErrors").Value Then
                           Exit Sub
                        End If
                     Else
                        ItemCount = ItemCount + 1
                        If ItemCount > ActiveSheet.Range("rangeMaxItems").Value Then
                           ItemCount = 0
                           Sleep
                        End If
                     End If
                  Next Index
               End If
            Else
               Count = Count - 1
               For Index = 0 To Count
                  Application.StatusBar = "Writing value " & Index + 1 & " of " & Count + 1 & " to target PI server..."
                  ReturnCode = piar_putvalue(TargetPointNum, Values(Index), 0, Times(Index), 0)
                  If ReturnCode <> SUCCESS Then
                     ErrorCount = ErrorCount + 1
                     If ErrorCount > ActiveSheet.Range("rangeMaxErrors").Value Then
                        Exit Sub
                     End If
                  Else
                     ItemCount = ItemCount + 1
                     If ItemCount > ActiveSheet.Range("rangeMaxItems").Value Then
                        ItemCount = 0
                        Sleep
                     End If
                  End If
               Next Index
            End If
         End If
      End If
   End If
   Exit Sub
errorHandler:
   MsgBox "Error occurred in the ProcessTagValues routine: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub Sleep()
   Const TIME_CONV_FACTOR As Double = 1 / 24 / 60 / 60
   Dim SleepStart As Date
   Dim SleepEnd As Date
   
   
   SleepStart = Now
   SleepEnd = SleepStart + (ActiveSheet.Range("rangeRestDuration").Value * TIME_CONV_FACTOR)
   Do While Now < SleepEnd
      Application.StatusBar = "Sleeping...  " & Format(Now, "hh:mm:ss")
      DoEvents
   Loop
End Sub

Private Sub ConvertDigStates(ByVal SourceTagName As String, _
                             ByVal TargetTagName As String, _
                             ByRef DigStates() As Long)
   Dim numStates As Long
   Dim Index As Long
   Dim StateString As String
   Dim StateStrings() As String
   
   numStates = UBound(DigStates)
   ReDim StateStrings(numStates)
   
   SetActiveServer SOURCE_SERVER
   For Index = 0 To numStates
      StateStrings(Index) = GetDigitalStateString(DigStates(Index))
   Next
   
   SetActiveServer TARGET_SERVER
   For Index = 0 To numStates
      DigStates(Index) = GetDigitalStateCodeForTag(TargetTagName, StateStrings(Index))
   Next Index
End Sub
