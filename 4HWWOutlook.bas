Dim objNameSpace As Outlook.NameSpace

'Set Offline Status on Outlook Startup
Private Sub Application_Startup()
    Set objNameSpace = Application.GetNamespace("MAPI")

          If objNameSpace.Offline = False Then
          'set it offline
          ActiveExplorer().CommandBars.ExecuteMso ("ToggleOnline")
          End If

End Sub

'Change Online/Offline Status at Specific Time
Private Sub Application_Reminder(ByVal Item As Object)
    Dim objOfflineTask As Outlook.TaskItem
    Dim objOnlineTask As Outlook.TaskItem

    Set objNameSpace = Application.GetNamespace("MAPI")

    If TypeOf Item Is TaskItem Then
       If Item.Subject = "Offline" Then
          Set objOfflineTask = Item

          'If Outlook is online when "Offline" task reminder alerts
          If objNameSpace.Offline = False Then
             'Set Outlook offline
             ActiveExplorer().CommandBars.ExecuteMso ("ToggleOnline")
          End If

          'Clear the reminder by marking task complete
          objOfflineTask.MarkComplete

       ElseIf Item.Subject = "Online" Then
          Set objOnlineTask = Item

          'If Outlook is offline when "Online" task reminder alerts
          If objNameSpace.Offline = True Then
             'Set Outlook online
             ActiveExplorer().CommandBars.ExecuteMso ("ToggleOnline")
          End If

          objOnlineTask.MarkComplete

       End If
    End If
End Sub
