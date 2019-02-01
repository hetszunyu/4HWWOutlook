# 4HWWOutlook
Implementation of the 4-Hour Workweek Email Rules in Outlook 2016
## Sources and Modifications
I took [this code and the instructions](https://www.datanumen.com/blogs/auto-set-outlook-online-offline-based-working-hours/) and updated the Work Offline toggle for Outlook 2016 compatibility using [this code](https://www.experts-exchange.com/questions/28945158/Excel-VBA-to-Toggle-Office-365-Outlook-Offline-Online.html)

To adapt it to 4HWW rules, Outlook automatically starts in Work Offline mode and only switches to online for specified short periods during the day to avoid interruptions.
## Instructions
1. Enable macros in Outlook (File-Options-Trust Center-Trust Center Settings-Macro Settings). I recommend Notifications for All Macros if you want to stay safe.
2. Follow the instructions in the first link in the Sources to set up tasks with reminders according to your taste. If you want to follow 4HWW rules, set up an "Online" task for 12:00 and 16:00 and an "Offline" task for 13:00 and 17:00. This way you can only check emails between noon-1pm and 4pm-5pm.
3. Open the VBA editor using Alt+F11. Find and open the ThisOutlookSession Project.
4. Copypaste the code into this project.
5. Restart Outlook and enjoy your distraction-free life.
