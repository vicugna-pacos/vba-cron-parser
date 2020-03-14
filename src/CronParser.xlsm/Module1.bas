Attribute VB_Name = "Module1"
Option Explicit

Private Sub test()
    Dim cron1 As CronExpression
    
    Set cron1 = New CronExpression
    cron1.Parse "0 0 12 ? * WED"
End Sub

