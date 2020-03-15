Attribute VB_Name = "Module1"
Option Explicit

Private Sub test()
    Dim cron1 As CronExpression
    
    Set cron1 = New CronExpression
    cron1.Parse "0 0 12 ? * WED"
    
    ' 日と曜日のチェック --------------------
    Debug.Print "日と曜日のチェック"
    
    cron1.Parse "0 0 12 1 * WED"
    
    If cron1.IsError And cron1.ErrorMessage = "日と曜日のいずれかにワイルドカードを指定してください" Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    
End Sub

