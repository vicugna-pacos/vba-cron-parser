Attribute VB_Name = "Module1"
Option Explicit

Private Sub test()
    Dim cron1 As CronExpression
    
    Set cron1 = New CronExpression
    cron1.Parse "0 0 12 ? * WED"
    
    ' ���Ɨj���̃`�F�b�N --------------------
    Debug.Print "���Ɨj���̃`�F�b�N"
    
    cron1.Parse "0 0 12 1 * WED"
    
    If cron1.IsError And cron1.ErrorMessage = "���Ɨj���̂����ꂩ�Ƀ��C���h�J�[�h���w�肵�Ă�������" Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    
End Sub

