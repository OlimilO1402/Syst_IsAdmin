Attribute VB_Name = "MShell32"
Option Explicit

'  http://www.vbforums.com/showthread.php?794967-RESOLVED-Application-Run-With-Admin-Rights

Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer

Public Function IsAdmin() As Boolean
    
    IsAdmin = CBool(IsUserAnAdmin)
    
End Function
