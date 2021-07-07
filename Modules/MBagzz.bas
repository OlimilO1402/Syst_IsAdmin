Attribute VB_Name = "MBagzz"
Option Explicit
'https://foren.activevb.de/archiv/vb-classic/thread-407651/beitrag-407701/Re-Pruefen-ob-Programm-als-Admi/
Private Const SECURITY_BUILTIN_DOMAIN_RID       As Long = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS           As Long = &H220

Private Declare Function AllocateAndInitializeSid Lib "advapi32" (pIdentifierAuthority As Any, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Sub FreeSid Lib "advapi32" (ByVal pSid As Long)
Private Declare Function CheckTokenMembership Lib "advapi32" (ByVal hToken As Long, ByVal pSidToCheck As Long, pbIsMember As Long) As Long

Private Type SID_IDENT_AUTH
    Value(0 To 5)            As Byte
End Type

Public Function IsAdmin() As Boolean
    Dim pSidAdmins As Long
    Dim sia As SID_IDENT_AUTH: sia.Value(5) = 5
    If AllocateAndInitializeSid(sia, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, pSidAdmins) <> 0 Then
        Dim rv As Long
        If CheckTokenMembership(0, pSidAdmins, rv) <> 0 Then
            IsAdmin = (rv <> 0)
        End If
        Call FreeSid(pSidAdmins)
    End If
End Function
