Attribute VB_Name = "MHapCod"
Option Explicit

'  http://visualbasic.happycodings.com/code-snippets/code5.html

Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long

Public Function IsAdmin() As Boolean
   
   IsAdmin = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))

End Function
