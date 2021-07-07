VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnIsAdminBagzz 
      Caption         =   "Bagzz.IsAdmin?"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton BtnIsAdminShell32 
      Caption         =   "Shell32.IsAdmin?"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton BtnIsAdminHapCod 
      Caption         =   "HappyC.IsAdmin?"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton BtnIsAdminDevX 
      Caption         =   "DevX.IsAdmin?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const csa As String = "App started as admin?: "
Const csu As String = "User is admin?: "

Private Sub BtnIsAdminHapCod_Click()
    
    MsgBox csa & MHapCod.IsAdmin
    
End Sub

Private Sub BtnIsAdminShell32_Click()
    
    MsgBox csa & MShell32.IsAdmin
    
End Sub

Private Sub BtnIsAdminBagzz_Click()
    
    MsgBox csa & MBagzz.IsAdmin
    
End Sub

Private Sub BtnIsAdminDevX_Click()
    
    MsgBox csu & MDevX.IsAdmin
    
End Sub

Private Sub Form_Load()
    Label1.Caption = csa & MHapCod.IsAdmin
    Label2.Caption = csa & MShell32.IsAdmin
    Label3.Caption = csa & MBagzz.IsAdmin
    Label4.Caption = csu & MDevX.IsAdmin
End Sub
