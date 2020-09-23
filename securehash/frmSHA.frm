VERSION 5.00
Begin VB.Form frmSHA 
   Caption         =   "Secure Hash Algorithm Test"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tboPass2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox tboPass1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label labStatus 
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label labHash2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label labHash1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tboPass1_GotFocus()
  tboPass1.SelStart = 0
  tboPass1.SelLength = Len(tboPass1.Text)
End Sub

Private Sub tboPass1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    tboPass2.SetFocus
  End If
End Sub

Private Sub tboPass1_LostFocus()
  HashPasswords
End Sub

Private Sub tboPass2_GotFocus()
  tboPass2.SelStart = 0
  tboPass2.SelLength = Len(tboPass2.Text)
End Sub

Private Sub tboPass2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    tboPass1.SetFocus
  End If
End Sub

Private Sub tboPass2_LostFocus()
  HashPasswords
End Sub

Private Sub HashPasswords()
  labHash1.Caption = ""
  labHash2.Caption = ""
  labStatus.Caption = ""
  
  If tboPass1.Text <> "" Then
    labHash1.Caption = SecureHash(tboPass1.Text)
  End If
  If tboPass2.Text <> "" Then
    labHash2.Caption = SecureHash(tboPass2.Text)
  End If
  
  If tboPass1.Text <> "" And tboPass2.Text <> "" Then
    If labHash1.Caption = labHash2.Caption Then
      labStatus.Caption = "Passwords match"
    Else
      labStatus.Caption = "Passwords do not match"
    End If
  End If
End Sub
