VERSION 5.00
Begin VB.Form test_frm 
   Caption         =   "save_data"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "create textbox fileds"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save data"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TEST SAVE DATA ON RECORD"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "test_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this utility help you to save data from any form in the project
' just call the function
' provide it with frm_name \ tbl_name .





Private Sub Command1_Click()
If Command2.Enabled = True Then Exit Sub

If empty_data(Me) Then
MsgBox "must fill all data "
Exit Sub
End If

If SAVE_DATA(Me, "students") = True Then
MsgBox vbTab & "data saved" & vbTab, vbInformation
Else
MsgBox error_var
End If
End Sub


Private Sub Command2_Click()
create_textbox Me, "students"

Me.Command2.Enabled = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

'width=97/233
' height=25
Private Sub Form_Load()
Me.KeyPreview = True
End Sub






