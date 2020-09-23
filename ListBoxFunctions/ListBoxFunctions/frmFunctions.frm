VERSION 5.00
Begin VB.Form FrmFunctions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ListBox Functions - [http://www.8op.com/leaderx]"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5010
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmFunctions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3195
      Picture         =   "frmFunctions.frx":0E42
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Empty Items"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1010
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "String Doubles"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Doubles"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Text            =   "Number1"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E X I T"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton RemoveIt 
      Caption         =   "&Remove >>"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton AddItemz 
      Caption         =   "<< &Add Items"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "frmFunctions.frx":1EE4
      Left            =   120
      List            =   "frmFunctions.frx":1EE6
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Website 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddItemz_Click()
Dim x As Integer
If AddItemz.Tag = 0 Then AddItemz.Tag = 1: GoTo here
If AddItemz.Tag = 1 Then AddItemz.Tag = 2: GoTo here
If AddItemz.Tag = 2 Then AddItemz.Tag = 3: GoTo here
If AddItemz.Tag = 3 Then AddItemz.Tag = 4: GoTo here
If AddItemz.Tag = 4 Then AddItemz.Tag = 5: GoTo here
If AddItemz.Tag = 5 Then AddItemz.Tag = 6: GoTo here
If AddItemz.Tag = 6 Then Exit Sub
here:
If Option2.Value = True Then GoTo here2
List1.AddItem "<START>"
For x = 1 To 100
List1.AddItem "Number" & x
Next x
List1.AddItem "************   The Middle   ************"
z = 101
For x = 1 To 100
z = z - 1
List1.AddItem "Number" & z
Next x
List1.AddItem "<END>"
List1.ListIndex = 0
Exit Sub
here2:
List1.AddItem "100 Empty Items"
For x = 1 To 100
List1.AddItem " "
List1.AddItem "Number1" & x
Next x
List1.ListIndex = 0
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
AddItemz.Tag = 0
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
End Sub

Private Sub Option2_Click()
Text1.Enabled = False
End Sub

Private Sub Option3_Click()
Text1.Enabled = True
End Sub


Private Sub Picture1_Click()
Shell "explorer.exe http://www.8op.com/leaderx"
End Sub

Private Sub RemoveIt_Click()
If Option1.Value = True Then RemoveDoubles List1: List1.ListIndex = 0
If Option2.Value = True Then RemoveEmptyItems List1: List1.ListIndex = 0
If Option3.Value = True Then RemoveDoubleString List1, Text1: List1.ListIndex = 0
End Sub

Private Sub Website_Click()
Shell "explorer.exe http://www.8op.com/leaderx"
End Sub
