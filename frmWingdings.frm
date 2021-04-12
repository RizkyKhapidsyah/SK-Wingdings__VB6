VERSION 5.00
Begin VB.Form frmWingdings 
   BackColor       =   &H00000000&
   Caption         =   "Wingdings"
   ClientHeight    =   3495
   ClientLeft      =   1515
   ClientTop       =   1935
   ClientWidth     =   9030
   Icon            =   "frmWingdings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9030
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Columns         =   5
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click the Wingding You Want to Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
   Begin VB.Image CmdDone 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "frmWingdings.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmWingdings.frx":0884
      ToolTipText     =   "Apply To All Lines"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image CmdApplySingle 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "frmWingdings.frx":0B8E
      MousePointer    =   99  'Custom
      Picture         =   "frmWingdings.frx":0FD0
      ToolTipText     =   "Apply at Insertion"
      Top             =   1560
      Width           =   480
   End
End
Attribute VB_Name = "frmWingdings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdApplySingle_Click()
If Text1.Text <> "" Then
Temp = "<font face=" & Q & "wingdings" & Q & ">" & Text1.Text & "</font>"
Clipboard.Clear
Clipboard.SetText Temp
Title = "Done!"
Style = 64
' vbOKOnly + vbInformation + vbDefaultButton1 + None
Message = "The HTML code for the wingding you selected has been copied to the Windows Clipboard."
Response = MsgBox(Message, Style, Title)
End If
End Sub

Private Sub CmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
List1.Clear
For X = 33 To 255
Temp = Chr(X)
List1.AddItem Temp
Next
End Sub

Private Sub List1_DblClick()
For X = 0 To List1.ListCount - 1
If List1.Selected(X) = True Then
Counter = X + 33
GoTo Out
End If
Next
Out:
Text1.Text = "&#" & Counter & ";"


End Sub
