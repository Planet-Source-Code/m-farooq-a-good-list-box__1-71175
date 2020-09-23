VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1950
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   -30
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   2835
      TabIndex        =   11
      Top             =   -10
      Width           =   2865
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2CEB0&
      ForeColor       =   &H00000000&
      Height          =   1395
      ItemData        =   "Form1.frx":2330
      Left            =   3030
      List            =   "Form1.frx":2332
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E4DC&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2340
      Picture         =   "Form1.frx":2334
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   1515
      Width           =   345
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   2
         Left            =   45
         TabIndex        =   9
         Top             =   -60
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   90
      Picture         =   "Form1.frx":4664
      ScaleHeight     =   225
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   1515
      Width           =   2265
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   90
      Picture         =   "Form1.frx":6994
      ScaleHeight     =   225
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   975
      Width           =   2265
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E4DC&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2340
      Picture         =   "Form1.frx":8CC4
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   975
      Width           =   345
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   1
         Left            =   45
         TabIndex        =   5
         Top             =   -60
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9E4DC&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2340
      Picture         =   "Form1.frx":AFF4
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   435
      Width           =   345
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   -60
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   90
      Picture         =   "Form1.frx":D324
      ScaleHeight     =   225
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   435
      Width           =   2265
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You can also open the listbox by pressing CTRL + 0 for 1st , CTRL + 1 for 2nd and CTRL + 2 for 3rd list box"
      Height          =   675
      Left            =   60
      TabIndex        =   12
      Top             =   3240
      Width           =   2685
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F2CEB0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press {ESC} to Exit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   -30
      TabIndex        =   10
      Top             =   3930
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mIndex As Integer

Const FLAGS As Long = SND_ASYNC Or SND_FILENAME Or SND_NOSTOP

Private Sub Form_Click()
    List1.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If List1.Visible = True Then
            List1.Visible = False
        Else
            Unload Me
        End If
    End If
    
    If KeyCode = vbKey0 And Shift = 2 Then
            Label1_Click (0)
        ElseIf KeyCode = vbKey1 And Shift = 2 Then
            Label1_Click (1)
        ElseIf KeyCode = vbKey2 And Shift = 2 Then
            Label1_Click (2)
    End If

End Sub
Private Sub Form_Load()
    CForm Me
    WriteOnPictureBox.PrintToCenter "List Box Example", Picture1(3)
 
End Sub

Private Sub Form_Resize()
    Me.Height = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Respected sir !" & vbCr + vbCr & "If you can make bug free .CTL for this list box" & vbCr & "I shall be grateful to you", vbInformation, "HELP"
End Sub

Private Sub Label1_Click(Index As Integer)
    
    PlaySound App.Path & "\click.wav", 0, FLAGS
    
    Select Case Index
        Case 0
            List1.Left = Picture1(Index).Left
            List1.Top = Picture1(Index).Top + 240
            List1.Visible = True
            List1.SetFocus
            List1.Clear
            mIndex = Index
            For a = 65 To 70
                List1.AddItem Chr(a)
            Next
            
        Case 1
            List1.Left = Picture1(Index).Left
            List1.Top = Picture1(Index).Top + 240
            List1.Visible = True
            List1.SetFocus
            List1.Clear
            mIndex = Index
            
            For a = 101 To 105
                List1.AddItem Chr(a)
            Next
    
        Case 2
            List1.Left = Picture1(Index).Left
            List1.Top = Picture1(Index).Top + 240
            List1.Visible = True
            List1.SetFocus
            List1.Clear
            mIndex = Index
            
            For a = 1 To 10
                List1.AddItem a
            Next
    
    End Select
End Sub



Private Sub List1_DblClick()
    Picture1(mIndex).Cls
    WriteOnPictureBox.PrintToCenter List1.Text, Picture1(mIndex)
    List1.Visible = False

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        List1_DblClick
    End If
End Sub

Private Sub Picture1_Click(Index As Integer)
    Select Case Index
        Case 0
            Label1_Click (Index)
        Case 1
            Label1_Click (Index)
        Case 2
            Label1_Click (Index)
    
    End Select
    
        
End Sub
