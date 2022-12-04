VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00CD6733&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mysoftware"
   ClientHeight    =   3060
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6525
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin Secureactivation.XpBs XpBs1 
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Set Trial to Zero"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Secureactivation.Xp_ProgressBar Xp_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      ProgressLook    =   6
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: You are in trial mode. If you want to use use this software permanently then you must activate this software."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "out of 100 executions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Menu file 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", ".": End
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: End
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function

Private Sub Form_Load()
TrialTime Form2, "The trial of " & Me.Caption & " has expired. Please activate this product to get the full version.", "Trial Expired", vbCritical, 110, True
'Activates the trial counter. True to count up and False to reset the Trial count
    Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
'Display times open
Xp_ProgressBar1.Value = Label1.Caption
End Sub


Private Sub XpBs1_Click()
  SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub
