VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My software 1.1"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":23D2
   ScaleHeight     =   7275
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CD6733&
      Caption         =   "Version Check/News"
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
      Height          =   2055
      Left            =   1320
      TabIndex        =   7
      Top             =   3600
      Width           =   6375
      Begin Secureactivation.XpBs command1 
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "Check for update"
         ButtonStyle     =   3
         Picture         =   "Form1.frx":8C4D
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   16
         OriginalPicSizeH=   16
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00CD6733&
         Caption         =   "Newer Version"
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
         Height          =   1095
         Left            =   2160
         TabIndex        =   11
         Top             =   600
         Width           =   1695
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "?.?.?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CD6733&
         Caption         =   "Current version"
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
         Height          =   1095
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1695
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "1.1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5640
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: You must be connected to internet to use this service."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
   End
   Begin Secureactivation.XpBs XpBs3 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "About this author"
      ButtonStyle     =   4
      Picture         =   "Form1.frx":B02F
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "Form1.frx":B1BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin Secureactivation.XpBs XpBs2 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      Caption         =   "I would like to activate MY Software through internet"
      ButtonStyle     =   4
      Picture         =   "Form1.frx":B45B
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "Form1.frx":B5E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin Secureactivation.XpBs XpBs1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Caption         =   "I would like to use Trial Version"
      ButtonStyle     =   4
      Picture         =   "Form1.frx":B887
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   14
      OriginalPicSizeH=   14
      PictureHover    =   "Form1.frx":BA12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPColor_Pressed =   13461299
      XPColor_Hover   =   13461299
      BackColor       =   13461299
      ForeColor       =   16777215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   700
      Left            =   0
      Picture         =   "Form1.frx":BCB3
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   6000
      Width           =   700
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":E754
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
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":E7F5
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   6480
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to MySoftware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Function HyperJump(ByVal URL As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub command1_Click()
'This section is used to check version

'This function assume files "application.ver", "news.txt" and "application.zip"
'on server http://server.com/user (change "server.com/user" by your server name and path)
'Inspect contain of files "news.txt" and "application.ver" at examples
Dim version As String, News As String
    'On Error GoTo ErrorMessage
    Me.MousePointer = 11
    'now assign content of file application.ver to variable Version
    version = Inet1.OpenURL("File://" + App.Path & "/" & "files" & "/" & "application.ver")
    Label8.Caption = version
    'You can try this function online, but You must change adresses:
    'for example: "http://server.com/yourname/application.ver"
    '===================================
   
    If version = "" Then GoTo Skip 'if file not found or file is empty then exit
    If version <= App.Major & "." & App.Minor Then
        Label6.Caption = "No newer version was released"
        Label8.Caption = version
        GoTo Skip
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    News = Inet1.OpenURL("file://" + App.Path & "/" & "files" & "/" & "news.txt")
    
    If MsgBox(News, vbYesNo, Me.Caption) = vbYes Then
        HyperJump "file://" + App.Path & "/" & "files" & "/" & "application.zip" 'this will run default download manager (probable also open default browser)
        
    End If
Skip:
    Me.MousePointer = 0
    Exit Sub
ErrorMessage:
    Me.MousePointer = 0
    MsgBox "An error has occured. Update failed." & Chr(10) & "You must download new version of this application manually at http://server.com.", vbCritical
End Sub


Private Sub Form_Load()
Unload Form2
'to verify the registration file if registered ,verifies _check.ini and compares Reg ID and Product ID
'_check.ini file will be creted when activated
'if _check.ini is not available then Trial is not diabled
Close #1
Dim regname
Dim productid
On Error GoTo errors
Open App.Path & "\" & "_check.ini" For Input As #1
Dim Code1 As Single
Dim i
Dim zip
Dim final
Line Input #1, regname
Line Input #1, productid
For i = 1 To Len(regname) - 1
    Code1 = Format(Asc(Right(regname, Len(regname) - i)) * 2 + (79 / i) + (i + 3 / 71), "#.#")
    zip = zip & Code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    Code1 = Format(Asc(Right(zip, Len(zip) - i)) * 2 + (1 / i) + (i + 1 / 9), "#00")
    final = final & Code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(regname)
If final = productid Then
Form2.Label1.Caption = 0
Form2.Label.Caption = "Registered"
Form2.Xp_ProgressBar1.Visible = False
Form2.Label4.Visible = True
Form2.Label.Visible = False
Form2.Label2.Visible = False
Form2.XpBs1.Visible = False
Form2.Label4.Caption = "Now you are in registered mode. Delete _check.ini in apps folder to setback trial"
XpBs1.Caption = "Enter Registered Software        "

Close #1
End If
errors: 'Form1.Show
End Sub

Private Sub XpBs1_Click()
Form2.Show
Unload Me

End Sub

Private Sub XpBs2_Click()
Form3.Show
Unload Me

End Sub

Private Sub XpBs3_Click()
Form4.Show
Unload Me
End Sub


