VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lift Napster Ban coded by Navarchy"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Navarchy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3440
      MouseIcon       =   "Form1.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   510
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   $"Form1.frx":0FD4
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1620
      MouseIcon       =   "Form1.frx":1100
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1395
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":140A
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Napster.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      MouseIcon       =   "Form1.frx":14AF
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2715
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Now all you have to do is reinstall Napster from Napster.com and enjoy the free music once again!"
      Height          =   855
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Step 4:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "LIFT THE BAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4920
      MouseIcon       =   "Form1.frx":17B9
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "This is the easy part. Click ""LIFT THE BAN"""
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Step 3:"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":1AC3
      Height          =   735
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Step 2:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vote As Boolean
Private Sub Form_Load()
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vote = GetSetting("napster", "vote", "vote", "false")
If vote = "false" Then Call Label9_Click
End Sub

Private Sub Label11_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:Navarchy@AOL.com", vbHide
End Sub

Private Sub Label5_Click()
bSetRegValue HKEY_LOCAL_MACHINE, "Software\Napster", "CurrentUser", ""
bSetRegValue HKEY_CLASSES_ROOT, "CLSID\{CAD8C813-1F34-1B3E-00CEAE43FF0AAD}", "id", ""
bSetRegValue HKEY_CURRENT_USER, "Software\Classes\CLSID\{CAD8C813-1F34-1B3E-00CEAE43FF0AAD}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{CAD8C813-1F34-1B3E-00CEAE43FF0AAD}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_USERS, ".DEFAULT\Software\Classes\CLSID\{CAD8C813-1F34-1B3E-00CEAE43FF0AAD}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_CLASSES_ROOT, "CLSID\{35D38C13-1434-AB7E-003483943341AA}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_USERS, ".DEFAULT\Software\Classes\CLSID\{35D38C13-1434-AB7E-003483943341AA}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_CURRENT_USER, "Software\Classes\CLSID\{35D38C13-1434-AB7E-003483943341AA}", "id", "zzzzzzzzzzz74324"
bSetRegValue HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{35D38C13-1434-AB7E-003483943341AA}", "id", "zzzzzzzzzzz74324"
End Sub

Private Sub Label8_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.napster.com/download/", vbMaximizedFocus
End Sub

Private Sub Label9_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE https://planet-source-code.com/ads/authentication/login.asp?lngWId=1&blnOutsideOfVBSubWeb=False&txtReturnURL=http%3A%2F%2Fplanet%2Dsource%2Dcode%2Ecom%2Fvb%2Fscripts%2Fvoting%2FVoteOnCodeRating%2Easp%3FlngWId%3D1%26optCodeRatingValue=5%26cmdRateIt=Rate%A0It%21%26txtCodeId=22101%26txtCodeName=Remove+Napster+Ban+From+Your+Computer%21%26intUserRatingTotal=0%26intNumOfUserRatings=0&txtCancelURL=%2Fvb%2Fdefault%2Easp%3FlngWId%3D1", vbMaximizedFocus
SaveSetting "napster", "vote", "vote", "true"
End Sub
