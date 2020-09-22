VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shimoon Ultra Spy v1.0"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0442
   ScaleHeight     =   5130
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Editing Info"
      Height          =   2805
      Left            =   90
      TabIndex        =   8
      Top             =   2250
      Width           =   3435
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Text            =   "<null>"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Text            =   "<null>"
         Top             =   1710
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Text            =   "<null>"
         Top             =   2340
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "NOTE: You cannot move ""+"" over this frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   2985
      End
      Begin VB.Label Label3 
         Caption         =   "Window text you wish to edit"
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   810
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Name of window to enable"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label Label7 
         Caption         =   "hWnd you wish to edit"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   2070
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "hWnd Information"
      Height          =   1545
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4155
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "hWnd:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Window Text:"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Parent:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3780
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   2340
      Width           =   555
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Drag the ""+"" over an object to see info. Then you can enter the required data to edit the strings."
      Height          =   465
      Left            =   90
      TabIndex        =   16
      Top             =   1710
      Width           =   4245
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3870
      Picture         =   "Form1.frx":0884
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "WINDOWS CRACKED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      TabIndex        =   15
      Top             =   4230
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Dim verylong As String * 100
Dim gParent As String * 100
Dim sndmsg As String * 100
Dim windowname As String * 100
Dim sztext As String * 100
Private Type POINTAPI
    X As Long
    Y As Long
    End Type

Dim mousemove As Boolean


Private Sub Form_Load()
Picture1.Picture = LoadResPicture(101, vbResIcon)

mousemove = False


End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Nothing
Form1.MousePointer = 99
mousemove = True

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim cursorpos1 As POINTAPI
   Dim wintext As String
   Dim garmon As String
   Dim gIcon As Image
   Dim OldX As Integer
   Dim OldY As Integer
   
   
 If mousemove = True Then
 
    r = GetCursorPos(cursorpos1)
    hwnd1 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    r = GetClassName(hwnd1, sztext, 100)
    hwnd2 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    p = GetWindowText(hwnd2, windowname, 100)
    hwnd3 = WindowFromPoint(cursorpos1.X, cursorpos1.Y)
    q = GetParent(hwnd3)
    
    Text4.Text = q
    Text2.Text = windowname
    Text1.Text = sztext


    If Text6.Text = Text2.Text Then
        
        r = EnableWindow(hwnd1, 1)
        Text6.Text = "<null>"
    End If
    
    If Text2.Text = Text3.Text Then
        garmon = InputBox("Choose new window title", "Title")
        q = SetWindowText(hwnd2, garmon)
        Picture1.Picture = LoadResPicture(101, vbResIcon)
        Form1.MousePointer = 0
        mousemove = False
        Text3.Text = "<null>"
    End If
    
    If Text1.Text = Text7.Text Then
        garmon = InputBox("Enter new string", "Title")
        z = SetWindowText(hwnd1, garmon)
        Picture1.Picture = LoadResPicture(101, vbResIcon)
        Form1.MousePointer = 0
        mousemove = False
        Text7.Text = "<null>"
    End If

If Text2.Text = "Editing Info" Then
   Form1.MousePointer = 0
        mousemove = False

End If
    
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = LoadResPicture(101, vbResIcon)
Form1.MousePointer = 0
mousemove = False
End Sub




