VERSION 5.00
Begin VB.Form frmAnimation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's BitBlt Animation Example!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   240
      Top             =   600
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   120
      Picture         =   "frmAnimation.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   2
      Top             =   2160
      Width           =   7680
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3600
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   1
      Top             =   120
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Animation!"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Width: 64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Height: 64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "x: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-Mike Canejo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: Mike@dev-center.com      for your Questions or comments!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Frame:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *******************************************
'*   How to use BitBlt with SRCCOPY to get   *
'*   sections of a picture for animation.    *
'*
'*   By: Mike Canejo
'*   AIM: Mike3dd or TheLeadX
'*   Email: mike@dev-center.com
'*   Website: Http://www.8op.com/leaderx
' *******************************************

'* The Layout and Animation picture used in this example
'* Was created by a guy named Bradley. I made this alot
'* Easier and better to understand. Thanks Bradley....

'****************************
' set up bit block transfers
'****************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'BitBlt is how this works
'It works by find a location on a picture..cutting the
'section out and pasting it to a destination.
'BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'DestDC is where you want the cut-out picture to go. (picturebox)
'X is the horizontal coordinate of where to start the "cut" from the hSrcDC [xSrc]
'Y is the vertical coordinate of where to start the "cut" from the hSrcDC [ySrc]
'nWidth is the width of the section you wanna cut-out
'nHeight is the height of the section you wanan cut-out
'hSrcDC is where the cut-out is comming from [the source]...example PictureBox
'xSrc is the x-axis coordinate  of where to start cutting
'ySrc is the y-axis coordinate of where to start cutting
'dwRop is where to store the cut-out...SRCCOPY is what is usually used to store the cut-out

'Thats it..Diffacult? it looks like it at first but isn't
'Really..If u just dont get this then email me yur question
'at Mike@dev-center.com. Thanks for reading and i hope
'You have a better idea on how to use BitBlt for Animation

'--Mike Canejo

Private Const SRCCOPY = &HCC0020 ' Holds the section of the picture
Dim CountUP As Integer 'Counts up [useless]
Dim AnimationCount As Long ' dim an long value that holds the current X position of the animation sequence

'******************************************************
Private Sub Command1_Click()

If Not Timer1.Enabled Then  ' check to see whether or not the timer is on
    Timer1.Enabled = True   ' if the timer is not on, turn it on
Else
    Timer1.Enabled = False  ' if the timer is on, turn it off
End If

End Sub

Private Sub Form_Load()
AnimationCount = 0 'make sure the X count is at 0
CountUP = 0 ' make sure the Counup value is 0
End Sub

Private Sub Label4_Click()
Shell "explorer.exe http://www.8op.com/leaderx", vbNormalFocus 'Opens up my website :)
End Sub

Private Sub Label5_Click()
Shell "explorer.exe http://www.8op.com/leaderx", vbNormalFocus 'Opens up my website :)
End Sub

Private Sub Timer1_Timer()
Dim ReturnResult& ' dim a variable to hold the return value that the BitBlt function will send
ReturnResult = BitBlt(Picture1.hDC, 0, 0, 64, 64, Picture2.hDC, AnimationCount&, 0, SRCCOPY) 'BitBlt using SRCCOPY to copy information to the 0,0 X Y location of Picture1, and make the copy 64 x 64 pixels large, and tell BitBlt that the destination X location of Picture2 is held in the variable AnimationCount&, and the Y location is 0
Label3.Left = AnimationCount 'Shows current Frame being cut-out
Picture1.Refresh 'refresh Picture1 to reflect the SRCCOPY
AnimationCount& = AnimationCount& + 64 'increase the X position by 64 pixels, which will be SRCCOPY'ied in the next pass of this subroutine
If AnimationCount& = 512 Then AnimationCount& = 0: CountUP = 0 'if animationcount& is equal to the last frame of animation, reset it to 0
CountUP = Val(CountUP) + 1 'Counts frames (or sections) cut out of the picturebox
Label1.Caption = CountUP 'Displays it
Label6.Caption = "x: " & AnimationCount& 'Displays the current X coordinate of the cut-out
End Sub
