VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Progressbar 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   4200
   End
   Begin VB.Label Loading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4560
      Width           =   5295
   End
End
Attribute VB_Name = "Progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Loading.FontName = "Tahoma"
    Loading.FontSize = 10
    
End Sub

Private Sub Timer1_Timer()
    Loading.Caption = "Loading Please Wait...." & " " & ProgressBar1.Value & "%"
    ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
            Timer1.Enabled = False
            Unload Me
        End If
End Sub
