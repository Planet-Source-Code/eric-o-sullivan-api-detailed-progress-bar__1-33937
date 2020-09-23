VERSION 5.00
Object = "*\AProgress Bar.vbp"
Begin VB.Form frmTestProgressBar 
   Caption         =   "Testing Progress Bar"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   0
   End
   Begin CtlProgressBar.ctlProgBar ctlProgBar2 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1720
      Value           =   50
      Caption         =   "testing here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CtlProgressBar.ctlProgBar ctlProgBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      Value           =   50
      Caption         =   "hello"
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTestProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    Static sngCounter As Single
    
    If sngCounter > ctlProgBar2.Max Then
        sngCounter = ctlProgBar2.Min
    End If
    ctlProgBar1.Value = sngCounter
    ctlProgBar2.Value = sngCounter
    sngCounter = sngCounter + 0.1
End Sub
