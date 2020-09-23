VERSION 5.00
Begin VB.Form frmAboutScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Information"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line lnSpacer 
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAboutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'This screen was first created on the 17/11/2001 and was intended
'for use in several future programs. The idea was that I should only
'have to create this screen once and be able to integrate it into any
'other project seemlessly. I wanted to do this instead of creating a
'new about screen for every project where I wanted one.
'
'A note on this About Screen :
'This screen requires the module APIGraphics (APIGraphics.bas) to
'operate the display.
'
'Eric O'Sullivan
'email DiskJunky@hotmail.com
'============================================================

Private mstrAllText As String
Private mblnStart As Boolean

Private Sub cmdOk_Click()
    'exit screen
    Unload Me
End Sub

Private Sub Form_Load()
    Call SetText
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub timText_Timer()
    'This timer will scroll the animated text
    
    Const Wait = 50 'wait 15 ticks before drawing the next frame
    
    Dim udtFont As FontStruc
    Dim udtBmp As BitmapStruc
    Dim udtMask As BitmapStruc
    Dim udtBmpSize As Rect
    Dim intResult As Integer
    Dim intTextHeight As Integer
    Dim lngStartingTick As Long
    
    Static udtSurphase As BitmapStruc
    Static intScroll As Integer
    
    'find out how much time it takes to draw a frame
    lngStartingTick = GetTickCount
    
    'set the bitmap dimensions and create them
    udtBmpSize.Right = picText.ScaleWidth
    udtBmpSize.Bottom = picText.ScaleHeight
    
    Call RectToPixels(udtBmpSize)
    
    udtMask.Area = udtBmpSize
    udtSurphase.Area = udtBmpSize
    udtBmp.Area = udtBmpSize
    
    'set font variables
    udtFont.Alignment = vbCentreAlign
    udtFont.Name = picText.FontName
    udtFont.Bold = picText.FontBold
    udtFont.Colour = vbWhite 'picText.ForeColor
    udtFont.Italic = picText.FontItalic
    udtFont.StrikeThru = picText.FontStrikethru
    udtFont.PointSize = picText.FontSize
    udtFont.Underline = picText.FontUnderline
    
    'test code - not currently used
    'Call MakeText(picText.hDc, "Hello World!", 0, 0, 40, 180, udtFont, InPixels)
    
    intTextHeight = GetTextHeight(picText.hDc) * LineCount(mstrAllText)
    
    intScroll = intScroll - Screen.TwipsPerPixelY
    If (intScroll < -(intTextHeight * Screen.TwipsPerPixelY)) _
       Or (Not mblnStart) Then
        intScroll = picText.ScaleHeight
        mblnStart = True
    End If
    
    'only create the surphase if necessary
    If udtSurphase.hDcMemory = 0 Then
        Call CreateNewBitmap(udtSurphase.hDcMemory, _
                             udtSurphase.hDcBitmap, _
                             udtSurphase.hDcPointer, _
                             udtSurphase.Area, _
                             frmAboutScreen.hDc, _
                             picText.ForeColor, _
                             InPixels)
        
        'create the surphase
        'text fade in
        Call Gradient(udtSurphase.hDcMemory, _
                      picText.ForeColor, _
                      picText.FillColor, _
                      0, _
                      (udtSurphase.Area.Bottom - ((intTextHeight / LineCount(mstrAllText)) * 2)), _
                      udtSurphase.Area.Right, _
                      (intTextHeight / LineCount(mstrAllText) * 2), _
                      GradHorizontal, InPixels)
        'text fade out
        Call Gradient(udtSurphase.hDcMemory, _
                      picText.FillColor, _
                      picText.ForeColor, _
                      0, _
                      0, _
                      udtSurphase.Area.Right, _
                      (intTextHeight / LineCount(mstrAllText)) * 2, _
                      GradHorizontal, _
                      InPixels)
    End If
    
    Call CreateNewBitmap(udtMask.hDcMemory, _
                         udtMask.hDcBitmap, _
                         udtMask.hDcPointer, _
                         udtMask.Area, _
                         frmAboutScreen.hDc, _
                         vbBlack, _
                         InPixels)
    Call CreateNewBitmap(udtBmp.hDcMemory, _
                         udtBmp.hDcBitmap, _
                         udtBmp.hDcPointer, _
                         udtBmp.Area, _
                         frmAboutScreen.hDc, _
                         vbWhite, _
                         InPixels)
    
    'draw the text onto the mask in black
    Call MakeText(udtMask.hDcMemory, _
                  mstrAllText, _
                  (intScroll / Screen.TwipsPerPixelY), _
                  0, _
                  intTextHeight, _
                  udtBmp.Area.Right, _
                  udtFont, _
                  InPixels)
    
    'copy the surphase onto the background
    intResult = BitBlt(udtBmp.hDcMemory, _
                       0, _
                       0, _
                       udtBmp.Area.Right, _
                       udtBmp.Area.Bottom, _
                       udtSurphase.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    'place the mask onto the background
    intResult = BitBlt(udtBmp.hDcMemory, _
                       0, _
                       0, _
                       udtBmp.Area.Right, _
                       udtBmp.Area.Bottom, _
                       udtMask.hDcMemory, _
                       0, _
                       0, _
                       SRCAND)
    
    'copy the result to the screen
    intResult = BitBlt(frmAboutScreen.hDc, _
                       0, _
                       0, _
                       udtBmp.Area.Right, _
                       udtBmp.Area.Bottom, _
                       udtBmp.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    'remove the bitmaps created
    Call DeleteBitmap(udtBmp.hDcMemory, _
                      udtBmp.hDcBitmap, _
                      udtBmp.hDcPointer)
    Call DeleteBitmap(udtMask.hDcMemory, _
                      udtMask.hDcBitmap, _
                      udtMask.hDcPointer)
    
    'wait X ticks minus the time it took to draw the frame
    Call Pause(Wait - (GetTickCount - lngStartingTick))
End Sub

Private Sub SetText()
    'This procedure is used to setting the text displayed in the picture box
    
    '" & vbCrLf & "
    
    'please note that ProductName can be set by going to
    'Project, Project Properties,Make tab. You should see a list box about
    'half way down on the left side. Scroll down until you come to
    'Product Name and enter some text into the text box on the right
    'side of the list box.
    mstrAllText = App.ProductName & vbCrLf & _
                  "Version " & App.Major & "." & _
                               App.Minor & "." & _
                               App.Revision & vbCrLf & _
                  "" & vbCrLf & _
                  "This program was made by" & vbCrLf & _
                  "Eric O'Sullivan." & vbCrLf & _
                  "" & vbCrLf & _
                  "Copyright 2002" & vbCrLf & _
                  "All rights reserved" & vbCrLf & _
                  "" & vbCrLf & _
                  "For more information, email" & vbCrLf & _
                  "DiskJunky@hotmail.com"
End Sub

Private Function LineCount(ByVal strText As String) _
                           As Integer
    'This function will return the number of lines
    'in the strText
    
    Dim intTemp As Integer
    Dim intCounter As Integer
    Dim intLastPos As Integer
    
    intLastPos = 1
    
    Do
        intTemp = intLastPos
        intLastPos = InStr(intLastPos + Len(vbCrLf), strText, vbCrLf)
        
        If intTemp <> intLastPos Then
            'a line was found
            intCounter = intCounter + 1
        End If
    Loop Until intLastPos = 0 'intLastPos will =0 when InStr cannot find any more occurances of vbCrlf
    
    LineCount = intCounter
End Function
