VERSION 5.00
Begin VB.UserControl ctlProgBar 
   Appearance      =   0  'Flat
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   HitBehavior     =   0  'None
   PropertyPages   =   "ctlProgBar.ctx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ToolboxBitmap   =   "ctlProgBar.ctx":002F
   Begin VB.PictureBox picSurphase 
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "ctlProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Option Explicit

'the border style of the control
Public Enum BorderEnum
    pgrNone = 0
    pgrFixed_Single = 1
End Enum

'default colours
Private Const vbDarkBlue = &H800000

'property variables
Private msngMax As Single           'the Max value of the progress bar
Private msngMin As Single           'the Min value of the progress bar
Private msngValue As Single         'the position of the progress bar
Private mstrCaption As String       'the user defined caption for the progress bar
Private mblnDefCapt As Boolean      'the defualt caption is the percentage of the progress bar taken up by Value
Private mlngBackColour As Long      'the initial background of the progress bar when Value = Min
Private mlngFillColour As Long      'the colour of the progress bar as it takes up space on the screen
Private mlngTextColour As Long      'the initial caption colour for the text
Private mlngOverColour As Long      'the colour when the progress bar moves over the caption
Private mfntCaption As FontStruc    'the font of the caption
Private menmBorder As BorderEnum    'the border style of the user control

'general variables
Private mudtBackBmp As BitmapStruc  'the background of the user control

'Event Declarations:
Public Event Click()                       'MappingInfo=UserControl,UserControl,-1,Click
Public Event DblClick()                    'MappingInfo=UserControl,UserControl,-1,DblClick
Public Event MouseDown(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)        'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)        'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, _
                     Shift As Integer, _
                     X As Single, _
                     Y As Single)          'MappingInfo=UserControl,UserControl,-1,MouseUp

'properties

Public Property Get PercentCaption() As Boolean
Attribute PercentCaption.VB_Description = "If set to True, the progress bar text will display the percentage of the control covered by the progress bar"
Attribute PercentCaption.VB_ProcData.VB_Invoke_Property = ";Text"
    'return whether or not to display the
    'default caption
    PercentCaption = mblnDefCapt
End Property

Public Property Let PercentCaption(ByVal blnNewValue As Boolean)
    'set the default caption
    mblnDefCapt = blnNewValue
    Call Refresh
    PropertyChanged "PercentCaption"
End Property

Public Property Get Max() As Single
Attribute Max.VB_Description = "The upper range of the progress bar"
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Scale"
    'return the Max value
    Max = msngMax
End Property

Public Property Let Max(ByVal sngNewValue As Single)
    'set the Max value
    
    If sngNewValue < msngMin Then
        Exit Property
    End If
    
    If sngNewValue < msngValue Then
        msnvalue = sngNewValue
    End If
    
    msngMax = sngNewValue
    Call Refresh
    PropertyChanged "Max"
End Property

Public Property Get Value() As Single
Attribute Value.VB_Description = "This sets the position of the progress bar. It must be between the Max and Min values"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Scale"
Attribute Value.VB_MemberFlags = "200"
    'return the Value
    Value = msngValue
End Property

Public Property Let Value(ByVal sngNewValue As Single)
    'set the Value
    
    If sngNewValue > msngMax Then
        sngNewValue = msngMax
    End If
    If sngNewValue < msngMin Then
        sngNewValue = msngMin
    End If
    
    msngValue = sngNewValue
    Call Refresh
    PropertyChanged "Value"
End Property

Public Property Get Min() As Single
Attribute Min.VB_Description = "The lower range of the progress bar"
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Scale"
    'return the Min value
    Min = msngMin
End Property

Public Property Let Min(ByVal sngNewValue As Single)
    'set the Min value
    
    If sngNewValue > msngMax Then
        Exit Property
    End If
    
    If sngNewValue > msngValue Then
        msngValue = sngNewValue
    End If
    
    msngMin = sngNewValue
    Call Refresh
    PropertyChanged "Min"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets the text displayed in the progress bar. This is ignored if PercentCaption is set to True"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    'return the caption
    Caption = mstrCaption
End Property

Public Property Let Caption(ByVal strNewValue As String)
    'set the caption
    mstrCaption = strNewValue
    Call Refresh
    PropertyChanged "Caption"
End Property

Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "Sets the background colour of the area not covered by the progress bar"
Attribute BackColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColour.VB_UserMemId = -501
    'get the background colour
    BackColour = mlngBackColour
End Property

Public Property Let BackColour(ByVal lngNewValue As OLE_COLOR)
    'set the background colour
    mlngBackColour = lngNewValue
    Call Refresh
    PropertyChanged "BackColour"
End Property

Public Property Get FillColour() As OLE_COLOR
Attribute FillColour.VB_Description = "This sets the colour of the progress bar"
Attribute FillColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FillColour.VB_UserMemId = -510
    'get the fill colour
    FillColour = mlngFillColour
End Property

Public Property Let FillColour(ByVal lngNewValue As OLE_COLOR)
    'set the fill colour
    mlngFillColour = lngNewValue
    Call Refresh
    PropertyChanged "FillColour"
End Property

Public Property Get TextColour() As OLE_COLOR
Attribute TextColour.VB_Description = "This is the colour of the Caption when the progress bar is not covering it"
Attribute TextColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute TextColour.VB_UserMemId = -513
    'get the text colour
    TextColour = mlngTextColour
End Property

Public Property Let TextColour(ByVal lngNewValue As OLE_COLOR)
    'set the text colour
    mlngTextColour = lngNewValue
    Call Refresh
    PropertyChanged "TextColour"
End Property

Public Property Get OverColour() As OLE_COLOR
Attribute OverColour.VB_Description = "This set the colour of the text when the progress bar moves over it"
Attribute OverColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'get the text over colour
    OverColour = mlngOverColour
End Property

Public Property Let OverColour(ByVal lngNewValue As OLE_COLOR)
    'set the text over colour
    mlngOverColour = lngNewValue
    picSurphase.Appearance = menmBorder
    Call Refresh
    PropertyChanged "OverColour"
End Property

Public Property Get BorderStyle() As BorderEnum
Attribute BorderStyle.VB_Description = "Sets the border style of the progress bar"
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -520
    'get the border style
    BorderStyle = menmBorder
End Property

Public Property Let BorderStyle(ByVal enmNewValue As BorderEnum)
    'set the border style
    menmBorder = enmNewValue
    Call Refresh
    PropertyChanged "BorderStyle"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "This sets the font attributes of the text displayed in the progress bar"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    'get a new font
    Call GetFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal fntNewValue As Font)
    'set the new font
    
    Set UserControl.Font = fntNewValue
    Call SetFont
    
    Call Refresh
    PropertyChanged "Font"
End Property

'methods/procedures

Private Sub GetFont()
    'This will retrieve the font settings from
    'the user control so that they can be used
    'to display the text properly
    
    If mfntCaption.Name = "" Then
        Exit Sub
    End If
    
    With mfntCaption
        UserControl.Font.Bold = .Bold
        UserControl.Font.Italic = .Italic
        UserControl.Font.Name = .Name
        UserControl.Font.Size = .PointSize
        UserControl.Font.Strikethrough = .StrikeThru
        UserControl.Font.Underline = .Underline
    End With
End Sub

Private Sub SetFont()
    'this will enter the settings of the font
    'structure into the user control so that
    'they can be saved in the property bag
    
    With mfntCaption
        .Bold = UserControl.Font.Bold
        .Italic = UserControl.Font.Italic
        .Name = UserControl.Font.Name
        .PointSize = UserControl.Font.Size
        .StrikeThru = UserControl.Font.Strikethrough
        .Underline = UserControl.Font.Underline
    End With
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Displays the about box for this form"
Attribute ShowAbout.VB_UserMemId = -552
    frmAboutScreen.Show vbModal
End Sub

Private Sub BuildProgressBar()
    'This will rebuild the progress bar picture
    'from scratch, using the current properties
    
    Dim udtTempFill As BitmapStruc      'this is a bitmap of the section of the progress bar that is being filled
    Dim udtTempBack As BitmapStruc      'this is a bitmap of the section of the progress bar that is not yet filled
    Dim strCaption As String            'the text to display in the progress bar
    Dim sngProgress As Single           'the amount of relative space taken up by Value
    Dim lngTextHeight As Long           'the height of the text
    Dim lngResult As Long               'any error value returned by the api call
    
    'set the border style if necessary
    If picSurphase.Appearance <> menmBorder Then
        picSurphase.Appearance = menmBorder
    End If
    
    'get the point to fill the progress bar to
    If msngMax > 0 Then
        sngProgress = (1 / (msngMax - msngMin)) _
                      * (msngValue - msngMin)
    End If
    
    'set the text to the default if necessary
    If mblnDefCapt Then
        'get the current percentage
        strCaption = Int(sngProgress * 100) & "%"
    Else
        strCaption = mstrCaption
    End If
    
    'set up the bitmaps
    With udtTempFill
        .Area = mudtBackBmp.Area
        
        'create the bitmap
        Call CreateNewBitmap(.hDcMemory, _
                             .hDcBitmap, _
                             .hDcPointer, _
                             .Area, _
                             UserControl.hDc, _
                             mlngFillColour)
    End With
    
    With udtTempBack
        .Area = udtTempFill.Area
        
        'create the bitmap
        Call CreateNewBitmap(.hDcMemory, _
                             .hDcBitmap, _
                             .hDcPointer, _
                             .Area, _
                             UserControl.hDc, _
                             mlngBackColour)
    End With
    
    'get the text width
    lngTextHeight = UserControl.TextHeight(strCaption)
    
    'draw the text on the filled section
    mfntCaption.Colour = mlngOverColour
    Call MakeText(udtTempFill.hDcMemory, _
                  strCaption, _
                  (udtTempFill.Area.Bottom - lngTextHeight) / 2, _
                  0, _
                  lngTextHeight, _
                  udtTempFill.Area.Right, _
                  mfntCaption, _
                  InPixels)
    
    mfntCaption.Colour = mlngTextColour
    Call MakeText(udtTempBack.hDcMemory, _
                  strCaption, _
                  (udtTempBack.Area.Bottom - lngTextHeight) / 2, _
                  0, _
                  lngTextHeight, _
                  udtTempBack.Area.Right, _
                  mfntCaption, _
                  InPixels)
    
    'copy the proper sections of the bitmaps
    'onto the background bitmap
    lngResult = BitBlt(mudtBackBmp.hDcMemory, _
                       0, _
                       0, _
                       Int(udtTempFill.Area.Right * sngProgress), _
                       udtTempFill.Area.Bottom, _
                       udtTempFill.hDcMemory, _
                       0, _
                       0, _
                       SRCCOPY)
    
    lngResult = BitBlt(mudtBackBmp.hDcMemory, _
                       Int(udtTempBack.Area.Right * sngProgress), _
                       0, _
                       udtTempBack.Area.Right, _
                       udtTempBack.Area.Bottom, _
                       udtTempBack.hDcMemory, _
                       Int(udtTempBack.Area.Right * sngProgress), _
                       0, _
                       SRCCOPY)
    
    'remove the bitmaps from memory
    Call DeleteBitmap(udtTempFill.hDcMemory, _
                      udtTempFill.hDcBitmap, _
                      udtTempFill.hDcPointer)
    Call DeleteBitmap(udtTempBack.hDcMemory, _
                      udtTempBack.hDcBitmap, _
                      udtTempBack.hDcPointer)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Rebuilds the progress bar from scratch"
    'This will update the display on the
    'progress bar
    Call BuildProgressBar
    Call ShowProgressBar
End Sub

Private Sub ShowProgressBar()
    'display the background
    
    Dim lngResult As Long   'holds any error value returned from the api call
    
    With mudtBackBmp
        lngResult = BitBlt(picSurphase.hDc, _
                           -menmBorder, _
                           -menmBorder, _
                           .Area.Right, _
                           .Area.Bottom, _
                           .hDcMemory, _
                           0, _
                           0, _
                           SRCCOPY)
    End With
End Sub

'events

Private Sub picSurphase_Click()
    'raise a click event for the progress bar
    RaiseEvent Click
End Sub

Private Sub picSurphase_DblClick()
    'raise a double click event for the progress bar
    RaiseEvent DblClick
End Sub

Private Sub picSurphase_Paint()
    Call ShowProgressBar
End Sub

Private Sub UserControl_Initialize()
    'set up the default values before
    'reading the property bag
    
    'create the background bitmap
    With mudtBackBmp
        .Area.Right = UserControl.ScaleWidth - UserControl.ScaleLeft
        .Area.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop
        
        Call CreateNewBitmap(.hDcMemory, _
                             .hDcBitmap, _
                             .hDcPointer, _
                             .Area, _
                             UserControl.hDc, _
                             vbWhite)
    End With
End Sub

Private Sub picSurphase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse down event for the progress bar
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picSurphase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse move event for the progress bar
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picSurphase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'raise a mouse up event for the progress bar
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
    'set the default properties for the control
    
    menmBorder = pgrFixed_Single
    msngMax = 100
    msngMin = 0
    msngValue = 0
    mblnDefCapt = True
    mstrCaption = ""
    mlngBackColour = vbWhite
    mlngFillColour = vbDarkBlue
    mlngTextColour = vbBlack
    mlngOverColour = vbWhite
    mfntCaption.Alignment = vbCentreAlign
    Set UserControl.Font = Ambient.Font
    Call SetFont
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'get the properties from the property bag
    
    With PropBag
        msngMax = .ReadProperty("Max", 100)
        msngMin = .ReadProperty("Min", 0)
        msngValue = .ReadProperty("Value", 0)
        mblnDefCapt = .ReadProperty("PercentCaption", True)
        mstrCaption = .ReadProperty("Caption", "")
        mlngBackColour = .ReadProperty("BackColour", vbWhite)
        mlngFillColour = .ReadProperty("FillColour", vbDarkBlue)
        mlngTextColour = .ReadProperty("TextColour", vbBlack)
        mlngOverColour = .ReadProperty("OverColour", vbWhite)
        menmBorder = .ReadProperty("BorderStyle", pgrFixed_Single)
        mfntCaption.Alignment = vbCentreAlign
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        Call SetFont
    End With
    
    'refresh the progress bar display
    Call Refresh
End Sub

Private Sub UserControl_Resize()
    'set the bitmap dimensions
    With mudtBackBmp
        .Area.Right = UserControl.ScaleWidth - (UserControl.ScaleLeft * 2)
        .Area.Bottom = UserControl.ScaleHeight - (UserControl.ScaleTop * 2)
    End With
    Call picSurphase.Move(0, _
                          0, _
                          UserControl.ScaleWidth, _
                          UserControl.ScaleHeight)
    Call BuildProgressBar
    Call ShowProgressBar
End Sub

Private Sub UserControl_Terminate()
    'remove the background bitmap from memory
    With mudtBackBmp
        Call DeleteBitmap(.hDcMemory, _
                          .hDcBitmap, _
                          .hDcPointer)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save the current properties
    
    With PropBag
        Call .WriteProperty("Max", msngMax, 100)
        Call .WriteProperty("Min", msngMin, 0)
        Call .WriteProperty("Value", msngValue, 0)
        Call .WriteProperty("Caption", mstrCaption, "")
        Call .WriteProperty("PercentCaption", mblnDefCapt, True)
        Call .WriteProperty("BackColour", mlngBackColour, vbWhite)
        Call .WriteProperty("FillColour", mlngFillColour, vbDarkBlue)
        Call .WriteProperty("TextColour", mlngTextColour, vbBlack)
        Call .WriteProperty("OverColour", mlngOverColour, vbWhite)
        Call .WriteProperty("BorderStyle", menmBorder, pgrFixed_Single)
        Call GetFont
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
    End With
End Sub
